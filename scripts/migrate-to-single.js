#!/usr/bin/env node
/**
 * 共有の本体 (サイロ型マルチテナント) から、先生専用の配置へ授業とボードの設定を引っ越す。
 *
 * 元の DB は先生から見えない (SA にだけ共有) ので、元の本体の admin API で読み出し
 * (lesson.list → lesson.review / getUserConfig)、先生専用の配置の admin API
 * (migrate.importLesson / migrate.importConfig、src/Migration.js) へ渡す。
 * 元の本体のデータは消さない (写すだけ)。何度流しても二重にならない。
 *
 * 児童の回答はこのスクリプトの中を通るだけで、画面には授業名と件数しか出さない。
 *
 * 注意: 元の本体の SA の鍵は 2026-10-29 に削除した。流す前に answer-board-db-manager の鍵を 1 本だけ
 *   作り直して元の本体の SERVICE_ACCOUNT_CREDS に入れ、元の本体を一時的に配置し直す。終わったら
 *   配置を外して鍵を消す (元の DB は SA にだけ共有されているため)。
 *
 * 使い方:
 *   node scripts/migrate-to-single.js --name <先生専用の名前> --userId <元の本体の userId> [--dry-run]
 *
 * 前提: 先生が先生専用の配置の管理パネルを 1 度開いている (登録済み)。
 */
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { getAccessToken, postJSONSync, getConfig } = require('./lib/gas-auth');

const TENANTS = path.join(__dirname, 'single-tenants.json');

function parseArgs(argv) {
  const out = {};
  for (let i = 0; i < argv.length; i++) {
    if (!argv[i].startsWith('--')) continue;
    const key = argv[i].slice(2);
    const next = argv[i + 1];
    if (next === undefined || next.startsWith('--')) out[key] = true;
    else { out[key] = next; i++; }
  }
  return out;
}

function main() {
  const args = parseArgs(process.argv.slice(2));
  if (!args.name || !args.userId) throw new Error('--name と --userId が必要です');
  const tenantsCfg = JSON.parse(fs.readFileSync(TENANTS, 'utf8'));
  const tenant = tenantsCfg.tenants.find((t) => t.name === args.name);
  if (!tenant || !tenant.deploymentId) throw new Error(`先生専用の配置「${args.name}」が見つかりません (deploy-single.js create を先に)`);

  // 元の本体 (scripts/config.json)
  const oldCfg = getConfig(null);
  const oldToken = getAccessToken(oldCfg.tokenName);
  const callOld = (operation, params) => {
    const res = postJSONSync(oldCfg.prodUrl, { action: 'adminApi', apiKey: oldCfg.apiKey, operation, params }, oldToken);
    if (!res || !res.success) throw new Error(`元の本体 ${operation} が失敗: ${(res && (res.message || res.error)) || 'no response'}`);
    return res;
  };

  // 先生専用の配置
  const domain = tenant.adminEmail.split('@')[1];
  const newUrl = `https://script.google.com/a/macros/${domain}/s/${tenant.deploymentId}/exec`;
  const newToken = getAccessToken(tenant.user || 'organization');
  if (!tenant.apiKey) {
    // 初回だけ: 管理者本人 (先生) のトークンで API キーを決める (main.js handleSetupApiKeyAction_)。
    const key = crypto.randomBytes(24).toString('base64url');
    const res = postJSONSync(newUrl, { action: 'setupApiKey', apiKey: key }, newToken);
    if (!res || !res.success) throw new Error('先生専用の配置の API キーを設定できませんでした: ' + ((res && (res.message || res.error)) || ''));
    tenant.apiKey = key;
    fs.writeFileSync(TENANTS, JSON.stringify(tenantsCfg, null, 2) + '\n');
    console.log('✓ 先生専用の配置の API キーを設定しました (scripts/single-tenants.json に保存)');
  }
  const callNew = (operation, params) => postJSONSync(newUrl, { action: 'adminApi', apiKey: tenant.apiKey, operation, params }, newToken);

  // 1. 授業 (古い順に取り込む)
  const lessons = callOld('lesson.list', { userId: args.userId }).data.lessons
    .slice().sort((a, b) => String(a.createdAt).localeCompare(String(b.createdAt)));
  console.log(`授業 ${lessons.length} 件`);
  for (const item of lessons) {
    const lesson = callOld('lesson.review', { userId: args.userId, lessonId: item.lessonId }).data.lesson;
    const snaps = (lesson.lessonJson && lesson.lessonJson.snapshots) || [];
    const rows = snaps.reduce((n, s) => n + ((s.rows && s.rows.length) || 0), 0);
    if (args['dry-run']) {
      console.log(`  (dry-run) ${item.state}\t${lesson.name}\tアーカイブ ${rows} 行`);
      continue;
    }
    const res = callNew('migrate.importLesson', { lesson });
    if (!res || !res.success) throw new Error(`「${lesson.name}」の引っ越しに失敗: ${(res && (res.message || res.error)) || 'no response'}`);
    const d = res.data || {};
    console.log(d.skipped
      ? `  - ${lesson.name}: 既に引っ越し済み`
      : `  ✓ ${lesson.name}: 回答シート ${d.sheets} 枚・アーカイブ ${d.archiveRows} 行`);
  }

  // 2. ボードの設定 (授業の後: ボードにつながっている授業をかけ直すため)
  const config = callOld('getUserConfig', { userId: args.userId }).data.config;
  if (args['dry-run']) {
    console.log(`(dry-run) ボードの設定: 公開 ${Boolean(config.isPublished)}・ボードの授業 ${config.activeLessonId || 'なし'}`);
    return;
  }
  const res = callNew('migrate.importConfig', { config });
  if (!res || !res.success) throw new Error('ボードの設定の引っ越しに失敗: ' + ((res && (res.message || res.error)) || ''));
  console.log(`✓ ボードの設定: 公開 ${res.data.published}・かけ直した授業 ${res.data.reappliedLessonId || 'なし'}`);
}

try {
  main();
} catch (e) {
  console.error('✗', e.message);
  process.exit(1);
}
