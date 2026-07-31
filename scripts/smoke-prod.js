#!/usr/bin/env node
/**
 * smoke-prod.js — 本番デプロイに対する read-only 統合スモークテスト。
 *
 * なぜ必要か:
 *   tests/*.test.cjs は vm.createContext で GAS API を stub したユニットテストなので、
 *   「stub と実 GAS の挙動が違う」クラスのバグを一切検出できない。実際 v2875 の
 *   isBoardCollaborator envelope バグ (getUserConfig の戻り値を unwrap し忘れ) は
 *   ユニットテストを書いて初めて発覚した。本スクリプトは **実 GAS / 実 Sheets /
 *   実 SA pool** を通して、デプロイ後の本番が本当に生きているかを検証する。
 *
 * 安全性:
 *   **read-only 操作のみ**。書き込み系 (setUserConfig / toggleUserBoard / publish 等) は
 *   一切呼ばない。いつ実行しても本番状態を変えない。
 *
 * 使い方:
 *   npm run smoke              # 那覇 (default)
 *   npm run smoke:open         # 沖縄 (open)
 *   node scripts/smoke-prod.js --env open --verbose
 *
 * CI には入れない: 認証情報 (scripts/config.json) と本番疎通が必要で、GAS quota も消費するため。
 * デプロイ直後の手動ゲートとして使う (`npm run deploy:all && npm run smoke`)。
 */

'use strict';

const { spawnSync } = require('node:child_process');
const path = require('node:path');

const ARGS = process.argv.slice(2);
const VERBOSE = ARGS.includes('--verbose');
const envIdx = ARGS.indexOf('--env');
const ENV = envIdx >= 0 && ARGS[envIdx + 1] ? ARGS[envIdx + 1] : null;
const API = path.join(__dirname, 'admin-api.js');

const results = [];
function record(name, ok, detail) {
  results.push({ name, ok, detail });
  console.log(`  ${ok ? '✅' : '❌'}  ${name}`);
  if (detail) console.log(`        ${detail}`);
}

// メールは本番の個人情報なのでログに残さない。ドメインと先頭 2 文字だけ出す。
function maskEmail(email) {
  const s = String(email || '');
  const at = s.indexOf('@');
  if (at <= 0) return '(不明)';
  return `${s.slice(0, 2)}***${s.slice(at)}`;
}

// admin-api.js を spawn して JSON を返す。失敗時は { __error } を返し呼び出し側で扱う。
function callApi(operation, extraArgs = []) {
  const argv = [API, operation, ...extraArgs];
  if (ENV) argv.push('--env', ENV);
  const r = spawnSync('node', argv, { encoding: 'utf8', timeout: 120000 });
  if (r.error) return { __error: `spawn 失敗: ${r.error.message}` };
  const out = String(r.stdout || '');
  const start = out.indexOf('{');
  if (start < 0) {
    return { __error: `JSON が返らなかった (exit=${r.status}) ${String(r.stderr || '').slice(0, 200)}` };
  }
  try {
    return JSON.parse(out.slice(start));
  } catch (e) {
    return { __error: `JSON parse 失敗: ${e.message}` };
  }
}

function main() {
  const label = ENV ? `env=${ENV}` : 'env=default (那覇)';
  console.log('\n══════════════════════════════════════════════════════════════');
  console.log(`  本番スモークテスト (read-only) — ${label}`);
  console.log('══════════════════════════════════════════════════════════════\n');

  // 1. アプリが有効か (緊急停止されていないか)
  const status = callApi('getAppStatus');
  const statusOk = !status.__error && status.success === true && status.data
    && status.data.isDisabled === false;
  record('1. アプリが有効 (緊急停止されていない)', statusOk,
    status.__error || `status=${status.data && status.data.status}`);

  // 2. システム診断が全項目 PASS (DB 接続 / Core props / デプロイ)
  const diag = callApi('systemDiagnosis');
  let diagOk = false;
  let diagDetail = diag.__error || '';
  if (!diag.__error && diag.success) {
    const json = JSON.stringify(diag);
    const fails = (json.match(/"status"\s*:\s*"FAIL"/g) || []).length;
    const passes = (json.match(/"status"\s*:\s*"PASS"/g) || []).length;
    diagOk = fails === 0 && passes > 0;
    diagDetail = `PASS=${passes}, FAIL=${fails}`;
  }
  record('2. systemDiagnosis 全項目 PASS', diagOk, diagDetail);

  // 3. users DB が実 Sheets から読め、レコード形状が正しい
  //    (SA pool → 実 Spreadsheet 読み取りを通す。stub では検証できない経路)
  const users = callApi('getUsers');
  const list = (!users.__error && (users.users || users.data)) || [];
  const arr = Array.isArray(list) ? list : [];
  const REQUIRED_USER_KEYS = ['userId', 'userEmail', 'configJson'];
  const shapeOk = arr.length > 0 && arr.every((u) => u && REQUIRED_USER_KEYS.every((k) => k in u));
  record('3. users DB 読み取り + レコード形状', shapeOk,
    users.__error || `${arr.length} 件, 必須キー=${REQUIRED_USER_KEYS.join('/')}`);

  // 4. configJson が全件 JSON として妥当 (破損 config は viewer 側で事故になる)
  let broken = 0;
  for (const u of arr) {
    const raw = u && u.configJson;
    if (!raw) continue;
    if (typeof raw === 'object') continue;
    try { JSON.parse(raw); } catch (_) { broken++; }
  }
  record('4. configJson が全件 parse 可能', broken === 0,
    `破損 ${broken} 件 / ${arr.length} 件`);

  // 5. Script Properties に必須キーが揃っている
  const props = callApi('listProperties');
  const REQUIRED_PROPS = ['ADMIN_EMAIL', 'DATABASE_SPREADSHEET_ID', 'SERVICE_ACCOUNT_CREDS', 'ADMIN_API_KEY'];
  const propsJson = props.__error ? '' : JSON.stringify(props);
  const missing = REQUIRED_PROPS.filter((k) => !propsJson.includes(k));
  record('5. 必須 Script Properties が存在', !props.__error && missing.length === 0,
    props.__error || (missing.length ? `欠落: ${missing.join(', ')}` : REQUIRED_PROPS.join(', ')));

  // 6. previewBoard が実データを返す — 本スモークの中核。
  //    SA pool 経由で実 Spreadsheet を開き、列マッピングを適用して viewer 視点の
  //    データを組み立てる経路を丸ごと通す。stub テストでは決して再現できない。
  const candidate = arr.find((u) => {
    if (!u || u.isActive === false) return false;
    try {
      const cfg = typeof u.configJson === 'object' ? u.configJson : JSON.parse(u.configJson || '{}');
      return Boolean(cfg && cfg.spreadsheetId);
    } catch (_) { return false; }
  });
  if (!candidate) {
    record('6. previewBoard で実 Sheets 読み取り', false,
      'データソース接続済みユーザーが 1 人もいない (検証不能)');
  } else {
    const pv = callApi('previewBoard', ['--userId', candidate.userId]);
    const pvOk = !pv.__error && pv.success === true;
    record('6. previewBoard で実 Sheets 読み取り', pvOk,
      pv.__error || `user=${maskEmail(candidate.userEmail)} / ${pvOk ? 'データ取得 OK' : (pv.message || '失敗')}`);
    if (VERBOSE && pvOk) console.log(`        ${JSON.stringify(pv).slice(0, 300)}`);
  }

  // 7. キャッシュ / 外部サービスが利用可能 (SA token, Sheets API 疎通)
  const perf = callApi('perfMetrics');
  const perfJson = perf.__error ? '' : JSON.stringify(perf);
  const perfOk = !perf.__error && perf.success === true
    && perfJson.includes('available') && !perfJson.includes('"unavailable"');
  record('7. Cache / 外部サービスが available', perfOk,
    perf.__error || 'scriptCache / SpreadsheetApp / UrlFetchApp');

  const passed = results.filter((r) => r.ok).length;
  console.log('\n──────────────────────────────────────────────────────────────');
  console.log(`  結果: ${passed} / ${results.length} PASS`);
  console.log('══════════════════════════════════════════════════════════════\n');

  if (passed !== results.length) {
    console.error('本番スモークテスト失敗。上の ❌ を確認してください。');
    console.error('  調査: npm run logs:cloud -- --severity ERROR --hours 1');
    process.exit(1);
  }
}

main();
