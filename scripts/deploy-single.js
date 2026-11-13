#!/usr/bin/env node
/**
 * 先生ごとの配置 (スプレッドシート紐づけ・先生として実行) を作る / 更新する。これが唯一の配置の道具。
 *
 * Why: 2026-10-29 10:20 ごろから、児童アカウントは「開いた人として実行」の Apps Script を
 *   中身に関係なく開けなくなった (管理側の設定)。src/ をそのまま、先生ごとに「先生として実行」
 *   (src/appsscript.json の webapp) で配置する。差し替えるのは Deployment.js だけ:
 *   その先生のメールと、データを入れるスプレッドシート (DB) を焼き込む。
 *
 * 使い方:
 *   node scripts/deploy-single.js create --name <名前> --adminEmail <先生のメール> [--user organization]
 *   node scripts/deploy-single.js deploy --name <名前>      # 1 人分を更新 (URL はそのまま)
 *   node scripts/deploy-single.js deploy --all              # 全員分を更新
 *   node scripts/deploy-single.js list
 *   scripts/ops-report.local.json (エラーの報告先) が無いと配置しない。報告なしで配置するなら --no-ops-report。
 *
 * 先生ごとの値は scripts/single-tenants.json (git に入れない)。
 */
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

const ROOT = path.resolve(__dirname, '..');
const SRC = path.join(ROOT, 'src');
const CONFIG = path.join(__dirname, 'single-tenants.json');
// clasp は作業フォルダがリポジトリの外 (/private/tmp 等) だと「escapes project root」で止まるので、
//   リポジトリの中の git に入らないフォルダで組み立てる。
const STAGE_ROOT = path.join(ROOT, '.single-stage');
// 配置の名前はフォルダ名に使う (.single-stage/<名前>/)
const TENANT_NAME_RE = /^[\w-]+$/;

function parseArgs(argv) {
  const out = { _: [] };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a.startsWith('--')) {
      const key = a.slice(2);
      const next = argv[i + 1];
      if (next === undefined || next.startsWith('--')) out[key] = true;
      else { out[key] = next; i++; }
    } else {
      out._.push(a);
    }
  }
  return out;
}

function loadConfig() {
  if (!fs.existsSync(CONFIG)) return { tenants: [] };
  return JSON.parse(fs.readFileSync(CONFIG, 'utf8'));
}

function saveConfig(cfg) {
  fs.writeFileSync(CONFIG, JSON.stringify(cfg, null, 2) + '\n');
}

function clasp(stageDir, user, args) {
  // -P / -I で作業フォルダの .clasp.json / .claspignore を明示する。付けないと clasp は
  //   親フォルダ (リポジトリ直下) の本体の .clasp.json を拾い、本体に push してしまう。
  // Windows の npx は .cmd なので shell なしでは起動できない (ENOENT)。node に同梱の npx-cli.js を node で直接呼ぶ。
  const [cmd, pre] = process.platform === 'win32'
    ? [process.execPath, [path.join(path.dirname(process.execPath), 'node_modules', 'npm', 'bin', 'npx-cli.js')]]
    : ['npx', []];
  // 絶対パスで渡す: clasp 3.4 は -P の場所を相対のまま持ち、絶対パスにした srcDir と比べるので、
  //   '.' だと「srcDir "" escapes project root」で止まる (3.1 では通っていた)。
  return execFileSync(cmd, [...pre, 'clasp', '--user', user, '-P', stageDir, '-I', stageDir, ...args], {
    cwd: stageDir, encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe']
  });
}

function checkGeneratedFresh() {
  for (const script of ['gen-pagehead.js', 'gen-utilities.js']) {
    execFileSync('node', [path.join(__dirname, script), '--check'], { stdio: 'inherit' });
  }
}

// ops-monitor への報告の送り先と鍵 (全配置で共通。git に入れない)。
//   形: { "url": "https://…/report", "token": "ops_…" }。値は画面に出さない。
//   Why 無ければ止める: このファイルは git に入らないので、別の PC に clone しただけでは無い。
//   無いまま配置すると送り先が空で焼き込まれ、本番のエラーの報告が黙って止まる (2026-11-11 に気づいた)。
//   報告なしで配置したいときだけ --no-ops-report を付ける。
const OPS_REPORT_CONFIG = path.join(__dirname, 'ops-report.local.json');
function loadOpsReport(opts) {
  const o = opts || {};
  if (o.skip) return { url: '', token: '' };
  const file = o.file || OPS_REPORT_CONFIG;
  const hint = '配置すると本番のエラーの報告が止まるので、配置しません。' +
    '配置している PC の scripts/ops-report.local.json をこの PC の scripts/ に写してください ' +
    '(報告なしで配置するなら --no-ops-report)。';
  if (!fs.existsSync(file)) throw new Error('scripts/ops-report.local.json がありません。' + hint);
  let j = null;
  try {
    j = JSON.parse(fs.readFileSync(file, 'utf8'));
  } catch (e) {
    // JSON.parse のエラー文は中身の一部を含むことがある (鍵が画面に出る) ので、文は出さない
    throw new Error('scripts/ops-report.local.json を JSON として読めません。' + hint);
  }
  // 文字列でないもの ({} など) を String() にすると '[object Object]' になり、空でないので通ってしまう
  const text = (v) => (typeof v === 'string' ? v.trim() : '');
  const url = text(j && j.url);
  const token = text(j && j.token);
  if (!url || !token) throw new Error('scripts/ops-report.local.json に url か token がありません。' + hint);
  // new URL のエラー文は入力を含むことがあるので、文は出さない
  let parsed = null;
  try { parsed = new URL(url); } catch (e) { parsed = null; }
  if (!parsed || parsed.protocol !== 'https:' || !parsed.hostname || /\s/.test(url)) {
    throw new Error('scripts/ops-report.local.json の url が https:// の URL ではありません。' + hint);
  }
  return { url, token };
}

// src/ を写し、Deployment.js だけ差し替える。
function stage(tenant, ops) {
  // この下でフォルダごと消すので、名前が STAGE_ROOT の外を指さないことを確かめる
  //   (create は名前を検証するが、single-tenants.json を手で書き換えた値はここまで素通りする)
  if (!TENANT_NAME_RE.test(String(tenant.name || ''))) throw new Error('single-tenants.json の name が英数字・-・_ だけではありません');
  const dir = path.join(STAGE_ROOT, tenant.name);
  fs.rmSync(dir, { recursive: true, force: true });
  fs.mkdirSync(dir, { recursive: true });
  for (const f of fs.readdirSync(SRC)) {
    if (!/\.(js|html)$/.test(f)) continue;
    fs.copyFileSync(path.join(SRC, f), path.join(dir, f));
  }
  fs.copyFileSync(path.join(SRC, 'appsscript.json'), path.join(dir, 'appsscript.json'));
  const str = (v) => JSON.stringify(String(v || ''));
  fs.writeFileSync(path.join(dir, 'Deployment.js'),
    '// scripts/deploy-single.js が書いたファイル。src/Deployment.js の説明を参照。\n' +
    `const DEPLOY_ADMIN_EMAIL = ${str(tenant.adminEmail)};\n` +
    `const DEPLOY_DATABASE_SPREADSHEET_ID = ${str(tenant.spreadsheetId)};\n` +
    `const DEPLOY_OPS_REPORT_URL = ${str(ops.url)};\n` +
    `const DEPLOY_OPS_REPORT_TOKEN = ${str(ops.token)};\n`);
  // 検証用の配置にだけ足すファイル (single-tenants.json の extraFiles、リポジトリからの相対パス)。
  //   例: scripts/verify/DevInvoke.js (画面の処理を admin API から直接呼ぶ入口)。本番の先生には付けない。
  for (const rel of tenant.extraFiles || []) {
    fs.copyFileSync(path.join(ROOT, rel), path.join(dir, path.basename(rel)));
  }
  fs.writeFileSync(path.join(dir, '.claspignore'), '**/**\n!appsscript.json\n!*.js\n!*.html\n');
  if (tenant.scriptId) {
    fs.writeFileSync(path.join(dir, '.clasp.json'), JSON.stringify({ scriptId: tenant.scriptId, rootDir: '' }, null, 2));
  }
  return dir;
}

function create(args, ops) {
  const name = args.name;
  const adminEmail = args.adminEmail;
  const user = args.user || 'organization';
  if (!name || !adminEmail) throw new Error('--name と --adminEmail が必要です');
  if (!TENANT_NAME_RE.test(name)) throw new Error('--name は英数字・-・_ だけにしてください (フォルダ名に使う)');
  const cfg = loadConfig();
  if (cfg.tenants.some((t) => t.name === name)) throw new Error(`「${name}」は既にあります (deploy を使ってください)`);

  const dir = path.join(STAGE_ROOT, name);
  fs.rmSync(dir, { recursive: true, force: true });
  fs.mkdirSync(dir, { recursive: true });
  const title = args.title || 'みんなの回答ボード';
  clasp(dir, user, ['create', '--type', 'sheets', '--title', title, '--rootDir', '.']);
  const project = JSON.parse(fs.readFileSync(path.join(dir, '.clasp.json'), 'utf8'));
  const tenant = { name, adminEmail: adminEmail.toLowerCase(), user, scriptId: project.scriptId, spreadsheetId: project.parentId, deploymentId: '' };
  cfg.tenants.push(tenant);
  saveConfig(cfg);
  console.log(`✓ スプレッドシートとスクリプトを作りました: ${name}`);
  deployOne(tenant, cfg, ops);
}

function deployOne(tenant, cfg, ops) {
  const dir = stage(tenant, ops);
  const user = tenant.user || 'organization';
  console.log(`▶ ${tenant.name}: push`);
  clasp(dir, user, ['push', '--force']);
  const desc = `deploy-single ${new Date().toISOString().slice(0, 16)}`;
  const out = tenant.deploymentId
    ? clasp(dir, user, ['deploy', '-i', tenant.deploymentId, '-d', desc])
    : clasp(dir, user, ['deploy', '-d', desc]);
  const m = /Deployed (\S+) @(\d+)/.exec(out);
  if (!m) throw new Error('デプロイの結果を読めませんでした: ' + out.trim());
  if (!tenant.deploymentId) {
    tenant.deploymentId = m[1];
    saveConfig(cfg);
  }
  console.log(`✓ ${tenant.name}: @${m[2]}  https://script.google.com/a/macros/${tenant.adminEmail.split('@')[1]}/s/${tenant.deploymentId}/exec`);
}

function main() {
  const args = parseArgs(process.argv.slice(2));
  const cmd = args._[0];
  if (cmd === 'list') {
    for (const t of loadConfig().tenants) console.log(`${t.name}\t${t.adminEmail}\t${t.deploymentId || '(未デプロイ)'}`);
    return;
  }
  if (cmd !== 'create' && cmd !== 'deploy') throw new Error('使い方: create / deploy / list (ファイル冒頭の説明を参照)');
  // 何かを作る・push する前に確かめる (1 人目だけ配置して止まる、を避ける)
  const ops = loadOpsReport({ skip: args['no-ops-report'] === true });
  if (!ops.url) console.log('⚠ --no-ops-report: エラーの報告先を空にして配置します (この配置からは報告が届きません)');
  checkGeneratedFresh();
  if (cmd === 'create') return create(args, ops);
  if (cmd === 'deploy') {
    const cfg = loadConfig();
    const targets = args.all ? cfg.tenants : cfg.tenants.filter((t) => t.name === args.name);
    if (!targets.length) throw new Error('対象がありません (--name <名前> か --all)');
    for (const t of targets) deployOne(t, cfg, ops);
  }
}

if (require.main === module) {
  try {
    main();
  } catch (e) {
    console.error('✗', e.message);
    process.exit(1);
  }
}

module.exports = { loadOpsReport, stage };
