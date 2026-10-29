#!/usr/bin/env node
/**
 * 先生 1 人専用の配置 (スプレッドシート紐づけ・先生として実行) を作る / 更新する。
 *
 * Why: 2026-10-29 10:20 ごろから、児童アカウントは「開いた人として実行」の Apps Script を
 *   中身に関係なく開けなくなった (管理側の設定)。本体のコード (src/) をそのまま使い、
 *   次の 2 ファイルだけを差し替えて、先生ごとに「先生として実行」で配置する。
 *     - appsscript.json: webapp.executeAs = USER_DEPLOYING (先生として実行)
 *     - Tenancy.js: DEPLOY_TENANCY = 'single' と、その先生のメール・DB にするスプレッドシート
 *   画面・機能は本体と同じ。違いは「データが先生のスプレッドシート 1 つに入る」ことだけ。
 *
 * 使い方:
 *   node scripts/deploy-single.js create --name <名前> --adminEmail <先生のメール> [--user organization]
 *   node scripts/deploy-single.js deploy --name <名前>      # 1 人分を更新 (URL はそのまま)
 *   node scripts/deploy-single.js deploy --all              # 全員分を更新
 *   node scripts/deploy-single.js list
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
  return execFileSync('npx', ['clasp', '--user', user, '-P', '.', '-I', '.', ...args], {
    cwd: stageDir, encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe']
  });
}

function checkGeneratedFresh() {
  for (const script of ['gen-pagehead.js', 'gen-utilities.js']) {
    execFileSync('node', [path.join(__dirname, script), '--check'], { stdio: 'inherit' });
  }
}

// src/ を写し、2 ファイルだけ差し替える。
function stage(tenant) {
  const dir = path.join(STAGE_ROOT, tenant.name);
  fs.rmSync(dir, { recursive: true, force: true });
  fs.mkdirSync(dir, { recursive: true });
  for (const f of fs.readdirSync(SRC)) {
    if (!/\.(js|html)$/.test(f)) continue;
    fs.copyFileSync(path.join(SRC, f), path.join(dir, f));
  }
  const manifest = JSON.parse(fs.readFileSync(path.join(SRC, 'appsscript.json'), 'utf8'));
  manifest.webapp = { executeAs: 'USER_DEPLOYING', access: 'DOMAIN' };
  fs.writeFileSync(path.join(dir, 'appsscript.json'), JSON.stringify(manifest, null, 2) + '\n');
  const str = (v) => JSON.stringify(String(v || ''));
  fs.writeFileSync(path.join(dir, 'Tenancy.js'),
    '// scripts/deploy-single.js が書いたファイル (先生専用の配置)。src/Tenancy.js の説明を参照。\n' +
    `const DEPLOY_TENANCY = 'single';\n` +
    `const DEPLOY_ADMIN_EMAIL = ${str(tenant.adminEmail)};\n` +
    `const DEPLOY_DATABASE_SPREADSHEET_ID = ${str(tenant.spreadsheetId)};\n`);
  fs.writeFileSync(path.join(dir, '.claspignore'), '**/**\n!appsscript.json\n!*.js\n!*.html\n');
  if (tenant.scriptId) {
    fs.writeFileSync(path.join(dir, '.clasp.json'), JSON.stringify({ scriptId: tenant.scriptId, rootDir: '' }, null, 2));
  }
  return dir;
}

function create(args) {
  const name = args.name;
  const adminEmail = args.adminEmail;
  const user = args.user || 'organization';
  if (!name || !adminEmail) throw new Error('--name と --adminEmail が必要です');
  if (!/^[\w-]+$/.test(name)) throw new Error('--name は英数字・-・_ だけにしてください (フォルダ名に使う)');
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
  deployOne(tenant, cfg);
}

function deployOne(tenant, cfg) {
  const dir = stage(tenant);
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
  checkGeneratedFresh();
  if (cmd === 'create') return create(args);
  if (cmd === 'deploy') {
    const cfg = loadConfig();
    const targets = args.all ? cfg.tenants : cfg.tenants.filter((t) => t.name === args.name);
    if (!targets.length) throw new Error('対象がありません (--name <名前> か --all)');
    for (const t of targets) deployOne(t, cfg);
    return;
  }
  throw new Error('使い方: create / deploy / list (ファイル冒頭の説明を参照)');
}

try {
  main();
} catch (e) {
  console.error('✗', e.message);
  process.exit(1);
}
