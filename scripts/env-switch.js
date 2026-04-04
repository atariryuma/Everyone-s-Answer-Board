/**
 * 環境切替スクリプト
 * .clasp.{env}.json → .clasp.json にスワップしてコマンドを実行し、終了後に復元する
 * clasp v3 は token フィールドを無視して常に default トークンを使うため、
 * .clasprc.json の default トークンもスワップする
 *
 * Usage: node scripts/env-switch.js <env> <command...>
 * Example: node scripts/env-switch.js open npx clasp push --force
 */
const fs = require('fs');
const path = require('path');
const os = require('os');
const { execFileSync } = require('child_process');

const ROOT = path.join(__dirname, '..');
const CLASP_JSON = path.join(ROOT, '.clasp.json');
const CLASP_BACKUP = path.join(ROOT, '.clasp.json.bak');

function findClasprc() {
  const local = path.join(ROOT, '.clasprc.json');
  if (fs.existsSync(local)) return local;
  const global = path.join(os.homedir(), '.clasprc.json');
  if (fs.existsSync(global)) return global;
  return null;
}

function swapDefaultToken(clasprcPath, tokenName) {
  if (!clasprcPath || !tokenName) return null;
  const data = JSON.parse(fs.readFileSync(clasprcPath, 'utf8'));
  if (!data.tokens?.[tokenName]) {
    console.error(`Token "${tokenName}" not found in .clasprc.json`);
    return null;
  }
  const originalDefault = data.tokens.default;
  data.tokens.default = data.tokens[tokenName];
  fs.writeFileSync(clasprcPath, JSON.stringify(data, null, 2));
  return originalDefault;
}

function restoreDefaultToken(clasprcPath, originalDefault) {
  if (!clasprcPath || !originalDefault) return;
  const data = JSON.parse(fs.readFileSync(clasprcPath, 'utf8'));
  data.tokens.default = originalDefault;
  fs.writeFileSync(clasprcPath, JSON.stringify(data, null, 2));
}

function main() {
  const [env, ...cmdArgs] = process.argv.slice(2);

  if (!env || !cmdArgs.length) {
    console.error('Usage: node scripts/env-switch.js <env> <command...>');
    process.exit(1);
  }

  const envClaspFile = path.join(ROOT, `.clasp.${env}.json`);
  if (!fs.existsSync(envClaspFile)) {
    console.error(`Environment file not found: .clasp.${env}.json`);
    console.error(`Copy from .clasp.${env}.json.template and fill in values.`);
    process.exit(1);
  }

  if (fs.existsSync(CLASP_BACKUP)) {
    console.error('ERROR: .clasp.json.bak exists from a previous interrupted run.');
    console.error('Restore manually: mv .clasp.json.bak .clasp.json');
    process.exit(1);
  }

  // .clasp.json スワップ
  const hadOriginal = fs.existsSync(CLASP_JSON);
  if (hadOriginal) {
    fs.copyFileSync(CLASP_JSON, CLASP_BACKUP);
  }
  fs.copyFileSync(envClaspFile, CLASP_JSON);

  // .clasprc.json の default トークンをスワップ
  const envClasp = JSON.parse(fs.readFileSync(envClaspFile, 'utf8'));
  const tokenName = envClasp.token;
  const clasprcPath = findClasprc();
  const originalDefault = swapDefaultToken(clasprcPath, tokenName);

  console.log(`[env-switch] Activated: ${env}`);

  try {
    execFileSync(cmdArgs[0], cmdArgs.slice(1), { stdio: 'inherit', cwd: ROOT });
  } catch (e) {
    process.exitCode = e.status || 1;
  } finally {
    // .clasp.json 復元
    if (hadOriginal && fs.existsSync(CLASP_BACKUP)) {
      fs.copyFileSync(CLASP_BACKUP, CLASP_JSON);
      fs.unlinkSync(CLASP_BACKUP);
    } else if (fs.existsSync(CLASP_BACKUP)) {
      fs.unlinkSync(CLASP_BACKUP);
    }
    // .clasprc.json 復元
    restoreDefaultToken(clasprcPath, originalDefault);
    console.log(`[env-switch] Restored`);
  }
}

main();
