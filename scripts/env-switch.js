/**
 * 環境切替スクリプト
 * .clasp.{env}.json → .clasp.json にスワップしてコマンドを実行し、終了後に復元する
 *
 * Usage: node scripts/env-switch.js <env> <command...>
 * Example: node scripts/env-switch.js open npx clasp push --force
 */
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

const ROOT = path.join(__dirname, '..');
const CLASP_JSON = path.join(ROOT, '.clasp.json');
const CLASP_BACKUP = path.join(ROOT, '.clasp.json.bak');

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

  const hadOriginal = fs.existsSync(CLASP_JSON);
  if (hadOriginal) {
    fs.copyFileSync(CLASP_JSON, CLASP_BACKUP);
  }
  fs.copyFileSync(envClaspFile, CLASP_JSON);

  console.log(`[env-switch] Activated: ${env}`);

  try {
    execFileSync(cmdArgs[0], cmdArgs.slice(1), { stdio: 'inherit', cwd: ROOT });
  } catch (e) {
    process.exitCode = e.status || 1;
  } finally {
    if (hadOriginal && fs.existsSync(CLASP_BACKUP)) {
      fs.copyFileSync(CLASP_BACKUP, CLASP_JSON);
      fs.unlinkSync(CLASP_BACKUP);
      console.log(`[env-switch] Restored original .clasp.json`);
    } else if (fs.existsSync(CLASP_BACKUP)) {
      fs.unlinkSync(CLASP_BACKUP);
    }
  }
}

main();
