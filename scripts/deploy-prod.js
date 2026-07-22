/**
 * 本番デプロイスクリプト
 * 既存のデプロイURLを維持したまま最新コードに更新する。
 *
 *   1. preflight: git 作業ツリーのクリーンチェック
 *   2. clasp push: ソースを HEAD に上げる
 *   3. Apps Script REST API で新しい Version を作成
 *   4. 既存 Deployment を新 Version に付け替え（URL/ID を維持）
 *
 * Usage:
 *   npm run deploy:prod                # 本番
 *   npm run deploy:prod -- --env open  # 別環境
 *   npm run deploy:prod -- --force     # 未コミット変更を許容してデプロイ
 */
const { execSync, execFileSync } = require('child_process');
const { parseEnvFromArgs, requestJSON, loadScriptContext } = require('./lib/gas-auth');

const { env, remainingArgs } = parseEnvFromArgs(process.argv.slice(2));
const force = remainingArgs.includes('--force');

// デプロイ成功後に HEAD コミットの vNNNN で git tag を打つ (best-effort)。
// Why: GAS の versionNumber (2870 等) と git コミットの vNNNN (2868 等) は別カウンタで、
//   2 テナント (naha/open) は同一コミットを別 GAS version として deploy する。tag は
//   「どのコミットが本番に出たか」を git 側で辿るためのものなので、GAS version ではなく
//   HEAD コミットの vNNNN を採用する。両テナントが同じコミットを打っても idempotent。
// 失敗してもデプロイ自体は成功済みなので、ここでは warn だけして exit code を変えない。
function tagDeployedVersionBestEffort() {
  try {
    const subject = execFileSync('git', ['log', '-1', '--format=%s'], { encoding: 'utf8' }).trim();
    const m = subject.match(/\(v(\d+)\)\s*$/);
    if (!m) return;
    const tag = `v${m[1]}`;
    const existing = execFileSync('git', ['tag', '-l', tag], { encoding: 'utf8' }).trim();
    if (existing) return; // 既存 tag は触らない
    execFileSync('git', ['tag', tag], { encoding: 'utf8' });
    console.log(`Tagged HEAD as ${tag} (push with: git push origin ${tag})`);
  } catch (e) {
    console.warn(`(tag skipped: ${e.message})`);
  }
}

function preflightGitClean() {
  let status;
  try {
    status = execFileSync('git', ['status', '--porcelain'], { encoding: 'utf8' }).trim();
  } catch (e) {
    // Not a git repo or git unavailable — skip preflight rather than blocking.
    return;
  }
  if (!status) return;

  if (force) {
    console.warn('⚠️  Uncommitted changes detected; deploying anyway (--force):');
    console.warn(status.split('\n').map((l) => `   ${l}`).join('\n'));
    return;
  }

  console.error('❌ Uncommitted changes detected. Commit or stash before deploying,');
  console.error('   or re-run with --force to bypass:');
  console.error(status.split('\n').map((l) => `   ${l}`).join('\n'));
  console.error('');
  console.error('   Example: npm run deploy:prod -- --force');
  process.exit(1);
}

try {
  preflightGitClean();

  const { config, scriptId, token } = loadScriptContext(env);

  // 1. Push latest code
  // clasp はグローバル default ユーザーで push するため、別テナントの作業で default が
  // 切り替わると 403 (The caller does not have permission) になる。REST 呼び出しと同じ
  // tokenName の named credential (~/.clasprc.json tokens.<name>) を明示して環境非依存にする。
  const claspUser = config.claspUser || config.tokenName || 'default';
  console.log(`Pushing code to GAS (clasp user: ${claspUser})...`);
  execSync(`npx clasp --user ${claspUser} push --force`, { stdio: 'inherit' });

  // 2. Create new version
  console.log('Creating new version...');
  const versionData = requestJSON(
    'POST',
    `https://script.googleapis.com/v1/projects/${scriptId}/versions`,
    { description: 'deploy:prod' },
    token
  );
  const version = versionData.versionNumber;
  if (!version) throw new Error(`Failed to create version: ${versionData.error?.message || versionData.error || 'unknown'}`);
  console.log(`Created version: ${version}`);

  // 3. Update existing deployment (URL stays the same)
  const deployId = config.prodDeployId;
  console.log(`Updating production deployment ${deployId.substring(0, 20)}...`);
  const deployData = requestJSON(
    'PUT',
    `https://script.googleapis.com/v1/projects/${scriptId}/deployments/${deployId}`,
    { deploymentConfig: { scriptId, versionNumber: version } },
    token
  );
  if (deployData.error) throw new Error(`Deploy API error: ${deployData.error.message || deployData.error}`);

  console.log(`\nDeploy complete! v${version}`);
  console.log(`URL: ${config.prodUrl}`);

  // 本番反映に成功したので、HEAD コミットの vNNNN を git tag で辿れるようにする (best-effort)。
  tagDeployedVersionBestEffort();
} catch (error) {
  console.error('Deploy failed:', error.message);
  process.exit(1);
}
