/**
 * Apps Script デプロイ一覧表示
 *
 * 「本番の deploymentId を失くした」「どの version がどこに紐づいてるか
 * 確認したい」ときに使う。scripts/config.json の prodDeployId を復旧する
 * 際の recovery tool。
 *
 * Usage:
 *   npm run deploy:list
 *   npm run deploy:list -- --env open
 */
const { requestJSON, parseEnvFromArgs, loadScriptContext } = require('./lib/gas-auth');

const { env } = parseEnvFromArgs(process.argv.slice(2));

try {
  const { config, scriptId, token } = loadScriptContext(env);
  const resp = requestJSON(
    'GET',
    `https://script.googleapis.com/v1/projects/${scriptId}/deployments?pageSize=20`,
    null,
    token
  );

  if (resp.error) throw new Error(`Deploy list API error: ${resp.error.message || resp.error}`);

  const deployments = resp.deployments || [];
  if (deployments.length === 0) {
    console.log('No deployments found.');
    process.exit(0);
  }

  console.log(`${deployments.length} deployment(s) for scriptId ${scriptId.substring(0, 20)}...\n`);
  for (const d of deployments) {
    const id = d.deploymentId || '';
    const cfg = d.deploymentConfig || {};
    const version = cfg.versionNumber !== undefined ? `v${cfg.versionNumber}` : 'HEAD';
    const desc = cfg.description || '';
    const updateTime = d.updateTime || '';
    const marker = id === config.prodDeployId ? ' ← prod' : '';
    console.log(`${id}  ${version.padEnd(6)}  ${updateTime.substring(0, 19)}  ${desc}${marker}`);
  }
} catch (error) {
  console.error('deploy:list failed:', error.message);
  process.exit(1);
}
