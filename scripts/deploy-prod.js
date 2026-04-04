/**
 * 本番デプロイスクリプト
 * 既存のデプロイURLを維持したまま最新コードに更新する
 *
 * Usage: npm run deploy:prod
 */
const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');
const { getAccessToken, getConfig, parseEnvFromArgs } = require('./lib/gas-auth');

function run(cmd) {
  return execSync(cmd, { encoding: 'utf8' }).trim();
}

const { env } = parseEnvFromArgs(process.argv.slice(2));

try {
  const config = getConfig(env);
  const claspJson = JSON.parse(fs.readFileSync(path.join(process.cwd(), '.clasp.json'), 'utf8'));
  const scriptId = claspJson.scriptId;

  // 1. Push latest code
  console.log('Pushing code to GAS...');
  run('npx clasp push --force');

  // 2. Create new version
  console.log('Creating new version...');
  const token = getAccessToken(config.tokenName);

  const versionRaw = execSync(
    `curl -s -X POST "https://script.googleapis.com/v1/projects/${scriptId}/versions" ` +
    `-H "Content-Type: application/json" ` +
    `-H "Authorization: Bearer ${token}" ` +
    `-d '{"description":"deploy:prod"}'`,
    { encoding: 'utf8' }
  );
  let versionData;
  try { versionData = JSON.parse(versionRaw); } catch (e) {
    throw new Error(`Failed to parse version response: ${versionRaw.substring(0, 200)}`);
  }
  const version = versionData.versionNumber;
  if (!version) throw new Error(`Failed to create version: ${versionData.error?.message || 'unknown'}`);
  console.log(`Created version: ${version}`);

  // 3. Update existing deployment (URL stays the same)
  const deployId = config.prodDeployId;
  console.log(`Updating production deployment ${deployId.substring(0, 20)}...`);

  const deployBody = JSON.stringify({ deploymentConfig: { scriptId, versionNumber: version } });
  const deployPayload = path.join(require('os').tmpdir(), `gas-deploy-${Date.now()}.json`);
  fs.writeFileSync(deployPayload, deployBody, { mode: 0o600 });
  let deployRaw;
  try {
    deployRaw = execSync(
      `curl -s -X PUT "https://script.googleapis.com/v1/projects/${scriptId}/deployments/${deployId}" ` +
      `-H "Content-Type: application/json" ` +
      `-H "Authorization: Bearer ${token}" ` +
      `-d @"${deployPayload}"`,
      { encoding: 'utf8' }
    );
  } finally {
    try { fs.unlinkSync(deployPayload); } catch (_) {}
  }
  let deployData;
  try { deployData = JSON.parse(deployRaw); } catch (e) {
    throw new Error(`Failed to parse deploy response: ${deployRaw.substring(0, 200)}`);
  }
  if (deployData.error) throw new Error(`Deploy API error: ${deployData.error.message}`);

  console.log(`\nDeploy complete! v${version}`);
  console.log(`URL: ${config.prodUrl}`);
} catch (error) {
  console.error('Deploy failed:', error.message);
  process.exit(1);
}
