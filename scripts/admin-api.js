/**
 * Admin API クライアント
 *
 * Usage:
 *   npm run api -- getAppStatus
 *   npm run api -- systemDiagnosis
 *   npm run api -- getUsers
 *   npm run api -- getLogs --limit 20
 *   npm run api -- perfMetrics --category api
 *   npm run api -- listProperties
 */
const { getAccessToken, postJSONSync, getConfig, parseEnvFromArgs } = require('./lib/gas-auth');

const OPERATIONS = [
  'getUsers', 'toggleUserActive', 'toggleUserBoard',
  'getLogs', 'disableApp', 'enableApp', 'getAppStatus',
  'systemDiagnosis', 'autoRepair', 'cacheReset',
  'perfMetrics', 'perfDiagnosis',
  'getProperty', 'setProperty', 'listProperties'
];

function parseArgs(args) {
  const operation = args[0];
  const params = {};
  for (let i = 1; i < args.length; i += 2) {
    const key = args[i]?.replace(/^--/, '');
    const value = args[i + 1];
    if (key && value !== undefined) {
      params[key] = isNaN(value) ? value : Number(value);
    }
  }
  return { operation, params };
}

const { env, remainingArgs } = parseEnvFromArgs(process.argv.slice(2));
if (!remainingArgs[0] || remainingArgs[0] === '--help') {
  console.log('Usage: npm run api -- <operation> [--env <env>] [--param value ...]');
  console.log('\nOperations:');
  OPERATIONS.forEach(op => console.log(`  ${op}`));
  process.exit(0);
}

const { operation, params } = parseArgs(remainingArgs);

try {
  const config = getConfig(env);
  const token = getAccessToken(config.tokenName);
  const payload = { action: 'adminApi', apiKey: config.apiKey, operation, params };
  const json = postJSONSync(config.prodUrl, payload, token);
  console.log(JSON.stringify(json, null, 2));
  process.exit(json.success ? 0 : 1);
} catch (error) {
  console.error('Error:', error.message);
  process.exit(1);
}
