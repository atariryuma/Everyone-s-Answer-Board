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
const { getAccessToken, postJSONSync, getConfig } = require('./lib/gas-auth');

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

const args = process.argv.slice(2);
if (!args[0] || args[0] === '--help') {
  console.log('Usage: npm run api -- <operation> [--param value ...]');
  console.log('\nOperations:');
  OPERATIONS.forEach(op => console.log(`  ${op}`));
  process.exit(0);
}

const { operation, params } = parseArgs(args);

try {
  const config = getConfig();
  const token = getAccessToken();
  const payload = { action: 'adminApi', apiKey: config.apiKey, operation, params };
  const json = postJSONSync(config.prodUrl, payload, token);
  console.log(JSON.stringify(json, null, 2));
  process.exit(json.success ? 0 : 1);
} catch (error) {
  console.error('Error:', error.message);
  process.exit(1);
}
