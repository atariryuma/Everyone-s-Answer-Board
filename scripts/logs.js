/**
 * GASアプリのCloud Loggingログを取得
 *
 * Usage:
 *   npm run logs:cloud                            # 直近6hのWARNING以上
 *   npm run logs:cloud -- --severity ERROR         # ERRORのみ
 *   npm run logs:cloud -- --severity INFO          # INFO以上
 *   npm run logs:cloud -- --hours 24 --limit 50    # 過去24時間、50件
 *   npm run logs:cloud -- --function doPost        # 特定関数のログ
 *   npm run logs:cloud -- --user 35t22             # 特定ユーザーのログ
 *   npm run logs:cloud -- --json                   # JSON形式で出力
 */
const { getAccessToken, postJSONSync, getConfig, parseEnvFromArgs } = require('./lib/gas-auth');

function parseArgs(args) {
  const opts = { severity: 'WARNING', limit: 20, hours: 6, user: null, func: null, json: false };
  for (let i = 0; i < args.length; i++) {
    switch (args[i]) {
      case '--severity': opts.severity = args[++i]; break;
      case '--limit': opts.limit = parseInt(args[++i], 10); break;
      case '--hours': opts.hours = parseInt(args[++i], 10); break;
      case '--user': opts.user = args[++i]; break;
      case '--function': opts.func = args[++i]; break;
      case '--json': opts.json = true; break;
    }
  }
  return opts;
}

const { env, remainingArgs } = parseEnvFromArgs(process.argv.slice(2));
const opts = parseArgs(remainingArgs);

try {
  const config = getConfig(env);
  const token = getAccessToken(config.tokenName);

  const cutoff = new Date(Date.now() - opts.hours * 3600000).toISOString();
  let filter = `resource.type="app_script_function" severity>=${opts.severity} timestamp>="${cutoff}"`;
  if (opts.user) filter += ` jsonPayload.message=~"${opts.user.replace(/"/g, '\\"')}"`;
  if (opts.func) filter += ` resource.labels.function_name="${opts.func.replace(/"/g, '\\"')}"`;

  const resp = postJSONSync('https://logging.googleapis.com/v2/entries:list', {
    resourceNames: [`projects/${config.gcpProjectId}`],
    filter,
    orderBy: 'timestamp desc',
    pageSize: opts.limit
  }, token);

  if (resp.error) {
    throw new Error(`Cloud Logging API error: ${resp.error.message}`);
  }

  const entries = resp.entries || [];

  if (opts.json) {
    console.log(JSON.stringify(entries, null, 2));
    process.exit(0);
  }

  if (entries.length === 0) {
    console.log(`No logs found (severity>=${opts.severity}, past ${opts.hours}h)`);
    process.exit(0);
  }

  console.log(`${entries.length} log entries (severity>=${opts.severity}, past ${opts.hours}h)\n`);
  for (const e of entries) {
    const ts = (e.timestamp || '').substring(0, 19).replace('T', ' ');
    const sev = (e.severity || '').padEnd(8);
    const func = e.resource?.labels?.function_name || 'unknown';
    const msg = e.jsonPayload?.message || '';
    console.log(`${ts}  ${sev}  ${func.padEnd(12)}  ${msg}`);
  }
} catch (error) {
  console.error('Error:', error.message);
  process.exit(1);
}
