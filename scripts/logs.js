/**
 * GASアプリのCloud Loggingログを取得・集計
 *
 * Quick usage (よく使う順):
 *   npm run logs:cloud                            # 直近6hのWARNING以上
 *   npm run logs:cloud -- --summary               # エラーをシグネチャ別に集計（★推奨★）
 *   npm run logs:cloud -- --tail                  # 直近10分の ERROR のみ（デプロイ後の確認に）
 *   npm run logs:cloud -- --severity ERROR --hours 3
 *
 * Filters:
 *   --severity DEFAULT|DEBUG|INFO|WARNING|ERROR|CRITICAL   (default: WARNING)
 *   --hours <N>                                            (default: 6)
 *   --limit <N>                                            (default: 20)
 *   --function <name>                                      (e.g. doPost, getPublishedSheetData)
 *   --user <frag>                                          (メールの一部で絞り込み)
 *
 * Output modes:
 *   (default)      1行/エントリの読みやすい形式
 *   --brief        1行/エントリ（ts + severity + 先頭100文字）。コピペ/LLM 向け
 *   --summary      メッセージシグネチャで集計、件数多い順
 *   --tail         --severity ERROR --hours 0.17 --brief と同じ（直近10分）
 *   --json         生の JSON
 */
const { getAccessToken, postJSONSync, getConfig, parseEnvFromArgs } = require('./lib/gas-auth');

function parseArgs(args) {
  const opts = {
    severity: 'WARNING',
    limit: 20,
    hours: 6,
    user: null,
    func: null,
    json: false,
    brief: false,
    summary: false,
    tail: false,
  };
  for (let i = 0; i < args.length; i++) {
    switch (args[i]) {
      case '--severity': opts.severity = args[++i]; break;
      case '--limit': opts.limit = parseInt(args[++i], 10); break;
      case '--hours': opts.hours = parseFloat(args[++i]); break;
      case '--user': opts.user = args[++i]; break;
      case '--function': opts.func = args[++i]; break;
      case '--json': opts.json = true; break;
      case '--brief': opts.brief = true; break;
      case '--summary': opts.summary = true; break;
      case '--tail':
        opts.tail = true;
        opts.severity = 'ERROR';
        opts.hours = 1 / 6; // 10分
        opts.brief = true;
        opts.limit = 30;
        break;
    }
  }
  return opts;
}

/**
 * メッセージから一意のシグネチャを抽出してグループ化キーに使う。
 * 具体値（タイムスタンプ・ID・ユーザー名等）を正規化してまとめる。
 */
function signatureOf(message) {
  return String(message || '')
    .replace(/\b[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\b/gi, '<uuid>')
    .replace(/\b[A-Za-z0-9_-]{30,60}\b/g, '<id>') // スプレッドシートID等
    .replace(/\b\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}[0-9.Z+:-]*/g, '<ts>')
    .replace(/\b\d+ms\b/g, '<N>ms')
    .replace(/\b\d+\b/g, '<n>')
    .replace(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g, '<email>')
    .replace(/\s+/g, ' ')
    .trim()
    .substring(0, 200);
}

function fetchEntries(opts) {
  const { env, remainingArgs: _remainingArgs } = parseEnvFromArgs(process.argv.slice(2));
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

  if (resp.error) throw new Error(`Cloud Logging API error: ${resp.error.message}`);
  return resp.entries || [];
}

function renderBrief(entries) {
  for (const e of entries) {
    const ts = (e.timestamp || '').substring(11, 19); // HH:MM:SS
    const sev = (e.severity || '').charAt(0);        // W/E/I
    const func = (e.resource?.labels?.function_name || '?').substring(0, 10);
    const msg = String(e.jsonPayload?.message || '').replace(/\s+/g, ' ').substring(0, 120);
    console.log(`${ts} ${sev} ${func.padEnd(10)} ${msg}`);
  }
}

function renderDefault(entries, opts) {
  console.log(`${entries.length} log entries (severity>=${opts.severity}, past ${opts.hours}h)\n`);
  for (const e of entries) {
    const ts = (e.timestamp || '').substring(0, 19).replace('T', ' ');
    const sev = (e.severity || '').padEnd(8);
    const func = (e.resource?.labels?.function_name || 'unknown').padEnd(12);
    const msg = e.jsonPayload?.message || '';
    console.log(`${ts}  ${sev}  ${func}  ${msg}`);
  }
}

function renderSummary(entries, opts) {
  const groups = new Map();
  for (const e of entries) {
    const key = `${e.severity}|${e.resource?.labels?.function_name || '?'}|${signatureOf(e.jsonPayload?.message)}`;
    const g = groups.get(key) || {
      severity: e.severity,
      func: e.resource?.labels?.function_name || '?',
      signature: signatureOf(e.jsonPayload?.message),
      count: 0,
      firstTs: e.timestamp,
      lastTs: e.timestamp,
      sample: e.jsonPayload?.message || ''
    };
    g.count++;
    if (e.timestamp < g.firstTs) g.firstTs = e.timestamp;
    if (e.timestamp > g.lastTs) g.lastTs = e.timestamp;
    groups.set(key, g);
  }

  const sorted = [...groups.values()].sort((a, b) => b.count - a.count);
  console.log(`Summary: ${entries.length} entries in ${groups.size} groups (severity>=${opts.severity}, past ${opts.hours}h)\n`);
  console.log('count  sev   function       first -> last                     signature');
  console.log('-----  ----  -------------  -------------------------------   ---------');
  for (const g of sorted) {
    const first = g.firstTs.substring(11, 19);
    const last = g.lastTs.substring(11, 19);
    const sev = (g.severity || '').padEnd(4);
    const func = (g.func || '').substring(0, 13).padEnd(13);
    const sig = g.signature.substring(0, 80);
    console.log(`${String(g.count).padStart(5)}  ${sev}  ${func}  ${first} -> ${last}                 ${sig}`);
  }
}

try {
  const { remainingArgs } = parseEnvFromArgs(process.argv.slice(2));
  const opts = parseArgs(remainingArgs);
  const entries = fetchEntries(opts);

  if (opts.json) {
    console.log(JSON.stringify(entries, null, 2));
    process.exit(0);
  }

  if (entries.length === 0) {
    const label = opts.tail ? 'tail (past 10 min, ERROR)' : `severity>=${opts.severity}, past ${opts.hours}h`;
    console.log(`No logs found (${label})`);
    process.exit(0);
  }

  if (opts.summary) return renderSummary(entries, opts);
  if (opts.brief || opts.tail) return renderBrief(entries);
  return renderDefault(entries, opts);

} catch (error) {
  console.error('Error:', error.message);
  process.exit(1);
}
