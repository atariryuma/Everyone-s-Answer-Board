/**
 * Sheets API の利用量を Cloud Monitoring から取り、上限に対する余裕を出す。
 *
 * Why: 429 (Read requests per minute per user) が出てから原因を追うのではなく、授業のたびに
 *   「1 分あたりの最大読み取り回数が上限のどれくらいか」を数字で見るため。Google は上限超過を
 *   2026 年後半から課金対象にする予定と明記している
 *   (https://developers.google.com/workspace/sheets/api/limits)。
 *
 * 上限 (同ページ): 読み取り・書き込みとも 300 回/分/プロジェクト、60 回/分/ユーザー。
 *   サービスアカウントは 1 アカウント = 1 ユーザーとして数えられる。
 *   ここでは credential (サービスアカウント等) ごとの 1 分最大を「ユーザー」の値として比べる。
 *
 * 使い方:
 *   npm run quota:report                 # 直近 7 日
 *   npm run quota:report -- --days 30
 *   npm run quota:report -- --day 2026-10-22     # その日 (JST) の 1 分ごとの上位
 *
 * 注意: consumed_api は「プロジェクトの全クライアント」の数。同じ GCP project に同居している
 *   別アプリの分も含まれる (credential が別なので、credential ごとの値で見分けられる)。
 */
const { requestJSON, loadScriptContext, parseEnvFromArgs } = require('./lib/gas-auth');

const LIMIT_PER_PROJECT = 300;
const LIMIT_PER_USER = 60;
const JST_MS = 9 * 3600 * 1000;

function parseArgs(argv) {
  const o = { days: 7, day: null, project: null };
  for (let i = 0; i < argv.length; i++) {
    if (argv[i] === '--days') o.days = Math.max(1, parseInt(argv[++i], 10) || 7);
    else if (argv[i] === '--day') o.day = argv[++i];
    else if (argv[i] === '--project') o.project = argv[++i];
  }
  return o;
}

// 一時的な通信断 (接続が切れた / 5xx / 429) は、指数バックオフ + Full Jitter で 4 回まで投げ直す。
//   読み取り専用の GET なので、投げ直しても副作用は無い。
function requestWithRetry(url, token) {
  const MAX_ATTEMPTS = 4;
  let lastError = null;
  for (let attempt = 0; attempt < MAX_ATTEMPTS; attempt++) {
    try {
      const r = requestJSON('GET', url, null, token);
      const code = r && r.error && Number(r.error.code);
      if (code && (code === 429 || code >= 500) && attempt < MAX_ATTEMPTS - 1) {
        lastError = new Error(r.error.message || String(code));
      } else {
        return r;
      }
    } catch (e) {
      lastError = e;
      if (attempt === MAX_ATTEMPTS - 1) break;
    }
    const waitMs = Math.round(Math.random() * Math.min(16000, 1000 * Math.pow(2, attempt)));
    Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, waitMs);
  }
  throw lastError || new Error('request failed');
}

function fetchSeries(token, project, startIso, endIso, alignSec, groupBy) {
  const filter = encodeURIComponent(
    'metric.type="serviceruntime.googleapis.com/api/request_count" AND resource.type="consumed_api" ' +
    'AND resource.labels.service="sheets.googleapis.com"');
  let url = `https://monitoring.googleapis.com/v3/projects/${project}/timeSeries?filter=${filter}` +
    `&interval.startTime=${startIso}&interval.endTime=${endIso}` +
    `&aggregation.alignmentPeriod=${alignSec}s&aggregation.perSeriesAligner=ALIGN_DELTA` +
    '&aggregation.crossSeriesReducer=REDUCE_SUM';
  groupBy.forEach((g) => { url += `&aggregation.groupByFields=${g}`; });
  const out = [];
  let pageToken = '';
  do {
    const r = requestWithRetry(url + (pageToken ? `&pageToken=${pageToken}` : ''), token);
    if (r.error) throw new Error(`Monitoring API: ${r.error.message || JSON.stringify(r.error)}`);
    (r.timeSeries || []).forEach((s) => out.push(s));
    pageToken = r.nextPageToken || '';
  } while (pageToken);
  return out;
}

// 1 分ごとの点を、日 (JST) → { total, c429, perMin: {minute: n}, perCred: {minute|cred: n} } に畳む。
function foldByDay(series) {
  const days = {};
  for (const s of series) {
    const code = s.metric.labels.response_code || '';
    const cred = s.resource.labels.credential_id || '-';
    for (const p of s.points) {
      const v = Number(p.value.int64Value || 0);
      if (!v) continue;
      const jst = new Date(new Date(p.interval.startTime).getTime() + JST_MS).toISOString();
      const day = jst.slice(0, 10);
      const minute = jst.slice(11, 16);
      const d = days[day] || (days[day] = { total: 0, c429: 0, perMin: {}, perCred: {}, minutes429: {} });
      d.total += v;
      d.perMin[minute] = (d.perMin[minute] || 0) + v;
      d.perCred[minute + '|' + cred] = (d.perCred[minute + '|' + cred] || 0) + v;
      if (code === '429') { d.c429 += v; d.minutes429[minute] = (d.minutes429[minute] || 0) + v; }
    }
  }
  return days;
}

function maxEntry(obj) {
  let best = ['-', 0];
  Object.entries(obj).forEach((e) => { if (e[1] > best[1]) best = e; });
  return best;
}

function pct(v, limit) { return Math.round((v / limit) * 100) + '%'; }

function main() {
  const args = parseArgs(process.argv.slice(2));
  const { token, config } = loadScriptContext(parseEnvFromArgs(process.argv.slice(2)).env);
  const project = args.project || config.gcpProjectId;
  if (!project) throw new Error('gcpProjectId が scripts/config.json にありません (--project で指定可)');

  let start; let end;
  if (args.day) {
    start = new Date(new Date(args.day + 'T00:00:00Z').getTime() - JST_MS);
    end = new Date(start.getTime() + 24 * 3600 * 1000);
  } else {
    end = new Date();
    start = new Date(end.getTime() - args.days * 24 * 3600 * 1000);
  }
  // 1 日ずつ取る。1 分粒度で数十日を 1 回で取ると応答が大きすぎて読み切れない。
  const series = [];
  const DAY_MS = 24 * 3600 * 1000;
  for (let t = start.getTime(); t < end.getTime(); t += DAY_MS) {
    const chunkEnd = new Date(Math.min(t + DAY_MS, end.getTime()));
    fetchSeries(token, project, new Date(t).toISOString(), chunkEnd.toISOString(), 60,
      ['metric.labels.response_code', 'resource.labels.credential_id']).forEach((s) => series.push(s));
  }
  const days = foldByDay(series);

  console.log(`Sheets API 利用量 (${project})  上限: ${LIMIT_PER_PROJECT}/分/プロジェクト, ${LIMIT_PER_USER}/分/ユーザー (SA は 1 つ = 1 ユーザー)`);
  console.log('日付 (JST)   合計   429   1 分最大 (対プロジェクト上限)   1 分最大/credential (対ユーザー上限)');
  Object.keys(days).sort().forEach((day) => {
    const d = days[day];
    const [pm, pv] = maxEntry(d.perMin);
    const [pk, cv] = maxEntry(d.perCred);
    console.log(`${day}  ${String(d.total).padStart(6)} ${String(d.c429).padStart(5)}   ` +
      `${String(pv).padStart(4)} @${pm} (${pct(pv, LIMIT_PER_PROJECT)})          ` +
      `${String(cv).padStart(4)} @${pk.split('|')[0]} (${pct(cv, LIMIT_PER_USER)})`);
  });

  if (args.day && days[args.day]) {
    const d = days[args.day];
    console.log(`\n${args.day} の 1 分ごとの上位 10:`);
    Object.entries(d.perMin).sort((a, b) => b[1] - a[1]).slice(0, 10)
      .forEach(([m, v]) => console.log(`  ${m}  ${v}${d.minutes429[m] ? `  (429: ${d.minutes429[m]})` : ''}`));
  }
  const any429 = Object.values(days).some((d) => d.c429 > 0);
  console.log(any429
    ? '\n429 が出た日があります。`--day <日付>` で 1 分ごとの内訳を見てください。'
    : '\n期間中の 429 は 0 件です。');
}

try {
  main();
} catch (e) {
  console.error('quota-report failed:', e.message);
  process.exit(1);
}
