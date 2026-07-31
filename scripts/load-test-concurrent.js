/**
 * 同時接続シミュレーション (SA pool / board cache / lock 競合 を測定)。
 *
 * 700 人規模の同時アクセスを想定した設計 (SA pool round-robin + 30s cooldown /
 * board data cache 10s / per-row CAS lock) が実際に効いているかを本番で実測する。
 * ユニットテストでも smoke テストでも「並列時に崩れる」挙動は検出できないため、
 * 授業前など負荷が読めるタイミングで手動実行する。
 *
 * Usage:
 *   npm run loadtest                        # 那覇 / N=30 / previewBoard
 *   npm run loadtest:open                   # 沖縄
 *   node scripts/load-test-concurrent.js --n 50 --env open --max-p90 20000
 *
 * オプション:
 *   --n <N>          並列数 (既定 30)
 *   --op <operation> 叩く admin API (既定 previewBoard = viewer の重い経路)
 *   --userId <id>    対象ユーザー。省略時は getUsers から自動選択 (テナント非依存)
 *   --env <env>      open 等。省略で default (那覇)
 *   --max-p90 <ms>   p90 latency のしきい値 (既定 20000)。超えたら FAIL
 *
 * 判定: 全リクエスト成功 かつ p90 <= max-p90 で exit 0。どちらか崩れたら exit 1。
 *
 * 注意: **本番に実負荷をかける**。GAS quota / Sheets API quota を消費するので
 *       授業中には実行しない。read-only 操作のみだが N に比例して quota を使う。
 *
 * 実装: 動作実績ある admin-api.js を N 個 spawn し、 ms 単位の latency を集計。
 *       各 process は別 OAuth fetch + python urllib 経由で本番 prodUrl を叩く。
 */
const { spawn, spawnSync } = require('child_process');
const path = require('path');

const args = process.argv.slice(2);
function getArg(name, dflt) {
  const i = args.indexOf('--' + name);
  if (i >= 0 && args[i + 1]) return args[i + 1];
  return dflt;
}
const N = parseInt(getArg('n', '30'), 10);
const operation = getArg('op', 'previewBoard');
const ENV = getArg('env', null);
const MAX_P90 = parseInt(getArg('max-p90', '20000'), 10);
const apiScript = path.join(__dirname, 'admin-api.js');

// --env を admin-api.js に引き継ぐ共通引数。
const envArgs = ENV ? ['--env', ENV] : [];

// userId 未指定なら getUsers から「データソース接続済みの active ユーザー」を選ぶ。
// ハードコードした UUID はテナントを跨げず、そのユーザーが消えると黙って壊れるため。
function resolveUserId() {
  const explicit = getArg('userId', null);
  if (explicit) return explicit;
  const r = spawnSync('node', [apiScript, 'getUsers', ...envArgs], { encoding: 'utf8', timeout: 120000 });
  const out = String(r.stdout || '');
  const start = out.indexOf('{');
  if (start < 0) throw new Error('getUsers が JSON を返さなかった (認証 / config.json を確認)');
  const json = JSON.parse(out.slice(start));
  const list = json.users || json.data || [];
  const arr = Array.isArray(list) ? list : [];
  const hit = arr.find((u) => {
    if (!u || u.isActive === false) return false;
    try {
      const cfg = typeof u.configJson === 'object' ? u.configJson : JSON.parse(u.configJson || '{}');
      return Boolean(cfg && cfg.spreadsheetId);
    } catch (_) { return false; }
  });
  if (!hit) throw new Error('データソース接続済みユーザーが見つからない (--userId で明示指定してください)');
  return hit.userId;
}

const userId = resolveUserId();

function runOne(idx) {
  return new Promise((resolve) => {
    const t0 = Date.now();
    const child = spawn('node', [apiScript, operation, '--userId', userId, ...envArgs], {
      env: process.env,
      stdio: ['ignore', 'pipe', 'pipe']
    });
    let stdout = '';
    let stderr = '';
    child.stdout.on('data', (d) => { stdout += d.toString('utf8'); });
    child.stderr.on('data', (d) => { stderr += d.toString('utf8'); });
    child.on('close', (code) => {
      const elapsed = Date.now() - t0;
      let json = null;
      try { json = JSON.parse(stdout); } catch (_) {}
      resolve({
        idx, elapsed, code, json,
        success: code === 0 && json && json.success,
        bodySnippet: stdout.substring(0, 200),
        stderr: stderr.substring(0, 200)
      });
    });
  });
}

(async () => {
  // ウォームアップ: 計測前に 1 回だけ捨てリクエストを投げる。
  //   GAS はアイドル後の初回で spin-up が入り (実測 22s vs ウォーム 4.6s)、
  //   そのまま測ると初回実行が常に false FAIL になりツールが信用されなくなる。
  //   --no-warmup でコールドスタート自体を測ることもできる。
  if (!args.includes('--no-warmup')) {
    process.stdout.write('Warm-up (1 req, 計測から除外)... ');
    const w = await runOne(-1);
    console.log(`${w.elapsed}ms ${w.success ? 'OK' : '(失敗 — 本計測で再判定)'}`);
  }

  console.log(`Spawning ${N} concurrent admin-api ${operation} (userId=${userId})...`);
  const tStart = Date.now();
  const results = await Promise.all(
    Array.from({ length: N }, (_, i) => runOne(i))
  );
  const tEnd = Date.now();

  const latencies = results.map((r) => r.elapsed).sort((a, b) => a - b);
  const success = results.filter((r) => r.success).length;
  const failed = results.filter((r) => !r.success);
  const sample = failed[0];

  const pct = (p) => latencies[Math.min(latencies.length - 1, Math.floor(latencies.length * p))];

  console.log('');
  console.log('=== Latency (ms) ===');
  console.log(`  min:  ${latencies[0]}`);
  console.log(`  p50:  ${pct(0.50)}`);
  console.log(`  p90:  ${pct(0.90)}`);
  console.log(`  p99:  ${pct(0.99)}`);
  console.log(`  max:  ${latencies[latencies.length - 1]}`);
  console.log(`  wall: ${tEnd - tStart}ms  (= ${N} req in this window)`);
  console.log(`  cumulative: ${latencies.reduce((a, b) => a + b, 0)}ms`);
  console.log(`  parallelism gain: ${(latencies.reduce((a, b) => a + b, 0) / (tEnd - tStart)).toFixed(1)}x`);
  console.log('');
  console.log('=== Results ===');
  console.log(`  success: ${success}/${N}`);
  console.log(`  failed:  ${failed.length}`);
  if (sample) {
    console.log(`  sample fail: code=${sample.code} body=${sample.bodySnippet} stderr=${sample.stderr}`);
  }

  // 合否判定: 全件成功 かつ p90 がしきい値以内。
  //   成功率だけ見ていると「全部返ったが 1 件 60 秒」を見逃す (授業では実質失敗)。
  const p90 = pct(0.90);
  const allOk = success === N;
  const latencyOk = p90 <= MAX_P90;
  console.log('');
  console.log('=== Verdict ===');
  console.log(`  ${allOk ? '✅' : '❌'} 成功率: ${success}/${N}`);
  console.log(`  ${latencyOk ? '✅' : '❌'} p90 latency: ${p90}ms (しきい値 ${MAX_P90}ms)`);
  if (!allOk || !latencyOk) {
    console.log('');
    console.log('  調査: npm run logs:cloud -- --severity ERROR --hours 1');
    console.log('        npm run api -- getServiceAccountUsage   # SA pool の cooldown 状況');
  }
  process.exit(allOk && latencyOk ? 0 : 1);
})();
