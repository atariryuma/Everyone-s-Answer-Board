/**
 * 同時接続シミュレーション (SA pool / board cache / lock 競合 を測定)。
 *
 * Usage:
 *   node scripts/load-test-concurrent.js [--n 30] [--userId <id>] [--op previewBoard]
 *
 * 実装: 動作実績ある admin-api.js を N 個 spawn し、 ms 単位の latency を集計。
 *       各 process は別 OAuth fetch + python urllib 経由で本番 prodUrl を叩く。
 */
const { spawn } = require('child_process');
const path = require('path');

const args = process.argv.slice(2);
function getArg(name, dflt) {
  const i = args.indexOf('--' + name);
  if (i >= 0 && args[i + 1]) return args[i + 1];
  return dflt;
}
const N = parseInt(getArg('n', '30'), 10);
const operation = getArg('op', 'previewBoard');
const userId = getArg('userId', 'adb24b94-8244-4d3a-a1c3-e409f81e40a0');
const apiScript = path.join(__dirname, 'admin-api.js');

function runOne(idx) {
  return new Promise((resolve) => {
    const t0 = Date.now();
    const child = spawn('node', [apiScript, operation, '--userId', userId], {
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
  process.exit(success === N ? 0 : 1);
})();
