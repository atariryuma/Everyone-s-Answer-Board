#!/usr/bin/env node
/**
 * test:coverage の素直な実装。
 *
 * Why custom (not just --experimental-test-coverage):
 *   tests/ は `vm.runInContext(fs.readFileSync(src), ctx)` で src/*.js を
 *   サンドボックスにロードする (GAS グローバルスタブを与えるため)。Node の
 *   --experimental-test-coverage は require/import で読まれたファイルのみ
 *   instrument するため、vm 経由でロードされた src/ は計測対象から漏れる。
 *   結果として 100.00% という嘘の数字が表示されていた (誤った安心感)。
 *
 *   本スクリプトは:
 *     1) node --test を素で走らせて pass 数を出す
 *     2) src/*.js / src/*.html を粗くスキャンして "tested" マーカー (tests/ 内で
 *        loadXxxContext(...).fnName(...) のように呼ばれている関数名 / load した
 *        ファイル) を集計
 *     3) "ファイル単位の有効テスト存在率" を出力
 *
 *   これは真の line/branch coverage ではないが、嘘の 100% よりは正直で、
 *   どの src/ ファイルが unit test から触られていないかをすぐに把握できる。
 *   line coverage がどうしても必要なら c8 を入れるが、プロジェクト方針
 *   (zero-dep) と相反するので採用しない。
 */
const fs = require('fs');
const path = require('path');
const { execSync, spawnSync } = require('child_process');

const ROOT = path.resolve(__dirname, '..');

console.log('▶ Running test suite...');
const testRun = spawnSync('node', ['--test', ...listGlob('tests/*.test.cjs')], {
  cwd: ROOT,
  stdio: 'inherit',
});

if (testRun.status !== 0) {
  console.error('\n✗ tests failed; not computing coverage estimate');
  process.exit(testRun.status || 1);
}

console.log('\n▶ Computing per-file test reach (heuristic)...');
const srcFiles = listGlob('src/*.js')
  .concat(listGlob('src/*.html'))
  .map((rel) => ({ rel, base: path.basename(rel) }));

const testBlob = listGlob('tests/*.test.cjs')
  .map((rel) => fs.readFileSync(path.join(ROOT, rel), 'utf8'))
  .join('\n');

let loaded = 0;
const untouched = [];
for (const { rel, base } of srcFiles) {
  // tests/ から fs.readFileSync で参照されているか、ファイル名や module suffix で
  // 言及されているか。十分粗いがファイル単位の "tested vs untested" は分かる。
  const re = new RegExp(escapeRe(base.replace(/\.html$/, '').replace(/\.js$/, '')), 'i');
  if (testBlob.includes(rel) || testBlob.includes(base) || re.test(testBlob)) {
    loaded++;
  } else {
    untouched.push(rel);
  }
}

const total = srcFiles.length;
const pct = total === 0 ? 0 : (loaded / total) * 100;

console.log('');
console.log(`File-reach (heuristic): ${loaded}/${total} = ${pct.toFixed(1)}%`);
if (untouched.length > 0) {
  console.log('\nFiles with no test reference:');
  for (const f of untouched) console.log(`  - ${f}`);
}
console.log('');
console.log('注意: これはファイル単位の "tested かどうか" の粗推定であり、line/branch');
console.log('      coverage ではない。詳細は scripts/test-coverage.js の冒頭コメント参照。');

function listGlob(pattern) {
  const [dir, glob] = pattern.split('/');
  const re = new RegExp('^' + escapeRe(glob).replace(/\\\*/g, '.*') + '$');
  return fs.readdirSync(path.join(ROOT, dir))
    .filter((f) => re.test(f))
    .map((f) => path.join(dir, f));
}

function escapeRe(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
