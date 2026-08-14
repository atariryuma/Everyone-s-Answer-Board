#!/usr/bin/env node
/**
 * theme-verify.js — ダーク/ライト切替の完成度を end-to-end で機械的に検証。
 *
 * チェック項目 (12 軸):
 *   1.  Token Coverage          (--theme-* token が dark + light で全定義済)
 *   2.  CSS Hardcoded           (CSS 内ハードコード hex/rgba ゼロ)
 *   3.  Tailwind Unpaired       (gray-family class が theme token 化済み)
 *   4.  Inline Style Hex        (HTML inline style に hex ハードコードなし)
 *   5.  WCAG AA (dark mode)     (主要 11 ペア全部 ≥4.5/3.0:1)
 *   6.  WCAG AA (light mode)    (同上)
 *   7.  utility CSS 静的同梱    (UtilityStyles.css.html に同梱 / 外部読込ゼロ)
 *   8.  themeManager API        (mount + apply + persist 完備)
 *   8b. class 同期              (themeManager の付与 class と CSS セレクタが一致)
 *   8c. brand alias 不在        (--theme-* 直参照に統一)
 *   9.  UI mount 点             (mode=view ヘッダー + mode=admin 設定)
 *   10. Unit tests              (themeManager + 全テスト PASS)
 *
 * スコア: 各 10 点満点 → 合計 120 点。 0 違反かつ 120/120 で「完成」。
 *
 * 注: 7 は v2901 で「Tailwind darkMode 設定」から差し替わった。Tailwind CDN と
 *   SharedTailwindConfig.html は廃止済みで、dark/light の判定は themeManager が付ける
 *   class と UnifiedStyles.css の body.theme-light セレクタの整合 (8b) で見る。
 *
 * 使い方:
 *   npm run theme:verify             # ヒューマンリーダブル出力
 *   npm run theme:verify -- --json   # JSON (CI 用)
 */

'use strict';

const fs = require('node:fs');
const path = require('node:path');
const { spawnSync } = require('node:child_process');
const { checkBrandAliasResidue } = require('./lib/brand-alias-check');

const ROOT = path.resolve(__dirname, '..');
const SRC_DIR = path.join(ROOT, 'src');
const ARGS = process.argv.slice(2);
const FLAG_JSON = ARGS.includes('--json');

const checks = [];

function record(name, weight, ok, details) {
  checks.push({ name, weight, ok, score: ok ? weight : 0, details });
}

// =====================================================================
// 1-4. theme-audit を呼び出して結果を取り込み
// =====================================================================
function runAuditSub() {
  const result = spawnSync('node', ['scripts/theme-audit.js', '--json'], {
    cwd: ROOT, encoding: 'utf8'
  });
  if (result.status === null) {
    return { error: 'audit script failed to run' };
  }
  try {
    return JSON.parse(result.stdout);
  } catch (e) {
    return { error: 'audit JSON parse failed: ' + e.message, stdout: result.stdout.slice(0, 200) };
  }
}

const audit = runAuditSub();
if (audit.error) {
  console.error('theme-audit 実行失敗:', audit.error);
  process.exit(2);
}

const auditByName = {};
for (const c of audit.checks) auditByName[c.name] = c;

record(
  '1. Token Coverage',
  10,
  auditByName['Token Coverage'].violations.length === 0,
  auditByName['Token Coverage'].summary
);
record(
  '2. CSS Hardcoded Colors',
  10,
  auditByName['Hardcoded CSS Colors'].violations.length === 0,
  auditByName['Hardcoded CSS Colors'].summary
);
record(
  '3. Tailwind Unpaired',
  10,
  auditByName['Tailwind Hardcoded Colors'].violations.length === 0,
  auditByName['Tailwind Hardcoded Colors'].summary
);
record(
  '4. Inline Style Hex',
  10,
  auditByName['Inline Style Colors'].violations.length === 0,
  auditByName['Inline Style Colors'].summary
);

// =====================================================================
// 5-6. WCAG コントラスト
// =====================================================================
function runContrastSub() {
  const result = spawnSync('node', ['scripts/theme-contrast.js', '--json'], {
    cwd: ROOT, encoding: 'utf8'
  });
  try {
    return JSON.parse(result.stdout);
  } catch (e) {
    return { error: 'contrast JSON parse failed' };
  }
}

const contrast = runContrastSub();
if (contrast.error) {
  record('5. WCAG AA (dark mode)', 10, false, contrast.error);
  record('6. WCAG AA (light mode)', 10, false, contrast.error);
} else {
  const darkFail = contrast.dark.filter(r => r.status === 'fail').length;
  const lightFail = contrast.light.filter(r => r.status === 'fail').length;
  const darkPass = contrast.dark.filter(r => r.status === 'pass').length;
  const lightPass = contrast.light.filter(r => r.status === 'pass').length;
  record(
    '5. WCAG AA (dark mode)',
    10,
    darkFail === 0,
    `${darkPass}/${darkPass + darkFail} pass`
  );
  record(
    '6. WCAG AA (light mode)',
    10,
    lightFail === 0,
    `${lightPass}/${lightPass + lightFail} pass`
  );
}

// =====================================================================
// 7. utility CSS が静的同梱されているか (CDN 非依存)
//    v2901: cdn.tailwindcss.com を廃止し、使用クラスだけを
//    UtilityStyles.css.html に静的生成する方式へ移行。
//    light/dark は --theme-* token + body.theme-light で表現するため
//    Tailwind の darkMode 設定は不要になった。
// =====================================================================
function checkUtilityCss() {
  const cssPath = path.join(SRC_DIR, 'UtilityStyles.css.html');
  if (!fs.existsSync(cssPath)) return { hasCss: false, hasThemeColors: false, cdnFree: false, ok: false };
  const cfg = fs.readFileSync(cssPath, 'utf8');
  const hasThemeColors = /var\(--theme-bg-base\)/.test(cfg) && /var\(--theme-text-primary\)/.test(cfg);
  // どのページも外部ホストから script/link を読んでいないこと
  let cdnFree = true;
  for (const f of fs.readdirSync(SRC_DIR)) {
    if (!f.endsWith('.html')) continue;
    const t = fs.readFileSync(path.join(SRC_DIR, f), 'utf8');
    if (/<(?:script|link)[^>]+(?:src|href)="https?:\/\/(?!script\.google\.com)[^"]+"/.test(t)) cdnFree = false;
  }
  return { hasCss: true, hasThemeColors, cdnFree, ok: hasThemeColors && cdnFree };
}

const twConfig = checkUtilityCss();
record(
  '7. utility CSS 静的同梱 (CDN 非依存)',
  10,
  twConfig.ok,
  twConfig.ok
    ? 'UtilityStyles.css.html に theme token utility を同梱 / 外部読込 0'
    : `themeColors:${twConfig.hasThemeColors ? '✓' : '✗'} cdnFree:${twConfig.cdnFree ? '✓' : '✗'}`
);

// =====================================================================
// 8. themeManager API
// =====================================================================
function checkThemeManager() {
  const su = fs.readFileSync(path.join(SRC_DIR, 'SharedUtilities.html'), 'utf8');
  const has = {
    api: /window\.themeManager\s*=/.test(su),
    get: /\bget\s*\(\)\s*\{[^}]*readStorage/.test(su),
    set: /\bset\s*\(value\)\s*\{/.test(su),
    toggle: /\btoggle\s*\(\)\s*\{/.test(su),
    subscribe: /\bsubscribe\s*\(/.test(su),
    init: /\binit\s*\(\)\s*\{[^}]*apply/.test(su),
    mountToggle: /mountToggle\s*=\s*function/.test(su),
    autoMql: /matchMedia.*prefers-color-scheme/.test(su),
    storageKey: /['"]app-theme['"]/.test(su),
  };
  const missing = Object.entries(has).filter(([, v]) => !v).map(([k]) => k);
  return { has, ok: missing.length === 0, missing };
}

// =====================================================================
// 11. themeManager と Tailwind の class 整合性
//     themeManager が body に付与する class が Tailwind の darkMode 設定と一致するか。
//     不整合だと dark: variant が永久に発動せず、 chart 白背景バグになる
//     (2026-05-19 user 報告の根本原因)。
// =====================================================================
function checkClassSyncWithTailwind() {
  const su = fs.readFileSync(path.join(SRC_DIR, 'SharedUtilities.html'), 'utf8');
  // v2901: Tailwind config は廃止。dark/light の判定は themeManager が付ける
  //   class と UnifiedStyles の body.theme-light セレクタの整合で見る。
  const tw = fs.readFileSync(path.join(SRC_DIR, 'UnifiedStyles.css.html'), 'utf8');

  // themeManager.apply が html/body に追加する class 名を抽出
  // v2810+ 統一: classList.add(resolved, 'theme-' + resolved) — resolved は 'dark'/'light' 変数
  const themeMgrAddsDark = /classList\.add\(\s*['"]dark['"]/.test(su)
    || /classList\.add\(resolved\b/.test(su);
  const themeMgrAddsLight = /classList\.add\(\s*['"]light['"]/.test(su)
    || /classList\.add\(resolved\b/.test(su);
  const themeMgrAddsThemeDark = /classList\.add\(['"]theme-['"]\s*\+\s*resolved\)|classList\.add\(['"]theme-dark['"]|classList\.add\(resolved,\s*['"]theme-['"]\s*\+\s*resolved\)/.test(su);

  // Tailwind darkMode 設定: 'class' (= .dark), ['class', '.X'], 'selector'
  // CSS 側が light を body.theme-light で表現しているなら 'class-custom:theme-light' 相当。
  const tailwindMode = /body\.theme-light/.test(tw) ? 'class-custom:theme-light' : 'none';

  // 整合性判定
  let ok = false;
  let reason = '';
  if (tailwindMode === 'class-default') {
    // Tailwind は .dark を探す → themeManager は .dark を追加すべき
    ok = themeMgrAddsDark;
    reason = ok
      ? "Tailwind darkMode:'class' + themeManager .dark class 同期 ✓"
      : "Tailwind darkMode:'class' は .dark を探すが themeManager が .dark を付与していない (dark: variant 発動せず)";
  } else if (tailwindMode.startsWith('class-custom:')) {
    const selector = tailwindMode.slice('class-custom:'.length).replace(/^\./, '');
    // themeManager は `classList.add('theme-' + resolved)` の動的形で付与するため、
    //   リテラル一致だけでなく動的合成形も許容する (v2810+ の統一形)。
    const literal = new RegExp(`classList\\.add\\((?:[^)]*,\\s*)?['"]${selector}['"]`);
    const dynamic = /classList\.add\((?:[^)]*,\s*)?['"]theme-['"]\s*\+\s*resolved/;
    ok = literal.test(su) || dynamic.test(su);
    reason = ok
      ? `Tailwind darkMode:[class, .${selector}] + themeManager 同期 ✓`
      : `Tailwind は .${selector} を探すが themeManager が付与していない`;
  } else {
    reason = 'Tailwind darkMode が未設定';
  }

  return { ok, reason, themeMgrAddsDark, themeMgrAddsLight, themeMgrAddsThemeDark, tailwindMode };
}

const tm = checkThemeManager();
record(
  '8. themeManager API',
  10,
  tm.ok,
  tm.ok ? '9 機能完備 (get/set/toggle/subscribe/init/mountToggle/autoMql/storageKey/api)' :
    `欠落: ${tm.missing.join(', ')}`
);

// 8b. Tailwind class 同期 (chart 白背景バグの根本原因チェック)
const sync = checkClassSyncWithTailwind();
record(
  '8b. Tailwind/themeManager class 同期',
  10,
  sync.ok,
  sync.reason
);

// 8c. v2849+: brand-{background,surface,text,border} alias 廃止チェック (共通ヘルパー)。
//   theme-perfect (axis 14) と同じロジックを scripts/lib/brand-alias-check で共有。
const alias = checkBrandAliasResidue({ srcDir: SRC_DIR });
record(
  '8c. brand alias 不在 (--theme-* 直参照)',
  10,
  alias.aliasFree,
  alias.aliasFree
    ? 'alias 廃止済 (--theme-* 直参照に統一, 0 件残存)'
    : `alias 残存: def=${alias.hasDef}, usage in ${alias.usageFiles.join(', ') || 'none'}`
);

// =====================================================================
// 9. UI mount 点
// =====================================================================
function checkUiMounts() {
  const page = fs.readFileSync(path.join(SRC_DIR, 'Page.html'), 'utf8');
  const pageJs = fs.readFileSync(path.join(SRC_DIR, 'page.js.html'), 'utf8');
  const adminHtml = fs.readFileSync(path.join(SRC_DIR, 'AdminPanel.html'), 'utf8');
  const adminJs = fs.readFileSync(path.join(SRC_DIR, 'AdminPanel.js.html'), 'utf8');

  const has = {
    pageHostElement: /id=["']themeToggleHost["']/.test(page),
    pageJsMounts: /themeManager\.mountToggle\(\s*this\.elements\.themeToggleHost/.test(pageJs),
    adminSelectElement: /id=["']theme-select["']/.test(adminHtml),
    adminJsInit: /function\s+initThemeSelect\s*\(/.test(adminJs),
    adminJsCalledInInit: /try\s*\{\s*initThemeSelect\(\)/.test(adminJs),
  };
  const missing = Object.entries(has).filter(([, v]) => !v).map(([k]) => k);
  return { has, ok: missing.length === 0, missing };
}

const ui = checkUiMounts();
record(
  '9. UI mount 点 (view + admin)',
  10,
  ui.ok,
  ui.ok ? 'mode=view themeToggleHost + mode=admin theme-select 両方 mount 済'
        : `欠落: ${ui.missing.join(', ')}`
);

// =====================================================================
// 10. Unit tests
// =====================================================================
function runTests() {
  // npm を経由せず node --test を直接叩き、**レポーターも明示指定**する。
  //   node:test の既定レポーターは非 TTY だと Node 20 で tap (`# pass N`)、
  //   Node 22+ で spec (`ℹ pass N`) と変わる。`ℹ` 固定の regex だったため
  //   CI (Node 20) だけ pass 行を取りこぼし本軸が false FAIL していた
  //   (ローカル Node 24 では spec が出るので再現しなかった)。
  //   reporter を固定すれば Node バージョンに依存しない。
  //   glob はシェル無しでは展開されないので明示的にファイル列挙する。
  const testDir = path.join(ROOT, 'tests');
  const testFiles = fs.readdirSync(testDir)
    .filter((f) => f.endsWith('.test.cjs'))
    .map((f) => path.join('tests', f));
  const result = spawnSync('node', ['--test', '--test-reporter=tap', ...testFiles], {
    cwd: ROOT, encoding: 'utf8'
  });
  // tap レポーターの summary (`# pass N` / `# fail N`) から抽出
  const output = (result.stdout || '') + (result.stderr || '');
  const passMatch = output.match(/^#\s*pass\s+(\d+)/m);
  const failMatch = output.match(/^#\s*fail\s+(\d+)/m);
  const pass = passMatch ? Number(passMatch[1]) : 0;
  const fail = failMatch ? Number(failMatch[1]) : -1;
  return { pass, fail, ok: fail === 0 && pass > 0 };
}

const tests = runTests();
record(
  '10. Unit tests',
  10,
  tests.ok,
  `${tests.pass} PASS / ${tests.fail} FAIL`
);

// =====================================================================
// レポート出力
// =====================================================================
const totalScore = checks.reduce((acc, c) => acc + c.score, 0);
const maxScore = checks.reduce((acc, c) => acc + c.weight, 0);
const passedChecks = checks.filter(c => c.ok).length;

if (FLAG_JSON) {
  console.log(JSON.stringify({
    score: totalScore,
    maxScore,
    percent: Math.round((totalScore / maxScore) * 1000) / 10,
    passed: passedChecks,
    total: checks.length,
    checks,
  }, null, 2));
  process.exit(totalScore === maxScore ? 0 : 1);
}

console.log('═══════════════════════════════════════════════════════════════════');
console.log('  Theme System Completeness Verification');
console.log('═══════════════════════════════════════════════════════════════════');
console.log('');

for (const c of checks) {
  const marker = c.ok ? '✅' : '❌';
  const score = `${c.score}/${c.weight}`;
  console.log(`  ${marker} ${score.padStart(5)}  ${c.name}`);
  console.log(`           ${c.details}`);
}

console.log('');
console.log('───────────────────────────────────────────────────────────────────');
const percent = ((totalScore / maxScore) * 100).toFixed(1);
console.log(`  Score: ${totalScore} / ${maxScore} (${percent}%)`);
console.log(`  Passed: ${passedChecks} / ${checks.length}`);

if (totalScore === maxScore) {
  console.log('');
  console.log('  ✅✅✅  100% 完成 — ダーク/ライト切替 production ready  ✅✅✅');
} else {
  console.log('');
  console.log(`  ⚠️  未完成: ${checks.length - passedChecks} 項目に問題あり`);
}
console.log('═══════════════════════════════════════════════════════════════════');

process.exit(totalScore === maxScore ? 0 : 1);
