#!/usr/bin/env node
/**
 * theme-verify.js — ダーク/ライト切替の完成度を end-to-end で機械的に検証。
 *
 * チェック項目 (10 軸):
 *   1. Token Coverage          (--theme-* token が dark + light で全定義済)
 *   2. CSS Hardcoded           (CSS 内ハードコード hex/rgba ゼロ)
 *   3. Tailwind Unpaired       (gray-family class が light+dark:dark でペア化)
 *   4. Inline Style Hex        (HTML inline style に hex ハードコードなし)
 *   5. WCAG AA (dark mode)     (主要 11 ペア全部 ≥4.5/3.0:1)
 *   6. WCAG AA (light mode)    (同上)
 *   7. Tailwind darkMode 設定  (SharedTailwindConfig に darkMode:'class')
 *   8. themeManager API        (mount + apply + persist 完備)
 *   9. UI mount 点             (mode=view ヘッダー + mode=admin 設定)
 *  10. Unit tests              (themeManager 11 件 + 全 957 件 PASS)
 *
 * スコア: 各 10 点満点 → 合計 100 点。 0 違反かつ 100/100 で「完成」。
 *
 * 使い方:
 *   npm run theme:verify             # ヒューマンリーダブル出力
 *   npm run theme:verify -- --json   # JSON (CI 用)
 */

'use strict';

const fs = require('node:fs');
const path = require('node:path');
const { spawnSync } = require('node:child_process');

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
// 7. Tailwind darkMode 設定
// =====================================================================
function checkTailwindConfig() {
  const cfg = fs.readFileSync(path.join(SRC_DIR, 'SharedTailwindConfig.html'), 'utf8');
  const hasDarkMode = /darkMode\s*:\s*['"]class['"]/.test(cfg);
  // theme.extend.colors に CSS 変数 mapping があるかチェック (key 名は何でもよく、 value が var(--theme-*))
  const hasThemeColors = /var\(--theme-bg-base\)/.test(cfg) && /var\(--theme-text-primary\)/.test(cfg);
  return { hasDarkMode, hasThemeColors, ok: hasDarkMode && hasThemeColors };
}

const twConfig = checkTailwindConfig();
record(
  '7. Tailwind darkMode 設定',
  10,
  twConfig.ok,
  twConfig.ok
    ? "darkMode:'class' + extend.colors(theme-bg-base 等) 完備"
    : `darkMode:${twConfig.hasDarkMode ? '✓' : '✗'} colors:${twConfig.hasThemeColors ? '✓' : '✗'}`
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

const tm = checkThemeManager();
record(
  '8. themeManager API',
  10,
  tm.ok,
  tm.ok ? '9 機能完備 (get/set/toggle/subscribe/init/mountToggle/autoMql/storageKey/api)' :
    `欠落: ${tm.missing.join(', ')}`
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
  const result = spawnSync('npm', ['test', '--silent'], {
    cwd: ROOT, encoding: 'utf8'
  });
  // node:test 出力から pass/fail を抽出
  const output = (result.stdout || '') + (result.stderr || '');
  const passMatch = output.match(/ℹ pass (\d+)/);
  const failMatch = output.match(/ℹ fail (\d+)/);
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
