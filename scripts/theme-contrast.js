#!/usr/bin/env node
/**
 * theme-contrast.js — WCAG 2.1 contrast ratio 検証。
 *
 * 確認:
 *   - 本文 (--theme-text-primary on --theme-bg-base) ≥ 4.5:1 (AA)
 *   - secondary text ≥ 4.5:1
 *   - muted text ≥ 4.5:1 (子供向けに厳格)
 *   - accent (--theme-accent-cyan on bg) ≥ 4.5:1 (リンク・ボタンテキスト)
 *   - status (success/error/warning on bg) ≥ 3:1 (アイコン用)
 *
 * 終了コード: 全部 AA pass → 0、 fail があれば → 1
 */

'use strict';

const fs = require('node:fs');
const path = require('node:path');
const { parseColor, blendOnBg, relativeLuminance, extractBlockBody, extractTokensFromBody } = require('./lib/theme-core');

const SRC_DIR = path.resolve(__dirname, '../src');
const ARGS = process.argv.slice(2);
const FLAG_JSON = ARGS.includes('--json');

// =====================================================================
// WCAG 2.1 contrast ratio
//   parseColor / blendOnBg / relativeLuminance は scripts/lib/theme-core.js に
//   共通化。contrastRatio は theme-contrast 固有の「丸めない生 ratio」なので残す。
// =====================================================================

function contrastRatio(fg, bg) {
  const l1 = relativeLuminance(fg);
  const l2 = relativeLuminance(bg);
  const lighter = Math.max(l1, l2);
  const darker = Math.min(l1, l2);
  return (lighter + 0.05) / (darker + 0.05);
}

// =====================================================================
// Token 抽出 (extractBlockBody / token 抽出ロジックは
//   scripts/lib/theme-core.js に共通化)
// =====================================================================

function extractTokens(css, startRegex) {
  return extractTokensFromBody(extractBlockBody(css, startRegex));
}

function resolveToken(tokens, name) {
  let value = tokens[name];
  if (!value) return null;
  // var(--other) 参照を解決
  const varMatch = value.match(/var\(--([a-z0-9-]+)(?:,\s*([^)]+))?\)/);
  if (varMatch) {
    const referenced = `--${varMatch[1]}`;
    if (tokens[referenced]) return resolveToken(tokens, referenced);
    if (varMatch[2]) value = varMatch[2].trim();
  }
  return value;
}

// =====================================================================
// Main check
// =====================================================================

function main() {
  const cssPath = path.join(SRC_DIR, 'UnifiedStyles.css.html');
  const css = fs.readFileSync(cssPath, 'utf8');

  const dark = extractTokens(css, /:root\s*\{/);
  const light = extractTokens(css, /body\.theme-light[^{]*\{/);

  const tests = [
    // pair: { fg-token, bg-token, label, min ratio, large-text? }
    // 注: body.theme-light で上書きされない token は dark の値が使われる
    { fg: '--theme-text-primary', bg: '--theme-bg-base', label: '本文 (primary text)', min: 4.5 },
    { fg: '--theme-text-secondary', bg: '--theme-bg-base', label: 'secondary text', min: 4.5 },
    { fg: '--theme-text-muted', bg: '--theme-bg-base', label: 'muted text', min: 4.5 },
    { fg: '--theme-accent-cyan', bg: '--theme-bg-base', label: 'accent cyan (links / button text)', min: 4.5 },
    { fg: '--theme-spinner', bg: '--theme-bg-base', label: 'spinner color', min: 3.0 },
    { fg: '--status-success', bg: '--theme-bg-base', label: 'status success', min: 3.0 },
    { fg: '--status-error', bg: '--theme-bg-base', label: 'status error', min: 3.0 },
    { fg: '--status-warning', bg: '--theme-bg-base', label: 'status warning', min: 3.0 },
    { fg: '--brand-like', bg: '--theme-bg-base', label: 'brand-like (LIKE reaction)', min: 3.0 },
    { fg: '--brand-understand', bg: '--theme-bg-base', label: 'brand-understand (amber)', min: 3.0 },
    { fg: '--brand-curious', bg: '--theme-bg-base', label: 'brand-curious (emerald)', min: 3.0 },
  ];

  const results = { dark: [], light: [] };

  for (const mode of ['dark', 'light']) {
    const tokens = mode === 'dark' ? dark : { ...dark, ...light };  // light は dark に上書き
    for (const t of tests) {
      const fgValue = resolveToken(tokens, t.fg);
      const bgValue = resolveToken(tokens, t.bg);
      if (!fgValue || !bgValue) {
        results[mode].push({ ...t, status: 'skip', reason: 'token unresolved', fgValue, bgValue });
        continue;
      }
      let fg = parseColor(fgValue);
      let bg = parseColor(bgValue);
      if (!fg || !bg) {
        results[mode].push({ ...t, status: 'skip', reason: 'parse fail', fgValue, bgValue });
        continue;
      }
      if (fg.a < 1) fg = blendOnBg(fg, bg);
      if (bg.a < 1) bg = blendOnBg(bg, { r: 255, g: 255, b: 255 });  // assume white below
      const ratio = contrastRatio(fg, bg);
      results[mode].push({
        ...t,
        ratio: Math.round(ratio * 100) / 100,
        status: ratio >= t.min ? 'pass' : 'fail',
        fgValue,
        bgValue,
      });
    }
  }

  if (FLAG_JSON) {
    console.log(JSON.stringify(results, null, 2));
    const totalFail = [...results.dark, ...results.light].filter(r => r.status === 'fail').length;
    process.exit(totalFail === 0 ? 0 : 1);
  }

  console.log('=== WCAG Contrast Audit ===\n');
  let totalFail = 0;
  for (const mode of ['dark', 'light']) {
    console.log(`【${mode.toUpperCase()} mode】`);
    for (const r of results[mode]) {
      const marker = r.status === 'pass' ? '✓' : r.status === 'fail' ? '✗' : '⚠';
      const ratio = r.ratio ? r.ratio.toFixed(2).padStart(6) : '  N/A';
      console.log(`  ${marker} ${ratio} (target ≥${r.min.toFixed(1)})  ${r.label}`);
      if (r.status === 'fail') {
        console.log(`         fg=${r.fgValue}  bg=${r.bgValue}`);
        totalFail++;
      }
    }
    console.log('');
  }

  if (totalFail === 0) {
    console.log('✅ WCAG AA: 全 contrast pass (dark + light 両モード)');
  } else {
    console.log(`✗ WCAG AA: ${totalFail} 件 fail`);
  }
  process.exit(totalFail === 0 ? 0 : 1);
}

main();
