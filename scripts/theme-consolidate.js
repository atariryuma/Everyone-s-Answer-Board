#!/usr/bin/env node
/**
 * theme-consolidate.js — Tailwind `bg-X dark:bg-Y` ペアを theme token utility に統一。
 *
 * 設計思想 (v2818+):
 *   - dark: modifier 全廃止し theme token utility のみで両モード対応
 *   - 色を変えたいとき UnifiedStyles.css.html の 1 トークン書き換えで全 HTML に反映
 *   - GitHub Primer / shadcn / Radix 標準パターンに完全準拠
 *
 * 置換マップ (semantic 推測ベース):
 *   text-slate-900 dark:text-slate-100  → text-theme (本文)
 *   text-slate-700 dark:text-slate-300  → text-theme-secondary
 *   text-slate-600 dark:text-slate-400  → text-theme-muted
 *   bg-white dark:bg-slate-800           → bg-theme-surface (カード)
 *   bg-slate-50 dark:bg-slate-900        → bg-theme-base (page bg)
 *   border-slate-200 dark:border-slate-700  → border-theme
 *   border-slate-300 dark:border-slate-700  → border-theme-normal
 *
 * Skip (brand identity 保持):
 *   - text-pink-* (question heading 専用 brand)
 *   - text-cyan-* / text-red-* / text-green-* / text-yellow-* (brand 色)
 *   - bg-X-600+ (solid action button)
 *   - hover: / focus: 修飾子付き (状態 transition で意図的)
 */

'use strict';
const fs = require('node:fs');
const path = require('node:path');

const SRC = path.resolve(__dirname, '..', 'src');
const DRY = process.argv.includes('--dry-run');
const FILES = fs.readdirSync(SRC).filter(f => f.endsWith('.html') || f.endsWith('.js'));

// 置換ルール: [pattern, replacement, label]
// pattern は class 内空白区切りの token 連続を想定。 順序は固定 (X then dark:Y) で書かれている前提。
const PAIRS = [
  // ─── text (slate scale) ─────────────────
  ['text-slate-900 dark:text-slate-100', 'text-theme', 'primary text 16:1'],
  ['text-slate-900 dark:text-slate-200', 'text-theme', 'primary text'],
  ['text-slate-900 dark:text-white',     'text-theme', 'primary text'],
  ['text-slate-800 dark:text-slate-100', 'text-theme', 'primary text'],
  ['text-slate-800 dark:text-slate-200', 'text-theme', 'primary text'],
  ['text-slate-700 dark:text-slate-200', 'text-theme', 'primary leaning'],
  ['text-slate-700 dark:text-slate-300', 'text-theme-secondary', 'secondary text'],
  ['text-slate-600 dark:text-slate-300', 'text-theme-secondary', 'secondary text'],
  ['text-slate-600 dark:text-slate-400', 'text-theme-muted', 'muted/caption'],
  ['text-slate-500 dark:text-slate-400', 'text-theme-muted', 'muted'],
  ['text-slate-500 dark:text-slate-500', 'text-theme-muted', 'muted'],

  // ─── bg ─────────────────────────────────
  ['bg-white dark:bg-slate-800',         'bg-theme-surface', 'card surface'],
  ['bg-white dark:bg-slate-700',         'bg-theme-elevated', 'elevated'],
  ['bg-slate-50 dark:bg-slate-900',      'bg-theme-base', 'page base'],
  ['bg-slate-50 dark:bg-slate-800',      'bg-theme-surface', 'card surface'],
  ['bg-slate-100 dark:bg-slate-900',     'bg-theme-base', 'page base'],
  ['bg-slate-100 dark:bg-slate-800',     'bg-theme-surface', 'card surface'],
  ['bg-slate-100 dark:bg-slate-700',     'bg-theme-elevated', 'elevated'],
  ['bg-slate-200 dark:bg-slate-800',     'bg-theme-surface', 'card surface'],
  ['bg-slate-200 dark:bg-slate-700',     'bg-theme-elevated', 'elevated'],

  // ─── border ─────────────────────────────
  ['border-slate-200 dark:border-slate-700', 'border-theme', 'subtle border'],
  ['border-slate-200 dark:border-slate-600', 'border-theme', 'subtle border'],
  ['border-slate-300 dark:border-slate-700', 'border-theme-normal', 'normal border'],
  ['border-slate-300 dark:border-slate-600', 'border-theme-normal', 'normal border'],
  ['border-slate-300 dark:border-slate-500', 'border-theme-normal', 'normal border'],
  ['border-slate-400 dark:border-slate-500', 'border-theme-strong', 'strong border'],
  ['border-slate-400 dark:border-slate-600', 'border-theme-strong', 'strong border'],

  // ─── 順序逆 (dark: が先) も念のため ─────
  ['dark:text-slate-100 text-slate-900', 'text-theme', 'primary text (rev)'],
  ['dark:text-slate-300 text-slate-700', 'text-theme-secondary', 'secondary (rev)'],
  ['dark:text-slate-400 text-slate-600', 'text-theme-muted', 'muted (rev)'],
  ['dark:bg-slate-800 bg-white',         'bg-theme-surface', 'surface (rev)'],

  // ─── secondary button bg (slate-50/100 light + slate-600/700 dark) ─────
  ['bg-slate-50 dark:bg-slate-700',      'bg-theme-elevated', 'sec button bg'],
  ['bg-slate-100 dark:bg-slate-600',     'bg-theme-elevated', 'sec button bg'],
  ['bg-slate-100 dark:bg-slate-500',     'bg-theme-elevated', 'sec button bg'],

  // ─── divide ───
  ['divide-slate-200 dark:divide-slate-700', 'divide-theme', 'divider'],

  // ─── 単独 dark:text-slate-X (paired light なし — 単独利用は theme token に) ───
  // ※注意: 単独 dark:text-slate-X は light で何も適用されない潜在バグ。
  //   → 確実に text-theme-* に統一。
  ['text-slate-400 dark:text-slate-500', 'text-theme-muted', 'muted edge'],
  ['text-slate-500 dark:text-slate-600', 'text-theme-muted', 'muted edge'],
  ['text-slate-400 dark:text-slate-600', 'text-theme-muted', 'reverse paired muted'],

  // ─── 残り light-side gray-* (v2817 で dark のみ slate 化、 light gray 残存) ───
  ['text-gray-700 dark:text-slate-300', 'text-theme-secondary', 'mixed gray/slate'],
  ['text-gray-700 dark:text-slate-200', 'text-theme-secondary', 'mixed gray/slate'],
  ['text-gray-600 dark:text-slate-400', 'text-theme-muted', 'mixed gray/slate'],
  ['text-gray-500 dark:text-slate-400', 'text-theme-muted', 'mixed gray/slate'],
  ['bg-gray-100 dark:bg-slate-700',     'bg-theme-elevated', 'mixed gray/slate'],
  ['bg-gray-100 dark:bg-slate-800',     'bg-theme-surface', 'mixed gray/slate'],
  ['bg-gray-50 dark:bg-slate-900',      'bg-theme-base', 'mixed gray/slate'],
  ['bg-gray-50 dark:bg-slate-800',      'bg-theme-surface', 'mixed gray/slate'],
  ['border-gray-300 dark:border-slate-600', 'border-theme-normal', 'mixed gray/slate'],
  ['border-gray-300 dark:border-slate-700', 'border-theme-normal', 'mixed gray/slate'],
  ['border-gray-200 dark:border-slate-700', 'border-theme', 'mixed gray/slate'],
];

function processFile(filePath) {
  const original = fs.readFileSync(filePath, 'utf8');
  let result = original;
  const changes = {};
  for (const [from, to, label] of PAIRS) {
    // 単純な文字列置換 (class 内で連続している前提)
    const before = result;
    result = result.split(from).join(to);
    const count = (before.length - result.length) > 0 ? (before.split(from).length - 1) : 0;
    if (count > 0) {
      changes[label] = (changes[label] || 0) + count;
    }
  }
  if (!DRY && Object.keys(changes).length > 0) {
    fs.writeFileSync(filePath, result, 'utf8');
  }
  return { file: path.basename(filePath), changes, total: Object.values(changes).reduce((a, b) => a + b, 0) };
}

function main() {
  console.log(`Theme consolidate — Tailwind dark: → theme utility (${DRY ? 'DRY RUN' : 'WRITING'})\n`);
  let grand = 0;
  const allChanges = {};
  for (const f of FILES) {
    const r = processFile(path.join(SRC, f));
    if (r.total > 0) {
      console.log(`  ${r.file}: ${r.total} 件`);
      for (const [label, n] of Object.entries(r.changes)) {
        allChanges[label] = (allChanges[label] || 0) + n;
      }
      grand += r.total;
    }
  }
  console.log(`\n─── 内訳 ───`);
  for (const [label, n] of Object.entries(allChanges).sort((a, b) => b[1] - a[1])) {
    console.log(`  ${label.padEnd(30)} ${n} 件`);
  }
  console.log(`\nTotal: ${grand} 件`);
}

main();
