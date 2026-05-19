#!/usr/bin/env node
/**
 * theme-tokenize.js — CSS 内の hardcoded slate/gray rgba を semantic theme token に置換する。
 *
 * 何故書いたか:
 *   `body.theme-light` 上書きを class ごとに書くより、 token に集約した方が保守性が高い。
 *   多くの dark-only ハードコード (rgba(30,41,59,0.55) 等) は意味的に「card 背景」だが、
 *   light mode override が無いので broken。 これらを `--theme-card-*` 等の semantic token に
 *   付け替えることで、 token を一箇所いじれば全箇所が theme 追従する。
 *
 *   `--theme-card-1/2/3` および `--theme-border-subtle/normal` 等は既に
 *   `:root` (dark) と `body.theme-light` (light) の両方で定義済み。
 *
 * 使い方:
 *   node scripts/theme-tokenize.js --dry-run   # 何を置換するか確認
 *   node scripts/theme-tokenize.js             # 実際に書き換える
 *
 * 制約:
 *   - `:root { ... }` および `body.theme-light { ... }` ブロック内は touch しない
 *     (token 定義自身が rgba 値を持つので置換すると無限再帰になる)
 *   - `theme:exempt` コメントが付いた行は touch しない
 */

'use strict';

const fs = require('node:fs');
const path = require('node:path');

const SRC_DIR = path.resolve(__dirname, '../src');
const DRY = process.argv.includes('--dry-run');

const TARGETS = ['UnifiedStyles.css.html', 'page.css.html', 'page.viz.css.html'];

// 置換マッピング — slate / gray scale rgba を semantic token に
//   (元値, 置換後 token, コメント)
const REPLACEMENTS = [
  // card 系背景 (mid elevation)
  ['rgba(30, 41, 59, 0.7)',   'var(--theme-card-1)',       'slate-800 70% → card-1'],
  ['rgba(30, 41, 59, 0.6)',   'var(--theme-card-1)',       'slate-800 60% → card-1'],
  ['rgba(30, 41, 59, 0.55)',  'var(--theme-card-1)',       'slate-800 55% → card-1'],
  ['rgba(30, 41, 59, 0.5)',   'var(--theme-card-1)',       'slate-800 50% → card-1'],
  ['rgba(30, 41, 59, 0.3)',   'var(--theme-card-1)',       'slate-800 30% → card-1'],
  // deep elevation (popup / modal 背面)
  ['rgba(15, 23, 42, 0.5)',   'var(--theme-card-2)',       'slate-900 50% → card-2'],
  ['rgba(15, 23, 42, 0.35)',  'var(--theme-card-2)',       'slate-900 35% → card-2'],
  ['rgba(15, 23, 42, 0.32)',  'var(--theme-card-2)',       'slate-900 32% → card-2'],
  // solid panel (popup 完全不透明)
  ['rgba(15, 23, 42, 0.98)',  'var(--theme-card-3)',       'slate-900 98% → card-3'],
  ['rgba(15, 23, 42, 0.95)',  'var(--theme-card-3)',       'slate-900 95% → card-3'],
  ['rgba(15, 23, 42, 0.9)',   'var(--theme-card-3)',       'slate-900 90% → card-3'],
  ['rgba(15, 23, 42, 0.8)',   'var(--theme-card-3)',       'slate-900 80% → card-3'],
  // border (slate-600 alpha)
  ['rgba(71, 85, 105, 0.6)',  'var(--theme-border-normal)', 'slate-600 60% → border-normal'],
  ['rgba(71, 85, 105, 0.4)',  'var(--theme-border-subtle)', 'slate-600 40% → border-subtle'],
  ['rgba(71, 85, 105, 0.3)',  'var(--theme-border-subtle)', 'slate-600 30% → border-subtle'],
];

// :root / body.theme-light ブロック内のオフセット範囲を計算 (置換除外)
function findExcludedRanges(text) {
  const ranges = [];
  const SELECTOR_PATTERNS = [/:root\s*\{/g, /body\.theme-light[^{]*\{/g, /:root\[data-theme="light"\]\s*\{/g];
  for (const re of SELECTOR_PATTERNS) {
    let m;
    while ((m = re.exec(text)) !== null) {
      const start = m.index;
      let depth = 0, i = start;
      while (i < text.length) {
        if (text[i] === '{') depth++;
        else if (text[i] === '}') {
          depth--;
          if (depth === 0) { ranges.push([start, i + 1]); break; }
        }
        i++;
      }
    }
  }
  return ranges;
}

function inExcluded(ranges, idx) {
  return ranges.some(([s, e]) => idx >= s && idx < e);
}

function processFile(filePath) {
  const original = fs.readFileSync(filePath, 'utf8');
  const excluded = findExcludedRanges(original);
  let result = original;
  let totalChanges = 0;
  const changes = [];

  for (const [from, to, _label] of REPLACEMENTS) {
    let offset = 0;
    const next = [];
    let i = 0;
    while (i < result.length) {
      const found = result.indexOf(from, i);
      if (found === -1) {
        next.push(result.slice(i));
        break;
      }
      // line に theme:exempt が含まれているか
      const lineStart = result.lastIndexOf('\n', found) + 1;
      const lineEnd = result.indexOf('\n', found);
      const line = result.slice(lineStart, lineEnd === -1 ? result.length : lineEnd);
      const skip = /theme:exempt/i.test(line) || inExcluded(excluded, found - offset);
      if (skip) {
        next.push(result.slice(i, found + from.length));
      } else {
        next.push(result.slice(i, found));
        next.push(to);
        totalChanges++;
        changes.push({ from, to, line: result.slice(0, found).split('\n').length });
      }
      i = found + from.length;
    }
    result = next.join('');
  }

  if (!DRY && totalChanges > 0) {
    fs.writeFileSync(filePath, result, 'utf8');
  }
  return { file: path.basename(filePath), totalChanges, changes };
}

function main() {
  console.log(`Theme tokenize — ${DRY ? 'DRY RUN (no writes)' : 'WRITING'}\n`);
  let grand = 0;
  for (const f of TARGETS) {
    const fp = path.join(SRC_DIR, f);
    if (!fs.existsSync(fp)) continue;
    const r = processFile(fp);
    console.log(`  ${r.file}: ${r.totalChanges} replacements`);
    if (DRY) r.changes.slice(0, 10).forEach(c => console.log(`    L${c.line}: ${c.from} → ${c.to}`));
    grand += r.totalChanges;
  }
  console.log(`\nTotal: ${grand} replacements`);
}

main();
