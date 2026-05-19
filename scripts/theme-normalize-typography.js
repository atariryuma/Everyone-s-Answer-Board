#!/usr/bin/env node
/**
 * theme-normalize-typography.js — 中途半端な font-size / border-radius 値を
 * Tailwind 互換スケール (= --font-size-* / --radius-* token) に正規化。
 *
 * 何を直す:
 *   - 0.7-0.78rem → --font-size-xs (0.75rem)
 *   - 0.8-0.92rem → --font-size-sm (0.875rem)
 *   - 0.95rem → --font-size-base (1rem)
 *   - 1.05-1.15rem → --font-size-lg (1.125rem)
 *   - 1.2-1.3rem → --font-size-xl (1.25rem)
 *   - 1.4-1.6rem → --font-size-2xl (1.5rem)
 *   - 999px → 9999px (rounded-full)
 *   - 0.3-0.45rem border-radius → 0.5rem (rounded)
 *
 * 視覚インパクト最小 (差は 1-3px 程度)。 全て Tailwind スケールに合流。
 *
 * 制約:
 *   - :root / body.theme-light の token 定義ブロックは触らない
 *   - theme:exempt が同行にある行は skip
 *
 * 使い方:
 *   node scripts/theme-normalize-typography.js --dry-run
 *   node scripts/theme-normalize-typography.js
 */

'use strict';
const fs = require('node:fs');
const path = require('node:path');

const SRC = path.resolve(__dirname, '..', 'src');
const DRY = process.argv.includes('--dry-run');

// CSS / HTML 全部対象
const FILES = fs.readdirSync(SRC).filter(f => f.endsWith('.html') || f.endsWith('.js'));

// 値マップ: 正確な regex (number + unit) → 置換値
const FONT_SIZE_MAP = [
  // xs
  [/font-size:\s*0\.7rem\b/g,    'font-size: 0.75rem' /* var(--font-size-xs) */],
  [/font-size:\s*0\.72rem\b/g,   'font-size: 0.75rem'],
  [/font-size:\s*0\.73rem\b/g,   'font-size: 0.75rem'],
  [/font-size:\s*0\.78rem\b/g,   'font-size: 0.75rem'],
  // sm
  [/font-size:\s*0\.8rem\b/g,    'font-size: 0.875rem'],
  [/font-size:\s*0\.82rem\b/g,   'font-size: 0.875rem'],
  [/font-size:\s*0\.85rem\b/g,   'font-size: 0.875rem'],
  [/font-size:\s*0\.88rem\b/g,   'font-size: 0.875rem'],
  [/font-size:\s*0\.9rem\b/g,    'font-size: 0.875rem'],
  [/font-size:\s*0\.92rem\b/g,   'font-size: 0.875rem'],
  // base
  [/font-size:\s*0\.95rem\b/g,   'font-size: 1rem'],
  [/font-size:\s*0\.95em\b/g,    'font-size: 1em'],  // em は近接合致
  // lg
  [/font-size:\s*1\.05rem\b/g,   'font-size: 1.125rem'],
  [/font-size:\s*1\.1rem\b/g,    'font-size: 1.125rem'],
  [/font-size:\s*1\.15rem\b/g,   'font-size: 1.125rem'],
  // xl
  [/font-size:\s*1\.2rem\b/g,    'font-size: 1.25rem'],
  [/font-size:\s*1\.3rem\b/g,    'font-size: 1.25rem'],
  // 2xl
  [/font-size:\s*1\.4rem\b/g,    'font-size: 1.5rem'],
  [/font-size:\s*1\.6rem\b/g,    'font-size: 1.5rem'],
  [/font-size:\s*1\.65rem\b/g,   'font-size: 1.5rem'],
  // 3xl
  [/font-size:\s*1\.7rem\b/g,    'font-size: 1.875rem'],
  [/font-size:\s*1\.85rem\b/g,   'font-size: 1.875rem'],
];

const RADIUS_MAP = [
  // rounded-full
  [/border-radius:\s*999px\b/g,  'border-radius: 9999px'],
  // rounded-sm (0.25rem)
  [/border-radius:\s*0\.2rem\b/g,  'border-radius: 0.25rem'],
  // rounded (0.5rem) — 0.3-0.45 は近接して見た目変化最小
  [/border-radius:\s*0\.3rem\b/g,  'border-radius: 0.5rem'],
  [/border-radius:\s*0\.35rem\b/g, 'border-radius: 0.5rem'],
  [/border-radius:\s*0\.375rem\b/g,'border-radius: 0.5rem'],
  [/border-radius:\s*0\.4rem\b/g,  'border-radius: 0.5rem'],
  [/border-radius:\s*0\.45rem\b/g, 'border-radius: 0.5rem'],
  // rounded-lg (0.75rem) — 0.6 は近接
  [/border-radius:\s*0\.6rem\b/g,  'border-radius: 0.75rem'],
  // px → rem 単位の小さな値も Tailwind 風に
  [/border-radius:\s*3px\b/g,    'border-radius: 0.25rem'],
  [/border-radius:\s*4px\b/g,    'border-radius: 0.25rem'],
  [/border-radius:\s*6px\b/g,    'border-radius: 0.375rem'],
  [/border-radius:\s*8px\b/g,    'border-radius: 0.5rem'],
  [/border-radius:\s*10px\b/g,   'border-radius: 0.625rem'],
  [/border-radius:\s*12px\b/g,   'border-radius: 0.75rem'],
];

function findExcludedRanges(text) {
  const ranges = [];
  const RES = [/:root\s*\{/g, /body\.theme-light[^{]*\{/g, /:root\[data-theme="light"\]\s*\{/g];
  for (const re of RES) {
    re.lastIndex = 0;
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

function processFile(filePath) {
  const original = fs.readFileSync(filePath, 'utf8');
  const excluded = findExcludedRanges(original);
  let result = original;
  let totalChanges = 0;

  for (const [pat, rep] of [...FONT_SIZE_MAP, ...RADIUS_MAP]) {
    pat.lastIndex = 0;
    result = result.replace(pat, (match, ...args) => {
      const offset = typeof args[args.length - 2] === 'number' ? args[args.length - 2] : 0;
      // exempt range 内 or theme:exempt 行ならスキップ
      if (excluded.some(([s, e]) => offset >= s && offset < e)) return match;
      // 同じ行に theme:exempt
      const lineStart = result.lastIndexOf('\n', offset) + 1;
      const lineEnd = result.indexOf('\n', offset);
      const line = result.slice(lineStart, lineEnd === -1 ? result.length : lineEnd);
      if (/theme:exempt/i.test(line)) return match;
      totalChanges++;
      return rep;
    });
  }

  if (!DRY && totalChanges > 0) fs.writeFileSync(filePath, result, 'utf8');
  return { file: path.basename(filePath), totalChanges };
}

function main() {
  console.log(`Typography normalize — ${DRY ? 'DRY RUN' : 'WRITING'}\n`);
  let grand = 0;
  for (const f of FILES) {
    const r = processFile(path.join(SRC, f));
    if (r.totalChanges > 0) {
      console.log(`  ${r.file}: ${r.totalChanges}`);
      grand += r.totalChanges;
    }
  }
  console.log(`\nTotal: ${grand}`);
}

main();
