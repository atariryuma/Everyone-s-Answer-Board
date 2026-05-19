#!/usr/bin/env node
/**
 * theme-pair-tailwind.js — Tailwind の単色 utility に `dark:` ペアを付与する。
 *
 * 何故書いたか:
 *   `text-cyan-400` は dark mode で適切 (light cyan on dark bg) だが、 light mode で
 *   white bg 上 1.6:1 で WCAG AA 不可。 `text-cyan-700 dark:text-cyan-400` のように
 *   両モード対応にする必要がある。 HTML 側の Tailwind class 自動ペアリング tool。
 *
 * 対象:
 *   - text-{cyan|yellow|green|blue|violet|red|pink|purple}-{200|300|400} → 対応する 700/800 と pair
 *   - bg-{violet|yellow|blue|red}-900/{X} → light の bg-{X}-100 と pair
 *   - skip: solid action button bg (bg-X-600+ with text-white in same class) — 既に視認 OK
 *   - skip: pure brand utility (bg-cyan-600 等のボタン)
 *   - skip: 既に `dark:` ペアが存在するもの
 *
 * 使い方:
 *   node scripts/theme-pair-tailwind.js --dry-run
 *   node scripts/theme-pair-tailwind.js
 */

'use strict';

const fs = require('node:fs');
const path = require('node:path');
const glob = (pattern) => {
  // simple expansion: only handles src/*.html or src/*.js
  const dir = path.join(__dirname, '..', 'src');
  return fs.readdirSync(dir).filter(f => f.endsWith('.html') || f.endsWith('.js')).map(f => path.join(dir, f));
};

const DRY = process.argv.includes('--dry-run');

// text-color pairing: 各 (color, shade) を light mode 用 (color, dark-shade) と pair
// 設計: 200/300/400 は light mode で読みづらい → 700 or 800 にペア。
//   500 はちょうど中間で両モード可。 600+ は dark mode で読みづらいので light 側専用扱いも検討。
const TEXT_PAIRS = {
  // light pale colors (dark mode に最適だが light mode で fail)
  '100': '800',  // 100 (extremely pale) ↔ 800 (very dark)
  '200': '800',  // 200 (very pale) ↔ 800 (very dark) — for hero accents like text-cyan-200
  '300': '700',  // 300 (pale)      ↔ 700 (dark)      — typical accent text
  '400': '700',  // 400 (mid-light) ↔ 700 (dark)      — most common
};

const PAIRABLE_COLORS = new Set([
  'cyan', 'sky', 'blue', 'indigo', 'violet', 'purple', 'fuchsia', 'pink', 'rose',
  'red', 'orange', 'amber', 'yellow', 'lime', 'green', 'emerald', 'teal',
]);

// bg-X-900/{X} は dark mode の depressed bg、 light mode では bg-X-100 系に
const BG_DEPRESSED_PAIRS = {
  // bg-X-900/Y → light: bg-X-100/Y (alpha 同じ)
  '900': '100',
  '800': '100',
};

// "this class is on a solid action button" detector: in the class string
// もし bg-{color}-{500+} + text-white が同じ class 文字列にあれば、 そのボタンは brand action
// (両モードで bg は brand 色のまま、 text-white は brand 色 ≥ 600 上で必ず読める)
function isSolidActionButtonClassStr(classStr) {
  const m = classStr.match(/\bbg-(cyan|sky|blue|indigo|violet|purple|pink|rose|red|orange|amber|green|emerald|teal)-(\d{3})(?:\/\d+)?/);
  if (!m) return false;
  const shade = parseInt(m[2], 10);
  if (shade < 500) return false;
  // text-white が同じ class 内に
  if (!/\btext-white\b/.test(classStr)) return false;
  return true;
}

// hover / focus / active 等の prefix を保持しつつ、 prop-color-shade のみ抽出
const PREFIX_PARTS = /(?:hover:|focus:|focus-visible:|active:|disabled:|group-hover:|focus-within:|peer-checked:)*/;
const CLASS_RE = /class=(["'])([^"']+)\1/g;

function pairClassString(classStr) {
  if (isSolidActionButtonClassStr(classStr)) return { result: classStr, changed: 0 };
  let changed = 0;
  // 全 token を space で split
  const tokens = classStr.split(/\s+/);
  const out = [];
  for (let i = 0; i < tokens.length; i++) {
    const tok = tokens[i];
    if (!tok) { out.push(tok); continue; }
    // 既に dark: ペアが next に居るか、または このトークンが dark: prefix なら skip
    const m = tok.match(/^((?:hover:|focus:|focus-visible:|active:|disabled:|group-hover:|focus-within:|peer-checked:)*)(text|bg|border|ring|from|via|to|divide|placeholder|fill|stroke|caret|accent|outline|decoration|shadow)-([a-z]+)-(\d{2,3})(\/\d+)?$/);
    if (!m) { out.push(tok); continue; }
    const [, prefix, prop, color, shade, alpha] = m;
    // dark: / light: prefix は skip
    if (/^(dark|light):/.test(tok)) { out.push(tok); continue; }
    if (!PAIRABLE_COLORS.has(color)) { out.push(tok); continue; }
    // 既にペアがあるか確認 (前後 + 全 token に dark:{prefix}{prop}-{color} が居れば skip)
    const hasDarkPair = tokens.some(t => {
      const dm = t.match(/^dark:(.*)/);
      if (!dm) return false;
      // dark prefix を剥がした残りで同 prop/color が居るかチェック
      return new RegExp(`^${prefix}${prop}-${color}-`).test(dm[1]);
    });
    if (hasDarkPair) { out.push(tok); continue; }
    // text-{color}-{200/300/400} のみペア化対象 (それ以下/以上はスキップ)
    if (prop === 'text') {
      const lightShade = TEXT_PAIRS[shade];
      if (!lightShade) { out.push(tok); continue; }
      // light mode 用 base を先に置く、 dark mode 用は dark: prefix で追加
      const lightToken = `${prefix}${prop}-${color}-${lightShade}${alpha || ''}`;
      out.push(lightToken);
      out.push(`dark:${prefix}${prop}-${color}-${shade}${alpha || ''}`);
      changed++;
      continue;
    }
    if (prop === 'bg' && (shade === '900' || shade === '800')) {
      const lightShade = BG_DEPRESSED_PAIRS[shade];
      const lightToken = `${prefix}${prop}-${color}-${lightShade}${alpha || ''}`;
      out.push(lightToken);
      out.push(`dark:${prefix}${prop}-${color}-${shade}${alpha || ''}`);
      changed++;
      continue;
    }
    if (prop === 'border' && shade === '500' && alpha && parseInt(alpha.slice(1)) < 50) {
      // border-X-500/30 系の subtle border — light でも border-X-300/30 にしておくと AA に近づく
      const lightToken = `${prefix}${prop}-${color}-300${alpha}`;
      out.push(lightToken);
      out.push(`dark:${prefix}${prop}-${color}-${shade}${alpha}`);
      changed++;
      continue;
    }
    out.push(tok);
  }
  return { result: out.join(' '), changed };
}

function processFile(filePath) {
  const original = fs.readFileSync(filePath, 'utf8');
  let totalChanges = 0;
  const changes = [];
  const result = original.replace(CLASS_RE, (full, quote, classStr) => {
    const { result: paired, changed } = pairClassString(classStr);
    if (changed > 0) {
      totalChanges += changed;
      const line = original.slice(0, original.indexOf(full)).split('\n').length;
      changes.push({ line, before: classStr.slice(0, 80), after: paired.slice(0, 80) });
    }
    return `class=${quote}${paired}${quote}`;
  });
  if (!DRY && totalChanges > 0) {
    fs.writeFileSync(filePath, result, 'utf8');
  }
  return { file: path.basename(filePath), totalChanges, changes };
}

function main() {
  console.log(`Theme pair Tailwind — ${DRY ? 'DRY RUN' : 'WRITING'}\n`);
  let grand = 0;
  for (const fp of glob()) {
    const r = processFile(fp);
    if (r.totalChanges > 0) {
      console.log(`  ${r.file}: ${r.totalChanges} pairings`);
      if (DRY) r.changes.slice(0, 5).forEach(c => {
        console.log(`    L${c.line}: ${c.before}`);
        console.log(`         → ${c.after}`);
      });
      grand += r.totalChanges;
    }
  }
  console.log(`\nTotal pairings: ${grand}`);
}

main();
