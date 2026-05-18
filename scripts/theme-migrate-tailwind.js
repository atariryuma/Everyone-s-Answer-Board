#!/usr/bin/env node
/**
 * theme-migrate-tailwind.js — Tailwind ハードコード class を light/dark ペアに変換。
 *
 * 戦略:
 *   既存の dark 専用 class (`text-gray-300`, `bg-gray-800` 等) を
 *   `light-class dark:dark-class` 形式に変換し、 body.theme-light で light が
 *   適用されるようにする (Tailwind darkMode: 'class' 設定下)。
 *
 * 例:
 *   text-gray-300  →  text-slate-700 dark:text-gray-300
 *   bg-gray-800    →  bg-white dark:bg-gray-800
 *   border-gray-600 → border-slate-300 dark:border-gray-600
 *
 * 注意:
 *   - すでに dark: / light: prefix が付いている class はスキップ
 *   - 文字列リテラル (JS の '...' / "...") 内の class も対象 (page.js / AdminPanel.js)
 *   - 慎重を期するため変更前後の差分を表示。 --apply で書き込み。
 */

'use strict';

const fs = require('node:fs');
const path = require('node:path');

const SRC_DIR = path.resolve(__dirname, '../src');
const APPLY = process.argv.includes('--apply');
const FILE_FILTER = process.argv.find(a => a.startsWith('--file='));
const TARGET_FILE = FILE_FILTER ? FILE_FILTER.slice('--file='.length) : null;

// Light → Dark mapping (light class → dark class を併記)
// slate ファミリーで統一 (cool, projector-friendly)。
//
// 設計原則:
//   - Dark の color level を反転 (gray-800 ≈ slate-50, gray-300 ≈ slate-700 等)
//   - 不変であるべき class (white on colored bg 等) は対象外
const MAPPING = {
  // text-gray-N : N が高い = dark mode で明るい text。 light では low N = 暗い text。
  'text-gray-100': 'text-slate-900',
  'text-gray-200': 'text-slate-800',
  'text-gray-300': 'text-slate-700',
  'text-gray-400': 'text-slate-600',
  'text-gray-500': 'text-slate-500',
  'text-gray-600': 'text-slate-400',
  'text-white':    'text-slate-900',   // light 上では slate-900 が "primary text"
  // bg-gray-N : 高 N = 暗い。 light では low N (明るい)。
  'bg-gray-900':   'bg-slate-100',
  'bg-gray-800':   'bg-white',
  'bg-gray-700':   'bg-slate-50',
  'bg-gray-600':   'bg-slate-100',
  'bg-gray-500':   'bg-slate-200',
  // border-gray-N
  'border-gray-900': 'border-slate-200',
  'border-gray-800': 'border-slate-200',
  'border-gray-700': 'border-slate-300',
  'border-gray-600': 'border-slate-300',
  'border-gray-500': 'border-slate-400',
};

// 「変換しない」class — brand-colored bg 上の白テキスト、 reaction の text-red/yellow/green 等
const SKIP_CONTEXTS = [
  /bg-(red|blue|green|yellow|cyan|purple|indigo|pink|emerald|amber|rose|orange|sky|violet|fuchsia)-/,
  /bg-gradient/,
  /from-/,
  /to-/,
];

function migrateContent(content) {
  let modified = content;
  const changes = [];

  for (const [darkClass, lightClass] of Object.entries(MAPPING)) {
    // 既に dark: / light: prefix 付いている場合は skip
    // 単独 class として class="..." 内に出現するか check
    // class 区切りは whitespace / " / '
    const escaped = darkClass.replace(/-/g, '\\-');
    const re = new RegExp(`(^|[\\s"'])(${escaped})(?=[\\s"'>])`, 'g');
    modified = modified.replace(re, (match, pre, cls, offset) => {
      // 前後を確認: dark: または light: の直後ではないか
      const before = content.slice(Math.max(0, offset - 6), offset + pre.length);
      if (/dark:$|light:$/.test(before.trim())) return match;
      // 同行に既に dark:darkClass がある = ペア確定済みかも
      const lineStart = content.lastIndexOf('\n', offset);
      const lineEnd = content.indexOf('\n', offset);
      const line = content.slice(lineStart, lineEnd > 0 ? lineEnd : content.length);
      if (line.includes(`dark:${cls}`)) return match;
      // brand bg ペアならスキップ (class="bg-red-600 text-white" 等)
      // 同 attribute 値内の他 class を取得
      // 安全側: text-white は常に変換 (どんな bg でも light テーマでは slate-900)
      // ただし「bg-red/blue/...」の同行に text-white があれば skip
      if (cls === 'text-white') {
        const classMatch = line.match(/class="[^"]*"|className="[^"]*"|class='[^']*'/);
        if (classMatch && SKIP_CONTEXTS.some(re => re.test(classMatch[0]))) return match;
      }
      changes.push({ from: cls, to: `${lightClass} dark:${cls}` });
      return `${pre}${lightClass} dark:${cls}`;
    });
  }

  return { modified, changes };
}

function main() {
  let files = fs.readdirSync(SRC_DIR).filter(f => f.endsWith('.html'));
  if (TARGET_FILE) files = files.filter(f => f === TARGET_FILE || `src/${f}` === TARGET_FILE);
  // CSS のみ含む HTML / Tailwind 設定 / d3 vendor は対象外
  const CSS_ONLY = new Set([
    'SharedTailwindConfig.html',
    'UnifiedStyles.css.html',
    'page.css.html',
    'page.viz.css.html',
    'SharedTailwindConfig.html',
    'd3.min.html',
    'tinySegmenter.html',
  ]);
  files = files.filter(f => !CSS_ONLY.has(f));

  const reports = {};
  let totalChanges = 0;

  for (const file of files) {
    const fullPath = path.join(SRC_DIR, file);
    const content = fs.readFileSync(fullPath, 'utf8');
    const { modified, changes } = migrateContent(content);
    if (changes.length === 0) continue;
    reports[file] = changes;
    totalChanges += changes.length;
    if (APPLY && modified !== content) {
      fs.writeFileSync(fullPath, modified, 'utf8');
    }
  }

  console.log(`=== Tailwind dark: migration ${APPLY ? '(applied)' : '(dry-run)'} ===\n`);
  for (const [file, changes] of Object.entries(reports)) {
    const counts = {};
    for (const c of changes) counts[c.from] = (counts[c.from] || 0) + 1;
    console.log(`src/${file}: ${changes.length} changes`);
    for (const [from, n] of Object.entries(counts).sort((a, b) => b[1] - a[1])) {
      const to = MAPPING[from];
      console.log(`  ${String(n).padStart(3)} × ${from} → ${to} dark:${from}`);
    }
  }
  console.log(`\nTotal: ${totalChanges} migrations across ${Object.keys(reports).length} files.`);
  if (!APPLY) {
    console.log('\nRe-run with --apply to write changes.');
  }
}

main();
