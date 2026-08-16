#!/usr/bin/env node
/**
 * tokenize-dimensions.js — CSS 内の寸法の生値を、既存の設計スケール token に置換する。
 *
 * 何故書いたか:
 *   token 体系 (--radius-* 6 個 / --font-size-* 8 個 / --space-* 12 個) は整備済みなのに、
 *   実際の CSS は 85-89% が生値で書かれていた (border-radius 15/106、font-size 20/172、
 *   padding 20/126)。同じ「0.5rem の角丸」を 48 箇所で個別に書いている状態で、
 *   スケールを変えたくても一括で変えられない。値を token に寄せて管理点を 1 つにする。
 *
 *   対象は寸法 (radius / font-size / space) のみ。色は semantic theme token として
 *   既に統一済みで、theme-perfect が生値の混入を検出する。
 *
 * 安全性:
 *   token と完全に同値のものだけを置換する (0.5rem → var(--radius-md))。
 *   スケールに無い値 (0.6rem 等) は「そのスケールを使っていない意図がある」可能性が
 *   あるので触らない。よって描画は 1px も変わらない。
 *
 * 使い方:
 *   node scripts/tokenize-dimensions.js --dry-run   # 何を置換するか確認
 *   node scripts/tokenize-dimensions.js             # 実際に書き換える
 *   node scripts/tokenize-dimensions.js --check     # 生値が残っていれば exit 1 (CI 用)
 */
const fs = require('fs');
const path = require('path');
const { listSourceHtml } = require('./lib/src-files');
const { scaleMap } = require('./lib/theme-core');

const SRC = path.resolve(__dirname, '../src');
const DRY = process.argv.includes('--dry-run');
const CHECK = process.argv.includes('--check');

// 生成物と vendor の判定は lib/src-files が持つ (名前ではなく中身の【自動生成】で判定)

// 対象プロパティ → 使う変換表。接頭辞から token 名を組み立てる。
const RULES = [
  { props: ['border-radius'], prefix: 'radius' },
  { props: ['font-size'], prefix: 'font-size' },
  {
    props: ['gap', 'row-gap', 'column-gap', 'padding', 'padding-top', 'padding-right',
      'padding-bottom', 'padding-left', 'margin-top', 'margin-right', 'margin-bottom', 'margin-left'],
    prefix: 'space',
  },
].map((r) => ({
  ...r,
  map: scaleMap(SRC, r.prefix),
  // 正規表現は起動時に 1 回だけ組む (以前はファイル x セグメント x ルールで再生成していた)
  re: new RegExp(`(^|[;{\\s])(${r.props.join('|')})\\s*:\\s*([^;{}!]+)(;|\\})`, 'g'),
}));

const findings = [];

function convertDeclValue(raw, map, prefix) {
  // 複数値 (padding: 0.5rem 1rem) は 1 つずつ見る。すべて変換できる場合のみ置換する
  //   (一部だけ token になると、かえって読みにくくなるため)。
  const parts = raw.trim().split(/\s+/);
  if (!parts.length || parts.length > 4) return null;
  const out = parts.map((p) => (Object.prototype.hasOwnProperty.call(map, p) ? `var(--${prefix}-${map[p]})` : null));
  if (out.some((x) => x === null)) return null;
  return out.join(' ');
}

function processFile(file) {
  const full = path.join(SRC, file);
  const src = fs.readFileSync(full, 'utf8');
  let changed = 0;

  // <style> ブロック内だけを対象にする (HTML 属性や JS 文字列は触らない)
  const out = src.replace(/<style>([\s\S]*?)<\/style>/g, (whole, css) => {
    // コメント内は書き換えない (説明文の数値を壊さないため)
    const segments = css.split(/(\/\*[\s\S]*?\*\/)/);
    const converted = segments.map((seg) => {
      if (seg.startsWith('/*')) return seg;
      let s = seg;
      for (const { map, prefix, re } of RULES) {
        s = s.replace(re, (m, pre, prop, val, post) => {
          const conv = convertDeclValue(val, map, prefix);
          if (!conv) return m;
          changed += 1;
          findings.push(`${file}: ${prop}: ${val.trim()} → ${conv}`);
          return `${pre}${prop}: ${conv}${post}`;
        });
      }
      return s;
    }).join('');
    return `<style>${converted}</style>`;
  });

  if (changed && !DRY && !CHECK) fs.writeFileSync(full, out);
  return changed;
}

const files = listSourceHtml(SRC);
let total = 0;
for (const f of files) total += processFile(f);

if (CHECK) {
  if (total > 0) {
    console.error(`✗ 設計スケールと同値の生値が ${total} 箇所残っています。`);
    console.error('  node scripts/tokenize-dimensions.js で token に寄せてください。');
    findings.slice(0, 12).forEach((x) => console.error(`    ${x}`));
    if (findings.length > 12) console.error(`    … 他 ${findings.length - 12} 件`);
    process.exit(1);
  }
  console.log('✓ 寸法はすべて設計スケール token 経由です。');
} else if (DRY) {
  console.log(`置換候補: ${total} 箇所`);
  findings.slice(0, 30).forEach((x) => console.log(`  ${x}`));
  if (findings.length > 30) console.log(`  … 他 ${findings.length - 30} 件`);
} else {
  console.log(`✓ ${total} 箇所を token に置換しました (${files.length} ファイル走査)`);
}
