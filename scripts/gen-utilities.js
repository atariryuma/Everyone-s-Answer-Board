#!/usr/bin/env node
/**
 * gen-utilities.js — src/ で実際に使われている Tailwind クラスだけを静的 CSS に書き出す。
 *
 * Why: 従来は cdn.tailwindcss.com (Tailwind Play CDN) を全ページで読み込んでいた。
 *   これはブラウザ上で JIT コンパイルする開発用ビルドで、Tailwind 公式が本番利用を
 *   明示的に非推奨としている (コード側もその警告を文字列マッチで握り潰していた)。
 *   学校ネットワークが CDN を遮断すると全画面が崩壊し、フォールバックも無かった。
 *   実測すると JIT 固有機能はほぼ使っておらず (任意値 7 / md: 9 / sm: 1 / dark: 0)、
 *   静的な utility CSS で完全に置き換えられる。
 *
 * 方針:
 *   - src/**.html から class 属性を走査し、Tailwind 由来のクラスだけを抽出
 *   - Tailwind v3 の既定値と同じ宣言を生成する (見た目を変えないため)
 *   - 自前 CSS に既に定義があるクラス名は生成しない (二重定義を避ける)
 *   - 出力は src/UtilityStyles.css.html。SharedPageHead が UnifiedStyles の後に読む
 *     (Tailwind と同じく utility が primitive を上書きする順序)
 *
 * 使い方:
 *   npm run gen:utilities        # 生成して上書き
 *   npm run gen:utilities -- --check   # 差分があれば exit 1 (CI 用)
 */

'use strict';
const fs = require('fs');
const path = require('path');
const { listSourceHtml } = require('./lib/src-files');

const SRC = path.resolve(__dirname, '../src');
const OUT = path.join(SRC, 'UtilityStyles.css.html');
const CHECK = process.argv.includes('--check');

// 生成物 / vendor の判定は lib/src-files が持つ。生成物を走査すると「生成物が入力に
// 混ざる」フィードバックループになり、生成のたびに結果が育つ (v2910 で実際に起きた)。
const htmlFiles = listSourceHtml(SRC);

// ── 1. 使用クラスの収集 ────────────────────────────────────────────
const used = new Set();
for (const f of htmlFiles) {
  const s = fs.readFileSync(path.join(SRC, f), 'utf8');
  const add = (str) => str.split(/\s+/).forEach((t) => { if (t) used.add(t); });
  for (const m of s.matchAll(/class="([^"]*)"/g)) add(m[1]);
  for (const m of s.matchAll(/class='([^']*)'/g)) add(m[1]);
  for (const m of s.matchAll(/className\s*=\s*['"]([^'"]*)['"]/g)) add(m[1]);
  for (const m of s.matchAll(/classList\.(?:add|remove|toggle)\(([^)]*)\)/g)) {
    for (const q of m[1].matchAll(/['"]([A-Za-z0-9:_/\[\].%-]+)['"]/g)) used.add(q[1]);
  }
}

// ── 1b. 動的に組み立てられたクラスの検出 ──────────────────────────
//   `'lg:grid-cols-' + value` のような文字列連結は、静的走査では実際の
//   クラス名が分からない。Tailwind CDN の JIT は実行時にコンパイルしていたので
//   動いていたが、静的 CSS では「生成されないので効かない」無言の不具合になる
//   (実例: 列数スライダーが 5・6 で効かなかった)。ここで検出して警告する。
const dynamicSuspects = [];
for (const f of htmlFiles) {
  const s = fs.readFileSync(path.join(SRC, f), 'utf8');
  // 'xxx-' + var  /  `xxx-${...}`  の形で、xxx が Tailwind っぽい接頭辞のもの
  const re = /['"`]([a-z]+(?::[a-z-]+)?-)['"`]?\s*(?:\+|\$\{)/g;
  for (const m of s.matchAll(re)) {
    const prefix = m[1];
    if (!/(grid-cols|col-span|gap|p|px|py|m|mx|my|w|h|text|bg|border|rounded|z|opacity|translate|scale)-$/.test(prefix)) continue;
    dynamicSuspects.push(`${f}:${(s.slice(0, m.index).match(/\n/g) || []).length + 1} '${prefix}' + …`);
  }
}

// ── 2. 自前 CSS に既にある名前は除外 ────────────────────────────────
//   注意: 「コメント内でクラス名に言及しているだけ」を実定義と誤認しないこと。
//   CSS コメントを除去し、宣言ブロック `{...}` の外 (= セレクタ部) だけを走査する。
let ownCss = '';
for (const f of htmlFiles) {
  const s = fs.readFileSync(path.join(SRC, f), 'utf8');
  for (const m of s.matchAll(/<style>([\s\S]*?)<\/style>/g)) ownCss += m[1] + '\n';
}
const ownClasses = new Set();
{
  const noComment = ownCss.replace(/\/\*[\s\S]*?\*\//g, ' ');
  // 「自前で定義している」とみなすのは、そのクラス単独が主体のセレクタだけ
  //   (.card / .card:hover / .card::before)。
  //
  // Why: 以前は複合セレクタに出てくる名前も「定義済み」として除外していた。
  //   すると `#boardInfoFooter .flex.items-center.gap-2.flex-shrink-0 { … }` のような
  //   「utility を狙い撃ちする」ルールを 1 本書いただけで、.flex / .items-center /
  //   .gap-2 / .flex-shrink-0 がアプリ全体から生成されなくなる。
  //   実際 .hidden が `body .loading-overlay:not(.hidden)` のせいで生成されず、
  //   hidden 指定の 19 要素中 18 個が画面に出たままになっていた。
  //   複合セレクタは「既存 utility に上乗せする」意図であって定義ではない。
  for (const m of noComment.matchAll(/(^|[};])([^{};]+)\{/g)) {
    for (const part of m[2].split(',')) {
      const solo = part.trim().match(/^\.([A-Za-z][A-Za-z0-9_-]*)(?:::?[a-zA-Z-]+(?:\([^)]*\))?)*$/);
      if (solo) ownClasses.add(solo[1]);
    }
  }
}

// ── 3. Tailwind v3 の既定スケール ──────────────────────────────────
const SPACE = {
  '0': '0px', '0.5': '0.125rem', '1': '0.25rem', '1.5': '0.375rem', '2': '0.5rem',
  '2.5': '0.625rem', '3': '0.75rem', '3.5': '0.875rem', '4': '1rem', '5': '1.25rem',
  '6': '1.5rem', '7': '1.75rem', '8': '2rem', '9': '2.25rem', '10': '2.5rem',
  '11': '2.75rem', '12': '3rem', '14': '3.5rem', '16': '4rem', '20': '5rem',
  '24': '6rem', '28': '7rem', '32': '8rem', '40': '10rem', '48': '12rem',
  '56': '14rem', '64': '16rem', '80': '20rem', '96': '24rem', 'px': '1px',
  'auto': 'auto', 'full': '100%', 'screen': '100vw', 'min': 'min-content',
  'max': 'max-content', 'fit': 'fit-content',
  '1/2': '50%', '1/3': '33.333333%', '2/3': '66.666667%', '1/4': '25%',
  '3/4': '75%', '1/5': '20%'
};
const FONT_SIZE = {
  xs: ['0.75rem', '1rem'], sm: ['0.875rem', '1.25rem'], base: ['1rem', '1.5rem'],
  lg: ['1.125rem', '1.75rem'], xl: ['1.25rem', '1.75rem'], '2xl': ['1.5rem', '2rem'],
  '3xl': ['1.875rem', '2.25rem'], '4xl': ['2.25rem', '2.5rem'], '5xl': ['3rem', '1'],
  '6xl': ['3.75rem', '1']
};
const FONT_WEIGHT = { thin: 100, light: 300, normal: 400, medium: 500, semibold: 600, bold: 700, extrabold: 800 };
const RADIUS = {
  none: '0px', sm: '0.125rem', '': '0.25rem', md: '0.375rem', lg: '0.5rem',
  xl: '0.75rem', '2xl': '1rem', '3xl': '1.5rem', full: '9999px'
};
const SHADOW = {
  sm: '0 1px 2px 0 rgb(0 0 0 / 0.05)',
  '': '0 1px 3px 0 rgb(0 0 0 / 0.1), 0 1px 2px -1px rgb(0 0 0 / 0.1)',
  md: '0 4px 6px -1px rgb(0 0 0 / 0.1), 0 2px 4px -2px rgb(0 0 0 / 0.1)',
  lg: '0 10px 15px -3px rgb(0 0 0 / 0.1), 0 4px 6px -4px rgb(0 0 0 / 0.1)',
  xl: '0 20px 25px -5px rgb(0 0 0 / 0.1), 0 8px 10px -6px rgb(0 0 0 / 0.1)',
  '2xl': '0 25px 50px -12px rgb(0 0 0 / 0.25)',
  none: '0 0 #0000'
};
const LEADING = {
  none: '1', tight: '1.25', snug: '1.375', normal: '1.5', relaxed: '1.625', loose: '2'
};
// Tailwind の既定パレット。src で実際に使われているものだけ。
const PALETTE = {
  white: '#ffffff', black: '#000000', transparent: 'transparent', current: 'currentColor',
  'blue-400': '#60a5fa', 'blue-600': '#2563eb', 'blue-700': '#1d4ed8',
  'cyan-200': '#a5f3fc', 'cyan-400': '#22d3ee', 'cyan-600': '#0891b2', 'cyan-700': '#0e7490',
  'green-400': '#4ade80', 'green-500': '#22c55e', 'green-600': '#16a34a', 'green-700': '#15803d',
  'emerald-600': '#059669', 'emerald-700': '#047857',
  'orange-500': '#f97316', 'orange-600': '#ea580c', 'orange-700': '#c2410c',
  'purple-600': '#9333ea', 'purple-700': '#7e22ce',
  'red-400': '#f87171', 'red-500': '#ef4444', 'red-600': '#dc2626', 'red-700': '#b91c1c',
  'slate-500': '#64748b', 'slate-600': '#475569',
  'yellow-400': '#facc15', 'yellow-500': '#eab308'
};
// プロジェクト独自のテーマ色 (SharedTailwindConfig の theme.extend.colors と同値)
const THEME_COLORS = JSON.parse(fs.readFileSync(path.join(__dirname, 'theme-colors.json'), 'utf8'));

// ── 4. クラス名 → CSS 宣言 ─────────────────────────────────────────
function colorValue(name) {
  if (PALETTE[name]) return PALETTE[name];
  // theme- / brand- / status- プレフィックス付きのカスタム色
  for (const prefix of ['theme-', 'brand-', 'status-']) {
    if (name.startsWith(prefix)) {
      const key = name.slice(prefix.length);
      if (THEME_COLORS[key]) return THEME_COLORS[key];
    }
  }
  if (name === 'theme') return THEME_COLORS.DEFAULT;
  // `red-500/10` のような alpha 記法
  const am = name.match(/^(.+)\/(\d+)$/);
  if (am) {
    const base = colorValue(am[1]);
    if (base) return `color-mix(in srgb, ${base} ${am[2]}%, transparent)`;
  }
  return null;
}

// #rrggbb を "r g b" に。bg-opacity-* と組み合わせるため Tailwind と同じ形にする。
function rgbChannels(hex) {
  const m = /^#([0-9a-f]{6})$/i.exec(hex);
  if (!m) return null;
  const n = parseInt(m[1], 16);
  return `${(n >> 16) & 255} ${(n >> 8) & 255} ${n & 255}`;
}

function arbitrary(v) {
  return v.replace(/^\[/, '').replace(/\]$/, '').replace(/_/g, ' ');
}

function declFor(cls) {
  const d = (k, v) => ({ [k]: v });
  // 任意値
  const arb = cls.match(/^([a-z-]+)-(\[.+\])$/);
  if (arb) {
    const [, prop, val] = arb;
    const v = arbitrary(val);
    const map = {
      z: 'z-index', 'max-h': 'max-height', 'min-h': 'min-height',
      'max-w': 'max-width', 'min-w': 'min-width', w: 'width', h: 'height',
      p: 'padding', m: 'margin', top: 'top', left: 'left', right: 'right', bottom: 'bottom'
    };
    if (map[prop]) return d(map[prop], v);
    return null;
  }

  // display / position / misc（値を持たない単独クラス）
  const SIMPLE = {
    block: d('display', 'block'), 'inline-block': d('display', 'inline-block'),
    inline: d('display', 'inline'), flex: d('display', 'flex'),
    'inline-flex': d('display', 'inline-flex'), grid: d('display', 'grid'),
    hidden: d('display', 'none'), table: d('display', 'table'),
    static: d('position', 'static'), fixed: d('position', 'fixed'),
    absolute: d('position', 'absolute'), relative: d('position', 'relative'),
    sticky: d('position', 'sticky'),
    'flex-row': d('flex-direction', 'row'), 'flex-col': d('flex-direction', 'column'),
    'flex-wrap': d('flex-wrap', 'wrap'), 'flex-nowrap': d('flex-wrap', 'nowrap'),
    'flex-1': d('flex', '1 1 0%'), 'flex-auto': d('flex', '1 1 auto'),
    'flex-none': d('flex', 'none'), 'flex-grow': d('flex-grow', '1'),
    'flex-shrink-0': d('flex-shrink', '0'), 'shrink-0': d('flex-shrink', '0'),
    'grow-0': d('flex-grow', '0'),
    'items-start': d('align-items', 'flex-start'), 'items-center': d('align-items', 'center'),
    'items-end': d('align-items', 'flex-end'), 'items-stretch': d('align-items', 'stretch'),
    'items-baseline': d('align-items', 'baseline'),
    'justify-start': d('justify-content', 'flex-start'),
    'justify-center': d('justify-content', 'center'),
    'justify-end': d('justify-content', 'flex-end'),
    'justify-between': d('justify-content', 'space-between'),
    'justify-around': d('justify-content', 'space-around'),
    'self-start': d('align-self', 'flex-start'), 'self-center': d('align-self', 'center'),
    'text-left': d('text-align', 'left'), 'text-center': d('text-align', 'center'),
    'text-right': d('text-align', 'right'),
    truncate: { overflow: 'hidden', 'text-overflow': 'ellipsis', 'white-space': 'nowrap' },
    italic: d('font-style', 'italic'), 'not-italic': d('font-style', 'normal'),
    underline: d('text-decoration-line', 'underline'),
    'no-underline': d('text-decoration-line', 'none'),
    'font-mono': d('font-family', 'ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace'),
    'whitespace-nowrap': d('white-space', 'nowrap'),
    'whitespace-pre-wrap': d('white-space', 'pre-wrap'),
    'break-all': d('word-break', 'break-all'),
    'break-words': d('overflow-wrap', 'break-word'),
    'cursor-pointer': d('cursor', 'pointer'),
    'cursor-not-allowed': d('cursor', 'not-allowed'),
    'select-none': d('user-select', 'none'),
    'pointer-events-none': d('pointer-events', 'none'),
    'pointer-events-auto': d('pointer-events', 'auto'),
    'appearance-none': d('appearance', 'none'),
    'resize-none': d('resize', 'none'),
    'overflow-hidden': d('overflow', 'hidden'), 'overflow-auto': d('overflow', 'auto'),
    'overflow-visible': d('overflow', 'visible'),
    'overflow-x-auto': d('overflow-x', 'auto'), 'overflow-y-auto': d('overflow-y', 'auto'),
    'overflow-x-hidden': d('overflow-x', 'hidden'), 'overflow-y-hidden': d('overflow-y', 'hidden'),
    'list-disc': d('list-style-type', 'disc'), 'list-decimal': d('list-style-type', 'decimal'),
    'list-inside': d('list-style-position', 'inside'),
    'mx-auto': { 'margin-left': 'auto', 'margin-right': 'auto' },
    'ml-auto': d('margin-left', 'auto'), 'mr-auto': d('margin-right', 'auto'),
    'inset-0': { top: '0px', right: '0px', bottom: '0px', left: '0px' },
    'min-h-screen': d('min-height', '100vh'),
    'h-screen': d('height', '100vh'), 'w-screen': d('width', '100vw'),
    'col-span-full': d('grid-column', '1 / -1'),
    'basis-full': d('flex-basis', '100%'), 'basis-auto': d('flex-basis', 'auto'),
    'border-dashed': d('border-style', 'dashed'),
    'border-solid': d('border-style', 'solid'),
    'transition-all': { 'transition-property': 'all', 'transition-timing-function': 'cubic-bezier(0.4, 0, 0.2, 1)', 'transition-duration': '150ms' },
    'transition-colors': { 'transition-property': 'color, background-color, border-color, fill, stroke', 'transition-timing-function': 'cubic-bezier(0.4, 0, 0.2, 1)', 'transition-duration': '150ms' },
    'transition-transform': { 'transition-property': 'transform', 'transition-timing-function': 'cubic-bezier(0.4, 0, 0.2, 1)', 'transition-duration': '150ms' },
    'transition-opacity': { 'transition-property': 'opacity', 'transition-timing-function': 'cubic-bezier(0.4, 0, 0.2, 1)', 'transition-duration': '150ms' },
    'ease-in-out': d('transition-timing-function', 'cubic-bezier(0.4, 0, 0.2, 1)'),
    'ease-out': d('transition-timing-function', 'cubic-bezier(0, 0, 0.2, 1)'),
    'ease-in': d('transition-timing-function', 'cubic-bezier(0.4, 0, 1, 1)'),
    'outline-none': { outline: '2px solid transparent', 'outline-offset': '2px' },
    'animate-spin': d('animation', 'tw-spin 1s linear infinite'),
    'animate-pulse': d('animation', 'tw-pulse 2s cubic-bezier(0.4, 0, 0.6, 1) infinite')
  };
  if (Object.prototype.hasOwnProperty.call(SIMPLE, cls)) return SIMPLE[cls];

  let m;
  // spacing: p/px/py/pt/pr/pb/pl/m/... /gap
  const SPACE_PROP = {
    p: ['padding'], px: ['padding-left', 'padding-right'], py: ['padding-top', 'padding-bottom'],
    pt: ['padding-top'], pr: ['padding-right'], pb: ['padding-bottom'], pl: ['padding-left'],
    m: ['margin'], mx: ['margin-left', 'margin-right'], my: ['margin-top', 'margin-bottom'],
    mt: ['margin-top'], mr: ['margin-right'], mb: ['margin-bottom'], ml: ['margin-left'],
    gap: ['gap'], 'gap-x': ['column-gap'], 'gap-y': ['row-gap'],
    w: ['width'], h: ['height'], 'min-w': ['min-width'], 'min-h': ['min-height'],
    'max-w': ['max-width'], 'max-h': ['max-height'],
    top: ['top'], right: ['right'], bottom: ['bottom'], left: ['left']
  };
  m = cls.match(/^(gap-x|gap-y|min-w|min-h|max-w|max-h|p[xytrbl]?|m[xytrbl]?|gap|w|h|top|right|bottom|left)-(.+)$/);
  if (m && SPACE_PROP[m[1]]) {
    const MAXW = {
      xs: '20rem', sm: '24rem', md: '28rem', lg: '32rem', xl: '36rem', '2xl': '42rem',
      '3xl': '48rem', '4xl': '56rem', '5xl': '64rem', '6xl': '72rem', '7xl': '80rem',
      full: '100%', none: 'none', prose: '65ch'
    };
    let v = null;
    if (m[1] === 'max-w' && MAXW[m[2]]) v = MAXW[m[2]];
    else if (SPACE[m[2]]) v = SPACE[m[2]];
    if (v) {
      const out = {};
      for (const p of SPACE_PROP[m[1]]) out[p] = v;
      return out;
    }
  }
  // space-x / space-y は子要素へのマージン
  m = cls.match(/^space-([xy])-(.+)$/);
  if (m && SPACE[m[2]]) {
    return { __selector: `.${cssEscape(cls)} > * + *`,
      [m[1] === 'x' ? 'margin-left' : 'margin-top']: SPACE[m[2]] };
  }
  // grid-cols-N
  m = cls.match(/^grid-cols-(\d+)$/);
  if (m) return d('grid-template-columns', `repeat(${m[1]}, minmax(0, 1fr))`);
  // text-{size} / text-{color}
  m = cls.match(/^text-(.+)$/);
  if (m) {
    if (FONT_SIZE[m[1]]) {
      return { 'font-size': FONT_SIZE[m[1]][0], 'line-height': FONT_SIZE[m[1]][1] };
    }
    const c = colorValue(m[1]);
    if (c) return d('color', c);
  }
  // font-{weight}
  m = cls.match(/^font-(.+)$/);
  if (m && FONT_WEIGHT[m[1]]) return d('font-weight', String(FONT_WEIGHT[m[1]]));
  // leading-{n}
  m = cls.match(/^leading-(.+)$/);
  if (m && LEADING[m[1]]) return d('line-height', LEADING[m[1]]);
  // bg-{color} / bg-opacity-{n}
  m = cls.match(/^bg-opacity-(\d+)$/);
  if (m) return d('--tw-bg-opacity', String(Number(m[1]) / 100));
  m = cls.match(/^bg-(.+)$/);
  if (m) {
    const c = colorValue(m[1]);
    if (c) {
      // Tailwind は bg-opacity-* と合成できるよう rgb(r g b / var(--tw-bg-opacity)) を出す。
      //   固定 hex のパレット色だけがこの形を取れる (var(--...) は分解できない)。
      const ch = rgbChannels(c);
      return d('background-color', ch ? `rgb(${ch} / var(--tw-bg-opacity, 1))` : c);
    }
  }
  // border 系
  if (cls === 'border') return { 'border-width': '1px', 'border-style': 'solid' };
  m = cls.match(/^border-(\d)$/);
  if (m) return { 'border-width': `${m[1]}px`, 'border-style': 'solid' };
  // Why 片側だけに style を当てるか: border-style は 4 辺まとめて効くショートハンド。
  //   border-t に 'border-style: solid' を付けると、幅を指定していない 3 辺が
  //   既定値 (medium ≒ 1.5px) で描画され、上線のつもりが箱になる。
  //   Tailwind は preflight で全要素に border-width:0 を敷いてこれを回避しているが、
  //   本プロジェクトは CDN 廃止時にその preflight を持たない (下の base 参照)。
  //   divide-y は元から border-top-style で正しく書かれていた。
  m = cls.match(/^border-([trbl])-(\d)$/);
  if (m) {
    const side = { t: 'top', r: 'right', b: 'bottom', l: 'left' }[m[1]];
    return { [`border-${side}-width`]: `${m[2]}px`, [`border-${side}-style`]: 'solid' };
  }
  m = cls.match(/^border-([trbl])$/);
  if (m) {
    const side = { t: 'top', r: 'right', b: 'bottom', l: 'left' }[m[1]];
    return { [`border-${side}-width`]: '1px', [`border-${side}-style`]: 'solid' };
  }
  m = cls.match(/^border-(.+)$/);
  if (m) {
    const c = colorValue(m[1]);
    if (c) return d('border-color', c);
  }
  // divide-y / divide-{color}
  if (cls === 'divide-y') {
    return { __selector: `.${cssEscape(cls)} > * + *`, 'border-top-width': '1px', 'border-top-style': 'solid' };
  }
  m = cls.match(/^divide-(.+)$/);
  if (m) {
    const c = colorValue(m[1]);
    if (c) return { __selector: `.${cssEscape(cls)} > * + *`, 'border-color': c };
  }
  // rounded
  if (cls === 'rounded') return d('border-radius', RADIUS['']);
  m = cls.match(/^rounded-(.+)$/);
  if (m && RADIUS[m[1]]) return d('border-radius', RADIUS[m[1]]);
  // shadow
  if (cls === 'shadow') return d('box-shadow', SHADOW['']);
  m = cls.match(/^shadow-(.+)$/);
  if (m && SHADOW[m[1]]) return d('box-shadow', SHADOW[m[1]]);
  // opacity
  m = cls.match(/^opacity-(\d+)$/);
  if (m) return d('opacity', String(Number(m[1]) / 100));
  // z-index
  m = cls.match(/^z-(\d+)$/);
  if (m) return d('z-index', m[1]);
  // duration
  m = cls.match(/^duration-(\d+)$/);
  if (m) return d('transition-duration', `${m[1]}ms`);
  // ring
  m = cls.match(/^ring-(\d)$/);
  if (m) return d('box-shadow', `0 0 0 ${m[1]}px var(--tw-ring-color, rgb(59 130 246 / 0.5))`);
  m = cls.match(/^ring-(.+)$/);
  if (m) {
    const c = colorValue(m[1]);
    if (c) return d('--tw-ring-color', c);
  }
  m = cls.match(/^ring-offset-(\d)$/);
  if (m) return d('--tw-ring-offset-width', `${m[1]}px`);
  // underline-offset
  m = cls.match(/^underline-offset-(\d+)$/);
  if (m) return d('text-underline-offset', `${m[1]}px`);
  return null;
}

// CSS セレクタで使えるようにエスケープ (`/` `:` `[` `]` `.` `%`)
function cssEscape(cls) {
  return cls.replace(/([:/[\].%])/g, '\\$1');
}

// ── 5. 生成 ────────────────────────────────────────────────────────
const BREAKPOINT = { sm: '640px', md: '768px', lg: '1024px', xl: '1280px', '2xl': '1536px' };
const STATE = { hover: ':hover', focus: ':focus', 'focus-visible': ':focus-visible',
  active: ':active', disabled: ':disabled', 'group-hover': '' };

const base = [];              // prefix 無し
const states = [];            // hover: 等
const media = {};             // sm: md: 等
const skipped = [];

for (const cls of [...used].sort()) {
  // 明らかに Tailwind ではないもの (JS の断片 / 自前クラス) は除外
  if (!/^[a-z0-9:/[\].%_-]+$/i.test(cls)) continue;
  const parts = cls.split(':');
  const bare = parts.pop();
  const prefixes = parts;
  if (ownClasses.has(cls) || (prefixes.length === 0 && ownClasses.has(bare))) continue;
  const decl = declFor(bare);
  if (!decl) { if (!ownClasses.has(bare)) skipped.push(cls); continue; }

  const sel = decl.__selector ? decl.__selector.replace(cssEscape(bare), cssEscape(cls)) : `.${cssEscape(cls)}`;
  // hidden だけは !important を付ける。
  //   Why: ページ固有 CSS (page.css 等) は SharedPageHead より後に読まれるため、
  //   同じ詳細度なら後勝ちで utility が負ける。display を持つ既存クラス
  //   (.eab-identity-pill { display: inline-flex } 等) に .hidden を足しても
  //   隠れず、JS の表示制御が無言で効かなくなる。「隠す」はレイアウト指定に
  //   負けてよい種類の指定ではないので、ここだけ強制する。
  const forceImportant = bare === 'hidden';
  const body = Object.entries(decl).filter(([k]) => k !== '__selector')
    .map(([k, v]) => `${k}: ${v}${forceImportant ? ' !important' : ''};`).join(' ');

  if (prefixes.length === 0) { base.push(`  ${sel} { ${body} }`); continue; }
  const bp = prefixes.find((p) => BREAKPOINT[p]);
  const st = prefixes.filter((p) => STATE[p]).map((p) => STATE[p]).join('');
  const finalSel = sel + st;
  if (bp) {
    (media[bp] ||= []).push(`    ${finalSel} { ${body} }`);
  } else {
    states.push(`  ${finalSel} { ${body} }`);
  }
}

const header = `<!-- =====================================================================
  UtilityStyles — Tailwind 互換の静的 utility CSS (自動生成)

  【自動生成】このファイルは scripts/gen-utilities.js が生成する。直接編集しない。
     クラスを増やしたいときは HTML 側で使ってから \`npm run gen:utilities\` を実行する。

  Why: 以前は cdn.tailwindcss.com (Tailwind Play CDN) をブラウザで読み込み、
    JIT コンパイルしていた。Tailwind 公式が本番非推奨としているビルドであり、
    学校ネットワークが CDN を遮断すると全画面が崩壊する構造だった。
    実測では JIT 固有機能をほぼ使っていなかったため、使用クラスだけを
    静的 CSS として同梱し、外部ネットワーク依存をゼロにする。

  読み込み順序: UnifiedStyles.css の後 (Tailwind と同じく utility が primitive を
    上書きする順序を保つ)。SharedPageHead が管理する。
===================================================================== -->
<style>
  /* ── preflight: box-sizing ──
     Why: Tailwind の preflight が全要素に border-box を当てていたが、CDN 廃止時に
       utility クラスだけを移植して preflight を落としてしまった。既定の content-box では
       w-full (width:100%) と px-* を併用した要素が padding 分だけ親からはみ出す。
       実際 #main-container (w-full + md:px-6) が desktop で 48px、mobile で 32px
       溢れ、ボードの右端カードが切れて横スクロールが出ていた。
       他の utility と違い「使われているクラス」から検出できない類なので、常に出力する。 */
  *, ::before, ::after { box-sizing: border-box; }
  /* 画像がカードや親要素を突き破らないようにする (教材画像は任意サイズで入ってくる) */
  img, video { max-width: 100%; height: auto; }
  /* 押せるものはカーソルで分かるようにする */
  button, [role="button"] { cursor: pointer; }
  /* ボタンの既定外観を消す。
     Why: 背景を明示していないボタンは UA 既定の灰色 (rgb(107,107,107)) と
     outset の立体枠になる。実際 回答カードの削除ボタンが「灰色の丸」として
     本文の上に乗り、90 年代のフォーム部品のように見えていた。
     色や枠は各 primitive (.btn / .reaction-btn 等) が与える。 */
  button, input[type="button"], input[type="submit"], input[type="reset"] {
    background-color: transparent;
    background-image: none;
    border-style: solid;
  }
  table { border-collapse: collapse; }
  /* hidden 属性はブラウザ既定の [hidden]{display:none} がクラス指定に負ける。
     .class-chip-row{display:flex} の要素に hidden を付けても消えず、JS の
     表示制御が無言で効かなくなっていた。.hidden クラスと同じ扱いに揃える。 */
  [hidden] { display: none !important; }

  /* ── keyframes (animate-spin / animate-pulse) ── */
  @keyframes tw-spin { to { transform: rotate(360deg); } }
  @keyframes tw-pulse { 50% { opacity: 0.5; } }

  /* ── base utilities ── */
`;

let out = header + base.join('\n') + '\n';
if (states.length) out += '\n  /* ── state variants (hover / focus / disabled) ── */\n' + states.join('\n') + '\n';
for (const bp of ['sm', 'md', 'lg', 'xl', '2xl']) {
  if (!media[bp]) continue;
  out += `\n  /* ── ${bp} (>= ${BREAKPOINT[bp]}) ── */\n  @media (min-width: ${BREAKPOINT[bp]}) {\n${media[bp].join('\n')}\n  }\n`;
}
out += '</style>\n';

const prev = fs.existsSync(OUT) ? fs.readFileSync(OUT, 'utf8') : '';
if (CHECK) {
  if (prev !== out) {
    console.error('✗ UtilityStyles.css.html が最新ではありません。`npm run gen:utilities` を実行してください。');
    process.exit(1);
  }
  console.log('✓ UtilityStyles.css.html は最新です。');
} else {
  fs.writeFileSync(OUT, out);
  console.log(`✓ ${path.relative(process.cwd(), OUT)} を生成`);
  console.log(`  base ${base.length} / state ${states.length} / media ${Object.values(media).flat().length} ルール`);
  if (dynamicSuspects.length) {
    console.log(`\n  ⚠ 動的に組み立てられている疑いのあるクラス ${dynamicSuspects.length} 件:`);
    for (const d of dynamicSuspects) console.log(`      ${d}`);
    console.log('      静的 CSS では実行時に作られるクラスは効きません。CSS 変数などに置き換えてください。');
  }
  if (skipped.length) {
    console.log(`\n  未対応 (Tailwind でも無効か、自前 CSS 側の名前) ${skipped.length} 件:`);
    console.log('   ', skipped.join(' '));
  }
}
