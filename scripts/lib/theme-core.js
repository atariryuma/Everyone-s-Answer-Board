'use strict';

/**
 * theme-core.js — theme-*.js CLI 群で重複していた色/CSS token ユーティリティの
 * single source of truth。
 *
 * 収録するのは「複数ツールで完全同一 (または正常入力で出力同一)」の関数のみ。
 * 実装が本質的に異なるもの (theme-contrast の生 contrastRatio と theme-matrix の
 * 丸め込み contrast、theme-contrast の部分一致 resolveToken と theme-matrix の
 * 完全一致 resolveVar) は各ツール固有の挙動なので **ここには入れない** —
 * theme:perfect が子ツールの stdout を parse するため、出力を 1 byte も変えない。
 *
 * CommonJS (scripts/ は package.json の type 指定なし = CJS)。
 */

// hex (#rgb / #rrggbb) / rgba() / rgb() → {r,g,b,a} | null。
// 不正な hex は null を返す (theme-matrix.js 版に統一。theme-contrast.js の旧版は
//   妥当性チェックが無く NaN を返していたが、正常な token 入力では両者同一出力)。
function parseColor(value) {
  if (!value) return null;
  value = String(value).trim();
  if (value.startsWith('#')) {
    const hex = value.slice(1);
    const expanded = hex.length === 3
      ? hex.split('').map(c => c + c).join('')
      : hex.length === 6 ? hex : hex.slice(0, 6);
    if (!/^[0-9a-fA-F]{6}$/.test(expanded)) return null;
    return {
      r: parseInt(expanded.slice(0, 2), 16),
      g: parseInt(expanded.slice(2, 4), 16),
      b: parseInt(expanded.slice(4, 6), 16),
      a: 1,
    };
  }
  if (value.startsWith('rgba(') || value.startsWith('rgb(')) {
    const m = value.match(/rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*(?:,\s*([\d.]+)\s*)?\)/);
    if (!m) return null;
    return {
      r: Number(m[1]),
      g: Number(m[2]),
      b: Number(m[3]),
      a: m[4] !== undefined ? Number(m[4]) : 1,
    };
  }
  return null;
}

// alpha 合成: foreground (with alpha) を background に重ねた最終色。両ツール完全同一。
function blendOnBg(fg, bg) {
  if (!fg || !bg) return fg;
  if (fg.a >= 1) return fg;
  const a = fg.a;
  return {
    r: Math.round(fg.r * a + bg.r * (1 - a)),
    g: Math.round(fg.g * a + bg.g * (1 - a)),
    b: Math.round(fg.b * a + bg.b * (1 - a)),
    a: 1,
  };
}

// WCAG 2.1 相対輝度。両ツール数式同一 (theme-matrix は relLum 名で同一実装)。
function relativeLuminance({ r, g, b }) {
  const toLinear = (c) => {
    const v = c / 255;
    return v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4);
  };
  return 0.2126 * toLinear(r) + 0.7152 * toLinear(g) + 0.0722 * toLinear(b);
}

// CSS ブロック本体を depth カウントで抽出 (:root { ... } の中身)。両ツール完全同一。
function extractBlockBody(css, startRegex) {
  const startMatch = css.match(startRegex);
  if (!startMatch) return '';
  const startIdx = startMatch.index + startMatch[0].lastIndexOf('{');
  let depth = 1;
  let i = startIdx + 1;
  while (i < css.length && depth > 0) {
    if (css[i] === '{') depth++;
    else if (css[i] === '}') depth--;
    i++;
  }
  return css.slice(startIdx + 1, i - 1);
}

// CSS ブロック本体から --token: value を抽出。両ツール同一 regex。
function extractTokensFromBody(body) {
  const tokens = {};
  if (!body) return tokens;
  const re = /--([a-z0-9-]+)\s*:\s*([^;]+);/gi;
  let m;
  while ((m = re.exec(body)) !== null) {
    tokens[`--${m[1]}`] = m[2].trim();
  }
  return tokens;
}

module.exports = {
  parseColor,
  blendOnBg,
  relativeLuminance,
  extractBlockBody,
  extractTokensFromBody,
};
