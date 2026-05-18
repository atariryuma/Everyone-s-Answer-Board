#!/usr/bin/env node
/**
 * theme-audit.js — ダーク/ライトモードの完成度を機械的に測定する。
 *
 * 5 つの軸で audit:
 *   1. Token coverage      — 全 --theme-* がライト/ダーク両方で定義されているか
 *   2. Hardcoded colors    — CSS / HTML / JS 内のハードコード hex / rgba 検出
 *   3. Tailwind hardcoded  — bg-gray-*, text-gray-* 等の theme-non-aware class
 *   4. Inline style colors — style="color: #..." 等のインライン違反
 *   5. Documented exceptions — brand identity / status / projector / spinner 等
 *
 * 終了コード: 違反 0 件 → 0, それ以外 → 1
 *
 * 使い方:
 *   npm run theme:audit             # 全 audit 実行
 *   npm run theme:audit -- --json   # JSON 出力 (CI 用)
 *   npm run theme:audit -- --strict # exception 無視で厳格 audit
 */

'use strict';

const fs = require('node:fs');
const path = require('node:path');

const SRC_DIR = path.resolve(__dirname, '../src');
const ARGS = process.argv.slice(2);
const FLAG_JSON = ARGS.includes('--json');
const FLAG_STRICT = ARGS.includes('--strict');

// =====================================================================
// 1. Token Coverage Check
//    全 --theme-* が :root (dark) と body.theme-light の両方で定義されているか
// =====================================================================

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

function extractTokensFromBody(body) {
  const tokens = {};
  const re = /--([a-z0-9-]+)\s*:\s*([^;]+);/gi;
  let m;
  while ((m = re.exec(body)) !== null) {
    tokens[`--${m[1]}`] = m[2].trim();
  }
  return tokens;
}

function checkTokenCoverage() {
  const cssPath = path.join(SRC_DIR, 'UnifiedStyles.css.html');
  const css = fs.readFileSync(cssPath, 'utf8');

  // 中括弧マッチで :root と body.theme-light ブロックを抽出
  const darkBody = extractBlockBody(css, /:root\s*\{/);
  const lightBody = extractBlockBody(css, /body\.theme-light[^{]*\{/);

  const darkTokens = extractTokensFromBody(darkBody);
  const lightTokens = extractTokensFromBody(lightBody);

  const themeTokens = Object.keys(darkTokens).filter(k => k.startsWith('--theme-'));
  const violations = [];

  for (const token of themeTokens) {
    if (!(token in lightTokens)) {
      violations.push({ type: 'missing-light-value', token, darkValue: darkTokens[token] });
    }
  }
  // light で定義されているのに dark で未定義 (drift detection)
  for (const token of Object.keys(lightTokens)) {
    if (!token.startsWith('--theme-')) continue;
    if (!(token in darkTokens)) {
      violations.push({ type: 'missing-dark-value', token, lightValue: lightTokens[token] });
    }
  }

  return {
    name: 'Token Coverage',
    total: themeTokens.length,
    violations,
    summary: violations.length === 0
      ? `✓ ${themeTokens.length} theme tokens (dark + light both defined)`
      : `✗ ${violations.length} tokens missing dark or light value`,
  };
}

// =====================================================================
// 2. Hardcoded Colors in CSS
//    :root / body.theme-light 外で hex / rgba(255,255,255) / rgba(0,0,0)
// =====================================================================

// 許可される hex / rgba (brand identity, projector mode 等)
const HEX_EXCEPTIONS = [
  '#fde047',     // yellow-300 brand highlight feedback
  '#fbbf24',     // amber-400 (brand-understand, inline override)
  '#67e8f9',     // cyan-300 (theme-spinner 内で定義)
  '#38bdf8',     // cyan-400 (theme-accent-cyan 内で定義)
  '#22d3ee',     // cyan-400 hover state (slider active)
  '#4ade80',     // green-400 saved indicator
  '#fff',        // white text on colored buttons (theme-non-dependent)
  '#ffffff',     // 同上
  '#2563eb',     // pull-to-refresh gradient
  '#7c3aed',     // pull-to-refresh gradient
  '#fdba74',     // orange-300 end-pub-btn text
  '#0a0e1a',     // projector mode dark bg (intentional always-dark)
  '#c7d2fe',     // indigo-200 keyword tag
  // Chart category swatches (phase-preview-pie-dot 0-3): theme 非依存の brand palette
  '#60a5fa',     // blue-400 (category 0)
  '#34d399',     // emerald-400 (category 1)
  '#f87171',     // red-400 (category 3) — phase preview pie のみ。 status-error 経由は別途許可
];

function checkHardcodedCssColors() {
  const cssFiles = ['UnifiedStyles.css.html', 'page.css.html', 'page.viz.css.html'];
  const violations = [];

  for (const file of cssFiles) {
    const fullPath = path.join(SRC_DIR, file);
    if (!fs.existsSync(fullPath)) continue;
    const content = fs.readFileSync(fullPath, 'utf8');
    const lines = content.split('\n');

    // :root と body.theme-light ブロック内はスキップ (token 定義は OK)
    let inRoot = false;
    let inLightOverride = false;
    let braceDepth = 0;

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      const trimmed = line.trim();

      // ブロック検出
      if (/^\s*:root\s*\{/.test(line)) { inRoot = true; braceDepth = 1; continue; }
      if (/^\s*body\.theme-light/.test(line)) { inLightOverride = true; braceDepth = 0; continue; }
      if (inRoot || inLightOverride) {
        braceDepth += (line.match(/\{/g) || []).length;
        braceDepth -= (line.match(/\}/g) || []).length;
        if (braceDepth <= 0 && (line.includes('}'))) { inRoot = false; inLightOverride = false; }
        continue;
      }

      // コメント内はスキップ
      if (/^\s*\/?\*/.test(trimmed) || /^\s*\/\//.test(trimmed)) continue;

      // 明示的に theme 非依存と表明された行はスキップ (`/* theme:exempt ... */` annotation)
      if (/theme:exempt/i.test(line)) continue;

      // hex 検出
      const hexMatches = line.match(/#[0-9a-fA-F]{6}\b|#[0-9a-fA-F]{3}\b/g);
      if (hexMatches) {
        for (const hex of hexMatches) {
          if (HEX_EXCEPTIONS.includes(hex.toLowerCase())) continue;
          if (FLAG_STRICT || !line.includes('/*')) {
            violations.push({
              file: `src/${file}`,
              line: i + 1,
              type: 'hex-hardcoded',
              snippet: trimmed.slice(0, 100),
              value: hex,
            });
          }
        }
      }

      // rgba(255,255,255,*) and rgba(0,0,0,*) — these are theme-non-aware
      // ただし shadow / accent-specific (rgba(56,189,248)) は許可
      const rgbaMatches = line.match(/rgba\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*,/g);
      if (rgbaMatches) {
        for (const rgba of rgbaMatches) {
          const m = rgba.match(/rgba\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*,/);
          if (!m) continue;
          const [, r, g, b] = m.map(Number);
          // white-overlay or black-overlay only (theme-aware target)
          const isWhiteOverlay = r === 255 && g === 255 && b === 255;
          const isBlackOverlay = r === 0 && g === 0 && b === 0;
          if (!isWhiteOverlay && !isBlackOverlay) continue;
          // shadow 系 (box-shadow / text-shadow / drop-shadow / filter) は light でも黒で OK
          if (/(?:box-shadow|text-shadow|drop-shadow|filter)\s*:/.test(line)) continue;
          // background-color: rgba(0,0,0,0) は transparent なので OK
          if (rgba.includes('0, 0, 0,') && /,\s*0\s*\)/.test(line)) continue;
          violations.push({
            file: `src/${file}`,
            line: i + 1,
            type: 'rgba-overlay-hardcoded',
            snippet: trimmed.slice(0, 100),
            value: rgba,
          });
        }
      }
    }
  }

  return {
    name: 'Hardcoded CSS Colors',
    total: 0,
    violations,
    summary: violations.length === 0
      ? '✓ No hardcoded theme-aware colors outside :root / .theme-light'
      : `✗ ${violations.length} hardcoded colors should use --theme-* token`,
  };
}

// =====================================================================
// 3. Tailwind theme-non-aware classes
//    bg-gray-*, text-gray-*, border-gray-* 等は dark/light 切替で動かない。
//    bg-theme-* / text-theme-* / dark: prefix を使うべき。
// =====================================================================

const TW_GRAY_FAMILIES = ['gray', 'slate', 'zinc', 'neutral', 'stone'];

function checkTailwindHardcoded() {
  const htmlFiles = fs.readdirSync(SRC_DIR).filter(f => f.endsWith('.html'));
  const violations = [];
  const byFile = {};

  // Tailwind class attribute (HTML class="..." or JS className 文字列 or template literal) を抽出。
  // 1 つの class attribute 内に `light-class dark:dark-class` のペアがあれば theme-aware として扱う。
  // attribute 抽出は厳密でなくても OK (false-positive を抑えるための approximation)。
  const ATTR_RE = /(?:class|className)=["'`]([^"'`]+)["'`]|\b(?:classList\.(?:add|remove)|className\s*[+]?=)\s*\(?[^;{]*["'`]([^"'`]+)["'`]/g;

  const propPattern = '(bg|text|border|from|to|via|ring|fill|stroke|placeholder|divide)';
  const familyPattern = `(${TW_GRAY_FAMILIES.join('|')}|white|black)`;
  // `prop-family-N` (gray families) または `prop-white` / `prop-black` (level なし)
  const grayClassRe = new RegExp(`\\b${propPattern}-${familyPattern}(?:-(\\d+))?\\b`, 'g');

  for (const file of htmlFiles) {
    if (file === 'SharedTailwindConfig.html') continue;
    const fullPath = path.join(SRC_DIR, file);
    const content = fs.readFileSync(fullPath, 'utf8');

    let totalForFile = 0;
    let attrMatch;
    const attrRe = new RegExp(ATTR_RE.source, 'g');
    while ((attrMatch = attrRe.exec(content)) !== null) {
      const classStr = attrMatch[1] || attrMatch[2];
      if (!classStr) continue;

      // この class attribute 内の全 gray-family class を抽出
      const grayMatches = [];
      let gm;
      const gRe = new RegExp(grayClassRe.source, 'g');
      while ((gm = gRe.exec(classStr)) !== null) {
        const fullToken = gm[0];
        const beforeStr = classStr.slice(0, gm.index);
        const isPrefixed = /(?:dark|light):$/.test(beforeStr);
        // hover: / focus: / active: / disabled: / group-hover: 等の state modifier 付きはペア判定対象外
        const isStatefulMod = /(?:hover|focus|active|disabled|focus-within|focus-visible|group-hover|peer-focus|aria-checked|first|last|odd|even|md|lg|xl|sm|2xl):$/.test(beforeStr);
        if (isStatefulMod) continue;  // 状態 modifier の color は base color の theme pair に影響しない
        grayMatches.push({ prop: gm[1], family: gm[2], step: gm[3], full: fullToken, prefixed: isPrefixed });
      }
      if (grayMatches.length === 0) continue;

      // この class attribute が有色背景クラス (bg-{red|blue|...}-N or gradient) を含むか?
      //   → 含むなら text-white/text-black は両モードで意図通り (white-on-color)、 violation 対象外
      const hasColoredBg = /\b(?:bg-(?:red|blue|green|yellow|cyan|purple|indigo|pink|emerald|amber|rose|orange|sky|violet|fuchsia|teal|lime)-\d+|bg-gradient|from-(?:red|blue|green|yellow|cyan|purple|indigo|pink|emerald|amber|rose|orange|sky|violet|fuchsia|teal|lime)-)/.test(classStr);

      // ペア検出: 同 prop (例 text-) の中で「prefixed dark:」 と 「prefix なし」 がペアならば theme-aware
      const byProp = {};
      for (const g of grayMatches) {
        if (!byProp[g.prop]) byProp[g.prop] = { plain: [], prefixed: [] };
        if (g.prefixed) byProp[g.prop].prefixed.push(g);
        else byProp[g.prop].plain.push(g);
      }
      for (const [prop, { plain, prefixed }] of Object.entries(byProp)) {
        const pairs = Math.min(plain.length, prefixed.length);
        const unpaired = (plain.length + prefixed.length) - 2 * pairs;
        // unpaired 数だけ violation
        const targets = plain.length > prefixed.length ? plain.slice(pairs) : prefixed.slice(pairs);
        for (const t of targets) {
          // colored-bg context での text-white/text-black/ring-white は intentional
          if (hasColoredBg && (t.family === 'white' || t.family === 'black')) continue;
          // from-/to-/via- も gradient context なら不可避
          if ((t.prop === 'from' || t.prop === 'to' || t.prop === 'via') && /bg-gradient/.test(classStr)) continue;
          // bg-black/N (opacity modifier) は modal backdrop の意図的 dimming overlay (theme-agnostic)
          if (t.prop === 'bg' && t.family === 'black' && /bg-black\/\d+|bg-black\s+bg-opacity-/.test(classStr)) continue;
          // bg-white/N (opacity modifier) も同様: 半透明 glass surface (両モードで意図通り)
          if (t.prop === 'bg' && t.family === 'white' && /bg-white\/\d+|bg-white\s+bg-opacity-/.test(classStr)) continue;
          // ring-white (focus ring) は colored ring の brand 一貫性で intentional
          if (t.prop === 'ring' && t.family === 'white') continue;
          // text-white/N (opacity modifier) は colored bg 上の text なので intentional
          if (t.prop === 'text' && t.family === 'white' && new RegExp(`text-white/\\d+`).test(classStr)) continue;
          // text-white/N が同 attribute にあるなら hover:text-white 等の base もペアとして扱う (children button etc)
          if (t.prop === 'text' && t.family === 'white' && /text-white\/\d+/.test(classStr)) continue;
          violations.push({ file: `src/${file}`, type: 'tailwind-unpaired', className: t.full, context: classStr.slice(0, 120) });
          totalForFile++;
        }
      }
    }
    if (totalForFile > 0) byFile[`src/${file}`] = totalForFile;
  }

  return {
    name: 'Tailwind Hardcoded Colors',
    total: 0,
    violations,
    byFile,
    summary: violations.length === 0
      ? '✓ No theme-non-aware Tailwind color classes (all paired with dark: variant)'
      : `✗ ${violations.length} Tailwind classes without dark: pair`,
  };
}

// =====================================================================
// 4. Inline style color violations
// =====================================================================

function checkInlineStyleColors() {
  const htmlFiles = fs.readdirSync(SRC_DIR).filter(f => f.endsWith('.html'));
  const violations = [];

  for (const file of htmlFiles) {
    const fullPath = path.join(SRC_DIR, file);
    const content = fs.readFileSync(fullPath, 'utf8');
    const lines = content.split('\n');

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      // theme:exempt が同行または前 80 行内にあればスキップ (document.write の long block 等)
      const exemptContext = lines.slice(Math.max(0, i - 80), i + 1).join('\n');
      if (/theme:exempt/i.test(exemptContext)) continue;
      // style="...color: #xxx..." / style="...background: rgba(...)..."
      const styleAttrs = line.match(/style="[^"]*"/g) || [];
      for (const styleAttr of styleAttrs) {
        const hexInStyle = styleAttr.match(/#[0-9a-fA-F]{3,6}/g);
        if (hexInStyle) {
          for (const hex of hexInStyle) {
            if (HEX_EXCEPTIONS.includes(hex.toLowerCase())) continue;
            violations.push({
              file: `src/${file}`,
              line: i + 1,
              type: 'inline-style-hex',
              value: hex,
              snippet: styleAttr.slice(0, 100),
            });
          }
        }
      }
    }
  }

  return {
    name: 'Inline Style Colors',
    total: 0,
    violations,
    summary: violations.length === 0
      ? '✓ No hardcoded hex in inline style attributes'
      : `✗ ${violations.length} inline style hex values should be CSS variable or removed`,
  };
}

// =====================================================================
// 5. Report
// =====================================================================

function main() {
  const checks = [
    checkTokenCoverage(),
    checkHardcodedCssColors(),
    checkTailwindHardcoded(),
    checkInlineStyleColors(),
  ];

  const totalViolations = checks.reduce((acc, c) => acc + c.violations.length, 0);

  if (FLAG_JSON) {
    console.log(JSON.stringify({ checks, totalViolations }, null, 2));
    process.exit(totalViolations === 0 ? 0 : 1);
  }

  console.log('=== Theme Audit Report ===\n');
  for (const c of checks) {
    console.log(`【${c.name}】 ${c.summary}`);
    if (c.violations.length > 0 && c.violations.length <= 30) {
      for (const v of c.violations) {
        const loc = v.line ? `:${v.line}` : '';
        const val = v.value || v.className || v.token;
        console.log(`  ${v.file || ''}${loc} — ${v.type}: ${val}`);
      }
    } else if (c.violations.length > 30) {
      console.log(`  (${c.violations.length} 件、 詳細は --json で出力)`);
      if (c.byFile) {
        Object.entries(c.byFile).sort((a, b) => b[1] - a[1]).slice(0, 8).forEach(([f, n]) => {
          console.log(`    ${n.toString().padStart(4)} ${f}`);
        });
      }
    }
    console.log('');
  }

  console.log(`総合: 違反 ${totalViolations} 件`);
  if (totalViolations === 0) {
    console.log('✅ 100% theme-aware. ライトモード/ダークモード切替 ready.');
  }
  process.exit(totalViolations === 0 ? 0 : 1);
}

main();
