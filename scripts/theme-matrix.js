#!/usr/bin/env node
/**
 * theme-matrix.js — テーマ全設定の網羅マトリクス & 視認性評価。
 *
 * 出力:
 *   1. 全 theme token (dark / light 両値) の一覧
 *   2. 主要な fg×bg 組合せの WCAG コントラスト評価
 *   3. CSS ファイル全体の hardcoded color を分類:
 *        - exempt (brand identity / status / 既知の theme:exempt コメント)
 *        - covered (body.theme-light 上書きが既に存在)
 *        - **uncovered** (どちらかのテーマで視認性が怪しいが override が無い)
 *   4. 改善推奨リスト (修正パッチの候補)
 *
 * 使い方:
 *   node scripts/theme-matrix.js              # 人間向け表示
 *   node scripts/theme-matrix.js --json       # CI 連携用 JSON
 *   node scripts/theme-matrix.js --uncovered  # 未対応のみ表示 (修正用 worklist)
 */

'use strict';

const fs = require('node:fs');
const path = require('node:path');

const SRC_DIR = path.resolve(__dirname, '../src');
const ARGS = process.argv.slice(2);
const FLAG_JSON = ARGS.includes('--json');
const FLAG_UNCOVERED = ARGS.includes('--uncovered');

// ─────────────────────────────────────────────────────────────
//  カラー処理 (parse / blend / contrast — theme-contrast.js と同じ)
// ─────────────────────────────────────────────────────────────

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

function relLum({ r, g, b }) {
  const lin = (c) => {
    const v = c / 255;
    return v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4);
  };
  return 0.2126 * lin(r) + 0.7152 * lin(g) + 0.0722 * lin(b);
}

function contrast(fgVal, bgVal, defaultBg = null) {
  let fg = parseColor(fgVal);
  let bg = parseColor(bgVal);
  if (!fg || !bg) return null;
  if (bg.a < 1 && defaultBg) bg = blendOnBg(bg, parseColor(defaultBg));
  if (bg.a < 1) bg = blendOnBg(bg, { r: 255, g: 255, b: 255, a: 1 });
  if (fg.a < 1) fg = blendOnBg(fg, bg);
  const l1 = relLum(fg);
  const l2 = relLum(bg);
  return Math.round(((Math.max(l1, l2) + 0.05) / (Math.min(l1, l2) + 0.05)) * 100) / 100;
}

// ─────────────────────────────────────────────────────────────
//  Token 抽出
// ─────────────────────────────────────────────────────────────

function extractBlock(css, startRegex) {
  const m = css.match(startRegex);
  if (!m) return '';
  const startIdx = m.index + m[0].lastIndexOf('{');
  let depth = 1, i = startIdx + 1;
  while (i < css.length && depth > 0) {
    if (css[i] === '{') depth++;
    else if (css[i] === '}') depth--;
    i++;
  }
  return css.slice(startIdx + 1, i - 1);
}

function extractTokens(body) {
  const tokens = {};
  const re = /--([a-z0-9-]+)\s*:\s*([^;]+);/gi;
  let m;
  while ((m = re.exec(body)) !== null) {
    tokens[`--${m[1]}`] = m[2].trim();
  }
  return tokens;
}

function resolveVar(tokens, name, depth = 0) {
  if (depth > 8) return tokens[name] || null;
  const val = tokens[name];
  if (!val) return null;
  const m = val.match(/^var\(--([a-z0-9-]+)(?:,\s*([^)]+))?\)$/);
  if (!m) return val;
  const ref = `--${m[1]}`;
  if (tokens[ref]) return resolveVar(tokens, ref, depth + 1);
  if (m[2]) return m[2].trim();
  return null;
}

// ─────────────────────────────────────────────────────────────
//  Hardcoded color スキャン
// ─────────────────────────────────────────────────────────────

const CSS_FILES = [
  'UnifiedStyles.css.html',
  'page.css.html',
  'page.viz.css.html',
  'SetupPage.css.html',
];

const HEX_RE = /#[0-9a-fA-F]{3,8}\b/g;
const RGBA_RE = /rgba?\([^)]+\)/g;

// brand identity / 確定 exempt な値 (theme 非依存で OK)
const EXEMPT_COLORS = new Set([
  // brand reaction colors (両モード共通でアイデンティティ)
  '#e11d48', '#fbbf24', '#059669', '#9333ea', '#c084fc',
  '#fde047', '#4ade80', '#fbbf24', '#fdba74', '#86efac',
  '#3b82f6', '#1d4ed8', '#8be9fd',
  // status (両モード共通)
  '#10b981', '#ef4444', '#f59e0b',
  // google brand
  '#4285f4', '#34a853', '#fbbc05', '#ea4335',
  // 純白/純黒 alpha (ベタ重ねの透明オーバーレイ; 殆ど両モード OK)
  '#fff', '#ffffff', '#fff0', '#000', '#000000',
]);

function scanHardcoded(filePath) {
  const text = fs.readFileSync(filePath, 'utf8');
  const lines = text.split('\n');
  const findings = [];
  lines.forEach((line, idx) => {
    // theme:exempt コメント付き行はスキップ
    if (/theme:exempt/i.test(line)) return;
    // CSS 変数定義行 (`--xxx: value;`) はスキップ — token 自身は分析対象
    if (/^\s*--[a-z0-9-]+\s*:/.test(line)) return;
    // comment line skip
    const stripped = line.replace(/\/\*.*?\*\//g, '');
    const hexMatches = stripped.match(HEX_RE) || [];
    const rgbaMatches = stripped.match(RGBA_RE) || [];
    [...hexMatches, ...rgbaMatches].forEach(color => {
      const norm = color.toLowerCase();
      if (EXEMPT_COLORS.has(norm)) return;
      findings.push({ file: path.basename(filePath), line: idx + 1, color, context: stripped.trim().slice(0, 100) });
    });
  });
  return findings;
}

// ─────────────────────────────────────────────────────────────
//  Main
// ─────────────────────────────────────────────────────────────

function main() {
  const unifiedPath = path.join(SRC_DIR, 'UnifiedStyles.css.html');
  const unifiedCss = fs.readFileSync(unifiedPath, 'utf8');

  const rootBody = extractBlock(unifiedCss, /:root\s*\{/);
  const lightBody = extractBlock(unifiedCss, /body\.theme-light[^{]*\{/);
  const darkTokens = extractTokens(rootBody);
  const lightOverrides = extractTokens(lightBody);
  const lightTokens = { ...darkTokens, ...lightOverrides };

  // ─── 1. Token 一覧 ───
  const themeTokenNames = Object.keys(darkTokens).filter(n => n.startsWith('--theme-') || n.startsWith('--brand-') || n.startsWith('--status-'));
  const tokenTable = themeTokenNames.map(name => {
    const dval = resolveVar(darkTokens, name);
    const lval = resolveVar(lightTokens, name);
    const inOverride = lightOverrides[name] !== undefined;
    return {
      token: name,
      dark: dval,
      light: lval,
      themed: inOverride,
      family: name.startsWith('--theme-') ? 'theme' : name.startsWith('--brand-') ? 'brand' : 'status',
    };
  });

  // ─── 2. Contrast マトリクス ───
  // 主要な fg×bg 組合せ — テキスト系 vs 背景系
  const TEXT_TOKENS = ['--theme-text-primary', '--theme-text-secondary', '--theme-text-muted', '--theme-accent-cyan', '--theme-spinner', '--brand-like', '--brand-understand', '--brand-curious', '--brand-highlight', '--status-success', '--status-error', '--status-warning'];
  const BG_TOKENS = ['--theme-bg-base', '--theme-bg-surface', '--theme-bg-elevated'];

  const matrix = { dark: [], light: [] };
  for (const mode of ['dark', 'light']) {
    const tokens = mode === 'dark' ? darkTokens : lightTokens;
    const baseBg = resolveVar(tokens, '--theme-bg-base');
    for (const fg of TEXT_TOKENS) {
      for (const bg of BG_TOKENS) {
        const fgVal = resolveVar(tokens, fg);
        const bgVal = resolveVar(tokens, bg);
        if (!fgVal || !bgVal) continue;
        const ratio = contrast(fgVal, bgVal, baseBg);
        if (ratio == null) continue;
        // 評価基準: text-* は 4.5, accent/status は 3.0
        const isStrictText = fg.includes('text-');
        const target = isStrictText ? 4.5 : 3.0;
        const isLargeText = !isStrictText;  // accent/status は icon/large-text 想定
        matrix[mode].push({
          fg, bg, fgVal, bgVal, ratio, target,
          status: ratio >= target ? 'pass' : ratio >= 3.0 && isLargeText ? 'pass-large' : 'fail',
          isLargeText,
        });
      }
    }
  }

  // ─── 3. Hardcoded color スキャン ───
  const allFindings = [];
  for (const file of CSS_FILES) {
    const fp = path.join(SRC_DIR, file);
    if (!fs.existsSync(fp)) continue;
    allFindings.push(...scanHardcoded(fp));
  }
  // body.theme-light 上書きが「同じ class」に存在する場合は covered とみなす
  const lightOverrideClasses = new Set();
  const lightBlockRe = /body\.theme-light[^{]*\.(\w[\w-]*)/g;
  for (const file of CSS_FILES) {
    const fp = path.join(SRC_DIR, file);
    if (!fs.existsSync(fp)) continue;
    const text = fs.readFileSync(fp, 'utf8');
    let m;
    while ((m = lightBlockRe.exec(text)) !== null) lightOverrideClasses.add(m[1]);
  }

  // findings に「同じ class が override されているか」フラグを付ける
  const findingsByFile = {};
  for (const f of allFindings) {
    const filePath = path.join(SRC_DIR, f.file);
    const text = fs.readFileSync(filePath, 'utf8');
    // 行から逆走してこの行を含むセレクタを探す
    const upTo = text.split('\n').slice(0, f.line).join('\n');
    const lastBrace = upTo.lastIndexOf('{');
    const selectorStart = upTo.lastIndexOf('}', lastBrace) + 1;
    const selector = upTo.slice(selectorStart, lastBrace).trim();
    const classMatch = selector.match(/\.([\w-]+)/);
    const cls = classMatch ? classMatch[1] : null;
    const covered = cls && lightOverrideClasses.has(cls);
    const enriched = { ...f, selector: selector.slice(0, 80), cls, covered };
    if (!findingsByFile[f.file]) findingsByFile[f.file] = [];
    findingsByFile[f.file].push(enriched);
  }

  if (FLAG_JSON) {
    console.log(JSON.stringify({ tokens: tokenTable, matrix, findings: findingsByFile }, null, 2));
    return;
  }

  if (FLAG_UNCOVERED) {
    console.log('\n=== Uncovered hardcoded colors (light theme で override が無い) ===\n');
    let total = 0;
    for (const [file, list] of Object.entries(findingsByFile)) {
      const un = list.filter(x => !x.covered);
      if (un.length === 0) continue;
      console.log(`◆ ${file}  (${un.length} 件)`);
      un.slice(0, 30).forEach(x => {
        console.log(`  L${String(x.line).padStart(4)}  ${x.color.padEnd(28)} .${x.cls || '?'}`);
      });
      if (un.length > 30) console.log(`  ... and ${un.length - 30} more`);
      total += un.length;
    }
    console.log(`\n合計 uncovered: ${total} 件`);
    return;
  }

  // ─── Human-readable report ───
  console.log('\n══════════════════════════════════════════════════════════════════');
  console.log('  Theme Matrix — 全 token / contrast / hardcoded color の網羅レポート');
  console.log('══════════════════════════════════════════════════════════════════\n');

  console.log('【1. Theme Token 一覧 (dark / light 値)】\n');
  const families = { theme: [], brand: [], status: [] };
  tokenTable.forEach(t => families[t.family].push(t));
  for (const fam of ['theme', 'brand', 'status']) {
    console.log(`  ── ${fam.toUpperCase()} ──`);
    families[fam].forEach(t => {
      const dark = (t.dark || '?').padEnd(36);
      const light = (t.light || '?').padEnd(36);
      const flag = t.themed ? ' ✓themed' : (t.dark === t.light ? '' : ' (=dark)');
      console.log(`    ${t.token.padEnd(36)} dark=${dark} light=${light}${flag}`);
    });
    console.log('');
  }

  console.log('【2. WCAG Contrast マトリクス (主要 fg × bg)】\n');
  for (const mode of ['dark', 'light']) {
    console.log(`  ── ${mode.toUpperCase()} mode ──`);
    matrix[mode].forEach(r => {
      const mark = r.status === 'pass' ? '✓' : r.status === 'pass-large' ? '◆' : '✗';
      const ratio = String(r.ratio).padStart(6);
      console.log(`    ${mark} ${ratio} (>${r.target})  ${r.fg.padEnd(28)} on  ${r.bg}`);
    });
    console.log('');
  }

  console.log('【3. Hardcoded color スキャン (CSS ファイル全体)】\n');
  for (const [file, list] of Object.entries(findingsByFile)) {
    const covered = list.filter(x => x.covered).length;
    const uncovered = list.filter(x => !x.covered).length;
    console.log(`  ◆ ${file}: 計 ${list.length} 件  (theme-light covered=${covered}, uncovered=${uncovered})`);
  }
  const totalUncovered = Object.values(findingsByFile).reduce((s, l) => s + l.filter(x => !x.covered).length, 0);
  console.log(`\n  ▶ 合計 uncovered: ${totalUncovered} 件  ※詳細は --uncovered フラグで`);
  console.log('');

  // 評価サマリ
  let fails = 0;
  for (const mode of ['dark', 'light']) {
    matrix[mode].forEach(r => { if (r.status === 'fail') fails++; });
  }
  console.log('【4. 評価サマリ】\n');
  console.log(`  Contrast pass: ${matrix.dark.length + matrix.light.length - fails}/${matrix.dark.length + matrix.light.length}`);
  console.log(`  Contrast fail: ${fails}`);
  console.log(`  Hardcoded uncovered: ${totalUncovered}`);
  console.log('');
}

main();
