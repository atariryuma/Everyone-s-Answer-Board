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
const { parseColor, blendOnBg, relativeLuminance, extractBlockBody, extractTokensFromBody } = require('./lib/theme-core');

const SRC_DIR = path.resolve(__dirname, '../src');
const ARGS = process.argv.slice(2);
const FLAG_JSON = ARGS.includes('--json');
const FLAG_UNCOVERED = ARGS.includes('--uncovered');

// ─────────────────────────────────────────────────────────────
//  カラー処理 (parseColor / blendOnBg / relativeLuminance は
//   scripts/lib/theme-core.js に共通化。contrast は matrix 固有の
//   丸め込み + defaultBg blend なのでここに残す)
// ─────────────────────────────────────────────────────────────

function contrast(fgVal, bgVal, defaultBg = null) {
  let fg = parseColor(fgVal);
  let bg = parseColor(bgVal);
  if (!fg || !bg) return null;
  if (bg.a < 1 && defaultBg) bg = blendOnBg(bg, parseColor(defaultBg));
  if (bg.a < 1) bg = blendOnBg(bg, { r: 255, g: 255, b: 255, a: 1 });
  if (fg.a < 1) fg = blendOnBg(fg, bg);
  const l1 = relativeLuminance(fg);
  const l2 = relativeLuminance(bg);
  return Math.round(((Math.max(l1, l2) + 0.05) / (Math.min(l1, l2) + 0.05)) * 100) / 100;
}

// ─────────────────────────────────────────────────────────────
//  Token 抽出 (extractBlockBody / extractTokensFromBody は
//   scripts/lib/theme-core.js に共通化。resolveVar は matrix 固有の
//   完全一致 regex + depth ガードなのでここに残す)
// ─────────────────────────────────────────────────────────────

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

// 全ての <style> ブロック を持つ HTML を対象に — 漏れ無くスキャンする
const SCAN_FILES = [
  'UnifiedStyles.css.html',
  'page.css.html',
  'page.viz.css.html',
  'AdminPanel.html',
  'AdminPanel.js.html',
  'AccessRestricted.html',
  'LoginPage.html',
  'login.js.html',
  'TeacherManual.html',
  'Unpublished.html',
  'AppSetupPage.html',
  'ErrorBoundary.html',
  'LessonWorkspace.html',
  'SetupPage.html',
  'SharedErrorHandling.html',
  'SharedUtilities.html',
];
const CSS_FILES = SCAN_FILES;  // 互換用エイリアス

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

// 色を意味カテゴリに分類:
//   - shadow:        rgba(0,0,0,X) / rgba(255,255,255,X) — 両モード安全な影/光
//   - brand:         brand identity の RGB を alpha < 1 で被せた accent overlay (両モード OK)
//   - theme-bleed:   slate/gray の中間 RGB 単色 (light mode で broken の可能性高い)
//   - other:         未分類 (要確認)
const BRAND_RGB = [
  [56, 189, 248],   // sky-400 / accent cyan
  [34, 211, 238],   // cyan-400
  [139, 233, 253],  // accent-light cyan
  [125, 211, 252],  // sky-300
  [103, 232, 249],  // cyan-300
  [251, 191, 36],   // amber-400 / understand
  [245, 158, 11],   // amber-500
  [253, 224, 71],   // yellow-300
  [34, 197, 94],    // green-500 / curious
  [16, 185, 129],   // emerald-500 / success
  [248, 113, 113],  // red-400
  [239, 68, 68],    // red-500 / error
  [220, 38, 38],    // red-600
  [225, 29, 72],    // rose-600 / like
  [168, 85, 247],   // purple-500 / highlight
  [147, 51, 234],   // purple-600
  [192, 132, 252],  // purple-300
  [59, 130, 246],   // blue-500 / primary
  [96, 165, 250],   // blue-400
  [139, 92, 246],   // violet-500
  [6, 182, 212],    // cyan-600
  [202, 138, 4],    // yellow-700
  [22, 163, 74],    // green-600
  [14, 116, 144],   // cyan-700
  [55, 48, 163],    // indigo-800
  [199, 210, 254],  // indigo-200
  [234, 88, 12],    // orange-600
];
// slate/gray のニュートラル RGB — alpha 持ち solid bg は theme-bleed の疑い高
const NEUTRAL_RGB = [
  [148, 163, 184],  // slate-400
  [100, 116, 139],  // slate-500
  [71, 85, 105],    // slate-600
  [51, 65, 85],     // slate-700
  [30, 41, 59],     // slate-800
  [15, 23, 42],     // slate-900
  [26, 27, 38],     // tokyo-night base
  [107, 114, 128],  // gray-500
  [156, 163, 175],  // gray-400
  [209, 213, 219],  // gray-300
  [55, 65, 81],     // gray-700
  [75, 85, 99],     // gray-600
  [17, 24, 39],     // gray-900
  [241, 245, 249],  // slate-100
  [226, 232, 240],  // slate-200
];

function classifyColor(colorStr) {
  const m = colorStr.match(/(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/);
  if (!m) {
    // hex
    const hexMatch = colorStr.match(/#([0-9a-fA-F]{3,8})/);
    if (!hexMatch) return 'other';
    let h = hexMatch[1];
    if (h.length === 3) h = h.split('').map(c => c + c).join('');
    if (h.length !== 6) return 'other';
    const r = parseInt(h.slice(0, 2), 16);
    const g = parseInt(h.slice(2, 4), 16);
    const b = parseInt(h.slice(4, 6), 16);
    return classifyRGB(r, g, b);
  }
  return classifyRGB(+m[1], +m[2], +m[3]);
}
function rgbDistance(a, b) {
  return Math.sqrt((a[0]-b[0])**2 + (a[1]-b[1])**2 + (a[2]-b[2])**2);
}
function classifyRGB(r, g, b) {
  if ((r < 20 && g < 20 && b < 20) || (r > 240 && g > 240 && b > 240)) return 'shadow';
  for (const c of BRAND_RGB) if (rgbDistance([r,g,b], c) < 12) return 'brand';
  for (const c of NEUTRAL_RGB) if (rgbDistance([r,g,b], c) < 12) return 'theme-bleed';
  return 'other';
}

function scanHardcoded(filePath) {
  const text = fs.readFileSync(filePath, 'utf8');
  const lines = text.split('\n');
  const findings = [];
  // theme:exempt-block-start .. -block-end の範囲を計算
  const exemptRanges = [];
  let blockStart = -1;
  lines.forEach((line, idx) => {
    if (/theme:exempt-block-start|document\.write\(/i.test(line)) {
      if (blockStart === -1) blockStart = idx;
    }
    if (blockStart !== -1 && /theme:exempt-block-end|`\);/.test(line)) {
      exemptRanges.push([blockStart, idx]);
      blockStart = -1;
    }
  });
  // :root { ... } / body.theme-light { ... } ブロック内も exempt (token override の中身は値そのもの)
  const SELECTOR_PATTERNS = [/:root\s*\{/g, /body\.theme-light[^{]*\{/g, /:root\[data-theme="light"\]\s*\{/g];
  for (const re of SELECTOR_PATTERNS) {
    re.lastIndex = 0;
    let m;
    while ((m = re.exec(text)) !== null) {
      const start = m.index;
      let depth = 0, i = start;
      while (i < text.length) {
        if (text[i] === '{') depth++;
        else if (text[i] === '}') {
          depth--;
          if (depth === 0) {
            const startLine = text.slice(0, start).split('\n').length - 1;
            const endLine = text.slice(0, i + 1).split('\n').length - 1;
            exemptRanges.push([startLine, endLine]);
            break;
          }
        }
        i++;
      }
    }
  }
  const inExemptBlock = (idx) => exemptRanges.some(([s, e]) => idx >= s && idx <= e);

  lines.forEach((line, idx) => {
    if (inExemptBlock(idx)) return;
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
      const category = classifyColor(color);
      findings.push({ file: path.basename(filePath), line: idx + 1, color, category, context: stripped.trim().slice(0, 100) });
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

  const rootBody = extractBlockBody(unifiedCss, /:root\s*\{/);
  const lightBody = extractBlockBody(unifiedCss, /body\.theme-light[^{]*\{/);
  const darkTokens = extractTokensFromBody(rootBody);
  const lightOverrides = extractTokensFromBody(lightBody);
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
    let total = 0, bleed = 0;
    for (const [file, list] of Object.entries(findingsByFile)) {
      const un = list.filter(x => !x.covered);
      if (un.length === 0) continue;
      const byCat = { 'theme-bleed': [], 'brand': [], 'shadow': [], 'other': [] };
      un.forEach(x => byCat[x.category || 'other'].push(x));
      console.log(`◆ ${file}  (${un.length} 件: bleed=${byCat['theme-bleed'].length}, brand=${byCat['brand'].length}, shadow=${byCat['shadow'].length}, other=${byCat['other'].length})`);
      // bleed (actionable) を全件表示、 brand/shadow は件数のみ
      byCat['theme-bleed'].forEach(x => {
        console.log(`  ⚠ L${String(x.line).padStart(4)} BLEED ${x.color.padEnd(28)} .${x.cls || '?'}`);
      });
      byCat['other'].slice(0, 10).forEach(x => {
        console.log(`  ? L${String(x.line).padStart(4)} OTHER ${x.color.padEnd(28)} .${x.cls || '?'}`);
      });
      total += un.length;
      bleed += byCat['theme-bleed'].length;
    }
    console.log(`\n合計 uncovered: ${total} 件 (うち actionable theme-bleed: ${bleed} 件)`);
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
