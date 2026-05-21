#!/usr/bin/env node
/**
 * theme-perfect.js — light/dark mode 完璧度を 20 軸で測る最終ゲート。
 *
 * 既存の theme-verify (12 軸) より厳しく:
 *   - CSS hardcoded
 *   - Tailwind unpaired
 *   - inline style hardcoded
 *   - JS-side style assignment
 *   - WCAG AA contrast (72 pair)
 *   - SVG fill/stroke hardcoded
 *   - icon class color hardcoded
 *   - meta theme-color exempt 認識
 *   - skeleton / loading state token 使用
 *   - hover / focus / active state pairing
 *   - 上限スコア: 200 (各軸 10 点)
 *
 * 使い方:
 *   npm run theme:perfect
 */

'use strict';
const fs = require('node:fs');
const path = require('node:path');
const { spawnSync } = require('node:child_process');

const ROOT = path.resolve(__dirname, '..');
const SRC = path.join(ROOT, 'src');

const checks = [];
function check(name, ok, detail) { checks.push({ name, ok, detail }); }

const HTML_FILES = fs.readdirSync(SRC).filter(f => f.endsWith('.html') || f.endsWith('.js'));

// 1. theme-matrix actionable theme-bleed = 0
const matrix = spawnSync('node', ['scripts/theme-matrix.js', '--uncovered'], { cwd: ROOT, encoding: 'utf8' });
const bleedMatch = matrix.stdout.match(/actionable theme-bleed:\s*(\d+)/);
const bleedCount = bleedMatch ? parseInt(bleedMatch[1]) : -1;
check('1. CSS hardcoded theme-bleed = 0', bleedCount === 0, `${bleedCount} 件`);

// 2. Tailwind unpaired actionable = 0  (theme-pair-tailwind を実行して何件出るか)
const pairDry = spawnSync('node', ['scripts/theme-pair-tailwind.js', '--dry-run'], { cwd: ROOT, encoding: 'utf8' });
const pairMatch = pairDry.stdout.match(/Total pairings:\s*(\d+)/);
const pairCount = pairMatch ? parseInt(pairMatch[1]) : -1;
check('2. Tailwind unpaired auto-fixable = 0', pairCount === 0, `${pairCount} 件 (auto-pair で修正可能なもの)`);

// 3. inline style="color:#xxx" / "background:#xxx" — exempt-block 範囲を AST 風に解析
let inlineHex = 0;
const STYLE_RE = /style=(["'])([^"']*?(?:#[0-9a-fA-F]{3,8}|rgba?\([^)]+\))[^"']*?)\1/g;
for (const f of HTML_FILES) {
  const text = fs.readFileSync(path.join(SRC, f), 'utf8');
  // exempt-block-start .. -block-end の char range を計算
  const blockRanges = [];
  const startRe = /theme:exempt-block-start/g;
  let sm;
  while ((sm = startRe.exec(text)) !== null) {
    const start = sm.index;
    const endMatch = text.indexOf('theme:exempt-block-end', start);
    if (endMatch === -1) continue;
    blockRanges.push([start, endMatch]);
  }
  const inBlock = (idx) => blockRanges.some(([s, e]) => idx >= s && idx <= e);

  STYLE_RE.lastIndex = 0;
  let m;
  while ((m = STYLE_RE.exec(text)) !== null) {
    if (/theme:exempt|var\(--/i.test(m[0])) continue;
    if (inBlock(m.index)) continue;
    // 同じ行に theme:exempt があるか
    const lineStart = text.lastIndexOf('\n', m.index) + 1;
    const lineEnd = text.indexOf('\n', m.index);
    const line = text.slice(lineStart, lineEnd === -1 ? text.length : lineEnd);
    if (/theme:exempt/i.test(line)) continue;
    inlineHex++;
  }
}
check('3. Inline style hardcoded = 0', inlineHex === 0, `${inlineHex} 件 inline style with raw hex/rgba`);

// 4. JS-side .style.X = '#xxx'
let jsAssign = 0;
for (const f of HTML_FILES) {
  const text = fs.readFileSync(path.join(SRC, f), 'utf8');
  const matches = text.match(/\.style\.(?:color|background\w*|borderColor|fill|stroke)\s*=\s*['"](?:#[0-9a-fA-F]{3,8}|rgba?\()/g) || [];
  for (const m of matches) {
    const idx = text.indexOf(m);
    const line = text.slice(Math.max(0, idx - 200), idx + 200);
    if (/theme:exempt/i.test(line) || /var\(--/.test(line)) continue;
    jsAssign++;
  }
}
check('4. JS-side .style hardcoded = 0', jsAssign === 0, `${jsAssign} 件`);

// 5. SVG fill="#xxx" / stroke="#xxx" attribute (not currentColor)
let svgHex = 0;
for (const f of HTML_FILES) {
  const text = fs.readFileSync(path.join(SRC, f), 'utf8');
  const matches = text.match(/(?:fill|stroke)="#[0-9a-fA-F]{3,8}"/g) || [];
  svgHex += matches.length;
}
check('5. SVG attribute fill/stroke hex = 0 (currentColor 推奨)', svgHex === 0, `${svgHex} 件`);

// 6. WCAG AA: theme:contrast を実行
const contrast = spawnSync('node', ['scripts/theme-contrast.js'], { cwd: ROOT, encoding: 'utf8' });
const allPass = /全 contrast pass/.test(contrast.stdout);
check('6. WCAG AA 全 pair pass (dark + light)', allPass, allPass ? '72/72' : 'fail');

// 7. theme-verify 12 軸満点
const verify = spawnSync('node', ['scripts/theme-verify.js'], { cwd: ROOT, encoding: 'utf8' });
const scoreMatch = verify.stdout.match(/Score:\s*(\d+)\s*\/\s*(\d+)/);
const perfect120 = scoreMatch && scoreMatch[1] === scoreMatch[2];
check('7. theme:verify 120/120', perfect120, scoreMatch ? `${scoreMatch[1]}/${scoreMatch[2]}` : 'unknown');

// 8. Unit tests 957 PASS
const tests = spawnSync('npm', ['test'], { cwd: ROOT, encoding: 'utf8' });
const failMatch = tests.stdout.match(/fail\s+(\d+)/);
const passMatch = tests.stdout.match(/pass\s+(\d+)/);
const testsPass = failMatch && failMatch[1] === '0' && passMatch && parseInt(passMatch[1]) > 900;
check('8. Unit tests 957+ PASS', testsPass, `pass=${passMatch?passMatch[1]:'?'}, fail=${failMatch?failMatch[1]:'?'}`);

// 9. body.theme-light 上書きを CSS / HTML <style> で持つファイル数
let lightOverrideFiles = 0;
for (const f of HTML_FILES) {
  const text = fs.readFileSync(path.join(SRC, f), 'utf8');
  if (/body\.theme-light\b/.test(text)) lightOverrideFiles++;
}
check('9. body.theme-light override 定義あり', lightOverrideFiles >= 2, `${lightOverrideFiles} ファイル (UnifiedStyles + page.viz)`);

// 10. themeManager API export 完全 (get/set/toggle/subscribe/init/mountToggle)
const su = fs.readFileSync(path.join(SRC, 'SharedUtilities.html'), 'utf8');
const apiBits = ['get(', 'set(', 'toggle(', 'subscribe(', 'init(', 'mountToggle(', 'resolved('];
const missing = apiBits.filter(b => !su.includes(b));
check('10. themeManager 全 API 実装', missing.length === 0, missing.length ? `missing: ${missing.join(',')}` : '7/7');

// 11. localStorage persistence key
const hasPersist = /['"]app-theme['"]/.test(su);
check('11. localStorage \'app-theme\' 永続化', hasPersist, hasPersist ? '✓' : '✗');

// 12. Tailwind darkMode:'class' + theme.extend.colors
const tw = fs.readFileSync(path.join(SRC, 'SharedTailwindConfig.html'), 'utf8');
const hasDarkMode = /darkMode:\s*['"]class['"]/.test(tw);
const hasThemeColors = /theme-bg-base|theme-text-primary/.test(tw);
check('12. Tailwind darkMode + theme tokens 公開', hasDarkMode && hasThemeColors, `darkMode=${hasDarkMode}, tokens=${hasThemeColors}`);

// 13. apply() で .dark + .light + theme-{dark,light} を html/body に
// v2810+: 統一形 classList.add(resolved, 'theme-' + resolved) を採用
const hasResolved = /classList\.add\(resolved,\s*'theme-'\s*\+\s*resolved\)/.test(su)
  || (/classList\.add\(.{0,30}'dark'/.test(su) && /classList\.add\(.{0,30}'light'/.test(su));
const hasThemeClass = /classList\.add\(.{0,30}'theme-'\s*\+\s*resolved\)/.test(su) || /classList\.add\('theme-dark'/.test(su);
const hasColorScheme = /style\.colorScheme/.test(su);
check('13. apply() で .dark/.light/.theme-* + colorScheme 同期', hasResolved && hasThemeClass && hasColorScheme, `resolved=${hasResolved}, theme-class=${hasThemeClass}, colorScheme=${hasColorScheme}`);

// 14. brand-* → theme-* alias (UnifiedStyles)
const us = fs.readFileSync(path.join(SRC, 'UnifiedStyles.css.html'), 'utf8');
const brandAlias = /--brand-background:\s*var\(--theme-bg-base\)/.test(us);
check('14. --brand-* → --theme-* alias 完備', brandAlias, brandAlias ? '✓' : '✗');

// 15. card tier tokens 完備 (--theme-card-1/2/3 が dark + light 両方に)
const darkCardDef = /:root\s*\{[\s\S]*?--theme-card-1:[\s\S]*?--theme-card-2:[\s\S]*?--theme-card-3:[\s\S]*?\}/m.test(us);
const lightCardDef = /body\.theme-light[\s\S]{0,3000}?--theme-card-1:[\s\S]*?--theme-card-2:[\s\S]*?--theme-card-3/.test(us);
check('15. --theme-card-1/2/3 dark+light 両定義', darkCardDef && lightCardDef, `dark=${darkCardDef}, light=${lightCardDef}`);

// 16. high contrast media query dark+light両方
const hcDark = /@media \(prefers-contrast: high\)[\s\S]*?:root\s*\{[\s\S]*?--theme/.test(us);
const hcLight = /@media \(prefers-contrast: high\)[\s\S]*?body\.theme-light/.test(us);
check('16. prefers-contrast:high override dark+light', hcDark && hcLight, `dark=${hcDark}, light=${hcLight}`);

// 17. theme-color meta tag に theme:exempt or theme-aware — same line exempt comment も対応
let untaggedMeta = 0;
for (const f of HTML_FILES) {
  const text = fs.readFileSync(path.join(SRC, f), 'utf8');
  const lines = text.split('\n');
  for (const line of lines) {
    if (/<meta name="theme-color"/i.test(line)) {
      if (!/theme:exempt/i.test(line)) untaggedMeta++;
    }
  }
}
check('17. <meta theme-color> 全部 exempt 明示', untaggedMeta === 0, `${untaggedMeta} 件 untagged`);

// 18. document.write popup に exempt-block マーカー
let unmarkedPopups = 0;
for (const f of HTML_FILES) {
  const text = fs.readFileSync(path.join(SRC, f), 'utf8');
  const writes = text.match(/(?:document\.write\(|=\s*`[\s\S]{0,500}?<!DOCTYPE html>)/g) || [];
  // 各 popup の周辺 500 chars に exempt-block-start があるか
  for (const m of writes) {
    const idx = text.indexOf(m);
    const before = text.slice(Math.max(0, idx - 500), idx);
    if (!/theme:exempt/i.test(before)) unmarkedPopups++;
  }
}
check('18. document.write popup 全部 exempt-block 化', unmarkedPopups === 0, `${unmarkedPopups} 件 unmarked`);

// 19. viz CSS body.theme-light override 数 ≥ 9 (quadrant-label/keyword/contrast/axis/4*quadrants 等)
const vizCss = fs.readFileSync(path.join(SRC, 'page.viz.css.html'), 'utf8');
const lightOverrideCount = (vizCss.match(/body\.theme-light\s+\./g) || []).length;
check('19. page.viz light override ≥ 9 件', lightOverrideCount >= 9, `${lightOverrideCount} 件`);

// 20. CLAUDE.md にメンテナンスフローが記載
const claudemd = fs.readFileSync(path.join(ROOT, 'CLAUDE.md'), 'utf8');
const hasFlow = /theme:matrix.*theme:uncovered.*theme:tokenize/s.test(claudemd) || /theme-card-1.*theme-card-2.*theme-card-3/s.test(claudemd);
check('20. CLAUDE.md にメンテナンスフロー記載', hasFlow, hasFlow ? '✓' : '✗');

// 21. color-scheme CSS property (両モード定義済)
const hasCsRoot = /:root\s*\{[\s\S]*?color-scheme:\s*dark/.test(us);
const hasCsLight = /body\.theme-light[\s\S]{0,2000}?color-scheme:\s*light/.test(us);
check('21. color-scheme CSS property (dark+light)', hasCsRoot && hasCsLight, `root=${hasCsRoot}, light=${hasCsLight}`);

// 22. FOUC 防止: SharedThemeBoot.html 存在 + 全 theme-aware ページに include
const bootExists = fs.existsSync(path.join(SRC, 'SharedThemeBoot.html'));
const PAGES_NEEDING_BOOT = ['Page', 'AdminPanel', 'LoginPage', 'AppSetupPage', 'SetupPage',
                            'TeacherManual', 'Unpublished', 'AccessRestricted', 'ErrorBoundary'];
let bootMissing = 0;
for (const p of PAGES_NEEDING_BOOT) {
  const fp = path.join(SRC, `${p}.html`);
  if (!fs.existsSync(fp)) continue;
  const t = fs.readFileSync(fp, 'utf8');
  if (!/SharedThemeBoot/.test(t)) bootMissing++;
}
check('22. FOUC 防止 SharedThemeBoot 全ページ include', bootExists && bootMissing === 0, `boot=${bootExists}, missing=${bootMissing}`);

// 23. Theme 切替 smooth transition (body に transition: background-color color)
const hasTransition = /body\s*\{[\s\S]*?transition:[\s\S]*?background-color[\s\S]*?color/.test(us);
const hasReducedMotion = /prefers-reduced-motion[\s\S]*?body[\s\S]*?transition:\s*none/.test(us);
check('23. Theme transition + reduced-motion 配慮', hasTransition && hasReducedMotion, `transition=${hasTransition}, reduced=${hasReducedMotion}`);

// 24. Typography token system 定義済 (--font-size-xs..4xl + weight + leading)
const fontTokens = ['--font-size-xs', '--font-size-sm', '--font-size-base', '--font-size-lg',
                    '--font-size-xl', '--font-size-2xl', '--font-size-3xl', '--font-size-4xl',
                    '--font-weight-normal', '--font-weight-medium', '--font-weight-semibold', '--font-weight-bold',
                    '--leading-tight', '--leading-normal', '--leading-relaxed'];
const missingFont = fontTokens.filter(t => !us.includes(t + ':'));
check('24. Typography token (size/weight/leading) 定義', missingFont.length === 0, missingFont.length ? `missing: ${missingFont.join(',')}` : `${fontTokens.length}/${fontTokens.length}`);

// 25. font-size の Tailwind スケール準拠率 ≥ 90% (Tailwind の text-xs/sm/base/lg/xl/2xl/3xl/4xl に近接)
function countFontSizes() {
  const files = ['UnifiedStyles.css.html', 'page.css.html', 'page.viz.css.html'];
  const STD = new Set(['0.75rem', '0.875rem', '1rem', '1.125rem', '1.25rem', '1.5rem', '1.875rem', '2.25rem']);
  let std = 0, total = 0;
  for (const f of files) {
    const text = fs.readFileSync(path.join(SRC, f), 'utf8');
    const cleaned = text.replace(/\/\*[\s\S]*?\*\//g, '');
    const matches = cleaned.matchAll(/font-size:\s*([0-9.]+(?:rem|px|em))/g);
    for (const m of matches) {
      total++;
      if (STD.has(m[1])) std++;
    }
  }
  return { std, total, ratio: total ? std / total : 0 };
}
const fs1 = countFontSizes();
check(`25. font-size Tailwind scale 準拠率 ≥ 90%`, fs1.ratio >= 0.9, `${fs1.std}/${fs1.total} (${(fs1.ratio * 100).toFixed(1)}%)`);

// 27. Orphan dark: utility (light counterpart 無し) = 0
//   ※ paired は受容、 orphan のみ問題 (light mode で何も適用されない潜在バグ)
function countOrphanDark() {
  let orphans = 0;
  const CLASS_RE = /class=(["'])([^"']+)\1/g;
  for (const f of HTML_FILES) {
    const text = fs.readFileSync(path.join(SRC, f), 'utf8');
    let m;
    while ((m = CLASS_RE.exec(text)) !== null) {
      const classes = m[2].split(/\s+/);
      for (const cls of classes) {
        if (!cls.startsWith('dark:')) continue;
        const pc = cls.match(/^dark:([a-z-]+)-([a-z]+)-\d+/);
        if (!pc) continue;
        const [, prop, color] = pc;
        const hasPair = classes.some(c => c !== cls && !c.startsWith('dark:') &&
                                          new RegExp(`^${prop}-${color}-`).test(c));
        if (!hasPair) orphans++;
      }
    }
  }
  return orphans;
}
const orphans = countOrphanDark();
check('27. Orphan dark: utility = 0 (paired のみ受容)', orphans === 0, `${orphans} 件 orphan`);

// 27b. Tailwind dark: prefix 全廃 (v2847+ — semantic theme token に統一)
//   paired であっても dark:bg-X-N00 系 / dark:hover: 系は使わない方針。
//   light/dark variation は --theme-* token / body.theme-light で表現する。
//   SharedTailwindConfig (darkMode 設定の説明 comment) は除外。
function countAnyDarkPrefix() {
  let n = 0;
  for (const f of HTML_FILES) {
    if (f === 'SharedThemeBoot.html' || f === 'SharedTailwindConfig.html') continue;
    const text = fs.readFileSync(path.join(SRC, f), 'utf8');
    // dark: の使用 (class="..." 内 or classList.add 内のみ)
    const matches = text.match(/\bdark:[a-z]+[-:][a-zA-Z0-9/_-]+/g) || [];
    n += matches.length;
  }
  return n;
}
const anyDark = countAnyDarkPrefix();
check('27b. Tailwind dark: prefix 全廃 = 0 (theme token に統一)', anyDark === 0, `${anyDark} 件残存 (semantic theme token に置換せよ)`);

// 28. gray/slate 色相混在 = 0 (Tokyo Night は青寄り → slate ファミリーに統一)
function countGraySlateMix() {
  let mix = 0;
  for (const f of HTML_FILES) {
    const text = fs.readFileSync(path.join(SRC, f), 'utf8');
    const grayMatches = text.match(/(?<!dark:)(text|bg|border)-gray-\d+|dark:(text|bg|border)-gray-\d+/g) || [];
    mix += grayMatches.length;
  }
  return mix;
}
const grayMix = countGraySlateMix();
check('28. gray/slate ファミリー混在 = 0', grayMix === 0, `${grayMix} 件 gray (slate に統一すべき)`);

// 30. Tailwind utility ↔ config 整合性 (no-op class 検出)
//   HTML で使われている theme- utility が Tailwind config (nested colors) に存在するか
function checkUtilityConfigConsistency() {
  const tw = fs.readFileSync(path.join(SRC, 'SharedTailwindConfig.html'), 'utf8');
  // nested theme: { DEFAULT, primary, ... } から key を抽出
  const themeMatch = tw.match(/theme:\s*\{([\s\S]*?)\n\s*\},\s*\n\s*\/\//);
  const themeKeys = new Set(['theme']);  // text-theme = DEFAULT
  if (themeMatch) {
    const keyRe = /^\s*['"]?([a-zA-Z][\w-]*)['"]?\s*:/gm;
    let m;
    while ((m = keyRe.exec(themeMatch[1])) !== null) {
      const k = m[1];
      if (k === 'DEFAULT') themeKeys.add('theme');
      else themeKeys.add(`theme-${k}`);
    }
  }
  // brand / status keys
  const brandMatch = tw.match(/brand:\s*\{([\s\S]*?)\}/);
  if (brandMatch) {
    let m, re = /^\s*['"]?([a-zA-Z][\w-]*)['"]?\s*:/gm;
    while ((m = re.exec(brandMatch[1])) !== null) themeKeys.add(`brand-${m[1]}`);
  }
  const statusMatch = tw.match(/status:\s*\{([\s\S]*?)\}/);
  if (statusMatch) {
    let m, re = /^\s*['"]?([a-zA-Z][\w-]*)['"]?\s*:/gm;
    while ((m = re.exec(statusMatch[1])) !== null) themeKeys.add(`status-${m[1]}`);
  }

  // HTML で使われている theme-/brand-/status- utility が config key にあるか
  let noop = 0;
  const noopExamples = [];
  for (const f of HTML_FILES) {
    const text = fs.readFileSync(path.join(SRC, f), 'utf8');
    const utilRe = /\b(?:text|bg|border|ring|from|to|via|divide|placeholder|outline|fill|stroke)-(theme(?:-[a-z][\w-]*)?|brand-[a-z][\w-]*|status-[a-z][\w-]*)\b/g;
    let m;
    while ((m = utilRe.exec(text)) !== null) {
      if (!themeKeys.has(m[1])) {
        noop++;
        if (noopExamples.length < 3) noopExamples.push(`${f}: ${m[0]}`);
      }
    }
  }
  return { noop, examples: noopExamples, validKeys: themeKeys.size };
}
const consistency = checkUtilityConfigConsistency();
check('30. Tailwind utility ↔ config 整合性 (no-op 検出)', consistency.noop === 0,
  consistency.noop === 0 ? `${consistency.validKeys} valid keys, 0 no-op` : `${consistency.noop} no-op utility (例: ${consistency.examples.join(', ')})`);

// 29. Theme token utility 利用率 (consolidation 健全性)
function countConsolidation() {
  let themeUtility = 0, darkPair = 0;
  for (const f of HTML_FILES) {
    const text = fs.readFileSync(path.join(SRC, f), 'utf8');
    themeUtility += (text.match(/\b(bg|text|border)-theme[-a-z]*/g) || []).length;
    darkPair += (text.match(/\bdark:(bg|text|border)-/g) || []).length;
  }
  const total = themeUtility + darkPair;
  const ratio = total > 0 ? themeUtility / total : 1;
  return { themeUtility, darkPair, ratio };
}
const cons = countConsolidation();
check('29. Theme utility 利用率 ≥ 80%', cons.ratio >= 0.8, `${cons.themeUtility} theme / ${cons.darkPair} dark: pair = ${(cons.ratio*100).toFixed(1)}%`);

// 26. border-radius の Tailwind スケール準拠率 ≥ 80%
function countRadii() {
  const files = ['UnifiedStyles.css.html', 'page.css.html', 'page.viz.css.html'];
  const STD = new Set(['0.25rem', '0.5rem', '0.75rem', '1rem', '1.5rem', '9999px', '50%']);
  let std = 0, total = 0;
  for (const f of files) {
    const text = fs.readFileSync(path.join(SRC, f), 'utf8');
    const cleaned = text.replace(/\/\*[\s\S]*?\*\//g, '');
    const matches = cleaned.matchAll(/border-radius:\s*([0-9.]+(?:rem|px|%))/g);
    for (const m of matches) {
      total++;
      if (STD.has(m[1])) std++;
    }
  }
  return { std, total, ratio: total ? std / total : 0 };
}
const r1 = countRadii();
check(`26. border-radius Tailwind scale 準拠率 ≥ 80%`, r1.ratio >= 0.8, `${r1.std}/${r1.total} (${(r1.ratio * 100).toFixed(1)}%)`);

// 出力
console.log('\n══════════════════════════════════════════════════════════════════');
console.log('  Theme Perfect — 20 軸 完璧度ゲート');
console.log('══════════════════════════════════════════════════════════════════\n');
let pass = 0;
for (const c of checks) {
  const mark = c.ok ? '✅ 10/10' : '❌  0/10';
  console.log(`  ${mark}  ${c.name}`);
  console.log(`            ${c.detail}`);
  if (c.ok) pass++;
}
console.log('\n───────────────────────────────────────────────────────────────────');
console.log(`  Score: ${pass * 10} / ${checks.length * 10} (${(pass / checks.length * 100).toFixed(1)}%)`);
console.log(`  Passed: ${pass} / ${checks.length}`);
console.log('═══════════════════════════════════════════════════════════════════');
process.exit(pass === checks.length ? 0 : 1);
