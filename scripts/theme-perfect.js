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
const { checkBrandAliasResidue } = require('./lib/brand-alias-check');
const { extractBlockBody } = require('./lib/theme-core');

const ROOT = path.resolve(__dirname, '..');
const SRC = path.join(ROOT, 'src');

const checks = [];
function check(name, ok, detail) { checks.push({ name, ok, detail }); }

const HTML_FILES = fs.readdirSync(SRC).filter(f => f.endsWith('.html') || f.endsWith('.js'));
const FILE_SET = new Set(HTML_FILES);

// SRC ファイル content を最初のアクセス時にだけ読む。 13 軸が同じファイルを独立 read
// していたのを 1 回に集約 (~25 reads × 13 ≈ 300 syscall → 25 syscall)。
const FILE_CACHE = new Map();
function readSrc(name) {
  let v = FILE_CACHE.get(name);
  if (v === undefined) {
    v = fs.readFileSync(path.join(SRC, name), 'utf8');
    FILE_CACHE.set(name, v);
  }
  return v;
}

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
  const text = readSrc(f);
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
  const text = readSrc(f);
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
  const text = readSrc(f);
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
// 子プロセスの score だけ拾うと「どの軸が落ちたか」が親のログに残らず、CI でしか
//   再現しない失敗 (npm バージョン差など) の診断ができない。失敗時のみ theme-verify の
//   ❌ 行を親のログへ転記する。
const verifyScoreText = scoreMatch ? `${scoreMatch[1]}/${scoreMatch[2]}` : 'unknown';
const verifyFailedAxes = (verify.stdout || '').split('\n').filter((l) => l.includes('❌')).join(' / ');
check('7. theme:verify 120/120', perfect120,
  perfect120 ? verifyScoreText : `${verifyScoreText} — 失敗軸: ${verifyFailedAxes || '(取得不可)'}`);

// 8. Unit tests 957 PASS
// npm を経由せず reporter を固定して叩く。既定レポーターは非 TTY だと Node 20=tap /
//   Node 22+=spec と変わり、出力形式 (`# pass N` / `ℹ pass N`) に依存した解析は
//   Node バージョンで壊れる (theme-verify axis 10 が CI で実際に false FAIL した)。
const testFileList = fs.readdirSync(path.join(ROOT, 'tests'))
  .filter((f) => f.endsWith('.test.cjs'))
  .map((f) => path.join('tests', f));
const tests = spawnSync('node', ['--test', '--test-reporter=tap', ...testFileList], { cwd: ROOT, encoding: 'utf8' });
const failMatch = (tests.stdout || '').match(/^#\s*fail\s+(\d+)/m);
const passMatch = (tests.stdout || '').match(/^#\s*pass\s+(\d+)/m);
const testsPass = failMatch && failMatch[1] === '0' && passMatch && parseInt(passMatch[1]) > 900;
check('8. Unit tests 957+ PASS', testsPass, `pass=${passMatch?passMatch[1]:'?'}, fail=${failMatch?failMatch[1]:'?'}`);

// 9. body.theme-light 上書きを CSS / HTML <style> で持つファイル数
let lightOverrideFiles = 0;
for (const f of HTML_FILES) {
  const text = readSrc(f);
  if (/body\.theme-light\b/.test(text)) lightOverrideFiles++;
}
check('9. body.theme-light override 定義あり', lightOverrideFiles >= 2, `${lightOverrideFiles} ファイル (UnifiedStyles + page.viz)`);

// 10. themeManager API export 完全 (get/set/toggle/subscribe/init/mountToggle)
const su = readSrc('SharedUtilities.html');
const apiBits = ['get(', 'set(', 'toggle(', 'subscribe(', 'init(', 'mountToggle(', 'resolved('];
const missing = apiBits.filter(b => !su.includes(b));
check('10. themeManager 全 API 実装', missing.length === 0, missing.length ? `missing: ${missing.join(',')}` : '7/7');

// 11. localStorage persistence key
const hasPersist = /['"]app-theme['"]/.test(su);
check('11. localStorage \'app-theme\' 永続化', hasPersist, hasPersist ? '✓' : '✗');

// 12. theme token が utility class として公開されているか
//   v2901: Tailwind CDN を廃止し、使用クラスだけを静的生成する方式へ移行した。
//   検査対象を Tailwind config → 生成済み UtilityStyles.css.html に切り替える。
//   light/dark の切替は --theme-* token と body.theme-light が担うので darkMode 設定は不要。
const tw = readSrc('UtilityStyles.css.html');
const hasThemeColors = /var\(--theme-bg-base\)|var\(--theme-text-primary\)/.test(tw);
// CDN 由来の <script src> / <link href> が 1 つも無いこと (コメント内の言及は対象外)。
let cdnTags = 0;
for (const f of HTML_FILES) {
  if (!f.endsWith('.html')) continue;
  const t = readSrc(f);
  cdnTags += (t.match(/<(?:script|link)[^>]+(?:src|href)="https?:\/\/(?!script\.google\.com)[^"]+"/g) || []).length;
}
check('12. theme token が utility class として公開 (CDN 非依存)', hasThemeColors && cdnTags === 0, `tokens=${hasThemeColors}, 外部読込タグ=${cdnTags}`);

// 13. apply() で .dark + .light + theme-{dark,light} を html/body に
// v2810+: 統一形 classList.add(resolved, 'theme-' + resolved) を採用
const hasResolved = /classList\.add\(resolved,\s*'theme-'\s*\+\s*resolved\)/.test(su)
  || (/classList\.add\(.{0,30}'dark'/.test(su) && /classList\.add\(.{0,30}'light'/.test(su));
const hasThemeClass = /classList\.add\(.{0,30}'theme-'\s*\+\s*resolved\)/.test(su) || /classList\.add\('theme-dark'/.test(su);
const hasColorScheme = /style\.colorScheme/.test(su);
check('13. apply() で .dark/.light/.theme-* + colorScheme 同期', hasResolved && hasThemeClass && hasColorScheme, `resolved=${hasResolved}, theme-class=${hasThemeClass}, colorScheme=${hasColorScheme}`);

// 14. v2849+: --brand-{background,surface,text,border} alias 廃止チェック (共通ヘルパー)。
//     theme-verify (axis 8c) と同じロジックを scripts/lib/brand-alias-check で共有。
const us = readSrc('UnifiedStyles.css.html');  // 後続 axis 15/16/21/23/24 で再利用
const aliasRes = checkBrandAliasResidue({ srcDir: SRC, htmlFiles: HTML_FILES, reader: readSrc });
check('14. brand alias 不在 (--theme-* 直参照に統一)',
  aliasRes.aliasFree,
  aliasRes.aliasFree ? '✓ 0 件' : `def=${aliasRes.hasDef}, usage=${aliasRes.usageFiles.join(',')}`);

// 15. card tier tokens 完備 (--theme-card-1/2/3 が dark + light 両方に)
//   旧実装は body.theme-light から 3000 文字の固定窓で検索していたが、light ブロックが
//   6000 文字超に育ち card token (実距離 3202 文字) を取りこぼして false FAIL していた。
//   窓を広げるのではなくブロック本体を正しく切り出して判定する (誤検知でゲートの信頼を落とさない)。
const CARD_TIER_TOKENS = ['--theme-card-1:', '--theme-card-2:', '--theme-card-3:'];
const darkTokenBlock = extractBlockBody(us, /:root\s*\{/);
const lightTokenBlock = extractBlockBody(us, /body\.theme-light[^{]*\{/);
const darkCardDef = CARD_TIER_TOKENS.every((t) => darkTokenBlock.includes(t));
const lightCardDef = CARD_TIER_TOKENS.every((t) => lightTokenBlock.includes(t));
check('15. --theme-card-1/2/3 dark+light 両定義', darkCardDef && lightCardDef, `dark=${darkCardDef}, light=${lightCardDef}`);

// 16. high contrast media query dark+light両方
const hcDark = /@media \(prefers-contrast: high\)[\s\S]*?:root\s*\{[\s\S]*?--theme/.test(us);
const hcLight = /@media \(prefers-contrast: high\)[\s\S]*?body\.theme-light/.test(us);
check('16. prefers-contrast:high override dark+light', hcDark && hcLight, `dark=${hcDark}, light=${hcLight}`);

// 17. theme-color meta tag に theme:exempt or theme-aware — same line exempt comment も対応
let untaggedMeta = 0;
for (const f of HTML_FILES) {
  const text = readSrc(f);
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
  const text = readSrc(f);
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
const vizCss = readSrc('page.viz.css.html');
const lightOverrideCount = (vizCss.match(/body\.theme-light\s+\./g) || []).length;
check('19. page.viz light override ≥ 9 件', lightOverrideCount >= 9, `${lightOverrideCount} 件`);

// 20. メンテナンスフローが記載 (CLAUDE.md または docs/THEME.md)
//   2026-07 の docs 再編で theme の詳細 (保守 CLI フロー含む) を CLAUDE.md から
//   docs/THEME.md へ移譲し、CLAUDE.md からはリンクで辿る構成にした。本 axis の意図は
//   「フローが文書化され追跡可能なこと」なので、移譲先の docs/THEME.md も判定対象に含める。
const hasMaintenanceFlow = (text) =>
  /theme:matrix.*theme:uncovered.*theme:tokenize/s.test(text) ||
  /theme-card-1.*theme-card-2.*theme-card-3/s.test(text);
const claudemd = fs.readFileSync(path.join(ROOT, 'CLAUDE.md'), 'utf8');
const themeDocPath = path.join(ROOT, 'docs', 'THEME.md');
const themeDoc = fs.existsSync(themeDocPath) ? fs.readFileSync(themeDocPath, 'utf8') : '';
const hasFlow = hasMaintenanceFlow(claudemd) || hasMaintenanceFlow(themeDoc);
check('20. メンテナンスフロー記載 (CLAUDE.md / docs/THEME.md)', hasFlow, hasFlow ? '✓' : '✗');

// 21. color-scheme CSS property (両モード定義済)
const hasCsRoot = /:root\s*\{[\s\S]*?color-scheme:\s*dark/.test(us);
const hasCsLight = /body\.theme-light[\s\S]{0,2000}?color-scheme:\s*light/.test(us);
check('21. color-scheme CSS property (dark+light)', hasCsRoot && hasCsLight, `root=${hasCsRoot}, light=${hasCsLight}`);

// 22. FOUC 防止: SharedThemeBoot が全 theme-aware ページに届いているか
//   v2899 で各ページの <head> 定型を SharedPageHead に集約したため、ページは
//   SharedThemeBoot を直接ではなく SharedPageHead 経由で読む。どちらの経路でも
//   「paint 前に theme class が付く」という目的は満たされるので両方を許容する。
const bootExists = FILE_SET.has('SharedThemeBoot.html');
const headBundlesBoot = FILE_SET.has('SharedPageHead.html')
  && /SharedThemeBoot/.test(readSrc('SharedPageHead.html'));
const PAGES_NEEDING_BOOT = ['Page', 'AdminPanel', 'LoginPage', 'AppSetupPage', 'SetupPage',
                            'TeacherManual', 'Unpublished', 'AccessRestricted', 'ErrorBoundary'];
let bootMissing = 0;
for (const p of PAGES_NEEDING_BOOT) {
  const name = `${p}.html`;
  if (!FILE_SET.has(name)) continue;
  const src = readSrc(name);
  const reachesBoot = /SharedThemeBoot/.test(src)
    || (headBundlesBoot && /SharedPageHead/.test(src));
  if (!reachesBoot) bootMissing++;
}
check('22. FOUC 防止 SharedThemeBoot 全ページ include', bootExists && bootMissing === 0, `boot=${bootExists}, via SharedPageHead=${headBundlesBoot}, missing=${bootMissing}`);

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
    const text = readSrc(f);
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
    const text = readSrc(f);
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
    if (f === 'SharedThemeBoot.html') continue;
    const text = readSrc(f);
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
    const text = readSrc(f);
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
  // v2901: Tailwind config を廃止したので、実際に CSS で定義されたルールを正とする。
  //   theme utility は 2 箇所に分かれる:
  //     - UnifiedStyles.css.html : 以前から手書きしていた status/soft 系
  //     - UtilityStyles.css.html : gen-utilities.js が生成する分
  const tw = readSrc('UtilityStyles.css.html') + '\n' + readSrc('UnifiedStyles.css.html');
  const themeKeys = new Set(['theme']);
  {
    // `.text-theme-x` だけでなく `.hover\:text-theme-x:hover` のような
    //   variant 付きセレクタも拾えるよう、直前が `.` または `\:` を許す。
    const keyRe = /[.:](?:text|bg|border|ring|divide)-(theme[\w-]*?)(?:\\\/\d+)?(?=[\s{:,])/g;
    let m;
    while ((m = keyRe.exec(tw)) !== null) {
      themeKeys.add(m[1]);
    }
  }
  // brand- / status- も同様に生成済みルールから拾う
  {
    const keyRe = /[.:](?:text|bg|border|ring|divide)-((?:brand|status)[\w-]*?)(?:\\\/\d+)?(?=[\s{:,])/g;
    let m;
    while ((m = keyRe.exec(tw)) !== null) themeKeys.add(m[1]);
  }

  // HTML で使われている theme-/brand-/status- utility が config key にあるか
  let noop = 0;
  const noopExamples = [];
  for (const f of HTML_FILES) {
    const text = readSrc(f);
    // `--border-status-success` のような CSS カスタムプロパティ名は utility class ではない。
    //   直前が `-` (= `--` の一部) の場合は除外する。
    const utilRe = /(^|[^-\w])(?:text|bg|border|ring|from|to|via|divide|placeholder|outline|fill|stroke)-(theme(?:-[a-z][\w-]*)?|brand-[a-z][\w-]*|status-[a-z][\w-]*)\b/g;
    let m;
    while ((m = utilRe.exec(text)) !== null) {
      if (!themeKeys.has(m[2])) {
        noop++;
        if (noopExamples.length < 3) noopExamples.push(`${f}: ${m[0].trim()}`);
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
    const text = readSrc(f);
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
    const text = readSrc(f);
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

// 31. Status + Accent token rgba hardcoded = 0 (v2851+: color-mix() 経由に統一)
//   status feedback / brand accent / cyan-blue family の半透明利用はすべて
//   `color-mix(in srgb, var(--xxx) N%, transparent)` 経由。
//   token 1 箇所変えれば全 alpha variant + 全 light/dark variant が自動追従。
//   v2849 で cyan/blue 6 系列、 v2851 で status (success/warning/error/purple) +
//   destructive/indigo を追加し計 22 系列 を network。
//   exempt: token 定義行 / `theme:exempt` 注釈 / login.js.html (standalone popup) /
//           `var(--token, rgba(...))` fallback (defensive, allowed)。
const PRIMARY_RGBA_TARGETS = [
  // cyan / blue 系 (v2849-v2850)
  { re: /rgba\(56,\s*189,\s*248/,  alt: '--theme-accent-cyan' },         // sky-400
  { re: /rgba\(59,\s*130,\s*246/,  alt: '--brand-primary' },             // blue-500
  { re: /rgba\(34,\s*211,\s*238/,  alt: '--theme-accent-cyan-strong' },  // cyan-400
  { re: /rgba\(8,\s*145,\s*178/,   alt: '--theme-accent-cyan-deep' },    // cyan-600
  { re: /rgba\(96,\s*165,\s*250/,  alt: '--brand-primary-light' },       // blue-400
  { re: /rgba\(103,\s*232,\s*249/, alt: '--theme-accent-cyan-hover' },   // cyan-300
  { re: /rgba\(14,\s*116,\s*144/,  alt: '--theme-accent-cyan-deep' },    // cyan-700 (light value)
  { re: /rgba\(139,\s*233,\s*253/, alt: '--brand-accent-light' },        // light cyan
  // status feedback box (v2851 — bg + border ペアで token 化)
  { re: /rgba\(34,\s*197,\s*94/,   alt: '--status-success' },            // green-500
  { re: /rgba\(16,\s*185,\s*129/,  alt: '--status-success' },            // emerald-500
  { re: /rgba\(187,\s*247,\s*208/, alt: '--status-success' },            // green-200
  { re: /rgba\(21,\s*128,\s*61/,   alt: '--status-success' },            // emerald-700 (light)
  { re: /rgba\(74,\s*222,\s*128/,  alt: '--status-success' },            // green-400
  { re: /rgba\(245,\s*158,\s*11/,  alt: '--status-warning' },            // amber-500
  { re: /rgba\(251,\s*191,\s*36/,  alt: '--brand-accent' },              // amber-400
  { re: /rgba\(253,\s*224,\s*71/,  alt: '--theme-status-warning-strong' }, // yellow-300
  { re: /rgba\(239,\s*68,\s*68/,   alt: '--status-error' },              // red-500
  { re: /rgba\(248,\s*113,\s*113/, alt: '--status-error' },              // red-400
  { re: /rgba\(168,\s*85,\s*247/,  alt: '--brand-highlight' },           // purple-500
  { re: /rgba\(147,\s*51,\s*234/,  alt: '--brand-highlight' },           // purple-600
  { re: /rgba\(192,\s*132,\s*252/, alt: '--brand-highlight-alt' },       // purple-400
  // accent: destructive / indigo (v2851 — 新規追加 token)
  { re: /rgba\(234,\s*88,\s*12/,   alt: '--theme-accent-destructive' },  // orange-600
  { re: /rgba\(99,\s*102,\s*241/,  alt: '--theme-accent-indigo' },       // indigo-500
];
// standalone popup window で書かれており theme system を持たない (window 外なので
// CSS 変数が定義されていない)。 axis 31 のスキャン対象外。
const RGBA_EXEMPT_FILES = new Set(['login.js.html']);

function countPrimaryRgbaHardcoded() {
  const violations = [];
  for (const f of HTML_FILES) {
    if (RGBA_EXEMPT_FILES.has(f)) continue;
    const lines = readSrc(f).split('\n');
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      if (/^\s*--(?:theme|brand)-/.test(line)) continue;
      if (/theme:exempt/.test(line)) continue;
      if (/var\(--[a-z0-9-]+,\s*rgba/.test(line)) continue;  // defensive fallback OK
      if (/^\s*\*/.test(line)) continue;                      // CSS multi-line comment continuation
      for (const t of PRIMARY_RGBA_TARGETS) {
        if (t.re.test(line)) violations.push(`${f}:${i+1} → ${t.alt}`);
      }
    }
  }
  return violations;
}
const primaryRgba = countPrimaryRgbaHardcoded();
check('31. Status + Accent rgba hardcoded = 0 (color-mix 統一)',
  primaryRgba.length === 0,
  primaryRgba.length === 0
    ? '✓ 0 件 (color-mix 経由に統一)'
    : `${primaryRgba.length} 件 残存 (例: ${primaryRgba.slice(0,3).join(', ')})`);

// 出力
console.log('\n══════════════════════════════════════════════════════════════════');
console.log(`  Theme Perfect — ${checks.length} 軸 完璧度ゲート`);
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
