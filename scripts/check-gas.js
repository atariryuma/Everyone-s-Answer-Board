#!/usr/bin/env node
/**
 * check-gas.js — GAS 特有の「テストでは落ちないが本番で落ちる」参照切れを静的検査する。
 *
 * Why: このアプリは単一グローバルスコープで動くため、モジュール解決の仕組みが無い。
 *   関数名を 1 つ消すと、それを呼ぶ別ファイルは何のエラーも出さずに本番で
 *   ReferenceError / undefined is not a function を投げる。node --test は各ファイルを
 *   個別に vm へ流し込むので、ファイル間の参照切れを構造的に検出できない。
 *   デプロイ前にここを機械的に潰す。
 *
 * 検査項目:
 *   1. include('X') の対象ファイルが実在するか
 *   2. runServer / google.script.run が呼ぶ関数がサーバ側に実在するか
 *   3. HTML テンプレートの scriptlet <?!= fn() ?> が呼ぶ関数が実在するか
 *   4. /* global ... *\/ で宣言された名前が実際にどこかで定義されているか
 *   5. doPost の action allowlist と dispatch テーブルが一致しているか
 *   6. HTML の要素 ID を JS が参照しているとき、その ID が存在するか
 *   7. appsscript.json が JSON として妥当か
 *   8. *.css.html の中身が <style> の外にはみ出していないか
 *   9. HTML コメント内に scriptlet が残っていないか (GAS はコメント内も評価する)
 *  10. include 対象が scriptlet を持たないか (include はテンプレート評価しない)
 *  11. <use href="#i-X"> の X が SVG sprite に定義されているか
 *  12. アイコンを使うページが sprite を include しているか
 *
 * 使い方:
 *   npm run check:gas          # 全項目
 *   node scripts/check-gas.js  # 同上
 *
 * 終了コード: 問題があれば 1。CI / pre-deploy ゲートに使える。
 */

'use strict';
const fs = require('fs');
const path = require('path');

const SRC = path.resolve(__dirname, '../src');
const files = fs.readdirSync(SRC);
const jsFiles = files.filter((f) => f.endsWith('.js'));
// d3 / tinySegmenter は外部ライブラリなので解析対象外。
const VENDOR = new Set(['d3.min.html', 'tinySegmenter.html']);
const htmlFiles = files.filter((f) => f.endsWith('.html') && !VENDOR.has(f));

const read = (f) => fs.readFileSync(path.join(SRC, f), 'utf8');

// コメント内の記述を「実際の呼び出し」と誤認しないための除去。
//   例: JSDoc の使用例 `themeManager.mountToggle(document.getElementById('header-actions'))`
//   長さを保つため中身を空白に置換し、行番号がずれないようにする。
function stripComments(src) {
  return src
    .replace(/\/\*[\s\S]*?\*\//g, (m) => m.replace(/[^\n]/g, ' '))
    .replace(/(^|[^:])\/\/[^\n]*/g, (m, p1) => p1 + ' '.repeat(m.length - p1.length));
}
const jsSrc = new Map(jsFiles.map((f) => [f, read(f)]));
const htmlSrc = new Map(htmlFiles.map((f) => [f, read(f)]));
const allSrc = [...jsSrc.values(), ...htmlSrc.values()].join('\n');

const problems = [];
const results = [];
function check(name, ok, detail) {
  results.push({ name, ok, detail });
  if (!ok) problems.push(`${name}: ${detail}`);
}

// HTML コメントを同じ長さの空白に潰す (行番号を保つため)。
//   コメント内の記述を「実際のコード」と誤認しないために各所で使う。
const blankComments = (src) => src.replace(/<!--[\s\S]*?-->/g, (m) => m.replace(/[^\n]/g, ' '));

// scriptlet 内の include('x') が指す名前を列挙する。
//   Why: 同じ 3 行 (コメント潰し → matchAll → .html 落とし) を項目 10/11/13 が
//   それぞれ書いており、片方だけ直すと検査がずれる状態だった。
const includesOf = (src) => [...blankComments(src)
  .matchAll(/<\?[^?]*?include\(\s*'([^']+)'/g)].map((m) => m[1].replace(/\.html$/, ''));

// 行番号の算出（エラーメッセージ用）
function lineOf(src, index) {
  let n = 1;
  for (let i = 0; i < index; i++) if (src.charCodeAt(i) === 10) n++;
  return n;
}

// ── 1. include() の対象が実在するか ────────────────────────────────
{
  const missing = [];
  for (const [f, s] of [...jsSrc, ...htmlSrc]) {
    for (const m of s.matchAll(/include\(\s*'([^']+)'/g)) {
      const name = m[1];
      const exists = fs.existsSync(path.join(SRC, name)) || fs.existsSync(path.join(SRC, `${name}.html`));
      if (!exists) missing.push(`${f}:${lineOf(s, m.index)} include('${name}')`);
    }
  }
  check('1. include() の対象が実在する', missing.length === 0, missing.join(', '));
}

// ── 2. client が呼ぶサーバ関数が実在するか ──────────────────────────
// サーバ側のグローバル関数（行頭 function 宣言）を収集。
const serverFns = new Set();
for (const s of jsSrc.values()) {
  for (const m of s.matchAll(/^function\s+([A-Za-z_$][\w$]*)\s*\(/gm)) serverFns.add(m[1]);
}
{
  const missing = [];
  for (const [f, raw] of htmlSrc) {
    const s = stripComments(raw);
    const calls = new Map();
    // runServer('name', ...) — 本アプリの標準経路
    for (const m of s.matchAll(/runServer\(\s*['"]([A-Za-z_$][\w$]*)['"]/g)) {
      calls.set(m[1], lineOf(s, m.index));
    }
    // google.script.run[.withXxx(...)].name(...) — 生の GAS API 経路。
    //   withSuccessHandler / withFailureHandler 等の builder は呼び出し対象ではないので除く。
    const BUILDER = /^with[A-Z]/;
    for (const m of s.matchAll(/google\.script\.run((?:\s*\.\s*[A-Za-z_$][\w$]*\s*\([^()]*\))+)/g)) {
      const last = [...m[1].matchAll(/\.\s*([A-Za-z_$][\w$]*)\s*\(/g)].map((x) => x[1]).pop();
      if (last && !BUILDER.test(last)) calls.set(last, lineOf(s, m.index));
    }
    for (const [name, line] of calls) {
      if (!serverFns.has(name)) missing.push(`${f}:${line} ${name}()`);
    }
  }
  check('2. client が呼ぶサーバ関数が実在する', missing.length === 0, missing.join(', '));
}

// ── 3. テンプレート scriptlet が呼ぶ関数が実在するか ─────────────────
{
  const missing = [];
  for (const [f, s] of htmlSrc) {
    // scriptlet は任意の JS を書けるので、制御構文・組み込みは呼び出し対象から除く。
    const JS_KEYWORDS = new Set(['if', 'for', 'while', 'switch', 'return', 'typeof', 'catch',
      'function', 'else', 'do', 'new', 'delete', 'void', 'String', 'Number', 'Boolean',
      'Array', 'Object', 'JSON', 'Math', 'Date']);
    for (const m of s.matchAll(/<\?!?=?\s*([A-Za-z_$][\w$]*)\s*\(/g)) {
      const name = m[1];
      if (name === 'include' || JS_KEYWORDS.has(name)) continue; // 1 で検査済 / 構文
      if (!serverFns.has(name)) missing.push(`${f}:${lineOf(s, m.index)} <?= ${name}() ?>`);
    }
  }
  check('3. テンプレート scriptlet の関数が実在する', missing.length === 0, missing.join(', '));
}

// ── 4. /* global ... */ 宣言が実体を伴うか ──────────────────────────
{
  // 定義済みシンボル = 関数宣言 + const/let/var のトップレベル宣言
  const defined = new Set(serverFns);
  for (const s of jsSrc.values()) {
    for (const m of s.matchAll(/^(?:const|let|var)\s+([A-Za-z_$][\w$]*)/gm)) defined.add(m[1]);
  }
  const missing = [];
  for (const [f, s] of jsSrc) {
    const m = s.match(/\/\*\s*global\s+([^*]+)\*\//);
    if (!m) continue;
    for (const raw of m[1].split(',')) {
      const name = raw.trim();
      if (!name) continue;
      if (!defined.has(name)) missing.push(`${f} が宣言する ${name}`);
    }
  }
  check('4. /* global */ 宣言に実体がある', missing.length === 0, missing.join(', '));
}

// ── 5. doPost の allowlist と dispatch テーブルが一致するか ──────────
{
  const main = jsSrc.get('main.js') || '';
  const allowMatch = main.match(/const allowedActions\s*=\s*\[([^\]]+)\]/);
  const tableMatch = main.match(/DO_POST_ACTION_HANDLERS\s*=\s*(?:Object\.freeze\()?\{([\s\S]*?)\n\}/);
  if (!allowMatch) {
    check('5. doPost の allowlist と handler が対応', false, 'allowedActions が見つからない');
  } else {
    const allowed = [...allowMatch[1].matchAll(/'([^']+)'/g)].map((m) => m[1]);
    const handled = tableMatch ? [...tableMatch[1].matchAll(/^\s*([A-Za-z_$][\w$]*)\s*:/gm)].map((m) => m[1]) : [];
    // allowlist にあるが handler も個別分岐も無い action を検出。
    //   setupApiKey / adminApi は doPost 内で早期 return する特別扱いなので除外。
    const SPECIAL = new Set(['setupApiKey', 'adminApi']);
    const orphan = allowed.filter((a) => !SPECIAL.has(a) && !handled.includes(a));
    // 逆向き: handler があるのに allowlist に無い = 到達不能
    const unreachable = handled.filter((h) => !allowed.includes(h));
    const detail = [
      orphan.length ? `handler 無し: ${orphan.join(', ')}` : '',
      unreachable.length ? `到達不能な handler: ${unreachable.join(', ')}` : ''
    ].filter(Boolean).join(' / ');
    check('5. doPost の allowlist と handler が対応', detail === '', detail);
  }
}

// ── 6. JS が参照する要素 ID が HTML に存在するか ─────────────────────
{
  // HTML 側で定義される ID をすべて収集（テンプレート生成分も含む）
  const definedIds = new Set();
  for (const s of htmlSrc.values()) {
    for (const m of s.matchAll(/id="([A-Za-z0-9_-]+)"/g)) definedIds.add(m[1]);
    // JS が動的生成する要素 (el.id = 'x' / id="${...}" 直書き) も定義とみなす
    for (const m of s.matchAll(/\.id\s*=\s*['"]([A-Za-z0-9_-]+)['"]/g)) definedIds.add(m[1]);
    for (const m of s.matchAll(/setAttribute\(\s*'id'\s*,\s*'([A-Za-z0-9_-]+)'/g)) definedIds.add(m[1]);
  }
  const missing = [];
  for (const [f, raw] of htmlSrc) {
    const s = stripComments(raw);
    for (const m of s.matchAll(/getElementById\(\s*'([A-Za-z0-9_-]+)'\s*\)/g)) {
      if (!definedIds.has(m[1])) missing.push(`${f}:${lineOf(s, m.index)} #${m[1]}`);
    }
  }
  check('6. JS が参照する要素 ID が存在する', missing.length === 0, missing.join(', '));
}

// ── 7. appsscript.json が妥当か ────────────────────────────────────
{
  const p = path.join(SRC, 'appsscript.json');
  let ok = false; let detail = 'appsscript.json が無い';
  if (fs.existsSync(p)) {
    try {
      const manifest = JSON.parse(fs.readFileSync(p, 'utf8'));
      ok = typeof manifest.timeZone === 'string' && manifest.runtimeVersion === 'V8';
      detail = ok ? '' : `timeZone/runtimeVersion が不正 (runtime=${manifest.runtimeVersion})`;
    } catch (e) {
      detail = `JSON parse 失敗: ${e.message}`;
    }
  }
  check('7. appsscript.json が妥当', ok, detail);
}

// ── 8. *.css.html が <style> の外にはみ出していないか ──────────────
// Why: これらは <head> 内に include される CSS 断片。ファイル途中で </style> が
//   閉じると、それ以降の CSS は「ただのテキスト」になる。ブラウザは <head> 内の
//   非空白テキストを見た時点で head を打ち切って body に送るため、
//     (a) その CSS 規則がすべて無効になる
//     (b) CSS のソースが画面上部に生テキストで表示される
//     (c) 後続の include (page.viz.css 等) まで body へ押し出される
//   という三重の事故になる。しかも lint も test も theme ゲートも
//   「ファイル内にその規則が書かれている」ことしか見ないので全部 pass する。
//   実際 v2905/v2906 で追加した .answers-grid 等 80 行がこの状態だった。
{
  const cssFiles = htmlFiles.filter((f) => f.endsWith('.css.html'));
  const bad = [];
  for (const f of cssFiles) {
    const src = htmlSrc.get(f);
    // <style> ブロックと HTML コメント (ファイル冒頭の説明) は正当。それ以外が
    // 残れば、CSS がテキストとして <head> に流れ出している。
    // 位置を報告するため、除去部分は同じ長さの空白に置換して行番号を保つ。
    const blank = (m) => m.replace(/[^\n]/g, ' ');
    const outside = src
      .replace(/<style\b[^>]*>[\s\S]*?<\/style>/gi, blank)
      .replace(/<!--[\s\S]*?-->/g, blank);
    if (outside.trim()) {
      const strayLines = outside.split('\n').filter((l) => l.trim()).length;
      bad.push(`${f}: ${strayLines} 行が <style> の外 (${lineOf(src, outside.search(/\S/))} 行目付近)`);
    }
  }
  check('8. *.css.html が <style> の中に収まっている', bad.length === 0, bad.join(', '));
}

// ── 9. HTML コメント内に生きた scriptlet が残っていないか ────────────
// Why: ページ本体は doGet が createTemplateFromFile().evaluate() で描画するため、
//   テンプレートエンジンは HTML より先に走る。つまり <!-- --> の中に書いた
//   <? ... ?> も「コメントだから無害」ではなく、そのまま評価される。
//   使い方の説明をファイル冒頭のコメントに書く運用と相性が最悪で、用例のつもりで
//   <?!= include('SelfName'); ?> と書くと自分自身を無限 include して全ページが落ちる。
//   SharedThemeBoot はこれを知っていて用例から記法を外しているが、規約が人の記憶に
//   依存していたため SharedPageHead (v2899) で再発した。構造で止める。
{
  const bad = [];
  for (const [f, src] of htmlSrc) {
    for (const c of src.matchAll(/<!--[\s\S]*?-->/g)) {
      for (const s of c[0].matchAll(/<\?[\s\S]*?\?>/g)) {
        const at = c.index + s.index;
        bad.push(`${f}:${lineOf(src, at)} ${s[0].trim().slice(0, 60)}`);
      }
    }
  }
  check('9. HTML コメント内に生きた scriptlet がない', bad.length === 0, bad.join(', '));
}

// ── 10. include 対象が scriptlet を持たないか ───────────────────────
// Why: include() は createHtmlOutputFromFile().getContent() (main.js) で、テンプレート評価を
//   しない。これは include 対象 13 ファイル (計 963KB) が scriptlet を 1 つも持たないことを
//   前提にした最適化で、この前提が崩れると scriptlet が生テキストとして画面に出る。
//   9/10 が「評価される側」の事故を防ぐのに対し、これは「評価されない側」の前提を守る。
//   入れ子 include もここで弾かれるため、include の深さは常に 1 で、循環は構造的に
//   発生し得ない。以前は循環検出を別項目で持っていたが、この項目に完全に包含される
//   (常に true を返すだけの) 検査だったので畳んだ。
{
  const targets = new Set();
  for (const [, src] of [...jsSrc, ...htmlSrc]) for (const n of includesOf(src)) targets.add(n);
  const bad = [];
  for (const name of targets) {
    const f = `${name}.html`;
    const src = htmlSrc.get(f);
    if (src === undefined) continue; // 1 が実在チェックを担当
    const s = src.match(/<\?/);
    if (s) bad.push(`${f}:${lineOf(src, s.index)} include 対象に scriptlet がある`);
  }
  check('10. include 対象が scriptlet を持たない', bad.length === 0, bad.join(', '));
}

// ── 11. SVG sprite の参照先が実在するか ─────────────────────────
// Why: <use href="#i-xxx"> は参照先が無くても例外を出さず、ただ「何も描画されない」。
//   コンソールにも出ないので、アイコンだけが消えた画面が本番に出ても気付けない。
//   実際 i-alert-triangle (ログイン画面とエラー画面の警告)、i-grid-2x2 (列設定)、
//   i-link (アプリ設定) の 3 つが参照だけ残って欠落していた。
//   getIcon(name) は PAIRED な名前だけ -outline/-solid に展開するため、その規則も再現する。
{
  const sprite = htmlSrc.get('SharedIcons.html') || '';
  const defined = new Set([...sprite.matchAll(/<symbol[^>]*\bid="([^"]+)"/g)].map((m) => m[1]));
  // PAIRED 集合は page.js の getIcon() 定義から読む。書き写すと、対に増減があったとき
  //   検査だけが古い規則で展開し、存在しない symbol を「ある」と誤判定する。
  const pairedLiteral = (htmlSrc.get('page.js.html') || '').match(/PAIRED\s*=\s*new Set\(\[([^\]]*)\]/);
  const PAIRED = new Set(pairedLiteral ? [...pairedLiteral[1].matchAll(/'([^']+)'/g)].map((m) => m[1]) : []);
  if (!pairedLiteral) problems.push('11. page.js.html の PAIRED 定義を読めなかった (getIcon の実装が変わった?)');
  const missing = [];
  const want = (id, f, s, idx) => {
    if (!defined.has(id)) missing.push(`${f}:${lineOf(s, idx)} #${id}`);
  };
  for (const [f, s] of [...jsSrc, ...htmlSrc]) {
    for (const m of s.matchAll(/getIcon\(\s*'([a-zA-Z0-9_-]+)'/g)) {
      const n = m[1];
      if (PAIRED.has(n)) { want(`i-${n}-outline`, f, s, m.index); want(`i-${n}-solid`, f, s, m.index); }
      else want(`i-${n}`, f, s, m.index);
    }
    if (f === 'SharedIcons.html') continue; // sprite 自身の内部参照は対象外
    for (const m of s.matchAll(/href="#(i-[a-zA-Z0-9_-]+)"/g)) want(m[1], f, s, m.index);
  }
  check('11. SVG sprite の参照先が実在する', missing.length === 0, missing.join(', '));
}

// ── 12. アイコンを使うページが sprite を読み込んでいるか ─────────
// Why: symbol が定義されていても、そのページが SharedIcons を include していなければ
//   <use> は解決先を見つけられず、やはり無言で何も描画されない。12 とは別の失敗経路。
//   実際 ErrorBoundary と Unpublished が参照だけ持っていて sprite を読んでいなかった。
{
  // ページの列挙は直書きしない。新しいページを足した人がこの配列を更新し忘れると、
  //   検査は「全ページ OK」と言いながらそのページだけ素通しする = 無言で失敗するゲートになる。
  //   <body> を持つファイル = 単体で描画される画面、という構造から導出する。
  const PAGES = [...htmlSrc]
    .filter(([, src]) => /<body[\s>]/.test(src))
    .map(([f]) => f.replace(/\.html$/, ''));
  const depsOf = (name) => includesOf(htmlSrc.get(`${name}.html`) || '');
  const treeOf = (name, seen = new Set()) => {
    if (seen.has(name)) return seen;
    seen.add(name);
    for (const d of depsOf(name)) treeOf(d, seen);
    return seen;
  };
  const bad = [];
  for (const page of PAGES) {
    const tree = treeOf(page);
    const usesIcon = [...tree].some((n) => {
      if (n === 'SharedIcons') return false;
      const s = htmlSrc.get(`${n}.html`) || '';
      return /href="#i-/.test(s) || /getIcon\(/.test(s);
    });
    if (usesIcon && !tree.has('SharedIcons')) bad.push(`${page}.html が SharedIcons を include していない`);
  }
  check('12. アイコン使用ページが sprite を読み込む', bad.length === 0, bad.join(', '));
}

// ── 出力 ──────────────────────────────────────────────────────────
console.log('\nGAS 参照整合性チェック\n');
for (const r of results) {
  console.log(`  ${r.ok ? '✓' : '✗'} ${r.name}`);
  if (!r.ok && r.detail) {
    for (const line of r.detail.split(', ').slice(0, 20)) console.log(`      ${line}`);
    const total = r.detail.split(', ').length;
    if (total > 20) console.log(`      … 他 ${total - 20} 件`);
  }
}
console.log('');
if (problems.length) {
  console.log(`✗ ${problems.length} 項目に問題があります。デプロイ前に解消してください。\n`);
  process.exit(1);
}
console.log(`✓ 全 ${results.length} 項目 pass。参照切れなし。\n`);
