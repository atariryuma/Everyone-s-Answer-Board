#!/usr/bin/env node
/**
 * gen-pagehead.js — 全ページ共通の <head> を、include を含まない 1 枚に平坦化する。
 *
 * Why: include() は本来テンプレート評価を必要としない。src/ の include 対象 17 ファイル
 *   (計 963KB: d3.min 273KB / page.js 184KB / AdminPanel.js 173KB / UnifiedStyles.css 91KB …) を
 *   調べると、変数を使う scriptlet は 1 つも無く、入れ子 include を持つのも SharedPageHead だけだった。
 *   にもかかわらず include() が createTemplateFromFile().evaluate() だったため、ページ描画のたびに
 *   963KB がテンプレートコンパイラを通っていた。さらに二つの潜在リスクを抱えていた:
 *
 *     (a) 循環 include — v2908 で実際に全ページ停止を起こしかけた
 *     (b) vendor ライブラリ (d3.min / tinySegmenter) が将来 "<?" を含むと誤解析されて壊れる
 *
 *   SharedPageHead を平坦化すれば「include 対象は scriptlet を一切持たない」が不変条件になり、
 *   include() を createHtmlOutputFromFile().getContent() に落とせる。(a) は構造的に発生不能になり、
 *   (b) はテンプレートエンジンを通らなくなるので消える。
 *
 * 読み込み順序はここが唯一の定義。CLAUDE.md の「include order は固定。変更には独立コミット」は
 * この配列を動かすことを指す。
 *
 * Usage:
 *   npm run gen:pagehead            # src/SharedPageHead.html を生成
 *   npm run gen:pagehead -- --check # 差分があれば exit 1 (CI 用)
 */
const fs = require('fs');
const path = require('path');

const SRC = path.resolve(__dirname, '../src');
const OUT = path.join(SRC, 'SharedPageHead.html');
const CHECK = process.argv.includes('--check');

// この順序に意味がある。動かすときは独立コミットで。
const ORDER = [
  ['SharedSecurityHeaders', 'CSP。他の何よりも先に効かせる'],
  ['SharedThemeBoot', 'paint 前に <html> へ theme class を付け FOUC を防ぐ'],
  ['UnifiedStyles.css', 'semantic token とプリミティブの本体'],
  ['UtilityStyles.css', 'Tailwind 互換の静的 utility (自動生成)。primitive を上書きできるよう必ず最後'],
];

const header = `<!-- =====================================================================
  SharedPageHead — 全ページ共通の <head> 定型

  【自動生成】このファイルは scripts/gen-pagehead.js が生成する。直接編集しない。
     中身を変えたいときは元ファイル (下記) を編集して \`npm run gen:pagehead\` を実行する。

  Why: 9 つのページテンプレートが同じ meta / base / viewport と 4 つの CSS/JS を各自で
    書いており、順序もページごとに揺れていた。順序の定義を 1 箇所に閉じ込め、規約を
    構造で強制する。さらに include を展開済みにすることで、この <head> 全体が
    テンプレート評価を必要としなくなり、include() が単なるファイル連結で済むようになる。

  読み込み順序 (scripts/gen-pagehead.js の ORDER が唯一の定義):
${ORDER.map(([n, why], i) => `    ${i + 1}. ${n.padEnd(22)} — ${why}`).join('\n')}

  含めないもの:
    - <title> : ページごとに異なるため各ページで書く
    - SharedIcons : SVG sprite は <use> 解決のため <body> 先頭に置く必要がある
===================================================================== -->
<meta charset="UTF-8" />
<base target="_top" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
`;

const parts = [header];
for (const [name] of ORDER) {
  const p = fs.existsSync(path.join(SRC, name)) ? path.join(SRC, name) : path.join(SRC, `${name}.html`);
  if (!fs.existsSync(p)) {
    console.error(`✗ 元ファイルが見つかりません: ${name}`);
    process.exit(1);
  }
  const body = fs.readFileSync(p, 'utf8');
  if (/<\?/.test(body)) {
    console.error(`✗ ${name} に scriptlet が含まれています。平坦化できません。`);
    process.exit(1);
  }
  parts.push(`<!-- ── ${name} ── -->\n`, body.endsWith('\n') ? body : `${body}\n`);
}
const out = parts.join('');

if (CHECK) {
  const current = fs.existsSync(OUT) ? fs.readFileSync(OUT, 'utf8') : '';
  if (current !== out) {
    console.error('✗ SharedPageHead.html が古くなっています。`npm run gen:pagehead` を実行してください。');
    process.exit(1);
  }
  console.log('✓ SharedPageHead.html は最新です。');
} else {
  fs.writeFileSync(OUT, out);
  console.log(`✓ SharedPageHead.html を生成しました (${(out.length / 1024).toFixed(0)} KB, ${ORDER.length} ファイルを展開)`);
}
