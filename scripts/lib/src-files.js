/**
 * src-files.js — 「このファイルは何か」の判定を 1 箇所に持つ。
 *
 * Why: 「生成物だから走査対象外」「vendor だから対象外」という同じ判断が、
 *   gen-utilities / tokenize-dimensions / theme-perfect / lint など 6 箇所で
 *   ファイル名の直書きリストとして重複していた。生成物を 1 つ増やすたびに
 *   6 箇所を探して回る必要があり、1 つ漏らすと「生成物を入力として走査する」
 *   フィードバックループが復活する (実際 v2910 で起きた)。
 *
 *   生成物は自分のヘッダーに【自動生成】と書いているので、名前ではなく中身で判定する。
 *   新しい生成物が増えても、そのヘッダーさえ書けば自動的に対象外になる。
 */
const fs = require('fs');
const path = require('path');

// 外部ライブラリ。中身は我々の管理外なので、どの検査でも対象にしない。
const VENDOR = new Set(['d3.min.html', 'tinySegmenter.html']);

/** 生成物マーカーはヘッダーコメントに置く決まりなので、先頭だけ見れば足りる。 */
function isGeneratedSource(text) {
  return /【自動生成】/.test(text.slice(0, 800));
}

function isVendor(file) {
  return VENDOR.has(path.basename(file));
}

/**
 * src/ の .html を列挙する。
 * @param {string} srcDir
 * @param {{ includeGenerated?: boolean, includeVendor?: boolean }} [opts]
 */
function listSourceHtml(srcDir, opts = {}) {
  const { includeGenerated = false, includeVendor = false } = opts;
  return fs.readdirSync(srcDir)
    .filter((f) => f.endsWith('.html'))
    .filter((f) => includeVendor || !isVendor(f))
    .filter((f) => {
      if (includeGenerated) return true;
      return !isGeneratedSource(fs.readFileSync(path.join(srcDir, f), 'utf8'));
    });
}

module.exports = { VENDOR, isGeneratedSource, isVendor, listSourceHtml };
