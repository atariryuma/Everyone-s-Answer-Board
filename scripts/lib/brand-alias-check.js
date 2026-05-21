'use strict';
/**
 * brand-alias-check — `--brand-{background,surface,text,border}` の
 * 「定義 + usage が 0 件」 を検証する共有ヘルパー。
 *
 * Why: v2849+ で 4 alias 廃止 → --theme-* 直接参照に統一。 alias は 2-hop indirection
 *      で読み手の認知負荷が高く、 実利用 0 件を確認した上で削除済。
 *      theme:perfect (axis 14) と theme:verify (axis 8c) が同じロジックを要求するため、
 *      drift 防止に共通化した。
 */

const fs = require('node:fs');
const path = require('node:path');

const ALIAS_NAMES = ['brand-background', 'brand-surface', 'brand-text', 'brand-border'];
const ALIAS_DEF_RE = new RegExp(`--(?:${ALIAS_NAMES.join('|')}):\\s*var`);
const ALIAS_USE_RE = new RegExp(`var\\(--(?:${ALIAS_NAMES.join('|')})[^a-z-]`);

/**
 * @param {object} opts
 * @param {string} opts.srcDir          src ディレクトリの絶対パス
 * @param {string[]} [opts.htmlFiles]   走査対象の basename 一覧 (未指定なら readdir)
 * @param {(name: string) => string} [opts.reader]  ファイル content を返す関数 (cache 注入用)
 * @returns {{aliasFree: boolean, hasDef: boolean, usageFiles: string[]}}
 */
function checkBrandAliasResidue({ srcDir, htmlFiles, reader }) {
  const read = reader || ((name) => fs.readFileSync(path.join(srcDir, name), 'utf8'));
  const files = htmlFiles || fs.readdirSync(srcDir).filter((f) => /\.(html|js|css)$/.test(f));

  const hasDef = ALIAS_DEF_RE.test(read('UnifiedStyles.css.html'));
  const usageFiles = files.filter((f) => ALIAS_USE_RE.test(read(f)));

  return { aliasFree: !hasDef && usageFiles.length === 0, hasDef, usageFiles };
}

module.exports = { ALIAS_NAMES, checkBrandAliasResidue };
