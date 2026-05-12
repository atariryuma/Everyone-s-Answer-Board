#!/usr/bin/env node
/**
 * CLAUDE.md ルールの違反を機械的に検出する軽量カスタム静的解析。
 *
 * Why: ESLint だけでは検出しづらい GAS 特有のアンチパターン（V8 ランタイムでも
 *   動作しない API、権限昇格、性能アンチパターン、XSS リスク）を、
 *   行番号付きで指摘する。違反は重大度別に集計し、CI に組み込めるよう
 *   非ゼロ終了コードを返す。
 *
 * 使い方:
 *   node scripts/lint.js                 # 全 src/ をスキャン
 *   node scripts/lint.js src/main.js     # 特定ファイル
 *   node scripts/lint.js --severity error  # ERROR のみ表示
 *
 * 終了コード:
 *   0  違反なし、もしくは WARN のみ
 *   1  1 件以上の ERROR
 *   2  使い方エラー
 */

const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const SRC_DIR = path.join(ROOT, 'src');

/**
 * ルール定義。
 *   - id           : 識別子（CI ログで grep しやすく）
 *   - severity     : 'error' | 'warn'
 *   - patterns     : RegExp の配列。複数行マッチ可
 *   - message      : 違反時のメッセージ
 *   - filter       : (filepath, content) => boolean。false なら適用しない
 *   - allowedFiles : 例外的に許容するファイル名サフィックス（テスト等）
 */
const RULES = [
  {
    id: 'no-set-timeout',
    severity: 'error',
    message: 'setTimeout/setInterval は GAS V8 では動作しない。Utilities.sleep() か Trigger を使用。',
    patterns: [/\bsetTimeout\s*\(/g, /\bsetInterval\s*\(/g],
    // フロント HTML 内の JS はブラウザで動くので許可。.js（GAS バックエンド）のみ検査。
    filter: (filepath) => filepath.endsWith('.js') && !filepath.includes('/tests/'),
  },
  {
    id: 'no-effective-user',
    severity: 'error',
    message: 'Session.getEffectiveUser() は権限昇格リスク。Session.getActiveUser() のみ使用。',
    patterns: [/Session\s*\.\s*getEffectiveUser\s*\(/g],
    filter: (filepath) => !filepath.includes('/tests/'),
  },
  {
    id: 'no-inner-html-assignment-with-variable',
    severity: 'warn',
    message: 'innerHTML = <変数> は XSS リスク。escapeHtml() / textContent を使用するか、テンプレートリテラル内の値を必ずエスケープ。',
    // innerHTML = `${...}` または innerHTML = someVar の形を検出
    // ただし innerHTML = '...固定文字列...' は許可
    patterns: [
      /\binnerHTML\s*=\s*[a-zA-Z_$][\w$.[\]]*\s*;/g,
    ],
    filter: (filepath) => /\.html$/.test(filepath) && !filepath.includes('/tests/'),
  },
  {
    id: 'no-get-value-loop',
    severity: 'warn',
    message: 'getValue() をループ内で使うと N+1 で遅い。getRange().getValues() でバッチ取得。',
    // for ... { ... .getValue() ... } を雑に検出（行内 for + getValue 同一ファイル内連続）
    // 行ごとには検出困難なので、行内シンプル形を検出（要 review）
    patterns: [/for\s*\([^)]+\)\s*\{[^}]{0,400}\.getValue\s*\(\s*\)/g],
    filter: (filepath) => filepath.endsWith('.js') && !filepath.includes('/tests/'),
  },
  {
    id: 'no-direct-property-fetch',
    severity: 'warn',
    message: 'PropertiesService を直接呼ばず getCachedProperty() / setCachedProperty() を使用（30s メモリキャッシュ）。',
    patterns: [/PropertiesService\s*\.\s*getScriptProperties\s*\(\s*\)\s*\.\s*getProperty\s*\(/g],
    filter: (filepath) => filepath.endsWith('.js') && !filepath.includes('/tests/')
      // helpers.js 自身の実装は許可
      && !filepath.endsWith('/helpers.js'),
  },
  {
    id: 'top-level-side-effects',
    severity: 'warn',
    message: 'HTML テンプレート JS の top-level に google.script.run / DOM 操作 / setInterval があると ReferenceError や race condition の原因。init() 内へ。',
    patterns: [
      // <script> 直後（DOMContentLoaded 待ちなし）に google.script.run.X(...) を雑に検出
      /<script[^>]*>\s*\n[\s\S]{0,200}?google\.script\.run\.[a-zA-Z_$][\w$]*\s*\(/g,
    ],
    filter: (filepath) => /\.html$/.test(filepath) && !filepath.includes('/tests/'),
  },
];

function listSourceFiles(target) {
  const out = [];
  function walk(dir) {
    for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
      const full = path.join(dir, e.name);
      if (e.isDirectory()) walk(full);
      else if (e.isFile() && /\.(js|html)$/.test(e.name)) out.push(full);
    }
  }
  if (fs.statSync(target).isDirectory()) walk(target);
  else out.push(path.resolve(target));
  return out;
}

function findLineNumber(content, index) {
  // index 位置の前にある \n の数 + 1
  let n = 1;
  for (let i = 0; i < index; i++) if (content.charCodeAt(i) === 10) n++;
  return n;
}

function snippet(content, index, span = 80) {
  // 行頭から行末まで（だが span に切り詰め）
  const lineStart = content.lastIndexOf('\n', index) + 1;
  const lineEnd = content.indexOf('\n', index);
  const line = content.substring(lineStart, lineEnd === -1 ? content.length : lineEnd);
  return line.length > span ? line.substring(0, span) + '...' : line;
}

/**
 * Suppression コメントの判定。
 *
 *   <code> // lint-disable-line <rule-id>            ← 同じ行を suppress
 *   <code> // lint-disable-line <rule-id> -- 理由    ← 同じ行を suppress (理由付き)
 *   // lint-disable-next-line <rule-id>              ← 次の行を suppress
 *   // lint-disable-next-line <rule-id> -- 理由      ← 次の行を suppress (理由付き)
 *
 * Why: ESLint 互換の suppress 表記。1 ファイル全体を無効化する手段は意図的に提供しない
 *     (見落としやすいため)。理由 (`-- ...`) は強制ではないが、書くと grep 可能になる。
 */
function isSuppressed(content, finding) {
  const lines = content.split('\n');
  // 1-indexed line に合わせる
  const sameLine = lines[finding.line - 1] || '';
  const prevLine = lines[finding.line - 2] || '';

  const sameLineRe = new RegExp(`//\\s*lint-disable-line\\s+${finding.rule}(\\s|$|;)`);
  const prevLineRe = new RegExp(`//\\s*lint-disable-next-line\\s+${finding.rule}(\\s|$|;)`);

  return sameLineRe.test(sameLine) || prevLineRe.test(prevLine);
}

function runLint(targets) {
  const findings = [];
  for (const file of targets) {
    const rel = path.relative(ROOT, file);
    let content;
    try { content = fs.readFileSync(file, 'utf8'); }
    catch (_) { continue; }

    for (const rule of RULES) {
      if (rule.filter && !rule.filter(file, content)) continue;

      for (const pat of rule.patterns) {
        pat.lastIndex = 0;
        let m;
        while ((m = pat.exec(content)) !== null) {
          const candidate = {
            file: rel,
            line: findLineNumber(content, m.index),
            rule: rule.id,
            severity: rule.severity,
            message: rule.message,
            snippet: snippet(content, m.index).trim()
          };
          if (!isSuppressed(content, candidate)) {
            findings.push(candidate);
          }
          // 無限ループ防止（zero-width 対策）
          if (pat.lastIndex === m.index) pat.lastIndex++;
        }
      }
    }
  }
  return findings;
}

function main() {
  const args = process.argv.slice(2);
  let severityFilter = null;
  const fileArgs = [];
  for (let i = 0; i < args.length; i++) {
    if (args[i] === '--severity') { severityFilter = args[++i]; continue; }
    if (args[i] === '--help' || args[i] === '-h') {
      console.log('Usage: node scripts/lint.js [--severity error|warn] [files...]');
      process.exit(0);
    }
    fileArgs.push(args[i]);
  }

  const targets = fileArgs.length === 0
    ? listSourceFiles(SRC_DIR)
    : fileArgs.flatMap(f => listSourceFiles(f));

  const findings = runLint(targets);
  const filtered = severityFilter
    ? findings.filter(f => f.severity === severityFilter)
    : findings;

  if (filtered.length === 0) {
    console.log(`✓ No CLAUDE.md rule violations${severityFilter ? ' (severity=' + severityFilter + ')' : ''} across ${targets.length} files.`);
    process.exit(0);
  }

  // ファイル別にグループ化して見やすく出力
  const byFile = new Map();
  for (const f of filtered) {
    if (!byFile.has(f.file)) byFile.set(f.file, []);
    byFile.get(f.file).push(f);
  }

  let errorCount = 0;
  let warnCount = 0;
  for (const [file, items] of byFile) {
    console.log(`\n${file}`);
    for (const it of items.sort((a, b) => a.line - b.line)) {
      const tag = it.severity === 'error' ? '✗ ERROR' : '⚠ WARN ';
      console.log(`  ${file}:${it.line}  ${tag}  [${it.rule}]`);
      console.log(`    ${it.message}`);
      console.log(`    > ${it.snippet}`);
      if (it.severity === 'error') errorCount++; else warnCount++;
    }
  }

  console.log(`\nTotal: ${errorCount} error, ${warnCount} warn  (across ${byFile.size} files)`);
  process.exit(errorCount > 0 ? 1 : 0);
}

if (require.main === module) main();

module.exports = { runLint, listSourceFiles, RULES };
