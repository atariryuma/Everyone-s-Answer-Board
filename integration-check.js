#!/usr/bin/env node

/**
 * Page.htmlとpage.js.htmlの統合後の包括的チェックスクリプト
 * CLAUDE.md準拠、重複チェック、未定義、ずれのチェックを実行
 */

const fs = require('fs');
const path = require('path');

// 統合チェック結果
const results = {
  duplicateCheck: { status: 'OK', issues: [] },
  claudeCompliance: { status: 'OK', issues: [] },
  undefinedCheck: { status: 'OK', issues: [] },
  consistencyCheck: { status: 'OK', issues: [] },
  summary: { errors: 0, warnings: 0 }
};

console.log('🔍 JavaScript統合後の包括的チェックを開始...\n');

// ファイル読み込み
let pageHtml, pageJs;

try {
  pageHtml = fs.readFileSync('src/Page.html', 'utf8');
  pageJs = fs.readFileSync('src/page.js.html', 'utf8');
} catch (error) {
  console.error('❌ ファイル読み込みエラー:', error.message);
  process.exit(1);
}

// 1. 重複チェック
console.log('📋 1. 重複チェック');

// Page.htmlにスクリプトタグがないことを確認
const pageHtmlScriptMatches = pageHtml.match(/<script[\s\S]*?<\/script>/gi);
if (pageHtmlScriptMatches && pageHtmlScriptMatches.length > 0) {
  results.duplicateCheck.issues.push('Page.htmlに不要なscriptタグが残存している');
  results.duplicateCheck.status = 'ERROR';
  results.summary.errors++;
}

// page.js.htmlに正しいincludeがあることを確認
const includeMatch = pageHtml.includes("include('page.js')");
if (!includeMatch) {
  results.duplicateCheck.issues.push('Page.htmlにpage.jsのinclude文が見つからない');
  results.duplicateCheck.status = 'ERROR';
  results.summary.errors++;
}

// テンプレート変数がpage.js.htmlに移行されたか確認
if (!pageJs.includes('parseTemplateVars')) {
  results.duplicateCheck.issues.push('page.js.htmlにテンプレート変数解析システムが見つからない');
  results.duplicateCheck.status = 'ERROR';
  results.summary.errors++;
}

if (!pageJs.includes('TEMPLATE_VARS')) {
  results.duplicateCheck.issues.push('page.js.htmlにTEMPLATE_VARS定義が見つからない');
  results.duplicateCheck.status = 'ERROR';
  results.summary.errors++;
}

console.log(`   ${results.duplicateCheck.status === 'OK' ? '✅' : '❌'} 重複チェック: ${results.duplicateCheck.status}`);
if (results.duplicateCheck.issues.length > 0) {
  results.duplicateCheck.issues.forEach(issue => console.log(`     - ${issue}`));
}

// 2. CLAUDE.md準拠チェック
console.log('\n📋 2. CLAUDE.md準拠チェック');

// const優先チェック
const varCount = (pageJs.match(/\bvar\s+/g) || []).length;
if (varCount > 0) {
  results.claudeCompliance.issues.push(`var使用が${varCount}箇所で発見 (CLAUDE.md規範違反)`);
  results.claudeCompliance.status = 'ERROR';
  results.summary.errors++;
}

// Object.freeze使用チェック
const freezeCount = (pageJs.match(/Object\.freeze\(/g) || []).length;
if (freezeCount < 5) {
  results.claudeCompliance.issues.push(`Object.freeze使用が少ない: ${freezeCount}箇所 (推奨: 5箇所以上)`);
  results.claudeCompliance.status = 'WARNING';
  results.summary.warnings++;
}

// 統一データソースチェック
if (pageJs.includes('publishedSpreadsheetId')) {
  results.claudeCompliance.issues.push('重複フィールド publishedSpreadsheetId が見つかった (統一データソース原則違反)');
  results.claudeCompliance.status = 'ERROR';
  results.summary.errors++;
}

// 必須定数の存在チェック
const requiredConstants = ['USER_ID', 'SHEET_NAME', 'ICONS', 'RENDER_BATCH_SIZE', 'CHUNK_SIZE'];
requiredConstants.forEach(constant => {
  if (!pageJs.includes(`const ${constant}`)) {
    results.claudeCompliance.issues.push(`必須定数 ${constant} が見つからない`);
    results.claudeCompliance.status = 'ERROR';
    results.summary.errors++;
  }
});

console.log(`   ${results.claudeCompliance.status === 'OK' ? '✅' : results.claudeCompliance.status === 'WARNING' ? '⚠️' : '❌'} CLAUDE.md準拠: ${results.claudeCompliance.status}`);
if (results.claudeCompliance.issues.length > 0) {
  results.claudeCompliance.issues.forEach(issue => console.log(`     - ${issue}`));
}

// 3. 未定義参照チェック
console.log('\n📋 3. 未定義参照チェック');

// StudyQuestAppクラスの存在チェック
if (!pageJs.includes('class StudyQuestApp')) {
  results.undefinedCheck.issues.push('StudyQuestAppクラス定義が見つからない');
  results.undefinedCheck.status = 'ERROR';
  results.summary.errors++;
}

// window.studyQuestAppの初期化チェック
if (!pageJs.includes('window.studyQuestApp = new StudyQuestApp')) {
  results.undefinedCheck.issues.push('StudyQuestAppインスタンス初期化が見つからない');
  results.undefinedCheck.status = 'ERROR';
  results.summary.errors++;
}

// テンプレート変数のグローバルエクスポートチェック
if (!pageJs.includes('Object.assign(window, CONFIG)')) {
  results.undefinedCheck.issues.push('CONFIG変数のグローバルエクスポートが見つからない');
  results.undefinedCheck.status = 'ERROR';
  results.summary.errors++;
}

console.log(`   ${results.undefinedCheck.status === 'OK' ? '✅' : '❌'} 未定義参照: ${results.undefinedCheck.status}`);
if (results.undefinedCheck.issues.length > 0) {
  results.undefinedCheck.issues.forEach(issue => console.log(`     - ${issue}`));
}

// 4. 整合性チェック
console.log('\n📋 4. 整合性チェック');

// ファイルサイズの合理性チェック
const pageHtmlLines = pageHtml.split('\n').length;
const pageJsLines = pageJs.split('\n').length;

if (pageHtmlLines > 1000) {
  results.consistencyCheck.issues.push(`Page.htmlが大きすぎる: ${pageHtmlLines}行 (目標: 1000行以下)`);
  results.consistencyCheck.status = 'WARNING';
  results.summary.warnings++;
}

if (pageJsLines < 5000) {
  results.consistencyCheck.issues.push(`page.js.htmlが小さすぎる: ${pageJsLines}行 (期待: 5000行以上)`);
  results.consistencyCheck.status = 'WARNING';
  results.summary.warnings++;
}

// include文の正確性チェック
const hasPageJsInclude = pageHtml.includes("include('page.js')");

if (!hasPageJsInclude) {
  results.consistencyCheck.issues.push('page.js のinclude文が見つからない');
  results.consistencyCheck.status = 'ERROR';
  results.summary.errors++;
}

console.log(`   ${results.consistencyCheck.status === 'OK' ? '✅' : results.consistencyCheck.status === 'WARNING' ? '⚠️' : '❌'} 整合性: ${results.consistencyCheck.status}`);
if (results.consistencyCheck.issues.length > 0) {
  results.consistencyCheck.issues.forEach(issue => console.log(`     - ${issue}`));
}

// 最終結果サマリー
console.log('\n📊 チェック結果サマリー');
console.log('================================');

if (results.summary.errors === 0 && results.summary.warnings === 0) {
  console.log('🎉 全てのチェックをパス！');
  console.log('✨ JavaScript統合が正常に完了しました');
} else {
  console.log(`⚠️ エラー: ${results.summary.errors}個`);
  console.log(`⚠️ 警告: ${results.summary.warnings}個`);
  
  if (results.summary.errors > 0) {
    console.log('\n❌ エラーが検出されました。修正が必要です。');
  } else {
    console.log('\n✅ エラーはありませんが、警告を確認してください。');
  }
}

console.log('\n🔧 統合後の構造:');
console.log('├── Page.html: 純粋なHTMLテンプレート (956行)');
console.log('├── <?!= include(\'page.js\'); ?>: 1つのJSインクルード');
console.log(`└── page.js.html: 統合JavaScript (${pageJsLines}行)`);

// エラーがあれば終了コード1で終了
process.exit(results.summary.errors > 0 ? 1 : 0);