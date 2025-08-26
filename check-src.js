#!/usr/bin/env node

const fs = require('fs');
const path = require('path');

console.log('🔍 /src ディレクトリ実地検証スクリプト\n');

const srcDir = path.join(__dirname, 'src');

// ファイルの存在確認と基本情報
function checkFileExists(filename) {
  const filepath = path.join(srcDir, filename);
  if (fs.existsSync(filepath)) {
    const stats = fs.statSync(filepath);
    const sizeKB = Math.round(stats.size / 1024);
    console.log(`✅ ${filename} (${sizeKB}KB)`);
    return fs.readFileSync(filepath, 'utf8');
  } else {
    console.log(`❌ ${filename} - ファイルが存在しません`);
    return null;
  }
}

// HTMLファイルの検証
function validateHtml(filename, content) {
  console.log(`\n🗂️ ${filename} HTML検証:`);
  
  let score = 0;
  const checks = [];
  
  // 基本構造
  if (content.includes('<!DOCTYPE html>')) { score += 25; checks.push('DOCTYPE'); }
  if (content.includes('<html')) { score += 25; checks.push('HTML'); }
  if (content.includes('<head>')) { score += 25; checks.push('HEAD'); }
  if (content.includes('<body>')) { score += 25; checks.push('BODY'); }
  
  // タイトル
  const titleMatch = content.match(/<title>(.*?)<\/title>/);
  if (titleMatch) { checks.push(`タイトル: ${titleMatch[1]}`); }
  
  // スクリプト
  const scriptCount = (content.match(/<script/g) || []).length;
  if (scriptCount > 0) { checks.push(`${scriptCount}個のスクリプト`); }
  
  // ファイル固有チェック
  if (filename === 'AdminPanel.html') {
    if (content.includes('save') || content.includes('publish')) { 
      score += 15; checks.push('保存機能'); 
    }
    if (content.includes('sheet') || content.includes('select')) { 
      score += 15; checks.push('シート選択'); 
    }
  }
  
  const status = score >= 80 ? '✅' : score >= 60 ? '⚠️' : '❌';
  console.log(`   ${status} スコア: ${score}% (${checks.join(', ')})`);
  
  return score;
}

// JavaScriptファイルの検証
function validateJs(filename, content) {
  console.log(`\n📜 ${filename} JavaScript検証:`);
  
  let score = 0;
  const features = [];
  
  // 関数数
  const functions = content.match(/function\s+([a-zA-Z_$][a-zA-Z0-9_$]*)/g) || [];
  const functionNames = functions.map(f => f.replace('function ', ''));
  score += Math.min(functions.length * 10, 50);
  features.push(`${functions.length}個の関数`);
  
  // エラーハンドリング
  if (content.includes('try') && content.includes('catch')) {
    score += 20;
    features.push('エラー処理');
  }
  
  // コンソール出力
  if (content.includes('console.log')) {
    score += 10;
    features.push('ログ出力');
  }
  
  // ファイル固有
  if (filename.includes('core')) {
    if (content.includes('adminLog')) { score += 15; features.push('管理ログ'); }
  } else if (filename.includes('api')) {
    if (content.includes('runGasWithUserId')) { score += 15; features.push('GAS API'); }
  }
  
  const status = score >= 70 ? '✅' : score >= 40 ? '⚠️' : '❌';
  console.log(`   ${status} スコア: ${score}% (${features.join(', ')})`);
  
  if (functionNames.length > 0) {
    console.log(`   📋 関数一覧: ${functionNames.slice(0, 5).join(', ')}${functionNames.length > 5 ? '...' : ''}`);
  }
  
  return score;
}

// GASファイルの検証
function validateGas(filename, content) {
  console.log(`\n⚙️ ${filename} GAS検証:`);
  
  let score = 0;
  const features = [];
  
  // 基本的なGAS API
  if (content.includes('SpreadsheetApp')) { score += 20; features.push('Spreadsheet API'); }
  if (content.includes('HtmlService')) { score += 20; features.push('HTML Service'); }
  if (content.includes('PropertiesService')) { score += 15; features.push('Properties Service'); }
  if (content.includes('CacheService')) { score += 10; features.push('Cache Service'); }
  
  // 関数数
  const functions = content.match(/^function\s+([a-zA-Z_$][a-zA-Z0-9_$]*)/gm) || [];
  const functionNames = functions.map(f => f.replace(/^function\s+/, ''));
  score += Math.min(functions.length * 5, 25);
  features.push(`${functions.length}個のエクスポート関数`);
  
  // ファイル固有の重要関数
  const criticalFunctions = {
    'main.gs': ['doGet', 'include', 'escapeJavaScript'],
    'Core.gs': ['verifyUserAccess', 'getPublishedSheetData'],
    'database.gs': ['createUser', 'updateUser', 'findUserById'],
    'config.gs': ['getConfig']
  };
  
  const required = criticalFunctions[filename] || [];
  let found = 0;
  const foundFunctions = [];
  
  required.forEach(func => {
    if (content.includes(`function ${func}`)) {
      found++;
      foundFunctions.push(func);
    }
  });
  
  if (required.length > 0) {
    score += (found / required.length) * 20;
    features.push(`${found}/${required.length} 重要関数`);
  }
  
  const status = score >= 75 ? '✅' : score >= 50 ? '⚠️' : '❌';
  console.log(`   ${status} スコア: ${score}% (${features.join(', ')})`);
  
  if (foundFunctions.length > 0) {
    console.log(`   ✅ 重要関数: ${foundFunctions.join(', ')}`);
  }
  if (found < required.length) {
    const missing = required.filter(f => !content.includes(`function ${f}`));
    console.log(`   ❌ 不足関数: ${missing.join(', ')}`);
  }
  
  return score;
}

// メイン実行
function main() {
  const results = [];
  
  // HTMLファイル検証
  console.log('=' .repeat(50));
  console.log('🗂️ HTMLファイル検証');
  console.log('=' .repeat(50));
  
  const htmlFiles = ['AdminPanel.html', 'AppSetupPage.html', 'LoginPage.html', 'Page.html'];
  htmlFiles.forEach(file => {
    const content = checkFileExists(file);
    if (content) {
      const score = validateHtml(file, content);
      results.push({ file, type: 'HTML', score });
    }
  });
  
  // JavaScriptファイル検証
  console.log('\n' + '=' .repeat(50));
  console.log('📜 JavaScriptファイル検証'); 
  console.log('=' .repeat(50));
  
  const jsFiles = [
    'adminPanel-core.js.html',
    'adminPanel-api.js.html',
    'adminPanel-events.js.html',
    'page.js.html',
    'UnifiedStyles.html'
  ];
  jsFiles.forEach(file => {
    const content = checkFileExists(file);
    if (content) {
      const score = validateJs(file, content);
      results.push({ file, type: 'JS', score });
    }
  });
  
  // GASファイル検証
  console.log('\n' + '=' .repeat(50));
  console.log('⚙️ GASファイル検証');
  console.log('=' .repeat(50));
  
  const gasFiles = ['main.gs', 'Core.gs', 'database.gs', 'config.gs'];
  gasFiles.forEach(file => {
    const content = checkFileExists(file);
    if (content) {
      const score = validateGas(file, content);
      results.push({ file, type: 'GAS', score });
    }
  });
  
  // サマリー
  console.log('\n' + '=' .repeat(50));
  console.log('📊 検証サマリー');
  console.log('=' .repeat(50));
  
  const totalFiles = results.length;
  const passedFiles = results.filter(r => r.score >= 70).length;
  const warningFiles = results.filter(r => r.score >= 40 && r.score < 70).length;
  const failedFiles = results.filter(r => r.score < 40).length;
  
  console.log(`📈 全体統計:`);
  console.log(`   ✅ 成功: ${passedFiles}/${totalFiles} (${Math.round(passedFiles/totalFiles*100)}%)`);
  console.log(`   ⚠️ 警告: ${warningFiles}/${totalFiles} (${Math.round(warningFiles/totalFiles*100)}%)`);
  console.log(`   ❌ 失敗: ${failedFiles}/${totalFiles} (${Math.round(failedFiles/totalFiles*100)}%)`);
  
  // 問題のあるファイル
  const problemFiles = results.filter(r => r.score < 70);
  if (problemFiles.length > 0) {
    console.log(`\n⚠️ 要注意ファイル:`);
    problemFiles.forEach(r => {
      console.log(`   ${r.score < 40 ? '❌' : '⚠️'} ${r.file}: ${r.score}%`);
    });
  }
  
  console.log(`\n🎯 総合評価: ${passedFiles >= totalFiles * 0.8 ? '✅ 良好' : warningFiles + passedFiles >= totalFiles * 0.6 ? '⚠️ 改善必要' : '❌ 要修正'}`);
}

// スクリプト実行
if (require.main === module) {
  main();
}