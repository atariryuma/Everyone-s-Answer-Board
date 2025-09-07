#!/usr/bin/env node

/**
 * GAS専用構文チェッカー
 * Google Apps Scriptファイル(.gs)の構文エラーを検出
 */

const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const SRC_DIR = path.join(__dirname, '..', 'src');
const TEMP_DIR = '/tmp';

// 色付きコンソール出力
const colors = {
  reset: '\x1b[0m',
  red: '\x1b[31m',
  green: '\x1b[32m',
  yellow: '\x1b[33m',
  blue: '\x1b[34m',
  magenta: '\x1b[35m'
};

function colorLog(message, color = 'reset') {
  console.log(colors[color] + message + colors.reset);
}

/**
 * .gsファイルを.jsファイルにコピーしてNode.js構文チェック
 */
function checkGasFileSyntax(filePath) {
  const fileName = path.basename(filePath, '.gs');
  const tempFile = path.join(TEMP_DIR, `${fileName}.js`);
  
  try {
    // .gsファイルを.jsファイルとしてコピー
    fs.copyFileSync(filePath, tempFile);
    
    // Node.jsで構文チェック
    execSync(`node -c "${tempFile}"`, { stdio: 'pipe' });
    
    // 一時ファイル削除
    fs.unlinkSync(tempFile);
    
    return { success: true, file: path.basename(filePath) };
    
  } catch (error) {
    // 一時ファイル削除（エラー時も実行）
    try {
      fs.unlinkSync(tempFile);
    } catch (cleanupError) {
      // ignore cleanup errors
    }
    
    return {
      success: false,
      file: path.basename(filePath),
      error: error.message
    };
  }
}

/**
 * srcディレクトリ内の全.gsファイルを検査
 */
function checkAllGasFiles() {
  colorLog('🔍 GAS構文チェック開始...', 'blue');
  
  if (!fs.existsSync(SRC_DIR)) {
    colorLog('❌ srcディレクトリが見つかりません', 'red');
    process.exit(1);
  }
  
  const gasFiles = fs.readdirSync(SRC_DIR)
    .filter(file => file.endsWith('.gs'))
    .map(file => path.join(SRC_DIR, file));
  
  if (gasFiles.length === 0) {
    colorLog('⚠️  .gsファイルが見つかりません', 'yellow');
    process.exit(0);
  }
  
  colorLog(`📄 検査対象: ${gasFiles.length}ファイル`, 'blue');
  
  const results = gasFiles.map(checkGasFileSyntax);
  const successes = results.filter(r => r.success);
  const failures = results.filter(r => !r.success);
  
  // 結果表示
  console.log('');
  successes.forEach(result => {
    colorLog(`✅ ${result.file}`, 'green');
  });
  
  failures.forEach(result => {
    colorLog(`❌ ${result.file}`, 'red');
    colorLog(`   エラー: ${result.error}`, 'red');
  });
  
  console.log('');
  colorLog(`📊 結果: ${successes.length}成功 / ${failures.length}失敗`, 
    failures.length === 0 ? 'green' : 'yellow');
  
  if (failures.length > 0) {
    colorLog('💡 推奨: 構文エラーを修正してから再実行してください', 'yellow');
    process.exit(1);
  } else {
    colorLog('🎉 全ての.gsファイルの構文チェックが完了しました！', 'green');
    process.exit(0);
  }
}

/**
 * clasp statusも表示（オプション）
 */
function showClaspStatus() {
  try {
    colorLog('\n📋 Clasp Status:', 'magenta');
    const status = execSync('clasp status', { encoding: 'utf8' });
    console.log(status);
  } catch (error) {
    colorLog('⚠️  clasp statusの実行に失敗しました', 'yellow');
  }
}

// メイン実行
if (require.main === module) {
  checkAllGasFiles();
  
  // --statusオプションが指定された場合
  if (process.argv.includes('--status')) {
    showClaspStatus();
  }
}

module.exports = {
  checkGasFileSyntax,
  checkAllGasFiles
};