#!/usr/bin/env node

/**
 * @fileoverview 重複関数検出プログラム
 * システム内の.gsファイルを解析し、重複している関数を検出します
 */

const fs = require('fs');
const path = require('path');

// 設定
const CONFIG = {
  sourceDir: path.join(__dirname, '..', 'src'),
  extensions: ['.gs'],
  ignorePatterns: [
    /^\/\*\*[\s\S]*?\*\//,  // JSDocコメント
    /^\/\*[\s\S]*?\*\//,    // マルチラインコメント
    /\/\/.*$/,              // 単行コメント
  ],
  minFunctionBodyLines: 3,   // 最小関数本体行数（短すぎる関数は除外）
};

/**
 * 関数情報の構造体
 * @typedef {Object} FunctionInfo
 * @property {string} name - 関数名
 * @property {string} signature - 関数シグネチャ
 * @property {string} body - 関数本体
 * @property {string} file - ファイルパス
 * @property {number} lineNumber - 開始行番号
 * @property {number} endLineNumber - 終了行番号
 * @property {string} hash - 内容のハッシュ値
 */

/**
 * ファイル内容からコメントを除去
 * @param {string} content - ファイル内容
 * @returns {string} コメントを除去した内容
 */
function removeComments(content) {
  let result = content;
  
  // マルチラインコメントを除去
  result = result.replace(/\/\*[\s\S]*?\*\//g, '');
  
  // 単行コメントを除去（文字列内は除外）
  result = result.replace(/\/\/(?=(?:[^"']*["'][^"']*["'])*[^"']*$).*$/gm, '');
  
  return result;
}

/**
 * 関数の内容をノーマライズ（空白・改行を統一）
 * @param {string} content - 関数内容
 * @returns {string} ノーマライズされた内容
 */
function normalizeFunctionContent(content) {
  return content
    .replace(/\s+/g, ' ')           // 連続する空白を1つに
    .replace(/\s*{\s*/g, '{')       // 開始括弧前後の空白除去
    .replace(/\s*}\s*/g, '}')       // 終了括弧前後の空白除去
    .replace(/\s*;\s*/g, ';')       // セミコロン前後の空白除去
    .replace(/\s*,\s*/g, ',')       // カンマ前後の空白除去
    .trim();
}

/**
 * 文字列の簡易ハッシュを生成
 * @param {string} str - ハッシュ化する文字列
 * @returns {string} ハッシュ値
 */
function simpleHash(str) {
  let hash = 0;
  if (str.length === 0) return hash.toString();
  
  for (let i = 0; i < str.length; i++) {
    const char = str.charCodeAt(i);
    hash = ((hash << 5) - hash) + char;
    hash = hash & hash; // 32bit整数に変換
  }
  
  return hash.toString();
}

/**
 * ファイルから関数を抽出
 * @param {string} filePath - ファイルパス
 * @returns {FunctionInfo[]} 関数情報の配列
 */
function extractFunctions(filePath) {
  const content = fs.readFileSync(filePath, 'utf-8');
  const cleanContent = removeComments(content);
  const lines = content.split('\n');
  const functions = [];

  // 関数定義のパターン（GAS/JavaScript）
  const functionPatterns = [
    // function functionName() {}
    /^(\s*)function\s+([a-zA-Z_$][a-zA-Z0-9_$]*)\s*\((.*?)\)\s*{/,
    // const functionName = function() {}
    /^(\s*)const\s+([a-zA-Z_$][a-zA-Z0-9_$]*)\s*=\s*function\s*\((.*?)\)\s*{/,
    // const functionName = () => {}
    /^(\s*)const\s+([a-zA-Z_$][a-zA-Z0-9_$]*)\s*=\s*\((.*?)\)\s*=>\s*{/,
    // const functionName = (params) => {}
    /^(\s*)const\s+([a-zA-Z_$][a-zA-Z0-9_$]*)\s*=\s*([a-zA-Z_$][a-zA-Z0-9_$,\s]*)\s*=>\s*{/,
    // メソッド定義 methodName() {}
    /^(\s*)([a-zA-Z_$][a-zA-Z0-9_$]*)\s*\((.*?)\)\s*{/,
  ];

  let i = 0;
  while (i < lines.length) {
    const line = lines[i];
    let match = null;
    let patternIndex = -1;

    // 各パターンをチェック
    for (let p = 0; p < functionPatterns.length; p++) {
      match = line.match(functionPatterns[p]);
      if (match) {
        patternIndex = p;
        break;
      }
    }

    if (match) {
      const indent = match[1] || '';
      const functionName = match[2];
      const params = match[3] || '';
      
      // 関数名のフィルタリング（明らかに関数でないものを除外）
      if (functionName.match(/^(if|for|while|switch|try|catch|else)$/)) {
        i++;
        continue;
      }

      const startLine = i;
      let braceCount = 1;
      let endLine = i;
      let functionBody = '';

      // 関数の終わりを探す（括弧の対応を数える）
      i++; // 次の行から開始
      while (i < lines.length && braceCount > 0) {
        const currentLine = lines[i];
        functionBody += currentLine + '\n';
        
        // 文字列内の括弧は除外して括弧をカウント
        let inString = false;
        let stringChar = '';
        
        for (let j = 0; j < currentLine.length; j++) {
          const char = currentLine[j];
          
          if (!inString && (char === '"' || char === "'" || char === '`')) {
            inString = true;
            stringChar = char;
          } else if (inString && char === stringChar && currentLine[j-1] !== '\\') {
            inString = false;
          } else if (!inString) {
            if (char === '{') braceCount++;
            if (char === '}') braceCount--;
          }
        }
        
        if (braceCount === 0) {
          endLine = i;
          break;
        }
        i++;
      }

      // 関数本体が十分な長さの場合のみ追加
      const bodyLines = functionBody.trim().split('\n').filter(l => l.trim().length > 0);
      if (bodyLines.length >= CONFIG.minFunctionBodyLines) {
        const signature = `function ${functionName}(${params})`;
        const normalizedBody = normalizeFunctionContent(functionBody);
        const hash = simpleHash(normalizedBody);

        functions.push({
          name: functionName,
          signature,
          body: functionBody.trim(),
          file: path.relative(CONFIG.sourceDir, filePath),
          lineNumber: startLine + 1,
          endLineNumber: endLine + 1,
          hash,
          normalizedBody
        });
      }
    } else {
      i++;
    }
  }

  return functions;
}

/**
 * 全ての.gsファイルを取得
 * @param {string} dir - 検索ディレクトリ
 * @returns {string[]} ファイルパスの配列
 */
function getAllGasFiles(dir) {
  const files = [];
  
  function scanDirectory(currentDir) {
    const entries = fs.readdirSync(currentDir);
    
    for (const entry of entries) {
      const fullPath = path.join(currentDir, entry);
      const stat = fs.statSync(fullPath);
      
      if (stat.isDirectory()) {
        scanDirectory(fullPath);
      } else if (CONFIG.extensions.includes(path.extname(entry))) {
        files.push(fullPath);
      }
    }
  }
  
  scanDirectory(dir);
  return files;
}

/**
 * 重複を検出
 * @param {FunctionInfo[]} functions - 全関数の配列
 * @returns {Object} 重複検出結果
 */
function detectDuplicates(functions) {
  // デバッグ: 関数の構造を確認
  console.log('デバッグ: 最初の関数の構造:', functions[0]);
  
  const results = {
    exactDuplicates: [],      // 完全一致
    similarFunctions: [],     // 類似関数
    nameConflicts: [],        // 同名関数
    stats: {
      totalFunctions: functions.length,
      uniqueNames: new Set(functions.map(f => f.name).filter(name => typeof name === 'string')).size,
      duplicateGroups: 0
    }
  };

  // 1. 完全一致の重複を検出（ハッシュベース）
  const hashGroups = {};
  functions.forEach(func => {
    if (!hashGroups[func.hash]) {
      hashGroups[func.hash] = [];
    }
    hashGroups[func.hash].push(func);
  });

  Object.values(hashGroups).forEach(group => {
    if (group.length > 1) {
      results.exactDuplicates.push({
        hash: group[0].hash,
        functions: group,
        type: 'exact_duplicate'
      });
      results.stats.duplicateGroups++;
    }
  });

  // 2. 同名関数を検出
  const nameGroups = new Map();
  functions.forEach(func => {
    const funcName = func.name;
    if (!nameGroups.has(funcName)) {
      nameGroups.set(funcName, []);
    }
    nameGroups.get(funcName).push(func);
  });

  Array.from(nameGroups.values()).forEach(group => {
    if (group.length > 1) {
      // ハッシュが異なる場合は同名関数として記録
      const uniqueHashes = new Set(group.map(f => f.hash));
      if (uniqueHashes.size > 1) {
        results.nameConflicts.push({
          name: group[0].name,
          functions: group,
          type: 'name_conflict'
        });
      }
    }
  });

  // 3. 類似関数を検出（関数名が異なるが内容が似ている）
  const processedHashes = new Set();
  functions.forEach((func1, i) => {
    if (processedHashes.has(func1.hash)) return;
    
    const similarGroup = [func1];
    
    functions.forEach((func2, j) => {
      if (i !== j && func1.name !== func2.name && func1.hash === func2.hash) {
        similarGroup.push(func2);
      }
    });

    if (similarGroup.length > 1) {
      results.similarFunctions.push({
        hash: func1.hash,
        functions: similarGroup,
        type: 'similar_content'
      });
      similarGroup.forEach(f => processedHashes.add(f.hash));
    }
  });

  return results;
}

/**
 * レポートを生成
 * @param {Object} results - 検出結果
 */
function generateReport(results) {
  console.log('\n🔍 ==============================================');
  console.log('     重複関数検出レポート');
  console.log('==============================================\n');

  console.log(`📊 統計情報:`);
  console.log(`   - 総関数数: ${results.stats.totalFunctions}`);
  console.log(`   - ユニーク関数名: ${results.stats.uniqueNames}`);
  console.log(`   - 重複グループ: ${results.stats.duplicateGroups}`);
  console.log('');

  // 完全一致の重複
  if (results.exactDuplicates.length > 0) {
    console.log('🚨 完全一致の重複関数:');
    results.exactDuplicates.forEach((group, index) => {
      console.log(`\n   ${index + 1}. 関数名: ${group.functions[0].name}`);
      console.log(`      ハッシュ: ${group.hash}`);
      group.functions.forEach(func => {
        console.log(`      📄 ${func.file}:${func.lineNumber}-${func.endLineNumber}`);
      });
      
      // 関数本体のプレビュー（最初の3行）
      const preview = group.functions[0].body.split('\n').slice(0, 3).join('\n');
      console.log(`      📋 プレビュー:\n${preview.replace(/^/gm, '         ')}...`);
    });
  } else {
    console.log('✅ 完全一致の重複関数: なし');
  }

  // 同名関数（内容が異なる）
  if (results.nameConflicts.length > 0) {
    console.log('\n⚠️  同名だが内容が異なる関数:');
    results.nameConflicts.forEach((group, index) => {
      console.log(`\n   ${index + 1}. 関数名: ${group.name}`);
      group.functions.forEach((func, i) => {
        console.log(`      📄 版${i + 1}: ${func.file}:${func.lineNumber}-${func.endLineNumber} (hash: ${func.hash.slice(0, 8)}...)`);
      });
    });
  } else {
    console.log('✅ 同名関数の競合: なし');
  }

  // 類似関数
  if (results.similarFunctions.length > 0) {
    console.log('\n🔄 類似関数（異なる名前だが同じ内容）:');
    results.similarFunctions.forEach((group, index) => {
      console.log(`\n   ${index + 1}. 共通ハッシュ: ${group.hash}`);
      group.functions.forEach(func => {
        console.log(`      📄 ${func.name}: ${func.file}:${func.lineNumber}-${func.endLineNumber}`);
      });
    });
  } else {
    console.log('✅ 類似関数: なし');
  }

  // 推奨アクション
  console.log('\n💡 推奨アクション:');
  
  if (results.exactDuplicates.length > 0) {
    console.log('   🗑️  完全一致の重複関数を削除または統合してください');
    results.exactDuplicates.forEach((group, i) => {
      console.log(`      - ${group.functions[0].name}: ${group.functions.length - 1}個の重複を削除`);
    });
  }
  
  if (results.nameConflicts.length > 0) {
    console.log('   🏷️  同名関数の名前を変更または統合してください');
    results.nameConflicts.forEach(group => {
      console.log(`      - ${group.name}: ${group.functions.length}個の版が存在`);
    });
  }
  
  if (results.similarFunctions.length > 0) {
    console.log('   🔧 類似関数の統合を検討してください');
    results.similarFunctions.forEach(group => {
      const names = group.functions.map(f => f.name).join(', ');
      console.log(`      - 統合候補: ${names}`);
    });
  }

  if (results.exactDuplicates.length === 0 && 
      results.nameConflicts.length === 0 && 
      results.similarFunctions.length === 0) {
    console.log('   🎉 重複関数は検出されませんでした！');
  }
}

/**
 * 詳細レポートをファイルに保存
 * @param {Object} results - 検出結果
 */
function saveDetailedReport(results) {
  const reportPath = path.join(__dirname, '..', 'scripts', 'reports', `duplicate-functions-${Date.now()}.json`);
  
  // reportsディレクトリを作成
  const reportsDir = path.dirname(reportPath);
  if (!fs.existsSync(reportsDir)) {
    fs.mkdirSync(reportsDir, { recursive: true });
  }

  const detailedReport = {
    timestamp: new Date().toISOString(),
    config: CONFIG,
    results: results,
    summary: {
      totalFunctions: results.stats.totalFunctions,
      exactDuplicates: results.exactDuplicates.length,
      nameConflicts: results.nameConflicts.length,
      similarFunctions: results.similarFunctions.length,
    }
  };

  fs.writeFileSync(reportPath, JSON.stringify(detailedReport, null, 2));
  console.log(`\n📄 詳細レポートを保存しました: ${path.relative(process.cwd(), reportPath)}`);
  
  return reportPath;
}

/**
 * メイン実行関数
 */
function main() {
  try {
    console.log('🔍 重複関数検出を開始しています...\n');
    
    // ソースディレクトリの確認
    if (!fs.existsSync(CONFIG.sourceDir)) {
      throw new Error(`ソースディレクトリが見つかりません: ${CONFIG.sourceDir}`);
    }

    // .gsファイルを取得
    const gasFiles = getAllGasFiles(CONFIG.sourceDir);
    console.log(`📁 検出された.gsファイル: ${gasFiles.length}個`);
    gasFiles.forEach(file => {
      console.log(`   - ${path.relative(CONFIG.sourceDir, file)}`);
    });

    // 各ファイルから関数を抽出
    console.log('\n🔍 関数を抽出しています...');
    const allFunctions = [];
    
    gasFiles.forEach(file => {
      try {
        const functions = extractFunctions(file);
        allFunctions.push(...functions);
        console.log(`   📄 ${path.relative(CONFIG.sourceDir, file)}: ${functions.length}個の関数`);
      } catch (error) {
        console.warn(`   ⚠️  ${path.relative(CONFIG.sourceDir, file)}: 解析エラー - ${error.message}`);
      }
    });

    console.log(`\n📊 総関数数: ${allFunctions.length}個`);

    // 重複を検出
    console.log('🕵️  重複を検出しています...');
    const results = detectDuplicates(allFunctions);

    // レポート生成
    generateReport(results);
    
    // 詳細レポートをファイルに保存
    const reportPath = saveDetailedReport(results);

    console.log('\n✅ 重複関数検出が完了しました！');
    
    return results;

  } catch (error) {
    console.error('❌ エラーが発生しました:', error.message);
    process.exit(1);
  }
}

// スクリプトが直接実行された場合のみmainを呼び出す
if (require.main === module) {
  main();
}

module.exports = {
  extractFunctions,
  detectDuplicates,
  generateReport,
  main
};