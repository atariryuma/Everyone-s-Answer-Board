#!/usr/bin/env node

/**
 * 手動確認が必要なvar変換プログラム
 * 残りのvarを安全にconst/letに変換
 */

const fs = require('fs');
const path = require('path');

class ManualVarConverter {
  constructor() {
    this.backupDir = path.join(__dirname, '../backups/manual-var-conversion-' + Date.now());
  }

  createBackup(filePath) {
    if (!fs.existsSync(this.backupDir)) {
      fs.mkdirSync(this.backupDir, { recursive: true });
    }
    
    const fileName = path.basename(filePath);
    const backupPath = path.join(this.backupDir, fileName);
    fs.copyFileSync(filePath, backupPath);
    console.log(`📁 バックアップ作成: ${backupPath}`);
  }

  /**
   * 変数の再代入パターンをより詳細に分析
   */
  analyzeVariableScope(lines, varLine, varName) {
    const startIndex = lines.indexOf(varLine);
    if (startIndex === -1) return { hasReassignment: false, scopeInfo: '' };
    
    let braceDepth = 0;
    let hasReassignment = false;
    let reassignmentLines = [];
    
    // 変数宣言の行から関数終了まで追跡
    for (let i = startIndex; i < lines.length; i++) {
      const line = lines[i].trim();
      
      // ブレースの深度を追跡
      const openBraces = (line.match(/{/g) || []).length;
      const closeBraces = (line.match(/}/g) || []).length;
      braceDepth += openBraces - closeBraces;
      
      // 関数スコープを出たら終了
      if (i > startIndex && braceDepth < 0) break;
      
      // 再代入パターンをチェック
      const reassignPattern = new RegExp(`\\b${varName}\\s*=(?!=)`, 'g');
      if (i > startIndex && line.match(reassignPattern) && !line.includes('var')) {
        hasReassignment = true;
        reassignmentLines.push(i + 1);
      }
    }
    
    return {
      hasReassignment,
      reassignmentLines,
      scopeInfo: hasReassignment ? '再代入あり' : '再代入なし'
    };
  }

  /**
   * database.gsの残りvar変換
   */
  convertDatabaseVars(content) {
    const lines = content.split('\n');
    let converted = content;
    const changes = [];

    // 安全にconstに変換できるパターン
    const safeConstPatterns = [
      // API呼び出しの結果代入（再代入されないもの）
      { pattern: /^(\s*)var\s+(dbId\s*=\s*props\.getProperty.*);?$/gm, name: 'dbId' },
      { pattern: /^(\s*)var\s+(sheetName\s*=\s*DB_CONFIG\.SHEET_NAME.*);?$/gm, name: 'sheetName' },
      { pattern: /^(\s*)var\s+(values\s*=\s*data\.valueRanges.*);?$/gm, name: 'values' },
      { pattern: /^(\s*)var\s+(headers\s*=\s*values\[0\].*);?$/gm, name: 'headers' },
      { pattern: /^(\s*)var\s+(userIdIndex\s*=\s*headers\.indexOf.*);?$/gm, name: 'userIdIndex' },
      { pattern: /^(\s*)var\s+(emailIndex\s*=\s*headers\.indexOf.*);?$/gm, name: 'emailIndex' },
      { pattern: /^(\s*)var\s+(isActiveIndex\s*=\s*headers\.indexOf.*);?$/gm, name: 'isActiveIndex' },
      { pattern: /^(\s*)var\s+(spreadsheetIdIndex\s*=\s*headers\.indexOf.*);?$/gm, name: 'spreadsheetIdIndex' },
      { pattern: /^(\s*)var\s+(userRows\s*=\s*values\.slice.*);?$/gm, name: 'userRows' },
      { pattern: /^(\s*)var\s+(response\s*=\s*UrlFetchApp\.fetch.*);?$/gm, name: 'response' },
      { pattern: /^(\s*)var\s+(url\s*=\s*[^;]+);?$/gm, name: 'url' },
      { pattern: /^(\s*)var\s+(responseCode\s*=\s*response\.getResponseCode.*);?$/gm, name: 'responseCode' },
      { pattern: /^(\s*)var\s+(responseText\s*=\s*response\.getContentText.*);?$/gm, name: 'responseText' },
      { pattern: /^(\s*)var\s+(baseUrl\s*=\s*service\.baseUrl.*);?$/gm, name: 'baseUrl' },
      { pattern: /^(\s*)var\s+(accessToken\s*=\s*service\.accessToken.*);?$/gm, name: 'accessToken' },
      { pattern: /^(\s*)var\s+(result\s*=.*JSON\.parse.*);?$/gm, name: 'result' },
      { pattern: /^(\s*)var\s+(sheetCount\s*=\s*result\.sheets\.length.*);?$/gm, name: 'sheetCount' },
    ];

    // letに変換すべきパターン（再代入があるもの）
    const letPatterns = [
      { pattern: /^(\s*)var\s+(rowIndex\s*=\s*-1.*);?$/gm, name: 'rowIndex' },
      { pattern: /^(\s*)var\s+(needsUpdate\s*=\s*false.*);?$/gm, name: 'needsUpdate' },
      { pattern: /^(\s*)var\s+(updateData\s*=\s*\{\}.*);?$/gm, name: 'updateData' },
      { pattern: /^(\s*)var\s+(updateSuccess\s*=\s*false.*);?$/gm, name: 'updateSuccess' },
      { pattern: /^(\s*)var\s+(retryCount\s*=\s*0.*);?$/gm, name: 'retryCount' },
      { pattern: /^(\s*)var\s+(userFound\s*=\s*false.*);?$/gm, name: 'userFound' },
      { pattern: /^(\s*)var\s+(userRowIndex\s*=\s*-1.*);?$/gm, name: 'userRowIndex' },
      { pattern: /^(\s*)var\s+(shouldClearAll\s*=\s*false.*);?$/gm, name: 'shouldClearAll' },
    ];

    // 特殊パターン（未初期化変数をletに）
    const uninitializedPatterns = [
      { pattern: /^(\s*)var\s+(accessToken);?\s*$/gm, name: 'accessToken' },
      { pattern: /^(\s*)var\s+(targetSheetId)\s*=\s*null;?$/gm, name: 'targetSheetId' },
    ];

    // const変換
    safeConstPatterns.forEach(({ pattern, name }) => {
      converted = converted.replace(pattern, (match, indent, declaration) => {
        changes.push(`const変換: var ${name} → const ${name}`);
        return `${indent}const ${declaration};`;
      });
    });

    // let変換
    letPatterns.forEach(({ pattern, name }) => {
      converted = converted.replace(pattern, (match, indent, declaration) => {
        changes.push(`let変換: var ${name} → let ${name}`);
        return `${indent}let ${declaration};`;
      });
    });

    // 未初期化変数のlet変換
    uninitializedPatterns.forEach(({ pattern, name }) => {
      converted = converted.replace(pattern, (match, indent, varName) => {
        changes.push(`let変換(未初期化): var ${name} → let ${name}`);
        return `${indent}let ${varName};`;
      });
    });

    return { converted, changes };
  }

  /**
   * ファイル処理
   */
  processFile(filePath) {
    console.log(`\n🔍 手動変換処理: ${path.basename(filePath)}`);
    
    const content = fs.readFileSync(filePath, 'utf8');
    const originalVarCount = (content.match(/\bvar\s+\w+/g) || []).length;
    
    console.log(`📊 変換前のvar数: ${originalVarCount}`);
    
    if (originalVarCount === 0) {
      console.log(`✅ 変換対象なし`);
      return;
    }

    this.createBackup(filePath);

    let result;
    if (filePath.includes('database.gs')) {
      result = this.convertDatabaseVars(content);
    } else {
      console.log(`ℹ️ ${path.basename(filePath)}は対応していません`);
      return;
    }

    const { converted, changes } = result;
    const finalVarCount = (converted.match(/\bvar\s+\w+/g) || []).length;
    
    console.log(`📊 変換後のvar数: ${finalVarCount}`);
    console.log(`📊 変換された数: ${originalVarCount - finalVarCount}`);
    
    if (changes.length > 0) {
      console.log(`\n✅ 適用された変更:`);
      changes.forEach(change => console.log(`   - ${change}`));
      
      fs.writeFileSync(filePath, converted);
      console.log(`💾 ファイル更新完了`);
      
      if (finalVarCount > 0) {
        console.log(`\n⚠️ 残りのvar (${finalVarCount}個) - 手動確認が必要:`);
        const remainingVars = converted.match(/^\s*var\s+\w+.*/gm);
        if (remainingVars) {
          remainingVars.slice(0, 10).forEach(line => {
            console.log(`   ${line.trim()}`);
          });
          if (remainingVars.length > 10) {
            console.log(`   ... 他${remainingVars.length - 10}件`);
          }
        }
      }
    } else {
      console.log(`ℹ️ 変換可能なパターンなし`);
    }
  }

  run(targetFiles) {
    console.log(`🔧 手動var変換プログラム開始`);
    console.log(`🛡️ バックアップディレクトリ: ${this.backupDir}`);
    
    targetFiles.forEach(filePath => {
      this.processFile(filePath);
    });
    
    console.log(`\n✅ 手動変換処理完了`);
  }
}

// メイン実行
if (require.main === module) {
  const targetFiles = [
    '/Users/ryuma/Everyone-s-Answer-Board/src/database.gs'
  ];
  
  const converter = new ManualVarConverter();
  converter.run(targetFiles);
}

module.exports = ManualVarConverter;