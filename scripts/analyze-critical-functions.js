#!/usr/bin/env node

/**
 * 重要関数の整合性チェックスクリプト（フォーカス版）
 * フロントエンド・バックエンド連携とCLAUDE.md準拠をチェック
 */

const fs = require('fs');
const path = require('path');

const SRC_DIR = path.join(__dirname, '../src');

class CriticalFunctionAnalyzer {
  constructor() {
    this.backendFunctions = new Map();  // .gsファイルの関数
    this.frontendCalls = new Map();     // HTMLファイルのgoogle.script.run呼び出し
    this.deprecatedPatterns = new Map(); // 非推奨パターン
    this.issues = [];
  }

  async analyze() {
    console.log('🎯 重要関数整合性チェック開始...\n');
    
    // 1. バックエンド関数定義収集
    await this.collectBackendFunctions();
    
    // 2. フロントエンドの関数呼び出し収集
    await this.collectFrontendCalls();
    
    // 3. CLAUDE.md準拠チェック
    await this.checkClaudeMdCompliance();
    
    // 4. 重要な不整合チェック
    this.checkCriticalMismatches();
    
    // 5. レポート生成
    this.generateReport();
    
    return { issues: this.issues };
  }

  async collectBackendFunctions() {
    const gsFiles = this.getFiles(SRC_DIR, ['.gs']);
    
    for (const filePath of gsFiles) {
      const content = fs.readFileSync(filePath, 'utf8');
      const relativePath = path.relative(SRC_DIR, filePath);
      
      // 関数定義パターン（GASファイル用）
      const functionPattern = /^(?:async\s+)?function\s+(\w+)\s*\(/gm;
      
      let match;
      while ((match = functionPattern.exec(content)) !== null) {
        const functionName = match[1];
        if (!this.backendFunctions.has(functionName)) {
          this.backendFunctions.set(functionName, []);
        }
        this.backendFunctions.get(functionName).push({
          file: relativePath,
          line: this.getLineNumber(content, match.index)
        });
      }
    }
    
    console.log(`✅ バックエンド関数定義: ${this.backendFunctions.size}個`);
  }

  async collectFrontendCalls() {
    const htmlFiles = this.getFiles(SRC_DIR, ['.html']);
    
    for (const filePath of htmlFiles) {
      const content = fs.readFileSync(filePath, 'utf8');
      const relativePath = path.relative(SRC_DIR, filePath);
      
      // google.script.run呼び出しパターン
      const patterns = [
        /google\.script\.run\s*\.\s*(\w+)\s*\(/g,
        /google\.script\.run\s*\[\s*['"`](\w+)['"`]\s*\]\s*\(/g
      ];
      
      for (const pattern of patterns) {
        let match;
        while ((match = pattern.exec(content)) !== null) {
          const functionName = match[1];
          if (this.isValidFunctionName(functionName)) {
            if (!this.frontendCalls.has(functionName)) {
              this.frontendCalls.set(functionName, []);
            }
            this.frontendCalls.get(functionName).push({
              file: relativePath,
              line: this.getLineNumber(content, match.index)
            });
          }
        }
      }
    }
    
    console.log(`✅ フロントエンド関数呼び出し: ${this.frontendCalls.size}個`);
  }

  async checkClaudeMdCompliance() {
    const gsFiles = this.getFiles(SRC_DIR, ['.gs']);
    
    for (const filePath of gsFiles) {
      const content = fs.readFileSync(filePath, 'utf8');
      const relativePath = path.relative(SRC_DIR, filePath);
      
      // 非推奨なuserInfo.*パターンをチェック
      const deprecatedPatterns = [
        { pattern: /userInfo\.spreadsheetId/g, issue: 'configJSON中心設計違反: userInfo.spreadsheetIdではなくconfig.spreadsheetIdを使用' },
        { pattern: /userInfo\.sheetName/g, issue: 'configJSON中心設計違反: userInfo.sheetNameではなくconfig.sheetNameを使用' },
        { pattern: /userInfo\.(?!configJson)(?!userId)(?!userEmail)(?!isActive)(?!lastModified)\w+/g, issue: 'configJSON中心設計違反: userInfo.フィールドは5フィールド構造以外禁止' }
      ];
      
      for (const { pattern, issue } of deprecatedPatterns) {
        let match;
        while ((match = pattern.exec(content)) !== null) {
          this.issues.push({
            type: 'CLAUDE_MD_VIOLATION',
            severity: 'WARNING',
            file: relativePath,
            line: this.getLineNumber(content, match.index),
            pattern: match[0],
            message: issue
          });
        }
      }
    }
  }

  checkCriticalMismatches() {
    // 重要な関数の整合性チェック
    const criticalFunctions = [
      'getConfig', 'getData', 'connectDataSource', 'publishApplication',
      'saveDraftConfiguration', 'getCurrentBoardInfoAndUrls', 'addReaction',
      'toggleHighlight', 'refreshBoardData'
    ];
    
    for (const functionName of criticalFunctions) {
      const hasBackendDefinition = this.backendFunctions.has(functionName);
      const hasFrontendCalls = this.frontendCalls.has(functionName);
      
      if (hasFrontendCalls && !hasBackendDefinition) {
        this.issues.push({
          type: 'CRITICAL_MISMATCH',
          severity: 'ERROR',
          function: functionName,
          message: `重要関数未定義: ${functionName} がフロントエンドから呼び出されていますが、バックエンドに定義がありません`,
          frontendCalls: this.frontendCalls.get(functionName)
        });
      }
      
      if (hasBackendDefinition && !hasFrontendCalls) {
        this.issues.push({
          type: 'UNUSED_BACKEND_FUNCTION',
          severity: 'WARNING',
          function: functionName,
          message: `未使用バックエンド関数: ${functionName} がバックエンドに定義されていますが、フロントエンドから呼び出されていません`,
          backendDefinitions: this.backendFunctions.get(functionName)
        });
      }
    }

    // 非推奨関数チェック
    const deprecatedFunctions = ['getCurrentConfig', 'getPublishedSheetData'];
    for (const functionName of deprecatedFunctions) {
      if (this.backendFunctions.has(functionName)) {
        this.issues.push({
          type: 'DEPRECATED_FUNCTION',
          severity: 'WARNING',
          function: functionName,
          message: `非推奨関数: ${functionName} は使用を停止してください`,
          backendDefinitions: this.backendFunctions.get(functionName)
        });
      }
    }
  }

  generateReport() {
    console.log('\n' + '='.repeat(80));
    console.log('🎯 重要関数整合性チェックレポート');
    console.log('='.repeat(80));

    // 統計情報
    const errorCount = this.issues.filter(i => i.severity === 'ERROR').length;
    const warningCount = this.issues.filter(i => i.severity === 'WARNING').length;
    
    console.log(`\n📈 検出した問題:`);
    console.log(`  ❌ エラー: ${errorCount}件`);
    console.log(`  ⚠️  警告: ${warningCount}件`);
    console.log(`  📊 総問題数: ${this.issues.length}件`);

    // エラー詳細
    const criticalErrors = this.issues.filter(i => i.severity === 'ERROR');
    if (criticalErrors.length > 0) {
      console.log('\n❌ 重要エラー:');
      criticalErrors.forEach((issue, index) => {
        console.log(`  ${index + 1}. ${issue.message}`);
        if (issue.frontendCalls) {
          const callLocations = issue.frontendCalls.map(c => `${c.file}:${c.line}`).join(', ');
          console.log(`     呼び出し元: ${callLocations}`);
        }
      });
    }

    // CLAUDE.md準拠違反
    const claudeMdViolations = this.issues.filter(i => i.type === 'CLAUDE_MD_VIOLATION');
    if (claudeMdViolations.length > 0) {
      console.log('\n⚠️  CLAUDE.md準拠違反:');
      
      // ファイル別にグループ化
      const violationsByFile = new Map();
      claudeMdViolations.forEach(violation => {
        if (!violationsByFile.has(violation.file)) {
          violationsByFile.set(violation.file, []);
        }
        violationsByFile.get(violation.file).push(violation);
      });
      
      violationsByFile.forEach((violations, file) => {
        console.log(`  📁 ${file}: ${violations.length}件`);
        violations.slice(0, 3).forEach(v => {
          console.log(`    - line ${v.line}: ${v.pattern} → ${v.message}`);
        });
        if (violations.length > 3) {
          console.log(`    - ... 他 ${violations.length - 3} 件`);
        }
      });
    }

    // 重要関数の状況
    console.log('\n🔄 重要関数の状況:');
    const criticalFunctions = ['getConfig', 'getData', 'connectDataSource', 'publishApplication'];
    
    criticalFunctions.forEach(functionName => {
      const backendDef = this.backendFunctions.get(functionName);
      const frontendCalls = this.frontendCalls.get(functionName);
      
      console.log(`  📋 ${functionName}:`);
      console.log(`    バックエンド定義: ${backendDef ? '✅' : '❌'} ${backendDef ? `(${backendDef.map(d => d.file).join(', ')})` : ''}`);
      console.log(`    フロントエンド呼び出し: ${frontendCalls ? '✅' : '❌'} ${frontendCalls ? `(${frontendCalls.length}件)` : ''}`);
    });

    console.log('\n' + '='.repeat(80));
    
    if (errorCount > 0) {
      console.log('❌ 重要なエラーがあります。修正が必要です。');
      return false;
    } else if (warningCount > 0) {
      console.log('⚠️  警告があります。改善をお勧めします。');
      return true;
    } else {
      console.log('✅ 重要な問題は見つかりませんでした！');
      return true;
    }
  }

  // ヘルパーメソッド
  getFiles(dir, extensions) {
    const files = [];
    const items = fs.readdirSync(dir);
    
    for (const item of items) {
      const fullPath = path.join(dir, item);
      const stat = fs.statSync(fullPath);
      
      if (stat.isDirectory()) {
        files.push(...this.getFiles(fullPath, extensions));
      } else if (extensions.some(ext => item.endsWith(ext))) {
        files.push(fullPath);
      }
    }
    
    return files;
  }

  isValidFunctionName(name) {
    return /^[a-zA-Z_$][a-zA-Z0-9_$]*$/.test(name) && name.length > 1;
  }

  getLineNumber(content, index) {
    return content.substring(0, index).split('\n').length;
  }
}

// メイン実行
async function main() {
  const analyzer = new CriticalFunctionAnalyzer();
  
  try {
    const result = await analyzer.analyze();
    
    // 結果をJSONファイルに保存
    const outputPath = path.join(__dirname, '../critical-analysis-results.json');
    fs.writeFileSync(outputPath, JSON.stringify(result, null, 2));
    console.log(`\n📄 詳細結果は ${outputPath} に保存されました`);
    
  } catch (error) {
    console.error('❌ 解析中にエラーが発生しました:', error);
    process.exit(1);
  }
}

if (require.main === module) {
  main();
}

module.exports = { CriticalFunctionAnalyzer };