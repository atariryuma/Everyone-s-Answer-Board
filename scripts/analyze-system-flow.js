#!/usr/bin/env node

/**
 * システム内フロー調査・関数整合性チェックスクリプト
 * CLAUDE.md準拠のシステム構造解析ツール
 */

const fs = require('fs');
const path = require('path');

const SRC_DIR = path.join(__dirname, '../src');

/**
 * システム内の全関数を解析
 */
class SystemFlowAnalyzer {
  constructor() {
    this.functions = new Map(); // 関数定義
    this.functionCalls = new Map(); // 関数呼び出し
    this.googleScriptCalls = new Map(); // google.script.run呼び出し
    this.issues = [];
    this.stats = {
      totalFunctions: 0,
      totalCalls: 0,
      unusedFunctions: 0,
      missingFunctions: 0,
      frontendBackendMismatches: 0
    };
  }

  /**
   * メイン解析処理
   */
  async analyze() {
    console.log('🔍 システムフロー解析開始...\n');
    
    // 1. 全ファイルの関数定義を収集
    await this.collectFunctionDefinitions();
    
    // 2. 全ファイルの関数呼び出しを収集
    await this.collectFunctionCalls();
    
    // 3. フロントエンドのgoogle.script.run呼び出しを収集
    await this.collectGoogleScriptCalls();
    
    // 4. 整合性チェック
    this.checkConsistency();
    
    // 5. レポート生成
    this.generateReport();
    
    return {
      functions: this.functions,
      calls: this.functionCalls,
      googleScriptCalls: this.googleScriptCalls,
      issues: this.issues,
      stats: this.stats
    };
  }

  /**
   * 関数定義を収集
   */
  async collectFunctionDefinitions() {
    const files = this.getFiles(SRC_DIR);
    
    for (const filePath of files) {
      const content = fs.readFileSync(filePath, 'utf8');
      const relativePath = path.relative(SRC_DIR, filePath);
      
      // JavaScript/GAS関数定義パターン
      const functionPatterns = [
        /^(?:async\s+)?function\s+(\w+)\s*\(/gm,           // function name()
        /^const\s+(\w+)\s*=\s*(?:async\s+)?\([^)]*\)\s*=>/gm, // const name = () =>
        /^let\s+(\w+)\s*=\s*(?:async\s+)?\([^)]*\)\s*=>/gm,   // let name = () =>
        /(\w+)\s*:\s*(?:async\s+)?function\s*\(/gm,       // name: function()
        /(\w+)\s*:\s*(?:async\s+)?\([^)]*\)\s*=>/gm       // name: () =>
      ];
      
      for (const pattern of functionPatterns) {
        let match;
        while ((match = pattern.exec(content)) !== null) {
          const functionName = match[1];
          if (this.isValidFunctionName(functionName)) {
            if (!this.functions.has(functionName)) {
              this.functions.set(functionName, []);
            }
            this.functions.get(functionName).push({
              file: relativePath,
              line: this.getLineNumber(content, match.index),
              type: this.getFunctionType(relativePath)
            });
          }
        }
      }
    }
    
    this.stats.totalFunctions = this.functions.size;
    console.log(`✅ 関数定義収集完了: ${this.stats.totalFunctions}個の関数`);
  }

  /**
   * 関数呼び出しを収集
   */
  async collectFunctionCalls() {
    const files = this.getFiles(SRC_DIR);
    
    for (const filePath of files) {
      const content = fs.readFileSync(filePath, 'utf8');
      const relativePath = path.relative(SRC_DIR, filePath);
      
      // 関数呼び出しパターン
      const callPatterns = [
        /(\w+)\s*\(/g,                    // name()
        /\.(\w+)\s*\(/g,                  // obj.name()
        /\[['"](\w+)['"]\]\s*\(/g        // obj['name']()
      ];
      
      for (const pattern of callPatterns) {
        let match;
        while ((match = pattern.exec(content)) !== null) {
          const functionName = match[1];
          if (this.isValidFunctionName(functionName) && !this.isBuiltinFunction(functionName)) {
            if (!this.functionCalls.has(functionName)) {
              this.functionCalls.set(functionName, []);
            }
            this.functionCalls.get(functionName).push({
              file: relativePath,
              line: this.getLineNumber(content, match.index),
              context: this.getContext(content, match.index)
            });
          }
        }
      }
    }
    
    this.stats.totalCalls = Array.from(this.functionCalls.values()).reduce((sum, calls) => sum + calls.length, 0);
    console.log(`✅ 関数呼び出し収集完了: ${this.stats.totalCalls}個の呼び出し`);
  }

  /**
   * google.script.run呼び出しを収集
   */
  async collectGoogleScriptCalls() {
    const files = this.getFiles(SRC_DIR).filter(f => f.endsWith('.html'));
    
    for (const filePath of files) {
      const content = fs.readFileSync(filePath, 'utf8');
      const relativePath = path.relative(SRC_DIR, filePath);
      
      // google.script.run呼び出しパターン
      const patterns = [
        /google\.script\.run\s*\.\s*(\w+)\s*\(/g,                    // google.script.run.functionName(
        /google\.script\.run\s*\[\s*['"](\w+)['"]\s*\]\s*\(/g,      // google.script.run['functionName'](
        /google\.script\.run\s*\[([^[\]]+)\]\s*\(/g                 // google.script.run[variable](
      ];
      
      for (const pattern of patterns) {
        let match;
        while ((match = pattern.exec(content)) !== null) {
          const functionName = match[1];
          if (functionName && this.isValidFunctionName(functionName)) {
            if (!this.googleScriptCalls.has(functionName)) {
              this.googleScriptCalls.set(functionName, []);
            }
            this.googleScriptCalls.get(functionName).push({
              file: relativePath,
              line: this.getLineNumber(content, match.index),
              context: this.getContext(content, match.index)
            });
          }
        }
      }
    }
    
    console.log(`✅ google.script.run呼び出し収集完了: ${this.googleScriptCalls.size}個の関数`);
  }

  /**
   * 整合性チェック
   */
  checkConsistency() {
    console.log('🔍 整合性チェック開始...\n');
    
    // 1. 未使用関数チェック
    this.checkUnusedFunctions();
    
    // 2. 未定義関数チェック
    this.checkUndefinedFunctions();
    
    // 3. フロントエンド・バックエンド不整合チェック
    this.checkFrontendBackendConsistency();
    
    // 4. CLAUDE.md準拠チェック
    this.checkClaudeMdCompliance();
    
    console.log(`✅ 整合性チェック完了: ${this.issues.length}個の問題を検出`);
  }

  /**
   * 未使用関数チェック
   */
  checkUnusedFunctions() {
    for (const [functionName, definitions] of this.functions) {
      const isCalled = this.functionCalls.has(functionName) || this.googleScriptCalls.has(functionName);
      const isEntryPoint = this.isEntryPointFunction(functionName);
      
      if (!isCalled && !isEntryPoint) {
        this.issues.push({
          type: 'UNUSED_FUNCTION',
          severity: 'WARNING',
          function: functionName,
          files: definitions.map(d => d.file),
          message: `未使用関数: ${functionName} が定義されていますが、どこからも呼び出されていません`
        });
        this.stats.unusedFunctions++;
      }
    }
  }

  /**
   * 未定義関数チェック
   */
  checkUndefinedFunctions() {
    // google.script.run呼び出しで未定義の関数をチェック
    for (const [functionName, calls] of this.googleScriptCalls) {
      if (!this.functions.has(functionName)) {
        this.issues.push({
          type: 'UNDEFINED_FUNCTION',
          severity: 'ERROR',
          function: functionName,
          calls: calls,
          message: `未定義関数: ${functionName} がフロントエンドから呼び出されていますが、バックエンドに定義がありません`
        });
        this.stats.missingFunctions++;
      }
    }

    // 一般的な関数呼び出しで未定義の関数をチェック
    for (const [functionName, calls] of this.functionCalls) {
      if (!this.functions.has(functionName) && !this.isBuiltinFunction(functionName)) {
        this.issues.push({
          type: 'UNDEFINED_FUNCTION',
          severity: 'ERROR',
          function: functionName,
          calls: calls,
          message: `未定義関数: ${functionName} が呼び出されていますが、定義が見つかりません`
        });
        this.stats.missingFunctions++;
      }
    }
  }

  /**
   * フロントエンド・バックエンド整合性チェック
   */
  checkFrontendBackendConsistency() {
    // 重要な関数の整合性チェック
    const criticalFunctions = ['getConfig', 'getData', 'connectDataSource', 'publishApplication'];
    
    for (const functionName of criticalFunctions) {
      const backendDefined = this.functions.has(functionName);
      const frontendCalls = this.googleScriptCalls.has(functionName);
      
      if (frontendCalls && !backendDefined) {
        this.issues.push({
          type: 'FRONTEND_BACKEND_MISMATCH',
          severity: 'ERROR',
          function: functionName,
          message: `フロントエンド・バックエンド不整合: ${functionName} がフロントエンドから呼び出されていますが、バックエンドに定義がありません`
        });
        this.stats.frontendBackendMismatches++;
      }
    }
  }

  /**
   * CLAUDE.md準拠チェック
   */
  checkClaudeMdCompliance() {
    // 非推奨パターンチェック
    const deprecatedPatterns = [
      { pattern: 'getCurrentConfig', replacement: 'getConfig', severity: 'WARNING' },
      { pattern: 'getPublishedSheetData', replacement: 'getData', severity: 'WARNING' }
    ];

    for (const { pattern, replacement, severity } of deprecatedPatterns) {
      if (this.functions.has(pattern)) {
        this.issues.push({
          type: 'DEPRECATED_FUNCTION',
          severity,
          function: pattern,
          replacement,
          message: `非推奨関数: ${pattern} は ${replacement} に置き換えてください`
        });
      }
    }

    // configJSON中心設計チェック
    this.checkConfigJsonUsage();
  }

  /**
   * configJSON使用状況チェック
   */
  checkConfigJsonUsage() {
    const files = this.getFiles(SRC_DIR);
    
    for (const filePath of files) {
      const content = fs.readFileSync(filePath, 'utf8');
      const relativePath = path.relative(SRC_DIR, filePath);
      
      // 非推奨なDB列アクセスパターン
      const deprecatedPatterns = [
        /userInfo\.spreadsheetId/g,
        /userInfo\.sheetName/g,
        /userInfo\.\w+(?!\.configJson)/g
      ];
      
      for (const pattern of deprecatedPatterns) {
        let match;
        while ((match = pattern.exec(content)) !== null) {
          this.issues.push({
            type: 'DEPRECATED_PATTERN',
            severity: 'WARNING',
            file: relativePath,
            line: this.getLineNumber(content, match.index),
            pattern: match[0],
            message: `非推奨パターン: ${match[0]} は configJSON中心設計に従って config.${match[0].split('.')[1]} を使用してください`
          });
        }
      }
    }
  }

  /**
   * レポート生成
   */
  generateReport() {
    console.log('\n' + '='.repeat(80));
    console.log('📊 システムフロー解析レポート');
    console.log('='.repeat(80));

    // 統計情報
    console.log('\n📈 統計情報:');
    console.log(`  総関数数: ${this.stats.totalFunctions}`);
    console.log(`  総呼び出し数: ${this.stats.totalCalls}`);
    console.log(`  未使用関数: ${this.stats.unusedFunctions}`);
    console.log(`  未定義関数: ${this.stats.missingFunctions}`);
    console.log(`  フロント・バック不整合: ${this.stats.frontendBackendMismatches}`);

    // 問題一覧
    if (this.issues.length > 0) {
      console.log('\n⚠️  問題一覧:');
      
      const errorIssues = this.issues.filter(i => i.severity === 'ERROR');
      const warningIssues = this.issues.filter(i => i.severity === 'WARNING');
      
      if (errorIssues.length > 0) {
        console.log(`\n❌ エラー (${errorIssues.length}件):`);
        errorIssues.forEach((issue, index) => {
          console.log(`  ${index + 1}. ${issue.message}`);
          if (issue.files) console.log(`     ファイル: ${issue.files.join(', ')}`);
          if (issue.calls) console.log(`     呼び出し元: ${issue.calls.map(c => `${c.file}:${c.line}`).join(', ')}`);
        });
      }
      
      if (warningIssues.length > 0) {
        console.log(`\n⚠️  警告 (${warningIssues.length}件):`);
        warningIssues.forEach((issue, index) => {
          console.log(`  ${index + 1}. ${issue.message}`);
          if (issue.replacement) console.log(`     推奨: ${issue.replacement}`);
        });
      }
    } else {
      console.log('\n✅ 問題は見つかりませんでした！');
    }

    // 重要な関数のフロー
    console.log('\n🔄 重要な関数のフロー:');
    const criticalFunctions = ['doGet', 'getConfig', 'getData', 'connectDataSource'];
    
    for (const functionName of criticalFunctions) {
      const definitions = this.functions.get(functionName) || [];
      const calls = this.functionCalls.get(functionName) || [];
      const googleCalls = this.googleScriptCalls.get(functionName) || [];
      
      console.log(`\n  📋 ${functionName}:`);
      console.log(`    定義: ${definitions.length > 0 ? definitions.map(d => `${d.file}:${d.line}`).join(', ') : '❌ なし'}`);
      console.log(`    バックエンド呼び出し: ${calls.length}件`);
      console.log(`    フロントエンド呼び出し: ${googleCalls.length}件`);
    }

    console.log('\n' + '='.repeat(80));
  }

  /**
   * ヘルパーメソッド群
   */
  getFiles(dir, extensions = ['.gs', '.html']) {
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
    return /^[a-zA-Z_$][a-zA-Z0-9_$]*$/.test(name) && 
           name.length > 1 && 
           !['if', 'for', 'while', 'function', 'const', 'let', 'var'].includes(name);
  }

  isBuiltinFunction(name) {
    const builtins = [
      'console', 'log', 'error', 'warn', 'info',
      'setTimeout', 'setInterval', 'clearTimeout', 'clearInterval',
      'JSON', 'parse', 'stringify',
      'Object', 'keys', 'values', 'entries', 'assign', 'freeze',
      'Array', 'map', 'filter', 'reduce', 'forEach', 'find', 'includes',
      'String', 'slice', 'substring', 'split', 'join', 'replace',
      'Math', 'floor', 'ceil', 'round', 'random',
      'Date', 'now', 'toISOString',
      'Promise', 'resolve', 'reject', 'all',
      'document', 'getElementById', 'querySelector', 'addEventListener',
      'window', 'location', 'reload',
      'SpreadsheetApp', 'DriveApp', 'UrlFetchApp', 'PropertiesService',
      'include', 'require'
    ];
    return builtins.includes(name);
  }

  isEntryPointFunction(name) {
    const entryPoints = ['doGet', 'doPost', 'onInstall', 'onOpen'];
    return entryPoints.includes(name);
  }

  getFunctionType(filePath) {
    if (filePath.endsWith('.html')) return 'frontend';
    if (filePath.endsWith('.gs')) return 'backend';
    return 'unknown';
  }

  getLineNumber(content, index) {
    return content.substring(0, index).split('\n').length;
  }

  getContext(content, index, contextLength = 50) {
    const start = Math.max(0, index - contextLength);
    const end = Math.min(content.length, index + contextLength);
    return content.substring(start, end).replace(/\s+/g, ' ').trim();
  }
}

/**
 * メイン実行
 */
async function main() {
  const analyzer = new SystemFlowAnalyzer();
  
  try {
    const result = await analyzer.analyze();
    
    // 結果をJSONファイルにも出力
    const outputPath = path.join(__dirname, '../analysis-results.json');
    fs.writeFileSync(outputPath, JSON.stringify(result, null, 2));
    console.log(`\n📄 詳細結果は ${outputPath} に保存されました`);
    
    // 問題がある場合は終了コード1で終了
    const hasErrors = result.issues.some(issue => issue.severity === 'ERROR');
    process.exit(hasErrors ? 1 : 0);
    
  } catch (error) {
    console.error('❌ 解析中にエラーが発生しました:', error);
    process.exit(1);
  }
}

// 直接実行の場合はメイン関数を実行
if (require.main === module) {
  main();
}

module.exports = { SystemFlowAnalyzer };