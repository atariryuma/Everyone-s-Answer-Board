#!/usr/bin/env node

/**
 * 未使用・孤立関数検出スクリプト
 * システム内でフローに関係ない関数を特定
 */

const fs = require('fs');
const path = require('path');

const SRC_DIR = path.join(__dirname, '../src');

class UnusedFunctionDetector {
  constructor() {
    this.functionDefinitions = new Map(); // 関数定義
    this.functionCalls = new Map(); // 関数呼び出し
    this.frontendCalls = new Map(); // フロントエンドからの呼び出し
    this.unusedFunctions = [];
    this.orphanedFunctions = [];
    this.deadCode = [];
  }

  async detect() {
    console.log('🔍 未使用・孤立関数検出開始...\n');
    
    // 1. 全関数定義を収集
    await this.collectFunctionDefinitions();
    
    // 2. 全関数呼び出しを収集
    await this.collectFunctionCalls();
    
    // 3. フロントエンドからの呼び出しを収集
    await this.collectFrontendCalls();
    
    // 4. 未使用関数を特定
    this.identifyUnusedFunctions();
    
    // 5. 孤立関数を特定
    this.identifyOrphanedFunctions();
    
    // 6. デッドコードを特定
    this.identifyDeadCode();
    
    // 7. レポート生成
    this.generateReport();
    
    return {
      unused: this.unusedFunctions,
      orphaned: this.orphanedFunctions,
      deadCode: this.deadCode,
      stats: {
        totalFunctions: this.functionDefinitions.size,
        usedFunctions: this.functionCalls.size,
        unusedCount: this.unusedFunctions.length
      }
    };
  }

  async collectFunctionDefinitions() {
    const gsFiles = this.getFiles(SRC_DIR, ['.gs']);
    
    for (const filePath of gsFiles) {
      const content = fs.readFileSync(filePath, 'utf8');
      const relativePath = path.relative(SRC_DIR, filePath);
      
      // 関数定義パターン
      const functionPattern = /^(?:async\s+)?function\s+(\w+)\s*\([^)]*\)/gm;
      let match;
      
      while ((match = functionPattern.exec(content)) !== null) {
        const functionName = match[1];
        if (this.isValidFunctionName(functionName)) {
          if (!this.functionDefinitions.has(functionName)) {
            this.functionDefinitions.set(functionName, []);
          }
          this.functionDefinitions.get(functionName).push({
            file: relativePath,
            line: this.getLineNumber(content, match.index),
            fullMatch: match[0],
            context: this.getContext(content, match.index)
          });
        }
      }
    }
    
    console.log(`📋 関数定義収集: ${this.functionDefinitions.size}個の関数`);
  }

  async collectFunctionCalls() {
    const files = this.getFiles(SRC_DIR, ['.gs', '.html']);
    
    for (const filePath of files) {
      const content = fs.readFileSync(filePath, 'utf8');
      const relativePath = path.relative(SRC_DIR, filePath);
      
      // 関数呼び出しパターン
      const callPatterns = [
        /(\w+)\s*\(/g,                    // functionName()
        /\.(\w+)\s*\(/g,                  // obj.functionName()
        /\[['"`](\w+)['"`]\]\s*\(/g       // obj['functionName']()
      ];
      
      for (const pattern of callPatterns) {
        let match;
        while ((match = pattern.exec(content)) !== null) {
          const functionName = match[1];
          if (this.isValidFunctionName(functionName) && 
              !this.isBuiltinFunction(functionName) &&
              this.functionDefinitions.has(functionName)) {
            
            if (!this.functionCalls.has(functionName)) {
              this.functionCalls.set(functionName, []);
            }
            this.functionCalls.get(functionName).push({
              file: relativePath,
              line: this.getLineNumber(content, match.index),
              context: this.getContext(content, match.index, 30)
            });
          }
        }
      }
    }
    
    console.log(`📞 関数呼び出し収集: ${this.functionCalls.size}個の関数が呼び出されている`);
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
    
    console.log(`🌐 フロントエンド呼び出し収集: ${this.frontendCalls.size}個の関数`);
  }

  identifyUnusedFunctions() {
    console.log('🔍 未使用関数特定中...');
    
    for (const [functionName, definitions] of this.functionDefinitions) {
      const hasBackendCalls = this.functionCalls.has(functionName);
      const hasFrontendCalls = this.frontendCalls.has(functionName);
      const isEntryPoint = this.isEntryPointFunction(functionName);
      const isExported = this.isExportedFunction(functionName);
      
      if (!hasBackendCalls && !hasFrontendCalls && !isEntryPoint && !isExported) {
        this.unusedFunctions.push({
          name: functionName,
          definitions: definitions,
          reason: 'どこからも呼び出されていない',
          category: 'COMPLETELY_UNUSED'
        });
      }
    }
  }

  identifyOrphanedFunctions() {
    console.log('🏝️  孤立関数特定中...');
    
    for (const [functionName, definitions] of this.functionDefinitions) {
      const hasBackendCalls = this.functionCalls.has(functionName);
      const hasFrontendCalls = this.frontendCalls.has(functionName);
      const isEntryPoint = this.isEntryPointFunction(functionName);
      
      // バックエンドでは使われているが、フロントエンドから呼ばれていない関数
      if (hasBackendCalls && !hasFrontendCalls && !isEntryPoint) {
        const callCount = this.functionCalls.get(functionName)?.length || 0;
        if (callCount < 3) { // 使用頻度が低い
          this.orphanedFunctions.push({
            name: functionName,
            definitions: definitions,
            callCount,
            reason: 'バックエンドでの使用頻度が低い',
            category: 'LOW_USAGE'
          });
        }
      }
      
      // フロントエンドから呼ばれているが、定義が見つからない
      if (!hasBackendCalls && hasFrontendCalls) {
        this.orphanedFunctions.push({
          name: functionName,
          definitions: definitions,
          frontendCalls: this.frontendCalls.get(functionName),
          reason: 'フロントエンドから呼ばれているが実装が不完全',
          category: 'FRONTEND_ORPHAN'
        });
      }
    }
  }

  identifyDeadCode() {
    console.log('💀 デッドコード特定中...');
    
    // 特定のパターンのデッドコードを検出
    const gsFiles = this.getFiles(SRC_DIR, ['.gs']);
    
    for (const filePath of gsFiles) {
      const content = fs.readFileSync(filePath, 'utf8');
      const relativePath = path.relative(SRC_DIR, filePath);
      
      // デッドコードパターン
      const deadCodePatterns = [
        {
          pattern: /\/\/\s*TODO:.*remove.*|\/\/\s*FIXME:.*remove.*/gi,
          type: 'TODO_REMOVE'
        },
        {
          pattern: /\/\*[\s\S]*?deprecated[\s\S]*?\*\//gi,
          type: 'DEPRECATED_BLOCK'
        },
        {
          pattern: /function\s+\w+.*?\{[\s\S]*?\/\/.*obsolete.*[\s\S]*?\}/gi,
          type: 'OBSOLETE_FUNCTION'
        }
      ];
      
      for (const {pattern, type} of deadCodePatterns) {
        let match;
        while ((match = pattern.exec(content)) !== null) {
          this.deadCode.push({
            file: relativePath,
            line: this.getLineNumber(content, match.index),
            type,
            code: match[0].substring(0, 100) + '...',
            reason: 'コメントで削除対象と明記されている'
          });
        }
      }
    }
  }

  generateReport() {
    console.log('\n' + '='.repeat(80));
    console.log('🔍 未使用・孤立関数検出レポート');
    console.log('='.repeat(80));

    const totalFunctions = this.functionDefinitions.size;
    const usedFunctions = this.functionCalls.size;
    const unusedCount = this.unusedFunctions.length;
    const orphanedCount = this.orphanedFunctions.length;
    
    console.log(`\n📊 統計情報:`);
    console.log(`  総関数数: ${totalFunctions}`);
    console.log(`  使用中関数数: ${usedFunctions}`);
    console.log(`  完全未使用関数: ${unusedCount}個`);
    console.log(`  孤立関数: ${orphanedCount}個`);
    console.log(`  使用効率: ${Math.round((usedFunctions / totalFunctions) * 100)}%`);

    // 完全未使用関数
    if (this.unusedFunctions.length > 0) {
      console.log(`\n❌ 完全未使用関数 (${this.unusedFunctions.length}個):`);
      console.log('   これらの関数は削除可能です');
      
      this.unusedFunctions.forEach((func, index) => {
        console.log(`\n  ${index + 1}. ${func.name}()`);
        func.definitions.forEach(def => {
          console.log(`     📁 ${def.file}:${def.line}`);
          console.log(`     📝 ${def.fullMatch}`);
        });
        console.log(`     💡 理由: ${func.reason}`);
      });
    }

    // 孤立関数
    if (this.orphanedFunctions.length > 0) {
      console.log(`\n🏝️  孤立関数 (${this.orphanedFunctions.length}個):`);
      console.log('   これらの関数は使用頻度が低いか、不完全です');
      
      this.orphanedFunctions.forEach((func, index) => {
        console.log(`\n  ${index + 1}. ${func.name}()`);
        if (func.definitions) {
          func.definitions.forEach(def => {
            console.log(`     📁 ${def.file}:${def.line}`);
          });
        }
        if (func.callCount !== undefined) {
          console.log(`     📞 呼び出し回数: ${func.callCount}回`);
        }
        if (func.frontendCalls) {
          console.log(`     🌐 フロントエンド呼び出し: ${func.frontendCalls.length}箇所`);
        }
        console.log(`     💡 理由: ${func.reason}`);
      });
    }

    // デッドコード
    if (this.deadCode.length > 0) {
      console.log(`\n💀 デッドコード (${this.deadCode.length}個):`);
      
      this.deadCode.forEach((code, index) => {
        console.log(`\n  ${index + 1}. ${code.file}:${code.line}`);
        console.log(`     📝 ${code.code}`);
        console.log(`     💡 理由: ${code.reason}`);
      });
    }

    // 推奨アクション
    console.log(`\n🚀 推奨アクション:`);
    
    if (this.unusedFunctions.length > 0) {
      console.log(`  1. 完全未使用関数 ${this.unusedFunctions.length} 個の削除`);
      console.log(`     💾 推定削減コード量: ${this.unusedFunctions.length * 15}行程度`);
    }
    
    if (this.orphanedFunctions.length > 0) {
      console.log(`  2. 孤立関数 ${this.orphanedFunctions.length} 個の見直し`);
      console.log(`     🔍 使用目的の確認と統合検討`);
    }
    
    if (this.deadCode.length > 0) {
      console.log(`  3. デッドコード ${this.deadCode.length} 箇所の削除`);
    }

    const cleanupBenefit = this.unusedFunctions.length + this.orphanedFunctions.length + this.deadCode.length;
    if (cleanupBenefit > 0) {
      console.log(`\n✨ クリーンアップ効果:`);
      console.log(`  コード品質向上: ${cleanupBenefit}項目の改善`);
      console.log(`  保守性向上: 不要コードの削除`);
      console.log(`  実行効率: 軽微な向上`);
    } else {
      console.log(`\n✅ 素晴らしい！フローに関係ない関数は見つかりませんでした。`);
      console.log(`  すべての関数が適切に使用されています。`);
    }

    console.log('\n' + '='.repeat(80));
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
      'include', 'require', 'push', 'pop', 'shift', 'unshift'
    ];
    return builtins.includes(name);
  }

  isEntryPointFunction(name) {
    const entryPoints = [
      'doGet', 'doPost', 'onInstall', 'onOpen', 'onEdit',
      'main', 'init', 'setup'
    ];
    return entryPoints.includes(name);
  }

  isExportedFunction(name) {
    // GASの特別な関数や明示的にエクスポートされた関数
    const exportedPatterns = [
      /^test/, /^debug/, /^admin/, /^system/,
      /^get/, /^set/, /^create/, /^update/, /^delete/
    ];
    return exportedPatterns.some(pattern => pattern.test(name));
  }

  getLineNumber(content, index) {
    return content.substring(0, index).split('\n').length;
  }

  getContext(content, index, length = 50) {
    const start = Math.max(0, index - length);
    const end = Math.min(content.length, index + length);
    return content.substring(start, end).replace(/\s+/g, ' ').trim();
  }
}

// メイン実行
async function main() {
  const detector = new UnusedFunctionDetector();
  
  try {
    const result = await detector.detect();
    
    // 結果をJSONファイルに保存
    const outputPath = path.join(__dirname, '../unused-functions-report.json');
    fs.writeFileSync(outputPath, JSON.stringify(result, null, 2));
    console.log(`\n📄 詳細結果は ${outputPath} に保存されました`);
    
  } catch (error) {
    console.error('❌ 検出中にエラーが発生しました:', error);
    process.exit(1);
  }
}

if (require.main === module) {
  main();
}

module.exports = { UnusedFunctionDetector };