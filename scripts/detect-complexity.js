#!/usr/bin/env node
/**
 * 重複・複雑な命名の機械的検出スクリプト
 * Usage: node scripts/detect-complexity.js
 */

const fs = require('fs');
const path = require('path');

// 検出対象パターン
const DETECTION_PATTERNS = {
  // 冗長・複雑な関数名
  VERBOSE_FUNCTIONS: [
    /function\s+([a-zA-Z_][a-zA-Z0-9_]*getCurrentConfig[a-zA-Z0-9_]*)/g,
    /function\s+([a-zA-Z_][a-zA-Z0-9_]*getAppConfig[a-zA-Z0-9_]*)/g,
    /function\s+([a-zA-Z_][a-zA-Z0-9_]*getUserConfig[a-zA-Z0-9_]*)/g,
    /function\s+([a-zA-Z_][a-zA-Z0-9_]*getPublishedSheetData[a-zA-Z0-9_]*)/g,
    /function\s+([a-zA-Z_][a-zA-Z0-9_]*getIncrementalSheetData[a-zA-Z0-9_]*)/g,
    /function\s+([a-zA-Z_][a-zA-Z0-9_]*addReactionBatch[a-zA-Z0-9_]*)/g,
  ],
  
  // 重複する変数宣言パターン
  DUPLICATE_VARS: [
    /const\s+currentUserEmail\s*=\s*UserManager\.getCurrentEmail\(\)/g,
    /const\s+spreadsheet\s*=\s*SpreadsheetApp\.openById/g,
    /const\s+userInfo\s*=\s*DB\.findUserByEmail/g,
    /const\s+config\s*=\s*ConfigManager\.getUserConfig/g,
  ],

  // 冗長な定数名
  VERBOSE_CONSTANTS: [
    /const\s+([A-Z_]*COLUMN_HEADERS[A-Z_]*)\s*=/g,
    /const\s+([A-Z_]*REACTION_KEYS[A-Z_]*)\s*=/g,
    /const\s+([A-Z_]*DELETE_LOG_SHEET_CONFIG[A-Z_]*)\s*=/g,
    /const\s+([A-Z_]*SYSTEM_CONSTANTS[A-Z_]*)\s*=/g,
  ],

  // 複雑な関数チェーン
  COMPLEX_CHAINS: [
    /App\.getConfig\(\)\.getUserConfig\([^)]*\)/g,
    /SYSTEM_CONSTANTS\.REACTIONS\.LABELS\.[A-Z_]+/g,
    /SYSTEM_CONSTANTS\.COLUMNS\.[A-Z_]+/g,
  ],

  // 長すぎる関数名（15文字以上）
  LONG_FUNCTION_NAMES: /function\s+([a-zA-Z_][a-zA-Z0-9_]{15,})/g,
};

// シンプルな名前の提案
const SIMPLIFICATION_SUGGESTIONS = {
  // 関数名の簡素化
  'getCurrentConfig': 'getConfig',
  'getAppConfig': 'getConfig', 
  'getUserConfig': 'getConfig',
  'getPublishedSheetData': 'getData',
  'getIncrementalSheetData': 'getIncrementalData',
  'addReactionBatch': 'addReactions',
  'getCurrentUserInfo': 'getUserInfo',
  'getAvailableSheets': 'getSheets',
  'refreshBoardData': 'refreshData',
  'clearActiveSheet': 'clearCache',
  'toggleHighlight': 'highlight',

  // 定数名の簡素化
  'SYSTEM_CONSTANTS': 'CONSTANTS',
  'COLUMN_HEADERS': 'HEADERS',
  'REACTION_KEYS': 'REACTIONS',
  'DELETE_LOG_SHEET_CONFIG': 'DELETE_CONFIG',

  // チェーンの簡素化
  'App.getConfig().getUserConfig': 'getConfig',
  'SYSTEM_CONSTANTS.REACTIONS.LABELS': 'REACTIONS',
  'SYSTEM_CONSTANTS.COLUMNS': 'COLUMNS',
};

class ComplexityDetector {
  constructor() {
    this.results = {
      verboseFunctions: [],
      duplicateVars: [],
      verboseConstants: [],
      complexChains: [],
      longNames: []
    };
    this.files = [];
  }

  // ファイル一覧を取得
  getSourceFiles(dir = 'src') {
    const files = [];
    const entries = fs.readdirSync(dir);
    
    for (const entry of entries) {
      const fullPath = path.join(dir, entry);
      const stat = fs.statSync(fullPath);
      
      if (stat.isDirectory()) {
        files.push(...this.getSourceFiles(fullPath));
      } else if (entry.match(/\.(gs|html)$/)) {
        files.push(fullPath);
      }
    }
    
    return files;
  }

  // パターン検出実行
  detectComplexity() {
    this.files = this.getSourceFiles();
    console.log(`🔍 検出対象ファイル: ${this.files.length}件`);

    for (const filePath of this.files) {
      try {
        const content = fs.readFileSync(filePath, 'utf8');
        this.analyzeFile(filePath, content);
      } catch (error) {
        console.error(`❌ ファイル読み込みエラー: ${filePath} - ${error.message}`);
      }
    }

    return this.generateReport();
  }

  // ファイル分析
  analyzeFile(filePath, content) {
    const lines = content.split('\n');
    
    // 冗長な関数名検出
    for (const pattern of DETECTION_PATTERNS.VERBOSE_FUNCTIONS) {
      let match;
      while ((match = pattern.exec(content)) !== null) {
        const lineNum = this.getLineNumber(content, match.index);
        this.results.verboseFunctions.push({
          file: filePath,
          line: lineNum,
          name: match[1],
          suggestion: this.getSuggestion(match[1])
        });
      }
    }

    // 重複変数検出
    for (const pattern of DETECTION_PATTERNS.DUPLICATE_VARS) {
      const matches = [...content.matchAll(pattern)];
      if (matches.length > 1) {
        for (const match of matches) {
          const lineNum = this.getLineNumber(content, match.index);
          this.results.duplicateVars.push({
            file: filePath,
            line: lineNum,
            pattern: match[0],
            count: matches.length
          });
        }
      }
    }

    // 複雑なチェーン検出
    for (const pattern of DETECTION_PATTERNS.COMPLEX_CHAINS) {
      let match;
      while ((match = pattern.exec(content)) !== null) {
        const lineNum = this.getLineNumber(content, match.index);
        this.results.complexChains.push({
          file: filePath,
          line: lineNum,
          chain: match[0],
          suggestion: this.getChainSuggestion(match[0])
        });
      }
    }

    // 長すぎる関数名検出
    let match;
    while ((match = DETECTION_PATTERNS.LONG_FUNCTION_NAMES.exec(content)) !== null) {
      const lineNum = this.getLineNumber(content, match.index);
      this.results.longNames.push({
        file: filePath,
        line: lineNum,
        name: match[1],
        length: match[1].length,
        suggestion: this.getSuggestion(match[1])
      });
    }
  }

  // 行番号を取得
  getLineNumber(content, index) {
    return content.substring(0, index).split('\n').length;
  }

  // 簡素化提案を取得
  getSuggestion(name) {
    // 直接マッチング
    if (SIMPLIFICATION_SUGGESTIONS[name]) {
      return SIMPLIFICATION_SUGGESTIONS[name];
    }
    
    // パターンマッチング
    for (const [pattern, suggestion] of Object.entries(SIMPLIFICATION_SUGGESTIONS)) {
      if (name.includes(pattern)) {
        return suggestion;
      }
    }

    // 自動簡素化
    return this.autoSimplify(name);
  }

  // チェーン簡素化提案
  getChainSuggestion(chain) {
    for (const [pattern, suggestion] of Object.entries(SIMPLIFICATION_SUGGESTIONS)) {
      if (chain.includes(pattern)) {
        return chain.replace(pattern, suggestion);
      }
    }
    return chain;
  }

  // 自動簡素化
  autoSimplify(name) {
    // 冗長な接頭辞・接尾辞を削除
    let simplified = name
      .replace(/^(get|set|create|delete|update|handle|process|manage)/, '')
      .replace(/(Data|Info|Config|Manager|Handler|Processor)$/, '')
      .replace(/([A-Z])[a-z]*And([A-Z])/g, '$1$2') // 'AndSomething' -> 略記
      .replace(/Current|Active|Available/g, '') // 冗長な形容詞削除
      .toLowerCase();

    // キャメルケースの最初を大文字に
    return simplified.charAt(0).toUpperCase() + simplified.slice(1);
  }

  // レポート生成
  generateReport() {
    const report = {
      summary: {
        totalFiles: this.files.length,
        verboseFunctions: this.results.verboseFunctions.length,
        duplicateVars: this.results.duplicateVars.length,
        complexChains: this.results.complexChains.length,
        longNames: this.results.longNames.length
      },
      details: this.results
    };

    console.log('\n📊 複雑性検出結果');
    console.log('================');
    console.log(`📁 対象ファイル: ${report.summary.totalFiles}件`);
    console.log(`🔤 冗長な関数名: ${report.summary.verboseFunctions}件`);
    console.log(`🔄 重複変数: ${report.summary.duplicateVars}件`);
    console.log(`⛓️  複雑なチェーン: ${report.summary.complexChains}件`);
    console.log(`📏 長すぎる名前: ${report.summary.longNames}件`);

    // 詳細出力
    if (report.summary.verboseFunctions > 0) {
      console.log('\n🔤 冗長な関数名:');
      for (const item of this.results.verboseFunctions.slice(0, 10)) {
        console.log(`  ${item.file}:${item.line} - ${item.name} → ${item.suggestion}`);
      }
    }

    if (report.summary.duplicateVars > 0) {
      console.log('\n🔄 重複変数パターン:');
      const grouped = this.groupByPattern(this.results.duplicateVars);
      for (const [pattern, items] of Object.entries(grouped)) {
        console.log(`  ${pattern} (${items.length}箇所)`);
      }
    }

    if (report.summary.complexChains > 0) {
      console.log('\n⛓️ 複雑なチェーン:');
      for (const item of this.results.complexChains.slice(0, 10)) {
        console.log(`  ${item.file}:${item.line} - ${item.chain} → ${item.suggestion}`);
      }
    }

    return report;
  }

  // パターンでグループ化
  groupByPattern(items) {
    const grouped = {};
    for (const item of items) {
      const pattern = item.pattern.replace(/const\s+\w+\s*=\s*/, '');
      if (!grouped[pattern]) grouped[pattern] = [];
      grouped[pattern].push(item);
    }
    return grouped;
  }

  // 修正用スクリプト生成
  generateFixScript() {
    const fixes = [];
    
    // 関数名修正
    for (const item of this.results.verboseFunctions) {
      fixes.push({
        type: 'rename_function',
        file: item.file,
        line: item.line,
        old: item.name,
        new: item.suggestion
      });
    }

    // チェーン簡素化
    for (const item of this.results.complexChains) {
      fixes.push({
        type: 'simplify_chain',
        file: item.file,
        line: item.line,
        old: item.chain,
        new: item.suggestion
      });
    }

    return fixes;
  }
}

// 実行
if (require.main === module) {
  const detector = new ComplexityDetector();
  const report = detector.detectComplexity();
  
  // 結果をJSON形式で保存
  fs.writeFileSync('complexity-report.json', JSON.stringify(report, null, 2));
  console.log('\n💾 詳細結果を complexity-report.json に保存しました');

  // 修正用スクリプト生成
  const fixes = detector.generateFixScript();
  fs.writeFileSync('auto-fixes.json', JSON.stringify(fixes, null, 2));
  console.log('🔧 修正用データを auto-fixes.json に保存しました');

  // 修正優先順位の表示
  console.log('\n📋 修正優先順位:');
  console.log('1. 重複変数の統一化 (即座に効果)');
  console.log('2. 複雑なチェーンの簡素化 (可読性向上)');
  console.log('3. 冗長な関数名の簡素化 (保守性向上)');
  console.log('4. 長すぎる名前の短縮 (タイピング効率)');
}

module.exports = ComplexityDetector;