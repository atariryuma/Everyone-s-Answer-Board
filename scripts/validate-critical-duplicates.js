#!/usr/bin/env node

/**
 * 🔍 クリティカル重複検証スクリプト - 実際の問題のある重複のみを検出
 * 誤検知を除去し、本当に修正が必要な重複のみをレポート
 */

const fs = require('fs');
const path = require('path');

// 実際の問題となる重複パターン
const CRITICAL_PATTERNS = {
  // グローバルスコープでの関数/変数/定数宣言の重複
  GLOBAL_DECLARATIONS: {
    CONST: /^(?:(?:\/\*[\s\S]*?\*\/|\/\/.*\n)\s*)*const\s+([A-Za-z_$][A-Za-z0-9_$]*)\s*=/gm,
    LET: /^(?:(?:\/\*[\s\S]*?\*\/|\/\/.*\n)\s*)*let\s+([A-Za-z_$][A-Za-z0-9_$]*)\s*=/gm,
    VAR: /^(?:(?:\/\*[\s\S]*?\*\/|\/\/.*\n)\s*)*var\s+([A-Za-z_$][A-Za-z0-9_$]*)\s*=/gm,
    FUNCTION: /^(?:(?:\/\*[\s\S]*?\*\/|\/\/.*\n)\s*)*function\s+([A-Za-z_$][A-Za-z0-9_$]*)\s*\(/gm,
    CLASS: /^(?:(?:\/\*[\s\S]*?\*\/|\/\/.*\n)\s*)*class\s+([A-Za-z_$][A-Za-z0-9_$]*)\s*(?:\s+extends\s+[A-Za-z_$][A-Za-z0-9_$]*)?\s*\{/gm
  },

  // Object.freeze/Object.seal内の最上位プロパティ（ネストではない）
  TOP_LEVEL_OBJECT_PROPERTIES: /^(?:(?:\/\*[\s\S]*?\*\/|\/\/.*\n)\s*)*const\s+([A-Za-z_$][A-Za-z0-9_$]*)\s*=\s*Object\.(?:freeze|seal)\s*\(/gm
};

// 除外すべき誤検知パターン
const EXCLUDE_FALSE_POSITIVES = {
  // Object内のプロパティ（これらは重複ではない）
  OBJECT_PROPERTIES: /\.\s*([A-Za-z_$][A-Za-z0-9_$]*)/g,
  
  // 文字列内の識別子
  STRING_LITERALS: /["'`](?:[^"'`\\]|\\.)*["'`]/g,
  
  // コメント内の識別子
  COMMENTS: /\/\*[\s\S]*?\*\/|\/\/.*$/gm,
  
  // プロパティキー（オブジェクトリテラル内）
  PROPERTY_KEYS: /^\s*([A-Za-z_$][A-Za-z0-9_$]*)\s*:/gm,
  
  // メソッド呼び出し
  METHOD_CALLS: /\.([A-Za-z_$][A-Za-z0-9_$]*)\s*\(/g
};

// 確実に重複してはならないクリティカル識別子
const CRITICAL_GLOBAL_IDENTIFIERS = [
  'ErrorHandler', 'DB', 'App', 'ConfigManager', 'SecurityValidator',
  'SYSTEM_CONSTANTS', 'DB_CONFIG', 'CORE', 'PROPS_KEYS',
  'UserManager', 'Services', 'AccessController'
];

class CriticalDuplicateValidator {
  constructor() {
    this.results = {
      criticalDuplicates: new Map(),
      potentialIssues: [],
      summary: {
        totalFiles: 0,
        criticalCount: 0,
        severity: 'NONE'
      }
    };
  }

  /**
   * メイン検証実行
   */
  async validateCriticalDuplicates() {
    console.log('🔍 クリティカル重複検証を開始...');
    console.log('='.repeat(60));
    
    const srcDir = path.join(__dirname, '../src');
    const files = this.getGasFiles(srcDir);
    
    this.results.summary.totalFiles = files.length;
    console.log(`📁 検証対象ファイル数: ${files.length}`);
    
    // 各ファイルからクリティカル識別子を抽出
    const globalDeclarations = new Map();
    
    for (const filePath of files) {
      const content = fs.readFileSync(filePath, 'utf8');
      const fileName = path.basename(filePath);
      
      // クリーンなコンテンツを作成（誤検知除去）
      const cleanContent = this.cleanContent(content);
      
      // グローバル宣言を抽出
      const declarations = this.extractGlobalDeclarations(cleanContent, fileName, content);
      
      // グローバルマップに追加
      for (const [name, locations] of declarations) {
        if (!globalDeclarations.has(name)) {
          globalDeclarations.set(name, []);
        }
        globalDeclarations.get(name).push(...locations);
      }
    }
    
    // クリティカル重複を検出
    this.detectCriticalDuplicates(globalDeclarations);
    
    // 結果出力
    this.outputResults();
    
    return this.results;
  }
  
  /**
   * GASファイル取得
   */
  getGasFiles(dir) {
    const files = [];
    const entries = fs.readdirSync(dir, { withFileTypes: true });
    
    for (const entry of entries) {
      if (entry.isFile() && (entry.name.endsWith('.gs') || entry.name.endsWith('.js'))) {
        files.push(path.join(dir, entry.name));
      }
    }
    
    return files.sort();
  }
  
  /**
   * コンテンツをクリーンアップ
   */
  cleanContent(content) {
    let cleaned = content;
    
    // コメントを除去
    cleaned = cleaned.replace(EXCLUDE_FALSE_POSITIVES.COMMENTS, '');
    
    // 文字列リテラルを除去
    cleaned = cleaned.replace(EXCLUDE_FALSE_POSITIVES.STRING_LITERALS, '""');
    
    return cleaned;
  }
  
  /**
   * グローバル宣言を抽出
   */
  extractGlobalDeclarations(cleanContent, fileName, originalContent) {
    const declarations = new Map();
    
    // 各宣言パターンをチェック
    for (const [type, pattern] of Object.entries(CRITICAL_PATTERNS.GLOBAL_DECLARATIONS)) {
      let match;
      pattern.lastIndex = 0; // Reset regex
      
      while ((match = pattern.exec(cleanContent)) !== null) {
        const name = match[1];
        const line = this.getLineNumber(originalContent, this.findOriginalIndex(originalContent, match[0], match.index));
        
        // 重要な識別子か、または複数ファイルにまたがる可能性があるもののみ記録
        if (this.shouldTrack(name)) {
          if (!declarations.has(name)) {
            declarations.set(name, []);
          }
          
          declarations.get(name).push({
            file: fileName,
            line: line,
            type: type,
            context: this.getContext(originalContent, this.findOriginalIndex(originalContent, match[0], match.index))
          });
        }
      }
    }
    
    return declarations;
  }
  
  /**
   * 追跡すべき識別子かを判定
   */
  shouldTrack(name) {
    // クリティカル識別子は必ず追跡
    if (CRITICAL_GLOBAL_IDENTIFIERS.includes(name)) {
      return true;
    }
    
    // 大文字で始まる識別子（クラス、定数など）
    if (/^[A-Z]/.test(name)) {
      return true;
    }
    
    // よく使われるグローバル関数名
    const commonGlobalFunctions = [
      'doGet', 'doPost', 'onOpen', 'onEdit', 'onFormSubmit',
      'generateUserId', 'createResponse', 'validateInput'
    ];
    
    return commonGlobalFunctions.includes(name);
  }
  
  /**
   * オリジナルコンテンツ内のインデックスを見つける
   */
  findOriginalIndex(originalContent, matchedText, cleanIndex) {
    // 簡易実装：マッチしたテキストの最初の出現箇所を探す
    return originalContent.indexOf(matchedText);
  }
  
  /**
   * 行番号取得
   */
  getLineNumber(content, index) {
    if (index < 0) return 1;
    return content.substring(0, index).split('\n').length;
  }
  
  /**
   * コンテキスト取得
   */
  getContext(content, index) {
    const lines = content.split('\n');
    const lineIndex = this.getLineNumber(content, index) - 1;
    const start = Math.max(0, lineIndex - 1);
    const end = Math.min(lines.length, lineIndex + 2);
    
    return lines.slice(start, end).join('\n');
  }
  
  /**
   * クリティカル重複を検出
   */
  detectCriticalDuplicates(globalDeclarations) {
    for (const [name, locations] of globalDeclarations) {
      if (locations.length > 1) {
        // 異なるファイル間での重複をチェック
        const uniqueFiles = [...new Set(locations.map(loc => loc.file))];
        
        if (uniqueFiles.length > 1) {
          // クリティカルな重複を検出
          const isCritical = CRITICAL_GLOBAL_IDENTIFIERS.includes(name);
          const severity = isCritical ? 'CRITICAL' : 'WARNING';
          
          this.results.criticalDuplicates.set(name, {
            locations,
            severity,
            isCritical,
            affectedFiles: uniqueFiles
          });
          
          this.results.summary.criticalCount++;
          
          if (isCritical && this.results.summary.severity !== 'CRITICAL') {
            this.results.summary.severity = 'CRITICAL';
          } else if (severity === 'WARNING' && this.results.summary.severity === 'NONE') {
            this.results.summary.severity = 'WARNING';
          }
        }
      }
    }
  }
  
  /**
   * 結果出力
   */
  outputResults() {
    console.log('\n📊 クリティカル重複検証結果');
    console.log('='.repeat(60));
    console.log(`総ファイル数: ${this.results.summary.totalFiles}`);
    console.log(`クリティカル重複数: ${this.results.summary.criticalCount}`);
    console.log(`重要度: ${this.results.summary.severity}`);
    
    if (this.results.criticalDuplicates.size === 0) {
      console.log('\n✅ クリティカルな重複は検出されませんでした！');
      console.log('🎉 システムの重複問題は解決済みです。');
      return;
    }
    
    console.log('\n🚨 検出されたクリティカル重複:');
    console.log('-'.repeat(60));
    
    // 重要度順にソート
    const sortedDuplicates = Array.from(this.results.criticalDuplicates.entries())
      .sort(([, a], [, b]) => {
        if (a.isCritical && !b.isCritical) return -1;
        if (!a.isCritical && b.isCritical) return 1;
        return 0;
      });
    
    for (const [name, info] of sortedDuplicates) {
      const icon = info.isCritical ? '🔴 CRITICAL' : '🟡 WARNING';
      
      console.log(`\n${icon} "${name}" が ${info.affectedFiles.length} ファイルで重複`);
      console.log(`   影響ファイル: ${info.affectedFiles.join(', ')}`);
      
      for (const location of info.locations) {
        console.log(`   📁 ${location.file}:${location.line} (${location.type})`);
        const contextLine = location.context.split('\n')[0].trim();
        if (contextLine) {
          console.log(`      ${contextLine.substring(0, 80)}${contextLine.length > 80 ? '...' : ''}`);
        }
      }
      
      if (info.isCritical) {
        console.log('   🚨 この重複は即座に修正が必要です！');
      }
    }
    
    // 修正提案
    if (this.results.summary.criticalCount > 0) {
      console.log('\n🔧 修正提案:');
      console.log('-'.repeat(40));
      
      const criticalCount = Array.from(this.results.criticalDuplicates.values())
        .filter(info => info.isCritical).length;
      
      if (criticalCount > 0) {
        console.log(`1. 🔴 CRITICAL: ${criticalCount}個の重複を最優先で修正`);
        console.log('   - 重複した識別子を統一または削除');
        console.log('   - 名前空間を活用して競合回避');
      }
      
      const warningCount = this.results.summary.criticalCount - criticalCount;
      if (warningCount > 0) {
        console.log(`2. 🟡 WARNING: ${warningCount}個の潜在的問題を確認`);
        console.log('   - 意図的な重複かを確認');
        console.log('   - 必要に応じてリファクタリング');
      }
      
      console.log('3. 修正後にこのスクリプトを再実行して検証');
    }
  }
}

// メイン実行
if (require.main === module) {
  const validator = new CriticalDuplicateValidator();
  validator.validateCriticalDuplicates().catch(console.error);
}

module.exports = { CriticalDuplicateValidator };