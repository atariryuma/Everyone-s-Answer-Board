#!/usr/bin/env node

/**
 * 🔍 重複検証スクリプト - システム全体の重複チェック
 * 機械的に重複を検証し、潜在的な問題を発見
 */

const fs = require('fs');
const path = require('path');

// 検証対象パターン
const VALIDATION_PATTERNS = {
  // 関数/変数/定数の重複
  IDENTIFIERS: {
    // const/let/var 宣言
    CONST_DECLARATIONS: /^(?:\/\*[\s\S]*?\*\/\s*)?(?:\/\/.*?\n\s*)*(?:export\s+)?const\s+([A-Za-z_$][A-Za-z0-9_$]*)\s*=/gm,
    LET_DECLARATIONS: /^(?:\/\*[\s\S]*?\*\/\s*)?(?:\/\/.*?\n\s*)*(?:export\s+)?let\s+([A-Za-z_$][A-Za-z0-9_$]*)\s*=/gm,
    VAR_DECLARATIONS: /^(?:\/\*[\s\S]*?\*\/\s*)?(?:\/\/.*?\n\s*)*(?:export\s+)?var\s+([A-Za-z_$][A-Za-z0-9_$]*)\s*=/gm,
    
    // 関数宣言
    FUNCTION_DECLARATIONS: /^(?:\/\*[\s\S]*?\*\/\s*)?(?:\/\/.*?\n\s*)*(?:export\s+)?function\s+([A-Za-z_$][A-Za-z0-9_$]*)\s*\(/gm,
    
    // クラス宣言
    CLASS_DECLARATIONS: /^(?:\/\*[\s\S]*?\*\/\s*)?(?:\/\/.*?\n\s*)*(?:export\s+)?class\s+([A-Za-z_$][A-Za-z0-9_$]*)\s*(?:\s+extends\s+[A-Za-z_$][A-Za-z0-9_$]*)?\s*\{/gm,
    
    // オブジェクトメソッド
    OBJECT_METHODS: /([A-Za-z_$][A-Za-z0-9_$]*)\s*:\s*function\s*\(/g,
    ARROW_METHODS: /([A-Za-z_$][A-Za-z0-9_$]*)\s*:\s*\([^)]*\)\s*=>/g,
    
    // Object.freeze内のプロパティ
    FREEZE_PROPERTIES: /([A-Za-z_$][A-Za-z0-9_$]*)\s*:\s*(?:Object\.freeze\s*\()?[{\[\'\"]?/g
  },
  
  // 特定の重要な識別子
  CRITICAL_IDENTIFIERS: [
    'ErrorHandler', 'ErrorManager', 'DB', 'App', 'ConfigManager', 'SecurityValidator',
    'SYSTEM_CONSTANTS', 'DB_CONFIG', 'CORE', 'PROPS_KEYS', 'SECURITY_CONFIG'
  ]
};

// 除外パターン
const EXCLUDE_PATTERNS = {
  COMMENTS: /\/\*[\s\S]*?\*\/|\/\/.*$/gm,
  STRINGS: /["'`](?:[^"'`\\]|\\.)*["'`]/g,
  OBJECT_PROPERTIES: /\.([A-Za-z_$][A-Za-z0-9_$]*)/g
};

class DuplicateValidator {
  constructor() {
    this.results = {
      duplicates: new Map(),
      summary: {
        totalFiles: 0,
        duplicateCount: 0,
        criticalDuplicates: 0
      }
    };
  }

  /**
   * メインの重複検証実行
   */
  async validateDuplicates() {
    console.log('🔍 システム全体の重複検証を開始...');
    console.log('='.repeat(60));
    
    const srcDir = path.join(__dirname, '../src');
    const files = this.getGasFiles(srcDir);
    
    this.results.summary.totalFiles = files.length;
    console.log(`📁 検証対象ファイル数: ${files.length}`);
    
    // 全識別子を収集
    const allIdentifiers = new Map();
    
    for (const filePath of files) {
      const content = fs.readFileSync(filePath, 'utf8');
      const fileName = path.basename(filePath);
      
      // ファイル内の識別子を抽出
      const identifiers = this.extractIdentifiers(content, fileName);
      
      // グローバルマップに追加
      for (const [name, locations] of identifiers) {
        if (!allIdentifiers.has(name)) {
          allIdentifiers.set(name, []);
        }
        allIdentifiers.get(name).push(...locations);
      }
    }
    
    // 重複を検出
    this.detectDuplicates(allIdentifiers);
    
    // 結果出力
    this.outputResults();
    
    return this.results;
  }
  
  /**
   * GASファイルを取得
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
   * ファイルから識別子を抽出
   */
  extractIdentifiers(content, fileName) {
    const identifiers = new Map();
    
    // コメントと文字列を除去してクリーンなコードを作成
    const cleanContent = this.cleanContent(content);
    
    // 各パターンで識別子を抽出
    for (const [patternName, pattern] of Object.entries(VALIDATION_PATTERNS.IDENTIFIERS)) {
      let match;
      while ((match = pattern.exec(cleanContent)) !== null) {
        const name = match[1];
        const line = this.getLineNumber(content, match.index);
        
        if (!identifiers.has(name)) {
          identifiers.set(name, []);
        }
        
        identifiers.get(name).push({
          file: fileName,
          line: line,
          type: patternName,
          context: this.getContext(content, match.index)
        });
      }
    }
    
    return identifiers;
  }
  
  /**
   * コンテンツをクリーンアップ（コメント・文字列除去）
   */
  cleanContent(content) {
    return content
      .replace(EXCLUDE_PATTERNS.COMMENTS, '') // コメント除去
      .replace(EXCLUDE_PATTERNS.STRINGS, '""'); // 文字列を空文字列に置換
  }
  
  /**
   * 行番号を取得
   */
  getLineNumber(content, index) {
    return content.substring(0, index).split('\n').length;
  }
  
  /**
   * コンテキストを取得（前後の行）
   */
  getContext(content, index) {
    const lines = content.split('\n');
    const lineIndex = this.getLineNumber(content, index) - 1;
    const start = Math.max(0, lineIndex - 1);
    const end = Math.min(lines.length, lineIndex + 2);
    
    return lines.slice(start, end).join('\n');
  }
  
  /**
   * 重複を検出
   */
  detectDuplicates(allIdentifiers) {
    for (const [name, locations] of allIdentifiers) {
      // 同じ識別子が複数の場所で定義されている場合
      if (locations.length > 1) {
        // 同じファイル内の重複は除外（オーバーロードなど）
        const uniqueFiles = new Set(locations.map(loc => loc.file));
        if (uniqueFiles.size > 1) {
          this.results.duplicates.set(name, locations);
          this.results.summary.duplicateCount++;
          
          // クリティカルな識別子の重複をチェック
          if (VALIDATION_PATTERNS.CRITICAL_IDENTIFIERS.includes(name)) {
            this.results.summary.criticalDuplicates++;
          }
        }
      }
    }
  }
  
  /**
   * 結果を出力
   */
  outputResults() {
    console.log('\n📊 重複検証結果');
    console.log('='.repeat(60));
    console.log(`総ファイル数: ${this.results.summary.totalFiles}`);
    console.log(`重複識別子数: ${this.results.summary.duplicateCount}`);
    console.log(`クリティカル重複数: ${this.results.summary.criticalDuplicates}`);
    
    if (this.results.duplicates.size === 0) {
      console.log('\n✅ 重複は検出されませんでした！');
      return;
    }
    
    console.log('\n⚠️ 検出された重複:');
    console.log('-'.repeat(60));
    
    // 重要度順にソート
    const sortedDuplicates = Array.from(this.results.duplicates.entries()).sort(([nameA], [nameB]) => {
      const isCriticalA = VALIDATION_PATTERNS.CRITICAL_IDENTIFIERS.includes(nameA);
      const isCriticalB = VALIDATION_PATTERNS.CRITICAL_IDENTIFIERS.includes(nameB);
      
      if (isCriticalA && !isCriticalB) return -1;
      if (!isCriticalA && isCriticalB) return 1;
      return nameA.localeCompare(nameB);
    });
    
    for (const [name, locations] of sortedDuplicates) {
      const isCritical = VALIDATION_PATTERNS.CRITICAL_IDENTIFIERS.includes(name);
      const priority = isCritical ? '🔴 CRITICAL' : '🟡 WARNING';
      
      console.log(`\n${priority} 重複識別子: "${name}"`);
      
      for (const location of locations) {
        console.log(`  📁 ${location.file}:${location.line} (${location.type})`);
        console.log(`     ${this.getFirstLine(location.context)}`);
      }
      
      if (isCritical) {
        console.log('  🚨 この重複は即座に修正が必要です！');
      }
    }
    
    // 修正提案
    if (this.results.summary.criticalDuplicates > 0) {
      console.log('\n🔧 修正提案:');
      console.log('-'.repeat(40));
      console.log('1. クリティカルな重複を優先的に修正');
      console.log('2. 重複した識別子を統一または別名に変更');
      console.log('3. 名前空間を活用して競合を回避');
      console.log('4. 修正後にこのスクリプトを再実行して検証');
    }
  }
  
  /**
   * コンテキストの最初の行を取得
   */
  getFirstLine(context) {
    return context.split('\n')[0].trim();
  }
}

// メイン実行
if (require.main === module) {
  const validator = new DuplicateValidator();
  validator.validateDuplicates().catch(console.error);
}

module.exports = { DuplicateValidator };