#!/usr/bin/env node

/**
 * 🔍 構文競合検証スクリプト - GAS特有の構文エラーと競合を検証
 * 実際にGASで問題となる構文エラーや名前空間の競合を検出
 */

const fs = require('fs');
const path = require('path');

// GAS特有の検証パターン
const GAS_VALIDATION_PATTERNS = {
  // 同一名前空間内での重複宣言（実際のエラーの原因）
  NAMESPACE_CONFLICTS: {
    // const/let/var/function/classの同一スコープ内重複
    GLOBAL_SCOPE: /^(?:const|let|var|function|class)\s+([A-Za-z_$][A-Za-z0-9_$]*)/gm
  },
  
  // GASで禁止されているパターン
  FORBIDDEN_PATTERNS: {
    // ES6 modules (GASでは使用不可)
    IMPORT_EXPORT: /^(?:import|export)\s/gm,
    
    // 非同期関数（GASでは制限あり）
    ASYNC_AWAIT: /(?:async\s+function|await\s)/g,
    
    // 予約されたGAS関数名の重複
    GAS_RESERVED: /^(?:function\s+)?(?:doGet|doPost|onOpen|onEdit|onFormSubmit|onInstall)\s*(?:\(|=)/gm
  },
  
  // 潜在的な問題パターン
  POTENTIAL_ISSUES: {
    // varの使用（CLAUDE.md準拠に反する）
    VAR_USAGE: /\bvar\s+[A-Za-z_$]/g,
    
    // Object.freeze忘れ（設定オブジェクトで）
    UNFROZEN_CONFIG: /const\s+([A-Z_][A-Z0-9_]*(?:CONFIG|CONSTANTS?))\s*=\s*\{[^}]*\}(?!\s*\))/g,
    
    // console.log内のトークン漏洩の可能性
    TOKEN_EXPOSURE: /console\.(?:log|info|warn|error)[^;]*(?:token|key|secret|password|credential)/gi
  }
};

// GAS予約関数名
const GAS_RESERVED_FUNCTIONS = [
  'doGet', 'doPost', 'onOpen', 'onEdit', 'onFormSubmit', 'onInstall',
  'include' // HtmlServiceで使用される
];

class SyntaxConflictValidator {
  constructor() {
    this.results = {
      conflicts: [],
      warnings: [],
      summary: {
        totalFiles: 0,
        conflictCount: 0,
        warningCount: 0,
        severity: 'NONE'
      }
    };
    
    // ファイル別のグローバル識別子を追跡
    this.globalIdentifiers = new Map();
  }

  /**
   * メイン検証実行
   */
  async validateSyntaxConflicts() {
    console.log('🔍 GAS構文競合検証を開始...');
    console.log('='.repeat(60));
    
    const srcDir = path.join(__dirname, '../src');
    const files = this.getGasFiles(srcDir);
    
    this.results.summary.totalFiles = files.length;
    console.log(`📁 検証対象ファイル数: ${files.length}`);
    
    // 各ファイルを個別に検証
    for (const filePath of files) {
      const content = fs.readFileSync(filePath, 'utf8');
      const fileName = path.basename(filePath);
      
      await this.validateFile(content, fileName);
    }
    
    // 全体的な競合をチェック
    this.detectGlobalConflicts();
    
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
      if (entry.isFile() && entry.name.endsWith('.gs')) {
        files.push(path.join(dir, entry.name));
      }
    }
    
    return files.sort();
  }
  
  /**
   * 単一ファイルの検証
   */
  async validateFile(content, fileName) {
    console.log(`🔍 ${fileName} を検証中...`);
    
    // グローバル識別子を収集
    this.collectGlobalIdentifiers(content, fileName);
    
    // 禁止パターンをチェック
    this.checkForbiddenPatterns(content, fileName);
    
    // 潜在的問題をチェック
    this.checkPotentialIssues(content, fileName);
    
    // ファイル内重複をチェック
    this.checkFileInternalDuplicates(content, fileName);
  }
  
  /**
   * グローバル識別子を収集
   */
  collectGlobalIdentifiers(content, fileName) {
    const identifiers = new Set();
    const pattern = GAS_VALIDATION_PATTERNS.NAMESPACE_CONFLICTS.GLOBAL_SCOPE;
    let match;
    
    pattern.lastIndex = 0;
    while ((match = pattern.exec(content)) !== null) {
      const name = match[1];
      identifiers.add(name);
    }
    
    this.globalIdentifiers.set(fileName, identifiers);
  }
  
  /**
   * 禁止パターンをチェック
   */
  checkForbiddenPatterns(content, fileName) {
    // ES6 modules
    const importExport = content.match(GAS_VALIDATION_PATTERNS.FORBIDDEN_PATTERNS.IMPORT_EXPORT);
    if (importExport) {
      this.addConflict(fileName, 'ES6_MODULES', 
        'GASではES6 modulesは使用できません', importExport);
    }
    
    // 非同期関数
    const asyncAwait = content.match(GAS_VALIDATION_PATTERNS.FORBIDDEN_PATTERNS.ASYNC_AWAIT);
    if (asyncAwait) {
      this.addWarning(fileName, 'ASYNC_USAGE',
        '非同期関数の使用が検出されました。GASでは制限があります', asyncAwait);
    }
    
    // GAS予約関数の重複
    const reserved = content.match(GAS_VALIDATION_PATTERNS.FORBIDDEN_PATTERNS.GAS_RESERVED);
    if (reserved) {
      // 複数ファイルで同じGAS関数が定義されていないかチェック
      for (const match of reserved) {
        const funcName = match.match(/(?:function\s+)?(\w+)/)[1];
        if (GAS_RESERVED_FUNCTIONS.includes(funcName)) {
          this.addConflict(fileName, 'GAS_RESERVED_DUPLICATE',
            `GAS予約関数 "${funcName}" が複数箇所で定義されている可能性`, [match]);
        }
      }
    }
  }
  
  /**
   * 潜在的問題をチェック
   */
  checkPotentialIssues(content, fileName) {
    // varの使用
    const varUsage = content.match(GAS_VALIDATION_PATTERNS.POTENTIAL_ISSUES.VAR_USAGE);
    if (varUsage && varUsage.length > 0) {
      this.addWarning(fileName, 'VAR_USAGE',
        `CLAUDE.md規範違反: ${varUsage.length}箇所でvarが使用されています`, varUsage.slice(0, 5));
    }
    
    // Object.freeze忘れ
    const unfrozenConfig = content.match(GAS_VALIDATION_PATTERNS.POTENTIAL_ISSUES.UNFROZEN_CONFIG);
    if (unfrozenConfig) {
      this.addWarning(fileName, 'UNFROZEN_CONFIG',
        'Object.freeze()されていない設定オブジェクトがあります', unfrozenConfig);
    }
    
    // トークン漏洩の可能性
    const tokenExposure = content.match(GAS_VALIDATION_PATTERNS.POTENTIAL_ISSUES.TOKEN_EXPOSURE);
    if (tokenExposure) {
      this.addConflict(fileName, 'SECURITY_TOKEN_EXPOSURE',
        '🚨 セキュリティ: トークンやキーがログに出力されている可能性', tokenExposure);
    }
  }
  
  /**
   * ファイル内重複をチェック
   */
  checkFileInternalDuplicates(content, fileName) {
    const identifierCount = new Map();
    const pattern = GAS_VALIDATION_PATTERNS.NAMESPACE_CONFLICTS.GLOBAL_SCOPE;
    let match;
    
    pattern.lastIndex = 0;
    while ((match = pattern.exec(content)) !== null) {
      const name = match[1];
      identifierCount.set(name, (identifierCount.get(name) || 0) + 1);
    }
    
    for (const [name, count] of identifierCount) {
      if (count > 1) {
        this.addConflict(fileName, 'FILE_INTERNAL_DUPLICATE',
          `ファイル内で "${name}" が ${count}回宣言されています`, [name]);
      }
    }
  }
  
  /**
   * グローバル競合を検出
   */
  detectGlobalConflicts() {
    const allIdentifiers = new Map();
    
    // 全ファイルの識別子を統合
    for (const [fileName, identifiers] of this.globalIdentifiers) {
      for (const identifier of identifiers) {
        if (!allIdentifiers.has(identifier)) {
          allIdentifiers.set(identifier, []);
        }
        allIdentifiers.get(identifier).push(fileName);
      }
    }
    
    // 複数ファイルで定義されている識別子をチェック
    for (const [identifier, files] of allIdentifiers) {
      if (files.length > 1) {
        this.addConflict('GLOBAL', 'CROSS_FILE_DUPLICATE',
          `"${identifier}" が複数ファイル(${files.join(', ')})で定義されています`, files);
      }
    }
  }
  
  /**
   * 競合を追加
   */
  addConflict(fileName, type, message, evidence) {
    this.results.conflicts.push({
      fileName,
      type,
      message,
      evidence: Array.isArray(evidence) ? evidence : [evidence],
      severity: 'ERROR'
    });
    
    this.results.summary.conflictCount++;
    this.results.summary.severity = 'CRITICAL';
  }
  
  /**
   * 警告を追加
   */
  addWarning(fileName, type, message, evidence) {
    this.results.warnings.push({
      fileName,
      type,
      message,
      evidence: Array.isArray(evidence) ? evidence : [evidence],
      severity: 'WARNING'
    });
    
    this.results.summary.warningCount++;
    if (this.results.summary.severity === 'NONE') {
      this.results.summary.severity = 'WARNING';
    }
  }
  
  /**
   * 結果出力
   */
  outputResults() {
    console.log('\n📊 GAS構文競合検証結果');
    console.log('='.repeat(60));
    console.log(`総ファイル数: ${this.results.summary.totalFiles}`);
    console.log(`競合数: ${this.results.summary.conflictCount}`);
    console.log(`警告数: ${this.results.summary.warningCount}`);
    console.log(`重要度: ${this.results.summary.severity}`);
    
    if (this.results.conflicts.length === 0 && this.results.warnings.length === 0) {
      console.log('\n✅ 構文競合は検出されませんでした！');
      console.log('🎉 GAS構文は完全に適合しています。');
      return;
    }
    
    // 競合を出力
    if (this.results.conflicts.length > 0) {
      console.log('\n🔴 検出された競合 (即座に修正が必要):');
      console.log('-'.repeat(50));
      
      for (const conflict of this.results.conflicts) {
        console.log(`\n❌ [${conflict.fileName}] ${conflict.type}`);
        console.log(`   ${conflict.message}`);
        
        if (conflict.evidence && conflict.evidence.length > 0) {
          console.log('   証拠:');
          for (const evidence of conflict.evidence.slice(0, 3)) {
            if (typeof evidence === 'string') {
              console.log(`     "${evidence.substring(0, 60)}${evidence.length > 60 ? '...' : ''}"`);
            } else {
              console.log(`     ${evidence}`);
            }
          }
          if (conflict.evidence.length > 3) {
            console.log(`     ... 他 ${conflict.evidence.length - 3} 件`);
          }
        }
      }
    }
    
    // 警告を出力
    if (this.results.warnings.length > 0) {
      console.log('\n🟡 検出された警告 (改善推奨):');
      console.log('-'.repeat(50));
      
      for (const warning of this.results.warnings) {
        console.log(`\n⚠️  [${warning.fileName}] ${warning.type}`);
        console.log(`   ${warning.message}`);
        
        if (warning.evidence && warning.evidence.length > 0) {
          const evidenceCount = warning.evidence.length;
          if (evidenceCount <= 2) {
            for (const evidence of warning.evidence) {
              if (typeof evidence === 'string') {
                console.log(`     "${evidence.substring(0, 60)}${evidence.length > 60 ? '...' : ''}"`);
              } else {
                console.log(`     ${evidence}`);
              }
            }
          } else {
            console.log(`     ${evidenceCount} 箇所で検出`);
          }
        }
      }
    }
    
    // 修正提案
    this.outputRecommendations();
  }
  
  /**
   * 修正提案出力
   */
  outputRecommendations() {
    if (this.results.conflicts.length === 0 && this.results.warnings.length === 0) {
      return;
    }
    
    console.log('\n🔧 修正提案:');
    console.log('-'.repeat(40));
    
    if (this.results.conflicts.length > 0) {
      console.log('🔴 CRITICAL (即座に修正):');
      
      const duplicates = this.results.conflicts.filter(c => 
        c.type.includes('DUPLICATE'));
      if (duplicates.length > 0) {
        console.log('   1. 重複宣言を統一または削除');
        console.log('   2. 名前空間を使用して競合回避');
      }
      
      const security = this.results.conflicts.filter(c => 
        c.type.includes('SECURITY'));
      if (security.length > 0) {
        console.log('   3. セキュリティ問題を即座に修正');
      }
    }
    
    if (this.results.warnings.length > 0) {
      console.log('🟡 WARNING (改善推奨):');
      
      const varUsage = this.results.warnings.filter(w => 
        w.type === 'VAR_USAGE');
      if (varUsage.length > 0) {
        console.log('   1. varをconst/letに変更');
      }
      
      const unfrozen = this.results.warnings.filter(w => 
        w.type === 'UNFROZEN_CONFIG');
      if (unfrozen.length > 0) {
        console.log('   2. 設定オブジェクトにObject.freeze()を追加');
      }
    }
    
    console.log('\n📋 修正後の検証:');
    console.log('   - このスクリプトを再実行');
    console.log('   - clasp push でGASにデプロイして動作確認');
    console.log('   - npm run test でテスト実行');
  }
}

// メイン実行
if (require.main === module) {
  const validator = new SyntaxConflictValidator();
  validator.validateSyntaxConflicts().catch(console.error);
}

module.exports = { SyntaxConflictValidator };