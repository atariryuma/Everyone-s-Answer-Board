#!/usr/bin/env node
/**
 * 自動修正スクリプト - 効率的なリファクタリング実行
 * Usage: node scripts/auto-fix.js
 */

const fs = require('fs');
const path = require('path');

class AutoFixer {
  constructor() {
    this.fixes = JSON.parse(fs.readFileSync('auto-fixes.json', 'utf8'));
    this.applied = [];
    this.errors = [];
  }

  // 優先度1: 重複変数の統一化 - 共通関数の作成と適用
  fixDuplicateVariables() {
    console.log('\n🔧 Phase 1: 重複変数の統一化');
    
    // 重複パターンを特定
    const duplicatePatterns = [
      {
        pattern: /const\s+currentUserEmail\s*=\s*UserManager\.getCurrentEmail\(\)/g,
        replacement: 'const { currentUserEmail } = App.getConfig().getCurrentUserInfo()',
        description: 'getCurrentUserInfo()使用への統一'
      },
      {
        pattern: /const\s+spreadsheet\s*=\s*SpreadsheetApp\.openById\(([^)]+)\)/g,
        replacement: 'const spreadsheet = App.getConfig().getSpreadsheet($1)',
        description: 'getSpreadsheet()使用への統一'
      },
      {
        pattern: /const\s+userInfo\s*=\s*DB\.findUserByEmail\(([^)]+)\)/g,
        replacement: 'const userInfo = App.getConfig().getCurrentUserInfo().userInfo',
        description: 'getCurrentUserInfo()からuserInfo取得への統一'
      }
    ];

    for (const fix of duplicatePatterns) {
      this.applyPatternFix(fix);
    }
  }

  // 優先度2: 複雑なチェーンの簡素化
  fixComplexChains() {
    console.log('\n🔧 Phase 2: 複雑なチェーンの簡素化');
    
    // 定数名の簡素化
    this.replaceGlobal('SYSTEM_CONSTANTS', 'CONSTANTS');
    
    // 各ファイルでconstants.gsの参照を簡素化
    const files = this.getSourceFiles();
    for (const filePath of files) {
      try {
        let content = fs.readFileSync(filePath, 'utf8');
        let modified = false;

        // SYSTEM_CONSTANTS → CONSTANTS
        if (content.includes('SYSTEM_CONSTANTS')) {
          content = content.replace(/SYSTEM_CONSTANTS/g, 'CONSTANTS');
          modified = true;
        }

        if (modified) {
          fs.writeFileSync(filePath, content);
          console.log(`  ✅ ${filePath} - CONSTANTS名に簡素化`);
          this.applied.push({
            file: filePath,
            type: 'chain_simplification',
            description: 'SYSTEM_CONSTANTS → CONSTANTS'
          });
        }
      } catch (error) {
        this.errors.push({
          file: filePath,
          error: error.message,
          type: 'chain_fix_error'
        });
      }
    }
  }

  // 優先度3: 冗長な関数名の簡素化
  fixVerboseFunctionNames() {
    console.log('\n🔧 Phase 3: 冗長な関数名の簡素化');
    
    const functionRenames = [
      { old: 'getCurrentConfig', new: 'getConfig' },
      { old: 'getAppConfig', new: 'getConfig' },
      { old: 'getPublishedSheetData', new: 'getData' },
      { old: 'getIncrementalSheetData', new: 'getIncrementalData' },
      { old: 'addReactionBatch', new: 'addReactions' },
      { old: 'getAvailableSheets', new: 'getSheets' },
      { old: 'refreshBoardData', new: 'refreshData' },
      { old: 'clearActiveSheet', new: 'clearCache' },
    ];

    for (const rename of functionRenames) {
      this.renameFunctionGlobally(rename.old, rename.new);
    }
  }

  // 優先度4: 長すぎる変数名の短縮
  fixLongVariableNames() {
    console.log('\n🔧 Phase 4: 長すぎる変数名の短縮');
    
    const variableRenames = [
      { old: 'currentUserEmail', new: 'userEmail' },
      { old: 'requestUserId', new: 'userId' },
      { old: 'targetUserId', new: 'userId' },
      { old: 'spreadsheetId', new: 'sheetId' },
      { old: 'configManager', new: 'config' },
      { old: 'executionManager', new: 'execution' },
    ];

    for (const rename of variableRenames) {
      this.renameVariableInContext(rename.old, rename.new);
    }
  }

  // パターン修正の適用
  applyPatternFix(fix) {
    const files = this.getSourceFiles();
    let totalReplacements = 0;

    for (const filePath of files) {
      try {
        let content = fs.readFileSync(filePath, 'utf8');
        const originalContent = content;
        
        content = content.replace(fix.pattern, fix.replacement);
        
        if (content !== originalContent) {
          fs.writeFileSync(filePath, content);
          const replacements = (originalContent.match(fix.pattern) || []).length;
          totalReplacements += replacements;
          
          console.log(`  ✅ ${filePath} - ${replacements}箇所修正`);
          this.applied.push({
            file: filePath,
            type: 'pattern_fix',
            description: fix.description,
            replacements
          });
        }
      } catch (error) {
        this.errors.push({
          file: filePath,
          error: error.message,
          type: 'pattern_fix_error'
        });
      }
    }

    console.log(`  📊 ${fix.description}: ${totalReplacements}箇所修正完了`);
  }

  // グローバル置換
  replaceGlobal(oldText, newText) {
    // constants.gsで定数名変更
    const constantsFile = 'src/constants.gs';
    if (fs.existsSync(constantsFile)) {
      let content = fs.readFileSync(constantsFile, 'utf8');
      content = content.replace(new RegExp(`const\\s+${oldText}`, 'g'), `const ${newText}`);
      fs.writeFileSync(constantsFile, content);
      console.log(`  ✅ ${constantsFile} - ${oldText} → ${newText}`);
    }
  }

  // 関数名のグローバル変更
  renameFunctionGlobally(oldName, newName) {
    const files = this.getSourceFiles();
    let totalReplacements = 0;

    for (const filePath of files) {
      try {
        let content = fs.readFileSync(filePath, 'utf8');
        const originalContent = content;
        
        // 関数定義の変更
        content = content.replace(
          new RegExp(`function\\s+${oldName}\\s*\\(`, 'g'),
          `function ${newName}(`
        );
        
        // 関数呼び出しの変更
        content = content.replace(
          new RegExp(`${oldName}\\s*\\(`, 'g'),
          `${newName}(`
        );

        if (content !== originalContent) {
          fs.writeFileSync(filePath, content);
          totalReplacements++;
          console.log(`  ✅ ${filePath} - ${oldName} → ${newName}`);
        }
      } catch (error) {
        this.errors.push({
          file: filePath,
          error: error.message,
          type: 'function_rename_error'
        });
      }
    }

    if (totalReplacements > 0) {
      console.log(`  📊 ${oldName} → ${newName}: ${totalReplacements}ファイル修正完了`);
    }
  }

  // 文脈を考慮した変数名変更
  renameVariableInContext(oldName, newName) {
    const files = this.getSourceFiles();
    let totalReplacements = 0;

    for (const filePath of files) {
      try {
        let content = fs.readFileSync(filePath, 'utf8');
        const originalContent = content;
        
        // 関数パラメータでの変更
        content = content.replace(
          new RegExp(`\\(([^)]*?)\\b${oldName}\\b`, 'g'),
          `($1${newName}`
        );
        
        // ローカル変数宣言での変更
        content = content.replace(
          new RegExp(`const\\s+${oldName}\\s*=`, 'g'),
          `const ${newName} =`
        );
        content = content.replace(
          new RegExp(`let\\s+${oldName}\\s*=`, 'g'),
          `let ${newName} =`
        );

        if (content !== originalContent) {
          fs.writeFileSync(filePath, content);
          totalReplacements++;
          console.log(`  ✅ ${filePath} - ${oldName} → ${newName}`);
        }
      } catch (error) {
        this.errors.push({
          file: filePath,
          error: error.message,
          type: 'variable_rename_error'
        });
      }
    }

    if (totalReplacements > 0) {
      console.log(`  📊 ${oldName} → ${newName}: ${totalReplacements}ファイル修正完了`);
    }
  }

  // ソースファイル一覧取得
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

  // 修正実行
  execute() {
    console.log('🚀 自動修正開始 - 保守性向上のためのリファクタリング');
    
    this.fixDuplicateVariables();
    this.fixComplexChains();
    this.fixVerboseFunctionNames();
    this.fixLongVariableNames();
    
    this.generateSummary();
  }

  // サマリー生成
  generateSummary() {
    console.log('\n📊 自動修正完了サマリー');
    console.log('========================');
    console.log(`✅ 適用された修正: ${this.applied.length}件`);
    console.log(`❌ エラー: ${this.errors.length}件`);
    
    if (this.applied.length > 0) {
      console.log('\n✅ 適用された修正:');
      const groupedByType = this.groupByType(this.applied);
      for (const [type, items] of Object.entries(groupedByType)) {
        console.log(`  ${type}: ${items.length}件`);
      }
    }
    
    if (this.errors.length > 0) {
      console.log('\n❌ エラー一覧:');
      for (const error of this.errors.slice(0, 5)) {
        console.log(`  ${error.file}: ${error.error}`);
      }
    }

    // 結果保存
    const summary = {
      applied: this.applied,
      errors: this.errors,
      timestamp: new Date().toISOString()
    };
    
    fs.writeFileSync('refactor-summary.json', JSON.stringify(summary, null, 2));
    console.log('\n💾 修正サマリーを refactor-summary.json に保存しました');

    // 次のステップ提案
    console.log('\n📋 次の推奨アクション:');
    console.log('1. npm run test - 修正後のテスト実行');
    console.log('2. npm run lint - コード品質チェック');
    console.log('3. git diff - 変更内容の確認');
    console.log('4. git commit - 変更のコミット');
  }

  // タイプでグループ化
  groupByType(items) {
    const grouped = {};
    for (const item of items) {
      if (!grouped[item.type]) grouped[item.type] = [];
      grouped[item.type].push(item);
    }
    return grouped;
  }
}

// 実行
if (require.main === module) {
  const fixer = new AutoFixer();
  fixer.execute();
}

module.exports = AutoFixer;