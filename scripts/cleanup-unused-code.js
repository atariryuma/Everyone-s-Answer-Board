#!/usr/bin/env node

/**
 * @fileoverview メイン未使用コードクリーンアップオーケストレーター
 * 依存関係解析、レポート生成、安全削除を統合した包括的なクリーンアップシステム
 */

const fs = require('fs');
const path = require('path');
const GasDependencyAnalyzer = require('./dependency-analyzer');
const SafeDeleteSystem = require('./safe-delete');
const CleanupReporter = require('./cleanup-reporter');

class UnusedCodeCleanup {
  constructor(srcDir = './src') {
    this.srcDir = path.resolve(srcDir);
    this.analyzer = new GasDependencyAnalyzer(this.srcDir);
    this.safeDelete = new SafeDeleteSystem(this.srcDir);
    this.reporter = new CleanupReporter(this.srcDir);
    this.config = {
      dryRun: false,
      interactive: true,
      autoCleanup: false,
      riskLevel: 'low', // 'low', 'medium', 'high'
      backupRetentionDays: 7
    };
  }

  /**
   * 設定を更新
   */
  configure(options) {
    this.config = { ...this.config, ...options };
    this.safeDelete.setDryRun(this.config.dryRun);
  }

  /**
   * メインクリーンアップ処理を実行
   */
  async run() {
    console.log('🚀 未使用コードクリーンアップシステム開始');
    console.log(`📂 対象ディレクトリ: ${this.srcDir}`);
    console.log(`⚙️ 設定: ${JSON.stringify(this.config, null, 2)}`);

    try {
      // Phase 1: 依存関係解析
      const analysisResult = await this.performAnalysis();
      
      // Phase 2: レポート生成
      const reports = await this.generateReports();
      
      // Phase 3: ユーザー確認（インタラクティブモード）
      if (this.config.interactive && !this.config.dryRun) {
        const userApproval = await this.getUserApproval(analysisResult);
        if (!userApproval) {
          console.log('❌ ユーザーによりクリーンアップがキャンセルされました');
          return false;
        }
      }

      // Phase 4: バックアップ作成
      if (!this.config.dryRun) {
        await this.safeDelete.createBackup();
      }

      // Phase 5: 安全削除実行
      const deletionResult = await this.performDeletion(analysisResult);
      
      // Phase 6: 結果レポートとクリーンアップ
      await this.finalizeCleanup(deletionResult);
      
      console.log('✅ 未使用コードクリーンアップ完了');
      return true;
      
    } catch (error) {
      console.error('❌ クリーンアップエラー:', error.message);
      await this.handleError(error);
      throw error;
    }
  }

  /**
   * Phase 1: 依存関係解析を実行
   */
  async performAnalysis() {
    console.log('\n📊 Phase 1: 依存関係解析中...');
    
    const analysisResult = await this.analyzer.analyze();
    
    console.log('✅ 解析完了:');
    console.log(`   総ファイル数: ${analysisResult.summary.totalFiles}`);
    console.log(`   未使用ファイル: ${analysisResult.summary.unusedFiles}`);
    console.log(`   総関数数: ${analysisResult.summary.totalFunctions}`);
    console.log(`   未使用関数: ${analysisResult.summary.unusedFunctions}`);
    
    return analysisResult;
  }

  /**
   * Phase 2: 詳細レポートを生成
   */
  async generateReports() {
    console.log('\n📄 Phase 2: レポート生成中...');
    
    const reports = await this.reporter.generateReport();
    
    console.log('✅ レポート生成完了:');
    for (const [type, report] of Object.entries(reports)) {
      console.log(`   ${type}: ${path.relative(process.cwd(), report.path)}`);
    }
    
    return reports;
  }

  /**
   * Phase 3: ユーザー承認を取得（インタラクティブモード）
   */
  async getUserApproval(analysisResult) {
    console.log('\n❓ Phase 3: ユーザー確認');
    
    const unusedFiles = analysisResult.details.unusedFiles;
    const unusedFunctions = analysisResult.details.unusedFunctions;
    
    if (unusedFiles.length === 0 && unusedFunctions.length === 0) {
      console.log('✅ 削除対象なし。クリーンアップの必要はありません。');
      return false;
    }

    console.log('\n削除予定項目:');
    
    // リスクレベルに基づいてフィルタリング
    const filesToDelete = this.filterByRiskLevel(unusedFiles, 'file');
    const functionsToDelete = this.filterByRiskLevel(unusedFunctions, 'function');
    
    if (filesToDelete.length > 0) {
      console.log(`📁 ファイル (${filesToDelete.length}個):`);
      filesToDelete.slice(0, 10).forEach(file => {
        const risk = this.reporter.assessFileDeletionRisk(file);
        const riskEmoji = risk === 'low' ? '🟢' : risk === 'medium' ? '🟡' : '🔴';
        console.log(`   ${riskEmoji} ${file.path} (${file.size} bytes)`);
      });
      if (filesToDelete.length > 10) {
        console.log(`   ... および ${filesToDelete.length - 10} 個の追加ファイル`);
      }
    }

    if (functionsToDelete.length > 0) {
      console.log(`🔧 関数 (${functionsToDelete.length}個):`);
      functionsToDelete.slice(0, 10).forEach(func => {
        const risk = this.reporter.assessFunctionDeletionRisk(func);
        const riskEmoji = risk === 'low' ? '🟢' : risk === 'medium' ? '🟡' : '🔴';
        console.log(`   ${riskEmoji} ${func.name} (${func.definedIn.join(', ')})`);
      });
      if (functionsToDelete.length > 10) {
        console.log(`   ... および ${functionsToDelete.length - 10} 個の追加関数`);
      }
    }

    console.log(`\n実行予定: ${this.config.riskLevel}リスク以下の項目を削除`);
    console.log('⚠️ バックアップが自動作成され、ロールバック可能です');
    
    return await this.promptUserConfirmation('\n続行しますか？ (yes/no): ');
  }

  /**
   * リスクレベルに基づいてアイテムをフィルタリング
   */
  filterByRiskLevel(items, itemType) {
    const riskLevels = ['low', 'medium', 'high'];
    const maxRiskIndex = riskLevels.indexOf(this.config.riskLevel);
    
    return items.filter(item => {
      let risk;
      if (itemType === 'file') {
        risk = this.reporter.assessFileDeletionRisk(item);
      } else {
        risk = this.reporter.assessFunctionDeletionRisk(item);
      }
      
      const riskIndex = riskLevels.indexOf(risk);
      return riskIndex <= maxRiskIndex;
    });
  }

  /**
   * Phase 4: 安全削除を実行
   */
  async performDeletion(analysisResult) {
    console.log('\n🗑️ Phase 4: 安全削除実行中...');
    
    const unusedFiles = this.filterByRiskLevel(analysisResult.details.unusedFiles, 'file');
    const unusedFunctions = this.filterByRiskLevel(analysisResult.details.unusedFunctions, 'function');
    
    let filesDeleted = 0;
    let functionsDeleted = 0;

    // ファイル削除
    for (const file of unusedFiles) {
      const success = await this.safeDelete.deleteFile(file.path);
      if (success) filesDeleted++;
    }

    // 関数削除
    for (const func of unusedFunctions) {
      for (const filePath of func.definedIn) {
        const success = await this.safeDelete.deleteFunctionFromFile(filePath, func.name);
        if (success) {
          functionsDeleted++;
          break; // 一つのファイルから削除できれば十分
        }
      }
    }

    console.log(`✅ 削除完了: ${filesDeleted}ファイル, ${functionsDeleted}関数`);
    
    return {
      filesDeleted,
      functionsDeleted,
      totalFilesProcessed: unusedFiles.length,
      totalFunctionsProcessed: unusedFunctions.length
    };
  }

  /**
   * Phase 5: 最終化とクリーンアップ
   */
  async finalizeCleanup(deletionResult) {
    console.log('\n📋 Phase 5: 最終処理中...');
    
    // 削除レポートを生成・保存
    const deleteReport = this.safeDelete.generateDeleteReport();
    this.safeDelete.saveDeleteReport(deleteReport);
    
    // 古いバックアップをクリーンアップ
    this.safeDelete.cleanupOldBackups(this.config.backupRetentionDays);
    
    // 削除後の検証（テスト実行など）
    if (!this.config.dryRun && deletionResult.filesDeleted > 0) {
      await this.performPostDeletionValidation();
    }
    
    console.log('✅ 最終処理完了');
  }

  /**
   * 削除後の検証を実行
   */
  async performPostDeletionValidation() {
    console.log('🧪 削除後検証中...');
    
    try {
      // package.jsonが存在する場合、テストを実行
      if (fs.existsSync('./package.json')) {
        const { execSync } = require('child_process');
        console.log('npmテストを実行中...');
        execSync('npm test', { stdio: 'inherit' });
        console.log('✅ テスト通過');
      } else {
        console.log('ℹ️ package.jsonが見つかりません。手動でテストを実行してください。');
      }
      
      // 構文チェック（GASファイル）
      await this.validateGasSyntax();
      
    } catch (error) {
      console.warn('⚠️ 検証中にエラーが発生しました:', error.message);
      console.warn('手動での確認をお勧めします。');
    }
  }

  /**
   * GAS構文チェック
   */
  async validateGasSyntax() {
    console.log('🔍 GAS構文チェック中...');
    
    const gasFiles = fs.readdirSync(this.srcDir)
      .filter(file => file.endsWith('.gs'))
      .map(file => path.join(this.srcDir, file));
    
    for (const gasFile of gasFiles) {
      try {
        const content = fs.readFileSync(gasFile, 'utf-8');
        // 基本的な構文チェック（完全ではないが、明らかなエラーを検出）
        eval(`(function() { ${content} })`);
      } catch (error) {
        console.warn(`⚠️ 構文エラーの可能性: ${gasFile}`);
        console.warn(`   エラー: ${error.message}`);
      }
    }
    
    console.log('✅ 構文チェック完了');
  }

  /**
   * エラー処理
   */
  async handleError(error) {
    console.error('💥 クリーンアップ中にエラーが発生しました:');
    console.error(error.stack);
    
    if (!this.config.dryRun) {
      console.log('🔄 問題が発生した場合は、以下のコマンドでロールバックできます:');
      console.log(`node scripts/rollback.js "${this.safeDelete.currentBackupDir}"`);
    }
  }

  /**
   * ユーザー確認プロンプト
   */
  async promptUserConfirmation(message) {
    return new Promise((resolve) => {
      const readline = require('readline');
      const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout
      });

      rl.question(message, (answer) => {
        rl.close();
        resolve(answer.toLowerCase() === 'yes' || answer.toLowerCase() === 'y');
      });
    });
  }
}

// コマンドライン引数処理
function parseArguments() {
  const args = process.argv.slice(2);
  const config = {
    dryRun: false,
    interactive: true,
    riskLevel: 'low',
    autoCleanup: false
  };

  for (let i = 0; i < args.length; i++) {
    switch (args[i]) {
      case '--dry-run':
        config.dryRun = true;
        break;
      case '--non-interactive':
        config.interactive = false;
        break;
      case '--auto':
        config.autoCleanup = true;
        config.interactive = false;
        break;
      case '--risk-level':
        config.riskLevel = args[++i] || 'low';
        break;
      case '--help':
        showHelp();
        process.exit(0);
    }
  }

  return config;
}

function showHelp() {
  console.log(`
未使用コードクリーンアップシステム

使用法:
  node scripts/cleanup-unused-code.js [オプション]

オプション:
  --dry-run              実際の削除を行わず、処理をシミュレート
  --non-interactive      ユーザー確認なしで実行
  --auto                 完全自動実行（非推奨：慎重に使用）
  --risk-level LEVEL     削除リスクレベル (low|medium|high) デフォルト: low
  --help                 このヘルプを表示

例:
  node scripts/cleanup-unused-code.js --dry-run
  node scripts/cleanup-unused-code.js --risk-level medium
  node scripts/cleanup-unused-code.js --non-interactive --risk-level low

注意:
  - 実行前にGitコミットまたは手動バックアップを推奨
  - 最初は--dry-runオプションでテスト実行することを強く推奨
  - 削除後は必要に応じてテストを実行してください
`);
}

// メイン実行部
if (require.main === module) {
  async function main() {
    try {
      const config = parseArguments();
      
      console.log('🔧 未使用コードクリーンアップシステム v1.0');
      
      if (config.dryRun) {
        console.log('🔍 ドライランモードで実行中...');
      }
      
      const cleanup = new UnusedCodeCleanup('./src');
      cleanup.configure(config);
      
      const success = await cleanup.run();
      
      if (success) {
        console.log('\n🎉 クリーンアップが正常に完了しました！');
        if (!config.dryRun) {
          console.log('📝 生成されたレポートを確認して、結果を検証してください。');
          console.log('🧪 念のため、アプリケーションの動作テストを実行することをお勧めします。');
        }
      }
      
    } catch (error) {
      console.error('❌ 実行エラー:', error.message);
      process.exit(1);
    }
  }
  
  main();
}

module.exports = UnusedCodeCleanup;