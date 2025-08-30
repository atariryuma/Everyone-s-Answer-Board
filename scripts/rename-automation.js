#!/usr/bin/env node

/**
 * 自動リネーミングスクリプト - システム全体の命名統一
 * 1,520箇所のログ文＋関数名＋ファイル名を一括処理
 */

const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

class NamingUnifier {
  constructor() {
    this.srcDir = path.join(__dirname, '../src');
    this.backupDir = path.join(__dirname, '../scripts/naming-backup-' + Date.now());
    this.dryRun = process.argv.includes('--dry-run');
    this.verbose = process.argv.includes('--verbose');
    
    this.stats = {
      filesProcessed: 0,
      replacements: 0,
      filesRenamed: 0
    };

    // Phase 1: ログシステム統一
    this.logReplacements = [
      { from: /console\.log\(/g, to: 'Log.info(', description: 'console.log → Log.info' },
      { from: /console\.info\(/g, to: 'Log.info(', description: 'console.info → Log.info' },
      { from: /console\.warn\(/g, to: 'Log.warn(', description: 'console.warn → Log.warn' },
      { from: /console\.error\(/g, to: 'Log.error(', description: 'console.error → Log.error' },
      { from: /console\.debug\(/g, to: 'Log.debug(', description: 'console.debug → Log.debug' },
      { from: /ULog\.(debug|info|warn|error|critical)/g, to: 'Log.$1', description: 'ULog.* → Log.*' },
      { from: /debugLog\(/g, to: 'Log.debug(', description: 'debugLog → Log.debug' }
    ];

    // Phase 2: 関数名統一
    this.functionReplacements = [
      { from: /getCurrentUserEmail\(/g, to: 'User.email(', description: 'getCurrentUserEmail → User.email' },
      { from: /getCurrentUserInfo\(/g, to: 'User.info(', description: 'getCurrentUserInfo → User.info' },
      { from: /checkApplicationAccess\(/g, to: 'Access.check(', description: 'checkApplicationAccess → Access.check' },
      { from: /getDeployUserDomainInfo\(/g, to: 'Deploy.domain(', description: 'getDeployUserDomainInfo → Deploy.domain' },
      { from: /isDeployUser\(/g, to: 'Deploy.isUser(', description: 'isDeployUser → Deploy.isUser' },
      { from: /findUserById\(/g, to: 'DB.findUserById(', description: 'findUserById → DB.findUserById' },
      { from: /findUserByEmail\(/g, to: 'DB.findUserByEmail(', description: 'findUserByEmail → DB.findUserByEmail' },
      { from: /createUser\(/g, to: 'DB.createUser(', description: 'createUser → DB.createUser' }
    ];

    // Phase 3: クラス名統一
    this.classReplacements = [
      { from: /UnifiedErrorHandler/g, to: 'ErrorHandler', description: 'UnifiedErrorHandler → ErrorHandler' },
      { from: /UnifiedCacheManager/g, to: 'CacheManager', description: 'UnifiedCacheManager → CacheManager' },
      { from: /UnifiedCacheAPI/g, to: 'CacheAPI', description: 'UnifiedCacheAPI → CacheAPI' },
      { from: /UnifiedExecutionCache/g, to: 'ExecutionCache', description: 'UnifiedExecutionCache → ExecutionCache' }
    ];

    // Phase 4: ファイル名変更
    this.fileRenamings = [
      { from: 'unifiedCacheManager.gs', to: 'cache.gs' },
      { from: 'unifiedSecurityManager.gs', to: 'security.gs' },
      { from: 'unifiedUtilities.gs', to: 'utils.gs' },
      { from: 'errorHandler.gs', to: 'error.gs' },
      { from: 'ulog.gs', to: 'log.gs' },
      { from: 'logConfig.gs', to: 'DELETE' }  // 削除対象
    ];
  }

  /**
   * メイン実行
   */
  async run() {
    console.log('🚀 システム全体の命名統一を開始...');
    console.log(`📁 対象: ${this.srcDir}`);
    console.log(`${this.dryRun ? '🔍 DRY RUN モード' : '✍️ 実行モード'}`);
    console.log('');

    try {
      // バックアップ作成
      if (!this.dryRun) {
        this.createBackup();
      }

      // Phase 1: ログシステム統一
      await this.executePhase('ログシステム統一', this.logReplacements);
      
      // Phase 2: 関数名統一  
      await this.executePhase('関数名統一', this.functionReplacements);
      
      // Phase 3: クラス名統一
      await this.executePhase('クラス名統一', this.classReplacements);

      // Phase 4: コアログクラス作成
      await this.createLogClass();
      
      // Phase 5: ファイル名変更
      await this.renameFiles();

      // 結果表示
      this.showResults();
      
    } catch (error) {
      console.error('❌ エラーが発生しました:', error.message);
      console.log('💡 バックアップから復元: cp -r', this.backupDir + '/* ./src/');
      process.exit(1);
    }
  }

  /**
   * バックアップ作成
   */
  createBackup() {
    console.log('📦 バックアップ作成中...');
    if (!fs.existsSync(this.backupDir)) {
      fs.mkdirSync(this.backupDir, { recursive: true });
    }
    execSync(`cp -r ${this.srcDir}/* ${this.backupDir}/`);
    console.log(`✅ バックアップ完了: ${this.backupDir}`);
    console.log('');
  }

  /**
   * フェーズ実行
   */
  async executePhase(phaseName, replacements) {
    console.log(`📝 ${phaseName} (${replacements.length}パターン):`);
    
    const files = this.getTargetFiles();
    let phaseReplacements = 0;

    for (const file of files) {
      let content = fs.readFileSync(file, 'utf-8');
      let fileReplacements = 0;

      for (const replacement of replacements) {
        const before = content;
        content = content.replace(replacement.from, replacement.to);
        
        const matches = (before.match(replacement.from) || []).length;
        if (matches > 0) {
          fileReplacements += matches;
          if (this.verbose) {
            console.log(`   ${path.basename(file)}: ${matches}箇所 (${replacement.description})`);
          }
        }
      }

      if (fileReplacements > 0) {
        phaseReplacements += fileReplacements;
        if (!this.dryRun) {
          fs.writeFileSync(file, content, 'utf-8');
        }
      }
    }

    console.log(`   ✅ ${phaseReplacements}箇所を処理完了`);
    this.stats.replacements += phaseReplacements;
    console.log('');
  }

  /**
   * Logクラス作成
   */
  async createLogClass() {
    console.log('📝 統一Logクラス作成:');
    
    const logClassContent = `/**
 * 統一ログシステム - 軽量・高性能
 * 全てのログ出力を統一管理
 */
const Log = {
  debug: (msg, data = '') => console.log(\`[DEBUG] \${msg}\`, data),
  info: (msg, data = '') => console.info(\`[INFO] \${msg}\`, data), 
  warn: (msg, data = '') => console.warn(\`[WARN] \${msg}\`, data),
  error: (msg, data = '') => console.error(\`[ERROR] \${msg}\`, data)
};

// グローバル登録
if (typeof global !== 'undefined') global.Log = Log;
`;

    const logFilePath = path.join(this.srcDir, 'log.gs');
    
    if (!this.dryRun) {
      fs.writeFileSync(logFilePath, logClassContent, 'utf-8');
    }
    
    console.log(`   ✅ 統一Logクラス作成: ${path.basename(logFilePath)}`);
    console.log('');
  }

  /**
   * ファイル名変更
   */
  async renameFiles() {
    console.log(`📁 ファイル名変更 (${this.fileRenamings.length}ファイル):`);
    
    for (const renaming of this.fileRenamings) {
      const fromPath = path.join(this.srcDir, renaming.from);
      const toPath = path.join(this.srcDir, renaming.to);
      
      if (fs.existsSync(fromPath)) {
        if (renaming.to === 'DELETE') {
          if (!this.dryRun) {
            fs.unlinkSync(fromPath);
          }
          console.log(`   🗑️ 削除: ${renaming.from}`);
        } else {
          if (!this.dryRun) {
            fs.renameSync(fromPath, toPath);
          }
          console.log(`   📝 ${renaming.from} → ${renaming.to}`);
        }
        this.stats.filesRenamed++;
      }
    }
    
    console.log(`   ✅ ${this.stats.filesRenamed}ファイルの名前変更完了`);
    console.log('');
  }

  /**
   * 対象ファイル取得
   */
  getTargetFiles() {
    const files = [];
    const entries = fs.readdirSync(this.srcDir);
    
    for (const entry of entries) {
      const fullPath = path.join(this.srcDir, entry);
      if (fs.statSync(fullPath).isFile() && /\.(gs|html)$/.test(entry)) {
        files.push(fullPath);
      }
    }
    
    this.stats.filesProcessed = files.length;
    return files;
  }

  /**
   * 結果表示
   */
  showResults() {
    console.log('🎉 命名統一完了!');
    console.log('');
    console.log('📊 処理結果:');
    console.log(`   📁 処理ファイル数: ${this.stats.filesProcessed}`);
    console.log(`   🔄 置換箇所数: ${this.stats.replacements}`);
    console.log(`   📝 ファイル名変更: ${this.stats.filesRenamed}`);
    console.log('');
    
    if (this.dryRun) {
      console.log('💡 実際に実行するには: node scripts/rename-automation.js');
    } else {
      console.log('✅ 統一完了! 次のコマンドで確認:');
      console.log('   npm run lint');
      console.log('   npm run test');
      console.log('');
      console.log('🔄 復元が必要な場合:');
      console.log(`   cp -r ${this.backupDir}/* ./src/`);
    }
  }
}

// 実行
if (require.main === module) {
  const unifier = new NamingUnifier();
  unifier.run().catch(console.error);
}

module.exports = NamingUnifier;