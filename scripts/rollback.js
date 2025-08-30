#!/usr/bin/env node

/**
 * @fileoverview ロールバックシステム
 * 安全削除システムのバックアップから完全復旧を行う
 */

const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

class RollbackSystem {
  constructor() {
    this.srcDir = path.resolve('./src');
    this.backupDir = path.resolve('./scripts/backups');
  }

  /**
   * 指定されたバックアップからロールバック実行
   */
  async rollback(backupPath) {
    const resolvedBackupPath = path.resolve(backupPath);
    
    // バックアップの存在確認
    if (!fs.existsSync(resolvedBackupPath)) {
      throw new Error(`バックアップディレクトリが見つかりません: ${resolvedBackupPath}`);
    }

    // メタデータの読み込み
    const metadataPath = path.join(resolvedBackupPath, 'backup-metadata.json');
    let metadata = null;
    if (fs.existsSync(metadataPath)) {
      metadata = JSON.parse(fs.readFileSync(metadataPath, 'utf-8'));
      console.log('📋 バックアップ情報:');
      console.log(`  作成日時: ${new Date(metadata.timestamp).toLocaleString()}`);
      console.log(`  元パス: ${metadata.originalPath}`);
      console.log(`  目的: ${metadata.purpose}`);
      if (metadata.gitCommit) {
        console.log(`  Gitコミット: ${metadata.gitCommit}`);
      }
    }

    // 現在の状態をバックアップ
    await this.createPreRollbackBackup();

    // ロールバック実行の確認
    console.log('\n⚠️ 以下の操作を実行します:');
    console.log(`  1. 現在の ${this.srcDir} を削除`);
    console.log(`  2. ${resolvedBackupPath}/src の内容を復元`);
    console.log('\n本当に実行しますか？ (yes/no):');
    
    // ユーザー確認（本番では自動実行も可能）
    const userConfirmation = await this.getUserConfirmation();
    if (userConfirmation !== 'yes') {
      console.log('❌ ロールバックがキャンセルされました');
      return false;
    }

    try {
      // 実際のロールバック実行
      await this.performRollback(resolvedBackupPath);
      
      console.log('✅ ロールバック完了');
      
      // ロールバック後の検証
      await this.verifyRollback(metadata);
      
      return true;
    } catch (error) {
      console.error('❌ ロールバックエラー:', error.message);
      throw error;
    }
  }

  /**
   * ロールバック前の現在状態をバックアップ
   */
  async createPreRollbackBackup() {
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const preRollbackDir = path.join(this.backupDir, `pre-rollback-${timestamp}`);
    
    console.log('💾 ロールバック前の現在状態をバックアップ中...');
    
    fs.mkdirSync(preRollbackDir, { recursive: true });
    
    try {
      execSync(`cp -R "${this.srcDir}" "${preRollbackDir}/"`, { stdio: 'pipe' });
      
      const metadata = {
        timestamp: new Date().toISOString(),
        originalPath: this.srcDir,
        backupPath: preRollbackDir,
        purpose: 'pre-rollback-backup',
        gitCommit: this.getCurrentGitCommit()
      };
      
      fs.writeFileSync(
        path.join(preRollbackDir, 'backup-metadata.json'),
        JSON.stringify(metadata, null, 2)
      );
      
      console.log(`✅ 現在状態のバックアップ完了: ${preRollbackDir}`);
    } catch (error) {
      throw new Error(`バックアップ作成エラー: ${error.message}`);
    }
  }

  /**
   * 実際のロールバック処理
   */
  async performRollback(backupPath) {
    const srcBackupPath = path.join(backupPath, 'src');
    
    if (!fs.existsSync(srcBackupPath)) {
      throw new Error(`バックアップ内にsrcディレクトリが見つかりません: ${srcBackupPath}`);
    }

    console.log('🔄 ロールバック実行中...');
    
    // 現在のsrcディレクトリを削除
    if (fs.existsSync(this.srcDir)) {
      execSync(`rm -rf "${this.srcDir}"`, { stdio: 'pipe' });
      console.log(`🗑️ 現在のディレクトリを削除: ${this.srcDir}`);
    }

    // バックアップから復元
    execSync(`cp -R "${srcBackupPath}" "${this.srcDir}"`, { stdio: 'pipe' });
    console.log(`📁 バックアップから復元: ${srcBackupPath} → ${this.srcDir}`);
  }

  /**
   * ロールバック後の検証
   */
  async verifyRollback(metadata) {
    console.log('🔍 ロールバック検証中...');
    
    const checks = [];
    
    // srcディレクトリの存在確認
    checks.push({
      name: 'srcディレクトリの存在',
      result: fs.existsSync(this.srcDir),
      critical: true
    });

    // 主要ファイルの存在確認
    const importantFiles = ['main.gs', 'appsscript.json'];
    for (const file of importantFiles) {
      const filePath = path.join(this.srcDir, file);
      checks.push({
        name: `${file}の存在`,
        result: fs.existsSync(filePath),
        critical: true
      });
    }

    // 検証結果の表示
    let allCriticalPassed = true;
    for (const check of checks) {
      const status = check.result ? '✅' : '❌';
      const level = check.critical ? '[重要]' : '[情報]';
      console.log(`  ${status} ${level} ${check.name}`);
      
      if (check.critical && !check.result) {
        allCriticalPassed = false;
      }
    }

    if (!allCriticalPassed) {
      console.warn('⚠️ 重要な検証項目が失敗しています。手動での確認をお勧めします。');
    } else {
      console.log('✅ ロールバック検証完了');
    }

    // Gitステータスの確認
    try {
      const gitStatus = execSync('git status --porcelain', { encoding: 'utf8' }).trim();
      if (gitStatus) {
        console.log('\n📝 Git変更検出:');
        console.log(gitStatus);
        console.log('必要に応じて git add . && git commit でコミットしてください');
      }
    } catch {
      console.log('ℹ️ Gitステータスの確認をスキップしました');
    }
  }

  /**
   * 現在のGitコミットハッシュを取得
   */
  getCurrentGitCommit() {
    try {
      return execSync('git rev-parse HEAD', { encoding: 'utf8' }).trim();
    } catch {
      return 'unknown';
    }
  }

  /**
   * ユーザー確認を取得（簡易実装）
   */
  async getUserConfirmation() {
    // Node.js環境での簡易的なユーザー入力
    // 実際の本番環境では適切な入力処理ライブラリを使用
    return new Promise((resolve) => {
      const readline = require('readline');
      const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout
      });

      rl.question('', (answer) => {
        rl.close();
        resolve(answer.toLowerCase());
      });
    });
  }

  /**
   * 利用可能なバックアップ一覧を表示
   */
  listAvailableBackups() {
    if (!fs.existsSync(this.backupDir)) {
      console.log('📂 バックアップディレクトリが存在しません');
      return [];
    }

    const backups = [];
    const backupDirs = fs.readdirSync(this.backupDir);
    
    console.log('📂 利用可能なバックアップ:');
    
    for (const backupName of backupDirs) {
      const backupPath = path.join(this.backupDir, backupName);
      if (fs.statSync(backupPath).isDirectory()) {
        const metadataPath = path.join(backupPath, 'backup-metadata.json');
        let metadata = { timestamp: 'unknown', purpose: 'unknown' };
        
        if (fs.existsSync(metadataPath)) {
          try {
            metadata = JSON.parse(fs.readFileSync(metadataPath, 'utf-8'));
          } catch {}
        }

        backups.push({
          name: backupName,
          path: backupPath,
          timestamp: metadata.timestamp,
          purpose: metadata.purpose
        });
        
        console.log(`  📁 ${backupName}`);
        console.log(`     日時: ${new Date(metadata.timestamp).toLocaleString()}`);
        console.log(`     目的: ${metadata.purpose}`);
        console.log(`     パス: ${backupPath}`);
        console.log();
      }
    }

    return backups;
  }
}

// コマンドライン実行時の処理
if (require.main === module) {
  async function main() {
    const args = process.argv.slice(2);
    const rollback = new RollbackSystem();

    try {
      if (args.length === 0 || args[0] === '--list') {
        // バックアップ一覧表示
        rollback.listAvailableBackups();
        console.log('使用法: node scripts/rollback.js <バックアップパス>');
        return;
      }

      const backupPath = args[0];
      console.log(`🔄 ロールバック開始: ${backupPath}`);
      
      const success = await rollback.rollback(backupPath);
      if (success) {
        console.log('\n🎉 ロールバック処理が正常に完了しました');
      }
      
    } catch (error) {
      console.error('❌ エラー:', error.message);
      process.exit(1);
    }
  }
  
  main();
}

module.exports = RollbackSystem;