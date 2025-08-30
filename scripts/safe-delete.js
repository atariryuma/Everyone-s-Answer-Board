#!/usr/bin/env node

/**
 * @fileoverview 安全削除システム
 * バックアップとロールバック機能付きの安全なファイル・関数削除システム
 */

const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

class SafeDeleteSystem {
  constructor(srcDir = './src') {
    this.srcDir = path.resolve(srcDir);
    this.backupDir = path.resolve('./scripts/backups');
    this.backupTimestamp = new Date().toISOString().replace(/[:.]/g, '-');
    this.currentBackupDir = path.join(this.backupDir, `backup-${this.backupTimestamp}`);
    this.deleteLog = [];
    this.dryRunMode = false;
  }

  /**
   * ドライランモードを設定
   */
  setDryRun(enabled) {
    this.dryRunMode = enabled;
    if (enabled) {
      console.log('🔍 ドライランモードが有効です（実際のファイル変更は行われません）');
    }
  }

  /**
   * 削除前バックアップの作成
   */
  async createBackup() {
    if (this.dryRunMode) {
      console.log(`🔍 [ドライラン] バックアップ作成: ${this.currentBackupDir}`);
      return;
    }

    console.log('💾 削除前バックアップを作成中...');
    
    // バックアップディレクトリの作成
    fs.mkdirSync(this.currentBackupDir, { recursive: true });
    
    try {
      // rsyncでファイルをバックアップ（シンボリックリンクや権限も保持）
      execSync(`cp -R "${this.srcDir}" "${this.currentBackupDir}/"`, { stdio: 'pipe' });
      
      // バックアップメタデータの記録
      const metadata = {
        timestamp: new Date().toISOString(),
        originalPath: this.srcDir,
        backupPath: this.currentBackupDir,
        gitCommit: this.getCurrentGitCommit(),
        nodeVersion: process.version,
        purpose: 'unused-code-cleanup'
      };
      
      fs.writeFileSync(
        path.join(this.currentBackupDir, 'backup-metadata.json'),
        JSON.stringify(metadata, null, 2)
      );
      
      console.log(`✅ バックアップ完了: ${this.currentBackupDir}`);
    } catch (error) {
      throw new Error(`バックアップ作成エラー: ${error.message}`);
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
   * ファイル全体を安全に削除
   */
  async deleteFile(filePath) {
    const fullPath = path.join(this.srcDir, filePath);
    const relativePath = path.relative(process.cwd(), fullPath);
    
    if (!fs.existsSync(fullPath)) {
      console.log(`⚠️ ファイルが存在しません: ${relativePath}`);
      return false;
    }

    const fileStats = fs.statSync(fullPath);
    const deleteInfo = {
      type: 'file',
      path: filePath,
      fullPath: fullPath,
      size: fileStats.size,
      timestamp: new Date().toISOString()
    };

    if (this.dryRunMode) {
      console.log(`🔍 [ドライラン] ファイル削除: ${relativePath} (${fileStats.size} bytes)`);
      this.deleteLog.push({ ...deleteInfo, dryRun: true });
      return true;
    }

    try {
      fs.unlinkSync(fullPath);
      console.log(`🗑️ ファイル削除: ${relativePath} (${fileStats.size} bytes)`);
      this.deleteLog.push(deleteInfo);
      return true;
    } catch (error) {
      console.error(`❌ ファイル削除エラー: ${relativePath} - ${error.message}`);
      return false;
    }
  }

  /**
   * ファイルから特定の関数を安全に削除
   */
  async deleteFunctionFromFile(filePath, functionName) {
    const fullPath = path.join(this.srcDir, filePath);
    const relativePath = path.relative(process.cwd(), fullPath);
    
    if (!fs.existsSync(fullPath)) {
      console.log(`⚠️ ファイルが存在しません: ${relativePath}`);
      return false;
    }

    try {
      const content = fs.readFileSync(fullPath, 'utf-8');
      const originalSize = content.length;
      
      // 関数削除を実行
      const { newContent, removed } = this.removeFunctionFromContent(content, functionName);
      
      if (!removed) {
        console.log(`⚠️ 関数が見つかりません: ${functionName} in ${relativePath}`);
        return false;
      }

      const deleteInfo = {
        type: 'function',
        path: filePath,
        functionName: functionName,
        originalSize: originalSize,
        newSize: newContent.length,
        timestamp: new Date().toISOString()
      };

      if (this.dryRunMode) {
        console.log(`🔍 [ドライラン] 関数削除: ${functionName} from ${relativePath}`);
        this.deleteLog.push({ ...deleteInfo, dryRun: true });
        return true;
      }

      // ファイルを更新
      fs.writeFileSync(fullPath, newContent, 'utf-8');
      console.log(`🔧 関数削除: ${functionName} from ${relativePath} (${originalSize - newContent.length} bytes削減)`);
      this.deleteLog.push(deleteInfo);
      return true;

    } catch (error) {
      console.error(`❌ 関数削除エラー: ${functionName} from ${relativePath} - ${error.message}`);
      return false;
    }
  }

  /**
   * 文字列から関数を削除
   */
  removeFunctionFromContent(content, functionName) {
    // 複数の関数定義パターンに対応
    const patterns = [
      // function functionName() { ... }
      new RegExp(`^\\s*function\\s+${functionName}\\s*\\([^)]*\\)\\s*\\{`, 'm'),
      // const functionName = function() { ... }
      new RegExp(`^\\s*const\\s+${functionName}\\s*=\\s*function\\s*\\([^)]*\\)\\s*\\{`, 'm'),
      // let functionName = function() { ... }
      new RegExp(`^\\s*let\\s+${functionName}\\s*=\\s*function\\s*\\([^)]*\\)\\s*\\{`, 'm'),
      // const functionName = () => { ... }
      new RegExp(`^\\s*const\\s+${functionName}\\s*=\\s*\\([^)]*\\)\\s*=>\\s*\\{`, 'm'),
      // methodName() { ... } (in class or object)
      new RegExp(`^\\s*${functionName}\\s*\\([^)]*\\)\\s*\\{`, 'm')
    ];

    for (const pattern of patterns) {
      const match = content.match(pattern);
      if (match) {
        const startIndex = match.index;
        const endIndex = this.findFunctionEnd(content, startIndex);
        
        if (endIndex > startIndex) {
          // 関数の前後の空行も含めて削除
          const beforeContent = content.slice(0, startIndex);
          const afterContent = content.slice(endIndex);
          
          // 前後の余分な空行を調整
          const cleanBeforeContent = beforeContent.replace(/\n\s*\n$/, '\n');
          const cleanAfterContent = afterContent.replace(/^\n\s*\n/, '\n');
          
          return {
            newContent: cleanBeforeContent + cleanAfterContent,
            removed: true
          };
        }
      }
    }

    return { newContent: content, removed: false };
  }

  /**
   * 関数の終了位置を見つける（波括弧のマッチング）
   */
  findFunctionEnd(content, startIndex) {
    let braceCount = 0;
    let inString = false;
    let stringChar = '';
    let inComment = false;
    let i = startIndex;

    // 最初の { を見つける
    while (i < content.length && content[i] !== '{') {
      i++;
    }
    
    if (i >= content.length) return startIndex;

    // 波括弧のマッチング開始
    braceCount = 1;
    i++;

    while (i < content.length && braceCount > 0) {
      const char = content[i];
      const prevChar = content[i - 1];
      const nextChar = content[i + 1];

      // コメントの処理
      if (!inString) {
        if (char === '/' && nextChar === '/') {
          inComment = 'line';
          i += 2;
          continue;
        }
        if (char === '/' && nextChar === '*') {
          inComment = 'block';
          i += 2;
          continue;
        }
        if (inComment === 'line' && char === '\n') {
          inComment = false;
        }
        if (inComment === 'block' && char === '*' && nextChar === '/') {
          inComment = false;
          i += 2;
          continue;
        }
        if (inComment) {
          i++;
          continue;
        }
      }

      // 文字列の処理
      if (!inComment) {
        if (!inString && (char === '"' || char === "'" || char === '`')) {
          inString = true;
          stringChar = char;
        } else if (inString && char === stringChar && prevChar !== '\\') {
          inString = false;
          stringChar = '';
        }
      }

      // 波括弧のカウント
      if (!inString && !inComment) {
        if (char === '{') {
          braceCount++;
        } else if (char === '}') {
          braceCount--;
        }
      }

      i++;
    }

    return i;
  }

  /**
   * 削除サマリーレポートを生成
   */
  generateDeleteReport() {
    const fileDeletes = this.deleteLog.filter(item => item.type === 'file');
    const functionDeletes = this.deleteLog.filter(item => item.type === 'function');
    
    const totalFilesDeleted = fileDeletes.length;
    const totalFunctionsDeleted = functionDeletes.length;
    const totalBytesFreed = fileDeletes.reduce((sum, item) => sum + (item.size || 0), 0) +
                           functionDeletes.reduce((sum, item) => sum + (item.originalSize - item.newSize), 0);

    const report = {
      summary: {
        totalFilesDeleted,
        totalFunctionsDeleted,
        totalBytesFreed,
        backupLocation: this.dryRunMode ? '(ドライランのためバックアップなし)' : this.currentBackupDir,
        timestamp: new Date().toISOString()
      },
      details: {
        deletedFiles: fileDeletes,
        deletedFunctions: functionDeletes
      },
      rollbackInstructions: this.dryRunMode ? null : {
        command: `node scripts/rollback.js "${this.currentBackupDir}"`,
        description: 'このコマンドで変更を完全にロールバックできます'
      }
    };

    return report;
  }

  /**
   * 削除レポートをファイルに保存
   */
  saveDeleteReport(report) {
    if (this.dryRunMode) {
      console.log('\n📊 削除レポート（ドライラン）:');
      console.log(`削除予定ファイル数: ${report.summary.totalFilesDeleted}`);
      console.log(`削除予定関数数: ${report.summary.totalFunctionsDeleted}`);
      console.log(`削減予定バイト数: ${report.summary.totalBytesFreed}`);
      return;
    }

    const reportPath = path.join(this.currentBackupDir, 'delete-report.json');
    fs.writeFileSync(reportPath, JSON.stringify(report, null, 2));
    
    console.log('\n📊 削除完了レポート:');
    console.log(`削除ファイル数: ${report.summary.totalFilesDeleted}`);
    console.log(`削除関数数: ${report.summary.totalFunctionsDeleted}`);
    console.log(`削減バイト数: ${report.summary.totalBytesFreed}`);
    console.log(`バックアップ場所: ${report.summary.backupLocation}`);
    console.log(`削除レポート: ${reportPath}`);
    
    if (report.rollbackInstructions) {
      console.log(`\n🔄 ロールバック方法: ${report.rollbackInstructions.command}`);
    }
  }

  /**
   * 古いバックアップを削除（保持期間を過ぎたもの）
   */
  cleanupOldBackups(retentionDays = 7) {
    if (this.dryRunMode) {
      console.log('🔍 [ドライラン] 古いバックアップの削除をスキップ');
      return;
    }

    if (!fs.existsSync(this.backupDir)) {
      return;
    }

    const cutoffTime = Date.now() - (retentionDays * 24 * 60 * 60 * 1000);
    const backupDirs = fs.readdirSync(this.backupDir);
    
    let deletedCount = 0;

    for (const backupName of backupDirs) {
      const backupPath = path.join(this.backupDir, backupName);
      if (fs.statSync(backupPath).isDirectory()) {
        const stats = fs.statSync(backupPath);
        if (stats.mtime.getTime() < cutoffTime) {
          try {
            execSync(`rm -rf "${backupPath}"`, { stdio: 'pipe' });
            deletedCount++;
            console.log(`🧹 古いバックアップを削除: ${backupName}`);
          } catch (error) {
            console.warn(`⚠️ バックアップ削除エラー: ${backupName} - ${error.message}`);
          }
        }
      }
    }

    if (deletedCount > 0) {
      console.log(`✅ ${deletedCount} 個の古いバックアップを削除しました`);
    }
  }
}

module.exports = SafeDeleteSystem;