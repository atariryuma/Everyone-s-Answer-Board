/**
 * @fileoverview 自動バックアップ機能 - データ損失リスク軽減
 */

/**
 * バックアップ設定
 */
const BACKUP_CONFIG = {
  RETENTION_DAYS: 30,          // バックアップ保持日数
  MAX_BACKUPS: 100,            // 最大バックアップ数
  BACKUP_INTERVAL_HOURS: 6,    // バックアップ間隔（時間）
  BACKUP_FOLDER_NAME: 'StudyQuest_Backups',
  ENABLE_AUTO_BACKUP: true
};

/**
 * バックアップメタデータ構造
 */
const BACKUP_METADATA = {
  timestamp: '',
  type: '',          // 'scheduled', 'manual', 'before_critical_operation'
  source: '',        // バックアップ元の識別子
  size: 0,
  checksum: '',
  description: '',
  userId: '',
  userEmail: ''
};

/**
 * メインバックアップ関数
 * @param {string} userId - ユーザーID
 * @param {string} type - バックアップタイプ
 * @param {string} description - バックアップの説明
 * @returns {object} バックアップ結果
 */
function createBackup(userId, type = 'manual', description = '') {
  if (!BACKUP_CONFIG.ENABLE_AUTO_BACKUP) {
    console.log('自動バックアップが無効化されています');
    return { success: false, message: 'バックアップが無効化されています' };
  }
  
  try {
    return safeExecute(() => {
      const userInfo = findUserById(userId);
      if (!userInfo || !userInfo.spreadsheetId) {
        throw new Error('バックアップ対象のスプレッドシートが見つかりません');
      }
      
      console.log(`🔄 バックアップ開始: ${type} - ${userInfo.adminEmail}`);
      
      // バックアップフォルダーを取得/作成
      const backupFolder = getOrCreateBackupFolder(userInfo.adminEmail);
      
      // 既存バックアップの確認
      if (type === 'scheduled' && !shouldCreateScheduledBackup(backupFolder)) {
        console.log('⏭️ スケジュールバックアップをスキップ（間隔未満）');
        return { success: false, message: 'バックアップ間隔未満のためスキップ' };
      }
      
      // スプレッドシートをコピー
      const backupFile = createSpreadsheetBackup(userInfo.spreadsheetId, backupFolder, type);
      
      // メタデータを作成
      const metadata = createBackupMetadata(backupFile, userInfo, type, description);
      
      // バックアップ履歴に記録
      recordBackupHistory(metadata);
      
      // 古いバックアップをクリーンアップ
      cleanupOldBackups(backupFolder);
      
      console.log(`✅ バックアップ完了: ${backupFile.getName()}`);
      
      return {
        success: true,
        message: 'バックアップが正常に作成されました',
        backupId: backupFile.getId(),
        backupName: backupFile.getName(),
        timestamp: metadata.timestamp
      };
      
    }, {
      function: 'createBackup',
      userId: userId,
      type: type,
      operation: 'backup_creation'
    });
    
  } catch (error) {
    console.error('バックアップ作成エラー:', error.message);
    return {
      success: false,
      message: error.message || 'バックアップの作成に失敗しました',
      error: error
    };
  }
}

/**
 * バックアップフォルダーを取得または作成
 * @param {string} userEmail - ユーザーメール
 * @returns {GoogleAppsScript.Drive.Folder} バックアップフォルダー
 */
function getOrCreateBackupFolder(userEmail) {
  try {
    const folderName = `${BACKUP_CONFIG.BACKUP_FOLDER_NAME}_${userEmail.replace(/[@.]/g, '_')}`;
    
    // 既存フォルダーを検索
    const folders = DriveApp.getFoldersByName(folderName);
    if (folders.hasNext()) {
      return folders.next();
    }
    
    // フォルダーが存在しない場合は作成
    const backupFolder = DriveApp.createFolder(folderName);
    console.log(`📁 バックアップフォルダー作成: ${folderName}`);
    
    return backupFolder;
    
  } catch (error) {
    throw new Error(`バックアップフォルダーの作成に失敗しました: ${error.message}`);
  }
}

/**
 * スケジュールバックアップが必要かチェック
 * @param {GoogleAppsScript.Drive.Folder} backupFolder - バックアップフォルダー
 * @returns {boolean} バックアップが必要かどうか
 */
function shouldCreateScheduledBackup(backupFolder) {
  try {
    const files = backupFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
    let lastBackupTime = 0;
    
    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      
      // ファイル名から日時を抽出（例: StudyQuest_Backup_20250714_120000.gsheet）
      const timeMatch = fileName.match(/(\d{8})_(\d{6})/);
      if (timeMatch) {
        const dateStr = timeMatch[1]; // YYYYMMDD
        const timeStr = timeMatch[2]; // HHMMSS
        
        const backupDate = new Date(
          parseInt(dateStr.substr(0, 4)),      // Year
          parseInt(dateStr.substr(4, 2)) - 1,  // Month (0-based)
          parseInt(dateStr.substr(6, 2)),      // Day
          parseInt(timeStr.substr(0, 2)),      // Hour
          parseInt(timeStr.substr(2, 2)),      // Minute
          parseInt(timeStr.substr(4, 2))       // Second
        );
        
        lastBackupTime = Math.max(lastBackupTime, backupDate.getTime());
      }
    }
    
    if (lastBackupTime === 0) {
      return true; // 初回バックアップ
    }
    
    const hoursSinceLastBackup = (Date.now() - lastBackupTime) / (1000 * 60 * 60);
    return hoursSinceLastBackup >= BACKUP_CONFIG.BACKUP_INTERVAL_HOURS;
    
  } catch (error) {
    console.error('バックアップ確認エラー:', error.message);
    return true; // エラー時はバックアップを作成
  }
}

/**
 * スプレッドシートバックアップを作成
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {GoogleAppsScript.Drive.Folder} backupFolder - バックアップ先フォルダー
 * @param {string} type - バックアップタイプ
 * @returns {GoogleAppsScript.Drive.File} バックアップファイル
 */
function createSpreadsheetBackup(spreadsheetId, backupFolder, type) {
  try {
    const sourceFile = DriveApp.getFileById(spreadsheetId);
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    const backupName = `StudyQuest_Backup_${type}_${timestamp}`;
    
    // スプレッドシートをコピー
    const backupFile = sourceFile.makeCopy(backupName, backupFolder);
    
    // 説明を追加
    backupFile.setDescription(`
      StudyQuest 自動バックアップ
      作成日時: ${new Date().toISOString()}
      バックアップタイプ: ${type}
      元ファイル: ${sourceFile.getName()}
    `.trim());
    
    return backupFile;
    
  } catch (error) {
    throw new Error(`スプレッドシートバックアップの作成に失敗しました: ${error.message}`);
  }
}

/**
 * バックアップメタデータを作成
 * @param {GoogleAppsScript.Drive.File} backupFile - バックアップファイル
 * @param {object} userInfo - ユーザー情報
 * @param {string} type - バックアップタイプ
 * @param {string} description - 説明
 * @returns {object} メタデータ
 */
function createBackupMetadata(backupFile, userInfo, type, description) {
  return {
    timestamp: new Date().toISOString(),
    type: type,
    source: userInfo.spreadsheetId,
    size: backupFile.getSize(),
    checksum: calculateFileChecksum(backupFile),
    description: description || `${type} backup for ${userInfo.adminEmail}`,
    userId: userInfo.userId,
    userEmail: userInfo.adminEmail,
    backupId: backupFile.getId(),
    backupName: backupFile.getName()
  };
}

/**
 * ファイルのチェックサムを計算（簡易版）
 * @param {GoogleAppsScript.Drive.File} file - ファイル
 * @returns {string} チェックサム
 */
function calculateFileChecksum(file) {
  try {
    // ファイルサイズと作成日時からシンプルなハッシュを生成
    const data = `${file.getSize()}_${file.getDateCreated().getTime()}_${file.getName()}`;
    const hash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, data, Utilities.Charset.UTF_8);
    return hash.map(byte => (byte + 256).toString(16).slice(-2)).join('').substr(0, 16);
  } catch (error) {
    console.warn('チェックサム計算エラー:', error.message);
    return 'unknown';
  }
}

/**
 * バックアップ履歴を記録
 * @param {object} metadata - バックアップメタデータ
 */
function recordBackupHistory(metadata) {
  try {
    // PropertiesService に履歴を保存（最新10件まで）
    const props = PropertiesService.getScriptProperties();
    const historyKey = `backup_history_${metadata.userId}`;
    
    let history = [];
    try {
      const existingHistory = props.getProperty(historyKey);
      if (existingHistory) {
        history = JSON.parse(existingHistory);
      }
    } catch (parseError) {
      console.warn('既存履歴の読み込みに失敗:', parseError.message);
    }
    
    // 新しいバックアップを履歴に追加
    history.unshift(metadata);
    
    // 最新10件のみ保持
    if (history.length > 10) {
      history = history.slice(0, 10);
    }
    
    props.setProperty(historyKey, JSON.stringify(history));
    console.log(`📝 バックアップ履歴記録: ${metadata.backupName}`);
    
  } catch (error) {
    console.error('バックアップ履歴記録エラー:', error.message);
    // 履歴記録の失敗はバックアップ全体の失敗とはしない
  }
}

/**
 * 古いバックアップをクリーンアップ
 * @param {GoogleAppsScript.Drive.Folder} backupFolder - バックアップフォルダー
 */
function cleanupOldBackups(backupFolder) {
  try {
    const files = backupFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
    const backupFiles = [];
    
    // バックアップファイルを収集
    while (files.hasNext()) {
      const file = files.next();
      backupFiles.push({
        file: file,
        created: file.getDateCreated().getTime()
      });
    }
    
    // 作成日時で降順ソート（新しい順）
    backupFiles.sort((a, b) => b.created - a.created);
    
    const cutoffTime = Date.now() - (BACKUP_CONFIG.RETENTION_DAYS * 24 * 60 * 60 * 1000);
    let deletedCount = 0;
    
    // 古いファイルまたは上限を超えたファイルを削除
    for (let i = 0; i < backupFiles.length; i++) {
      const shouldDelete = 
        i >= BACKUP_CONFIG.MAX_BACKUPS ||           // 上限を超過
        backupFiles[i].created < cutoffTime;        // 保持期間を超過
      
      if (shouldDelete) {
        try {
          backupFiles[i].file.setTrashed(true);
          deletedCount++;
          console.log(`🗑️ 古いバックアップを削除: ${backupFiles[i].file.getName()}`);
        } catch (deleteError) {
          console.error(`バックアップ削除エラー: ${deleteError.message}`);
        }
      }
    }
    
    if (deletedCount > 0) {
      console.log(`✅ ${deletedCount}個の古いバックアップを削除しました`);
    }
    
  } catch (error) {
    console.error('バックアップクリーンアップエラー:', error.message);
    // クリーンアップの失敗はバックアップ全体の失敗とはしない
  }
}

/**
 * ユーザーのバックアップ履歴を取得
 * @param {string} userId - ユーザーID
 * @returns {array} バックアップ履歴
 */
function getBackupHistory(userId) {
  try {
    const props = PropertiesService.getScriptProperties();
    const historyKey = `backup_history_${userId}`;
    const historyJson = props.getProperty(historyKey);
    
    if (!historyJson) {
      return [];
    }
    
    return JSON.parse(historyJson);
    
  } catch (error) {
    console.error('バックアップ履歴取得エラー:', error.message);
    return [];
  }
}

/**
 * 重要操作前の自動バックアップ
 * @param {string} userId - ユーザーID
 * @param {string} operation - 操作名
 * @returns {object} バックアップ結果
 */
function createPreOperationBackup(userId, operation) {
  console.log(`🛡️ 重要操作前バックアップ: ${operation}`);
  return createBackup(userId, 'before_critical_operation', `Before ${operation}`);
}

/**
 * スケジュール実行用のバックアップ関数
 * 全ユーザーの定期バックアップを実行
 */
function runScheduledBackups() {
  if (!BACKUP_CONFIG.ENABLE_AUTO_BACKUP) {
    console.log('自動バックアップが無効化されています');
    return;
  }
  
  try {
    console.log('🔄 スケジュールバックアップ開始');
    
    const allUsers = getAllUsers();
    let successCount = 0;
    let errorCount = 0;
    
    for (const user of allUsers) {
      if (user.spreadsheetId && user.isActive === 'true') {
        try {
          const result = createBackup(user.userId, 'scheduled');
          if (result.success) {
            successCount++;
          } else {
            errorCount++;
          }
        } catch (error) {
          console.error(`ユーザー ${user.adminEmail} のバックアップに失敗:`, error.message);
          errorCount++;
        }
      }
    }
    
    console.log(`✅ スケジュールバックアップ完了: 成功=${successCount}, エラー=${errorCount}`);
    
  } catch (error) {
    console.error('スケジュールバックアップエラー:', error.message);
  }
}

/**
 * バックアップからの復元（管理者用）
 * @param {string} backupFileId - バックアップファイルID
 * @param {string} targetUserId - 復元対象ユーザーID
 * @returns {object} 復元結果
 */
function restoreFromBackup(backupFileId, targetUserId) {
  if (!isDeployUser()) {
    throw new Error('この機能は管理者のみ利用できます');
  }
  
  try {
    return safeExecute(() => {
      const backupFile = DriveApp.getFileById(backupFileId);
      const userInfo = findUserById(targetUserId);
      
      if (!userInfo) {
        throw new Error('復元対象ユーザーが見つかりません');
      }
      
      // 現在のスプレッドシートをバックアップ
      createPreOperationBackup(targetUserId, 'restore_operation');
      
      // バックアップファイルをコピーして新しいスプレッドシートとして作成
      const restoredName = `${userInfo.adminEmail}_復元_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss')}`;
      const restoredFile = backupFile.makeCopy(restoredName);
      
      // ユーザー情報を更新
      updateUser(targetUserId, {
        spreadsheetId: restoredFile.getId(),
        spreadsheetUrl: `https://docs.google.com/spreadsheets/d/${restoredFile.getId()}/edit`,
        lastAccessedAt: new Date().toISOString()
      });
      
      console.log(`✅ バックアップから復元完了: ${restoredName}`);
      
      return {
        success: true,
        message: 'バックアップからの復元が完了しました',
        restoredSpreadsheetId: restoredFile.getId(),
        restoredSpreadsheetUrl: `https://docs.google.com/spreadsheets/d/${restoredFile.getId()}/edit`
      };
      
    }, {
      function: 'restoreFromBackup',
      backupFileId: backupFileId,
      targetUserId: targetUserId,
      operation: 'backup_restore'
    });
    
  } catch (error) {
    console.error('バックアップ復元エラー:', error.message);
    throw new Error(`バックアップからの復元に失敗しました: ${error.message}`);
  }
}

/**
 * バックアップ設定を更新
 * @param {object} newConfig - 新しい設定
 * @returns {object} 更新結果
 */
function updateBackupConfig(newConfig) {
  if (!isDeployUser()) {
    throw new Error('この機能は管理者のみ利用できます');
  }
  
  try {
    const props = PropertiesService.getScriptProperties();
    const currentConfigJson = props.getProperty('BACKUP_CONFIG') || '{}';
    const currentConfig = JSON.parse(currentConfigJson);
    
    // 設定をマージ
    const updatedConfig = Object.assign(currentConfig, newConfig);
    props.setProperty('BACKUP_CONFIG', JSON.stringify(updatedConfig));
    
    console.log('✅ バックアップ設定を更新しました:', updatedConfig);
    
    return {
      success: true,
      message: 'バックアップ設定が更新されました',
      config: updatedConfig
    };
    
  } catch (error) {
    console.error('バックアップ設定更新エラー:', error.message);
    throw new Error(`バックアップ設定の更新に失敗しました: ${error.message}`);
  }
}