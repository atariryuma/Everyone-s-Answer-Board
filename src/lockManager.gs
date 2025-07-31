/**
 * 統一ロック管理システム
 * 全フローでのロック制御を標準化し、デッドロックを防止
 */

// 統一ロックタイムアウト設定
const LOCK_TIMEOUTS = {
  READ_OPERATION: 5000,      // 読み取り専用操作: 5秒
  WRITE_OPERATION: 15000,    // 通常の書き込み操作: 15秒
  CRITICAL_OPERATION: 20000, // クリティカルな操作: 20秒
  BATCH_OPERATION: 30000     // バッチ処理: 30秒 (saveAndPublish等)
};

/**
 * 操作タイプに応じた適切なロックを取得
 * @param {string} operationType - 操作タイプ (READ_OPERATION, WRITE_OPERATION, CRITICAL_OPERATION, BATCH_OPERATION)
 * @param {string} [operationName] - デバッグ用操作名
 * @returns {Lock} Google Apps Script Lock オブジェクト
 */
function acquireStandardizedLock(operationType, operationName = 'unknown') {
  const timeout = LOCK_TIMEOUTS[operationType];
  if (!timeout) {
    throw new Error(`無効な操作タイプ: ${operationType}`);
  }
  
  const lock = LockService.getScriptLock();
  debugLog(`🔒 ロック取得試行: ${operationName} (${operationType}, ${timeout}ms)`);
  
  try {
    lock.waitLock(timeout);
    debugLog(`✅ ロック取得成功: ${operationName}`);
    return lock;
  } catch (error) {
    debugLog(`❌ ロック取得失敗: ${operationName} - ${error.message}`);
    throw new Error(`ロック取得がタイムアウトしました: ${operationName} (${timeout}ms)`);
  }
}

/**
 * ロックを安全に解放
 * @param {Lock} lock - 解放するロック
 * @param {string} [operationName] - デバッグ用操作名 
 */
function releaseStandardizedLock(lock, operationName = 'unknown') {
  try {
    lock.releaseLock();
    debugLog(`🔓 ロック解放成功: ${operationName}`);
  } catch (error) {
    debugLog(`⚠️ ロック解放エラー: ${operationName} - ${error.message}`);
  }
}

/**
 * 統一ロック管理でラップされた関数実行
 * @param {string} operationType - 操作タイプ
 * @param {string} operationName - 操作名
 * @param {Function} operation - 実行する関数
 * @returns {any} 操作の戻り値
 */
function executeWithStandardizedLock(operationType, operationName, operation) {
  const lock = acquireStandardizedLock(operationType, operationName);
  
  try {
    return operation();
  } finally {
    releaseStandardizedLock(lock, operationName);
  }
}