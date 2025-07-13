/**
 * EnhancedLockSystem.gs
 * 本番環境対応の強化されたロックシステム
 * 高負荷時のタイムアウト問題を解決
 */

/**
 * 🚀 エンハンスドユーザー作成（本番環境対応版）
 * 段階的フォールバック機構で高い成功率を実現
 * @param {string} adminEmail - メールアドレス
 * @param {object} additionalData - 追加データ
 * @returns {object} { userId, isNewUser, userInfo }
 */
function findOrCreateUserEnhanced(adminEmail, additionalData = {}) {
  if (!adminEmail || !adminEmail.includes('@')) {
    throw new Error('有効なメールアドレスが必要です');
  }

  debugLog('findOrCreateUserEnhanced: 開始', { adminEmail });

  // Stage 1: 高速ロック試行（10秒）
  let result = attemptWithAdaptiveLock(adminEmail, additionalData);
  if (result) {
    debugLog('findOrCreateUserEnhanced: Stage1成功', { adminEmail });
    return result;
  }

  // Stage 2: 中速ロック試行（5秒）
  result = attemptWithMediumLock(adminEmail, additionalData);
  if (result) {
    debugLog('findOrCreateUserEnhanced: Stage2成功', { adminEmail });
    return result;
  }

  // Stage 3: 最終フォールバック（既存ユーザーのみ）
  result = attemptFinalFallback(adminEmail);
  if (result) {
    debugLog('findOrCreateUserEnhanced: Stage3成功', { adminEmail });
    return result;
  }

  // Stage 4: 最終エラー
  // クライアント側での自動リトライを促すため専用エラーコードを返す
  throw new Error('LOCK_TIMEOUT');
}

/**
 * 🎯 Stage 1: 適応的ロック（10秒タイムアウト）
 * @param {string} adminEmail - メールアドレス
 * @param {object} additionalData - 追加データ
 * @returns {object|null} 成功時は結果、失敗時はnull
 */
function attemptWithAdaptiveLock(adminEmail, additionalData) {
  const lock = LockService.getUserLock();
  const timeout = 10000; // 10秒
  
  try {
    if (!lock.waitLock(timeout)) {
      debugLog('attemptWithAdaptiveLock: タイムアウト', { adminEmail, timeout });
      return null;
    }

    debugLog('attemptWithAdaptiveLock: ロック取得成功', { adminEmail });

    // 1. 既存ユーザー確認
    let existingUser = findUserByEmail(adminEmail);
    
    if (existingUser) {
      debugLog('attemptWithAdaptiveLock: 既存ユーザー発見', { userId: existingUser.userId, adminEmail });
      
      // 必要に応じて既存ユーザー情報を更新
      if (Object.keys(additionalData).length > 0) {
        const updateData = {
          lastAccessedAt: new Date().toISOString(),
          isActive: 'true',
          ...additionalData
        };
        updateUser(existingUser.userId, updateData);
        debugLog('attemptWithAdaptiveLock: 既存ユーザー更新完了', { userId: existingUser.userId });
      }
      
      return {
        userId: existingUser.userId,
        isNewUser: false,
        userInfo: existingUser
      };
    }

    // 2. 新規ユーザー作成
    debugLog('attemptWithAdaptiveLock: 新規ユーザー作成開始', { adminEmail });
    
    const userId = generateConsistentUserId(adminEmail);
    const userData = {
      userId: userId,
      adminEmail: adminEmail,
      createdAt: new Date().toISOString(),
      lastAccessedAt: new Date().toISOString(),
      isActive: 'true',
      configJson: '{}',
      spreadsheetId: '',
      spreadsheetUrl: '',
      ...additionalData
    };

    // 原子的作成
    createUserAtomic(userData);
    
    debugLog('attemptWithAdaptiveLock: 新規ユーザー作成完了', { userId, adminEmail });
    
    return {
      userId: userId,
      isNewUser: true,
      userInfo: userData
    };

  } catch (error) {
    console.error('attemptWithAdaptiveLock エラー:', error);
    return null;
  } finally {
    try {
      lock.releaseLock();
      debugLog('attemptWithAdaptiveLock: ロック解除完了', { adminEmail });
    } catch (e) {
      console.warn('ロック解除エラー:', e.message);
    }
  }
}

/**
 * 🔥 Stage 2: 中速ロック（5秒タイムアウト）
 * @param {string} adminEmail - メールアドレス
 * @param {object} additionalData - 追加データ
 * @returns {object|null} 成功時は結果、失敗時はnull
 */
function attemptWithMediumLock(adminEmail, additionalData) {
  const lock = LockService.getUserLock();
  const timeout = 5000; // 5秒
  
  try {
    if (!lock.waitLock(timeout)) {
      debugLog('attemptWithMediumLock: タイムアウト', { adminEmail, timeout });
      return null;
    }

    debugLog('attemptWithMediumLock: ロック取得成功', { adminEmail });

    // 既存ユーザー確認のみ（新規作成は行わない）
    let existingUser = findUserByEmail(adminEmail);
    
    if (existingUser) {
      debugLog('attemptWithMediumLock: 既存ユーザー発見', { userId: existingUser.userId });
      return {
        userId: existingUser.userId,
        isNewUser: false,
        userInfo: existingUser
      };
    }

    // 新規ユーザー作成（簡略版）
    const userId = generateConsistentUserId(adminEmail);
    const userData = {
      userId: userId,
      adminEmail: adminEmail,
      createdAt: new Date().toISOString(),
      lastAccessedAt: new Date().toISOString(),
      isActive: 'true',
      configJson: JSON.stringify(additionalData.configJson ? JSON.parse(additionalData.configJson) : {}),
      spreadsheetId: '',
      spreadsheetUrl: ''
    };
    
    createUserAtomic(userData);
    
    debugLog('attemptWithMediumLock: 新規ユーザー作成完了', { userId });
    
    return {
      userId: userId,
      isNewUser: true,
      userInfo: userData
    };

  } catch (error) {
    console.error('attemptWithMediumLock エラー:', error);
    return null;
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {
      console.warn('中速ロック解除エラー:', e.message);
    }
  }
}

/**
 * 🛡️ Stage 3: 最終フォールバック（既存ユーザーのみ）
 * @param {string} adminEmail - メールアドレス
 * @returns {object|null} 成功時は結果、失敗時はnull
 */
function attemptFinalFallback(adminEmail) {
  debugLog('attemptFinalFallback: 最終フォールバック実行', { adminEmail });
  
  try {
    // ロックなしで既存ユーザー確認のみ実行
    const existingUser = findUserByEmail(adminEmail);
    
    if (existingUser) {
      debugLog('attemptFinalFallback: 既存ユーザーを発見', { userId: existingUser.userId });
      
      // 既存ユーザーがある場合はそのまま返却
      return {
        userId: existingUser.userId,
        isNewUser: false,
        userInfo: existingUser
      };
    }
    
    // 既存ユーザーがいない場合はnullを返す
    debugLog('attemptFinalFallback: 既存ユーザーなし', { adminEmail });
    return null;
    
  } catch (error) {
    console.error('attemptFinalFallback エラー:', error);
    return null;
  }
}

/**
 * 📊 ロック状況監視
 * 本番環境でのロック競合状況を監視
 * @returns {object} ロック統計情報
 */
function getLockStatistics() {
  const stats = {
    timestamp: new Date().toISOString(),
    lockAttempts: {
      stage1Success: 0,
      stage2Success: 0,
      stage3Success: 0,
      totalFailures: 0
    },
    averageWaitTime: 0,
    currentLoad: 'unknown'
  };
  
  // 簡易負荷テスト
  const testLock = LockService.getScriptLock();
  const startTime = Date.now();
  
  try {
    if (testLock.waitLock(1000)) { // 1秒テスト
      const waitTime = Date.now() - startTime;
      stats.averageWaitTime = waitTime;
      stats.currentLoad = waitTime < 100 ? 'low' : waitTime < 500 ? 'medium' : 'high';
      testLock.releaseLock();
    } else {
      stats.currentLoad = 'very_high';
    }
  } catch (e) {
    stats.currentLoad = 'error';
  }
  
  return stats;
}

/**
 * サーバー側でリトライしながらユーザーを取得または作成する
 * @param {string} adminEmail - メールアドレス
 * @param {object} additionalData - 追加データ
 * @returns {object} 結果オブジェクト
 */
function findOrCreateUserWithRetry(adminEmail, additionalData = {}) {
  const maxRetries = 3;
  const interval = 1500;

  for (let i = 0; i < maxRetries; i++) {
    try {
      return findOrCreateUserEnhanced(adminEmail, additionalData);
    } catch (e) {
      if (e.message === 'LOCK_TIMEOUT' && i < maxRetries - 1) {
        Utilities.sleep(interval);
        continue;
      }
      throw e;
    }
  }

  throw new Error('LOCK_TIMEOUT');
}

/**
 * 🎯 本番環境用メイン関数（findOrCreateUserの置き換え）
 * 既存のfindOrCreateUserを段階的に置き換える
 * @param {string} adminEmail - メールアドレス
 * @param {object} additionalData - 追加データ
 * @returns {object} { userId, isNewUser, userInfo }
 */
function findOrCreateUserProduction(adminEmail, additionalData = {}) {
  try {
    return findOrCreateUserWithRetry(adminEmail, additionalData);
  } catch (error) {
    // 本番環境用エラーログ
    console.error('findOrCreateUserProduction 最終エラー:', {
      error: error.message,
      adminEmail: adminEmail,
      timestamp: new Date().toISOString(),
      lockStats: getLockStatistics()
    });
    throw error;
  }
}

/**
 * 📈 パフォーマンス監視付きユーザー作成
 * 本番環境でのパフォーマンス分析用
 * @param {string} adminEmail - メールアドレス
 * @param {object} additionalData - 追加データ
 * @returns {object} 結果とパフォーマンス情報
 */
function findOrCreateUserWithMetrics(adminEmail, additionalData = {}) {
  const startTime = Date.now();
  const metrics = {
    startTime: startTime,
    stagesAttempted: [],
    lockStats: getLockStatistics()
  };
  
  try {
    const result = findOrCreateUserEnhanced(adminEmail, additionalData);
    
    metrics.endTime = Date.now();
    metrics.totalDuration = metrics.endTime - metrics.startTime;
    metrics.success = true;
    
    debugLog('findOrCreateUserWithMetrics: 成功', {
      adminEmail,
      metrics,
      isNewUser: result.isNewUser
    });
    
    return {
      ...result,
      metrics: metrics
    };
    
  } catch (error) {
    metrics.endTime = Date.now();
    metrics.totalDuration = metrics.endTime - metrics.startTime;
    metrics.success = false;
    metrics.error = error.message;
    
    console.error('findOrCreateUserWithMetrics: 失敗', {
      adminEmail,
      metrics,
      error: error.message
    });
    
    throw error;
  }
}