/**
 * メール特化ロックシステム
 * 同一ユーザーの重複操作を防ぎ、他ユーザーとの競合のみロック対象とする
 */

// グローバルメール処理状態管理
const PROCESSING_EMAILS = {};

/**
 * メール特化ロック：同一メールの重複処理を防止
 * @param {string} adminEmail - メールアドレス
 * @returns {boolean} ロック取得成功かどうか
 */
function acquireEmailLock(adminEmail) {
  const lockKey = 'processing_' + adminEmail;
  
  // 既に同じメールで処理中の場合は拒否
  if (PROCESSING_EMAILS[lockKey]) {
    console.log('Email-specific lock: Already processing for', adminEmail);
    return false;
  }
  
  // ロック取得
  PROCESSING_EMAILS[lockKey] = {
    timestamp: Date.now()
  };
  
  console.log('Email-specific lock: Acquired for', adminEmail);
  return true;
}

/**
 * メール特化ロック解放
 * @param {string} adminEmail - メールアドレス
 */
function releaseEmailLock(adminEmail) {
  const lockKey = 'processing_' + adminEmail;
  delete PROCESSING_EMAILS[lockKey];
  console.log('Email-specific lock: Released for', adminEmail);
}

/**
 * メール特化ロック付きユーザー作成
 * @param {string} adminEmail - メールアドレス
 * @param {object} additionalData - 追加データ
 * @returns {object} 結果
 */
function findOrCreateUserWithEmailLock(adminEmail, additionalData = {}) {
  // Step 1: メール特化ロック取得
  if (!acquireEmailLock(adminEmail)) {
    throw new Error('EMAIL_ALREADY_PROCESSING');
  }
  
  try {
    // Step 2: 既存ユーザー確認（ロックなし）
    let existingUser = findUserByEmailNonBlocking(adminEmail);
    console.log('findOrCreateUserWithEmailLock: 既存ユーザー確認結果', { existingUser: !!existingUser, adminEmail });
    
    if (existingUser) {
      console.log('findOrCreateUserWithEmailLock: 既存ユーザー発見', { userId: existingUser.userId, adminEmail });
      
      // 必要に応じて既存ユーザー情報を更新
      if (Object.keys(additionalData).length > 0) {
        const updateData = {
          lastAccessedAt: new Date().toISOString(),
          isActive: 'true',
          ...additionalData
        };
        updateUser(existingUser.userId, updateData);
      }
      
      // 古いロック状態をクリーンアップ
      cleanupEmailLocks();
      
      return {
        userId: existingUser.userId,
        isNewUser: false,
        userInfo: existingUser
      };
    }
    
    // Step 3: 新規ユーザー作成（スクリプトロック使用 - ユーザーロック廃止）
    console.log('findOrCreateUserWithEmailLock: 新規ユーザー作成開始', { adminEmail });
    const lock = LockService.getScriptLock();
    const timeout = 10000; // 25秒→10秒に短縮
    
    if (!lock.waitLock(timeout)) {
      console.error('findOrCreateUserWithEmailLock: スクリプトロックタイムアウト', { adminEmail, timeout });
      throw new Error('SCRIPT_LOCK_TIMEOUT');
    }
    
    try {
      // ダブルチェック: ロック取得後に再度確認
      existingUser = findUserByEmailNonBlocking(adminEmail);
      if (existingUser) {
        return {
          userId: existingUser.userId,
          isNewUser: false,
          userInfo: existingUser
        };
      }
      
      // 新規ユーザー作成実行（既存のcreateUserAtomicを使用）
      const userId = generateUniqueUserId();
      const userData = {
        userId: userId,
        adminEmail: adminEmail,
        createdAt: new Date().toISOString(),
        lastAccessedAt: new Date().toISOString(),
        isActive: 'true',
        ...additionalData
      };
      
      // 既存の安全な作成関数を使用
      if (typeof createUserAtomic === 'function') {
        createUserAtomic(userData);
      } else {
        // フォールバック: 直接Sheetに追加
        const props = PropertiesService.getScriptProperties();
        const dbId = props.getProperty('DATABASE_SPREADSHEET_ID');
        if (!dbId) throw new Error('Database not configured');
        
        const sheet = SpreadsheetApp.openById(dbId).getSheetByName('ユーザー');
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        const newRow = headers.map(header => userData[header] || '');
        sheet.appendRow(newRow);
      }
      
      console.log('findOrCreateUserWithEmailLock: 新規ユーザー作成完了', { userId, adminEmail });
      
      return {
        userId: userId,
        isNewUser: true,
        userInfo: userData
      };
      
    } finally {
      lock.releaseLock();
    }
    
  } finally {
    // Step 4: メール特化ロック解放
    releaseEmailLock(adminEmail);
  }
}

/**
 * クリーンアップ: 古い処理状態をクリア
 * 5分以上古い処理状態は自動削除
 */
function cleanupEmailLocks() {
  const now = Date.now();
  const timeout = 5 * 60 * 1000; // 5分
  
  Object.keys(PROCESSING_EMAILS).forEach(key => {
    if (now - PROCESSING_EMAILS[key].timestamp > timeout) {
      delete PROCESSING_EMAILS[key];
      console.log('Email lock cleanup: Removed stale lock', key);
    }
  });
}