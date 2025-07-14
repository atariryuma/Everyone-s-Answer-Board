/**
 * メール特化ロックシステム
 * 同一ユーザーの重複操作を防ぎ、他ユーザーとの競合のみロック対象とする
 */

// メール処理状態はCacheServiceで管理（実行間で共有）
const EMAIL_LOCK_CACHE = CacheService.getScriptCache();
// スクリプトロックの保持者情報を共有
const SCRIPT_LOCK_INFO_CACHE = CacheService.getScriptCache();
const SCRIPT_LOCK_HOLDER_KEY = 'scriptLockHolder';

/**
 * メール特化ロック：同一メールの重複処理を防止
 * @param {string} adminEmail - メールアドレス
 * @returns {boolean} ロック取得成功かどうか
 */
function acquireEmailLock(adminEmail) {
  const lockKey = 'processing_' + adminEmail;

  // 既に同じメールで処理中の場合は拒否
  if (EMAIL_LOCK_CACHE.get(lockKey)) {
    console.log('Email-specific lock: Already processing for', adminEmail);
    return false;
  }

  // ロック取得（60秒有効）
  EMAIL_LOCK_CACHE.put(lockKey, String(Date.now()), 60);

  console.log('Email-specific lock: Acquired for', adminEmail);
  return true;
}

/**
 * メール特化ロック解放
 * @param {string} adminEmail - メールアドレス
 */
function releaseEmailLock(adminEmail) {
  const lockKey = 'processing_' + adminEmail;
  EMAIL_LOCK_CACHE.remove(lockKey);
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
    const timeout = 30000; // 30秒に延長（フォーム・スプレッドシート作成時間を考慮）

    if (!lock.waitLock(timeout)) {
      const holder = SCRIPT_LOCK_INFO_CACHE.get(SCRIPT_LOCK_HOLDER_KEY);
      console.error('findOrCreateUserWithEmailLock: スクリプトロックタイムアウト', {
        adminEmail,
        timeout,
        lockHolder: holder || '不明'
      });
      // タイムアウト時に関連するメールロックをクリーンアップ
      cleanupEmailLocks();
      throw new Error(`SCRIPT_LOCK_TIMEOUT: Lock held by ${holder || 'unknown'}`);
    }

    SCRIPT_LOCK_INFO_CACHE.put(SCRIPT_LOCK_HOLDER_KEY, adminEmail, 30);
    
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
      
      // 決定論的IDを事前生成
      const userId = generateConsistentUserId(adminEmail);
      const userData = {
        userId: userId,
        adminEmail: adminEmail,
        createdAt: new Date().toISOString(),
        lastAccessedAt: new Date().toISOString(),
        isActive: 'true',
        ...additionalData
      };
      
      // 楽観的ロック：作成試行→エラー時既存ユーザー検索
      try {
        createUserAtomic(userData);
        console.log('findOrCreateUserWithEmailLock: 新規ユーザー作成完了', { userId, adminEmail });
        
        return {
          userId: userId,
          isNewUser: true,
          userInfo: userData
        };
      } catch (createError) {
        // 作成失敗時（重複など）は既存ユーザーを返す
        console.log('findOrCreateUserWithEmailLock: 作成失敗、既存ユーザー検索', { adminEmail, error: createError.message });
        const fallbackUser = findUserByEmailNonBlocking(adminEmail);
        if (fallbackUser) {
          return {
            userId: fallbackUser.userId,
            isNewUser: false,
            userInfo: fallbackUser
          };
        }
        // それでも見つからない場合は元のエラーを再スロー
        throw createError;
      }
      
    } finally {
      lock.releaseLock();
      SCRIPT_LOCK_INFO_CACHE.remove(SCRIPT_LOCK_HOLDER_KEY);
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
  // CacheService のエントリは自動的に期限切れになるため明示的なクリーンアップは不要
}

/**
 * 軽量ユーザー作成（基本情報のみ）
 * フォーム・スプレッドシート作成は後回し
 * @param {string} adminEmail - メールアドレス
 * @param {object} additionalData - 追加データ
 * @returns {object} 結果
 */
function createLightweightUser(adminEmail, additionalData = {}) {
  // Step 1: メール特化ロック取得
  if (!acquireEmailLock(adminEmail)) {
    throw new Error('EMAIL_ALREADY_PROCESSING');
  }
  
  try {
    // Step 2: 既存ユーザー確認（軽量）
    let existingUser = findUserByEmailNonBlocking(adminEmail);
    console.log('createLightweightUser: 既存ユーザー確認結果', { existingUser: !!existingUser, adminEmail });
    
    if (existingUser) {
      return {
        userId: existingUser.userId,
        isNewUser: false,
        userInfo: existingUser
      };
    }
    
    // Step 3: 軽量ユーザー作成（基本情報のみ）
    console.log('createLightweightUser: 軽量ユーザー作成開始', { adminEmail });
    
    const userId = generateConsistentUserId(adminEmail);
    console.log('createLightweightUser: 生成されたユーザーID', { userId, adminEmail });
    
    const userData = {
      userId: userId,
      adminEmail: adminEmail,
      createdAt: new Date().toISOString(),
      lastAccessedAt: new Date().toISOString(),
      isActive: 'true',
      configJson: '{}', // 空の設定
      spreadsheetId: '', // 未設定
      spreadsheetUrl: '', // 未設定
      setupStatus: 'basic', // 基本作成完了
      ...additionalData
    };
    
    console.log('createLightweightUser: 作成するユーザーデータ', userData);
    
    // 楽観的ロック：作成試行→エラー時既存ユーザー検索
    try {
      createUserAtomic(userData);
      console.log('createLightweightUser: 軽量ユーザー作成完了', { userId, adminEmail });
      
      // 作成直後に確認
      const verificationUser = findUserByEmailNonBlocking(adminEmail);
      console.log('createLightweightUser: 作成後の確認結果', { 
        found: !!verificationUser, 
        foundUserId: verificationUser?.userId,
        expectedUserId: userId,
        match: verificationUser?.userId === userId
      });
      
      return {
        userId: userId,
        isNewUser: true,
        userInfo: userData
      };
    } catch (createError) {
      console.log('createLightweightUser: 作成失敗、既存ユーザー検索', { adminEmail, error: createError.message });
      const fallbackUser = findUserByEmailNonBlocking(adminEmail);
      if (fallbackUser) {
        return {
          userId: fallbackUser.userId,
          isNewUser: false,
          userInfo: fallbackUser
        };
      }
      throw createError;
    }
    
  } finally {
    // Step 4: メール特化ロック解放
    releaseEmailLock(adminEmail);
  }
}
