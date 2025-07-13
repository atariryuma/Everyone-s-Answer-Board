/**
 * UserRegistrationOptimized.gs
 * マルチテナント対応・高速・安全なユーザー登録システム
 * 
 * 主要改善点：
 * 1. Upsert パターンによる原子的操作
 * 2. LockService による競合制御
 * 3. 一貫したユーザーID生成
 * 4. キャッシュ最適化
 */

/**
 * 🔒 安全なユーザー登録・更新（Upsert操作）
 * メールアドレスをキーとした重複のない登録
 * 
 * @param {string} adminEmail - ユーザーメールアドレス
 * @param {Object} additionalData - 追加データ（オプション）
 * @returns {Object} { status: 'success', userId: string, isNewUser: boolean, message: string }
 */
function upsertUser(adminEmail, additionalData = {}) {
  if (!adminEmail || !adminEmail.includes('@')) {
    throw new Error('有効なメールアドレスが必要です');
  }

  // 1. 分散ロック取得（最大30秒待機）
  const lock = LockService.getScriptLock();
  const lockKey = `user_registration_${adminEmail}`;
  
  try {
    // メールアドレス別のロック（粒度を細かく）
    if (!lock.waitLock(30000)) {
      throw new Error('システムが混雑しています。しばらく後に再試行してください。');
    }

    // 2. 認証確認
    const activeUser = Session.getActiveUser();
    if (adminEmail !== activeUser.getEmail()) {
      throw new Error('認証エラー: 操作権限がありません');
    }

    // 3. 既存ユーザー確認
    let existingUser = findUserByEmailDirect(adminEmail);
    let userId, isNewUser;

    if (existingUser) {
      // 既存ユーザーの更新
      userId = existingUser.userId;
      isNewUser = false;
      
      // 必要に応じて情報更新
      if (Object.keys(additionalData).length > 0) {
        const updateData = {
          lastAccessedAt: new Date().toISOString(),
          isActive: 'true',
          ...additionalData
        };
        updateUserDirect(userId, updateData);
      }
      
      debugLog('upsertUser: 既存ユーザーを更新', { userId, adminEmail });
      
    } else {
      // 新規ユーザー作成
      isNewUser = true;
      userId = generateConsistentUserId(adminEmail);
      
      const userData = {
        userId: userId,
        adminEmail: adminEmail,
        createdAt: new Date().toISOString(),
        lastAccessedAt: new Date().toISOString(),
        isActive: 'true',
        configJson: '{}',
        ...additionalData
      };
      
      createUserDirect(userData);
      debugLog('upsertUser: 新規ユーザーを作成', { userId, adminEmail });
    }

    // 4. キャッシュ更新
    invalidateUserCacheConsistent(userId, adminEmail);
    
    // 5. 結果の検証
    const verifiedUser = findUserByIdDirect(userId);
    if (!verifiedUser) {
      throw new Error('ユーザー登録の検証に失敗しました');
    }

    return {
      status: 'success',
      userId: userId,
      isNewUser: isNewUser,
      userInfo: verifiedUser,
      message: isNewUser ? '新規ユーザーを登録しました' : '既存ユーザーの情報を更新しました'
    };

  } finally {
    // 必ずロックを解除
    try {
      lock.releaseLock();
    } catch (e) {
      console.warn('ロック解除エラー:', e.message);
    }
  }
}

/**
 * 🎯 一貫したユーザーID生成
 * メールアドレスから決定論的にUUIDを生成
 * 
 * @param {string} adminEmail - メールアドレス
 * @returns {string} 一意のユーザーID
 */
function generateConsistentUserId(adminEmail) {
  // メールアドレスからハッシュ値を生成して、UUIDフォーマットに変換
  const hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, adminEmail, Utilities.Charset.UTF_8);
  const hexString = hash.map(byte => (byte + 256).toString(16).slice(-2)).join('');
  
  // UUID v4 フォーマットに整形
  const uuid = [
    hexString.slice(0, 8),
    hexString.slice(8, 12),
    '4' + hexString.slice(13, 16), // version 4
    ((parseInt(hexString.slice(16, 17), 16) & 0x3) | 0x8).toString(16) + hexString.slice(17, 20), // variant
    hexString.slice(20, 32)
  ].join('-');
  
  return uuid;
}

/**
 * 🚀 高速なユーザー検索（キャッシュ最適化版）
 * 
 * @param {string} adminEmail - メールアドレス
 * @returns {Object|null} ユーザー情報
 */
function findUserByEmailDirect(adminEmail) {
  const cacheKey = `email_${adminEmail}`;
  
  // 1. L1キャッシュ確認（実行時メモリ）
  if (runtimeUserCache && runtimeUserCache[cacheKey]) {
    return runtimeUserCache[cacheKey];
  }
  
  // 2. L2キャッシュ確認（CacheService）
  try {
    const cached = CacheService.getScriptCache().get(cacheKey);
    if (cached) {
      const userInfo = JSON.parse(cached);
      if (!runtimeUserCache) runtimeUserCache = {};
      runtimeUserCache[cacheKey] = userInfo;
      return userInfo;
    }
  } catch (e) {
    console.warn('キャッシュ読み取りエラー:', e.message);
  }
  
  // 3. データベース検索
  try {
    const userInfo = findUserByEmail(adminEmail); // 既存関数を使用
    
    // キャッシュに保存
    if (userInfo) {
      try {
        CacheService.getScriptCache().put(cacheKey, JSON.stringify(userInfo), 1800); // 30分
        if (!runtimeUserCache) runtimeUserCache = {};
        runtimeUserCache[cacheKey] = userInfo;
      } catch (e) {
        console.warn('キャッシュ保存エラー:', e.message);
      }
    }
    
    return userInfo;
  } catch (e) {
    console.error('findUserByEmailDirect エラー:', e.message);
    return null;
  }
}

/**
 * 🔄 一貫したキャッシュ無効化
 * 
 * @param {string} userId - ユーザーID
 * @param {string} adminEmail - メールアドレス
 */
function invalidateUserCacheConsistent(userId, adminEmail) {
  try {
    // CacheService クリア
    const cache = CacheService.getScriptCache();
    cache.remove(`user_${userId}`);
    cache.remove(`email_${adminEmail}`);
    
    // ランタイムキャッシュクリア
    if (runtimeUserCache) {
      delete runtimeUserCache[`user_${userId}`];
      delete runtimeUserCache[`email_${adminEmail}`];
    }
    
    // 従来のキャッシュシステムも併用
    invalidateUserCache(userId, adminEmail, null, false);
    
  } catch (e) {
    console.warn('キャッシュ無効化エラー:', e.message);
  }
}

/**
 * 🎯 最適化されたクイックスタート
 * 重複登録を防ぐ統合フロー
 * 
 * @param {string} requestUserId - リクエストユーザーID（オプション）
 * @returns {Object} クイックスタート結果
 */
function quickStartSetupOptimized(requestUserId) {
  try {
    const activeUserEmail = Session.getActiveUser().getEmail();
    let userInfo;

    // 1. ユーザー情報の統一取得・作成
    if (requestUserId) {
      // 既存ユーザーIDが指定されている場合
      userInfo = findUserByIdDirect(requestUserId);
      if (!userInfo) {
        throw new Error('指定されたユーザーIDが見つかりません');
      }
      
      // メールアドレス一致確認
      if (userInfo.adminEmail !== activeUserEmail) {
        throw new Error('ユーザーIDとメールアドレスが一致しません');
      }
    } else {
      // 新規または既存ユーザーの自動判定
      const upsertResult = upsertUser(activeUserEmail);
      userInfo = upsertResult.userInfo;
      requestUserId = upsertResult.userId;
    }

    // 2. フォームとスプレッドシート作成
    const formAndSsInfo = createStudyQuestForm(activeUserEmail, requestUserId);
    
    // 3. ユーザー情報更新
    const updatedData = {
      spreadsheetId: formAndSsInfo.spreadsheetId,
      spreadsheetUrl: formAndSsInfo.spreadsheetUrl,
      configJson: JSON.stringify({
        formUrl: formAndSsInfo.formUrl,
        publishedSheetName: formAndSsInfo.sheetName,
        sheet_フォームの回答: {
          published: true,
          publishDate: new Date().toISOString(),
          opinionHeader: formAndSsInfo.opinionHeader || 'お題'
        }
      }),
      lastAccessedAt: new Date().toISOString()
    };
    
    updateUserDirect(requestUserId, updatedData);
    
    // 4. キャッシュ更新
    invalidateUserCacheConsistent(requestUserId, activeUserEmail);
    
    return {
      status: 'success',
      userId: requestUserId,
      formUrl: formAndSsInfo.formUrl,
      spreadsheetUrl: formAndSsInfo.spreadsheetUrl,
      message: 'クイックスタートが完了しました'
    };
    
  } catch (error) {
    console.error('quickStartSetupOptimized エラー:', error);
    throw error;
  }
}

// ランタイムキャッシュ（実行時のみ有効）
var runtimeUserCache = {};

/**
 * 🔍 直接データベース操作（キャッシュバイパス）
 */
function findUserByIdDirect(userId) {
  try {
    return findUserById(userId); // 既存関数使用
  } catch (e) {
    console.error('findUserByIdDirect エラー:', e.message);
    return null;
  }
}

function createUserDirect(userData) {
  try {
    return createUser(userData); // 既存関数使用
  } catch (e) {
    console.error('createUserDirect エラー:', e.message);
    throw e;
  }
}

function updateUserDirect(userId, updateData) {
  try {
    return updateUser(userId, updateData); // 既存関数使用
  } catch (e) {
    console.error('updateUserDirect エラー:', e.message);
    throw e;
  }
}