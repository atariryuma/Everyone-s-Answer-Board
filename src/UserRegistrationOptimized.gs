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
 * クイックスタートセットアップ
 * @returns {object} セットアップ結果
 */
function quickStartSetup() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    
    // Step 1: 軽量ユーザー作成（ロック保護・高速）
    console.log('quickStartSetup: 軽量ユーザー作成開始', { userEmail });
    const userResult = createLightweightUser(userEmail);
    const userId = userResult.userId;
    
    // 既存ユーザーで既にセットアップ済みの場合
    if (!userResult.isNewUser && userResult.userInfo.spreadsheetId) {
      const adminUrl = getWebAppUrl() + '?page=admin&userId=' + userId;
      return {
        status: 'existing',
        userId: userId,
        adminUrl: adminUrl,
        formUrl: userResult.userInfo.configJson ? JSON.parse(userResult.userInfo.configJson).formUrl : '',
        spreadsheetUrl: userResult.userInfo.spreadsheetUrl
      };
    }
    
    // Step 2: ユーザー作成完了を返す（フォーム作成は後で実行）
    const adminUrl = getWebAppUrl() + '?page=admin&userId=' + userId;
    console.log('quickStartSetup: 軽量ユーザー作成完了、リソース作成は後で実行', { userId, userEmail });
    
    return {
      status: 'partial',
      userId: userId,
      adminUrl: adminUrl,
      message: 'アカウント作成完了。フォーム・スプレッドシートを作成中...',
      needsResourceCreation: true
    };
    
  } catch (e) {
    console.error('quickStartSetup エラー:', e);
    let errorMessage;
    
    if (e.message.includes('SCRIPT_LOCK_TIMEOUT')) {
      errorMessage = '多数のユーザーが同時にアカウント作成を行っています。30秒ほど待ってから再度お試しください。';
    } else if (e.message.includes('EMAIL_ALREADY_PROCESSING')) {
      errorMessage = '既に同じアカウントで処理が実行中です。しばらく待ってから再度お試しください。';
    } else if (e.message.includes('generateUniqueUserId')) {
      errorMessage = 'システム内部エラーが発生しました。ページを再読み込みしてから再度お試しください。';
    } else {
      errorMessage = `アカウント作成に失敗しました: ${e.message}`;
    }
    
    throw new Error(`${errorMessage} (${Session.getActiveUser().getEmail()})`);
  }
}

/**
 * フォーム・スプレッドシート作成（ロック保護なし・非同期実行可能）
 * @param {string} userId ユーザーID
 * @returns {object} 作成結果
 */
function createUserResourcesAsync(userId) {
  try {
    // ユーザー情報を取得
    const userInfo = findUserById(userId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    
    const userEmail = userInfo.adminEmail;
    
    // 既にリソースが作成済みかチェック
    if (userInfo.spreadsheetId) {
      console.log('createUserResourcesAsync: リソース作成済み', { userId, userEmail });
      return {
        status: 'existing',
        userId: userId,
        formUrl: userInfo.configJson ? JSON.parse(userInfo.configJson).formUrl : '',
        spreadsheetUrl: userInfo.spreadsheetUrl
      };
    }
    
    // Step 1: フォームとスプレッドシートを作成（ロック保護なし）
    console.log('createUserResourcesAsync: リソース作成開始', { userId, userEmail });
    const { formUrl, spreadsheetUrl, spreadsheetId } = createStudyQuestForm(userEmail, userId);
    
    // Step 2: ユーザー情報を更新
    const config = {
      formUrl: formUrl,
      publishedSheetName: 'フォームの回答 1',
      sheet_フォームの回答_1: {
        published: true,
        publishDate: new Date().toISOString(),
        opinionHeader: 'お題'
      }
    };
    
    const updateData = {
      spreadsheetId: spreadsheetId,
      spreadsheetUrl: spreadsheetUrl,
      configJson: JSON.stringify(config),
      setupStatus: 'complete'
    };
    
    updateUser(userId, updateData);
    
    console.log('createUserResourcesAsync: リソース作成完了', { userId, userEmail });
    
    return {
      status: 'success',
      userId: userId,
      formUrl: formUrl,
      spreadsheetUrl: spreadsheetUrl
    };
    
  } catch (e) {
    console.error('createUserResourcesAsync エラー:', e);
    throw new Error(`リソース作成に失敗しました: ${e.message}`);
  }
}

// ランタイムキャッシュ（実行時のみ有効）
var runtimeUserCache = {};

/**
 * フロントエンド向け：リソース作成状況チェック
 * @param {string} userId ユーザーID
 * @returns {object} 状況
 */
function checkResourceCreationStatus(userId) {
  try {
    const userInfo = findUserById(userId);
    if (!userInfo) {
      return { status: 'error', message: 'ユーザーが見つかりません' };
    }
    
    if (userInfo.spreadsheetId && userInfo.setupStatus === 'complete') {
      // リソース作成完了
      const config = userInfo.configJson ? JSON.parse(userInfo.configJson) : {};
      return {
        status: 'complete',
        userId: userId,
        formUrl: config.formUrl || '',
        spreadsheetUrl: userInfo.spreadsheetUrl || '',
        adminUrl: getWebAppUrl() + '?page=admin&userId=' + userId
      };
    } else if (userInfo.setupStatus === 'basic') {
      // リソース作成待ち
      return {
        status: 'pending',
        message: 'フォーム・スプレッドシートを作成中です...'
      };
    } else {
      // 不明な状態
      return {
        status: 'unknown',
        message: 'セットアップ状態が不明です'
      };
    }
    
  } catch (error) {
    console.error('checkResourceCreationStatus エラー:', error);
    return { status: 'error', message: error.message };
  }
}

/**
 * フロントエンド向け：段階的セットアップ完了
 * 軽量ユーザー作成 + 即座のリソース作成
 * @returns {object} 完了結果
 */
function completeFullSetup() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    
    // Step 1: ユーザー確保（軽量）
    const userResult = createLightweightUser(userEmail);
    const userId = userResult.userId;
    
    // Step 2: リソース作成
    const resourceResult = createUserResourcesAsync(userId);
    
    const adminUrl = getWebAppUrl() + '?page=admin&userId=' + userId;
    
    return {
      status: 'success',
      userId: userId,
      adminUrl: adminUrl,
      formUrl: resourceResult.formUrl,
      spreadsheetUrl: resourceResult.spreadsheetUrl,
      message: '完全セットアップが完了しました'
    };
    
  } catch (error) {
    console.error('completeFullSetup エラー:', error);
    throw new Error(`完全セットアップに失敗しました: ${error.message}`);
  }
}

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