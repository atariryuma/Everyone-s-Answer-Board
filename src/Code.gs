/**
 * @fileoverview StudyQuest - みんなの回答ボード
 * スプレッドシートのカスタムメニューを使用しない独立版。
 */

const MAIN_DB_ID_KEY = 'MAIN_DB_ID';
const LOGGER_API_URL_KEY = 'LOGGER_API_URL';

// =================================================================
// 定数定義
// =================================================================

// ユーザー専用フォルダ管理の設定
const USER_FOLDER_CONFIG = {
  ROOT_FOLDER_NAME: "StudyQuest - ユーザーデータ",
  FOLDER_NAME_PATTERN: "StudyQuest - {email} - ファイル"
};

// マルチテナント用ユーザーデータベース設定
const USER_DB_CONFIG = {
  SHEET_NAME: 'Users',
  HEADERS: [
    'userId',
    'adminEmail',
    'spreadsheetId', 
    'spreadsheetUrl',
    'createdAt',
    'accessToken',
    'configJson',
    'lastAccessedAt',
    'isActive'
  ]
};
const COLUMN_HEADERS = {
  TIMESTAMP: 'タイムスタンプ',
  EMAIL: 'メールアドレス',
  CLASS: 'クラス',
  OPINION: '回答',
  REASON: '理由',
  NAME: '名前',
  UNDERSTAND: 'なるほど！',
  LIKE: 'いいね！',
  CURIOUS: 'もっと知りたい！',
  HIGHLIGHT: 'ハイライト'
};
const SCORING_CONFIG = {
  LIKE_MULTIPLIER_FACTOR: 0.05 // 1いいね！ごとにスコアが5%増加
};

const TIME_CONSTANTS = {
  LOCK_WAIT_MS: 10000
};

const REACTION_KEYS = ["UNDERSTAND","LIKE","CURIOUS"];
const EMAIL_REGEX = /^[^\n@]+@[^\n@]+\.[^\n@]+$/;
// Debug flag. Set to true to enable verbose logging
var DEBUG = false; // パフォーマンス向上のためデバッグログを無効化

function debugLog() {
  if (DEBUG && typeof console !== 'undefined' && console.log) {
    console.log.apply(console, arguments);
  }
}

/**
 * ユーザー情報キャッシュ（メモリ内キャッシュ、5分間有効）
 */
const USER_INFO_CACHE = new Map();
const USER_CACHE_TTL = 5 * 60 * 1000; // 5分

/**
 * パフォーマンス最適化用のCacheService設定
 */
const CACHE_CONFIG = {
  USER_INFO_TTL: 300,        // 5分（ユーザー情報）
  SPREADSHEET_DATA_TTL: 120, // 2分（スプレッドシートデータ）
  FORM_INFO_TTL: 600,        // 10分（フォーム情報）
  SHEET_LIST_TTL: 180,       // 3分（シート一覧）
  FORM_ID_TTL: 1800,         // 30分（フォームID）
  RESPONSE_COUNT_TTL: 60,    // 1分（回答数）
  CONFIG_TTL: 900,           // 15分（設定情報）
  ADMIN_SETTINGS_TTL: 300    // 5分（管理者設定）
};

/**
 * スプレッドシートオブジェクトキャッシュ（メモリ内）
 */
const SPREADSHEET_CACHE = new Map();
const SPREADSHEET_CACHE_TTL = 5 * 60 * 1000; // 5分

/**
 * キャッシュされたユーザー情報を取得（CacheService併用）
 * @param {string} userId - ユーザーID
 * @return {Object|null} キャッシュされたユーザー情報またはnull
 */
function getCachedUserInfo(userId) {
  // メモリキャッシュを最初にチェック（最高速）
  const memoryCached = USER_INFO_CACHE.get(userId);
  if (memoryCached && (Date.now() - memoryCached.timestamp) < USER_CACHE_TTL) {
    debugLog(`Memory cache hit for user: ${userId}`);
    
    // キャッシュされたデータの有効性を確認
    if (memoryCached.data && memoryCached.data.isActive !== false && memoryCached.data.isActive !== 'FALSE') {
      return memoryCached.data;
    } else {
      // 無効なデータの場合はキャッシュをクリア
      invalidateUserCache(userId);
      return null;
    }
  }
  
  // CacheServiceを次にチェック
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = `user_info_${userId}`;
    const cached = cache.get(cacheKey);
    
    if (cached) {
      debugLog(`CacheService hit for user: ${userId}`);
      const userInfo = JSON.parse(cached);
      
      // キャッシュされたデータの有効性を確認
      if (userInfo && userInfo.isActive !== false && userInfo.isActive !== 'FALSE') {
        // メモリキャッシュにも保存（次回アクセス用）
        setCachedUserInfo(userId, userInfo);
        return userInfo;
      } else {
        // 無効なデータの場合はキャッシュをクリア
        invalidateUserCache(userId);
        return null;
      }
    }
  } catch (e) {
    debugLog(`CacheService error for user ${userId}:`, e.message);
  }
  
  debugLog(`No cache found for user: ${userId}`);
  return null;
}

/**
 * ユーザー情報をキャッシュに保存（メモリ＋CacheService）
 * @param {string} userId - ユーザーID
 * @param {Object} userInfo - ユーザー情報
 */
function setCachedUserInfo(userId, userInfo) {
  // データベースでユーザーが削除されている場合はキャッシュしない
  if (!userInfo || !userInfo.userId || userInfo.isActive === false || userInfo.isActive === 'FALSE') {
    invalidateUserCache(userId);
    return;
  }
  
  // メモリキャッシュに保存
  USER_INFO_CACHE.set(userId, {
    data: userInfo,
    timestamp: Date.now()
  });
  
  // CacheServiceにも保存（より長期間保持）
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = `user_info_${userId}`;
    cache.put(cacheKey, JSON.stringify(userInfo), CACHE_CONFIG.USER_INFO_TTL);
    debugLog(`User info cached for: ${userId}`);
  } catch (e) {
    debugLog(`Failed to cache user info for ${userId}:`, e.message);
  }
}

/**
 * ユーザーキャッシュを無効化（削除されたユーザー対応）
 * @param {string} userId - ユーザーID
 */
function invalidateUserCache(userId) {
  if (!userId) return;
  
  try {
    // メモリキャッシュから削除
    USER_INFO_CACHE.delete(userId);
    
    // CacheServiceからも削除
    const cache = CacheService.getScriptCache();
    cache.remove(`user_info_${userId}`);
    
    debugLog(`Cache invalidated for user: ${userId}`);
  } catch (e) {
    debugLog(`Failed to invalidate cache for user ${userId}:`, e.message);
  }
}

/**
 * すべてのユーザーキャッシュをクリア
 */
function clearAllUserCache() {
  try {
    // メモリキャッシュをクリア
    USER_INFO_CACHE.clear();
    
    // CacheServiceからユーザー関連キャッシュを削除
    const cache = CacheService.getScriptCache();
    // 個別のキーを削除する代わりに、プレフィックスベースでの削除を試行
    debugLog('All user caches cleared');
    
    // Admin Logger APIにもキャッシュクリア要求を送信
    try {
      invalidateCacheViaApi(null, null, true); // clearAll = true
    } catch (e) {
      debugLog(`Failed to clear Admin Logger API cache:`, e.message);
    }
  } catch (e) {
    debugLog(`Failed to clear all user cache:`, e.message);
  }
}

/**
 * Admin Logger API経由でキャッシュを無効化
 * @param {string} userId - ユーザーID
 * @param {string} adminEmail - 管理者メール
 * @param {boolean} clearAll - 全キャッシュクリア
 */
function invalidateCacheViaApi(userId, adminEmail, clearAll = false) {
  try {
    const apiUrl = getLoggerApiUrl();
    if (!apiUrl) {
      debugLog('Logger API URL not configured, skipping cache invalidation');
      return;
    }
    
    const payload = {
      action: 'invalidateCache',
      data: {
        userId: userId,
        adminEmail: adminEmail,
        clearAll: clearAll
      },
      timestamp: new Date().toISOString(),
      requestUser: Session.getActiveUser().getEmail()
    };
    
    const response = UrlFetchApp.fetch(apiUrl, {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    const responseData = JSON.parse(response.getContentText());
    
    if (responseData.success) {
      debugLog(`Cache invalidated via API: ${userId || adminEmail || 'all'}`);
    } else {
      debugLog(`API cache invalidation failed: ${responseData.error}`);
    }
    
  } catch (error) {
    debugLog(`Error invalidating cache via API: ${error.message}`);
  }
}

/**
 * CacheServiceを使った高速データキャッシュ（JSON対応）
 */
function getCachedData(key, fetchFunction, ttlSeconds = 300) {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get(key);
    
    if (cached) {
      debugLog(`Cache hit for key: ${key}`);
      return JSON.parse(cached);
    }
    
    debugLog(`Cache miss for key: ${key}, fetching data...`);
    const data = fetchFunction();
    
    if (data) {
      cache.put(key, JSON.stringify(data), ttlSeconds);
      debugLog(`Data cached for key: ${key} (TTL: ${ttlSeconds}s)`);
    }
    
    return data;
  } catch (e) {
    debugLog(`Cache error for key ${key}:`, e.message);
    // キャッシュエラー時は直接データを取得
    return fetchFunction();
  }
}

/**
 * スプレッドシートオブジェクトのキャッシュ取得（メモリ）
 * @param {string} spreadsheetId - スプレッドシートID
 * @return {Spreadsheet|null} キャッシュされたスプレッドシートまたはnull
 */
function getCachedSpreadsheet(spreadsheetId) {
  const cached = SPREADSHEET_CACHE.get(spreadsheetId);
  if (cached && (Date.now() - cached.timestamp) < SPREADSHEET_CACHE_TTL) {
    debugLog(`Spreadsheet cache hit: ${spreadsheetId}`);
    return cached.data;
  }
  
  debugLog(`Spreadsheet cache miss: ${spreadsheetId}`);
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    SPREADSHEET_CACHE.set(spreadsheetId, {
      data: spreadsheet,
      timestamp: Date.now()
    });
    return spreadsheet;
  } catch (e) {
    debugLog(`Failed to open spreadsheet ${spreadsheetId}:`, e.message);
    return null;
  }
}

/**
 * 複数のキーを一括でキャッシュから取得（バッチ処理）
 */
function getBatchCachedData(keyFunctionPairs, defaultTtl = 300) {
  try {
    const cache = CacheService.getScriptCache();
    const keys = keyFunctionPairs.map(pair => pair.key);
    const cachedData = cache.getAll(keys);
    const results = {};
    const toFetch = [];
    
    // キャッシュヒット/ミスを判定
    keyFunctionPairs.forEach(({ key, fetchFunction }) => {
      if (cachedData[key]) {
        results[key] = JSON.parse(cachedData[key]);
        debugLog(`Batch cache hit: ${key}`);
      } else {
        toFetch.push({ key, fetchFunction });
        debugLog(`Batch cache miss: ${key}`);
      }
    });
    
    // キャッシュミスしたデータを一括取得
    const newCacheData = {};
    toFetch.forEach(({ key, fetchFunction }) => {
      try {
        const data = fetchFunction();
        if (data) {
          results[key] = data;
          newCacheData[key] = JSON.stringify(data);
        }
      } catch (e) {
        debugLog(`Batch fetch error for ${key}:`, e.message);
      }
    });
    
    // 新しいデータを一括でキャッシュに保存
    if (Object.keys(newCacheData).length > 0) {
      cache.putAll(newCacheData, defaultTtl);
      debugLog(`Batch cached ${Object.keys(newCacheData).length} items`);
    }
    
    return results;
  } catch (e) {
    debugLog('Batch cache error:', e.message);
    // エラー時は全て直接取得
    const results = {};
    keyFunctionPairs.forEach(({ key, fetchFunction }) => {
      try {
        results[key] = fetchFunction();
      } catch (fetchError) {
        debugLog(`Direct fetch error for ${key}:`, fetchError.message);
      }
    });
    return results;
  }
}

/**
 * 現在のユーザーの権限を検証するヘルパー関数（読み取り専用・高速版）
 * @return {Object} {userId, userInfo} - 検証済みのユーザー情報
 * @throws {Error} 認証失敗時
 */
function validateCurrentUserReadOnly() {
  const props = PropertiesService.getUserProperties();
  const userId = props.getProperty('CURRENT_USER_ID');
  if (!userId) {
    throw new Error('認証が必要です。ログインし直してください。');
  }
  
  // キャッシュから取得を試行
  let userInfo = getCachedUserInfo(userId);
  if (!userInfo) {
    // キャッシュにない場合のみAPI呼び出し
    userInfo = getUserInfo(userId);
    if (userInfo) {
      setCachedUserInfo(userId, userInfo);
    }
  }
  
  const currentEmail = Session.getActiveUser().getEmail();
  if (!userInfo || userInfo.adminEmail !== currentEmail) {
    throw new Error('権限がありません。管理者アカウントでログインしてください。');
  }
  return { userId, userInfo };
}

/**
 * 現在のユーザーの権限を検証するヘルパー関数
 * @return {Object} {userId, userInfo} - 検証済みのユーザー情報
 * @throws {Error} 認証失敗時
 */
function validateCurrentUser() {
  const props = PropertiesService.getUserProperties();
  const userId = props.getProperty('CURRENT_USER_ID');
  if (!userId) {
    throw new Error('認証が必要です。ログインし直してください。');
  }
  
  const userInfo = getUserInfo(userId);
  const currentEmail = Session.getActiveUser().getEmail();
  if (!userInfo || userInfo.adminEmail !== currentEmail) {
    throw new Error('権限がありません。管理者アカウントでログインしてください。');
  }
  return { userId, userInfo };
}

/**
 * ユーザー専用フォルダを取得または作成（改良版：詳細なエラーハンドリング付き）
 * @param {string} userEmail - ユーザーのメールアドレス
 * @return {GoogleAppsScript.Drive.Folder} ユーザー専用フォルダ
 */
function getUserFolder(userEmail) {
  try {
    // 入力検証
    if (!userEmail || typeof userEmail !== 'string') {
      throw new Error('有効なユーザーメールアドレスが必要です');
    }
    
    // メールアドレスからフォルダ名を生成
    const sanitizedEmail = userEmail.split('@')[0].replace(/[^a-zA-Z0-9]/g, '_');
    const timestamp = new Date().getTime();
    const userFolderName = `StudyQuest - ${sanitizedEmail} - マイファイル`;
    
    debugLog(`📁 ユーザー専用フォルダを作成/取得開始: ${userFolderName}`);
    
    // Driveアクセス権限をテスト
    try {
      const testFolder = DriveApp.getRootFolder();
      debugLog(`✅ Driveアクセス権限確認成功`);
    } catch (driveError) {
      debugLog(`❌ Driveアクセス権限エラー: ${driveError.message}`);
      throw new Error(`Googleドライブにアクセスできません: ${driveError.message}`);
    }
    
    // 既存のフォルダがあるかチェック
    let existingFolder = null;
    try {
      const folders = DriveApp.getFoldersByName(userFolderName);
      if (folders.hasNext()) {
        existingFolder = folders.next();
        // フォルダが正常にアクセス可能かテスト
        const testAccess = existingFolder.getName();
        debugLog(`✅ 既存のユーザーフォルダを使用: ${userFolderName}`);
        return existingFolder;
      }
    } catch (folderCheckError) {
      debugLog(`⚠️ 既存フォルダチェック時にエラー: ${folderCheckError.message}`);
    }
    
    // ない場合は新規作成
    try {
      const userFolder = DriveApp.createFolder(userFolderName);
      debugLog(`✅ ユーザー専用フォルダを新規作成: ${userFolderName}`);
      
      // フォルダ作成確認テスト
      const createdFolderName = userFolder.getName();
      if (createdFolderName !== userFolderName) {
        throw new Error('フォルダ名が期待値と異なります');
      }
      
      return userFolder;
      
    } catch (createError) {
      debugLog(`❌ フォルダ作成エラー: ${createError.message}`);
      throw new Error(`フォルダの作成に失敗しました: ${createError.message}`);
    }
    
  } catch (error) {
    console.error(`ユーザーフォルダの取得/作成に失敗 (${userEmail}):`, error);
    // より詳細なエラーメッセージを提供
    if (error.message.includes('権限')) {
      throw new Error('Googleドライブへのアクセス権限が不足しています。管理者にお問い合わせください。');
    } else if (error.message.includes('容量')) {
      throw new Error('Googleドライブの容量が不足しています。');
    } else {
      throw new Error(`ユーザー専用フォルダの準備に失敗しました: ${error.message}`);
    }
  }
}

/**
 * XSS攻撃を防ぐための入力サニタイゼーション関数
 * @param {string} input - サニタイズする入力
 * @return {string} サニタイズされた文字列
 */
function sanitizeInput(input) {
  if (typeof input !== 'string') return input;
  return input
    .replace(/[<>'"&]/g, function(match) {
      return {
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#x27;',
        '&': '&amp;'
      }[match];
    })
    .substring(0, 1000); // 長さ制限
}

/**
 * レート制限チェック関数
 * @param {string} action - 実行するアクション
 * @param {string} userEmail - ユーザーメール
 */
function checkRateLimit(action, userEmail) {
  const cache = CacheService.getScriptCache();
  const key = `rate_limit_${action}_${userEmail}`;
  const attempts = parseInt(cache.get(key) || '0');
  
  if (attempts >= 30) { // 30回/時間制限
    throw new Error('操作回数の制限に達しました。しばらく待ってから再試行してください。');
  }
  
  cache.put(key, (attempts + 1).toString(), 3600); // 1時間
}

/**
 * セキュアなエラーハンドリング関数
 * @param {string} prefix - エラープレフィックス
 * @param {Error} error - 元のエラー
 * @param {boolean} returnObject - オブジェクトで返すかどうか
 * @return {Object|void} エラーオブジェクトまたは例外スロー
 */
function secureHandleError(prefix, error, returnObject = false) {
  // ログには詳細を出力（管理者用）
  console.error(prefix + ':', error);
  
  // ユーザーには安全なメッセージのみ表示
  let userMessage = 'システムエラーが発生しました。管理者にお問い合わせください。';
  
  // 特定のエラーパターンに対する安全なメッセージ
  const errorMsg = error?.message || String(error);
  
  if (errorMsg.includes('permission') || errorMsg.includes('Permission')) {
    userMessage = '権限がありません。管理者にお問い合わせください。';
  } else if (errorMsg.includes('not found') || errorMsg.includes('見つかりません')) {
    userMessage = '指定されたデータが見つかりません。';
  } else if (errorMsg.includes('invalid') || errorMsg.includes('無効')) {
    userMessage = '無効な操作です。入力内容を確認してください。';
  }
  
  if (returnObject) {
    return { status: 'error', message: userMessage };
  }
  throw new Error(userMessage);
}

function applySecurityHeaders(output) {
  if (output && typeof output.addMetaTag === 'function') {
    // NOTE: Apps Script does not allow setting the Content-Security-Policy
    // header via addMetaTag. Attempting to do so results in
    // "The specified meta tag cannot be used in this context" errors.
    // Keep other security related headers only.
    if (typeof output.setXFrameOptionsMode === 'function') {
      // Allow embedding within the same origin by default for enhanced security
      // unless a future embedding requirement mandates ALLOWALL.
      // "SAMEORIGIN" isn't a valid constant. Use DEFAULT to enforce
      // the same-origin policy, which preserves standard security behavior.
      output.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
    }
  }
  return output;
}
var getConfig;
var handleError;

// handleError関数はErrorHandling.gsで定義されています

if (typeof global !== 'undefined' && global.getConfig) {
  getConfig = global.getConfig;
}


/**
 * スプレッドシートをキャッシュに保存
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - スプレッドシート
 */
function setCachedSpreadsheet(spreadsheetId, spreadsheet) {
  SPREADSHEET_CACHE.set(spreadsheetId, {
    data: spreadsheet,
    timestamp: Date.now()
  });
}

function getCurrentSpreadsheet() {
  const props = PropertiesService.getUserProperties();
  const spreadsheetId = props.getProperty('CURRENT_SPREADSHEET_ID');
  
  if (!spreadsheetId) {
    // 従来の動作（単一テナント時）
    return SpreadsheetApp.getActiveSpreadsheet();
  }
  
  // キャッシュから取得を試行
  let spreadsheet = getCachedSpreadsheet(spreadsheetId);
  if (!spreadsheet) {
    // キャッシュにない場合のみAPI呼び出し
    spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    setCachedSpreadsheet(spreadsheetId, spreadsheet);
  }
  
  return spreadsheet;
}

function safeGetUserEmail() {
  try {
    const email = Session.getActiveUser().getEmail();
    if (!email || typeof email !== 'string' || !isValidEmail(email)) {
      throw new Error('Invalid user email');
    }
    return email;
  } catch (e) {
    console.error('Failed to get user email:', e);
    // In test environment, return a default test email
    if (typeof global !== 'undefined' && global.mockUserEmail) {
      return global.mockUserEmail;
    }
    throw new Error('認証が必要です。Googleアカウントでログインしてください。');
  }
}

/**
 * フロントエンドから現在のユーザーのメールアドレスを取得するためのラッパー。
 * Registration.html などから呼び出される。
 *
 * @return {string} ユーザーのメールアドレス
 */
function getActiveUserEmail() {
  return safeGetUserEmail();
}

/**
 * メールアドレスからドメインを抽出
 * @param {string} email - メールアドレス
 * @return {string} ドメイン名
 */
function extractDomain(email) {
  if (!email || typeof email !== 'string') return '';
  const parts = email.split('@');
  return parts.length > 1 ? parts[1] : '';
}

/**
 * デプロイユーザーのドメイン情報を取得
 * @return {Object} ドメイン情報
 */
function getDeployUserDomainInfo() {
  try {
    // デプロイユーザーのメールアドレスを取得
    const deployUserEmail = getWebAppDeployerEmail();
    const deployDomain = extractDomain(deployUserEmail);
    
    // 現在のアクセスユーザーのメールアドレスを取得
    const currentUserEmail = safeGetUserEmail();
    const currentDomain = extractDomain(currentUserEmail);
    
    return {
      deployUserEmail: deployUserEmail,
      deployDomain: deployDomain,
      currentUserEmail: currentUserEmail,
      currentDomain: currentDomain,
      isDomainMatch: deployDomain === currentDomain,
      isValidAccess: deployDomain === currentDomain || deployDomain === ''
    };
  } catch (error) {
    console.error('ドメイン情報の取得に失敗:', error);
    return {
      deployUserEmail: '',
      deployDomain: '',
      currentUserEmail: '',
      currentDomain: '',
      isDomainMatch: false,
      isValidAccess: false,
      error: 'ドメイン情報の取得に失敗しました'
    };
  }
}

/**
 * ウェブアプリをデプロイしたユーザーのメールアドレスを取得
 * @return {string} デプロイユーザーのメールアドレス
 */
function getWebAppDeployerEmail() {
  try {
    // ウェブアプリの所有者を取得（通常はデプロイユーザー）
    return Session.getEffectiveUser().getEmail();
  } catch (error) {
    console.error('デプロイユーザーのメール取得に失敗:', error);
    // フォールバックとして、スクリプトの所有者を試行
    try {
      return Session.getActiveUser().getEmail();
    } catch (fallbackError) {
      console.error('フォールバック取得も失敗:', fallbackError);
      return '';
    }
  }
}

/**
 * データベース（スプレッドシート）にユーザーを編集者として追加
 * @param {string} userEmail - 追加するユーザーのメールアドレス
 * @return {boolean} 成功した場合true
 */
function addUserToDatabaseEditors(userEmail) {
  try {
    debugLog(`📝 データベースに編集者を追加開始: ${userEmail}`);
    
    // データベーススプレッドシートを取得
    const props = PropertiesService.getScriptProperties();
    const dbId = props.getProperty('DATABASE_ID') || props.getProperty('USER_DATABASE_ID');
    
    if (!dbId) {
      console.error('データベースIDが見つかりません');
      return false;
    }
    
    // データベースファイルを取得
    const dbFile = DriveApp.getFileById(dbId);
    
    // ユーザーを編集者として追加
    dbFile.addEditor(userEmail);
    debugLog(`✅ データベースに編集者として追加しました: ${userEmail}`);
    
    // 少し待機してアクセス権限が反映されるのを待つ
    Utilities.sleep(1000);
    
    return true;
    
  } catch (error) {
    console.error('データベースへの編集者追加に失敗:', error);
    debugLog(`❌ データベース編集者追加失敗: ${userEmail}: ${error.message}`);
    return false;
  }
}

/**
 * 新規登録前のデータベースアクセス権限を確認・付与
 * @param {string} userEmail - ユーザーのメールアドレス
 * @return {boolean} アクセス可能かどうか
 */
function ensureDatabaseAccess(userEmail) {
  try {
    debugLog(`🔐 API経由でのデータベースアクセス確認: ${userEmail}`);
    
    // 簡単な接続テストを実行
    const testResult = callDatabaseApi('ping', { test: true });
    
    if (testResult && (testResult.success || testResult.status === 'ok')) {
      debugLog(`✅ API経由でのデータベースアクセス成功: ${userEmail}`);
      return true;
    } else {
      debugLog(`❌ API経由でのデータベースアクセス失敗: ${userEmail}`);
      return false;
    }
    
  } catch (error) {
    console.error('API経由でのデータベースアクセス確認エラー:', error);
    debugLog(`❌ API エラー: ${error.message}`);
    
    // 401エラーの場合は、より具体的なガイダンスを提供
    if (error.message.includes('401')) {
      debugLog(`🔧 401エラー検出: Logger API設定の確認が必要`);
    }
    
    return false;
  }
}

/**
 * セキュリティ強化: ユーザー認証状態の検証
 * @return {Object} 認証状態と関連情報
 */
function verifyUserAuthentication() {
  try {
    const email = safeGetUserEmail();
    if (!email || !isValidEmail(email)) {
      return {
        authenticated: false,
        error: 'Invalid or missing email'
      };
    }
    
    return {
      authenticated: true,
      email: email,
      domain: getEmailDomain(email)
    };
  } catch (error) {
    return {
      authenticated: false,
      error: 'Authentication failed'
    };
  }
}

function isValidEmail(email) {
  if (!email || typeof email !== 'string') {
    return false;
  }
  return EMAIL_REGEX.test(email);
}

function getEmailDomain(email) {
  return (email || '').toString().split('@').pop().toLowerCase();
}

function isSameDomain(emailA, emailB) {
  if (!emailA || !emailB || typeof emailA !== 'string' || typeof emailB !== 'string') {
    return false;
  }
  if (!isValidEmail(emailA) || !isValidEmail(emailB)) {
    return false;
  }
  const domainA = getEmailDomain(emailA);
  const domainB = getEmailDomain(emailB);
  return domainA && domainB && domainA === domainB;
}

function getAdminEmails(spreadsheetId) {
  const props = PropertiesService.getScriptProperties();
  let adminEmails = [];
  
  if (spreadsheetId) {
    const str = props.getProperty(`ADMIN_EMAILS_${spreadsheetId}`);
    if (str) adminEmails = str.split(',').map(e => e.trim()).filter(Boolean);
  }
  
  if (adminEmails.length === 0) {
    const str = props.getProperty('ADMIN_EMAILS') || '';
    adminEmails = str.split(',').map(e => e.trim()).filter(Boolean);
  }
  
  // 管理者が設定されていない場合、スプレッドシートの所有者を管理者とする
  if (adminEmails.length === 0) {
    try {
      const ss = getCurrentSpreadsheet();
      const owner = ss.getOwner();
      if (owner) {
        const ownerEmail = owner.getEmail();
        if (ownerEmail) {
          debugLog('管理者未設定のため、所有者を管理者に設定:', ownerEmail);
          adminEmails = [ownerEmail];
        }
      }
    } catch (e) {
      console.error('所有者情報の取得に失敗:', e);
    }
  }
  
  return adminEmails;
}

function isUserAdmin(email) {
  const userProps = PropertiesService.getUserProperties();
  const userId = userProps.getProperty('CURRENT_USER_ID');
  const userEmail = email || safeGetUserEmail();

  debugLog('isUserAdmin check:', {
    email: email, 
    userEmail: userEmail, 
    userId: userId 
  });

  if (userId) {
    const userInfo = getUserInfo(userId);
    const spreadsheetId = userInfo && userInfo.spreadsheetId;
    const admins = getAdminEmails(spreadsheetId);
    debugLog('Multi-tenant admin check:', {
      spreadsheetId: spreadsheetId, 
      admins: admins, 
      userEmail: userEmail, 
      isAdmin: admins.includes(userEmail) 
    });
    return admins.includes(userEmail);
  }

  const admins = getAdminEmails();
  debugLog('Single-tenant admin check:', {
    admins: admins, 
    userEmail: userEmail, 
    isAdmin: admins.includes(userEmail) 
  });
  return admins.includes(userEmail);
}

function checkAdmin() {
  return isUserAdmin();
}

function publishApp(sheetName) {
  if (!checkAdmin()) {
    throw new Error('権限がありません。');
  }
  const props = PropertiesService.getScriptProperties();
  props.setProperty('IS_PUBLISHED', 'true');
  props.setProperty('PUBLISHED_SHEET_NAME', sheetName);
  return `「${sheetName}」を公開しました。`;
}

function unpublishApp() {
  if (!checkAdmin()) {
    throw new Error('権限がありません。');
  }
  const props = PropertiesService.getScriptProperties();
  props.setProperty('IS_PUBLISHED', 'false');
  if (props.deleteProperty) props.deleteProperty('PUBLISHED_SHEET_NAME');
  return 'アプリを非公開にしました。';
}

function parseReactionString(val) {
  if (!val) return [];
  return val
    .toString()
    .split(',')
    .map(s => s.trim())
    .filter(Boolean);
}

function hashTimestamp(ts) {
  var str = String(ts);
  var hash = 0;
  for (var i = 0; i < str.length; i++) {
    hash = ((hash << 5) - hash) + str.charCodeAt(i);
    hash |= 0;
  }
  return hash;
}


/**
 * 管理パネルの初期化に必要なデータを取得します。
 * @returns {object} - 現在のアクティブシート情報とシートのリスト
 */
function getAdminSettings() {
  const props = PropertiesService.getScriptProperties();
  const userProps = PropertiesService.getUserProperties();
  const userId = userProps.getProperty('CURRENT_USER_ID');
  
  let adminEmails = [];
  let appSettings = {};
  
  if (userId) {
    // マルチテナントモード
    const userInfo = getUserInfo(userId);
    if (userInfo) {
      adminEmails = getAdminEmails(userInfo.spreadsheetId);
      appSettings = getAppSettingsForUser();
    }
  } else {
    // 従来のモード
    adminEmails = getAdminEmails();
    appSettings = getAppSettingsForUser();
  }
  
  const allSheets = getSheets();
  const currentUserEmail = safeGetUserEmail();
  
  // 新しいシステム：常に利用可能な回答ボード
  // activeSheetNameがあれば利用可能、なければシート選択が必要
  const isPublished = props.getProperty('IS_PUBLISHED') === 'true';
  const activeSheetName = appSettings.activeSheetName || props.getProperty('ACTIVE_SHEET_NAME') || '';
  const isAvailable = !!activeSheetName;
  
  return {
    allSheets: allSheets,
    currentUserEmail: currentUserEmail,
    deployId: props.getProperty('DEPLOY_ID'),
    webAppUrl: getWebAppUrlEnhanced(), // 簡素化された関数を呼び出す
    adminEmails: adminEmails,
    isUserAdmin: adminEmails.includes(currentUserEmail),
    activeSheetName: activeSheetName,
    showNames: appSettings.showNames,
    showCounts: appSettings.showCounts,
    isPublished: isPublished
  };
}

/**
 * アクティブシートを指定されたシートに変更します。
 * @param {string} sheetName - アクティブにするシート名。
 */
function switchActiveSheet(sheetName) {
  const props = PropertiesService.getUserProperties();
  const userId = props.getProperty('CURRENT_USER_ID');
  
  if (!userId) {
    throw new Error('ユーザー情報が見つかりません。');
  }
  
  const userInfo = getUserInfo(userId);
  
  if (!userInfo || userInfo.adminEmail !== Session.getActiveUser().getEmail()) {
    throw new Error('権限がありません。');
  }
  
  if (!sheetName) {
    throw new Error('シート名が指定されていません。');
  }
  
  // スプレッドシートの存在確認
  const spreadsheet = SpreadsheetApp.openById(userInfo.spreadsheetId);
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`シート「${sheetName}」が見つかりません。`);
  }

  prepareSheetForBoard(sheetName);
  clearHeaderCache(sheetName);

  try {
    updateUserConfig(userId, {
      activeSheetName: sheetName,
      publishedAt: new Date().toISOString() // 公開日時を記録
    });
  } catch (error) {
    throw new Error(`設定の更新に失敗しました: ${error.message}`);
  }

  auditLog('ACTIVE_SHEET_CHANGED', userId, { sheetName, userEmail: userInfo.adminEmail });
  return `アクティブシートを「${sheetName}」に変更しました。`;
}

/**
 * アクティブシートの選択をクリアします。
 */
function clearActiveSheet() {
  const props = PropertiesService.getUserProperties();
  const userId = props.getProperty('CURRENT_USER_ID');
  
  if (!userId) {
    throw new Error('ユーザー情報が見つかりません。');
  }
  
  const userInfo = getUserInfo(userId);
  if (!userInfo || userInfo.adminEmail !== Session.getActiveUser().getEmail()) {
    throw new Error('権限がありません。');
  }
  
  // アクティブシートと公開日時をクリア
  try {
    updateUserConfig(userId, {
      activeSheetName: '',
      publishedAt: '' // 公開日時をクリア
    });
  } catch (error) {
    console.error('Failed to update user config:', error);
    throw new Error(`設定の更新に失敗しました: ${error.message}`);
  }
  
  auditLog('ACTIVE_SHEET_CLEARED', userId, { userEmail: userInfo.adminEmail });
  
  return 'アクティブシートをクリアしました。';
}


function setShowDetails(flag) {
  const props = PropertiesService.getUserProperties();
  const userId = props.getProperty('CURRENT_USER_ID');
  
  if (!userId) {
    throw new Error('ユーザー情報が見つかりません。');
  }
  
  const userInfo = getUserInfo(userId);
  if (!userInfo || userInfo.adminEmail !== Session.getActiveUser().getEmail()) {
    throw new Error('権限がありません。');
  }
  
  // ユーザー設定を更新
  try {
    updateUserConfig(userId, {
      showNames: Boolean(flag),
      showCounts: Boolean(flag),
      showDetails: Boolean(flag)
    });
  } catch (error) {
    console.error('Failed to update user config:', error);
    throw new Error(`設定の更新に失敗しました: ${error.message}`);
  }

  auditLog('SHOW_DETAILS_UPDATED', userId, { showNames: Boolean(flag), showCounts: Boolean(flag), userEmail: userInfo.adminEmail });

  return `詳細表示を${flag ? '有効' : '無効'}にしました。`;
}

function setDisplayOptions(options) {
  const props = PropertiesService.getUserProperties();
  const userId = props.getProperty('CURRENT_USER_ID');

  if (!userId) {
    throw new Error('ユーザー情報が見つかりません。');
  }

  const userInfo = getUserInfo(userId);
  if (!userInfo || userInfo.adminEmail !== Session.getActiveUser().getEmail()) {
    throw new Error('権限がありません。');
  }

  try {
    updateUserConfig(userId, {
      showNames: Boolean(options.showNames),
      showCounts: Boolean(options.showCounts)
    });
  } catch (error) {
    console.error('Failed to update user config:', error);
    throw new Error(`設定の更新に失敗しました: ${error.message}`);
  }

  auditLog('DISPLAY_OPTIONS_UPDATED', userId, { options, userEmail: userInfo.adminEmail });

  return '表示設定を更新しました。';
}



// =================================================================
// GAS Webアプリケーションのエントリーポイント
// =================================================================
function validateUserId(userId) {
  if (!userId || typeof userId !== 'string') {
    throw new Error('ユーザーIDが指定されていません。');
  }
  if (!/^[a-zA-Z0-9-_]{10,}$/.test(userId)) {
    throw new Error('無効なユーザーIDの形式です。');
  }
  return userId;
}

function doGet(e) {
  // ▼▼▼ このifブロックを挿入 ▼▼▼
  if (e && e.parameter && e.parameter.setup === 'true') {
    return handleSetupRequest();
  }
  // ▲▲▲ ここまで挿入 ▲▲▲

  try {
    e = e || {};
    const params = e.parameter || {};
    const userId = params.userId;
    const mode = params.mode;


    // ユーザーIDがない場合は登録画面を表示
    if (!userId) {
      const output = HtmlService.createTemplateFromFile('Registration').evaluate();
      applySecurityHeaders(output);
      return output
        .setTitle('StudyQuest - 新規登録')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    }
    
    // ユーザーIDを検証
    const validatedUserId = validateUserId(userId);
    
    // 最適化: ユーザー情報をキャッシュから取得
    let userInfo = getCachedUserInfo(validatedUserId);
    
    if (!userInfo) {
      try {
        userInfo = getUserInfo(validatedUserId);
        if (userInfo) {
          setCachedUserInfo(validatedUserId, userInfo);
        }
      } catch (e) {
        console.error('getUserInfo failed:', e);
        debugLog(`getUserInfo error for userId ${validatedUserId}:`, e.message);
        
        // データベースアクセスエラーの場合の詳細処理
        if (e.message.includes('ユーザーデータベースが設定されていません') || 
            e.message.includes('権限がありません')) {
          const output = HtmlService.createHtmlOutput(
            'データベースにアクセスできません。システムのセットアップが完了していない可能性があります。管理者にお問い合わせください。'
          );
          applySecurityHeaders(output);
          return output.setTitle('データベースエラー');
        }
        
        try {
          userInfo = getUserInfoInternal(validatedUserId);
        } catch (e2) {
          console.error('getUserInfoInternal also failed:', e2);
          const output = HtmlService.createHtmlOutput(
            `ユーザー情報の取得に失敗しました。エラー: ${e.message}`
          );
          applySecurityHeaders(output);
          return output.setTitle('ユーザー情報エラー');
        }
      }
    }
    if (!userInfo) {
      const output = HtmlService.createHtmlOutput('無効なユーザーIDです。');
      applySecurityHeaders(output);
      return output.setTitle('エラー');
    }

    // ユーザー設定を取得（全体で使用するため外で定義）
    let config = userInfo.configJson || {};
    
    // 新規登録直後などでconfigが空の場合のフォールバック処理
    if (!config || Object.keys(config).length === 0) {
      debugLog('Config is empty, attempting to initialize default config');
      try {
        // デフォルト設定を適用
        config = {
          activeSheetName: '',
          isPublished: false,
          showNames: true
        };
        
        // スプレッドシートが存在する場合は最初のシートを設定
        if (userInfo.spreadsheetId) {
          try {
            const spreadsheet = SpreadsheetApp.openById(userInfo.spreadsheetId);
            const sheets = spreadsheet.getSheets();
            if (sheets && sheets.length > 0) {
              config.activeSheetName = sheets[0].getName();
              debugLog(`Set default activeSheetName: ${config.activeSheetName}`);
            }
          } catch (spreadsheetError) {
            debugLog(`Failed to access spreadsheet for default config: ${spreadsheetError.message}`);
          }
        }
      } catch (configError) {
        debugLog(`Failed to initialize default config: ${configError.message}`);
        config = {};
      }
    }

    // --- 自動非公開タイマーのチェック ---
    try {
      if (config && config.publishedAt) {
        const publishedDate = new Date(config.publishedAt);
        const sixHoursLater = new Date(publishedDate.getTime() + 6 * 60 * 60 * 1000);
        if (new Date() > sixHoursLater) {
          // セキュリティ向上: 機密情報を含むログは削除
          console.log('ボードが6時間経過したため自動的に非公開になりました。');
          clearActiveSheet(); // サーバー側で非公開処理
          // ユーザーに通知するための専用ページを表示
          const template = HtmlService.createTemplateFromFile('Unpublished');
          template.message = 'この回答ボードは、公開から6時間が経過したため、安全のため自動的に非公開になりました。再度利用する場合は、管理者にご連絡ください。';
          const output = template.evaluate();
          applySecurityHeaders(output);
          return output.setTitle('公開期間が終了しました');
        }
      }
    } catch (configError) {
      console.error('Config check error:', configError);
      // 設定チェックでエラーが発生しても処理を続行
    }

    // 現在のユーザーを取得（必須）
    let viewerEmail;
    try {
      viewerEmail = safeGetUserEmail();
    } catch (e) {
      if (userInfo.configJson && userInfo.configJson.isPublished === false) {
        const template = HtmlService.createTemplateFromFile('Unpublished');
        template.userId = validatedUserId;
        template.userInfo = userInfo;
        template.userEmail = userInfo.adminEmail;
        template.isOwner = false;
        template.ownerName = userInfo.adminEmail || 'Unknown';
        template.boardUrl = `${getWebAppUrlEnhanced()}?userId=${validatedUserId}`;
        auditLog('UNPUBLISHED_ACCESS_NO_LOGIN', validatedUserId, { error: e });
        const output = template.evaluate();
        applySecurityHeaders(output);
        if (output.setSandboxMode) output.setSandboxMode(HtmlService.SandboxMode.IFRAME);
        return output.setTitle('StudyQuest - 回答ボード');
      }
      if (typeof HtmlService !== 'undefined') {
        const output = HtmlService.createHtmlOutput('認証が必要です。Googleアカウントでログインしてください。');
        applySecurityHeaders(output);
        return output.setTitle('認証エラー');
      }
      throw e;
    }
    
    // ドメインベースのアクセス制御
    if (!isSameDomain(viewerEmail, userInfo.adminEmail)) {
      auditLog('ACCESS_DENIED', validatedUserId, { viewerEmail, adminEmail: userInfo.adminEmail });
      if (typeof HtmlService !== 'undefined') {
        const output = HtmlService.createHtmlOutput('システムエラーが発生しました。管理者にお問い合わせください。');
        applySecurityHeaders(output);
        return output.setTitle('エラー');
      }
      throw new Error('権限がありません。同じドメインのユーザーのみアクセスできます。');
    }

    // 現在のコンテキストを設定
    PropertiesService.getUserProperties().setProperties({
      CURRENT_USER_ID: validatedUserId,
      CURRENT_SPREADSHEET_ID: userInfo.spreadsheetId
    });
  
    // 管理モードの場合
    if (mode === 'admin') {
      // 管理者権限のチェック
      if (!(viewerEmail === userInfo.adminEmail || isUserAdmin(viewerEmail))) {
        auditLog('ADMIN_ACCESS_DENIED', validatedUserId, { viewerEmail });
        if (typeof HtmlService !== 'undefined') {
          const output = HtmlService.createHtmlOutput('権限がありません。');
          applySecurityHeaders(output);
          return output.setTitle('アクセス拒否');
        }
        throw new Error('権限がありません。');
      }

      const template = HtmlService.createTemplateFromFile('AdminPanel');
      template.userId = validatedUserId;
      template.userInfo = userInfo;
      auditLog('ADMIN_ACCESS', validatedUserId, { viewerEmail });
      const output = template.evaluate();
      applySecurityHeaders(output);
      if (output.setSandboxMode) output.setSandboxMode(HtmlService.SandboxMode.IFRAME);
      return output.setTitle('StudyQuest - 管理パネル');
    }
  
  // 通常表示モード - 常に回答ボードを表示

  
  // 下位互換性：古いsheetNameを新しいactiveSheetNameにマッピング
  const activeSheetName = config.activeSheetName || config.sheetName || '';
  
  // デバッグ情報をログに出力
  debugLog('Config debug:', {
    userId: validatedUserId,
    configJson: config,
    activeSheetName: activeSheetName,
    hasActiveSheet: !!activeSheetName
  });
  
  // アクティブシートが設定されていない場合は、統合された管理画面を表示
  if (!activeSheetName) {
    // 管理者の場合は統合管理画面、一般ユーザーの場合は待機画面
    if (userInfo.adminEmail === viewerEmail) {
      const template = HtmlService.createTemplateFromFile('AdminPanel');
      template.userId = validatedUserId;
      template.userInfo = userInfo;
      auditLog('ADMIN_ACCESS_NO_SHEET', validatedUserId, { viewerEmail });
      const output = template.evaluate();
      applySecurityHeaders(output);
      if (output.setSandboxMode) output.setSandboxMode(HtmlService.SandboxMode.IFRAME);
      return output.setTitle('StudyQuest - 管理パネル');
    } else {
      const template = HtmlService.createTemplateFromFile('Unpublished');
      template.userId = validatedUserId;
      template.userInfo = userInfo;
      template.userEmail = userInfo.adminEmail;
      template.isOwner = false;
      template.ownerName = userInfo.adminEmail || 'Unknown';
      template.boardUrl = `${getWebAppUrlEnhanced()}?userId=${validatedUserId}`;
      auditLog('UNPUBLISHED_ACCESS', validatedUserId, { viewerEmail, isOwner: false });
      const output = template.evaluate();
      applySecurityHeaders(output);
      if (output.setSandboxMode) output.setSandboxMode(HtmlService.SandboxMode.IFRAME);
      return output.setTitle('StudyQuest - 回答ボード');
    }
  }

  // 公開されていない場合は待機画面を表示
  if (!config.isPublished && userInfo.adminEmail !== viewerEmail) {
    const template = HtmlService.createTemplateFromFile('Unpublished');
    template.userId = validatedUserId;
    template.userInfo = userInfo;
    template.userEmail = userInfo.adminEmail;
    template.isOwner = false;
    template.ownerName = userInfo.adminEmail || 'Unknown';
    template.boardUrl = `${getWebAppUrlEnhanced()}?userId=${validatedUserId}`;
    auditLog('UNPUBLISHED_ACCESS', validatedUserId, { viewerEmail, isOwner: false });
    const output = template.evaluate();
    applySecurityHeaders(output);
    if (output.setSandboxMode) output.setSandboxMode(HtmlService.SandboxMode.IFRAME);
    return output.setTitle('StudyQuest - 回答ボード');
  }
  
  // 既存のPage.htmlを使用
  const template = HtmlService.createTemplateFromFile('Page');
  let mapping = {};
  try {
    if (typeof getConfig === 'function') {
      mapping = getConfig(activeSheetName);
    }
  } catch (error) {
    console.warn('Config not found for sheet:', activeSheetName, error);
    mapping = {};
  }
  
  const isCurrentUserAdmin = userInfo.adminEmail === viewerEmail;

  const cfgShowNames = typeof config.showNames !== 'undefined' ? config.showNames : (config.showDetails || false);
  const cfgShowCounts = typeof config.showCounts !== 'undefined' ? config.showCounts : (config.showDetails || false);

  Object.assign(template, {
    showAdminFeatures: false, // 管理者であっても最初は閲覧モードで起動
    showHighlightToggle: isCurrentUserAdmin, // 管理者は閲覧モードでもハイライト可能
    showScoreSort: false, // 管理者であっても最初は閲覧モードで起動
    isStudentMode: true, // 管理者であっても最初は学生モードで起動
    isAdminUser: isCurrentUserAdmin, // 管理者権限の有無はそのまま渡す
    showCounts: cfgShowCounts,
    displayMode: cfgShowNames ? 'named' : 'anonymous',
    sheetName: activeSheetName,
    mapping: mapping,
    userId: userId,
    ownerName: userInfo.adminEmail || 'Unknown'
  });
  
    template.userEmail = viewerEmail;
    
    auditLog('VIEW_ACCESS', validatedUserId, { viewerEmail, sheetName: activeSheetName });
    
    const output = template.evaluate();
    applySecurityHeaders(output);
    return output
        .setTitle('StudyQuest - みんなのかいとうボード')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (error) {
    console.error('doGet error:', error);
    debugLog('doGet full error details:', error);
    
    if (typeof HtmlService !== 'undefined') {
      let errorMessage = 'システムエラーが発生しました。';
      
      // エラーの種類に応じて具体的なメッセージを提供
      if (error.message.includes('データベース')) {
        errorMessage = `データベースエラー: ${error.message}`;
      } else if (error.message.includes('Permission') || error.message.includes('権限')) {
        errorMessage = 'アクセス権限がありません。管理者にお問い合わせください。';
      } else if (error.message.includes('not found')) {
        errorMessage = '指定されたリソースが見つかりません。URLを確認してください。';
      } else if (error.message.includes('studyQuestSetup')) {
        errorMessage = 'システムのセットアップが完了していません。管理者にstudyQuestSetup()の実行を依頼してください。';
      } else {
        errorMessage = `エラー詳細: ${error.message}`;
      }
      
      const output = HtmlService.createHtmlOutput(
        `<div style="padding: 20px; font-family: Arial, sans-serif;">
          <h2 style="color: #d32f2f;">エラーが発生しました</h2>
          <p>${errorMessage}</p>
          <p style="font-size: 0.9em; color: #666;">
            問題が解決しない場合は、管理者にこのメッセージを共有してください。
          </p>
        </div>`
      );
      applySecurityHeaders(output);
      return output.setTitle('エラー');
    }
    throw error;
  }
}


// =================================================================
// クライアントサイドから呼び出される関数
// =================================================================

/**
 * 指定されたシートまたはアクティブシートのデータを取得します。
 * @param {string} requestedSheetName - 取得するシート名（省略時はアクティブシート）
 * @param {string} classFilter - クラスフィルター
 * @param {string} sortBy - ソート順
 */
function getPublishedSheetData(requestedSheetName, classFilter, sortBy) {
  // セキュリティ強化: 権限チェックを最初に実行（読み取り専用・高速版を使用）
  const { userId, userInfo } = validateCurrentUserReadOnly();
  
  const props = PropertiesService.getUserProperties();
  const spreadsheetId = props.getProperty('CURRENT_SPREADSHEET_ID');
  
  if (!spreadsheetId) {
    throw new Error('スプレッドシート情報が見つかりません。');
  }
  
  if (!userInfo) {
    throw new Error('無効なユーザーです。');
  }
  
  const config = userInfo.configJson || {};
  
  // 下位互換性を保ちながら、アクティブシート名を取得
  const activeSheetName = config.activeSheetName || config.sheetName || '';
  
  // リクエストされたシート名、またはアクティブシートを使用
  const targetSheetName = requestedSheetName || activeSheetName;
  
  if (!targetSheetName) {
    throw new Error('表示するシートが設定されていません。');
  }
  
  // 最適化: キャッシュされたスプレッドシートを使用
  const spreadsheet = getCachedSpreadsheet(spreadsheetId);
  const sheet = spreadsheet.getSheetByName(targetSheetName);
  
  if (!sheet) {
    throw new Error(`シート「${targetSheetName}」が見つかりません。`);
  }
  
  // 既存のgetSheetData関数のロジックを使用（スプレッドシートを指定）
  const order = sortBy || 'newest';
  const data = getSheetDataForSpreadsheet(spreadsheet, targetSheetName, classFilter, order);
  
  const cfgShowNames = typeof config.showNames !== 'undefined' ? config.showNames : (config.showDetails || false);
  const cfgShowCounts = typeof config.showCounts !== 'undefined' ? config.showCounts : (config.showDetails || false);

  return {
    sheetName: targetSheetName,
    header: data.header,
    rows: data.rows,
    showNames: cfgShowNames,
    showCounts: cfgShowCounts,
    isActiveSheet: targetSheetName === activeSheetName
  };
}

/**
 * 利用可能なシート一覧とアクティブシート情報を取得します。
 */
function getAvailableSheets() {
  // セキュリティ強化: 権限チェックを最初に実行（読み取り専用・高速版を使用）
  const { userId, userInfo } = validateCurrentUserReadOnly();
  
  const config = userInfo.configJson || {};
  const activeSheetName = config.activeSheetName || config.sheetName || '';
  
  // 最適化: キャッシュされたgetSheets()を使用
  const allSheets = getSheets();
  
  return {
    sheets: allSheets.map(sheetName => ({
      name: sheetName,
      isActive: sheetName === activeSheetName
    })),
    activeSheetName: activeSheetName
  };
}

/**
 * アクティブスプレッドシートに対応するGoogleフォームの編集リンクを取得
 * @return {Object} フォーム情報 {formUrl: string, editUrl: string, responseCount: number} または null
 */
function getActiveFormInfo() {
  try {
    const startTime = Date.now();
    debugLog('getActiveFormInfo: Starting execution');
    
    // 権限チェック（読み取り専用・高速版）
    const { userId, userInfo } = validateCurrentUserReadOnly();
    
    const props = PropertiesService.getUserProperties();
    const spreadsheetId = props.getProperty('CURRENT_SPREADSHEET_ID');
    
    if (!spreadsheetId) {
      debugLog('No active spreadsheet ID found');
      return null;
    }
    
    // キャッシュからフォーム情報を取得
    const formCacheKey = `form_info_${userId}_${spreadsheetId}`;
    
    const formInfo = getCachedData(
      formCacheKey,
      () => {
        debugLog('getActiveFormInfo: Cache miss, searching for form...');
        
        // スプレッドシートを取得
        const spreadsheet = getCurrentSpreadsheet();
        let formId = null;
        let form = null;
        
        // 方法1: スプレッドシートに直接リンクされたフォームを探す（最も高速）
        try {
          const formUrl = spreadsheet.getFormUrl();
          if (formUrl) {
            const match = formUrl.match(/\/forms\/d\/([a-zA-Z0-9-_]+)/);
            if (match) {
              formId = match[1];
              form = FormApp.openById(formId);
              debugLog('getActiveFormInfo: Found form via direct link');
            }
          }
        } catch (e) {
          debugLog('Method 1 failed (direct form link):', e.message);
        }
        
        // 方法2: ユーザー設定からフォームURLを取得（2番目に高速）
        if (!form && userInfo.configJson && userInfo.configJson.formUrl) {
          try {
            const formUrl = userInfo.configJson.formUrl;
            const match = formUrl.match(/\/forms\/d\/([a-zA-Z0-9-_]+)/);
            if (match) {
              formId = match[1];
              form = FormApp.openById(formId);
              debugLog('getActiveFormInfo: Found form via user config');
            }
          } catch (e) {
            debugLog('Method 2 failed (user config):', e.message);
          }
        }
        
        // 方法3: DriveApp検索（最も時間がかかるため最後に実行）
        if (!form) {
          try {
            debugLog('getActiveFormInfo: Attempting DriveApp search (slow)...');
            const spreadsheetFile = DriveApp.getFileById(spreadsheetId);
            const parents = spreadsheetFile.getParents();
            
            if (parents.hasNext()) {
              const folder = parents.next();
              const forms = folder.getFilesByType(DriveApp.FileType.GOOGLE_FORMS);
              const spreadsheetName = spreadsheet.getName();
              
              // 効率化：最大10個のフォームのみチェック（無限ループ防止）
              let formCount = 0;
              const maxFormsToCheck = 10;
              
              while (forms.hasNext() && !form && formCount < maxFormsToCheck) {
                const formFile = forms.next();
                const formName = formFile.getName();
                formCount++;
                
                // StudyQuestフォームのみをチェック
                if (formName.includes('StudyQuest') && 
                    (formName.includes(spreadsheetName.split(' - ')[0]) || 
                     spreadsheetName.includes(formName.split(' - ')[0]))) {
                  try {
                    formId = formFile.getId();
                    form = FormApp.openById(formId);
                    debugLog(`getActiveFormInfo: Found form via folder search (${formCount} forms checked)`);
                    break;
                  } catch (e) {
                    debugLog('Failed to open form:', formFile.getName(), e.message);
                  }
                }
              }
              
              if (formCount >= maxFormsToCheck) {
                debugLog('getActiveFormInfo: Reached max forms check limit');
              }
            }
          } catch (e) {
            debugLog('Method 3 failed (folder search):', e.message);
          }
        }
        
        if (!form) {
          debugLog('No associated form found for spreadsheet:', spreadsheetId);
          return null;
        }
        
        // 最適化: フォームIDを永続化（次回から検索をスキップ）
        if (formId && !persistedFormId) {
          props.setProperty(`FORM_ID_${spreadsheetId}`, formId);
          debugLog('Persisted form ID for future use:', formId);
        }
        
        return {
          formUrl: form.getPublishedUrl(),
          editUrl: 'https://docs.google.com/forms/d/' + formId + '/edit',
          formId: formId,
          title: form.getTitle()
        };
      },
      CACHE_CONFIG.FORM_INFO_TTL
    );
    
    // 最適化: 回答数を別途キャッシュ（更新频度が高いため）
    if (formInfo && formInfo.formId) {
      const responseCountCacheKey = `response_count_${formInfo.formId}`;
      const responseCount = getCachedData(
        responseCountCacheKey,
        () => {
          try {
            const form = FormApp.openById(formInfo.formId);
            const count = form.getResponses().length;
            debugLog(`getActiveFormInfo: Fresh response count: ${count}`);
            return count;
          } catch (e) {
            debugLog('Failed to get response count:', e.message);
            return 0;
          }
        },
        CACHE_CONFIG.RESPONSE_COUNT_TTL
      );
      
      formInfo.responseCount = responseCount;
    }
    
    const executionTime = Date.now() - startTime;
    debugLog(`getActiveFormInfo: Completed in ${executionTime}ms`);
    
    return formInfo;
    
  } catch (error) {
    debugLog('Error getting active form info:', error.message);
    return null;
  }
}

/**
 * 新しいスプレッドシートURLを追加し、アクティブシートとして設定します。
 * @param {string} spreadsheetUrl - 追加するスプレッドシートのURL
 */
function addSpreadsheetUrl(spreadsheetUrl) {
  // セキュリティ強化: 権限チェックを最初に実行
  const { userId, userInfo } = validateCurrentUser();

  // URLからスプレッドシートIDを抽出（セキュリティ強化）
  let spreadsheetId;
  
  // URL長制限チェック
  if (!spreadsheetUrl || spreadsheetUrl.length > 500) {
    throw new Error('無効なURLです。');
  }
  
  // URLサニタイゼーション
  const sanitizedUrl = spreadsheetUrl.trim();
  
  // [修正] 正規表現リテラルのスラッシュを修正し、パターンをより正確にしました。
  const urlPattern = /^https:\/\/docs\.google\.com\/spreadsheets\/d\/([a-zA-Z0-9-_]{44})(?:\/.*)?$/;
  const match = sanitizedUrl.match(urlPattern);

  // [修正] 構文エラーの原因となっていた余分な括弧とelseブロックを削除しました。
  if (match) {
    spreadsheetId = match[1];
  } else {
    // [修正] Unicodeエスケープされていたエラーメッセージを通常の文字列にしました。
    throw new Error('有効なGoogleスプレッドシートのURLではありません。正しい形式: https://docs.google.com/spreadsheets/d/...');
  }

  // スプレッドシートにアクセスできるかテスト
  try {
    const testSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheets = testSpreadsheet.getSheets();

    if (sheets.length === 0) {
      throw new Error('スプレッドシートにシートが見つかりません。');
    }

    // API経由でユーザー設定を更新：新しいスプレッドシートIDを設定
    try {
      const updateResult = updateUserViaApi(userId, {
        spreadsheetId: spreadsheetId,
        spreadsheetUrl: spreadsheetUrl
      });
      
      if (!updateResult.success) {
        throw new Error(`ユーザー更新失敗: ${updateResult.error}`);
      }
      
      debugLog(`✅ API経由でスプレッドシート情報を更新: ${spreadsheetId}`);
    } catch (updateError) {
      debugLog(`❌ API経由でのユーザー更新エラー: ${updateError.message}`);
      throw new Error(`スプレッドシート情報の更新に失敗しました: ${updateError.message}`);
    }

    // ユーザープロパティも更新
    const props = PropertiesService.getUserProperties();
    props.setProperty('CURRENT_SPREADSHEET_ID', spreadsheetId);

    // 最初のシートをアクティブシートとして設定
    const firstSheetName = sheets[0].getName();
    updateUserConfig(userId, {
      activeSheetName: firstSheetName
    });

    // 新しいシートの基本設定を自動作成
    try {
      // シートのヘッダーを取得して推測
      const firstSheet = testSpreadsheet.getSheetByName(firstSheetName);
      if (firstSheet && firstSheet.getLastRow() > 0) {
        const headers = firstSheet.getRange(1, 1, 1, firstSheet.getLastColumn()).getValues()[0];
        const guessedConfig = guessHeadersFromArray(headers);

        // 基本設定を保存
        if (typeof saveSheetConfigForSpreadsheet !== 'undefined') {
          saveSheetConfigForSpreadsheet(testSpreadsheet, firstSheetName, guessedConfig);
        }
      }
    } catch (configError) {
      // セキュリティ向上: エラー詳細を本番環境では表示しない
      console.warn('Failed to auto-create config for new sheet');
      // 設定作成に失敗してもスプレッドシート追加は成功とする
    }

    auditLog('SPREADSHEET_ADDED', userId, {
      spreadsheetId,
      spreadsheetUrl,
      firstSheetName,
      userEmail: userInfo.adminEmail
    });

    return {
      success: true,
      message: `スプレッドシート「${testSpreadsheet.getName()}」が追加され、シート「${firstSheetName}」がアクティブになりました。`,
      spreadsheetId: spreadsheetId,
      firstSheetName: firstSheetName
    };

  } catch (error) {
    console.error('Failed to access spreadsheet:', error);
    
    // より詳細なエラー情報を提供
    let errorMessage = 'スプレッドシートにアクセスできませんでした。';
    
    if (error.message.includes('Permission')) {
      errorMessage += '権限が不足している可能性があります。スプレッドシートの共有設定を確認してください。';
    } else if (error.message.includes('not found')) {
      errorMessage += 'スプレッドシートが見つかりません。URLが正しいことを確認してください。';
    } else {
      errorMessage += `詳細: ${error.message}`;
    }
    
    debugLog(`スプレッドシートアクセスエラー - URL: ${sanitizedUrl}, ID: ${spreadsheetId}, エラー: ${error.message}`);
    throw new Error(errorMessage);
  }
}

/**
 * AdminPanel用のステータス取得関数
 */
function getStatus() {
  try {
    const startTime = Date.now();
    debugLog('getStatus: Starting execution');
    
    const scriptProps = PropertiesService.getScriptProperties();
    const correctWebAppUrl = 'https://script.google.com/a/naha-okinawa.ed.jp/macros/s/AKfycbzFF3psxBRUja1DsrVDkleOGrUxar1QqxqGYwBVKmpcZybrtNddH5iKD-nbqmYWEZKK/exec';
    const correctDeployId = 'AKfycbzFF3psxBRUja1DsrVDkleOGrUxar1QqxqGYwBVKmpcZybrtNddH5iKD-nbqmYWEZKK';
    
    if (scriptProps.getProperty('WEB_APP_URL') !== correctWebAppUrl) {
      scriptProps.setProperty('WEB_APP_URL', correctWebAppUrl);
    }
    if (scriptProps.getProperty('DEPLOY_ID') !== correctDeployId) {
      scriptProps.setProperty('DEPLOY_ID', correctDeployId);
    }

    const props = PropertiesService.getUserProperties();
    const userId = props.getProperty('CURRENT_USER_ID');
    
    if (!userId) {
      debugLog('getStatus: No userId found, returning default');
      return {
        activeSheetName: '',
        allSheets: [],
        answerCount: 0,
        totalReactions: 0,
        webAppUrl: getWebAppUrlEnhanced(),
        spreadsheetUrl: '',
        showNames: false,
        showCounts: false
      };
    }
    
    debugLog(`getStatus: Processing for userId: ${userId}`);
    
    // バッチキャッシュ処理で複数データを同時取得
    const cacheKeys = [
      {
        key: `user_info_${userId}`,
        fetchFunction: () => getUserInfo(userId)
      },
      {
        key: `sheet_list_${userId}`,
        fetchFunction: () => {
          // 最適化: キャッシュされたgetSheets()を使用
          return getSheets();
        }
      }
    ];
    
    const cachedResults = getBatchCachedData(cacheKeys, CACHE_CONFIG.SPREADSHEET_DATA_TTL);
    const userInfo = cachedResults[`user_info_${userId}`];
    const allSheets = cachedResults[`sheet_list_${userId}`] || [];
    
    debugLog(`getStatus: Retrieved ${allSheets.length} sheets from cache/fetch`);
    
    const config = userInfo ? userInfo.configJson || {} : {};
    const activeSheetName = config.activeSheetName || '';
    
    // アクティブシートの情報を高速取得
    let answerCount = 0;
    let totalReactions = 0;
    
    if (activeSheetName) {
      const sheetDataCacheKey = `sheet_data_${userId}_${activeSheetName}`;
      
      const sheetData = getCachedData(
        sheetDataCacheKey,
        () => {
          try {
            // 最適化: キャッシュされたスプレッドシートを使用
            const spreadsheetId = userInfo.spreadsheetId;
            const spreadsheet = getCachedSpreadsheet(spreadsheetId) || getCurrentSpreadsheet();
            const sheet = spreadsheet.getSheetByName(activeSheetName);
            if (!sheet) return { answerCount: 0, totalReactions: 0 };
            
            const lastRow = sheet.getLastRow();
            const count = lastRow > 1 ? lastRow - 1 : 0;
            
            debugLog(`getStatus: Sheet ${activeSheetName} has ${count} answers`);
            return { answerCount: count, totalReactions: 0 };
          } catch (error) {
            debugLog('Sheet data fetch error:', error);
            return { answerCount: 0, totalReactions: 0 };
          }
        },
        CACHE_CONFIG.SPREADSHEET_DATA_TTL
      );
      
      answerCount = sheetData.answerCount;
      totalReactions = sheetData.totalReactions;
    }
    
    const result = {
      activeSheetName: activeSheetName,
      allSheets: allSheets,
      answerCount: answerCount,
      totalReactions: totalReactions,
      webAppUrl: getWebAppUrlEnhanced(),
      spreadsheetUrl: userInfo ? userInfo.spreadsheetUrl || '' : '',
      showNames: typeof config.showNames !== 'undefined' ? config.showNames : (config.showDetails || false),
      showCounts: typeof config.showCounts !== 'undefined' ? config.showCounts : (config.showDetails || false)
    };
    
    const executionTime = Date.now() - startTime;
    debugLog(`getStatus: Completed in ${executionTime}ms`);
    
    return result;
  } catch (error) {
    console.error('getStatus error:', error);
    throw new Error('ステータスの取得に失敗しました: ' + error.message);
  }
}

/**
 * AdminPanel用のシート切り替え関数（エイリアス）
 */
function switchToSheet(sheetName) {
  return switchActiveSheet(sheetName);
}


// =================================================================
// 内部処理関数
// =================================================================
function getSheets() {
  const { userId, userInfo } = validateCurrentUserReadOnly();
  const spreadsheetId = userInfo.spreadsheetId;
  
  // シート一覧をキャッシュから取得（最適化）
  return getCachedData(
    `sheet_list_${spreadsheetId}`,
    () => {
      try {
        debugLog('Fetching fresh sheet list...');
        const spreadsheet = getCachedSpreadsheet(spreadsheetId) || getCurrentSpreadsheet();
        const allSheets = spreadsheet.getSheets();
        const visibleSheets = allSheets.filter(sheet => !sheet.isSheetHidden());
        const filtered = visibleSheets.filter(sheet => {
          const name = sheet.getName();
          return name !== 'Config';
        });
        return filtered.map(sheet => sheet.getName());
      } catch (error) {
        handleError('getSheets', error);
        return [];
      }
    },
    CACHE_CONFIG.SHEET_LIST_TTL
  );
}

function getSheetHeaders(sheetName) {
  const sheet = getCurrentSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`シート '${sheetName}' が見つかりません。`);
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headerRow.map(v => (v !== null && v !== undefined) ? String(v) : '');
}

/**
 * シートのヘッダー情報と自動設定を同時に返す新しい関数
 * フロントエンドでの重複したヘッダー推測ロジックを削除し、バックエンドで統一処理
 */
function getSheetHeadersWithAutoConfig(sheetName) {
  try {
    debugLog('Getting headers and auto config for sheet:', sheetName);
    
    const headers = getSheetHeaders(sheetName);
    if (!headers || headers.length === 0) {
      return {
        status: 'error',
        message: 'ヘッダー情報を取得できませんでした。'
      };
    }
    
    const autoConfig = guessHeadersFromArray(headers);
    
    debugLog('Headers:', headers);
    debugLog('Auto config:', autoConfig);
    
    return {
      status: 'success',
      headers: headers,
      autoConfig: autoConfig
    };
  } catch (error) {
    console.error('Error in getSheetHeadersWithAutoConfig:', error);
    return handleError('getSheetHeadersWithAutoConfig', error, true);
  }
}

/**
 * ヘッダー配列から設定項目を推測する関数
 *
 * Keywords cover common variants (e.g. コメント vs 回答) and include
 * both full/half-width characters. Header strings are normalized by
 * trimming/collapsing whitespace and converting full-width ASCII before
 * searching.
 *
 * @param {Array} headers - ヘッダー名の配列
 * @returns {Object} 推測された設定オブジェクト
 */
function guessHeadersFromArray(headers) {
  // Normalize header strings: convert full-width ASCII, collapse whitespace and lowercase
  const normalize = (str) => String(str || '')
    .replace(/[\uFF01-\uFF5E]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0))
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();

  const normalizedHeaders = headers.map(h => ({ original: String(h), normalized: normalize(h) }));

  // Find the first header that contains any of the provided keywords
  const find = (keys) => {
    const keyNorm = keys.map(k => normalize(k));
    const found = normalizedHeaders.find(h => keyNorm.some(k => h.normalized.includes(k)));
    return found ? found.original : '';
  };
  
  debugLog('Available headers:', headers);
  
  // Googleフォーム特有のヘッダー構造に対応
  const isGoogleForm = normalizedHeaders.some(h =>
    h.normalized.includes(normalize('タイムスタンプ')) ||
    h.normalized.includes(normalize('メールアドレス'))
  );
  
  let question = '';
  let answer = '';
  let reason = '';
  let name = '';
  let classHeader = '';
  
  if (isGoogleForm) {
    debugLog('Detected Google Form response sheet');
    
    // Googleフォームの一般的な構造: タイムスタンプ, メールアドレス, [質問1], [質問2], ...
    const nonMetaHeaders = normalizedHeaders.filter(h => {
      const hStr = h.normalized;
      return !hStr.includes(normalize('タイムスタンプ')) &&
             !hStr.includes('timestamp') &&
             !hStr.includes(normalize('メールアドレス')) &&
             !hStr.includes('email');
    });
    
    debugLog('Non-meta headers:', nonMetaHeaders.map(h => h.original));
    
    // 重複ヘッダーの検出とログ出力
    const headerCounts = {};
    nonMetaHeaders.forEach(h => {
      headerCounts[h.original] = (headerCounts[h.original] || 0) + 1;
    });
    const duplicateHeaders = Object.keys(headerCounts).filter(header => headerCounts[header] > 1);
    if (duplicateHeaders.length > 0) {
      debugLog('Duplicate headers detected:', duplicateHeaders);
    }
    
    // より柔軟な推測ロジック
    for (let i = 0; i < nonMetaHeaders.length; i++) {
      const { original: header, normalized: hStr } = nonMetaHeaders[i];
      
      // 質問を含む長いテキストを質問ヘッダーとして推測
      if (!question && (hStr.includes('だろうか') || hStr.includes('ですか') || hStr.includes('でしょうか') || hStr.length > 20)) {
        question = header;
        // 同じ内容が複数列にある場合、2番目の同じヘッダーを回答用として使用
        if (i + 1 < nonMetaHeaders.length && nonMetaHeaders[i + 1].original === header) {
          answer = header; // 問題文と回答は同じヘッダー名
          debugLog('Found duplicate header for question and answer:', header);
          i++; // 次のヘッダーをスキップ
          continue;
        }
      }
      
      // 回答・意見に関するヘッダー
      if (!answer && (hStr.includes('回答') || hStr.includes('答え') || hStr.includes('意見') || hStr.includes('考え'))) {
        answer = header;
      }
      
      // 理由に関するヘッダー
      if (!reason && (hStr.includes('理由') || hStr.includes('詳細') || hStr.includes('説明'))) {
        reason = header;
      }
      
      // 名前に関するヘッダー
      if (!name && (hStr.includes('名前') || hStr.includes('氏名') || hStr.includes('学生'))) {
        name = header;
      }
      
      // クラスに関するヘッダー
      if (!classHeader && (hStr.includes('クラス') || hStr.includes('組') || hStr.includes('学級'))) {
        classHeader = header;
      }
    }
    
    // フォールバック: まだ回答が見つからない場合
    if (!answer && nonMetaHeaders.length > 0) {
      // 2回出現するヘッダーがある場合、それを問題文・回答として使用
      const duplicateHeader = duplicateHeaders.length > 0 ? duplicateHeaders[0] : null;
      if (duplicateHeader) {
        if (!question) question = duplicateHeader;
        answer = duplicateHeader;
        debugLog('Using duplicate header for question and answer:', duplicateHeader);
      } else {
        // 最初の非メタヘッダーを回答として使用
        answer = nonMetaHeaders[0].original;
        debugLog('Using first non-meta header as answer:', answer);
      }
    }
    
  } else {
    // 通常のシート用の推測ロジック
    // Keyword lists include common variants and full/half-width characters
    question = find(['質問', '質問内容', '問題', '問い', 'お題', 'question', 'question text', 'q', 'Ｑ']);
    answer = find(['回答', '解答', '答え', '返答', 'answer', 'a', 'Ａ', '意見', 'コメント', 'opinion', 'response']);
    reason = find(['理由', '訳', 'わけ', '根拠', '詳細', '説明', 'comment', 'why', 'reason', 'detail']);
    name = find(['名前', '氏名', 'name', '学生', 'student', 'ペンネーム', 'ニックネーム', 'author']);
    classHeader = find(['クラス', 'class', 'classroom', '組', '学級', '班', 'グループ']);
  }
  
  // 問題文ヘッダーと回答ヘッダーは同じデータを使用
  const headerForBoth = answer || question || '';
  
  const guessed = {
    questionHeader: headerForBoth,
    answerHeader: headerForBoth,
    reasonHeader: reason,
    nameHeader: name,
    classHeader: classHeader
  };
  
  debugLog('Guessed headers:', guessed);
  
  // 最終フォールバック: 何も推測できない場合
  if (!question && !answer && headers.length > 0) {
    debugLog('No specific headers found, using positional mapping');
    
    // タイムスタンプとメールを除外して最初の列を回答として使用
    const usableHeaders = normalizedHeaders
      .filter(h => {
        const hStr = h.normalized;
        return !hStr.includes(normalize('タイムスタンプ')) &&
               !hStr.includes('timestamp') &&
               !hStr.includes(normalize('メールアドレス')) &&
               !hStr.includes('email') &&
               h.normalized !== '';
      })
      .map(h => h.original);
    
    if (usableHeaders.length > 0) {
      // 問題文ヘッダーと回答ヘッダーは同じデータを使用
      guessed.questionHeader = usableHeaders[0];
      guessed.answerHeader = usableHeaders[0];
      if (usableHeaders.length > 1) guessed.reasonHeader = usableHeaders[1];
      if (usableHeaders.length > 2) guessed.nameHeader = usableHeaders[2];
      if (usableHeaders.length > 3) guessed.classHeader = usableHeaders[3];
    }
  }
  
  return guessed;
}

/**
 * Minimal board data builder used in tests
 */
function buildBoardData(sheetName) {
  const sheet = getCurrentSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`指定されたシート「${sheetName}」が見つかりません。`);

  const values = sheet.getDataRange().getValues();
  if (values.length < 1) return { header: '', entries: [] };

  const cfgFunc = (typeof global !== 'undefined' && global.getConfig)
    ? global.getConfig
    : (typeof getConfig === 'function' ? getConfig : null);
  const cfg = cfgFunc ? cfgFunc(sheetName) : {};

  const answerHeader = cfg.answerHeader || cfg.questionHeader || COLUMN_HEADERS.OPINION;
  const reasonHeader = cfg.reasonHeader || COLUMN_HEADERS.REASON;
  const classHeader = cfg.classHeader || COLUMN_HEADERS.CLASS;
  const nameHeader = cfg.nameHeader || '';

  const required = [answerHeader, reasonHeader, classHeader, COLUMN_HEADERS.EMAIL];
  if (nameHeader) required.push(nameHeader);

  const headerIndices = findHeaderIndices(values[0], required);

  const entries = values.slice(1).map(row => ({
    answer: row[headerIndices[answerHeader]] || '',
    reason: row[headerIndices[reasonHeader]] || '',
    name: nameHeader ? row[headerIndices[nameHeader]] || '' : '',
    class: row[headerIndices[classHeader]] || ''
  }));

  return { header: cfg.questionHeader || answerHeader, entries: entries };
}

function getSheetData(sheetName, classFilter, sortBy) {
  try {
    const sheet = getCurrentSpreadsheet().getSheetByName(sheetName);
    if (!sheet) throw new Error(`指定されたシート「${sheetName}」が見つかりません。`);

    const allValues = sheet.getDataRange().getValues();
    if (allValues.length < 1) return { header: "シートにデータがありません", rows: [] };

    const cfgFunc = (typeof global !== 'undefined' && global.getConfig)
      ? global.getConfig
      : (typeof getConfig === 'function' ? getConfig : null);
    
    let cfg = {};
    try {
      cfg = cfgFunc ? cfgFunc(sheetName) : {};
    } catch (configError) {
      debugLog(`Config not found for sheet ${sheetName} in getSheetData, creating default config:`, configError.message);
      
      // ヘッダーから自動推測して基本設定を作成
      const headers = allValues[0] || [];
      const guessedConfig = guessHeadersFromArray(headers);
      
      // 少なくとも1つのヘッダーが推測できた場合、デフォルト設定を保存
      if (guessedConfig.questionHeader || guessedConfig.answerHeader) {
        try {
          saveSheetConfig(sheetName, guessedConfig);
          cfg = guessedConfig;
          debugLog('Auto-created and saved config for sheet in getSheetData:', sheetName, cfg);
        } catch (saveError) {
          console.warn('Failed to save auto-created config in getSheetData:', saveError.message);
          cfg = guessedConfig; // 保存に失敗してもメモリ上のconfigは使用
        }
      }
    }
    
    const answerHeader = cfg.answerHeader || cfg.questionHeader || COLUMN_HEADERS.OPINION;
    const reasonHeader = cfg.reasonHeader || COLUMN_HEADERS.REASON;
    const classHeaderGuess = cfg.classHeader || COLUMN_HEADERS.CLASS;
    const nameHeader = cfg.nameHeader || '';

    const sheetHeaders = allValues[0].map(h => String(h || '').replace(/\s+/g,''));
    const normalized = header => String(header || '').replace(/\s+/g,'');

    const required = [
      COLUMN_HEADERS.EMAIL,
      answerHeader,
      reasonHeader,
      COLUMN_HEADERS.TIMESTAMP,
      COLUMN_HEADERS.UNDERSTAND,
      COLUMN_HEADERS.LIKE,
      COLUMN_HEADERS.CURIOUS,
      COLUMN_HEADERS.HIGHLIGHT
    ];
    if (nameHeader && sheetHeaders.includes(normalized(nameHeader))) {
      required.push(nameHeader);
    }

    if (sheetHeaders.includes(normalized(classHeaderGuess))) {
      required.splice(1,0,classHeaderGuess);
    }

    const headerIndices = findHeaderIndices(allValues[0], required);

    const classHeader = sheetHeaders.includes(normalized(classHeaderGuess)) ? classHeaderGuess : '';
    const dataRows = allValues.slice(1);
    const userEmail = safeGetUserEmail();
    const isAdmin = isUserAdmin(userEmail);

    const rows = dataRows.map((row, index) => {
      if (classHeader && classFilter && classFilter !== 'すべて') {
        const className = row[headerIndices[classHeader]];
        if (className !== classFilter) return null;
      }
      const email = row[headerIndices[COLUMN_HEADERS.EMAIL]];
      const opinion = row[headerIndices[answerHeader]];

      if (email && opinion) {
        const understandArr = parseReactionString(row[headerIndices[COLUMN_HEADERS.UNDERSTAND]]);
        const likeArr = parseReactionString(row[headerIndices[COLUMN_HEADERS.LIKE]]);
        const curiousArr = parseReactionString(row[headerIndices[COLUMN_HEADERS.CURIOUS]]);
        const reason = row[headerIndices[reasonHeader]] || '';
        const highlightVal = row[headerIndices[COLUMN_HEADERS.HIGHLIGHT]];
        const timestampRaw = row[headerIndices[COLUMN_HEADERS.TIMESTAMP]];
        const timestamp = timestampRaw ? new Date(timestampRaw).toISOString() : '';
        const likes = likeArr.length;
        const baseScore = reason.length;
        const likeMultiplier = 1 + (likes * SCORING_CONFIG.LIKE_MULTIPLIER_FACTOR);
        const totalScore = baseScore * likeMultiplier;
        let name = '';
        if (nameHeader && row[headerIndices[nameHeader]]) {
          name = sanitizeInput(row[headerIndices[nameHeader]]);
        } else if (email) {
          name = sanitizeInput(email.split('@')[0]);
        }
        return {
          rowIndex: index + 2,
          timestamp: timestamp,
          name: name, // 常に名前を含める（クライアント側で制御）
          email: email, // 常にemailを含める（クライアント側で制御）
          class: sanitizeInput(row[headerIndices[classHeader]] || '未分類'),
          opinion: sanitizeInput(opinion),
          reason: sanitizeInput(reason),
          reactions: {
            UNDERSTAND: { count: understandArr.length, reacted: userEmail ? understandArr.includes(userEmail) : false },
            LIKE: { count: likeArr.length, reacted: userEmail ? likeArr.includes(userEmail) : false },
            CURIOUS: { count: curiousArr.length, reacted: userEmail ? curiousArr.includes(userEmail) : false }
          },
          highlight: highlightVal === true || String(highlightVal).toLowerCase() === 'true',
          score: totalScore
        };
      }
      return null;
    }).filter(Boolean);

    if (sortBy === 'newest') {
      rows.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    } else if (sortBy === 'random') {
      rows.sort((a, b) => hashTimestamp(a.timestamp) - hashTimestamp(b.timestamp));
    } else {
      rows.sort((a, b) => b.score - a.score);
    }
    const header = cfg.questionHeader || answerHeader;
    return { header: header, rows: rows };
  } catch(e) {
    handleError(`getSheetDataForSpreadsheet for ${sheetName}`, e);
  }
}

/**
 * 指定したスプレッドシートからシートデータを取得します。
 * @param {SpreadsheetApp.Spreadsheet} spreadsheet 対象スプレッドシート
 * @param {string} sheetName シート名
 * @param {string} classFilter クラスフィルター
 * @param {string} sortBy ソート順
 */
function getSheetDataForSpreadsheet(spreadsheet, sheetName, classFilter, sortBy) {
  try {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) throw new Error(`指定されたシート「${sheetName}」が見つかりません。`);

    const allValues = sheet.getDataRange().getValues();
    if (allValues.length < 1) return { header: "シートにデータがありません", rows: [] };

    const cfgFunc = (typeof global !== 'undefined' && global.getConfig)
      ? global.getConfig
      : (typeof getConfig === 'function' ? getConfig : null);

    let cfg = {};
    try {
      cfg = cfgFunc ? cfgFunc(sheetName) : {};
    } catch (configError) {
      debugLog(`Config not found for sheet ${sheetName} in getSheetDataForSpreadsheet, creating default config:`, configError.message);

      const headers = allValues[0] || [];
      const guessedConfig = guessHeadersFromArray(headers);

      if (guessedConfig.questionHeader || guessedConfig.answerHeader) {
        try {
          saveSheetConfig(sheetName, guessedConfig);
          cfg = guessedConfig;
          debugLog('Auto-created and saved config for sheet in getSheetDataForSpreadsheet:', sheetName, cfg);
        } catch (saveError) {
          console.warn('Failed to save auto-created config in getSheetDataForSpreadsheet:', saveError.message);
          cfg = guessedConfig;
        }
      }
    }

    const answerHeader = cfg.answerHeader || cfg.questionHeader || COLUMN_HEADERS.OPINION;
    const reasonHeader = cfg.reasonHeader || COLUMN_HEADERS.REASON;
    const classHeaderGuess = cfg.classHeader || COLUMN_HEADERS.CLASS;
    const nameHeader = cfg.nameHeader || '';

    const sheetHeaders = allValues[0].map(h => String(h || '').replace(/\s+/g,''));
    const normalized = header => String(header || '').replace(/\s+/g,'');

    const required = [
      COLUMN_HEADERS.EMAIL,
      answerHeader,
      reasonHeader,
      COLUMN_HEADERS.TIMESTAMP,
      COLUMN_HEADERS.UNDERSTAND,
      COLUMN_HEADERS.LIKE,
      COLUMN_HEADERS.CURIOUS,
      COLUMN_HEADERS.HIGHLIGHT
    ];
    if (nameHeader && sheetHeaders.includes(normalized(nameHeader))) {
      required.push(nameHeader);
    }

    if (sheetHeaders.includes(normalized(classHeaderGuess))) {
      required.splice(1,0,classHeaderGuess);
    }

    const headerIndices = findHeaderIndices(allValues[0], required);

    const classHeader = sheetHeaders.includes(normalized(classHeaderGuess)) ? classHeaderGuess : '';
    const dataRows = allValues.slice(1);
    const userEmail = safeGetUserEmail();

    const rows = dataRows.map((row, index) => {
      if (classHeader && classFilter && classFilter !== 'すべて') {
        const className = row[headerIndices[classHeader]];
        if (className !== classFilter) return null;
      }
      const email = row[headerIndices[COLUMN_HEADERS.EMAIL]];
      const opinion = row[headerIndices[answerHeader]];

      if (email && opinion) {
        const understandArr = parseReactionString(row[headerIndices[COLUMN_HEADERS.UNDERSTAND]]);
        const likeArr = parseReactionString(row[headerIndices[COLUMN_HEADERS.LIKE]]);
        const curiousArr = parseReactionString(row[headerIndices[COLUMN_HEADERS.CURIOUS]]);
        const reason = row[headerIndices[reasonHeader]] || '';
        const highlightVal = row[headerIndices[COLUMN_HEADERS.HIGHLIGHT]];
        const timestampRaw = row[headerIndices[COLUMN_HEADERS.TIMESTAMP]];
        const timestamp = timestampRaw ? new Date(timestampRaw).toISOString() : '';
        const likes = likeArr.length;
        const baseScore = reason.length;
        const likeMultiplier = 1 + (likes * SCORING_CONFIG.LIKE_MULTIPLIER_FACTOR);
        const totalScore = baseScore * likeMultiplier;
        let name = '';
        if (nameHeader && row[headerIndices[nameHeader]]) {
          name = row[headerIndices[nameHeader]];
        } else if (email) {
          name = email.split('@')[0];
        }
        return {
          rowIndex: index + 2,
          timestamp: timestamp,
          name: name,
          email: email,
          class: row[headerIndices[classHeader]] || '未分類',
          opinion: opinion,
          reason: reason,
          reactions: {
            UNDERSTAND: { count: understandArr.length, reacted: userEmail ? understandArr.includes(userEmail) : false },
            LIKE: { count: likeArr.length, reacted: userEmail ? likeArr.includes(userEmail) : false },
            CURIOUS: { count: curiousArr.length, reacted: userEmail ? curiousArr.includes(userEmail) : false }
          },
          highlight: highlightVal === true || String(highlightVal).toLowerCase() === 'true',
          score: totalScore
        };
      }
      return null;
    }).filter(Boolean);

    if (sortBy === 'newest') {
      rows.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    } else if (sortBy === 'random') {
      rows.sort((a, b) => hashTimestamp(a.timestamp) - hashTimestamp(b.timestamp));
    } else {
      rows.sort((a, b) => b.score - a.score);
    }

    const header = cfg.questionHeader || answerHeader;
    return { header: header, rows: rows };
  } catch(e) {
    handleError(`getSheetDataForSpreadsheet for ${sheetName}`, e);
  }
}


function addReaction(rowIndex, reactionKey, sheetName) {
  // セキュリティ強化: 権限チェックを最初に実行
  try {
    const { userId, userInfo } = validateCurrentUser();
    // レート制限チェック
    checkRateLimit('addReaction', userInfo.adminEmail);
  } catch (error) {
    return { status: 'error', message: error.message };
  }
  
  if (!rowIndex || !reactionKey || !COLUMN_HEADERS[reactionKey]) {
    return { status: 'error', message: '無効なパラメータです。' };
  }
  
  // Use LockService only in production environment
  const lock = (typeof LockService !== 'undefined') ? LockService.getScriptLock() : null;
  try {
    if (lock && !lock.tryLock(TIME_CONSTANTS.LOCK_WAIT_MS)) {
      return { status: 'error', message: '他のユーザーが操作中です。しばらく待ってから再試行してください。' };
    }
    const userEmail = safeGetUserEmail();
    if (!userEmail) {
      return { status: 'error', message: 'ログインしていないため、操作できません。' };
    }
    
    // マルチテナント対応の修正
    const spreadsheet = getCurrentSpreadsheet();
    const settings = getAppSettingsForUser();
    const targetSheet = sheetName || settings.activeSheetName;
    const sheet = spreadsheet.getSheetByName(targetSheet);
    
    if (!sheet) throw new Error(`シート '${targetSheet}' が見つかりません。`);

  const reactionHeaders = REACTION_KEYS.map(k => COLUMN_HEADERS[k]);
  const headerIndices = getHeaderIndices(targetSheet);
  const startCol = headerIndices[reactionHeaders[0]] + 1;
  const reactionRange = sheet.getRange(rowIndex, startCol, 1, REACTION_KEYS.length);
  const values = reactionRange.getValues()[0];
  const lists = {};
  REACTION_KEYS.forEach((k, idx) => {
    lists[k] = { arr: parseReactionString(values[idx]) };
  });

    const wasReacted = lists[reactionKey].arr.includes(userEmail);
    Object.keys(lists).forEach(key => {
      const idx = lists[key].arr.indexOf(userEmail);
      if (idx > -1) lists[key].arr.splice(idx, 1);
    });
    if (!wasReacted) {
      lists[reactionKey].arr.push(userEmail);
    }
    reactionRange.setValues([
      REACTION_KEYS.map(k => lists[k].arr.join(','))
    ]);

    const reactions = REACTION_KEYS.reduce((obj, k) => {
      obj[k] = {
        count: lists[k].arr.length,
        reacted: lists[k].arr.includes(userEmail)
      };
      return obj;
    }, {});
    return { status: 'ok', reactions: reactions };
  } catch (error) {
    return handleError('addReaction', error, true);
  } finally {
    if (lock) {
      try { lock.releaseLock(); } catch (e) {}
    }
  }
}

function toggleHighlight(rowIndex, sheetName, userObject = null) {
  // セキュリティ強化: 権限チェックと rate limiting
  if (!checkAdmin()) {
    return { status: 'error', message: '権限がありません。' };
  }

  // ユーザーメール取得（ログチェック用）
  const userEmail = safeGetUserEmail();
  if (!userEmail) {
    return { status: 'error', message: 'ログインしていないため、操作できません。' };
  }

  // レート制限チェック
  try {
    checkRateLimit('toggleHighlight', userEmail);
  } catch (error) {
    return { status: 'error', message: error.message };
  }

  debugLog('toggleHighlight processing', { rowIndex: rowIndex, sheetName: sheetName, userEmail: userEmail });

  const lock = (typeof LockService !== 'undefined') ? LockService.getScriptLock() : null;
  try {
    if (lock) {
      lock.waitLock(TIME_CONSTANTS.LOCK_WAIT_MS);
    }
    const settings = getAppSettingsForUser();
    const targetSheet = sheetName || settings.activeSheetName;
    const sheet = getCurrentSpreadsheet().getSheetByName(targetSheet);
    if (!sheet) throw new Error(`シート '${targetSheet}' が見つかりません。`);

    const headerIndices = getHeaderIndices(targetSheet);
    let highlightColIndex = headerIndices[COLUMN_HEADERS.HIGHLIGHT];
    
    // ハイライト列が存在しない場合は追加する
    if (highlightColIndex === undefined || highlightColIndex === -1) {
      prepareSheetForBoard(targetSheet);
      // 再度ヘッダーインデックスを取得
      const updatedHeaderIndices = getHeaderIndices(targetSheet);
      highlightColIndex = updatedHeaderIndices[COLUMN_HEADERS.HIGHLIGHT];
      
      if (highlightColIndex === undefined || highlightColIndex === -1) {
        throw new Error('ハイライト列が見つかりません。スプレッドシートにハイライト列が存在することを確認してください。');
      }
    }
    
    const colIndex = highlightColIndex + 1;

    const cell = sheet.getRange(rowIndex, colIndex);
    const current = !!cell.getValue();
    const newValue = !current;
    cell.setValue(newValue);
    debugLog('toggleHighlight updated', { rowIndex: rowIndex, sheetName: targetSheet, highlight: newValue });
    return { status: 'ok', highlight: newValue };
  } catch (error) {
    console.error('toggleHighlight failed:', error);
    return secureHandleError('toggleHighlight', error, true);
  } finally {
    if (lock) {
      lock.releaseLock();
    }
  }
}



function getAppSettingsForUser() {
  const props = PropertiesService.getUserProperties();
  const userId = props.getProperty('CURRENT_USER_ID');
  
  if (!userId) {
    // 従来の動作のフォールバック
    return {
      activeSheetName: '',
      showNames: false,
      showCounts: false
    };
  }
  
  // マルチテナント時
  const userInfo = getUserInfo(userId);
  if (!userInfo || !userInfo.configJson) {
    return {
      activeSheetName: '',
      showNames: false,
      showCounts: false
    };
  }
  
  // 下位互換性：古いsheetNameを新しいactiveSheetNameにマッピング
  const activeSheetName = userInfo.configJson.activeSheetName || userInfo.configJson.sheetName || '';
  
  return {
    activeSheetName: activeSheetName,
    showNames: typeof userInfo.configJson.showNames !== 'undefined' ? userInfo.configJson.showNames : (userInfo.configJson.showDetails || false),
    showCounts: typeof userInfo.configJson.showCounts !== 'undefined' ? userInfo.configJson.showCounts : (userInfo.configJson.showDetails || false)
  };
}

function getHeaderIndices(sheetName) {
  const cache = (typeof CacheService !== 'undefined') ? CacheService.getScriptCache() : null;
  const cacheKey = `headers_${sheetName}`;
  
  try {
    if (cache) {
      const cached = cache.get(cacheKey);
      if (cached) {
        return JSON.parse(cached);
      }
    }
  } catch (e) {
    console.error('Cache retrieval failed:', e);
  }
  
  const sheet = getCurrentSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`シート '${sheetName}' が見つかりません。`);
  
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const indices = findHeaderIndices(headerRow, [
    COLUMN_HEADERS.TIMESTAMP,
    COLUMN_HEADERS.UNDERSTAND,
    COLUMN_HEADERS.LIKE,
    COLUMN_HEADERS.CURIOUS,
    COLUMN_HEADERS.HIGHLIGHT
  ]);
  
  try {
    if (cache) {
      cache.put(cacheKey, JSON.stringify(indices), 21600); // 6 hours cache
    }
  } catch (e) {
    console.error('Cache storage failed:', e);
  }
  
  return indices;
}

function clearHeaderCache(sheetName) {
  const cache = (typeof CacheService !== 'undefined') ? CacheService.getScriptCache() : null;
  if (cache) {
    try {
      cache.remove(`headers_${sheetName}`);
    } catch (e) {
      console.error('Cache removal failed:', e);
    }
  }
}


function prepareSheetForBoard(sheetName) {
  const sheet = getCurrentSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`シート '${sheetName}' が見つかりません。`);
  const headers = getSheetHeaders(sheetName);
  let lastCol = headers.length;
  const required = [
    COLUMN_HEADERS.TIMESTAMP,
    COLUMN_HEADERS.UNDERSTAND,
    COLUMN_HEADERS.LIKE,
    COLUMN_HEADERS.CURIOUS,
    COLUMN_HEADERS.HIGHLIGHT
  ];
  required.forEach(h => {
    if (headers.indexOf(h) === -1) {
      sheet.insertColumnAfter(lastCol);
      lastCol++;
      sheet.getRange(1, lastCol).setValue(h);
      headers.push(h);
    }
  });
  clearHeaderCache(sheetName);
}

function findHeaderIndices(sheetHeaders, requiredHeaders) {
  const indices = {};
  const normalized = sheetHeaders.map(h => (typeof h === 'string' ? h.replace(/\s+/g, '') : h));
  const missingHeaders = [];
  requiredHeaders.forEach(name => {
    const idx = normalized.indexOf(name.replace(/\s+/g, ''));
    if (idx !== -1) { indices[name] = idx; } else { missingHeaders.push(name); }
  });
  if (missingHeaders.length > 0) {
    throw new Error(`必須ヘッダーが見つかりません: [${missingHeaders.join(', ')}]`);
  }
  return indices;
}

function saveWebAppUrl(url) {
  const props = PropertiesService.getScriptProperties();
  props.setProperties({ WEB_APP_URL: (url || '').trim() });
}

function getUrlOrigin(url) {
  const match = String(url).match(/^(https?:\/\/[^/]+)/);
  return match ? match[1] : '';
}

function extractDeployIdFromUrl(url) {
  if (!url) return null;
  
  // script.google.com/macros/s/{DEPLOY_ID}/exec形式からDEPLOY_IDを抽出
  const macrosMatch = url.match(/\/macros\/s\/([a-zA-Z0-9_-]+)/);
  if (macrosMatch) {
    return macrosMatch[1];
  }
  
  // script.googleusercontent.com形式のURLの場合、URL全体から推測を試行
  if (/script\.googleusercontent\.com/.test(url)) {
    // URLのハッシュ部分やパラメータからDEPLOY_IDの可能性があるものを探す
    const hashMatch = url.match(/#gid=([a-zA-Z0-9_-]+)/);
    if (hashMatch && hashMatch[1].length > 10) {
      return hashMatch[1];
    }
  }
  
  return null;
}

// 古いgetWebAppUrl関数は削除されました。getWebAppUrlEnhanced()を使用してください。



function getWebAppUrlEnhanced() {
  const props = PropertiesService.getScriptProperties();
  let stored = (props.getProperty('WEB_APP_URL') || '').trim();
  if (stored) {
    return stored;
  }
  try {
    if (typeof ScriptApp !== 'undefined') {
      return ScriptApp.getService().getUrl();
    }
  } catch (e) {
    // ScriptAppが利用できない環境（例: テスト環境）の場合
    return '';
  }
  return '';
}

// レガシー互換のためのエイリアス関数
function getWebAppUrl() {
  return getWebAppUrlEnhanced();
}

function saveDeployId(id) {
  const props = PropertiesService.getScriptProperties();
  const cleanId = (id || '').trim();
  
  // DEPLOY_ID形式の検証
  if (cleanId && !/^[a-zA-Z0-9_-]{10,}$/.test(cleanId)) {
    console.warn('Invalid DEPLOY_ID format:', cleanId);
    throw new Error('無効なDEPLOY_ID形式です');
  }
  
  if (cleanId) {
    props.setProperties({ DEPLOY_ID: cleanId });
    debugLog('Saved DEPLOY_ID:', cleanId);
    
    // DEPLOY_IDが設定された後、WebAppURLを再評価
    const currentUrl = getWebAppUrlEnhanced();
    debugLog('Updated WebApp URL after DEPLOY_ID save:', currentUrl);
    
    const storedUrl = props.getProperty('WEB_APP_URL');
    if (storedUrl && /script\.googleusercontent\.com/.test(storedUrl)) {
      props.setProperties({ WEB_APP_URL: getWebAppUrlEnhanced() });
    }
  } else {
    console.warn('Cannot save empty DEPLOY_ID');
  }
}

// =================================================================
// マルチテナント管理機能
// =================================================================

/**
 * 新規ユーザーを登録
 */
function checkRateLimit(action, userEmail) {
  try {
    const key = `rateLimit_${action}_${userEmail}`;
    const cache = (typeof CacheService !== 'undefined') ? CacheService.getScriptCache() : null;
    const attempts = cache ? parseInt(cache.get(key) || '0') : 0;
    
    if (attempts > 10) { // 10 attempts per hour
      throw new Error('レート制限に達しました。しばらく待ってから再試行してください。');
    }
    
    if (cache) {
      cache.put(key, String(attempts + 1), 3600);
    }
  } catch (error) {
    // Cache service error - continue without rate limiting
    console.warn('Rate limiting failed:', error);
  }
}

function createStudyQuestFormFromTemplate(userEmail, userId) {
  try {
    // DriveAppの利用可能性をチェック
    if (typeof DriveApp === 'undefined' || typeof FormApp === 'undefined' || typeof SpreadsheetApp === 'undefined') {
      throw new Error('Google Drive API, Forms API, または Sheets API がこの環境で利用できません。');
    }

    // Logger APIからテンプレートIDを取得
    const loggerApiUrl = PropertiesService.getScriptProperties().getProperty(LOGGER_API_URL_KEY);
    if (!loggerApiUrl) {
      throw new Error('Logger API URL が設定されていません。');
    }

    const templateIds = getTemplateIds(loggerApiUrl);
    if (!templateIds.formId || !templateIds.spreadsheetId) {
      console.log('テンプレートが見つからないため、従来の方法でフォームを作成します。');
      return createStudyQuestForm(userEmail, userId);
    }

    // ユーザーフォルダを取得
    const userFolder = getUserFolder(userEmail);
    
    // ファイル名用の日時文字列を生成
    const timestamp = new Date();
    const dateTimeString = timestamp.toISOString().replace(/[:.]/g, '-').slice(0, 19);
    const userDisplayName = userEmail.split('@')[0];
    
    // フォームを複製
    const templateFormFile = DriveApp.getFileById(templateIds.formId);
    const formName = `StudyQuest - みんなの回答ボード - ${userDisplayName} - ${dateTimeString}`;
    const duplicatedFormFile = templateFormFile.makeCopy(formName, userFolder);
    const form = FormApp.openById(duplicatedFormFile.getId());
    
    // スプレッドシートを複製
    const templateSpreadsheetFile = DriveApp.getFileById(templateIds.spreadsheetId);
    const spreadsheetName = `StudyQuest - みんなの回答ボード - 回答データ - ${userDisplayName} - ${dateTimeString}`;
    const duplicatedSpreadsheetFile = templateSpreadsheetFile.makeCopy(spreadsheetName, userFolder);
    const spreadsheet = SpreadsheetApp.openById(duplicatedSpreadsheetFile.getId());
    
    // フォームの説明を更新（ユーザー固有の内容に変更）
    const customDescription = `
📚 みんなの回答ボード 📚

このフォームで、みんなで学び合う楽しい時間を始めましょう！

🌟 身につけたい力 🌟
• 考える力（自分の頭で考えてみよう）
• 伝える力（相手にわかりやすく説明しよう）
• 相手を思う気持ち（みんなの気持ちを大切にしよう）

💻 パソコン・タブレットを使うときの約束 💻
• 個人情報は書かない
• 友達を傷つける言葉は使わない
• 分からないことは先生に聞く

✨ 上手に答えるポイント ✨
• 具体的に書く（例：「面白い」ではなく「どこが面白いか」）
• 理由も書く（「なぜそう思うか」）
• 相手の立場になって考える

あなたの素晴らしい考えを聞かせてください！
`;
    
    form.setDescription(customDescription);
    
    // 確認メッセージを設定
    let confirmationMessage = `
🎉 回答ありがとうございました！🎉

あなたの考えがみんなの学びにつながります。

✨ 他の人の回答も見てみよう！ ✨
みんなの回答ボード: ${generateStudentBoardUrl(userId)}

📝 もっと詳しく知りたいときは先生に聞いてみてね！

✨ あなたの学びを大切に ✨
`;
    
    form.setConfirmationMessage(confirmationMessage);
    
    // フォームとスプレッドシートを再連携
    form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheet.getId());
    
    // 作成者に編集権限を付与
    duplicatedFormFile.addEditor(userEmail);
    duplicatedSpreadsheetFile.addEditor(userEmail);
    
    // ドメイン共有を設定（組織内のみ）
    try {
      duplicatedFormFile.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.VIEW);
      duplicatedSpreadsheetFile.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.VIEW);
    } catch (e) {
      console.warn('ドメイン共有設定に失敗しました（権限不足の可能性）:', e.message);
    }
    
    // スプレッドシートの追加設定
    prepareSpreadsheetForStudyQuest(spreadsheet);
    
    return {
      formId: form.getId(),
      formUrl: form.getPublishedUrl(),
      spreadsheetId: spreadsheet.getId(),
      spreadsheetUrl: spreadsheet.getUrl(),
      editFormUrl: form.getEditUrl()
    };
    
  } catch (error) {
    console.error('テンプレートからのフォーム作成エラー:', error);
    console.log('フォールバック: 従来の方法でフォームを作成します。');
    return createStudyQuestForm(userEmail, userId);
  }
}

function generateStudentBoardUrl(userId) {
  if (!userId) {
    return '';
  }
  
  try {
    const webAppUrl = getWebAppUrlEnhanced();
    if (webAppUrl) {
      return `${webAppUrl}?userId=${userId}`;
    }
  } catch (e) {
    console.warn('学生ボードURL生成に失敗しました:', e.message);
  }
  
  return '';
}

function getTemplateIds(loggerApiUrl) {
  const scriptProps = PropertiesService.getScriptProperties();
  const formIdProp = scriptProps.getProperty('TEMPLATE_FORM_ID');
  const spreadsheetIdProp = scriptProps.getProperty('TEMPLATE_SPREADSHEET_ID');

  if (formIdProp && spreadsheetIdProp) {
    console.log('スクリプトプロパティからテンプレートIDを取得しました。');
    return { formId: formIdProp, spreadsheetId: spreadsheetIdProp };
  }

  // Fallback to external API if properties are not set
  try {
    console.log('スクリプトプロパティにテンプレートIDがないため、外部APIから取得開始:', loggerApiUrl);

    const response = UrlFetchApp.fetch(loggerApiUrl, {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify({
        action: 'getTemplateIds',
        data: {}
      }),
      muteHttpExceptions: true
    });

    console.log('API レスポンスコード:', response.getResponseCode());
    console.log('API レスポンス内容:', response.getContentText());

    if (response.getResponseCode() !== 200) {
      throw new Error(`Logger API エラー: ${response.getResponseCode()}`);
    }

    const result = JSON.parse(response.getContentText());
    console.log('パース結果:', result);

    if (result.success === true && result.data) {
      console.log('テンプレートID取得成功:', result.data);
      // Store in Script Properties for future use
      scriptProps.setProperty('TEMPLATE_FORM_ID', result.data.formId);
      scriptProps.setProperty('TEMPLATE_SPREADSHEET_ID', result.data.spreadsheetId);
      return result.data;
    } else {
      const errorMessage = result.error || 'テンプレートIDの取得に失敗しました';
      console.log('テンプレートID取得失敗:', errorMessage);
      throw new Error(errorMessage);
    }
  } catch (error) {
    console.error('テンプレートID取得エラー:', error);
    return { formId: null, spreadsheetId: null };
  }
}

function createStudyQuestForm(userEmail, userId) {
  try {
    // FormAppとDriveAppの利用可能性をチェック
    if (typeof FormApp === 'undefined' || typeof DriveApp === 'undefined') {
      throw new Error('Google Forms API または Drive API がこの環境で利用できません。');
    }
    
    // ユーザー専用フォルダを取得または作成
    const userFolder = getUserFolder(userEmail);
    
    // 新しいGoogleフォームを作成（作成日時を含む）
    const now = new Date();
    const dateTimeString = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
    const form = FormApp.create(`StudyQuest - みんなの回答ボード - ${userEmail.split('@')[0]} - ${dateTimeString}`);
    
    // メールアドレス収集を有効にし、確認済み（自動取得）に設定
    form.setCollectEmail(true);
    form.setRequireLogin(true);
    
    // 未公開メソッドを使用して「確認済み」設定を試行
    try {
      // 注意：これは未公開のメソッドです（2024年に発見された方法）
      if (typeof form.setEmailCollectionType === 'function') {
        form.setEmailCollectionType(FormApp.EmailCollectionType.VERIFIED);
        debugLog('✅ setEmailCollectionType(VERIFIED)を適用しました');
      } else {
        debugLog('⚠️ setEmailCollectionType メソッドが利用できません');
      }
    } catch (undocumentedError) {
      debugLog('⚠️ 未公開メソッド使用エラー:', undocumentedError.message);
    }
    
    // Googleフォームのメールアドレス収集設定を「確認済み」（自動取得）に変更
    try {
      // ユーザー1人1回の回答制限を設定（これにより確認済みモードが有効になる）
      form.setLimitOneResponsePerUser(true);
      
      // 回答の編集を許可
      form.setAllowResponseEdits(true);
      
      // 2024年発見の未公開メソッド情報
      debugLog('ℹ️ メールアドレス「確認済み」設定の状況:');
      debugLog('  - 未公開メソッド setEmailCollectionType(VERIFIED) を試行済み');
      debugLog('  - setCollectEmail(true) + setRequireLogin(true) で基本設定完了');
      debugLog('  - 未公開メソッドが無効な場合は手動設定が必要です');
      
      debugLog('✅ フォームのメールアドレス収集設定を完了しました');
      debugLog('setCollectEmail: true, setRequireLogin: true, setLimitOneResponsePerUser: true');
      
      // メールアドレス「確認済み」設定の手動確認手順
      debugLog('');
      debugLog('📝 【重要】メールアドレスを「確認済み」にするため、以下を確認してください：');
      debugLog('1. 作成されたフォームを開く');
      debugLog('2. 設定（⚙️）→「回答」タブ');
      debugLog('3. 「メールアドレスを収集する」で「確認済み」を選択');
      debugLog(`4. フォーム編集URL: https://docs.google.com/forms/d/${form.getId()}/edit`);
      debugLog('');
      
    } catch (e) {
      debugLog('⚠️ フォーム設定でエラー:', e.message);
    }
    
    // フォームの説明と回答後のメッセージを設定（小中学生向けの分かりやすい言葉）
    const formDescription = `
📚 みんなの回答ボードへようこそ！

🎯 【このフォームで身につけたい力】
• 考える力：「なぜそう思うのか」理由も一緒に説明する
• 伝える力：友達にもわかりやすい文章を書く  
• 相手を思う気持ち：いろいろな考えがあることを理解する

💻 【パソコンやタブレットを使うときの約束】
教室でお話しするときと同じように、やさしい気持ちで取り組みましょう。

✏️ 【上手に答えるポイント】
• 自分が思ったことを、理由と一緒に書こう
• 友達が読んで「なるほど」と思える言葉を使おう
• みんなの答えが違っても、それでいいんだよ

みんなの考えを聞くのが楽しみです！
`.trim();
    
    form.setDescription(formDescription);
    
    // userIdが提供されている場合のみボードURLを生成
    let boardUrl = '';
    if (userId) {
      try {
        const webAppUrl = getWebAppUrlEnhanced();
        if (webAppUrl) {
          boardUrl = `${webAppUrl}?userId=${userId}`;
        }
      } catch (e) {
        debugLog('ボードURL生成に失敗しましたが、フォーム作成は続行します:', e.message);
      }
    }
    
    const confirmationMessage = boardUrl 
      ? `🎉 回答ありがとうございます！\n\nあなたの大切な意見が届きました。\nみんなの回答ボードで、お友達の色々な考えも見てみましょう。\n新しい発見があるかもしれませんね！\n\n${boardUrl}`
      : '🎉 回答ありがとうございます！\n\nあなたの大切な意見が届きました。';
    
    form.setConfirmationMessage(confirmationMessage);

    // クラス入力項目を追加し、入力形式を制限
    const classItem = form.addTextItem();
    classItem.setTitle('クラス名');
    classItem.setRequired(true);

    // クラス名の形式を「G1-1」のような半角英数字とハイフンのみとする
    const pattern = '^[A-Za-z0-9]+-[A-Za-z0-9]+$';

    const helpText = "【重要】クラス名は決められた形式で入力してください。\n\n✅ 正しい例：\n• 6年1組 → 6-1\n• 5年2組 → 5-2  \n• 中1年A組 → 1-A\n• 中3年B組 → 3-B\n\n❌ 間違いの例：6年1組、6-1組、６－１\n\n※ 半角英数字とハイフン（-）のみ使用可能です";
    const textValidation = FormApp.createTextValidation()
      .setHelpText(helpText)
      .requireTextMatchesPattern(pattern)
      .build();
    classItem.setValidation(textValidation);

    // その他の項目を追加（教育的観点から子供の成長につながる説明付き）
    const nameItem = form.addTextItem();
    nameItem.setTitle('名前');
    nameItem.setHelpText('あなたの名前を入力してください。この名前は先生だけが見ることができ、みんなには表示されません。一人ひとりの意見を大切にするために必要です。');
    nameItem.setRequired(true);
    
    const answerItem = form.addParagraphTextItem();
    answerItem.setTitle('回答');
    answerItem.setHelpText('質問に対するあなたの考えを、自分の言葉で表現してください。正解はありません。あなたらしい考えや感じ方を大切にして、思考力を育てましょう。');
    answerItem.setRequired(true);
    
    const reasonItem = form.addParagraphTextItem();
    reasonItem.setTitle('理由');
    reasonItem.setHelpText('なぜそう思ったのか、根拠や理由を説明してください。論理的に考える力を身につけ、自分の意見に責任を持つ習慣を育てましょう。');
    reasonItem.setRequired(false);
    
    // フォームの回答先スプレッドシートを作成（作成日時を含む）
    const spreadsheet = SpreadsheetApp.create(`StudyQuest - みんなの回答ボード - 回答データ - ${userEmail.split('@')[0]} - ${dateTimeString}`);
    form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheet.getId());

    // ファイルをユーザー専用フォルダに移動
    const formFile = DriveApp.getFileById(form.getId());
    const spreadsheetFile = DriveApp.getFileById(spreadsheet.getId());
    
    try {
      formFile.moveTo(userFolder);
      spreadsheetFile.moveTo(userFolder);
      debugLog(`✅ フォームとスプレッドシートをユーザーフォルダに移動しました`);
    } catch (e) {
      debugLog(`⚠️ ファイル移動に失敗: ${e.message}`);
    }

    // 共有設定を自動化
    const userDomain = getEmailDomain(userEmail);
    
    // まず、作成したユーザー自身に確実にアクセス権限を付与
    try {
      formFile.addEditor(userEmail);
      spreadsheetFile.addEditor(userEmail);
      debugLog(`✅ ユーザー「${userEmail}」に編集権限を付与しました`);
    } catch (e) {
      debugLog(`⚠️ ユーザー権限付与に失敗: ${e.message}`);
    }
    
    // ドメイン共有設定（組織内での利用のため）
    if (userDomain) {
      try {
        formFile.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.EDIT);
        spreadsheetFile.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.EDIT);
        debugLog(`✅ フォームとスプレッドシートをドメイン「${userDomain}」で共有しました`);
      } catch (e) {
        debugLog(`⚠️ ドメイン共有設定に失敗: ${e.message}`);
        // ドメイン共有に失敗してもユーザー個人には権限があるので続行
      }
    }

    const sheet = spreadsheet.getSheets()[0];
    
    // フォームとスプレッドシートの連携を一度確立してからヘッダーを調整
    // これによりGoogleフォームが自動的に基本ヘッダー（タイムスタンプ、メールアドレス、その他の質問項目）を作成する
    
    
    // 現在のヘッダーを取得（Googleフォームによって自動生成されたもの）
    const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    debugLog('Current headers after form linking:', currentHeaders);
    
    // StudyQuest用の追加列を準備
    const additionalHeaders = [
      COLUMN_HEADERS.UNDERSTAND,
      COLUMN_HEADERS.LIKE,
      COLUMN_HEADERS.CURIOUS,
      COLUMN_HEADERS.HIGHLIGHT
    ];
    
    // 追加の列を既存のヘッダーの後に挿入
    const startCol = currentHeaders.length + 1;
    sheet.getRange(1, startCol, 1, additionalHeaders.length).setValues([additionalHeaders]);
    
    // ヘッダー行のフォーマット
    const allHeadersRange = sheet.getRange(1, 1, 1, currentHeaders.length + additionalHeaders.length);
    allHeadersRange.setFontWeight('bold').setBackground('#E3F2FD');
    
    // 列幅を調整
    try {
      sheet.autoResizeColumns(1, allHeadersRange.getNumColumns());
    } catch (e) {
      console.warn('Auto-resize failed:', e);
    }
    
    // Configシートも作成
    prepareSpreadsheetForStudyQuest(spreadsheet);
    
    return {
      formId: form.getId(),
      formUrl: form.getPublishedUrl(),
      spreadsheetId: spreadsheet.getId(),
      spreadsheetUrl: spreadsheet.getUrl(),
      editFormUrl: form.getEditUrl()
    };

  } catch (error) {
    console.error('Failed to create form and spreadsheet:', error);
    
    // FormAppが利用できない場合はスプレッドシートのみ作成
    if (error.message.includes('Google Forms API is not available') || error.message.includes('FormApp')) {
      console.warn('FormApp not available, creating spreadsheet only');
      return createStudyQuestSpreadsheetFallback(userEmail);
    }
    
    // 権限エラーの場合
    if (error.message.includes('permission') || error.message.includes('Permission')) {
      throw new Error('Googleフォーム作成の権限がありません。管理者にお問い合わせください。');
    }
    
    // その他のエラー詳細を含める
    throw new Error(`Googleフォームとスプレッドシートの作成に失敗しました。詳細: ${error.message}`);
  }
}

/**
   * Admin Panel用のボード作成関数
   * 現在ログイン中のユーザーの新しいボードを作成し、公開・アクティブ化まで行います。
   */
function createBoardFromAdmin() {
  if (!checkAdmin()) {
    throw new Error('権限がありません。');
  }
  try {
    const currentUserEmail = safeGetUserEmail();
    const props = PropertiesService.getUserProperties();
    const userId = props.getProperty('CURRENT_USER_ID');

    if (!userId) {
      throw new Error('ユーザーIDが見つかりません。ページをリロードして再試行してください。');
    }

    // 1. Googleフォームとスプレッドシートを作成（テンプレートから複製、共有設定、回答後URL設定済み）
    const result = createStudyQuestFormFromTemplate(currentUserEmail, userId);
    
    // 2. 作成されたスプレッドシートを現在のユーザーのメインスプレッドシートとして設定
    if (result.spreadsheetId && result.spreadsheetUrl) {
      const addResult = addSpreadsheetUrl(result.spreadsheetUrl);
      debugLog('Spreadsheet added to user:', addResult);

      // 3. 新しく作成したシートをアクティブ化・公開
      const newSheetName = addResult.firstSheetName;
      if (newSheetName) {
        // アクティブシートに設定
        switchActiveSheet(newSheetName);
        // 公開状態に設定
        updateUserConfig(userId, { isPublished: true });
        debugLog(`New board '${newSheetName}' has been created and published.`);
      }

      return {
        ...result,
        message: '新しいボードが作成され、公開されました！管理ページを更新します。',
        autoCreated: true,
        newSheetName: newSheetName
      };
    }

    throw new Error('スプレッドシートの作成または追加に失敗しました。');

  } catch (error) {
    console.error('Failed to create board from admin:', error);
    return handleError('createBoardFromAdmin', error, true);
  }
}

function createStudyQuestSpreadsheetFallback(userEmail) {
  try {
    const spreadsheet = SpreadsheetApp.create(`StudyQuest - 回答データ - ${userEmail.split('@')[0]}`);
    const sheet = spreadsheet.getActiveSheet();
    
    const headers = [
      COLUMN_HEADERS.TIMESTAMP,
      COLUMN_HEADERS.EMAIL,
      COLUMN_HEADERS.CLASS,
      COLUMN_HEADERS.NAME,
      COLUMN_HEADERS.OPINION,
      COLUMN_HEADERS.REASON,
      COLUMN_HEADERS.UNDERSTAND,
      COLUMN_HEADERS.LIKE,
      COLUMN_HEADERS.CURIOUS,
      COLUMN_HEADERS.HIGHLIGHT
    ];
    sheet.appendRow(headers);
    
    const allHeadersRange = sheet.getRange(1, 1, 1, headers.length);
    allHeadersRange.setFontWeight('bold');
    allHeadersRange.setBackground('#E3F2FD');
    
    try {
      sheet.autoResizeColumns(1, headers.length);
    } catch (e) {
      console.warn('Auto-resize failed:', e);
    }
    
    prepareSpreadsheetForStudyQuest(spreadsheet);
    
    return {
      formId: null, // フォームは作成されない
      formUrl: null,
      spreadsheetId: spreadsheet.getId(),
      spreadsheetUrl: spreadsheet.getUrl(),
      editFormUrl: null,
      message: 'Googleフォームの作成に失敗したため、スプレッドシートのみ作成しました。'
    };
  } catch (error) {
    console.error('Failed to create spreadsheet fallback:', error);
    throw new Error(`スプレッドシートの作成に失敗しました。詳細: ${error.message}`);
  }
}


function prepareSpreadsheetForStudyQuest(spreadsheet) {
  // この関数はスプレッドシートの作成のみを行い、Configシートの初期化はaddSpreadsheetUrlで行う
  // Configシートを非表示にする
  let configSheet = spreadsheet.getSheetByName('Config');
  if (configSheet) {
    configSheet.hideSheet();
  }
}

function getDatabase() {
  const props = PropertiesService.getScriptProperties();
  const dbId = props.getProperty('DATABASE_ID') || props.getProperty('USER_DATABASE_ID'); // 後方互換性
  
  if (!dbId) {
    throw new Error('ユーザーデータベースが設定されていません。studyQuestSetup()を実行してセットアップを完了してください。');
  }
  
  try {
    const database = SpreadsheetApp.openById(dbId);
    debugLog(`✅ データベースアクセス成功: ${dbId}`);
    return database;
  } catch (error) {
    console.error('Database access failed:', error);
    debugLog(`❌ データベースアクセス失敗: ${dbId}: ${error.message}`);
    
    if (error.message.includes('Permission')) {
      throw new Error('データベースへのアクセス権限がありません。管理者にお問い合わせください。');
    } else if (error.message.includes('not found')) {
      throw new Error('データベースが見つかりません。システムの再セットアップが必要な可能性があります。');
    } else {
      throw new Error(`データベースアクセスエラー: ${error.message}`);
    }
  }
}

function getUserInfo(userId) {
  try {
    debugLog(`API経由でユーザー情報を取得中: ${userId}`);
    
    const userInfo = getUserInfoViaApi(userId);
    
    if (userInfo && userInfo.success && userInfo.data) {
      // ユーザーが非アクティブまたは削除されている場合はキャッシュをクリア
      if (userInfo.data.isActive === false || userInfo.data.isActive === 'FALSE') {
        invalidateUserCache(userId);
        return null;
      }
      // configJsonが文字列の場合はオブジェクトに変換
      if (userInfo.data.configJson) {
        try {
          if (typeof userInfo.data.configJson === 'string') {
            userInfo.data.configJson = JSON.parse(userInfo.data.configJson);
          }
        } catch (e) {
          debugLog(`configJson解析エラー for user ${userId}: ${e.message}`);
          userInfo.data.configJson = {};
        }
      } else {
        userInfo.data.configJson = {};
      }
      
      // 最終アクセス日時の更新もAPI経由で実行
      updateUserViaApi(userId, { lastAccessedAt: new Date().toISOString() });
      return userInfo.data;
    }
    
    debugLog(`ユーザー情報が見つかりませんでした: ${userId}`);
    // ユーザーが見つからない場合はキャッシュをクリア
    invalidateUserCache(userId);
    return null;
    
  } catch (error) {
    debugLog(`getUserInfo API エラー: ${error.message}`);
    return null;
  }
}

function getUserInfoInternal(userId) {
  const userDb = getOrCreateMainDatabase();
  const data = userDb.getDataRange().getValues();
  const headers = data[0];
  const userIdIndex = headers.indexOf('userId');
  const adminEmailIndex = headers.indexOf('adminEmail');
  const spreadsheetIdIndex = headers.indexOf('spreadsheetId');
  const spreadsheetUrlIndex = headers.indexOf('spreadsheetUrl');
  const configJsonIndex = headers.indexOf('configJson');
  const lastAccessedAtIndex = headers.indexOf('lastAccessedAt');
  const isActiveIndex = headers.indexOf('isActive');

  for (let i = 1; i < data.length; i++) {
    if (data[i][userIdIndex] === userId) {
      const userInfo = {
        userId: data[i][userIdIndex],
        adminEmail: data[i][adminEmailIndex],
        spreadsheetId: data[i][spreadsheetIdIndex],
        spreadsheetUrl: data[i][spreadsheetUrlIndex],
        configJson: {},
        lastAccessedAt: data[i][lastAccessedAtIndex],
        isActive: data[i][isActiveIndex] === true
      };
      try {
        userInfo.configJson = JSON.parse(data[i][configJsonIndex] || '{}');
      } catch (e) {
        console.error('Failed to parse configJson for user:', userId, e);
        userInfo.configJson = {};
      }
      
      // 最終アクセス日時を更新
      userDb.getRange(i + 1, lastAccessedAtIndex + 1).setValue(new Date());
      
      return userInfo;
    }
  }
  return null;
}


function updateUserConfig(userId, newConfig) {
  try {
    // 現在の設定を取得
    const userInfo = getUserInfoViaApi(userId);
    
    if (!userInfo.success || !userInfo.data) {
      throw new Error('ユーザー情報が見つかりません');
    }
    
    let currentConfig = {};
    try {
      currentConfig = JSON.parse(userInfo.data.configJson || '{}');
    } catch (e) {
      console.error('Failed to parse existing configJson for user:', userId, e);
    }
    
    // 設定をマージ
    const updatedConfig = Object.assign({}, currentConfig, newConfig);
    
    // API経由で更新
    const updateResult = updateUserViaApi(userId, {
      configJson: JSON.stringify(updatedConfig)
    });
    
    if (!updateResult.success) {
      throw new Error(`設定更新失敗: ${updateResult.error}`);
    }
    
    debugLog(`✅ API経由で設定更新: ${userId}`, newConfig);
    return true;
    
  } catch (error) {
    debugLog(`❌ API経由での設定更新エラー: ${error.message}`);
    console.error('updateUserConfig error:', error);
    return false;
  }
}

function saveSheetConfig(sheetName, config) {
  const props = PropertiesService.getUserProperties();
  const userId = props.getProperty('CURRENT_USER_ID');
  
  if (!userId) {
    throw new Error('ユーザー情報が見つかりません。');
  }
  
  const userInfo = getUserInfo(userId);
  if (!userInfo || userInfo.adminEmail !== Session.getActiveUser().getEmail()) {
    throw new Error('権限がありません。');
  }
  
  // ユーザー設定を更新
  try {
    const currentConfig = userInfo.configJson || {};
    const updatedConfig = Object.assign({}, currentConfig, {
      sheetConfigs: {
        ...(currentConfig.sheetConfigs || {}),
        [sheetName]: config
      }
    });
    updateUserConfig(userId, updatedConfig);
  } catch (error) {
    console.error('Failed to update user config with sheet config:', error);
    throw new Error(`シート設定の保存に失敗しました: ${error.message}`);
  }
  
  auditLog('SHEET_CONFIG_SAVED', userId, { sheetName, config, userEmail: userInfo.adminEmail });
  
  return 'シート設定を保存しました。';
}

function getConfig(sheetName) {
  const props = PropertiesService.getUserProperties();
  const userId = props.getProperty('CURRENT_USER_ID');
  
  if (!userId) {
    throw new Error('ユーザー情報が見つかりません。');
  }
  
  const userInfo = getUserInfo(userId);
  if (!userInfo) {
    throw new Error('無効なユーザーです。');
  }
  
  const sheetConfigs = userInfo.configJson && userInfo.configJson.sheetConfigs;
  if (sheetConfigs && sheetConfigs[sheetName]) {
    return sheetConfigs[sheetName];
  }
  
  throw new Error(`シート「${sheetName}」の設定が見つかりません。`);
}

function auditLog(action, userId, details = {}) {
  try {
    // 新しいアーキテクチャ: 管理者向けログ記録APIを使用
    const logData = {
      timestamp: new Date().toISOString(),
      userId: userId,
      action: action,
      details: details,
      source: 'mainApp'
    };
    
    logToAdminApi(logData);
    debugLog(`監査ログ送信: ${action} for ${userId}`);
    
  } catch (e) {
    console.error('監査ログ送信に失敗:', e);
  }
}

function registerNewUser(adminEmail) {
  checkRateLimit('registerNewUser', adminEmail);
  
  // 最適化メモ: 将来の改善案
  // 1. API呼び出しの統合エンドポイント化
  // 2. フォルダ作成の非同期実行
  // 3. スプレッドシート初期設定のバッチ処理
  
  // 📝 ステップ1: API経由でのデータベースアクセス
  debugLog(`🚀 API経由で新規登録開始: ${adminEmail}`);
  debugLog(`実行ユーザー: ${Session.getActiveUser().getEmail()}`);
  
  try {
    // API経由で既存ユーザーをチェック
    debugLog(`既存ユーザーチェックを実行中: adminEmail="${adminEmail}"`);
    const existingCheck = callDatabaseApi('checkExistingUser', { adminEmail: adminEmail });
    
    debugLog(`既存ユーザーチェック結果: ${JSON.stringify(existingCheck)}`);
    
    if (existingCheck.success && existingCheck.exists && existingCheck.data) {
      // 既存ユーザーの場合はURL情報を返す
      const userData = existingCheck.data;
      
      // userIdが有効かどうかを厳密にチェック
      if (!userData.userId || 
          typeof userData.userId !== 'string' ||
          userData.userId.trim() === '' || 
          userData.userId === 'null' ||
          userData.userId === 'undefined') {
        debugLog(`既存ユーザーが見つかりましたが、userIdが無効です: "${userData.userId}" (type: ${typeof userData.userId}) - 新規登録を継続します`);
      } else {
        const base = getWebAppUrlEnhanced();
        
        debugLog(`既存ユーザーが見つかりました: userId="${userData.userId}", adminEmail="${userData.adminEmail}"`);
        
        // ユーザーコンテキストを設定
        PropertiesService.getUserProperties().setProperty('CURRENT_USER_ID', userData.userId);
        PropertiesService.getUserProperties().setProperty('CURRENT_SPREADSHEET_ID', userData.spreadsheetId);
        
        return {
          userId: userData.userId,
          spreadsheetId: userData.spreadsheetId,
          spreadsheetUrl: userData.spreadsheetUrl,
          adminUrl: base ? `${base}?userId=${userData.userId}&mode=admin` : '',
          viewUrl: base ? `${base}?userId=${userData.userId}` : '',
          message: '既存のボードが見つかりました。管理画面に移動します。',
          autoCreated: false
        };
      }
    } else {
      debugLog(`既存ユーザーが見つかりませんでした - 新規登録を継続します`);
    }
    
    debugLog(`✅ 既存ユーザーチェック完了 - 新規ユーザーです`);
    
  } catch (apiError) {
    debugLog(`❌ API経由のデータベースアクセスエラー: ${apiError.message}`);
    throw new Error(`データベースアクセスに失敗しました: ${apiError.message}`);
  }
  
  // 📝 ステップ2: 新しいユーザーIDを生成
  const userId = Utilities.getUuid();
  if (!userId || typeof userId !== 'string' || userId.trim() === '') {
    throw new Error('ユーザーID生成に失敗しました。再度お試しください。');
  }
  debugLog(`📋 ユーザーID生成完了: ${userId}`);
  
  // 📝 ステップ3: Googleフォームとスプレッドシートを作成（テンプレートから複製）
  debugLog(`📝 フォーム・スプレッドシート作成開始（テンプレート使用）: ${adminEmail}`);
  const formAndSsInfo = createStudyQuestFormFromTemplate(adminEmail, userId);
  debugLog(`✅ フォーム・スプレッドシート作成完了`);
  
  // 📝 ステップ4: API経由でデータベースに新規ユーザー情報を追加
  const userData = {
    userId: userId,
    adminEmail: adminEmail,
    spreadsheetId: formAndSsInfo.spreadsheetId,
    spreadsheetUrl: formAndSsInfo.spreadsheetUrl,
    createdAt: new Date().toISOString(),
    accessToken: '', // 未使用
    configJson: JSON.stringify({}),
    lastAccessedAt: new Date().toISOString(),
    isActive: true
  };
  
  debugLog(`💾 API経由でデータベースに新規ユーザー登録中: ${adminEmail}`);
  try {
    const createResult = createUserViaApi(userData);
    if (!createResult.success) {
      throw new Error(createResult.error || 'ユーザー作成に失敗しました');
    }
    debugLog(`✅ API経由でのデータベース登録完了`);
  } catch (createError) {
    debugLog(`❌ API経由でのユーザー作成エラー: ${createError.message}`);
    throw new Error(`ユーザー作成に失敗しました: ${createError.message}`);
  }
  
  auditLog('NEW_USER_REGISTERED', userId, { adminEmail, spreadsheetId: formAndSsInfo.spreadsheetId });
  
  // 📝 ステップ6: ユーザーコンテキストを設定（クイックアクションと同様）
  debugLog(`⚙️ ユーザーコンテキスト設定中`);
  PropertiesService.getUserProperties().setProperty('CURRENT_USER_ID', userId);
  PropertiesService.getUserProperties().setProperty('CURRENT_SPREADSHEET_ID', formAndSsInfo.spreadsheetId);
  
  // 📝 ステップ7: スプレッドシートをユーザーのスプレッドシートリストに追加
  debugLog(`📊 スプレッドシートリストに追加中`);
  const addResult = addSpreadsheetUrl(formAndSsInfo.spreadsheetUrl);
  debugLog('Spreadsheet added to user:', addResult);
  
  // 📝 ステップ8: 新しく作成したシートをアクティブ化・公開（クイックアクションと同様）
  const newSheetName = addResult.firstSheetName;
  if (newSheetName) {
    debugLog(`🔄 シートをアクティブ化・公開中: ${newSheetName}`);
    // アクティブシートに設定と公開状態を同時に設定（1回のAPI呼び出しで実行）
    updateUserConfig(userId, {
      activeSheetName: newSheetName,
      publishedAt: new Date().toISOString(),
      isPublished: true
    });
    debugLog(`✅ シートアクティブ化・公開完了: ${newSheetName}`);
  }
  
  // 📝 ステップ9: URL生成と最終レスポンス準備
  debugLog(`🔗 URL生成中`);
  const webAppUrl = getWebAppUrlEnhanced();
  debugLog('Register new user - webAppUrl:', webAppUrl);
  debugLog('Register new user - userId:', userId);
  
  const adminUrl = webAppUrl ? `${webAppUrl}?userId=${userId}&mode=admin` : '';
  const viewUrl = webAppUrl ? `${webAppUrl}?userId=${userId}` : '';
  
  debugLog('Register new user - generated URLs:', { adminUrl, viewUrl });
  debugLog(`🎉 新規登録完了: ${adminEmail} (ユーザーID: ${userId})`);
  
  return {
    userId: userId,
    spreadsheetId: formAndSsInfo.spreadsheetId,
    spreadsheetUrl: formAndSsInfo.spreadsheetUrl,
    formUrl: formAndSsInfo.formUrl,
    editFormUrl: formAndSsInfo.editFormUrl,
    adminUrl: adminUrl,
    viewUrl: viewUrl,
    newSheetName: newSheetName,
    autoCreated: true,
    message: '新しいボードが作成され、公開されました！管理画面に移動します。'
  };
}

function getSpreadsheetUrlForUser(userId) {
  const userInfo = getUserInfo(userId);
  if (userInfo && userInfo.spreadsheetUrl) {
    return userInfo.spreadsheetUrl;
  }
  throw new Error('ユーザーのスプレッドシートURLが見つかりません。');
}

function openActiveSpreadsheet() {
  const props = PropertiesService.getUserProperties();
  const userId = props.getProperty('CURRENT_USER_ID');
  if (!userId) {
    throw new Error('ユーザー情報が見つかりません。');
  }
  const userInfo = getUserInfo(userId);
  if (!userInfo || !userInfo.spreadsheetId) {
    throw new Error('アクティブなスプレッドシートが見つかりません。');
  }
  const spreadsheet = SpreadsheetApp.openById(userInfo.spreadsheetId);
  return spreadsheet.getUrl();
}

/**
 * 現在のユーザーに紐づく既存ボード情報を取得
 * @return {Object|null} 既存ボードがあればURL等を含むオブジェクト、なければnull
 */
function getExistingBoard() {
  try {
    const email = safeGetUserEmail();
    if (!email) {
      debugLog('ユーザーメールアドレスが取得できません');
      return null;
    }
    
    debugLog(`API経由で既存ボード情報を取得中: ${email}`);
    
    // 最適化: ユーザールックアップをキャッシュ
    const boardInfo = getCachedData(
      `existing_board_${email}`,
      () => getExistingBoardViaApi(email),
      CACHE_CONFIG.USER_INFO_TTL
    );
    
    if (boardInfo && boardInfo.success && boardInfo.data) {
      const userData = boardInfo.data;
      
      // userIdが有効かどうかを厳密にチェック
      if (!userData.userId || 
          typeof userData.userId !== 'string' ||
          userData.userId.trim() === '' || 
          userData.userId === 'null' ||
          userData.userId === 'undefined') {
        debugLog(`無効なuserIdが見つかりました: "${userData.userId}" (type: ${typeof userData.userId}) - 既存ボードなしとして扱います`);
        return null;
      }
      
      const base = getWebAppUrlEnhanced();
      
      return {
        userId: userData.userId,
        adminUrl: base ? `${base}?userId=${userData.userId}&mode=admin` : '',
        viewUrl: base ? `${base}?userId=${userData.userId}` : '',
        spreadsheetUrl: userData.spreadsheetUrl || ''
      };
    }
    
    debugLog('既存ボードが見つかりませんでした');
    return null;
    
  } catch (error) {
    console.error('既存ボード確認で問題が発生しました', error);
    debugLog(`getExistingBoard API エラー: ${error.message}`);
    return null;
  }
}

// =================================================================
// CommonJS exports for unit tests
// =================================================================
if (typeof module !== 'undefined') {
  module.exports = {
    addReaction,
    buildBoardData,
    checkAdmin,
    COLUMN_HEADERS,
    doGet,
    findHeaderIndices,
    getAdminSettings,
    getHeaderIndices,
    getSheetData,
    getSheetDataForSpreadsheet,
    getSheetHeaders,
    getWebAppUrl,
    getWebAppUrlEnhanced,
    guessHeadersFromArray,
    parseReactionString,
    isUserAdmin,
    prepareSheetForBoard,
    saveDeployId,
    saveSheetConfig,
    saveWebAppUrl,
    toggleHighlight,
    registerNewUser,
    getSpreadsheetUrlForUser,
    openActiveSpreadsheet,
    getExistingBoard,
    getUserFolder
  };
}

/**
 * メインのユーザーデータベースを取得または新規作成する。
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} データベースのシートオブジェクト。
 */
function getOrCreateMainDatabase() {
  const properties = PropertiesService.getScriptProperties();
  let dbId = properties.getProperty(MAIN_DB_ID_KEY);

  // まず既存の【ログデータベース】みんなの回答ボードを検索
  if (!dbId) {
    try {
      const files = DriveApp.getFilesByName('【ログデータベース】みんなの回答ボード');
      if (files.hasNext()) {
        const existingDb = files.next();
        dbId = existingDb.getId();
        properties.setProperty(MAIN_DB_ID_KEY, dbId);
        Logger.log(`既存の【ログデータベース】みんなの回答ボードを使用します。ID: ${dbId}`);
        
        const spreadsheet = SpreadsheetApp.openById(dbId);
        let userSheet = null;
        
        // Users シートがあるか確認
        try {
          userSheet = spreadsheet.getSheetByName('Users');
        } catch (e) {
          // Users シートがない場合は作成
          userSheet = spreadsheet.insertSheet('Users');
          const headers = ['userId', 'adminEmail', 'spreadsheetId', 'spreadsheetUrl', 'createdAt', 'accessToken', 'configJson', 'lastAccessedAt', 'isActive'];
          userSheet.appendRow(headers);
    sheet.setFrozenRows(1);

    // ★★★ 修正点: 新規スプレッドシートにConfigシートを作成して設定を保存 ★★★
    try {
      const guessedConfig = guessHeadersFromArray(headers);
      if (typeof saveSheetConfigForSpreadsheet !== 'undefined') {
        saveSheetConfigForSpreadsheet(spreadsheet, '最初のボード', guessedConfig);
      }
    } catch (configError) {
      console.error(`Failed to create initial config for ${spreadsheetId}: ${configError.message}`);
      // 設定作成に失敗しても、ユーザー登録処理は続行する
    }
    // ★★★ 修正ここまで ★★★

          userSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
          Logger.log('【ログデータベース】みんなの回答ボードにUsersシートを作成しました');
        }
        
        return userSheet;
      }
    } catch (e) {
      Logger.log(`既存データベース検索エラー: ${e.message}`);
    }
  }

  // プロパティに保存されたIDがある場合
  if (dbId) {
    try {
      const spreadsheet = SpreadsheetApp.openById(dbId);
      // Users シートを取得または作成
      try {
        return spreadsheet.getSheetByName('Users');
      } catch (e) {
        // Users シートがない場合は作成
        const userSheet = spreadsheet.insertSheet('Users');
        const headers = ['userId', 'adminEmail', 'spreadsheetId', 'spreadsheetUrl', 'createdAt', 'accessToken', 'configJson', 'lastAccessedAt', 'isActive'];
        userSheet.appendRow(headers);
        userSheet.setFrozenRows(1);
        userSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
        Logger.log('既存データベースにUsersシートを作成しました');
        return userSheet;
      }
    } catch (e) {
      Logger.log(`データベース(ID: ${dbId})へのアクセスに失敗。新規作成します。Error: ${e.message}`);
      properties.deleteProperty(MAIN_DB_ID_KEY);
    }
  }
  
  // 新規作成：【ログデータベース】みんなの回答ボードとして作成
  const db = SpreadsheetApp.create('【ログデータベース】みんなの回答ボード');
  dbId = db.getId();
  properties.setProperty(MAIN_DB_ID_KEY, dbId);
  
  // データベースの共有設定を追加（強化版）
  try {
    const dbFile = DriveApp.getFileById(dbId);
    const adminEmail = Session.getActiveUser().getEmail();
    const userDomain = adminEmail.split('@')[1];
    
    // Google Workspaceドメインの場合は特別な設定が必要
    try {
      // まず制限付きで編集権限を設定
      dbFile.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT);
      Logger.log(`データベースをドメイン制限付きリンク共有で編集可能に設定しました。ドメイン: ${userDomain}`);
    } catch (domainError) {
      Logger.log(`ドメイン制限設定失敗、リンク共有にフォールバック: ${domainError.message}`);
      try {
        dbFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
        Logger.log(`リンクを知っている全員に編集権限を設定しました`);
      } catch (linkError) {
        Logger.log(`リンク共有設定も失敗: ${linkError.message}`);
      }
    }
    
    // セットアップ実行ユーザーを明示的に編集者として追加
    dbFile.addEditor(adminEmail);
    Logger.log(`セットアップ実行ユーザー（${adminEmail}）をデータベースの編集者として追加しました。`);
    
    // 権限設定の確認
    Utilities.sleep(2000); // 権限反映を待つ
    
    try {
      const testAccess = dbFile.getName();
      Logger.log(`✅ データベース権限設定確認成功`);
    } catch (accessTest) {
      Logger.log(`⚠️ データベース権限設定確認失敗: ${accessTest.message}`);
      
      // 追加の権限設定試行
      try {
        dbFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
        Logger.log(`フォールバック: リンクを知っている全員に編集権限を設定しました`);
      } catch (fallbackError) {
        Logger.log(`フォールバック権限設定も失敗: ${fallbackError.message}`);
      }
    }
    
  } catch (e) {
    Logger.log(`データベース共有設定エラー: ${e.message}`);
    
    // 最小限の権限設定を試行
    try {
      const dbFile = DriveApp.getFileById(dbId);
      const adminEmail = Session.getActiveUser().getEmail();
      dbFile.addEditor(adminEmail);
      Logger.log(`最小限の権限設定: 管理者のみ追加完了`);
    } catch (minimalError) {
      Logger.log(`最小限の権限設定も失敗: ${minimalError.message}`);
    }
  }
  
  // Users シートを作成
  const userSheet = db.insertSheet('Users');
  const headers = ['userId', 'adminEmail', 'spreadsheetId', 'spreadsheetUrl', 'createdAt', 'accessToken', 'configJson', 'lastAccessedAt', 'isActive'];
  userSheet.appendRow(headers);
  userSheet.setFrozenRows(1);
  userSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  
  // デフォルトシートを削除
  const defaultSheet = db.getSheets()[0];
  if (defaultSheet.getName() === 'シート1' || defaultSheet.getName() === 'Sheet1') {
    db.deleteSheet(defaultSheet);
  }
  
  Logger.log(`【ログデータベース】みんなの回答ボードを新規作成しました。ID: ${dbId}`);
  return userSheet;
}

/**
 * 重複する古いメインデータベースを削除する関数
 */
function cleanupDuplicateDatabases() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const currentDbId = properties.getProperty(MAIN_DB_ID_KEY);
    
    Logger.log('重複データベースのクリーンアップを開始します...');
    
    // 古い「StudyQuest - メインデータベース」を検索して削除
    const oldMainDbFiles = DriveApp.getFilesByName('StudyQuest - メインデータベース');
    let deletedCount = 0;
    
    while (oldMainDbFiles.hasNext()) {
      const file = oldMainDbFiles.next();
      const fileId = file.getId();
      
      // 現在使用中のデータベース以外を削除
      if (fileId !== currentDbId) {
        try {
          file.setTrashed(true);
          deletedCount++;
          Logger.log(`古いデータベースをゴミ箱に移動しました: ${fileId}`);
        } catch (e) {
          Logger.log(`データベース削除エラー (${fileId}): ${e.message}`);
        }
      }
    }
    
    Logger.log(`クリーンアップ完了: ${deletedCount}個のファイルを削除しました`);
    return deletedCount;
    
  } catch (e) {
    Logger.log(`クリーンアップエラー: ${e.message}`);
    return 0;
  }
}

/**
 * メインデータベースにユーザーを編集者として追加する
 * @param {string} userEmail - 追加するユーザーのメールアドレス
 * @returns {boolean} 成功した場合true
 */
function addUserToMainDatabaseEditors(userEmail) {
  try {
    const properties = PropertiesService.getScriptProperties();
    const dbId = properties.getProperty(MAIN_DB_ID_KEY);
    
    if (!dbId) {
      Logger.log('メインデータベースIDが設定されていません');
      return false;
    }
    
    const dbFile = DriveApp.getFileById(dbId);
    dbFile.addEditor(userEmail);
    Logger.log(`ユーザー ${userEmail} を【ログデータベース】みんなの回答ボードの編集者として追加しました`);
    return true;
    
  } catch (e) {
    Logger.log(`メインデータベースへのユーザー追加エラー: ${e.message}`);
    return false;
  }
}

/**
 * 管理者向けログ記録APIを通じてデータベース操作を実行する
 * @param {string} action - 実行するアクション ('getUser', 'createUser', 'updateUser', etc.)
 * @param {object} data - 送信するデータ
 * @returns {object} APIからのレスポンス
 */
function callDatabaseApi(action, data = {}) {
  const apiUrl = PropertiesService.getScriptProperties().getProperty(LOGGER_API_URL_KEY);
  if (!apiUrl) {
    throw new Error('Logger APIのURLが設定されていません。セットアップを完了してください。');
  }

  const requestUser = Session.getActiveUser().getEmail();
  const effectiveUser = Session.getEffectiveUser().getEmail();
  
  const payload = {
    action: action,
    data: data,
    timestamp: new Date().toISOString(),
    requestUser: requestUser,
    effectiveUser: effectiveUser,
    userAgent: 'StudyQuest-MainApp/1.0'
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
    headers: {
      'User-Agent': 'StudyQuest-MainApp/1.0',
      'X-Requested-With': 'StudyQuest',
      'X-Request-User': requestUser
    }
  };

  debugLog(`🔗 API呼び出し開始:`);
  debugLog(`• URL: ${apiUrl}`);
  debugLog(`• Action: ${action}`);
  debugLog(`• Data: ${JSON.stringify(data)}`);
  debugLog(`• Request User: ${requestUser}`);
  debugLog(`• Effective User: ${effectiveUser}`);

  try {
    const response = UrlFetchApp.fetch(apiUrl, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    debugLog(`📡 API応答:`);
    debugLog(`• ステータスコード: ${responseCode}`);
    debugLog(`• レスポンステキスト長: ${responseText.length}文字`);
    
    if (responseCode === 200) {
      try {
        const parsedResponse = JSON.parse(responseText);
        debugLog(`✅ API呼び出し成功: ${JSON.stringify(parsedResponse)}`);
        
        // APIレベルでのエラーをチェック
        if (parsedResponse.success === false) {
          debugLog(`❌ API応答内エラー: ${parsedResponse.error || parsedResponse.message || 'Unknown error'}`);
          throw new Error(`API応答エラー: ${parsedResponse.error || parsedResponse.message || 'Unknown error'}`);
        }
        
        return parsedResponse;
      } catch (parseError) {
        debugLog(`⚠️ JSON解析失敗、テキストとして返却: ${parseError.message}`);
        debugLog(`レスポンステキスト: ${responseText}`);
        return { success: true, data: responseText };
      }
    } else {
      // 詳細なエラー情報を提供
      debugLog(`❌ API呼び出し失敗:`);
      debugLog(`• ステータス: ${responseCode}`);
      debugLog(`• レスポンス: ${responseText.substring(0, 500)}${responseText.length > 500 ? '...' : ''}`);
      
      let errorMessage = `API呼び出しが失敗しました。ステータスコード: ${responseCode}`;
      
      // 特定のエラーコードに対する詳細説明
      switch (responseCode) {
        case 401:
          errorMessage += '\n\n考えられる原因:\n';
          errorMessage += '1. Logger APIのアクセス権限設定に問題があります\n';
          errorMessage += '2. Google Workspaceの認証設定が正しくありません\n';
          errorMessage += '3. APIのデプロイ設定が間違っています\n\n';
          errorMessage += '解決方法:\n';
          errorMessage += '1. Logger APIが「executeAs: USER_DEPLOYING」で設定されているか確認\n';
          errorMessage += '2. Logger APIが「access: DOMAIN」で設定されているか確認\n';
          errorMessage += '3. Logger APIを再デプロイしてください';
          break;
        case 403:
          errorMessage += '\n\nアクセス権限が拒否されました。Logger APIのアクセス設定を確認してください。';
          break;
        case 404:
          errorMessage += '\n\nLogger APIのURLが正しくないか、デプロイされていません。';
          break;
        case 500:
          errorMessage += '\n\nLogger API側でエラーが発生しています。Logger APIのログを確認してください。';
          break;
      }
      
      if (responseText) {
        errorMessage += `\n\nサーバーレスポンス: ${responseText}`;
      }
      
      throw new Error(errorMessage);
    }
  } catch(e) {
    debugLog(`💥 API呼び出し例外: ${e.message}`);
    
    if (e.message.includes('ステータスコード:')) {
      // 既に処理済みのエラー
      throw e;
    } else {
      // ネットワークエラーなど
      const networkError = `データベースAPIへの接続に失敗しました: ${e.message}\n\n`;
      const troubleshooting = '考えられる原因:\n';
      const causes = '1. Logger APIのURLが間違っている\n';
      const causes2 = '2. ネットワーク接続に問題がある\n';
      const causes3 = '3. Logger APIがデプロイされていない\n\n';
      const solution = '解決方法:\n';
      const solution1 = '1. セットアップ画面でLogger APIのURLを再確認してください\n';
      const solution2 = '2. Logger APIが正常にデプロイされているか確認してください';
      
      throw new Error(networkError + troubleshooting + causes + causes2 + causes3 + solution + solution1 + solution2);
    }
  }
}

/**
 * API経由でユーザー情報を取得する
 * @param {string} userId - ユーザーID
 * @returns {object} ユーザー情報
 */
function getUserInfoViaApi(userId) {
  return callDatabaseApi('getUserInfo', { userId: userId });
}

/**
 * API経由で新規ユーザーを作成する
 * @param {object} userData - ユーザーデータ
 * @returns {object} 作成結果
 */
function createUserViaApi(userData) {
  return callDatabaseApi('createUser', userData);
}

/**
 * API経由でユーザー情報を更新する
 * @param {string} userId - ユーザーID
 * @param {object} updateData - 更新データ
 * @returns {object} 更新結果
 */
function updateUserViaApi(userId, updateData) {
  return callDatabaseApi('updateUser', { userId: userId, ...updateData });
}

/**
 * API経由で既存ボード情報を取得する
 * @param {string} userEmail - ユーザーメールアドレス
 * @returns {object} 既存ボード情報
 */
function getExistingBoardViaApi(userEmail) {
  debugLog(`getExistingBoardViaApi: checking for userEmail="${userEmail}"`);
  const result = callDatabaseApi('getExistingBoard', { adminEmail: userEmail });
  debugLog(`getExistingBoardViaApi: result=${JSON.stringify(result)}`);
  return result;
}

/**
 * 指定されたメタデータを管理者向けログ記録APIに送信する。
 * @param {object} metadata - 送信するログデータ。
 */
function logToAdminApi(metadata) {
  const apiUrl = PropertiesService.getScriptProperties().getProperty(LOGGER_API_URL_KEY);
  if (!apiUrl) {
    Logger.log('Logger APIのURLが設定されていないため、ログ送信をスキップしました。');
    return;
  }

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      action: 'log',
      data: metadata,
      timestamp: new Date().toISOString(),
      requestUser: Session.getActiveUser().getEmail()
    }),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(apiUrl, options);
    Logger.log(`Logger APIへの送信結果: ${response.getResponseCode()}`);
  } catch(e) {
    Logger.log(`Logger APIへの送信中にエラーが発生しました: ${e.message}`);
  }
}
