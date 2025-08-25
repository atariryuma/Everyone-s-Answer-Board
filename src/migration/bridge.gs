/**
 * @fileoverview マイグレーションブリッジ - 既存コードと新サービスの橋渡し
 * 
 * このファイルは既存のコードから新しいサービスへの段階的な移行を可能にします。
 * 既存の関数名を維持しながら、内部で新しいサービスを呼び出します。
 */

// ===========================
// エラーハンドリング系のブリッジ
// ===========================

// 既存のエラーハンドラーがある場合は、それらを新しいサービスにマップ
if (typeof UnifiedErrorHandler !== 'undefined') {
  // 既存のUnifiedErrorHandlerを無効化
  UnifiedErrorHandler = null;
}

// ===========================
// キャッシュ系のブリッジ
// ===========================

// 既存のキャッシュマネージャーを新しいサービスにマップ
if (typeof cacheManager !== 'undefined') {
  // 既存のcacheManagerを上書き
  var cacheManager = {
    get: function(key, fetchFunc, options) {
      const cached = getCacheValue(key);
      if (cached !== null) return cached;
      
      if (fetchFunc) {
        const value = fetchFunc();
        setCacheValue(key, value, options?.ttl || CACHE_CONFIG.DEFAULT_TTL);
        return value;
      }
      return null;
    },
    set: function(key, value, ttl) {
      setCacheValue(key, value, ttl);
    },
    remove: function(key) {
      removeCacheValue(key);
    },
    clear: function() {
      clearAllCache();
    }
  };
}

// 既存の関数を新しいサービスにマップ
function clearExecutionUserInfoCache() {
  clearUserCache(getUserId());
}

function invalidateUserCache(userId) {
  clearUserCache(userId);
}

function performUnifiedCacheClear(userId) {
  clearUserCache(userId);
}

// ===========================
// データベース系のブリッジ
// ===========================

// 既存のfindUserById関数をマップ
function findUserById(userId) {
  return getUserById(userId);
}

// 既存のfindUserByEmail関数をマップ
function findUserByEmail(email) {
  return getUserByEmail(email);
}

// 既存のcreateOrUpdateUser関数をマップ
function createOrUpdateUser(userData) {
  const existingUser = getUserById(userData.userId);
  if (existingUser) {
    return updateUser(userData.userId, userData);
  } else {
    return createUser(userData);
  }
}

// ===========================
// セッション系のブリッジ
// ===========================

// 既存のgetActiveUserEmail関数をマップ
function getActiveUserEmail() {
  return SessionService.getActiveUserEmail();
}

// 既存のgetTemporaryActiveUserKey関数をマップ
function getTemporaryActiveUserKey() {
  return SessionService.getTemporaryActiveUserKey();
}

// ===========================
// プロパティ系のブリッジ
// ===========================

// 既存のgetScriptProperty関数をマップ
function getScriptProperty(key) {
  return Properties.getScriptProperty(key);
}

// 既存のsetScriptProperty関数をマップ
function setScriptProperty(key, value) {
  Properties.setScriptProperty(key, value);
}

// ===========================
// スプレッドシート系のブリッジ
// ===========================

// 既存のopenSpreadsheetOptimized関数をマップ
function openSpreadsheetOptimized(spreadsheetId) {
  return getSpreadsheetCached(spreadsheetId, (id) => Spreadsheet.openById(id));
}

// 既存のgetCachedSpreadsheet関数をマップ
function getCachedSpreadsheet(spreadsheetId) {
  return getSpreadsheetCached(spreadsheetId, (id) => Spreadsheet.openById(id));
}

// ===========================
// 認証系のブリッジ
// ===========================

// 既存のisSystemSetup関数をマップ
function isSystemSetup() {
  return isSystemInitialized();
}

// 既存のverifyUserAccess関数をマップ
function verifyUserAccess(userId) {
  const user = getUserInfo(userId);
  return user && user.isActive === 'true';
}

// ===========================
// Utilities系のブリッジ
// ===========================

// 既存のresilientUrlFetch関数をブリッジ（既存の関数がない場合のみ定義）
if (typeof resilientUrlFetch === 'undefined') {
  global.resilientUrlFetch = function(url, options) {
    try {
      return Http.fetch(url, options);
    } catch (error) {
      // リトライ
      Utils.sleep(1000);
      return Http.fetch(url, options);
    }
  };
}

// resilientExecutorは既存のresilientExecutor.gsで定義されているため、ここでは定義しない

// ===========================
// セキュリティ系のブリッジ
// ===========================

// サービスアカウント認証情報の取得
function getSecureServiceAccountCreds() {
  const creds = Properties.getScriptProperty('SERVICE_ACCOUNT_CREDS');
  if (!creds) {
    throw new Error('サービスアカウント認証情報が設定されていません');
  }
  return JSON.parse(creds);
}

// Sheets APIサービスの作成
function createSheetsService() {
  // 簡易実装：通常のSpreadsheetAppを返す
  return {
    spreadsheets: {
      get: function(spreadsheetId) {
        return Spreadsheet.openById(spreadsheetId);
      },
      values: {
        get: function(spreadsheetId, range) {
          const spreadsheet = Spreadsheet.openById(spreadsheetId);
          const sheet = spreadsheet.getSheetByName(range.split('!')[0]);
          return { values: sheet.getDataRange().getValues() };
        },
        update: function(spreadsheetId, range, valueRange) {
          const spreadsheet = Spreadsheet.openById(spreadsheetId);
          const sheet = spreadsheet.getSheetByName(range.split('!')[0]);
          sheet.getRange(range.split('!')[1]).setValues(valueRange.values);
          return { updatedCells: valueRange.values.length };
        }
      }
    }
  };
}

// ===========================
// ロック系のブリッジ
// ===========================

// 既存のwithLock関数をマップ
function withLock(func, timeout) {
  return Lock.withLock(func, timeout);
}

// ===========================
// 設定系のブリッジ
// ===========================

// デバッグモードの確認
function isDebugMode() {
  return shouldEnableDebugMode();
}

// システム管理者の確認
function isSystemAdmin() {
  return isDeployUser();
}

// ===========================
// その他のユーティリティ
// ===========================

// 現在のWebアプリURLを取得
function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

// タイムスタンプをフォーマット
function formatTimestamp(date) {
  return Utils.formatDate(date || new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
}

// オブジェクトをJSON文字列に変換（エラーハンドリング付き）
function safeStringify(obj) {
  try {
    return JSON.stringify(obj);
  } catch (error) {
    return String(obj);
  }
}

// JSON文字列をオブジェクトに変換（エラーハンドリング付き）
function safeParse(str) {
  try {
    return JSON.parse(str);
  } catch (error) {
    return {};
  }
}