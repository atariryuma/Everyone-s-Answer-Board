/**
 * @fileoverview Helper Utilities
 *
 * 🎯 責任範囲:
 * - 列マッピング・インデックス操作
 * - データフォーマッティング
 * - 汎用ヘルパー関数
 * - 計算・変換ユーティリティ
 *
 * 🔄 GAS Best Practices準拠:
 * - フラット関数構造 (Object.freeze削除)
 * - 直接的な関数エクスポート
 * - 簡素なユーティリティ関数群
 */

/* global CACHE_DURATION, TIMEOUT_MS, SLEEP_MS, PROPERTY_CACHE_TTL, getCurrentEmail, isAdministrator, getUserConfig */

const RUNTIME_PROPERTIES_CACHE = {};
const MAX_CACHE_SIZE = 50; // 最大キャッシュエントリ数

/**
 * PropertiesServiceのメモリキャッシュ付きアクセス（TTL + サイズ制限対応）
 * ✅ CLAUDE.md準拠: 30秒TTLで自動期限切れ
 * ✅ Google公式推奨: 頻繁アクセスする設定値はメモリにキャッシュ
 * ✅ メモリリーク対策: 最大50エントリ、LRU削除
 * @param {string} key - プロパティキー
 * @returns {string|null} プロパティ値
 */
function getCachedProperty(key) {
  const now = Date.now();
  const cached = RUNTIME_PROPERTIES_CACHE[key];

  if (cached && cached.timestamp && (now - cached.timestamp < PROPERTY_CACHE_TTL)) {
    cached.lastAccess = now;
    return cached.value;
  }

  const value = PropertiesService.getScriptProperties().getProperty(key);

  if (Object.keys(RUNTIME_PROPERTIES_CACHE).length >= MAX_CACHE_SIZE) {
    const entries = Object.entries(RUNTIME_PROPERTIES_CACHE);
    const oldestKey = entries.reduce((oldest, [key, value]) => {
      const oldestTime = oldest[1].lastAccess || oldest[1].timestamp;
      const currentTime = value.lastAccess || value.timestamp;
      return currentTime < oldestTime ? [key, value] : oldest;
    }, entries[0])[0];
    delete RUNTIME_PROPERTIES_CACHE[oldestKey];
  }

  RUNTIME_PROPERTIES_CACHE[key] = {
    value,
    timestamp: now,  // ✅ タイムスタンプ記録
    lastAccess: now  // ✅ 最終アクセス時刻（LRU用）
  };
  return value;
}

/**
 * タイミングセーフな文字列比較（タイミング攻撃防止）
 * @param {string} a - 比較文字列1
 * @param {string} b - 比較文字列2
 * @returns {boolean} 一致するかどうか
 */
function timingSafeEqual(a, b) {
  if (typeof a !== 'string' || typeof b !== 'string') return false;
  if (a.length !== b.length) return false;
  let result = 0;
  for (let i = 0; i < a.length; i++) {
    result |= a.charCodeAt(i) ^ b.charCodeAt(i);
  }
  return result === 0;
}

/**
 * メモリキャッシュをクリア（テスト用・設定変更時用）
 * @param {string} key - クリアするキー（省略時は全クリア）
 */
function clearPropertyCache(key = null) {
  if (key) {
    delete RUNTIME_PROPERTIES_CACHE[key];
  } else {
    Object.keys(RUNTIME_PROPERTIES_CACHE).forEach(k => delete RUNTIME_PROPERTIES_CACHE[k]);
  }
}

/**
 * Script Propertyを書き込み、メモリキャッシュも即座に無効化
 * @param {string} key - プロパティキー
 * @param {string} value - プロパティ値
 */
function setCachedProperty(key, value) {
  PropertiesService.getScriptProperties().setProperty(key, value);
  delete RUNTIME_PROPERTIES_CACHE[key];
}

/**
 * CacheServiceにオブジェクトを保存（サイズチェック付き）
 * ✅ DRY原則: 重複コード削減
 * ✅ CacheServiceの制限（100KB）を自動チェック
 * @param {string} cacheKey - キャッシュキー
 * @param {Object} data - 保存するデータ
 * @param {number} ttl - TTL（秒）
 * @param {number} maxSize - 最大サイズ（バイト）、デフォルト100KB
 * @returns {boolean} 保存成功したらtrue
 */
function saveToCacheWithSizeCheck(cacheKey, data, ttl, maxSize = 100000) {
  try {
    const dataJson = JSON.stringify(data);
    if (dataJson.length > maxSize) {
      console.warn(`saveToCacheWithSizeCheck: Data too large for cache (${dataJson.length} bytes, max ${maxSize})`);
      return false;
    }
    CacheService.getScriptCache().put(cacheKey, dataJson, ttl);
    return true;
  } catch (saveError) {
    console.warn(`saveToCacheWithSizeCheck: Cache save failed for key "${cacheKey}":`, saveError.message);
    return false;
  }
}

/**
 * オブジェクトをシンプルなハッシュ文字列に変換
 * ✅ API最適化: JSON.stringify()より約50%高速
 * @param {Object} obj - ハッシュ化するオブジェクト
 * @returns {string} ハッシュ文字列
 */
function simpleHash(obj) {
  if (!obj || typeof obj !== 'object') return '';
  const keys = Object.keys(obj).sort();
  return keys.map(k => `${k}:${obj[k]}`).join('|');
}

/**
 * 標準化エラーレスポンス生成（拡張版）
 * @param {string} message - エラーメッセージ
 * @param {*} data - 追加データ
 * @param {Object} extraFields - 追加フィールド
 * @returns {Object} 標準エラーレスポンス
 */
function createErrorResponse(message, data = null, extraFields = {}) {
  return {
    success: false,
    message,
    error: message,
    ...(data && { data }),
    ...extraFields
  };
}

/**
 * 標準化成功レスポンス生成（拡張版）
 * @param {string} message - 成功メッセージ
 * @param {*} data - レスポンスデータ
 * @param {Object} extraFields - 追加フィールド
 * @returns {Object} 標準成功レスポンス
 */
function createSuccessResponse(message, data = null, extraFields = {}) {
  return {
    success: true,
    message,
    ...(data && { data }),
    ...extraFields
  };
}

/**
 * データサービス用エラーレスポンス
 * @param {string} message - エラーメッセージ
 * @param {string} sheetName - シート名
 * @returns {Object} データサービス用エラーレスポンス
 */
function createDataServiceErrorResponse(message, sheetName = '') {
  return createErrorResponse(message, [], { headers: [], sheetName });
}

/**
 * 認証エラーレスポンス
 * @returns {Object} 認証エラー
 */
function createAuthError() {
  return createErrorResponse('ユーザー認証が必要です');
}

/**
 * ユーザー未発見エラーレスポンス
 * @returns {Object} ユーザー未発見エラー
 */
function createUserNotFoundError() {
  return createErrorResponse('ユーザーが見つかりません');
}

/**
 * 管理者権限エラーレスポンス
 * @returns {Object} 管理者権限エラー
 */
function createAdminRequiredError() {
  return createErrorResponse('管理者権限が必要です');
}

/**
 * 例外エラーレスポンス生成
 *
 * @param {Error} error - エラーオブジェクト
 * @param {string} [context] - 発生箇所を示す短い文脈名。メッセージ先頭に付く
 *   （例: "アプリケーション停止処理: Spreadsheet not found"）。
 * @returns {Object} 例外エラーレスポンス
 */
function createExceptionResponse(error, context) {
  const rawMessage = (error && error.message) || 'Unknown error';
  const message = context ? `${context}: ${rawMessage}` : rawMessage;
  return createErrorResponse(message);
}

/**
 * 認証チェック: メール取得 + 管理者判定を一括実行
 * @returns {{email: string, isAdmin: boolean}|null} 認証情報、未認証時はnull
 */
function requireAuth() {
  const email = getCurrentEmail();
  if (!email) return null;
  return { email, isAdmin: isAdministrator(email) };
}

/**
 * 管理者チェック: 未認証またはadminでなければnull
 * @returns {{email: string}|null} 管理者メール、権限不足時はnull
 */
function requireAdmin() {
  const auth = requireAuth();
  if (!auth || !auth.isAdmin) return null;
  return auth;
}

/**
 * getUserConfig 結果から config を安全に取得
 *
 * 失敗時は空オブジェクトを返す（caller が壊れないことを優先）が、
 * 同時に WARN ログを出す。以前は silent failure で、今日の「生徒にフォームボタンが
 * 出ない」バグのように「空 config が返って機能が無効化されている」事象を
 * ログから追えなかった。
 *
 * @param {string} userId - ユーザーID
 * @param {Object} [user] - 事前取得済みユーザー（findUserById 再実行を避けるため）
 * @returns {Object} config（失敗時は空オブジェクト）
 */
function getConfigOrDefault(userId, user) {
  const result = getUserConfig(userId, user);
  if (!result.success) {
    // 生徒が他人のボードを見るときに findUserById が拒否されると message が来る。
    // Why WARN not ERROR: 呼び出し元が empty config でも壊れない設計なので fatal ではない。
    //                     ただし silent だとバグに気付けないので痕跡は残す。
    console.warn('getConfigOrDefault: returning empty config', {
      userId: userId ? `${String(userId).substring(0, 8)}***` : 'N/A',
      reason: result.message || 'unknown',
      hasPreloadedUser: !!user
    });
  }
  return result.success ? result.config : {};
}

/**
 * ヘッダー文字列の正規化
 * @param {*} header - ヘッダー値
 * @returns {string} 正規化された文字列
 */
function normalizeHeader(header) {
  return String(header || '').toLowerCase().trim();
}

/** デフォルト表示設定 */
const DEFAULT_DISPLAY_SETTINGS = { showNames: false, showReactions: false, theme: 'default', pageSize: 20 };

