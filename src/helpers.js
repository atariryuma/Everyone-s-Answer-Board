/**
 * @fileoverview 共通ヘルパー: PropertiesService キャッシュ、レスポンスビルダー、
 *   認証ショートカット、emailToShortHash（仮名化）。
 */

/* global PROPERTY_CACHE_TTL, getCurrentEmail, isAdministrator, getUserConfig */

const RUNTIME_PROPERTIES_CACHE = {};
const MAX_CACHE_SIZE = 50; // 最大キャッシュエントリ数

// PropertiesService のメモリキャッシュ付きアクセス（30秒 TTL / 最大 50 エントリ LRU）。
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

  RUNTIME_PROPERTIES_CACHE[key] = { value, timestamp: now, lastAccess: now };
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
 * CacheService に JSON を保存（CacheService の 100KB 制限を超える場合は false）。
 * @param {string} cacheKey
 * @param {Object} data
 * @param {number} ttl - 秒
 * @param {number} [maxSize=100000]
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

// オブジェクトを `key:value|key:value` 形式に変換 (JSON.stringify より約 50% 速い、cache key 用)。
function simpleHash(obj) {
  if (!obj || typeof obj !== 'object') return '';
  const keys = Object.keys(obj).sort();
  return keys.map(k => `${k}:${obj[k]}`).join('|');
}

/**
 * 純粋データの deep clone (JSON で表現できる object/array のみ)。
 * GAS V8 では structuredClone 未サポートのため、 JSON 経路を helper に集約する。
 *
 * 注意: function / Date / undefined / Set / Map / 循環参照は除去/変換される。
 * lesson JSON / config JSON のような plain data 専用。
 */
function deepClone(value) {
  if (value === null || value === undefined) return value;
  return JSON.parse(JSON.stringify(value));
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
    // data:0 / '' / false を落とさず、 明示指定 (null/undefined 以外) なら含める。
    ...(data !== null && data !== undefined ? { data } : {}),
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
    // data:0 / '' / false を落とさず、 明示指定 (null/undefined 以外) なら含める。
    ...(data !== null && data !== undefined ? { data } : {}),
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
 * 共通エラーログ。 `console.error('funcName: msg', err.message)` の繰り返しを集約。
 *
 * @param {string} funcName  - 発生関数の名前 (Cloud Logging 検索用 prefix)
 * @param {Error|*} error    - Error オブジェクト or 値 (.message があれば抽出)
 * @param {Object} [context] - 追加の構造化コンテキスト (userId, configLength 等)
 *
 * Why context: 旧コードは `console.error('foo failed:', err.message, { ... })` で
 *   構造化ログを書いていたが、 prefix 形式がバラバラ。 全 catch を logError_ に
 *   移すために context を第 3 引数で受ける形に拡張 (2026-05-18 v2797)。
 */
function logError_(funcName, error, context) {
  const msg = (error && error.message) ? error.message : error;
  if (context !== undefined) {
    console.error(funcName + ' error:', msg, context);
  } else {
    console.error(funcName + ' error:', msg);
  }
}

/**
 * 安全な JSON.parse。 parse 失敗時は fallback を返す (例外 throw しない)。
 *
 * Why: configJson / SA creds / cache 読み戻し など、 信頼できない文字列を parse する箇所で
 *   try/catch を毎回書くのは冗長。 fallback を明示する形に統一する。
 * 既存 35 箇所の JSON.parse は try/catch で十分にガードされているので強制 migration はしない。
 *
 * @param {string|*} text
 * @param {*} fallback - parse 失敗時に返す値 (既定: null)
 * @returns {*}
 */
function safeJsonParse_(text, fallback) {
  if (text == null || text === '') return (fallback === undefined ? null : fallback);
  if (typeof text === 'object') return text;
  try {
    return JSON.parse(text);
  } catch (_) {
    return (fallback === undefined ? null : fallback);
  }
}

/**
 * 管理者チェック: 未認証または admin でなければ null。
 * @returns {{email: string, isAdmin: boolean}|null}
 */
function requireAdmin() {
  const email = getCurrentEmail();
  if (!email) return null;
  if (!isAdministrator(email)) return null;
  return { email, isAdmin: true };
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

/**
 * メールアドレスを salted SHA-1 で短縮ハッシュ化（仮名化）。
 *
 * Why: 可視化モード（M1/M2）で同一児童の再投稿を「揺らぎ」として追跡する目的。
 *      送信先の wire ペイロードに生メアドを載せたくないので、不可逆な短縮IDに置換する。
 *      salt は ScriptProperty に初回利用時に自動生成されるランダム値を使用。
 *      4 bytes (8 hex chars) ≒ 約 43 億通り。クラス〜学年規模の児童識別には十分。
 *
 *      仮名化であって匿名化ではない（GDPR/個人情報保護法の用語）。
 *      DB に元メアドは残るため、教師端末では同一児童の追跡が可能。
 *      児童端末側に送る wire ペイロードは showNames=false なら名前/メアド除去 + emailHash のみ。
 *
 * @param {string} email - メールアドレス
 * @returns {string|null} 8文字hex、入力空ならnull
 */
function emailToShortHash(email) {
  if (!email || typeof email !== 'string') return null;
  const trimmed = email.trim().toLowerCase();
  if (!trimmed) return null;

  let salt = getCachedProperty('EMAIL_HASH_SALT');
  if (!salt) {
    // Why: salt 未設定なら自動生成。スクリプト全体で 1 回だけ生成、以後同じ salt を使う。
    //      乱数源は Utilities.getUuid()（GAS 提供の crypto-grade UUID）。
    // Race 対策: 初回同時呼び出しが各々別 salt を生成して last-write-wins すると、
    //   loser request の hash が将来の hash と一致せず同一生徒が 2 identity に割れる。
    //   ScriptLock + lock 内再読みで「最初に生成した salt」を全員が共有するよう直列化する。
    let lock = null;
    try {
      lock = LockService.getScriptLock();
      lock.waitLock(5000);
    } catch (lockErr) {
      lock = null; // lock 取れなくても下で生成は試みる (劣化動作)
    }
    try {
      salt = getCachedProperty('EMAIL_HASH_SALT'); // lock 取得中に他 request が確定済みかも
      if (!salt) {
        salt = Utilities.getUuid().replace(/-/g, '');
        try {
          setCachedProperty('EMAIL_HASH_SALT', salt);
        } catch (e) {
          console.warn('emailToShortHash: failed to persist salt, using ephemeral value:', e.message);
        }
      }
    } finally {
      if (lock) { try { lock.releaseLock(); } catch (_) { /* ignore */ } }
    }
  }

  try {
    const bytes = Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_1,
      salt + ':' + trimmed,
      Utilities.Charset.UTF_8
    );
    // 4 bytes → 8 hex chars
    let hex = '';
    for (let i = 0; i < 4; i++) {
      const b = bytes[i] & 0xff;
      hex += (b < 16 ? '0' : '') + b.toString(16);
    }
    return hex;
  } catch (e) {
    console.warn('emailToShortHash: digest error', e.message);
    return null;
  }
}

