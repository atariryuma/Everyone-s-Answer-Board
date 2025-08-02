/**
 * @fileoverview URL管理 - GAS互換版
 */

// デバッグログ関数の定義（テスト環境対応）
if (typeof debugLog === 'undefined') {
  function debugLog(message, ...args) {
    console.log('[DEBUG]', message, ...args);
  }
}

if (typeof errorLog === 'undefined') {
  function errorLog(message, ...args) {
    console.error('[ERROR]', message, ...args);
  }
}

if (typeof warnLog === 'undefined') {
  function warnLog(message, ...args) {
    console.warn('[WARN]', message, ...args);
  }
}

if (typeof infoLog === 'undefined') {
  function infoLog(message, ...args) {
    console.log('[INFO]', message, ...args);
  }
}

// URL管理の定数
const URL_CACHE_KEY = 'WEB_APP_URL';
const URL_CACHE_TTL = 21600; // 6時間

/**
 * 余分なクォートを除去してURLを正規化します。
 * @param {string} url 処理対象のURL
 * @returns {string} 正規化されたURL
 */
function normalizeUrlString(url) {
  if (!url || typeof url !== 'string') {
    return url;
  }

  const cleaned = String(url).trim();

  if ((cleaned.startsWith('"') && cleaned.endsWith('"')) ||
      (cleaned.startsWith("'") && cleaned.endsWith("'"))) {
    cleaned = cleaned.slice(1, -1);
  }

  return cleaned.replace(/\\"/g, '"').replace(/\\'/g, "'");
}

/**
 * WebアプリのURLを取得（最適化版）
 * @returns {string} WebアプリURL
 */
function computeWebAppUrl() {
  try {
    let url = ScriptApp.getService().getUrl();
    if (!url) {
      warnLog('ScriptApp.getService().getUrl()がnullを返しました');
      return getFallbackUrl(); // Fallback if URL is null
    }

    // 末尾のスラッシュを除去
    url = url.replace(/\/$/, '');

    // Production URL should always be returned directly
    infoLog('✅ WebAppURL取得完了:', url);
    return url;

  } catch (e) {
    errorLog('WebアプリURL取得エラー:', e.message);
    return getFallbackUrl(); // Fallback on error
  }
}


function getWebAppUrlCached() {
  try {
    var webAppUrl = cacheManager.get(URL_CACHE_KEY, function() {
      debugLog('🔍 WebAppURL キャッシュmiss - 新規生成');
      return computeWebAppUrl();
    }, {
      ttl: 3600,
      enableMemoization: true
    });

    webAppUrl = normalizeUrlString(webAppUrl);

    if (webAppUrl) {
      infoLog('✅ 有効なキャッシュURL取得:', webAppUrl);
      return webAppUrl;
    }

    warnLog('⚠️ キャッシュURL無効、直接生成に切り替え');
    var fallbackUrl = computeWebAppUrl();

    if (fallbackUrl) {
      try {
        cacheManager.put(URL_CACHE_KEY, fallbackUrl, 3600);
        infoLog('✅ 新規URL生成・キャッシュ完了:', fallbackUrl);
      } catch (cacheError) {
        warnLog('キャッシュ保存エラー:', cacheError.message);
      }
    }

    return fallbackUrl || getFallbackUrl();

  } catch (e) {
    errorLog('❌ getWebAppUrlCached error:', e.message);
    return getFallbackUrl();
  }
}

/**
 * フォールバックURL生成
 * @returns {string} フォールバックURL
 */
function getFallbackUrl() {
  try {
    var scriptId = ScriptApp.getScriptId();
    if (scriptId) {
      return 'https://script.google.com/macros/s/' + scriptId + '/exec';
    }
  } catch (e) {
    errorLog('フォールバックURL生成エラー: ' + e.message);
  }
  return '';
}

/**
 * URLキャッシュをクリアして再初期化
 */
function clearUrlCache() {
  try {
    var cache = CacheService.getScriptCache();
    cache.remove(URL_CACHE_KEY);
    debugLog('URL cache cleared successfully');

    // 新しいURLを即座に生成してキャッシュ
    var newUrl = computeWebAppUrl();
    if (newUrl) {
      cache.put(URL_CACHE_KEY, newUrl, URL_CACHE_TTL);
      debugLog('New URL cached:', newUrl);
    }

    return newUrl;
  } catch (e) {
    errorLog('clearUrlCache error:', e.message);
    return getFallbackUrl();
  }
}

/**
 * WebアプリのベースURLを取得（キャッシュ利用）
 * @returns {string} WebアプリのベースURL
 */
function getWebAppBaseUrl() {
  return getWebAppUrlCached();
}

/**
 * 強制的にURLシステムをリセット（公開API）
 * フロントエンドから呼び出し可能
 */
function forceUrlSystemReset() {
  try {
    debugLog('Forcing URL system reset...');

    // 全てのURLキャッシュをクリア
    var cache = CacheService.getScriptCache();
    cache.remove(URL_CACHE_KEY);

    // 新しいURLを生成
    var newUrl = computeWebAppUrl();
    debugLog('New URL generated:', newUrl);

    // 新しいURLをキャッシュ
    if (newUrl) {
      cache.put(URL_CACHE_KEY, newUrl, URL_CACHE_TTL);
    }

    return {
      status: 'success',
      message: 'URLシステムがリセットされました',
      newUrl: newUrl
    };
  } catch (e) {
    errorLog('forceUrlSystemReset error:', e.message);
    return {
      status: 'error',
      message: 'URLシステムリセットに失敗しました: ' + e.message
    };
  }
}

/**
 * アプリケーション用のURL群を生成（最適化版）
 * @param {string} userId - ユーザーID
 * @returns {object} URL群
 */
function generateUserUrls(userId) {
  try {
    if (!userId || typeof userId !== 'string' || userId.trim() === '') {
      errorLog('generateUserUrls: 無効なuserId:', userId);
      return {
        webAppUrl: '',
        adminUrl: '',
        viewUrl: '',
        setupUrl: '',
        status: 'error',
        message: '無効なユーザーIDです'
      };
    }

    var webAppUrl = getWebAppUrlCached();

    if (!webAppUrl) {
      errorLog('WebAppURL取得失敗、フォールバック使用');
      webAppUrl = getFallbackUrl();
    }

    if (!webAppUrl) {
      return {
        webAppUrl: '',
        adminUrl: '',
        viewUrl: '',
        setupUrl: '',
        status: 'error',
        message: 'WebアプリURLの取得に失敗しました'
      };
    }

    var encodedUserId = encodeURIComponent(userId.trim());

    return {
      webAppUrl: webAppUrl,
      adminUrl: webAppUrl + '?mode=admin&userId=' + encodedUserId,
      viewUrl: webAppUrl + '?mode=view&userId=' + encodedUserId,
      setupUrl: webAppUrl + '?setup=true',
      status: 'success'
    };

  } catch (e) {
    errorLog('generateUserUrls error:', e.message);
    return {
      webAppUrl: '',
      adminUrl: '',
      viewUrl: '',
      setupUrl: '',
      status: 'error',
      message: 'URL生成エラー: ' + e.message
    };
  }
}


/**
 * キャッシュバスティング対応のURL生成
 * 非公開状態時の確実なリダイレクトを保証するため、キャッシュ無効化パラメータを追加
 * @param {string} baseUrl - ベースURL
 * @param {Object} options - オプション設定
 * @returns {string} キャッシュバスティング付きURL
 */
function addCacheBustingParams(baseUrl, options = {}) {
  try {
    if (!baseUrl || typeof baseUrl !== 'string') {
      warnLog('addCacheBustingParams: 無効なbaseUrlが渡されました:', baseUrl);
      return baseUrl;
    }

    const url = new URL(baseUrl);

    // キャッシュバスティングパラメータを追加
    if (options.forceFresh || options.unpublished) {
      // タイムスタンプベースのキャッシュバスティング
      url.searchParams.set('_cb', Date.now().toString());
      debugLog('🔄 Cache busting timestamp added:', Date.now());
    }

    if (options.sessionId) {
      // セッション固有のパラメータ
      url.searchParams.set('_sid', options.sessionId);
    }

    if (options.publicationStatus === 'unpublished') {
      // 非公開状態の明示的な指定
      url.searchParams.set('_ps', 'unpublished');
      url.searchParams.set('_t', Math.random().toString(36).substr(2, 9));
      debugLog('🚫 Unpublished state cache busting applied');
    }

    if (options.version) {
      // バージョン指定
      url.searchParams.set('_v', options.version);
    }

    return url.toString();

  } catch (error) {
    errorLog('addCacheBustingParams error:', error.message);
    return baseUrl; // エラー時は元のURLを返す
  }
}

/**
 * 非公開状態用の特別なURL生成
 * キャッシュを完全に無効化したアクセスを保証
 * @param {string} userId - ユーザーID
 * @returns {string} 非公開状態アクセス用URL
 */
function generateUnpublishedStateUrl(userId) {
  try {
    const baseUrl = getWebAppUrlCached();
    if (!baseUrl) {
      errorLog('generateUnpublishedStateUrl: ベースURLの取得に失敗');
      return '';
    }

    // 非公開状態用の強力なキャッシュバスティング
    const cacheBustedUrl = addCacheBustingParams(baseUrl, {
      forceFresh: true,
      publicationStatus: 'unpublished',
      sessionId: Session.getTemporaryActiveUserKey() || 'session_' + Date.now(),
      version: Date.now().toString()
    });

    // userIdパラメータを追加（mode=viewは除外して非公開ページに誘導）
    const url = new URL(cacheBustedUrl);
    if (userId) {
      url.searchParams.set('userId', userId);
    }

    debugLog('🚫 Unpublished state URL generated:', url.toString());
    return url.toString();

  } catch (error) {
    errorLog('generateUnpublishedStateUrl error:', error.message);
    // フォールバック: 基本的なURL
    const baseUrl = getWebAppUrlCached();
    return baseUrl + (userId ? '?userId=' + userId + '&_cb=' + Date.now() : '?_cb=' + Date.now());
  }
}

/**
 * パブリケーション状態を考慮したURL生成の拡張
 * @param {string} userId - ユーザーID
 * @param {Object} options - URL生成オプション
 * @returns {Object} 拡張されたURL群
 */
function generateUserUrlsWithCacheBusting(userId, options = {}) {
  try {
    const standardUrls = generateUserUrls(userId);

    if (standardUrls.status === 'error') {
      return standardUrls;
    }

    // キャッシュバスティング対応版のURL生成
    const cacheBustOptions = {
      forceFresh: options.forceFresh || false,
      publicationStatus: options.publicationStatus || 'unknown',
      sessionId: options.sessionId || Session.getTemporaryActiveUserKey() || 'session_' + Date.now()
    };

    return {
      ...standardUrls,
      // 既存URLにキャッシュバスティングを追加
      adminUrl: addCacheBustingParams(standardUrls.adminUrl, cacheBustOptions),
      viewUrl: addCacheBustingParams(standardUrls.viewUrl, cacheBustOptions),
      setupUrl: addCacheBustingParams(standardUrls.setupUrl, cacheBustOptions),
      // 非公開状態専用URL
      unpublishedUrl: generateUnpublishedStateUrl(userId),
      cacheBustingApplied: true,
      cacheBustOptions: cacheBustOptions
    };

  } catch (error) {
    errorLog('generateUserUrlsWithCacheBusting error:', error.message);
    return generateUserUrls(userId); // フォールバック
  }
}

