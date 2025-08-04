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
    // 複数の方法でURLを取得し検証
    let primaryUrl = null;
    let fallbackUrl = null;
    
    // 方法1: ScriptApp.getService().getUrl()
    try {
      primaryUrl = ScriptApp.getService().getUrl();
      if (primaryUrl) {
        primaryUrl = primaryUrl.replace(/\/$/, ''); // 末尾のスラッシュを除去
        debugLog('🔍 Primary URL取得:', primaryUrl);
      }
    } catch (e) {
      warnLog('Primary URL取得エラー:', e.message);
    }

    // 方法2: ScriptIdベースのフォールバック
    try {
      const scriptId = ScriptApp.getScriptId();
      if (scriptId) {
        fallbackUrl = 'https://script.google.com/macros/s/' + scriptId + '/exec';
        debugLog('🔄 Fallback URL生成:', fallbackUrl);
      }
    } catch (e) {
      warnLog('Fallback URL生成エラー:', e.message);
    }

    // 開発環境URLの検出（優先度順）
    const devIndicators = [
      'userCodeAppPanel',  // 最も確実な開発環境指標
      '/dev',              // 開発エンドポイント
      '/test',             // テストエンドポイント
      'googleusercontent.com' // Dev環境でよく使われる
    ];

    // Primary URLの検証
    if (primaryUrl) {
      // 開発環境URLかチェック
      let isDevelopmentUrl = false;
      for (var i = 0; i < devIndicators.length; i++) {
        if (primaryUrl.includes(devIndicators[i])) {
          warnLog('⚠️ Primary URL開発環境検出:', primaryUrl, '→ フォールバック使用');
          isDevelopmentUrl = true;
          break;
        }
      }

      if (!isDevelopmentUrl) {
        // URL形式の検証（Google Workspace対応）
        var isValidGoogleWorkspaceUrl = /^https:\/\/script\.google\.com\/a\/macros\/[^\/]+\/s\/[A-Za-z0-9_-]+\/exec$/.test(primaryUrl);
        var isValidStandardUrl = /^https:\/\/script\.google\.com\/macros\/s\/[A-Za-z0-9_-]+\/exec$/.test(primaryUrl);
        var isValidDeployUrl = /^https:\/\/[a-z0-9-]+\.googleusercontent\.com$/.test(primaryUrl);

        if (isValidGoogleWorkspaceUrl || isValidStandardUrl || isValidDeployUrl) {
          infoLog('✅ Primary URL検証完了 (Google Workspace対応):', primaryUrl);
          return primaryUrl;
        }

        // URL形式の補正（Google Workspace形式対応）
        var correctedUrl = primaryUrl.replace(
          /^https:\/\/script\.google\.com\/a\/([^\/]+)\/macros\//,
          'https://script.google.com/a/macros/$1/'
        );

        if (correctedUrl !== primaryUrl) {
          debugLog('🔧 URL形式を補正:', primaryUrl, '→', correctedUrl);
          // 補正後のURLを再検証
          if (/^https:\/\/script\.google\.com\/a\/macros\/[^\/]+\/s\/[A-Za-z0-9_-]+\/exec$/.test(correctedUrl)) {
            infoLog('✅ 補正されたURL検証完了:', correctedUrl);
            return correctedUrl;
          }
        }
      }
    }

    // Fallback URLの検証
    if (fallbackUrl) {
      // 開発環境でないかチェック
      let isFallbackDev = false;
      for (var i = 0; i < devIndicators.length; i++) {
        if (fallbackUrl.includes(devIndicators[i])) {
          warnLog('⚠️ Fallback URL開発環境検出:', fallbackUrl);
          isFallbackDev = true;
          break;
        }
      }
      
      if (!isFallbackDev && /^https:\/\/script\.google\.com\/macros\/s\/[A-Za-z0-9_-]+\/exec$/.test(fallbackUrl)) {
        infoLog('✅ Fallback URL使用:', fallbackUrl);
        return fallbackUrl;
      }
    }

    // 最終フォールバック
    errorLog('❌ 有効なURL生成に失敗 - Primary:', primaryUrl, 'Fallback:', fallbackUrl);
    return fallbackUrl || '';

  } catch (e) {
    errorLog('❌ computeWebAppUrl 重大エラー:', e.message);
    return getFallbackUrl();
  }
}

function getWebAppUrlCached() {
  try {
    // 統合キャッシュマネージャーを使用したURL管理（最適化版）
    var webAppUrl = cacheManager.get(URL_CACHE_KEY, function() {
      debugLog('🔍 WebAppURL キャッシュmiss - 新規生成');
      return computeWebAppUrl(); // 最適化されたURL生成を使用
    }, {
      ttl: 3600, // 1時間キャッシュ
      enableMemoization: true
    });

    // URL正規化
    webAppUrl = normalizeUrlString(webAppUrl);

    // キャッシュ済みURLの妥当性チェック
    if (webAppUrl) {
      // 開発環境URL検出（簡素化）
      var devIndicators = ['userCodeAppPanel', '/dev', '/test'];
      var isDevelopmentUrl = devIndicators.some(function(indicator) {
        return webAppUrl.includes(indicator);
      });

      if (isDevelopmentUrl) {
        warnLog('⚠️ キャッシュされた開発URL検出、クリアして再生成:', webAppUrl);
        cacheManager.remove(URL_CACHE_KEY);
        // 再帰的に呼び出して新しいURLを生成（最大1回）
        return computeWebAppUrl(); // 直接生成で無限ループ回避
      }

      infoLog('✅ 有効なキャッシュURL取得:', webAppUrl);
      return webAppUrl;
    }

    // フォールバック処理
    warnLog('⚠️ キャッシュURL無効、直接生成に切り替え');
    var fallbackUrl = computeWebAppUrl();

    // 有効なURLが生成された場合はキャッシュに保存
    if (fallbackUrl && !fallbackUrl.includes('userCodeAppPanel')) {
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
 * URLキャッシュをクリアして再初期化（強化版）
 */
function clearUrlCache() {
  try {
    infoLog('🗑️ URLキャッシュクリア開始');
    
    // 全てのキャッシュソースをクリア
    var cache = CacheService.getScriptCache();
    cache.remove(URL_CACHE_KEY);
    
    // 統合キャッシュマネージャーからもクリア
    if (typeof cacheManager !== 'undefined' && cacheManager.remove) {
      cacheManager.remove(URL_CACHE_KEY);
      debugLog('🧹 統合キャッシュマネージャーからもクリア完了');
    }

    infoLog('✅ 全URLキャッシュクリア完了');

    // 新しいURLを即座に生成
    var newUrl = computeWebAppUrl();
    
    if (newUrl) {
      // 生成されたURLの検証
      var isValidUrl = false;
      
      // Google Workspace URL形式
      if (/^https:\/\/script\.google\.com\/a\/macros\/[^\/]+\/s\/[A-Za-z0-9_-]+\/exec$/.test(newUrl)) {
        isValidUrl = true;
        infoLog('✅ Google WorkspaceURL形式で新規URL生成:', newUrl);
      }
      // 標準URL形式
      else if (/^https:\/\/script\.google\.com\/macros\/s\/[A-Za-z0-9_-]+\/exec$/.test(newUrl)) {
        isValidUrl = true;
        infoLog('✅ 標準URL形式で新規URL生成:', newUrl);
      }

      // 有効なURLの場合のみキャッシュ
      if (isValidUrl && !newUrl.includes('googleusercontent.com') && !newUrl.includes('userCodeAppPanel')) {
        cache.put(URL_CACHE_KEY, newUrl, URL_CACHE_TTL);
        
        // 統合キャッシュマネージャーにもキャッシュ
        if (typeof cacheManager !== 'undefined' && cacheManager.put) {
          cacheManager.put(URL_CACHE_KEY, newUrl, 3600);
        }
        
        infoLog('✅ 新規URL生成・キャッシュ完了:', newUrl);
      } else {
        warnLog('⚠️ 生成されたURLが無効またはdev環境のためキャッシュしません:', newUrl);
      }
    } else {
      warnLog('⚠️ 新しいURL生成に失敗');
    }

    return newUrl;
  } catch (e) {
    errorLog('❌ clearUrlCache error:', e.message);
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

    // 開発URLチェック
    if (newUrl && (newUrl.includes('googleusercontent.com') || newUrl.includes('userCodeAppPanel'))) {
      warnLog('Development URL detected, using fallback');
      newUrl = getFallbackUrl();
    }

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
 * URLの妥当性を検証
 * @param {string} url - 検証するURL
 * @returns {object} 検証結果
 */
function validateWebAppUrl(url) {
  if (!url || typeof url !== 'string') {
    return { isValid: false, reason: 'URLが空または無効な形式' };
  }

  // 開発環境URLの検出
  const devIndicators = ['userCodeAppPanel', '/dev', '/test', 'googleusercontent.com'];
  for (var i = 0; i < devIndicators.length; i++) {
    if (url.includes(devIndicators[i])) {
      return { isValid: false, reason: '開発環境URLが検出されました', type: 'development' };
    }
  }

  // Google Workspace URL形式
  if (/^https:\/\/script\.google\.com\/a\/macros\/[^\/]+\/s\/[A-Za-z0-9_-]+\/exec$/.test(url)) {
    const domain = url.match(/\/a\/macros\/([^\/]+)\//);
    return { 
      isValid: true, 
      type: 'google-workspace', 
      domain: domain ? domain[1] : 'unknown',
      scriptId: url.match(/\/s\/([A-Za-z0-9_-]+)\/exec/)?.[1]
    };
  }

  // 標準URL形式
  if (/^https:\/\/script\.google\.com\/macros\/s\/[A-Za-z0-9_-]+\/exec$/.test(url)) {
    return { 
      isValid: true, 
      type: 'standard',
      scriptId: url.match(/\/s\/([A-Za-z0-9_-]+)\/exec/)?.[1]
    };
  }

  return { isValid: false, reason: '認識されないURL形式', url: url };
}

/**
 * アプリケーション用のURL群を生成（最適化版）
 * @param {string} userId - ユーザーID
 * @returns {object} URL群
 */
function generateUserUrls(userId) {
  try {
    // userIdの妥当性チェック（簡素化）
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

    // 最適化済みURL取得（リトライ処理内包）
    var webAppUrl = getWebAppUrlCached();

    // 基本的な妥当性チェック
    if (!webAppUrl) {
      errorLog('WebAppURL取得失敗、フォールバック使用');
      webAppUrl = getFallbackUrl();
    }

    // URL検証システムを使用
    var validation = validateWebAppUrl(webAppUrl);
    if (!validation.isValid) {
      warnLog('⚠️ URL検証失敗:', validation.reason, 'URL:', webAppUrl);
      
      // フォールバックを試行
      var fallbackUrl = getFallbackUrl();
      var fallbackValidation = validateWebAppUrl(fallbackUrl);
      
      if (fallbackValidation.isValid) {
        webAppUrl = fallbackUrl;
        infoLog('✅ フォールバックURL使用:', webAppUrl);
      } else {
        errorLog('❌ フォールバックURLも無効:', fallbackValidation.reason);
        return {
          webAppUrl: webAppUrl, // エラーでも一応返す
          adminUrl: '',
          viewUrl: '',
          setupUrl: '',
          status: 'error',
          message: 'WebアプリURLの検証に失敗: ' + validation.reason,
          validation: validation
        };
      }
    } else {
      infoLog('✅ URL検証成功:', validation.type, validation.domain ? `(${validation.domain})` : '', webAppUrl);
    }

    // URLが取得できない場合のエラーハンドリング
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

    // URL生成（エンコード処理）
    var encodedUserId = encodeURIComponent(userId.trim());

    return {
      webAppUrl: webAppUrl,
      adminUrl: webAppUrl + '?mode=admin&userId=' + encodedUserId,
      viewUrl: webAppUrl + '?mode=view&userId=' + encodedUserId,
      setupUrl: webAppUrl + '?setup=true',
      status: 'success',
      validation: validation  // 検証情報も含める
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

