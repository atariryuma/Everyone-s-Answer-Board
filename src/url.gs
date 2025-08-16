/**
 * @fileoverview Web app URL utilities.
 * Uses Apps Script API deployments for stable URL retrieval.
 */

// Define basic logging helpers for GAS and test environments
// debugLog関数はdebugConfig.gsで統一定義されていますが、テスト環境でのfallback定義
// debugLog は debugConfig.gs で統一制御されるため、重複定義を削除

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

/**
 * Retrieve the deployed Web App URL via Apps Script API.
 * This avoids development URLs and supports silo-type multi-tenancy
 * by returning the domain specific URL of the current deployment.
 *
 * @returns {string} Web app URL or empty string on failure.
 */
function getWebAppUrl() {
  // キャッシュから結果を取得（5分間有効）
  const cacheKey = 'WEB_APP_URL_CACHE';
  try {
    const cachedUrl = CacheService.getScriptCache().get(cacheKey);
    if (cachedUrl) {
      debugLog('Web app URL retrieved from cache:', cachedUrl);
      return cachedUrl;
    }
  } catch (cacheError) {
    debugLog('Cache access failed:', cacheError.message);
  }

  let finalUrl = '';
  
  try {
    // AppsScriptオブジェクトの詳細な存在チェック
    if (typeof AppsScript === 'undefined') {
      debugLog('AppsScript API not defined, using fallback method');
      finalUrl = getFallbackWebAppUrl();
    } else if (!AppsScript.Script) {
      debugLog('AppsScript.Script not available, using fallback method');
      finalUrl = getFallbackWebAppUrl();
    } else if (!AppsScript.Script.Deployments) {
      debugLog('AppsScript.Script.Deployments not available, using fallback method');
      finalUrl = getFallbackWebAppUrl();
    } else {
      // AppsScript APIを使用してURL取得（本番環境での確実な動作を目指す）
      try {
        const scriptId = ScriptApp.getScriptId();
        
        // デプロイメント一覧取得を試行
        let deploymentsList;
        try {
          deploymentsList = AppsScript.Script.Deployments.list(scriptId, {
            fields: 'deployments(deploymentId,deploymentConfig(webApp(url)),updateTime)'
          });
        } catch (apiError) {
          debugLog('AppsScript API call failed, trying simple list:', apiError.message);
          // テスト環境等で引数なしの場合のフォールバック
          deploymentsList = AppsScript.Script.Deployments.list();
        }
        
        debugLog('AppsScript.Script.Deployments.list response:', JSON.stringify(deploymentsList, null, 2));
        
        const deployments = deploymentsList.deployments || [];
        
        if (deployments.length === 0) {
          infoLog('getWebAppUrl: デプロイメントが見つからないため、フォールバック処理へ');
          finalUrl = getFallbackWebAppUrl();
        } else {
          // 最新のWebアプリデプロイメントを取得
          const webAppDeployments = deployments
            .filter(d => d.deploymentConfig && d.deploymentConfig.webApp)
            .sort((a, b) => new Date(b.updateTime) - new Date(a.updateTime));
          
          if (webAppDeployments.length > 0) {
            finalUrl = webAppDeployments[0].deploymentConfig.webApp.url;
            infoLog('Web app URL obtained via AppsScript API:', finalUrl);
            debugLog('Selected deployment:', JSON.stringify(webAppDeployments[0], null, 2));
          } else {
            infoLog('getWebAppUrl: WebアプリデプロイメントなしファンドされないWebアプリ、フォールバック処理へ');
            finalUrl = getFallbackWebAppUrl();
          }
        }
      } catch (apiError) {
        infoLog('getWebAppUrl API詳細エラー:', apiError.message);
        finalUrl = getFallbackWebAppUrl();
      }
    }
  } catch (apiError) {
    infoLog('getWebAppUrl API error (normal fallback):', apiError.message);
    finalUrl = getFallbackWebAppUrl();
  }
  
  // 有効なURLが取得できた場合はキャッシュに保存（5分間）
  if (finalUrl) {
    try {
      CacheService.getScriptCache().put(cacheKey, finalUrl, 300); // 5分間キャッシュ
      debugLog('Web app URL cached for 5 minutes');
    } catch (cacheError) {
      debugLog('Failed to cache Web app URL:', cacheError.message);
    }
  }
  
  return finalUrl;
}

/**
 * ScriptApp.getService()を使用したフォールバックWebAppURL取得
 * @returns {string} Web app URL or empty string on failure
 */
function getFallbackWebAppUrl() {

  // Fallback to ScriptApp service URL if API is unavailable or returns nothing
  try {
    const serviceUrl = ScriptApp.getService && ScriptApp.getService().getUrl
      ? ScriptApp.getService().getUrl()
      : '';
    if (serviceUrl) {
      infoLog('Web app URL fallback obtained:', serviceUrl);
      return serviceUrl;
    }
    warnLog('getWebAppUrl fallback failed: service URL unavailable');
  } catch (e) {
    errorLog('getWebAppUrl fallback error:', e.message);
  }

  return '';
}

/**
 * Generate user specific URLs.
 *
 * @param {string} userId - unique user identifier
 * @returns {object} URLs for the user
 */
function generateUserUrls(userId) {
  if (!userId || typeof userId !== 'string' || userId.trim() === '') {
    errorLog('generateUserUrls: invalid userId:', userId);
    return {
      webAppUrl: '',
      adminUrl: '',
      viewUrl: '',
      setupUrl: '',
      status: 'error',
      message: '無効なユーザーIDです'
    };
  }

  const webAppUrl = getWebAppUrl();
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

  const encodedId = encodeURIComponent(userId.trim());
  return {
    webAppUrl: webAppUrl,
    adminUrl: `${webAppUrl}?mode=admin&userId=${encodedId}`,
    viewUrl: `${webAppUrl}?mode=view&userId=${encodedId}`,
    setupUrl: `${webAppUrl}?setup=true&userId=${encodedId}`,
    status: 'success'
  };
}

/**
 * Add cache busting parameters to a URL.
 *
 * @param {string} baseUrl
 * @param {Object} options
 * @returns {string}
 */
function addCacheBustingParams(baseUrl, options = {}) {
  try {
    if (!baseUrl || typeof baseUrl !== 'string') {
      warnLog('addCacheBustingParams: 無効なbaseUrlが渡されました:', baseUrl);
      return baseUrl;
    }

    const url = new URL(baseUrl);

    if (options.forceFresh || options.unpublished) {
      url.searchParams.set('_cb', Date.now().toString());
    }

    if (options.sessionId) {
      url.searchParams.set('_sid', options.sessionId);
    }

    if (options.publicationStatus === 'unpublished') {
      url.searchParams.set('_ps', 'unpublished');
      url.searchParams.set('_t', Math.random().toString(36).substr(2, 9));
    }

    if (options.version) {
      url.searchParams.set('_v', options.version);
    }

    return url.toString();
  } catch (error) {
    errorLog('addCacheBustingParams error:', error.message);
    return baseUrl;
  }
}

/**
 * Generate URL for unpublished state with aggressive cache busting.
 *
 * @param {string} userId
 * @returns {string}
 */
function generateUnpublishedStateUrl(userId) {
  try {
    const baseUrl = getWebAppUrl();
    if (!baseUrl) {
      errorLog('generateUnpublishedStateUrl: ベースURLの取得に失敗');
      return '';
    }

    const cacheBustedUrl = addCacheBustingParams(baseUrl, {
      forceFresh: true,
      publicationStatus: 'unpublished',
      sessionId: (typeof Session !== 'undefined' && Session.getTemporaryActiveUserKey)
        ? Session.getTemporaryActiveUserKey() || 'session_' + Date.now()
        : 'session_' + Date.now(),
      version: Date.now().toString()
    });

    const url = new URL(cacheBustedUrl);
    if (userId) {
      url.searchParams.set('userId', userId);
    }
    return url.toString();
  } catch (error) {
    errorLog('generateUnpublishedStateUrl error:', error.message);
    const baseUrl = getWebAppUrl();
    return baseUrl + (userId ? `?userId=${userId}&_cb=${Date.now()}` : `?_cb=${Date.now()}`);
  }
}

/**
 * Generate URLs with optional cache busting.
 *
 * @param {string} userId
 * @param {Object} options
 * @returns {Object}
 */
function generateUserUrlsWithCacheBusting(userId, options = {}) {
  try {
    const standardUrls = generateUserUrls(userId);
    if (standardUrls.status === 'error') {
      return standardUrls;
    }

    const cacheBustOptions = {
      forceFresh: options.forceFresh || false,
      publicationStatus: options.publicationStatus || 'unknown',
      sessionId: options.sessionId || (typeof Session !== 'undefined' && Session.getTemporaryActiveUserKey ? Session.getTemporaryActiveUserKey() : null)
    };

    return {
      ...standardUrls,
      adminUrl: addCacheBustingParams(standardUrls.adminUrl, cacheBustOptions),
      viewUrl: addCacheBustingParams(standardUrls.viewUrl, cacheBustOptions),
      setupUrl: addCacheBustingParams(standardUrls.setupUrl, cacheBustOptions),
      unpublishedUrl: generateUnpublishedStateUrl(userId),
      cacheBustingApplied: true,
      cacheBustOptions: cacheBustOptions
    };
  } catch (error) {
    errorLog('generateUserUrlsWithCacheBusting error:', error.message);
    return generateUserUrls(userId);
  }
}
