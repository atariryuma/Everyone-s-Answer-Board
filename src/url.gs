/**
 * @fileoverview Web app URL utilities.
 * Uses Apps Script API deployments for stable URL retrieval.
 */

// Define basic logging helpers for GAS and test environments
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

/**
 * Retrieve the deployed Web App URL via Apps Script API.
 * This avoids development URLs and supports silo-type multi-tenancy
 * by returning the domain specific URL of the current deployment.
 *
 * @returns {string} Web app URL or empty string on failure.
 */
function getWebAppUrl() {
  try {
    const scriptId = ScriptApp.getScriptId();
    const deployments = AppsScript.Script.Deployments.list(scriptId, {
      fields: 'deployments(deploymentConfig(webApp(url)))'
    }).deployments || [];
    const webApp = deployments.find(d => d.deploymentConfig && d.deploymentConfig.webApp);
    const url = webApp ? webApp.deploymentConfig.webApp.url : '';
    infoLog('Web app URL obtained:', url);
    return url;
  } catch (e) {
    errorLog('getWebAppUrl error:', e.message);
    return '';
  }
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
    setupUrl: `${webAppUrl}?setup=true`,
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
