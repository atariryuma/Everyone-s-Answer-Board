/**
 * @fileoverview URL管理 - シンプル版
 * 常に本番環境のWebアプリURLを生成します。
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

/**
 * 常に本番環境のWebアプリURLを生成します。
 * ScriptApp.getScriptId() を使用して、デプロイされたスクリプトの固定URLを返します。
 * @returns {string} 本番環境のWebアプリURL
 */
function getProductionWebAppUrl() {
  try {
    const scriptId = ScriptApp.getScriptId();
    if (!scriptId) {
      errorLog('Script ID not found. Cannot generate production URL.');
      return '';
    }
    const url = `https://script.google.com/macros/s/${scriptId}/exec`;
    infoLog('✅ Production WebApp URL generated:', url);
    return url;
  } catch (e) {
    errorLog('Error generating production WebApp URL:', e.message);
    return '';
  }
}

/**
 * WebアプリのベースURLを取得します。
 * キャッシュ機能付きでパフォーマンスを最適化。
 * @returns {string} WebアプリのベースURL
 */
function getWebAppBaseUrl() {
  try {
    return cacheManager.get('WEB_APP_URL', () => getProductionWebAppUrl());
  } catch (e) {
    warnLog('Cache error, falling back to direct URL generation:', e.message);
    return getProductionWebAppUrl();
  }
}

/**
 * ユーザー毎のアプリURLを生成します。
 * @param {string} userId ユーザーID
 * @returns {{adminUrl: string, viewUrl: string}} 生成されたURL
 */
function generateUserUrls(userId) {
  if (!userId) {
    errorLog('generateUserUrls: userId is required');
    return { adminUrl: '', viewUrl: '' };
  }
  
  const baseUrl = getWebAppBaseUrl();
  if (!baseUrl) {
    errorLog('generateUserUrls: Failed to get base URL');
    return { adminUrl: '', viewUrl: '' };
  }
  
  const encodedUserId = encodeURIComponent(userId);
  return {
    adminUrl: `${baseUrl}?mode=admin&userId=${encodedUserId}`,
    viewUrl: `${baseUrl}?mode=view&userId=${encodedUserId}`
  };
}

/**
 * 後方互換性のための廃止予定関数
 * @deprecated getProductionWebAppUrl() を直接使用してください
 */
function getWebAppUrl() {
  warnLog('getWebAppUrl() is deprecated. Use getProductionWebAppUrl() instead.');
  return getProductionWebAppUrl();
}

/**
 * 後方互換性のための廃止予定関数  
 * @deprecated getProductionWebAppUrl() を直接使用してください
 */
function getWebAppUrlCached() {
  warnLog('getWebAppUrlCached() is deprecated. Use getWebAppBaseUrl() instead.');
  return getWebAppBaseUrl();
}