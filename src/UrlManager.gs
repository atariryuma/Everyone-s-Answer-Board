/**
 * @fileoverview URL管理 - GAS互換版
 */

// URL管理の定数
var URL_CACHE_KEY = 'WEB_APP_URL';
var URL_CACHE_TTL = 3600; // 1時間

/**
 * WebアプリのURLを取得（キャッシュ利用）
 * @returns {string} WebアプリURL
 */
function getWebAppUrlCached() {
  return AdvancedCacheManager.smartGet(URL_CACHE_KEY, function() {
    try {
      var url = ScriptApp.getService().getUrl();
      if (!url) {
        console.warn('ScriptApp.getService().getUrl()がnullを返しました');
        return getFallbackUrl();
      }
      
      // URLの正規化
      return url.indexOf('/') === url.length - 1 ? url.slice(0, -1) : url;
    } catch (e) {
      console.error('WebアプリURL取得エラー: ' + e.message);
      return getFallbackUrl();
    }
  }, 'medium');
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
    console.error('フォールバックURL生成エラー: ' + e.message);
  }
  return '';
}

/**
 * アプリケーション用のURL群を生成
 * @param {string} userId - ユーザーID
 * @returns {object} URL群
 */
function generateAppUrls(userId) {
  try {
    var webAppUrl = getWebAppUrlCached();
    
    if (!webAppUrl) {
      return {
        webAppUrl: '',
        adminUrl: '',
        viewUrl: '',
        setupUrl: '',
        status: 'error',
        message: 'WebアプリURLが取得できませんでした'
      };
    }
    
    return {
      webAppUrl: webAppUrl,
      adminUrl: webAppUrl + '?userId=' + userId + '&mode=admin',
      viewUrl: webAppUrl + '?userId=' + userId,
      setupUrl: webAppUrl + '?setup=true',
      status: 'success'
    };
  } catch (e) {
    console.error('URL生成エラー: ' + e.message);
    return {
      webAppUrl: '',
      adminUrl: '',
      viewUrl: '',
      setupUrl: '',
      status: 'error',
      message: 'URLの生成に失敗しました: ' + e.message
    };
  }
}

/**
 * URLキャッシュをクリア
 */
function clearUrlCache() {
  removeCachedValue(URL_CACHE_KEY);
}