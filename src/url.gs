/**
 * @fileoverview URL管理 - GAS互換版
 */

// URL管理の定数
var URL_CACHE_KEY = 'WEB_APP_URL';
var URL_CACHE_TTL = 21600; // 6時間

/**
 * WebアプリのURLを取得（キャッシュ利用）
 * @returns {string} WebアプリURL
 */
function computeWebAppUrl() {
  try {
    var url = ScriptApp.getService().getUrl();
    if (!url) {
      console.warn('ScriptApp.getService().getUrl()がnullを返しました');
      return getFallbackUrl();
    }

    url = url.replace(/\/$/, '');

    // \"https://script.google.com/a/<domain>/macros/s/...\" 形式を
    // \"https://script.google.com/a/macros/<domain>/s/...\" に補正
    var wrongPattern = /^https:\/\/script\.google\.com\/a\/([^\/]+)\/macros\//;
    var match = url.match(wrongPattern);
    if (match) {
      url = url.replace(wrongPattern, 'https://script.google.com/a/macros/' + match[1] + '/');
    }

    return url;
  } catch (e) {
    console.error('WebアプリURL取得エラー: ' + e.message);
    return getFallbackUrl();
  }
}

function getWebAppUrlCached() {
  var cachedUrl = cacheManager.get(URL_CACHE_KEY, function() {
    return computeWebAppUrl();
  }, { ttl: URL_CACHE_TTL });

  try {
    var currentUrl = computeWebAppUrl();
    // キャッシュされたURLと現在のURLが異なる場合、キャッシュを更新
    if (cachedUrl !== currentUrl) {
      console.log('Cached URL is stale or incorrect. Updating cache.');
      cacheManager.remove(URL_CACHE_KEY); // 古いキャッシュを無効化
      cachedUrl = cacheManager.get(URL_CACHE_KEY, function() {
        return currentUrl; // 最新のURLをキャッシュに保存
      }, { ttl: URL_CACHE_TTL });
    }
  } catch (e) {
    console.warn('WebアプリURL更新チェックでエラー: ' + e.message);
  }

  return cachedUrl;
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
    // userIdの妥当性チェック
    if (!userId || userId === 'undefined' || typeof userId !== 'string' || userId.trim() === '') {
      console.error('generateAppUrls: 無効なuserIdが渡されました: ' + userId);
      return {
        webAppUrl: '',
        adminUrl: '',
        viewUrl: '',
        setupUrl: '',
        status: 'error',
        message: '無効なユーザーIDです。有効なIDを指定してください。'
      };
    }
    
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
    
    // URLエンコードして安全にユーザーIDを追加
    var encodedUserId = encodeURIComponent(userId.trim());
    
    return {
      webAppUrl: webAppUrl,
      adminUrl: webAppUrl.replace(/\/dev$/, '/exec') +
        '?userId=' + encodedUserId + '&mode=admin',
      viewUrl: webAppUrl + '?userId=' + encodedUserId + '&mode=view',
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

