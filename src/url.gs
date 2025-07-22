/**
 * @fileoverview URL生成・管理機能
 * アプリケーション全体のURL生成を統一管理
 */

/**
 * WebApp基本URL取得（キャッシュ対応）
 */
function getWebAppBaseUrl() {
  const cacheKey = 'webapp_base_url';
  return cacheManager.get(cacheKey, () => {
    try {
      const url = ScriptApp.getService().getUrl();
      if (!url) {
        throw new Error('WebApp URLが取得できませんでした');
      }
      console.log('WebApp基本URL取得成功:', url);
      return url;
    } catch (error) {
      console.error('WebApp基本URL取得エラー:', error.message);
      throw new Error('WebApp基本URLの取得に失敗しました: ' + error.message);
    }
  }, { ttl: 3600 }); // 1時間キャッシュ
}

/**
 * 管理パネルURLを生成
 * @param {string} userId - ユーザーID
 * @returns {string} 管理パネルURL
 */
function buildAdminPanelUrl(userId) {
  try {
    if (!userId || typeof userId !== 'string') {
      throw new Error('無効なユーザーIDです');
    }
    
    const baseUrl = getWebAppBaseUrl();
    const encodedUserId = encodeURIComponent(userId);
    const adminUrl = `${baseUrl}?mode=admin&userId=${encodedUserId}`;
    
    debugLog(`管理パネルURL生成: ${adminUrl}`);
    return adminUrl;
    
  } catch (error) {
    console.error('管理パネルURL生成エラー:', error.message);
    throw new Error('管理パネルURLの生成に失敗しました: ' + error.message);
  }
}

/**
 * 回答ボードURLを生成
 * @param {string} userId - ユーザーID
 * @returns {string} 回答ボードURL
 */
function buildViewBoardUrl(userId) {
  try {
    if (!userId || typeof userId !== 'string') {
      throw new Error('無効なユーザーIDです');
    }
    
    const baseUrl = getWebAppBaseUrl();
    const encodedUserId = encodeURIComponent(userId);
    const viewUrl = `${baseUrl}?mode=view&userId=${encodedUserId}`;
    
    debugLog(`回答ボードURL生成: ${viewUrl}`);
    return viewUrl;
    
  } catch (error) {
    console.error('回答ボードURL生成エラー:', error.message);
    throw new Error('回答ボードURLの生成に失敗しました: ' + error.message);
  }
}

/**
 * セットアップURLを生成
 * @param {string} [userId] - ユーザーID（オプション）
 * @returns {string} セットアップURL
 */
function buildSetupUrl(userId = null) {
  try {
    const baseUrl = getWebAppBaseUrl();
    let setupUrl = `${baseUrl}?setup=true`;
    
    if (userId && typeof userId === 'string') {
      const encodedUserId = encodeURIComponent(userId);
      setupUrl += `&userId=${encodedUserId}`;
    }
    
    debugLog(`セットアップURL生成: ${setupUrl}`);
    return setupUrl;
    
  } catch (error) {
    console.error('セットアップURL生成エラー:', error.message);
    throw new Error('セットアップURLの生成に失敗しました: ' + error.message);
  }
}

/**
 * ログインURLを生成
 * @returns {string} ログインURL
 */
function buildLoginUrl() {
  try {
    const baseUrl = getWebAppBaseUrl();
    debugLog(`ログインURL生成: ${baseUrl}`);
    return baseUrl;
    
  } catch (error) {
    console.error('ログインURL生成エラー:', error.message);
    throw new Error('ログインURLの生成に失敗しました: ' + error.message);
  }
}

/**
 * ユーザー用URL一式を生成（キャッシュ対応）
 * @param {string} userId - ユーザーID
 * @returns {object} URL一式
 */
function generateUserUrls(userId) {
  const cacheKey = `user_urls_${userId}`;
  return cacheManager.get(cacheKey, () => {
    try {
      if (!userId || typeof userId !== 'string') {
        throw new Error('無効なユーザーIDです');
      }
      
      const urls = {
        baseUrl: getWebAppBaseUrl(),
        adminUrl: buildAdminPanelUrl(userId),
        viewUrl: buildViewBoardUrl(userId),
        setupUrl: buildSetupUrl(userId),
        loginUrl: buildLoginUrl()
      };
      
      console.log(`ユーザーURL一式生成完了: ${userId}`);
      return urls;
      
    } catch (error) {
      console.error('ユーザーURL生成エラー:', error.message);
      throw new Error('ユーザーURLの生成に失敗しました: ' + error.message);
    }
  }, { ttl: 1800 }); // 30分キャッシュ
}

/**
 * HTTPリダイレクトレスポンスを生成
 * @param {string} targetUrl - リダイレクト先URL
 * @param {number} [statusCode=302] - HTTPステータスコード
 * @returns {HtmlOutput} リダイレクトレスポンス
 */
function createRedirectResponse(targetUrl, statusCode = 302) {
  try {
    if (!targetUrl || typeof targetUrl !== 'string') {
      throw new Error('無効なリダイレクトURLです');
    }
    
    // HTTPリダイレクトのHTMLを生成
    const redirectHtml = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta http-equiv="refresh" content="0;url=${targetUrl}">
  <title>リダイレクト中...</title>
  <style>
    body {
      font-family: 'Google Sans', Arial, sans-serif;
      display: flex;
      justify-content: center;
      align-items: center;
      height: 100vh;
      margin: 0;
      background-color: #f8f9fa;
    }
    .redirect-message {
      text-align: center;
      color: #5f6368;
    }
    .spinner {
      border: 3px solid #f3f3f3;
      border-top: 3px solid #4285f4;
      border-radius: 50%;
      width: 30px;
      height: 30px;
      animation: spin 1s linear infinite;
      margin: 0 auto 20px;
    }
    @keyframes spin {
      0% { transform: rotate(0deg); }
      100% { transform: rotate(360deg); }
    }
  </style>
</head>
<body>
  <div class="redirect-message">
    <div class="spinner"></div>
    <p>ページを移動しています...</p>
    <p><small>自動的に移動しない場合は<a href="${targetUrl}">こちら</a>をクリックしてください</small></p>
  </div>
  <script>
    // 即座にリダイレクト実行
    setTimeout(function() {
      window.location.href = '${targetUrl}';
    }, 100);
  </script>
</body>
</html>`;
    
    const output = HtmlService.createHtmlOutput(redirectHtml);
    output.setTitle('リダイレクト中...');
    
    console.log(`リダイレクトレスポンス生成: ${targetUrl}`);
    return output;
    
  } catch (error) {
    console.error('リダイレクトレスポンス生成エラー:', error.message);
    throw new Error('リダイレクトレスポンスの生成に失敗しました: ' + error.message);
  }
}

/**
 * URLパラメータを安全に解析
 * @param {object} e - doGet()のイベントオブジェクト
 * @returns {object} 解析されたパラメータ
 */
function parseUrlParameters(e) {
  try {
    const params = e.parameters || {};
    const parsed = {};
    
    // 各パラメータを安全に処理
    Object.keys(params).forEach(key => {
      const value = params[key];
      if (Array.isArray(value)) {
        parsed[key] = value[0] || ''; // 配列の場合は最初の値
      } else {
        parsed[key] = String(value || '').trim();
      }
    });
    
    // URLデコード
    if (parsed.userId) {
      try {
        parsed.userId = decodeURIComponent(parsed.userId);
      } catch (decodeError) {
        console.warn('userIDのデコードに失敗:', decodeError.message);
        parsed.userId = String(params.userId || '').trim();
      }
    }
    
    debugLog('URLパラメータ解析完了:', JSON.stringify(parsed));
    return parsed;
    
  } catch (error) {
    console.error('URLパラメータ解析エラー:', error.message);
    return {}; // エラー時は空オブジェクトを返す
  }
}

/**
 * URLの妥当性をチェック
 * @param {string} url - チェック対象URL
 * @returns {boolean} 妥当な場合true
 */
function isValidUrl(url) {
  try {
    if (!url || typeof url !== 'string') {
      return false;
    }
    
    // 基本的なURL形式チェック
    const urlPattern = /^https?:\/\/.+/;
    if (!urlPattern.test(url)) {
      return false;
    }
    
    // Google Apps Script WebApp URLかチェック
    const gasUrlPattern = /script\.google\.com.*\/exec/;
    return gasUrlPattern.test(url);
    
  } catch (error) {
    console.warn('URL妥当性チェックエラー:', error.message);
    return false;
  }
}

/**
 * 後方互換性のためのレガシー関数
 * @deprecated buildAdminPanelUrl()を使用してください
 */
function buildUserAdminUrl(userId) {
  console.warn('buildUserAdminUrl()は非推奨です。buildAdminPanelUrl()を使用してください。');
  return buildAdminPanelUrl(userId);
}

/**
 * 後方互換性のためのレガシー関数
 * @deprecated generateUserUrls()を使用してください
 */
function generateAppUrls(userId) {
  console.warn('generateAppUrls()は非推奨です。generateUserUrls()を使用してください。');
  return generateUserUrls(userId);
}