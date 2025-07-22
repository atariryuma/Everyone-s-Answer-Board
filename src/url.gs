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

    // 開発モードやテスト用の一時URLを検出して除外
    var devPatterns = [
      /^https:\/\/[a-z0-9-]+\.googleusercontent\.com\//, // 開発モード
      /\/userCodeAppPanel/, // テスト用パネル
      /\/dev$/, // 開発エンドポイント
      /\/test$/ // テストエンドポイント
    ];
    
    var isDevUrl = devPatterns.some(function(pattern) {
      return pattern.test(url);
    });
    
    if (isDevUrl) {
      console.warn('開発モードのURLを検出しました: ' + url + ' フォールバックURLを使用します');
      return getFallbackUrl();
    }

    // \"https://script.google.com/a/<domain>/macros/s/...\" 形式を
    // \"https://script.google.com/a/macros/<domain>/s/...\" に補正
    var wrongPattern = /^https:\/\/script\.google\.com\/a\/([^\/]+)\/macros\//;
    var match = url.match(wrongPattern);
    if (match) {
      url = url.replace(wrongPattern, 'https://script.google.com/a/macros/' + match[1] + '/');
    }

    // 有効なURLパターンかチェック
    var validPattern = /^https:\/\/script\.google\.com\/(a\/macros\/[^\/]+\/)?s\/[A-Za-z0-9_-]+\/(exec|dev)$/;
    if (!validPattern.test(url)) {
      console.warn('無効なURLパターンを検出しました: ' + url + ' フォールバックURLを使用します');
      return getFallbackUrl();
    }

    return url;
  } catch (e) {
    console.error('WebアプリURL取得エラー: ' + e.message);
    return getFallbackUrl();
  }
}

function getWebAppUrlCached() {
  try {
    // 統合キャッシュマネージャーを使用してURL取得・生成・キャッシュを一元化
    var webAppUrl = cacheManager.get(URL_CACHE_KEY, () => {
      console.log('🔍 WebAppURL キャッシュmiss - 新規生成開始');
      
      // 新しいURLを生成
      var freshUrl = ScriptApp.getService().getUrl();
      
      // 開発URLの検証
      if (freshUrl.includes('googleusercontent.com') ||
          freshUrl.includes('userCodeAppPanel') ||
          freshUrl.endsWith('/dev')) {
        console.warn('⚠️ 開発URLが検出されました、キャッシュしません:', freshUrl);
        return null; // 開発URLはキャッシュしない
      }
      
      console.log('✅ 有効なWebAppURL生成:', freshUrl);
      return freshUrl;
    }, { 
      ttl: 3600, // 1時間キャッシュ
      enableMemoization: true 
    });

    // キャッシュされたURLの検証（既存URLが開発URLになっていないかチェック）
    if (webAppUrl && (webAppUrl.includes('googleusercontent.com') ||
        webAppUrl.includes('userCodeAppPanel') ||
        webAppUrl.endsWith('/dev'))) {
      console.warn('⚠️ キャッシュされたURLが開発URLです、クリアして再生成:', webAppUrl);
      cacheManager.remove(URL_CACHE_KEY);
      // 再帰的に呼び出して新しいURLを生成
      return getWebAppUrlCached();
    }

    if (webAppUrl) {
      console.log('✅ 統合キャッシュから有効URL取得:', webAppUrl);
      return webAppUrl;
    }

    // フォールバック: 統合キャッシュマネージャーが失敗した場合の直接生成
    console.warn('⚠️ 統合キャッシュが利用できません、直接URL生成に切り替え');
    var currentUrl = computeWebAppUrl();
    
    if (currentUrl && !currentUrl.includes('googleusercontent.com') && !currentUrl.includes('userCodeAppPanel')) {
      console.log('✅ 新規URL生成成功（キャッシュなし）:', currentUrl);
      return currentUrl;
    } else {
      console.warn('⚠️ 無効なURL生成、フォールバックURLを使用:', currentUrl);
      return getFallbackUrl();
    }
    
  } catch (e) {
    console.error('❌ getWebAppUrlCached critical error:', e.message);
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
    console.error('フォールバックURL生成エラー: ' + e.message);
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
    console.log('URL cache cleared successfully');
    
    // 新しいURLを即座に生成してキャッシュ
    var newUrl = computeWebAppUrl();
    if (newUrl && !newUrl.includes('googleusercontent.com') && !newUrl.includes('userCodeAppPanel')) {
      cache.put(URL_CACHE_KEY, newUrl, URL_CACHE_TTL);
      console.log('New URL cached:', newUrl);
    }
    
    return newUrl;
  } catch (e) {
    console.error('clearUrlCache error:', e.message);
    return getFallbackUrl();
  }
}

/**
 * 強制的にURLシステムをリセット（公開API）
 * フロントエンドから呼び出し可能
 */
function forceUrlSystemReset() {
  try {
    console.log('Forcing URL system reset...');
    
    // 全てのURLキャッシュをクリア
    var cache = CacheService.getScriptCache();
    cache.remove(URL_CACHE_KEY);
    
    // 新しいURLを生成
    var newUrl = computeWebAppUrl();
    console.log('New URL generated:', newUrl);
    
    // 開発URLチェック
    if (newUrl && (newUrl.includes('googleusercontent.com') || newUrl.includes('userCodeAppPanel'))) {
      console.warn('Development URL detected, using fallback');
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
    console.error('forceUrlSystemReset error:', e.message);
    return {
      status: 'error',
      message: 'URLシステムリセットに失敗しました: ' + e.message
    };
  }
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
    
    // 最終的なURL検証を複数回実行
    var maxRetries = 3;
    for (var i = 0; i < maxRetries; i++) {
      if (!webAppUrl || webAppUrl.includes('googleusercontent.com') || webAppUrl.includes('userCodeAppPanel')) {
        console.warn('無効なURLが返されました（試行 ' + (i + 1) + '/' + maxRetries + '）: ' + webAppUrl);
        
        // キャッシュをクリアして再取得
        webAppUrl = clearUrlCache();
        
        if (i < maxRetries - 1) {
          // 再試行
          webAppUrl = computeWebAppUrl();
        } else {
          // 最後の試行でもダメな場合はフォールバック
          webAppUrl = getFallbackUrl();
        }
      } else {
        break;
      }
    }
    
    // 最終チェック: まだ開発URLが含まれている場合は強制的にフォールバック
    if (webAppUrl && (webAppUrl.includes('googleusercontent.com') || webAppUrl.includes('userCodeAppPanel'))) {
      console.error('開発URLが最終チェックで検出されました。フォールバックURLを使用します: ' + webAppUrl);
      webAppUrl = getFallbackUrl();
    }
    
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
      adminUrl: webAppUrl + '?userId=' + encodedUserId + '&mode=admin',
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

/**
 * 指定されたURLへサーバーサイドでリダイレクトします。
 * @param {string} url リダイレクト先のURL
 */
function redirectToUrl(url) {
  return HtmlService.createHtmlOutput('<script>window.top.location.href="' + url + '";</script>');
}

