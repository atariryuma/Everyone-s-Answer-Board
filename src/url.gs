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
  try {
    debugLog('🔗 generateUserUrls START:', {
      userId: userId,
      userIdType: typeof userId,
      userIdValid: !!userId,
      userIdLength: userId ? userId.length : 0
    });
    
    if (!userId) {
      errorLog('🚨 generateUserUrls: userId is required');
      return { adminUrl: '', viewUrl: '', error: 'userId_required' };
    }
    
    const baseUrl = getWebAppBaseUrl();
    debugLog('🔗 Base URL result:', {
      baseUrl: baseUrl,
      baseUrlType: typeof baseUrl,
      baseUrlValid: !!baseUrl,
      baseUrlLength: baseUrl ? baseUrl.length : 0
    });
    
    if (!baseUrl) {
      errorLog('🚨 generateUserUrls: Failed to get base URL');
      return { adminUrl: '', viewUrl: '', error: 'base_url_failed' };
    }
    
    const encodedUserId = encodeURIComponent(userId);
    debugLog('🔗 User ID encoding:', {
      original: userId,
      encoded: encodedUserId,
      encodingSuccessful: encodedUserId !== userId
    });
    
    const result = {
      adminUrl: `${baseUrl}?mode=admin&userId=${encodedUserId}`,
      viewUrl: `${baseUrl}?mode=view&userId=${encodedUserId}`
    };
    
    // URL検証
    const urlValidation = {
      adminUrlValid: result.adminUrl.startsWith('https://'),
      viewUrlValid: result.viewUrl.startsWith('https://'),
      adminUrlLength: result.adminUrl.length,
      viewUrlLength: result.viewUrl.length
    };
    
    debugLog('🔗 URL validation:', urlValidation);
    debugLog('✅ generateUserUrls RESULT:', result);
    
    // 最終検証
    if (!urlValidation.adminUrlValid || !urlValidation.viewUrlValid) {
      errorLog('🚨 Generated URLs are invalid');
      return { 
        adminUrl: '', 
        viewUrl: '', 
        error: 'invalid_urls_generated',
        details: urlValidation 
      };
    }
    
    return result;
    
  } catch (e) {
    errorLog('🚨 generateUserUrls: Unexpected error:', {
      message: e.message,
      stack: e.stack,
      userId: userId,
      userIdType: typeof userId
    });
    return { 
      adminUrl: '', 
      viewUrl: '', 
      error: 'unexpected_error',
      details: e.message 
    };
  }
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