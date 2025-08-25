/**
 * 構造的に堅牢なサーバーサイドシステム
 * エラーが起きない、起きても必ず復旧する設計
 */

// =============================================================================
// 堅牢なルーティングシステム - 必ず何らかのページを返す
// =============================================================================

/**
 * 堅牢なdoGet実装 - 絶対に失敗しない
 */
function robustDoGet(e) {
  const maxRetries = 3;
  let attempt = 0;
  
  while (attempt < maxRetries) {
    try {
      attempt++;
      logInfo(`doGet attempt ${attempt}/${maxRetries}`);
      
      // 基本的なルーティング処理を実行
      const result = processRequestWithFallback(e);
      
      if (result) {
        logInfo(`doGet succeeded on attempt ${attempt}`);
        return result;
      }
      
    } catch (error) {
      logError(`doGet attempt ${attempt} failed:`, error);
      
      // 最後の試行の場合は緊急ページを返す
      if (attempt >= maxRetries) {
        return createEmergencyPage(error, e);
      }
    }
  }
  
  // 理論上到達しないが、安全のため緊急ページを返す
  return createEmergencyPage(new Error('All retry attempts exhausted'), e);
}

/**
 * フォールバック付きリクエスト処理
 */
function processRequestWithFallback(e) {
  try {
    // 通常のルーティング処理
    const params = parseRequestParamsRobust(e);
    return routeRequestRobust(params);
    
  } catch (error) {
    logWarn('Normal routing failed, trying fallback routing:', error);
    
    // フォールバックルーティング
    return fallbackRouting(e, error);
  }
}

/**
 * 堅牢なパラメータ解析 - 必ず有効なオブジェクトを返す
 */
function parseRequestParamsRobust(e) {
  const defaultParams = {
    mode: 'default',
    userId: null,
    parameters: {},
    queryString: '',
    contentLength: 0
  };
  
  try {
    if (!e) return defaultParams;
    
    const params = {
      mode: (e.parameter && e.parameter.mode) || 'default',
      userId: (e.parameter && e.parameter.userId) || null,
      parameters: e.parameter || {},
      queryString: e.queryString || '',
      contentLength: e.contentLength || 0
    };
    
    // パラメータ値を検証・サニタイズ
    params.mode = sanitizeMode(params.mode);
    if (params.userId) {
      params.userId = sanitizeUserId(params.userId);
    }
    
    return params;
    
  } catch (error) {
    logError('Parameter parsing failed:', error);
    return defaultParams;
  }
}

/**
 * 堅牢なルーティング - 必ず適切なページを返す
 */
function routeRequestRobust(params) {
  const routeMap = {
    'default': () => handleDefaultRouteRobust(),
    'login': () => handleLoginModeRobust(),
    'appSetup': () => handleAppSetupModeRobust(),
    'admin': () => handleAdminModeRobust(params),
    'view': () => handleViewModeRobust(params)
  };
  
  try {
    const handler = routeMap[params.mode] || routeMap['default'];
    return handler();
    
  } catch (error) {
    logError(`Route handler failed for mode: ${params.mode}`, error);
    
    // フォールバック：デフォルトページを返す
    return handleDefaultRouteRobust();
  }
}

/**
 * フォールバックルーティング
 */
function fallbackRouting(e, originalError) {
  try {
    logWarn('Using fallback routing');
    
    // クエリ文字列から直接モードを抽出
    const queryMode = extractModeFromQuery(e);
    
    switch (queryMode) {
      case 'admin':
        return createBasicAdminPage();
      case 'view':
        return createBasicViewPage();
      case 'login':
        return createBasicLoginPage();
      default:
        return createBasicLandingPage();
    }
    
  } catch (error) {
    logError('Fallback routing failed:', error);
    return createEmergencyPage(error, e);
  }
}

/**
 * クエリ文字列からモード抽出
 */
function extractModeFromQuery(e) {
  try {
    const query = e.queryString || '';
    const modeMatch = query.match(/mode=([^&]+)/);
    return modeMatch ? modeMatch[1] : 'default';
  } catch (error) {
    return 'default';
  }
}

// =============================================================================
// 堅牢なハンドラー実装
// =============================================================================

/**
 * 堅牢なデフォルトルートハンドラー
 */
function handleDefaultRouteRobust() {
  try {
    // 現在のユーザーを安全に取得
    const userEmail = getCurrentUserEmailRobust();
    
    if (!userEmail) {
      return handleLoginModeRobust();
    }
    
    // ユーザー情報を安全に取得
    const userInfo = findUserByEmailRobust(userEmail);
    
    if (userInfo && userInfo.userId) {
      // 管理者モードにリダイレクト
      return redirectToAdminPanelRobust(userInfo.userId);
    } else {
      // 新規ユーザー：セットアップモードへ
      return handleAppSetupModeRobust();
    }
    
  } catch (error) {
    logError('Default route handling failed:', error);
    return createBasicLandingPage();
  }
}

/**
 * 堅牢な管理者モードハンドラー
 */
function handleAdminModeRobust(params) {
  try {
    if (!params.userId) {
      return createErrorPageRobust('ユーザーIDが指定されていません', 'パラメータエラー');
    }
    
    // ユーザー情報を安全に取得
    const userInfo = findUserByIdRobust(params.userId);
    
    if (!userInfo) {
      return createErrorPageRobust('ユーザーが見つかりません', 'アクセスエラー');
    }
    
    // アクセス権限を確認
    const hasAccess = verifyAdminAccessRobust(params.userId);
    
    if (!hasAccess) {
      return createErrorPageRobust('アクセス権限がありません', '認証エラー');
    }
    
    // 管理パネルを生成
    return createAdminPanelRobust(userInfo);
    
  } catch (error) {
    logError('Admin mode handling failed:', error);
    return createBasicAdminPage();
  }
}

/**
 * 堅牢なビューモードハンドラー
 */
function handleViewModeRobust(params) {
  try {
    if (!params.userId) {
      return createErrorPageRobust('ユーザーIDが指定されていません', 'パラメータエラー');
    }
    
    const userInfo = findUserByIdRobust(params.userId);
    
    if (!userInfo) {
      return createErrorPageRobust('ユーザーが見つかりません', 'データエラー');
    }
    
    // 公開状態を確認
    const config = parseUserConfigRobust(userInfo.configJson);
    
    if (!config.appPublished) {
      return createErrorPageRobust('このアプリは公開されていません', '公開設定エラー');
    }
    
    return createAnswerBoardRobust(userInfo, config);
    
  } catch (error) {
    logError('View mode handling failed:', error);
    return createBasicViewPage();
  }
}

// =============================================================================
// 堅牢なユーティリティ関数
// =============================================================================

/**
 * 堅牢なユーザーメール取得
 */
function getCurrentUserEmailRobust() {
  try {
    return Session.getActiveUser().getEmail();
  } catch (error) {
    logDebug('Failed to get current user email:', error);
    return null;
  }
}

/**
 * 堅牢なユーザー検索（メール）
 */
function findUserByEmailRobust(email) {
  if (!email) return null;
  
  try {
    const maxRetries = 3;
    
    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        const result = findUserByEmail(email);
        if (result) return result;
      } catch (error) {
        logDebug(`findUserByEmail attempt ${attempt} failed:`, error);
        if (attempt === maxRetries) throw error;
      }
    }
    
    return null;
    
  } catch (error) {
    logError('Robust user search by email failed:', error);
    return null;
  }
}

/**
 * 堅牢なユーザー検索（ID）
 */
function findUserByIdRobust(userId) {
  if (!userId) return null;
  
  try {
    const maxRetries = 3;
    
    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        const result = findUserById(userId);
        if (result) return result;
      } catch (error) {
        logDebug(`findUserById attempt ${attempt} failed:`, error);
        if (attempt === maxRetries) throw error;
      }
    }
    
    return null;
    
  } catch (error) {
    logError('Robust user search by ID failed:', error);
    return null;
  }
}

/**
 * 堅牢なアクセス権限確認
 */
function verifyAdminAccessRobust(userId) {
  try {
    return verifyAdminAccess(userId);
  } catch (error) {
    logError('Admin access verification failed:', error);
    // セキュリティ上、エラー時はfalse
    return false;
  }
}

/**
 * 堅牢な設定JSON解析
 */
function parseUserConfigRobust(configJson) {
  const defaultConfig = {
    setupStatus: 'pending',
    appPublished: false,
    displayMode: 'ANONYMOUS',
    showReactionCounts: false
  };
  
  try {
    if (!configJson) return defaultConfig;
    
    const config = JSON.parse(configJson);
    
    // 必須フィールドにデフォルト値を設定
    return Object.assign({}, defaultConfig, config);
    
  } catch (error) {
    logError('Config JSON parsing failed:', error);
    return defaultConfig;
  }
}

// =============================================================================
// 緊急ページ生成機能
// =============================================================================

/**
 * 緊急ページ生成
 */
function createEmergencyPage(error, requestData) {
  try {
    const errorMessage = error ? error.toString() : '不明なエラー';
    const timestamp = new Date().toISOString();
    
    const html = `<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>システム緊急モード - みんなの回答ボード</title>
  <style>
    body { font-family: Arial, sans-serif; background: #1a1b26; color: white; padding: 20px; }
    .container { max-width: 600px; margin: 0 auto; }
    .error-box { background: #dc2626; padding: 20px; border-radius: 8px; margin-bottom: 20px; }
    .info-box { background: #1f2937; padding: 20px; border-radius: 8px; margin-bottom: 20px; }
    .button { background: #3b82f6; color: white; padding: 10px 20px; text-decoration: none; border-radius: 4px; display: inline-block; margin: 5px; }
    .button:hover { background: #2563eb; }
    .code { background: #374151; padding: 10px; border-radius: 4px; font-family: monospace; font-size: 12px; }
  </style>
</head>
<body>
  <div class="container">
    <div class="error-box">
      <h1>🚨 システム緊急モード</h1>
      <p>システムの初期化に問題が発生しました。ご迷惑をおかけして申し訳ございません。</p>
    </div>
    
    <div class="info-box">
      <h2>対処方法</h2>
      <ol>
        <li>数分待ってからページを再読み込みしてください</li>
        <li>問題が続く場合は、管理者にお問い合わせください</li>
        <li>緊急の場合は、以下のリンクから基本機能にアクセスできます</li>
      </ol>
      
      <div style="margin-top: 20px;">
        <a href="javascript:location.reload()" class="button">🔄 ページを再読み込み</a>
        <a href="?mode=login" class="button">🔑 ログインページ</a>
        <a href="?mode=appSetup" class="button">⚙️ セットアップ</a>
      </div>
    </div>
    
    <div class="info-box">
      <h3>技術情報</h3>
      <div class="code">
        エラー: ${htmlEscape(errorMessage)}<br>
        タイムスタンプ: ${timestamp}<br>
        リクエストID: ${generateRequestId()}
      </div>
    </div>
  </div>
</body>
</html>`;
    
    return HtmlService.createHtmlOutput(html);
    
  } catch (emergencyError) {
    // 緊急ページの生成すら失敗した場合の最終フォールバック
    logError('Emergency page creation failed:', emergencyError);
    
    const fallbackHtml = `
      <!DOCTYPE html>
      <html><head><title>System Error</title></head>
      <body style="font-family:Arial;background:#000;color:#fff;padding:20px;">
        <h1>System Emergency Mode</h1>
        <p>Critical system error occurred. Please contact administrator.</p>
        <p><a href="javascript:location.reload()" style="color:#4fc3f7;">Reload Page</a></p>
      </body></html>`;
    
    return HtmlService.createHtmlOutput(fallbackHtml);
  }
}

/**
 * 基本管理ページ生成（フォールバック）
 */
function createBasicAdminPage() {
  try {
    const html = `<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>管理パネル（基本モード）</title>
  <style>
    body { font-family: Arial, sans-serif; background: #1a1b26; color: white; padding: 20px; }
    .container { max-width: 800px; margin: 0 auto; }
    .panel { background: #1f2937; padding: 20px; border-radius: 8px; margin-bottom: 20px; }
    .button { background: #3b82f6; color: white; padding: 10px 20px; text-decoration: none; border-radius: 4px; display: inline-block; margin: 5px; }
    .warning { background: #dc2626; padding: 15px; border-radius: 8px; margin-bottom: 20px; }
  </style>
</head>
<body>
  <div class="container">
    <div class="warning">
      <h2>⚠️ 基本モードで動作中</h2>
      <p>完全な管理機能にアクセスするには、システムの修復が必要です。</p>
    </div>
    
    <div class="panel">
      <h1>📊 みんなの回答ボード 管理パネル</h1>
      <p>基本的な管理機能のみ利用可能です。</p>
      
      <div style="margin-top: 20px;">
        <a href="javascript:location.reload()" class="button">🔄 完全モードに切り替え</a>
        <a href="?mode=appSetup" class="button">⚙️ システムセットアップ</a>
      </div>
    </div>
  </div>
</body>
</html>`;
    
    return HtmlService.createHtmlOutput(html);
    
  } catch (error) {
    logError('Basic admin page creation failed:', error);
    return createEmergencyPage(error, null);
  }
}

// =============================================================================
// ユーティリティ関数
// =============================================================================

function htmlEscape(str) {
  if (!str) return '';
  return str.toString()
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function generateRequestId() {
  return Date.now().toString(36) + Math.random().toString(36).substr(2, 5);
}

function sanitizeMode(mode) {
  const allowedModes = ['default', 'login', 'appSetup', 'admin', 'view'];
  return allowedModes.includes(mode) ? mode : 'default';
}

function sanitizeUserId(userId) {
  if (!userId || typeof userId !== 'string') return null;
  // 英数字とハイフンのみ許可
  return userId.replace(/[^a-zA-Z0-9\-_]/g, '') || null;
}

/**
 * 堅牢なdoGet - main.gsの既存doGetを置き換え
 */
function doGet(e) {
  return robustDoGet(e);
}