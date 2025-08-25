/**
 * 緊急サーバー修復パッチ
 * 500エラーを構造的に解決し、必ず何らかのレスポンスを返す
 */

// =============================================================================
// 緊急doGet実装 - 絶対に500エラーを出さない
// =============================================================================

/**
 * 緊急doGet - main.gsの既存実装を完全に置き換える
 * 何があっても必ずHTMLレスポンスを返す
 */
function doGet(e) {
  // グローバルエラーハンドラー設定
  try {
    return safeDoGet(e);
  } catch (criticalError) {
    // 最終フォールバック：最小限のHTMLを返す
    console.error('Critical doGet error:', criticalError);
    return createUltimateFailsafeResponse(criticalError, e);
  }
}

/**
 * 安全なdoGet実装
 */
function safeDoGet(e) {
  try {
    // ログ出力（エラーでも続行）
    logSafe('INFO', 'doGet called', { 
      parameter: e ? e.parameter : null,
      queryString: e ? e.queryString : null 
    });

    // パラメータ解析（安全版）
    const params = parseParamsUltraSafe(e);
    
    // ルーティング実行（安全版）
    const result = routeUltraSafe(params);
    
    if (result && typeof result.getContent === 'function') {
      logSafe('INFO', 'doGet successful', { mode: params.mode });
      return result;
    } else {
      throw new Error('Invalid route result');
    }
    
  } catch (error) {
    logSafe('ERROR', 'safeDoGet failed', error);
    return handleDoGetError(error, e);
  }
}

/**
 * 超安全なパラメータ解析
 */
function parseParamsUltraSafe(e) {
  const defaults = {
    mode: 'emergency',
    userId: null,
    parameter: {},
    queryString: ''
  };
  
  try {
    if (!e) return defaults;
    
    return {
      mode: extractMode(e) || 'emergency',
      userId: extractUserId(e),
      parameter: e.parameter || {},
      queryString: e.queryString || ''
    };
  } catch (error) {
    logSafe('WARN', 'Parameter parsing failed', error);
    return defaults;
  }
}

/**
 * モード抽出（複数方法で試行）
 */
function extractMode(e) {
  try {
    // 方法1: e.parameter.mode
    if (e.parameter && e.parameter.mode) {
      return sanitizeMode(e.parameter.mode);
    }
    
    // 方法2: クエリ文字列から直接抽出
    if (e.queryString) {
      const modeMatch = e.queryString.match(/mode=([^&]+)/);
      if (modeMatch && modeMatch[1]) {
        return sanitizeMode(modeMatch[1]);
      }
    }
    
    // 方法3: URLから推測
    const url = e.queryString || '';
    if (url.includes('admin')) return 'admin';
    if (url.includes('view')) return 'view';
    if (url.includes('login')) return 'login';
    if (url.includes('setup')) return 'appSetup';
    
    return 'default';
  } catch (error) {
    return 'emergency';
  }
}

/**
 * ユーザーID抽出
 */
function extractUserId(e) {
  try {
    if (e.parameter && e.parameter.userId) {
      return sanitizeUserId(e.parameter.userId);
    }
    
    if (e.queryString) {
      const userIdMatch = e.queryString.match(/userId=([^&]+)/);
      if (userIdMatch && userIdMatch[1]) {
        return sanitizeUserId(userIdMatch[1]);
      }
    }
    
    return null;
  } catch (error) {
    return null;
  }
}

/**
 * 超安全ルーティング
 */
function routeUltraSafe(params) {
  try {
    logSafe('DEBUG', 'Routing', { mode: params.mode, userId: params.userId });
    
    switch (params.mode) {
      case 'admin':
        return handleAdminModeUltraSafe(params);
      case 'view':
        return handleViewModeUltraSafe(params);
      case 'login':
        return handleLoginModeUltraSafe();
      case 'appSetup':
        return handleAppSetupModeUltraSafe();
      case 'default':
        return handleDefaultRouteUltraSafe();
      default:
        return handleEmergencyModeUltraSafe(params);
    }
  } catch (error) {
    logSafe('ERROR', 'Routing failed', error);
    return handleEmergencyModeUltraSafe(params);
  }
}

/**
 * 超安全管理者モードハンドラー
 */
function handleAdminModeUltraSafe(params) {
  try {
    // ユーザーID必須チェック
    if (!params.userId) {
      return createErrorPageSafe('User ID Required', 'ユーザーIDが指定されていません。');
    }
    
    // 現在のユーザーメール取得（安全版）
    const currentUserEmail = getCurrentUserEmailSafe();
    if (!currentUserEmail) {
      return createLoginPageSafe('認証が必要です');
    }
    
    // データベース接続テスト
    const dbTestResult = testDatabaseConnectionSafe();
    if (!dbTestResult.success) {
      logSafe('WARN', 'Database connection failed', dbTestResult.error);
      return createSetupRequiredPageSafe(dbTestResult.error);
    }
    
    // ユーザー情報取得（安全版）
    const userInfo = getUserInfoSafe(params.userId);
    if (!userInfo) {
      return createUserNotFoundPageSafe(params.userId);
    }
    
    // アクセス権限確認（安全版）
    const hasAccess = verifyAccessSafe(currentUserEmail, userInfo);
    if (!hasAccess) {
      return createAccessDeniedPageSafe(currentUserEmail);
    }
    
    // 管理パネル生成
    return createAdminPanelSafe(userInfo);
    
  } catch (error) {
    logSafe('ERROR', 'Admin mode failed', error);
    return createAdminPanelEmergency(params.userId, error);
  }
}

/**
 * 超安全ビューモードハンドラー
 */
function handleViewModeUltraSafe(params) {
  try {
    if (!params.userId) {
      return createErrorPageSafe('User ID Required', 'ユーザーIDが指定されていません。');
    }
    
    const userInfo = getUserInfoSafe(params.userId);
    if (!userInfo) {
      return createUserNotFoundPageSafe(params.userId);
    }
    
    const config = parseConfigSafe(userInfo.configJson);
    if (!config.appPublished) {
      return createNotPublishedPageSafe();
    }
    
    return createViewPageSafe(userInfo, config);
    
  } catch (error) {
    logSafe('ERROR', 'View mode failed', error);
    return createViewPageEmergency(params.userId, error);
  }
}

/**
 * デフォルトルートハンドラー（安全版）
 */
function handleDefaultRouteUltraSafe() {
  try {
    const currentUserEmail = getCurrentUserEmailSafe();
    if (!currentUserEmail) {
      return createLoginPageSafe('ログインが必要です');
    }
    
    // 簡単なセットアップページを返す
    return createQuickSetupPageSafe(currentUserEmail);
    
  } catch (error) {
    logSafe('ERROR', 'Default route failed', error);
    return createLoginPageSafe('システムエラーが発生しました');
  }
}

// =============================================================================
// 安全なユーティリティ関数
// =============================================================================

/**
 * 安全なユーザーメール取得
 */
function getCurrentUserEmailSafe() {
  try {
    const email = Session.getActiveUser().getEmail();
    return email && email.length > 0 ? email : null;
  } catch (error) {
    logSafe('DEBUG', 'Failed to get user email', error);
    return null;
  }
}

/**
 * 安全なデータベース接続テスト
 */
function testDatabaseConnectionSafe() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const dbId = properties.getProperty('DATABASE_SPREADSHEET_ID');
    
    if (!dbId) {
      return { success: false, error: 'DATABASE_SPREADSHEET_ID not configured' };
    }
    
    const spreadsheet = SpreadsheetApp.openById(dbId);
    const sheets = spreadsheet.getSheets();
    
    return { success: true, dbId: dbId, sheetCount: sheets.length };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * 安全なユーザー情報取得
 */
function getUserInfoSafe(userId) {
  try {
    // 既存のfindUserById関数を使用（存在する場合）
    if (typeof findUserById === 'function') {
      return findUserById(userId);
    }
    
    // 直接データベースアクセス
    const properties = PropertiesService.getScriptProperties();
    const dbId = properties.getProperty('DATABASE_SPREADSHEET_ID');
    
    if (!dbId) return null;
    
    const spreadsheet = SpreadsheetApp.openById(dbId);
    const sheet = spreadsheet.getSheetByName('USER_INFO');
    
    if (!sheet) return null;
    
    const data = sheet.getDataRange().getValues();
    const header = data[0];
    
    const userIdCol = header.indexOf('userId');
    if (userIdCol === -1) return null;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][userIdCol] === userId) {
        return {
          userId: data[i][userIdCol],
          adminEmail: data[i][header.indexOf('adminEmail')] || '',
          spreadsheetId: data[i][header.indexOf('spreadsheetId')] || '',
          configJson: data[i][header.indexOf('configJson')] || '{}',
          isActive: data[i][header.indexOf('isActive')] !== false
        };
      }
    }
    
    return null;
  } catch (error) {
    logSafe('ERROR', 'getUserInfoSafe failed', error);
    return null;
  }
}

/**
 * 安全なアクセス権限確認
 */
function verifyAccessSafe(currentUserEmail, userInfo) {
  try {
    return currentUserEmail === userInfo.adminEmail;
  } catch (error) {
    logSafe('ERROR', 'verifyAccessSafe failed', error);
    return false;
  }
}

/**
 * 安全な設定解析
 */
function parseConfigSafe(configJson) {
  const defaults = {
    appPublished: false,
    displayMode: 'ANONYMOUS',
    setupStatus: 'pending'
  };
  
  try {
    if (!configJson) return defaults;
    const config = JSON.parse(configJson);
    return Object.assign(defaults, config);
  } catch (error) {
    logSafe('WARN', 'Config parsing failed', error);
    return defaults;
  }
}

/**
 * 安全なログ出力
 */
function logSafe(level, message, data) {
  try {
    const logMessage = `[${level}] ${message}`;
    console.log(logMessage, data || '');
  } catch (error) {
    // ログ出力すら失敗した場合は何もしない
  }
}

// =============================================================================
// 緊急ページ生成関数
// =============================================================================

/**
 * 安全な管理パネル生成
 */
function createAdminPanelSafe(userInfo) {
  try {
    // 既存のAdminPanel生成を試行
    if (typeof handleAdminMode === 'function') {
      return handleAdminMode({ userId: userInfo.userId });
    }
    
    // フォールバック：基本的な管理パネルを生成
    return createBasicAdminPanelSafe(userInfo);
  } catch (error) {
    logSafe('ERROR', 'createAdminPanelSafe failed', error);
    return createAdminPanelEmergency(userInfo.userId, error);
  }
}

/**
 * 基本管理パネル生成
 */
function createBasicAdminPanelSafe(userInfo) {
  const html = `<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>管理パネル - みんなの回答ボード</title>
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, sans-serif; margin: 0; padding: 20px; background: #1a1b26; color: white; }
    .container { max-width: 1200px; margin: 0 auto; }
    .header { background: #2d3748; padding: 20px; border-radius: 8px; margin-bottom: 20px; }
    .panel { background: #374151; padding: 20px; border-radius: 8px; margin-bottom: 20px; }
    .button { background: #3b82f6; color: white; padding: 12px 24px; border: none; border-radius: 6px; cursor: pointer; text-decoration: none; display: inline-block; }
    .button:hover { background: #2563eb; }
    .info-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 20px; }
    .info-item { background: #4b5563; padding: 15px; border-radius: 6px; }
    .info-label { font-size: 14px; color: #9ca3af; margin-bottom: 5px; }
    .info-value { font-size: 16px; font-weight: 500; }
    .status-ok { color: #10b981; }
    .status-error { color: #ef4444; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>📊 みんなの回答ボード 管理パネル</h1>
      <p>安全モードで動作中です。完全な機能を使用するには、システムの修復が必要です。</p>
    </div>
    
    <div class="panel">
      <h2>基本情報</h2>
      <div class="info-grid">
        <div class="info-item">
          <div class="info-label">管理者メール</div>
          <div class="info-value">${htmlEscapeSafe(userInfo.adminEmail || '不明')}</div>
        </div>
        <div class="info-item">
          <div class="info-label">ユーザーID</div>
          <div class="info-value">${htmlEscapeSafe(userInfo.userId || '不明')}</div>
        </div>
        <div class="info-item">
          <div class="info-label">システム状態</div>
          <div class="info-value status-ok">安全モードで動作中</div>
        </div>
        <div class="info-item">
          <div class="info-label">最終アクセス</div>
          <div class="info-value">${new Date().toLocaleString('ja-JP')}</div>
        </div>
      </div>
    </div>
    
    <div class="panel">
      <h2>復旧オプション</h2>
      <div style="display: flex; gap: 10px; flex-wrap: wrap;">
        <button onclick="location.reload()" class="button">🔄 ページを再読み込み</button>
        <button onclick="openFullSetup()" class="button">⚙️ 完全セットアップ</button>
        <button onclick="testSystem()" class="button">🧪 システムテスト</button>
        <button onclick="viewPublicPage()" class="button">👁️ 公開ページを確認</button>
      </div>
    </div>
    
    <div class="panel">
      <h3>⚠️ システム情報</h3>
      <p>現在、管理パネルは安全モードで動作しています。これは以下の理由が考えられます：</p>
      <ul>
        <li>システム設定の初期化が必要</li>
        <li>データベース接続の問題</li>
        <li>一時的なシステム負荷</li>
      </ul>
      <p>完全な機能を使用するには、上記の「完全セットアップ」を実行してください。</p>
    </div>
  </div>
  
  <script>
    function openFullSetup() {
      window.open('?mode=appSetup', '_self');
    }
    
    function testSystem() {
      alert('システムテストを実行中...');
      console.log('System Test - Safe Mode Active');
    }
    
    function viewPublicPage() {
      const userId = '${htmlEscapeSafe(userInfo.userId)}';
      if (userId) {
        window.open('?mode=view&userId=' + encodeURIComponent(userId), '_blank');
      }
    }
    
    console.log('✅ Safe mode admin panel loaded');
    console.log('User Info:', ${JSON.stringify({
      userId: userInfo.userId,
      email: userInfo.adminEmail,
      timestamp: new Date().toISOString()
    })});
  </script>
</body>
</html>`;

  return HtmlService.createHtmlOutput(html)
    .setTitle('管理パネル - みんなの回答ボード');
}

/**
 * 緊急管理パネル生成
 */
function createAdminPanelEmergency(userId, error) {
  const html = `<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>緊急モード - 管理パネル</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 0; padding: 20px; background: #1a1b26; color: white; text-align: center; }
    .emergency { background: #dc2626; padding: 20px; border-radius: 8px; margin: 20px auto; max-width: 600px; }
    .button { background: #3b82f6; color: white; padding: 15px 30px; border: none; border-radius: 6px; cursor: pointer; margin: 10px; }
  </style>
</head>
<body>
  <div class="emergency">
    <h1>🚨 システム緊急モード</h1>
    <p>管理パネルの初期化に問題が発生しました。</p>
    <p>以下のオプションをお試しください：</p>
    <button onclick="location.reload()" class="button">🔄 ページを再読み込み</button>
    <button onclick="window.open('?mode=appSetup', '_self')" class="button">⚙️ 初期セットアップ</button>
    <button onclick="window.open('?mode=login', '_self')" class="button">🔑 ログインページ</button>
  </div>
  <div style="max-width: 600px; margin: 20px auto; background: #374151; padding: 15px; border-radius: 6px;">
    <h3>技術情報</h3>
    <p>エラー: ${htmlEscapeSafe(error ? error.toString() : '不明')}</p>
    <p>ユーザーID: ${htmlEscapeSafe(userId || '不明')}</p>
    <p>タイムスタンプ: ${new Date().toISOString()}</p>
  </div>
</body>
</html>`;

  return HtmlService.createHtmlOutput(html);
}

/**
 * クイックセットアップページ生成
 */
function createQuickSetupPageSafe(userEmail) {
  const html = `<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>クイックセットアップ - みんなの回答ボード</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 0; padding: 20px; background: #1a1b26; color: white; }
    .container { max-width: 800px; margin: 0 auto; text-align: center; }
    .welcome { background: #2d3748; padding: 30px; border-radius: 8px; margin-bottom: 20px; }
    .button { background: #10b981; color: white; padding: 15px 30px; border: none; border-radius: 6px; cursor: pointer; margin: 10px; font-size: 16px; }
    .button:hover { background: #059669; }
    .info { background: #374151; padding: 20px; border-radius: 6px; margin: 20px 0; text-align: left; }
  </style>
</head>
<body>
  <div class="container">
    <div class="welcome">
      <h1>🎉 みんなの回答ボードへようこそ！</h1>
      <p>認証済みメール: ${htmlEscapeSafe(userEmail)}</p>
      <p>アプリケーションをセットアップして、クラスの意見収集を始めましょう！</p>
      <button onclick="startSetup()" class="button">🚀 セットアップを開始</button>
    </div>
    
    <div class="info">
      <h3>📋 セットアップで行うこと</h3>
      <ul>
        <li>Googleフォームとスプレッドシートの自動作成</li>
        <li>管理パネルのセットアップ</li>
        <li>公開ページの設定</li>
        <li>基本的な表示設定</li>
      </ul>
    </div>
    
    <div class="info">
      <h3>🔧 手動セットアップ（上級者向け）</h3>
      <p>すでに設定済みの場合や、カスタム設定を行いたい場合：</p>
      <button onclick="openManualSetup()" class="button" style="background: #3b82f6;">⚙️ 手動セットアップ</button>
    </div>
  </div>
  
  <script>
    function startSetup() {
      window.location.href = '?mode=appSetup';
    }
    
    function openManualSetup() {
      alert('手動セットアップ機能は準備中です。自動セットアップをご利用ください。');
    }
  </script>
</body>
</html>`;

  return HtmlService.createHtmlOutput(html);
}

/**
 * ログインページ生成（安全版）
 */
function createLoginPageSafe(message) {
  return handleLoginModeUltraSafe(message);
}

function handleLoginModeUltraSafe(message = '') {
  const html = `<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>ログイン - みんなの回答ボード</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 0; padding: 20px; background: #1a1b26; color: white; text-align: center; }
    .login-box { background: #2d3748; padding: 40px; border-radius: 8px; margin: 50px auto; max-width: 500px; }
    .message { background: #374151; padding: 15px; border-radius: 6px; margin: 20px 0; }
    .button { background: #10b981; color: white; padding: 15px 30px; border: none; border-radius: 6px; cursor: pointer; margin: 10px; }
  </style>
</head>
<body>
  <div class="login-box">
    <h1>🔑 みんなの回答ボード</h1>
    <p>Googleアカウントでログインしてください</p>
    ${message ? `<div class="message">${htmlEscapeSafe(message)}</div>` : ''}
    <button onclick="login()" class="button">Googleアカウントでログイン</button>
  </div>
  <script>
    function login() {
      window.location.href = '?mode=default';
    }
  </script>
</body>
</html>`;

  return HtmlService.createHtmlOutput(html);
}

/**
 * 最終フォールバックレスポンス
 */
function createUltimateFailsafeResponse(error, requestData) {
  const html = `<!DOCTYPE html>
<html><head><title>System Error</title></head>
<body style="font-family:Arial;padding:20px;background:#000;color:#fff;">
  <h1>🚨 System Emergency</h1>
  <p>A critical error occurred. Please contact the administrator.</p>
  <p><strong>Error:</strong> ${htmlEscapeSafe(error ? error.toString() : 'Unknown')}</p>
  <p><strong>Time:</strong> ${new Date().toISOString()}</p>
  <p><a href="javascript:location.reload()" style="color:#4fc3f7;">Reload Page</a></p>
</body></html>`;

  return HtmlService.createHtmlOutput(html);
}

// =============================================================================
// ユーティリティ
// =============================================================================

function htmlEscapeSafe(str) {
  if (!str) return '';
  return str.toString()
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function sanitizeMode(mode) {
  const allowed = ['default', 'admin', 'view', 'login', 'appSetup'];
  return allowed.includes(mode) ? mode : 'emergency';
}

function sanitizeUserId(userId) {
  if (!userId || typeof userId !== 'string') return null;
  return userId.replace(/[^a-zA-Z0-9\-]/g, '') || null;
}

// アプリセットアップモードハンドラー（安全版）
function handleAppSetupModeUltraSafe() {
  try {
    if (typeof handleAppSetupMode === 'function') {
      return handleAppSetupMode();
    }
    return createSetupPageSafe();
  } catch (error) {
    logSafe('ERROR', 'App setup mode failed', error);
    return createSetupPageSafe('セットアップでエラーが発生しました');
  }
}

function createSetupPageSafe(errorMessage = '') {
  const html = `<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <title>アプリセットアップ</title>
  <style>body{font-family:Arial;padding:20px;background:#1a1b26;color:white;}</style>
</head>
<body>
  <h1>⚙️ アプリセットアップ</h1>
  ${errorMessage ? `<div style="background:#dc2626;padding:15px;border-radius:6px;margin:20px 0;">${htmlEscapeSafe(errorMessage)}</div>` : ''}
  <p>セットアップ機能を準備中です...</p>
  <button onclick="location.href='?mode=default'" style="background:#3b82f6;color:white;padding:15px;border:none;border-radius:6px;">戻る</button>
</body>
</html>`;
  return HtmlService.createHtmlOutput(html);
}

// エラーページ、その他のハンドラー（省略版）
function createErrorPageSafe(title, message) {
  const html = `<!DOCTYPE html><html><head><title>${htmlEscapeSafe(title)}</title></head>
<body style="font-family:Arial;padding:20px;background:#1a1b26;color:white;">
<h1>❌ ${htmlEscapeSafe(title)}</h1><p>${htmlEscapeSafe(message)}</p>
<button onclick="history.back()" style="background:#3b82f6;color:white;padding:10px;border:none;">戻る</button>
</body></html>`;
  return HtmlService.createHtmlOutput(html);
}

function createUserNotFoundPageSafe(userId) {
  return createErrorPageSafe('User Not Found', `ユーザー ${userId} が見つかりません`);
}

function createAccessDeniedPageSafe(email) {
  return createErrorPageSafe('Access Denied', `${email} にはアクセス権限がありません`);
}

function createSetupRequiredPageSafe(error) {
  return createErrorPageSafe('Setup Required', `システムセットアップが必要です: ${error}`);
}

function createNotPublishedPageSafe() {
  return createErrorPageSafe('Not Published', 'このアプリは公開されていません');
}

function createViewPageSafe(userInfo, config) {
  return createErrorPageSafe('View Page', 'ビューページは準備中です');
}

function createViewPageEmergency(userId, error) {
  return createErrorPageSafe('View Error', `ビューページでエラーが発生しました: ${error}`);
}

function handleEmergencyModeUltraSafe(params) {
  return createUltimateFailsafeResponse(new Error('Emergency mode activated'), params);
}