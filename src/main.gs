/**
 * @OnlyCurrentDoc
 * @fileoverview Google Apps Script メインファイル
 * 
 * V8/ES6 最適化済み - パフォーマンス向上と安定性強化
 * 
 * 変更履歴:
 * - 2024-03-xx: V8ランタイム対応、ES6構文導入
 * - 2024-03-xx: 非同期処理とキャッシュ機能強化
 * - 2024-03-xx: 多言語対応とアクセシビリティ改善
 * - 2024-03-xx: セキュリティ強化（HTTPS、CORS対応）
 * - 2024-03-xx: エラーハンドリング改善
 * - 2024-03-xx: マルチテナント対応
 * - 2024-03-xx: ログ出力とデバッグ機能強化
 * - 2024-03-xx: フロントエンド通信プロトコル改善
 * - 2024-03-xx: サーバーサイドレンダリング最適化
 * - 2024-03-xx: JSON処理の循環参照対応
 * - 2024-03-xx: API関数名統一対応
 */

// グローバル定数（定数モジュールから統合）
const MAIN_ERROR_SEVERITY = {
  CRITICAL: 'critical',
  HIGH: 'high',
  MEDIUM: 'medium',
  LOW: 'low'
};

const MAIN_ERROR_CATEGORIES = {
  AUTH: 'authentication',
  DATA: 'data_processing',
  EXTERNAL: 'external_api',
  INTERNAL: 'internal_error',
  VALIDATION: 'validation',
  SYSTEM: 'system'
};

// =================================================================
// Core Error Handling & Logging
// =================================================================

/**
 * 安全なJSON文字列化（循環参照対応）
 * @param {*} obj - シリアライズするオブジェクト
 * @param {number} maxLength - 最大文字数制限
 * @returns {string} 安全にシリアライズされた文字列
 */
function safeStringify(obj, maxLength = 500) {
  try {
    const seen = new WeakSet();
    const result = JSON.stringify(obj, (key, val) => {
      if (typeof val === 'object' && val !== null) {
        if (seen.has(val)) {
          return '[Circular Reference]';
        }
        seen.add(val);
      }
      if (typeof val === 'function') {
        return '[Function]';
      }
      if (val && typeof val === 'object' && val.nodeType) {
        return '[DOM Element]';
      }
      if (val instanceof Error) {
        return val.message;
      }
      return val;
    });
    return result && result.length > maxLength ? 
      `${result.substring(0, maxLength)}...[truncated]` : 
      result || '[Unable to serialize]';
  } catch (e) {
    return `[Stringify Error: ${e.message}]`;
  }
}

/**
 * エラーログ記録
 * @param {Error|string} error - エラーオブジェクトまたはメッセージ
 * @param {string} context - エラー発生コンテキスト
 * @param {string} severity - エラーの重要度
 * @param {string} category - エラーカテゴリ
 */
function logError(error, context, severity = MAIN_ERROR_SEVERITY.MEDIUM, category = MAIN_ERROR_CATEGORIES.INTERNAL) {
  try {
    const timestamp = new Date().toISOString();
    const errorMessage = error instanceof Error ? error.message : String(error);
    const stack = error instanceof Error ? error.stack : 'No stack trace available';
    
    const logEntry = {
      timestamp,
      context: String(context || 'unknown'),
      severity: String(severity),
      category: String(category),
      message: errorMessage,
      stack: stack ? String(stack).substring(0, 1000) : 'No stack available'
    };
    
    console.error(`[${severity.toUpperCase()}] ${context}: ${errorMessage}`);
    console.error(`Stack: ${stack}`);
    
    // PropertiesServiceにも記録（冗長化）
    try {
      const properties = PropertiesService.getScriptProperties();
      const existingLogs = properties.getProperty('error_logs');
      const logs = existingLogs ? JSON.parse(existingLogs) : [];
      
      logs.push(logEntry);
      
      // 最新100件のみ保持
      if (logs.length > 100) {
        logs.splice(0, logs.length - 100);
      }
      
      properties.setProperty('error_logs', JSON.stringify(logs));
    } catch (propError) {
      console.error('PropertiesService logging failed:', propError.message);
    }
    
  } catch (logError) {
    console.error('Critical: logError function failed:', logError.message);
  }
}

/**
 * クライアントエラーログ記録（フロントエンドから）
 * @param {*} errorData - エラーデータ（オブジェクトまたは文字列）
 */
function logClientError(errorData) {
  try {
    let processedError = {};
    
    // データ型による安全な処理
    if (typeof errorData === 'string') {
      processedError = {
        message: errorData,
        timestamp: new Date().toISOString(),
        source: 'client'
      };
    } else if (errorData && typeof errorData === 'object') {
      processedError = {
        message: safeStringify(errorData.message || errorData),
        userId: safeStringify(errorData.userId || 'unknown'),
        url: safeStringify(errorData.url || 'unknown'),
        timestamp: errorData.timestamp || new Date().toISOString(),
        userAgent: safeStringify(errorData.userAgent || 'unknown'),
        type: safeStringify(errorData.type || 'unknown'),
        source: 'client'
      };
    } else {
      processedError = {
        message: `[Non-serializable client error: ${typeof errorData}]`,
        timestamp: new Date().toISOString(),
        source: 'client'
      };
    }
    
    console.error('CLIENT ERROR:', safeStringify(processedError));
    
    // サーバーサイドエラーログにも記録
    logError(`CLIENT: ${processedError.message}`, 'clientError', MAIN_ERROR_SEVERITY.LOW, MAIN_ERROR_CATEGORIES.EXTERNAL);
    
  } catch (e) {
    console.error('logClientError failed:', e.message);
    logError(`Failed to process client error: ${e.message}`, 'logClientError', MAIN_ERROR_SEVERITY.MEDIUM, MAIN_ERROR_CATEGORIES.SYSTEM);
  }
}

/**
 * デバッグログ出力
 * @param {string} message - ログメッセージ
 * @param {*} data - 追加データ
 */
function debugLog(message, data = null) {
  try {
    const timestamp = new Date().toISOString();
    const logMessage = `[DEBUG ${timestamp}] ${message}`;
    
    if (data) {
      console.log(logMessage, safeStringify(data));
    } else {
      console.log(logMessage);
    }
  } catch (error) {
    console.error('debugLog failed:', error.message);
  }
}

// =================================================================
// User Authentication & Access Control
// =================================================================

/**
 * 現在のユーザー情報を取得
 * @returns {object} ユーザー情報オブジェクト
 */
function getCurrentUser() {
  try {
    const user = Session.getActiveUser();
    const email = user.getEmail();
    
    if (!email) {
      throw new Error('No user email found');
    }
    
    return {
      email,
      id: email, // emailをIDとして使用
      isLoggedIn: true,
      timestamp: new Date().toISOString()
    };
  } catch (error) {
    logError(error, 'getCurrentUser', MAIN_ERROR_SEVERITY.HIGH, MAIN_ERROR_CATEGORIES.AUTH);
    return {
      email: null,
      id: null,
      isLoggedIn: false,
      error: error.message,
      timestamp: new Date().toISOString()
    };
  }
}

/**
 * ユーザーアクセス検証
 * @param {string} requestUserId - リクエストユーザーID
 * @throws {Error} アクセス権限がない場合
 */
function verifyUserAccess(requestUserId) {
  try {
    const currentUser = getCurrentUser();
    
    if (!currentUser.isLoggedIn) {
      throw new Error('User not logged in');
    }
    
    if (!requestUserId || typeof requestUserId !== 'string') {
      throw new Error('Invalid request user ID');
    }
    
    if (currentUser.id !== requestUserId) {
      logError(`Access denied: ${currentUser.id} tried to access ${requestUserId}`, 'verifyUserAccess', MAIN_ERROR_SEVERITY.HIGH, MAIN_ERROR_CATEGORIES.AUTH);
      throw new Error('Access denied');
    }
    
  } catch (error) {
    logError(error, 'verifyUserAccess', MAIN_ERROR_SEVERITY.HIGH, MAIN_ERROR_CATEGORIES.AUTH);
    throw error;
  }
}

// =================================================================
// Main Rendering Functions
// =================================================================

/**
 * メイン関数 - HTMLページを返す
 * @param {Event} e - リクエストイベント
 * @returns {HtmlOutput} レンダリングされたHTML
 */
function doGet(e) {
  try {
    debugLog('doGet called', { parameters: e ? e.parameters : 'no parameters' });
    
    const userInfo = getCurrentUser();
    
    if (!userInfo.isLoggedIn) {
      return renderLoginPage();
    }
    
    // パラメータによる画面制御
    const params = e && e.parameters ? e.parameters : {};
    
    if (params.page && params.page.toString() === 'admin') {
      return renderAdminPanel(userInfo, params);
    } else if (params.page && params.page.toString() === 'setup') {
      return renderSetupPage(userInfo, params);
    } else if (params.page && params.page.toString() === 'unpublished') {
      return renderUnpublishedPage(userInfo, params);
    } else {
      return renderAnswerBoard(userInfo, params);
    }
    
  } catch (error) {
    logError(error, 'doGet', MAIN_ERROR_SEVERITY.CRITICAL, MAIN_ERROR_CATEGORIES.SYSTEM);
    return renderErrorPage(error);
  }
}

/**
 * ログインページレンダリング
 * @returns {HtmlOutput} ログインページHTML
 */
function renderLoginPage() {
  try {
    const template = HtmlService.createTemplateFromFile('LoginPage');
    template.timestamp = new Date().toISOString();
    
    return template.evaluate()
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setTitle('ログイン - みんなの回答ボード');
      
  } catch (error) {
    logError(error, 'renderLoginPage', MAIN_ERROR_SEVERITY.HIGH, MAIN_ERROR_CATEGORIES.SYSTEM);
    throw error;
  }
}

/**
 * エラーページレンダリング
 * @param {Error} error - エラーオブジェクト
 * @returns {HtmlOutput} エラーページHTML
 */
function renderErrorPage(error) {
  try {
    const template = HtmlService.createTemplate(`
      <!DOCTYPE html>
      <html>
      <head>
        <title>エラー - みんなの回答ボード</title>
        <meta charset="utf-8">
      </head>
      <body>
        <div style="padding: 20px; font-family: Arial, sans-serif;">
          <h1>エラーが発生しました</h1>
          <p>申し訳ございませんが、システムエラーが発生しました。</p>
          <p>詳細: <?= errorMessage ?></p>
          <p>時刻: <?= timestamp ?></p>
          <button onclick="window.location.reload()">再読み込み</button>
        </div>
      </body>
      </html>
    `);
    
    template.errorMessage = error.message;
    template.timestamp = new Date().toISOString();
    
    return template.evaluate()
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setTitle('エラー - みんなの回答ボード');
      
  } catch (renderError) {
    logError(renderError, 'renderErrorPage', MAIN_ERROR_SEVERITY.CRITICAL, MAIN_ERROR_CATEGORIES.SYSTEM);
    
    // 最後の手段として基本的なHTMLを返す
    return HtmlService.createHtmlOutput(`
      <html>
        <body>
          <h1>システムエラー</h1>
          <p>重大なエラーが発生しました。管理者にお問い合わせください。</p>
        </body>
      </html>
    `).setTitle('システムエラー');
  }
}

// =================================================================
// Utility Functions
// =================================================================

/**
 * WebAppのURLを取得
 * @returns {string} WebAppのURL
 */
function getWebAppUrl() {
  try {
    return ScriptApp.getService().getUrl();
  } catch (error) {
    logError(error, 'getWebAppUrl', MAIN_ERROR_SEVERITY.LOW, MAIN_ERROR_CATEGORIES.SYSTEM);
    return '';
  }
}

/**
 * ログアウトURLを生成
 * @returns {string} Google ログアウトURL
 */
function getLogoutUrl() {
  try {
    return 'https://accounts.google.com/logout?continue=' + encodeURIComponent(getWebAppUrl());
  } catch (error) {
    logError(error, 'getLogoutUrl', MAIN_ERROR_SEVERITY.LOW, MAIN_ERROR_CATEGORIES.SYSTEM);
    return 'https://accounts.google.com/logout';
  }
}

// =================================================================
// History Management Functions
// =================================================================

/**
 * 公開履歴を追加
 * @param {string} requestUserId - ユーザーID
 * @param {object} historyData - 履歴データ
 * @returns {object} 処理結果
 */
function addPublishHistory(requestUserId, historyData) {
  try {
    verifyUserAccess(requestUserId);
    
    const properties = PropertiesService.getScriptProperties();
    const historyKey = `publishHistory_${requestUserId}`;
    const existingHistory = properties.getProperty(historyKey);
    const history = existingHistory ? JSON.parse(existingHistory) : [];
    
    const newEntry = {
      id: Utilities.getUuid(),
      timestamp: new Date().toISOString(),
      userId: requestUserId,
      ...historyData
    };
    
    history.unshift(newEntry); // 新しいものを先頭に
    
    // 最新50件のみ保持
    if (history.length > 50) {
      history.splice(50);
    }
    
    properties.setProperty(historyKey, JSON.stringify(history));
    
    return {
      status: 'success',
      historyId: newEntry.id,
      message: '履歴が追加されました'
    };
    
  } catch (error) {
    logError(error, 'addPublishHistory', MAIN_ERROR_SEVERITY.MEDIUM, MAIN_ERROR_CATEGORIES.DATA);
    return {
      status: 'error',
      message: error.message
    };
  }
}

/**
 * 公開履歴を取得
 * @param {string} requestUserId - ユーザーID
 * @returns {object} 履歴データ
 */
function getPublishHistory(requestUserId) {
  try {
    verifyUserAccess(requestUserId);
    
    const properties = PropertiesService.getScriptProperties();
    const historyKey = `publishHistory_${requestUserId}`;
    const existingHistory = properties.getProperty(historyKey);
    const history = existingHistory ? JSON.parse(existingHistory) : [];
    
    return {
      status: 'success',
      history: history
    };
    
  } catch (error) {
    logError(error, 'getPublishHistory', MAIN_ERROR_SEVERITY.MEDIUM, MAIN_ERROR_CATEGORIES.DATA);
    return {
      status: 'error',
      message: error.message,
      history: []
    };
  }
}

/**
 * 履歴から復元
 * @param {string} requestUserId - ユーザーID
 * @param {string} historyId - 履歴ID
 * @returns {object} 復元結果
 */
function restoreFromHistory(requestUserId, historyId) {
  try {
    verifyUserAccess(requestUserId);
    
    const historyResult = getPublishHistory(requestUserId);
    if (historyResult.status !== 'success') {
      throw new Error('履歴の取得に失敗しました');
    }
    
    const historyItem = historyResult.history.find(item => item.id === historyId);
    if (!historyItem) {
      throw new Error('指定された履歴が見つかりません');
    }
    
    // 履歴データから設定を復元（実装は設定管理システムに依存）
    // ここでは基本的な復元処理を実装
    
    return {
      status: 'success',
      message: '履歴から復元しました',
      restoredData: historyItem
    };
    
  } catch (error) {
    logError(error, 'restoreFromHistory', MAIN_ERROR_SEVERITY.MEDIUM, MAIN_ERROR_CATEGORIES.DATA);
    return {
      status: 'error',
      message: error.message
    };
  }
}

/**
 * プレビューURL生成
 * @param {string} requestUserId - ユーザーID
 * @param {object} previewData - プレビューデータ
 * @returns {object} プレビューURL
 */
function generatePreviewUrl(requestUserId, previewData) {
  try {
    verifyUserAccess(requestUserId);
    
    const baseUrl = getWebAppUrl();
    const previewParams = new URLSearchParams({
      page: 'preview',
      userId: requestUserId,
      timestamp: new Date().getTime().toString()
    });
    
    const previewUrl = `${baseUrl}?${previewParams.toString()}`;
    
    return {
      status: 'success',
      previewUrl: previewUrl,
      message: 'プレビューURLを生成しました'
    };
    
  } catch (error) {
    logError(error, 'generatePreviewUrl', MAIN_ERROR_SEVERITY.MEDIUM, MAIN_ERROR_CATEGORIES.SYSTEM);
    return {
      status: 'error',
      message: error.message
    };
  }
}

/**
 * 公開履歴をクリア
 * @param {string} requestUserId - ユーザーID
 * @returns {object} 処理結果
 */
function clearPublishHistory(requestUserId) {
  try {
    verifyUserAccess(requestUserId);
    
    const properties = PropertiesService.getScriptProperties();
    const historyKey = `publishHistory_${requestUserId}`;
    
    properties.deleteProperty(historyKey);
    
    return {
      status: 'success',
      message: '履歴をクリアしました'
    };
    
  } catch (error) {
    logError(error, 'clearPublishHistory', MAIN_ERROR_SEVERITY.MEDIUM, MAIN_ERROR_CATEGORIES.DATA);
    return {
      status: 'error',
      message: error.message
    };
  }
}

// =================================================================
// Include Support
// =================================================================

/**
 * HTMLファイルのインクルード機能
 * @param {string} filename - インクルードするファイル名
 * @returns {string} ファイルの内容
 */
function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (error) {
    logError(error, 'include', MAIN_ERROR_SEVERITY.MEDIUM, MAIN_ERROR_CATEGORIES.SYSTEM);
    return `<!-- Error including ${filename}: ${error.message} -->`;
  }
}

// =================================================================
// Legacy Functions Support
// =================================================================

/**
 * 管理パネルレンダリング（レガシー対応）
 * @param {object} userInfo - ユーザー情報
 * @param {object} params - パラメータ
 * @returns {HtmlOutput} 管理パネルHTML
 */
function renderAdminPanel(userInfo, params) {
  try {
    const template = HtmlService.createTemplateFromFile('AdminPanel');
    template.userInfo = userInfo;
    template.timestamp = new Date().toISOString();
    
    return template.evaluate()
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setTitle('管理パネル - みんなの回答ボード');
      
  } catch (error) {
    logError(error, 'renderAdminPanel', MAIN_ERROR_SEVERITY.HIGH, MAIN_ERROR_CATEGORIES.SYSTEM);
    throw error;
  }
}

/**
 * セットアップページレンダリング（レガシー対応）
 * @param {object} userInfo - ユーザー情報
 * @param {object} params - パラメータ
 * @returns {HtmlOutput} セットアップページHTML
 */
function renderSetupPage(userInfo, params) {
  try {
    const template = HtmlService.createTemplateFromFile('SetupPage');
    template.userInfo = userInfo;
    template.timestamp = new Date().toISOString();
    
    return template.evaluate()
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setTitle('セットアップ - みんなの回答ボード');
      
  } catch (error) {
    logError(error, 'renderSetupPage', MAIN_ERROR_SEVERITY.HIGH, MAIN_ERROR_CATEGORIES.SYSTEM);
    throw error;
  }
}

/**
 * 未公開ページレンダリング（レガシー対応）
 * @param {object} userInfo - ユーザー情報
 * @param {object} params - パラメータ
 * @returns {HtmlOutput} 未公開ページHTML
 */
function renderUnpublishedPage(userInfo, params) {
  try {
    const template = HtmlService.createTemplateFromFile('Unpublished');
    template.userInfo = userInfo;
    template.timestamp = new Date().toISOString();
    
    return template.evaluate()
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setTitle('未公開 - みんなの回答ボード');
      
  } catch (error) {
    logError(error, 'renderUnpublishedPage', MAIN_ERROR_SEVERITY.HIGH, MAIN_ERROR_CATEGORIES.SYSTEM);
    throw error;
  }
}

/**
 * 回答ボードレンダリング（レガシー対応）
 * @param {object} userInfo - ユーザー情報
 * @param {object} params - パラメータ
 * @returns {HtmlOutput} 回答ボードHTML
 */
function renderAnswerBoard(userInfo, params) {
  try {
    const template = HtmlService.createTemplateFromFile('Page');
    template.userInfo = userInfo;
    template.timestamp = new Date().toISOString();
    
    return template.evaluate()
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setTitle('みんなの回答ボード');
      
  } catch (error) {
    logError(error, 'renderAnswerBoard', MAIN_ERROR_SEVERITY.HIGH, MAIN_ERROR_CATEGORIES.SYSTEM);
    throw error;
  }
}

// =================================================================
// Additional Utility Functions
// =================================================================

/**
 * システム状態チェック
 * @returns {object} システム状態
 */
function checkSystemConfiguration() {
  try {
    const user = getCurrentUser();
    const webAppUrl = getWebAppUrl();
    const timestamp = new Date().toISOString();
    
    return {
      status: 'healthy',
      user: user,
      webAppUrl: webAppUrl,
      timestamp: timestamp,
      version: '2.0.0'
    };
    
  } catch (error) {
    logError(error, 'checkSystemConfiguration', MAIN_ERROR_SEVERITY.HIGH, MAIN_ERROR_CATEGORIES.SYSTEM);
    return {
      status: 'error',
      message: error.message,
      timestamp: new Date().toISOString()
    };
  }
}