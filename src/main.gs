/**
 * @fileoverview メインエントリーポイント - 2024年最新技術の結集
 * V8ランタイム、最新パフォーマンス技術、安定性強化を統合
 */

/**
 * HTML ファイルを読み込む include ヘルパー
 * @param {string} path ファイルパス
 * @return {string} HTML content
 */
function include(path) {
  try {
    const tmpl = HtmlService.createTemplateFromFile(path);
    tmpl.include = include;
    return tmpl.evaluate().getContent();
  } catch (error) {
    logError(error, 'includeFile', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.SYSTEM, { filePath: path });
    return `<!-- Error including ${path}: ${error.message} -->`;
  }
}

/**
 * JavaScript文字列エスケープ関数 (URL対応版)
 * @param {string} str エスケープする文字列
 * @return {string} エスケープされた文字列
 */
function escapeJavaScript(str) {
  if (!str) return '';

  const strValue = str.toString();

  // URL判定: HTTP/HTTPSで始まり、すでに適切にエスケープされている場合は最小限の処理
  if (strValue.match(/^https?:\/\/[^\s<>"']+$/)) {
    // URLの場合はバックスラッシュと改行文字のみエスケープ
    return strValue
      .replace(/\\/g, '\\\\')
      .replace(/\n/g, '\\n')
      .replace(/\r/g, '\\r')
      .replace(/\t/g, '\\t');
  }

  // 通常のテキストの場合は従来通りの完全エスケープ
  return strValue
    .replace(/\\/g, '\\\\')
    .replace(/'/g, "\\'")
    .replace(/"/g, '\\"')
    .replace(/\n/g, '\\n')
    .replace(/\r/g, '\\r')
    .replace(/\t/g, '\\t');
}

// グローバル定数の定義（書き換え不可）
const SCRIPT_PROPS_KEYS = {
  SERVICE_ACCOUNT_CREDS: 'SERVICE_ACCOUNT_CREDS',
  DATABASE_SPREADSHEET_ID: 'DATABASE_SPREADSHEET_ID',
  ADMIN_EMAIL: 'ADMIN_EMAIL'
};

const DB_SHEET_CONFIG = {
  SHEET_NAME: 'Users',
  HEADERS: [
    'userId', 'adminEmail', 'spreadsheetId', 'spreadsheetUrl',
    'createdAt', 'configJson', 'lastAccessedAt', 'isActive'
  ]
};

const LOG_SHEET_CONFIG = {
  SHEET_NAME: 'Logs',
  HEADERS: ['timestamp', 'userId', 'action', 'details']
};

// 実行中のユーザー情報キャッシュ（パフォーマンス最適化用）
let _executionUserInfoCache = null;

/**
 * 実行中のユーザー情報キャッシュをクリア
 */
function clearExecutionUserInfoCache() {
  _executionUserInfoCache = null;

  // 統一キャッシュマネージャーの関連エントリもクリア
  if (typeof cacheManager !== 'undefined' && cacheManager) {
    try {
      // セッション関連キャッシュのクリア
      const currentEmail = getCurrentUserEmail();
      if (currentEmail) {
        cacheManager.remove('session_' + currentEmail);
      }

      debugLog('[Memory] 実行レベル + 統一キャッシュの関連エントリをクリアしました');
    } catch (error) {
      debugLog('[Memory] 統一キャッシュクリア中にエラー:', error.message);
    }
  } else {
    debugLog('[Memory] 実行レベルユーザー情報キャッシュをクリアしました');
  }
}

const COLUMN_HEADERS = {
  TIMESTAMP: 'タイムスタンプ',
  EMAIL: 'メールアドレス',
  CLASS: 'クラス',
  OPINION: '回答',
  REASON: '理由',
  NAME: '名前',
  UNDERSTAND: 'なるほど！',
  LIKE: 'いいね！',
  CURIOUS: 'もっと知りたい！',
  HIGHLIGHT: 'ハイライト'
};

// 表示モード定数
const DISPLAY_MODES = {
  ANONYMOUS: 'anonymous',
  NAMED: 'named'
};

// リアクション関連定数
const REACTION_KEYS = ['UNDERSTAND', 'LIKE', 'CURIOUS'];

// スコア計算設定
const SCORING_CONFIG = {
  LIKE_MULTIPLIER_FACTOR: 0.1,
  RANDOM_SCORE_FACTOR: 0.01
};

/**
 * 自動停止チェック関数
 * 公開期限をチェックし、期限切れの場合は自動的に非公開に変更
 * @param {Object} config - ユーザーのconfig情報
 * @param {Object} userInfo - ユーザー情報
 * @return {boolean} 自動停止実行の有無
 */
function checkAndHandleAutoStop(config, userInfo) {
  // appPublishedがfalseなら既に非公開
  if (!config.appPublished) {
    return false;
  }

  // 自動停止が無効、または必要な情報がない場合はスキップ
  if (!config.autoStopEnabled || !config.scheduledEndAt) {
    debugLog('🔍 自動停止チェック: 無効またはデータ不足', {
      autoStopEnabled: config.autoStopEnabled,
      hasScheduledEndAt: !!config.scheduledEndAt
    });
    return false;
  }

  const scheduledEndTime = new Date(config.scheduledEndAt);
  const now = new Date();

  debugLog('🔍 自動停止チェック:', {
    scheduledEndAt: config.scheduledEndAt,
    now: now.toISOString(),
    isOverdue: now >= scheduledEndTime
  });

  // 期限切れチェック
  if (now >= scheduledEndTime) {
    warnLog('⚠️ 期限切れ検出 - 自動停止を実行します');

    // 自動停止前に履歴を保存
    try {
      saveHistoryOnAutoStop(config, userInfo);
    } catch (historyError) {
      logError(historyError, 'autoStopHistorySave', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.DATABASE);
      // 履歴保存エラーは処理を継続
    }

    // 自動停止実行
    config.appPublished = false;
    config.autoStoppedAt = now.toISOString();
    config.autoStopReason = 'scheduled_timeout';

    try {
      // データベース更新
      updateUser(userInfo.userId, {
        configJson: JSON.stringify(config)
      });

      infoLog(`🔄 自動停止実行完了: ${userInfo.adminEmail} (期限: ${config.scheduledEndAt})`);
      return true; // 自動停止実行済み
    } catch (error) {
      logError(error, 'autoStopProcess', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.SYSTEM);
      return false;
    }
  }

  debugLog('✅ まだ期限内です');
  return false; // まだ期限内
}

/**
 * 自動停止時の履歴保存関数
 * @param {Object} config - ユーザーのconfig情報
 * @param {Object} userInfo - ユーザー情報
 */
function saveHistoryOnAutoStop(config, userInfo) {
  debugLog('💾 自動停止時履歴保存開始');

  // 履歴アイテムを作成
  const historyItem = {
    id: 'auto_' + Date.now(),
    questionText: getQuestionTextFromConfig(config, userInfo),
    sheetName: config.publishedSheetName || '',
    publishedAt: config.publishedAt || config.lastPublishedAt || new Date().toISOString(),
    endTime: new Date().toISOString(), // 実際の終了日時
    scheduledEndTime: config.scheduledEndAt || null, // 予定終了日時
    answerCount: getAnswerCountFromSheet(config, userInfo),
    reactionCount: 0, // 自動停止時はリアクション数取得を省略
    config: config,
    formUrl: userInfo.formUrl || config.formUrl || '',
    spreadsheetUrl: userInfo.spreadsheetUrl || '',
    setupType: determineSetupTypeFromConfig(config, userInfo),
    isActive: false,
    endReason: 'auto_timeout'
  };

  // サーバーサイドでの履歴保存（スプレッドシート）
  try {
    saveHistoryToSheet(historyItem, userInfo);
    infoLog('✅ 自動停止履歴保存完了:', historyItem.questionText);
  } catch (error) {
    logError(error, 'serverSideHistorySave', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.DATABASE);
  }
}

/**
 * configからメイン質問文を取得
 * @param {Object} config - config情報
 * @param {Object} userInfo - ユーザー情報
 * @return {string} 質問文
 */
function getQuestionTextFromConfig(config, userInfo) {
  // 1. sheet固有設定から取得
  if (config.publishedSheetName) {
    const sheetConfigKey = `sheet_${config.publishedSheetName}`;
    const sheetConfig = config[sheetConfigKey];
    if (sheetConfig && sheetConfig.opinionHeader) {
      return sheetConfig.opinionHeader;
    }
  }

  // 2. グローバル設定から取得
  if (config.opinionHeader) {
    return config.opinionHeader;
  }

  // 3. カスタムフォーム情報から取得
  if (userInfo.customFormInfo) {
    try {
      const customInfo = typeof userInfo.customFormInfo === 'string'
        ? JSON.parse(userInfo.customFormInfo)
        : userInfo.customFormInfo;
      if (customInfo.mainQuestion) {
        return customInfo.mainQuestion;
      }
    } catch (e) {
      warnLog('customFormInfo パースエラー:', e);
    }
  }

  return '（問題文未設定）';
}

/**
 * シートから回答数を取得
 * @param {Object} config - config情報
 * @param {Object} userInfo - ユーザー情報
 * @return {number} 回答数
 */
function getAnswerCountFromSheet(config, userInfo) {
  try {
    if (!userInfo.spreadsheetId || !config.publishedSheetName) {
      return 0;
    }

    const sheet = openSpreadsheetOptimized(userInfo.spreadsheetId).getSheetByName(config.publishedSheetName);
    if (!sheet) {
      return 0;
    }

    const lastRow = sheet.getLastRow();
    return Math.max(0, lastRow - 1); // ヘッダー行を除外

  } catch (error) {
    warnLog('回答数取得エラー:', error);
    return 0;
  }
}

/**
 * セットアップタイプを判定
 * @param {Object} config - config情報
 * @param {Object} userInfo - ユーザー情報
 * @return {string} セットアップタイプ
 */
function determineSetupTypeFromConfig(config, userInfo) {
  if (userInfo.customFormInfo) {
    return 'カスタムセットアップ';
  } else if (config.isQuickStart || config.setupType === 'quickstart') {
    return 'クイックスタート';
  } else if (config.isExternalResource) {
    return '外部リソース';
  } else {
    return 'unknown';
  }
}

/**
 * 履歴をスプレッドシートに保存
 * @param {Object} historyItem - 履歴アイテム
 * @param {Object} userInfo - ユーザー情報
 */
function saveHistoryToSheet(historyItem, userInfo) {
  debugLog('📋 サーバーサイド履歴保存開始:', historyItem.questionText);

  try {
    if (!userInfo || !userInfo.userId) {
      throw new Error('ユーザー情報が不正です');
    }

    // 既存のユーザー情報を取得
    const existingUser = findUserById(userInfo.userId);
    if (!existingUser) {
      throw new Error('ユーザーが見つかりません: ' + userInfo.userId);
    }

    // 現在のconfigJsonを取得・解析
    let configJson = getConfigJSON(existingUser);

    // 履歴配列を取得または初期化
    if (!Array.isArray(configJson.historyArray)) {
      configJson.historyArray = [];
    }

    // 新しい履歴アイテムを作成
    const serverHistoryItem = {
      id: historyItem.id || ('server_' + Date.now()),
      timestamp: new Date().toISOString(),
      questionText: historyItem.questionText || '（問題文未設定）',
      sheetName: historyItem.sheetName || '',
      publishedAt: historyItem.publishedAt || new Date().toISOString(),
      endTime: historyItem.endTime || new Date().toISOString(),
      scheduledEndTime: historyItem.scheduledEndTime || null,
      answerCount: historyItem.answerCount || 0,
      reactionCount: historyItem.reactionCount || 0,
      endReason: historyItem.endReason || 'manual',
      savedAt: new Date().toISOString()
    };

    // 履歴配列の先頭に追加
    configJson.historyArray.unshift(serverHistoryItem);

    // 最大50件まで保持
    if (configJson.historyArray.length > 50) {
      configJson.historyArray.splice(50);
    }

    // configJsonを更新
    configJson.lastModified = new Date().toISOString();

    // データベースに保存
    const updateResult = updateUser(userInfo.userId, {
      configJson: JSON.stringify(configJson),
      lastAccessedAt: new Date().toISOString()
    });

    if (updateResult.status === 'success') {
      infoLog('✅ サーバーサイド履歴保存完了:', {
        userId: userInfo.userId,
        questionText: serverHistoryItem.questionText,
        historyCount: configJson.historyArray.length
      });
    } else {
      throw new Error('データベース更新に失敗: ' + updateResult.message);
    }

  } catch (error) {
    logError(error, 'serverSideHistorySave', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.DATABASE);
    // エラーをログに記録するが、メイン処理は継続
  }
}

/**
 * 履歴をサーバーサイドに保存する（認証付きAPI）
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @param {Object} historyItem - 履歴アイテム
 * @returns {Object} 保存結果
 */
function saveHistoryToSheetAPI(requestUserId, historyItem) {
  try {
    verifyUserAccess(requestUserId);

    // ユーザー情報を取得
    const userInfo = findUserById(requestUserId);
    if (!userInfo) {
      throw new Error('ユーザーが見つかりません');
    }

    // 履歴保存を実行
    saveHistoryToSheet(historyItem, userInfo);

    return {
      status: 'success',
      message: '履歴がサーバーに保存されました',
      timestamp: new Date().toISOString()
    };

  } catch (error) {
    logError(error, 'saveHistoryToSheetAPI', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.DATABASE);
    return {
      status: 'error',
      message: error.message
    };
  }
}

/**
 * サーバーサイドから履歴を取得する（認証付きAPI）
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @returns {Object} 履歴データ
 */
function getHistoryFromServerAPI(requestUserId) {
  try {
    verifyUserAccess(requestUserId);

    // ユーザー情報を取得
    const userInfo = findUserById(requestUserId);
    if (!userInfo) {
      throw new Error('ユーザーが見つかりません');
    }

    // configJsonから履歴を取得
    let configJson;
    try {
      configJson = getConfigJSON(userInfo);
    } catch (parseError) {
      warnLog('configJson解析エラー:', parseError.message);
      configJson = {};
    }

    const historyArray = Array.isArray(configJson.historyArray) ? configJson.historyArray : [];

    return {
      status: 'success',
      historyArray: historyArray,
      count: historyArray.length,
      lastModified: configJson.lastModified || null
    };

  } catch (error) {
    logError(error, 'getHistoryFromServerAPI', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.DATABASE);
    return {
      status: 'error',
      message: error.message,
      historyArray: []
    };
  }
}

/**
 * サーバーサイドの履歴をクリアする（認証付きAPI）
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @returns {Object} クリア結果
 */
function clearHistoryFromServerAPI(requestUserId) {
  try {
    verifyUserAccess(requestUserId);

    // ユーザー情報を取得
    const userInfo = findUserById(requestUserId);
    if (!userInfo) {
      throw new Error('ユーザーが見つかりません');
    }

    // configJsonから履歴をクリア
    let configJson;
    try {
      configJson = getConfigJSON(userInfo);
    } catch (parseError) {
      warnLog('configJson解析エラー、新規作成します:', parseError.message);
      configJson = {};
    }

    // 履歴配列をクリア
    configJson.historyArray = [];
    configJson.lastModified = new Date().toISOString();

    // データベースに保存
    const updateResult = updateUser(requestUserId, {
      configJson: JSON.stringify(configJson),
      lastAccessedAt: new Date().toISOString()
    });

    if (updateResult.status === 'success') {
      infoLog('✅ サーバーサイド履歴クリア完了:', requestUserId);
      return {
        status: 'success',
        message: 'サーバーサイドの履歴をクリアしました',
        timestamp: new Date().toISOString()
      };
    } else {
      throw new Error('データベース更新に失敗: ' + updateResult.message);
    }

  } catch (error) {
    logError(error, 'clearHistoryFromServerAPI', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.DATABASE);
    return {
      status: 'error',
      message: error.message
    };
  }
}

const EMAIL_REGEX = /^[^\n@]+@[^\n@]+\.[^\n@]+$/;
const DEBUG = PropertiesService.getScriptProperties()
  .getProperty('DEBUG_MODE') === 'true';

/**
 * Determine if a value represents boolean true.
 * Accepts boolean true, 'true', or 'TRUE'.
 * @param {any} value
 * @returns {boolean}
 */
function isTrue(value) {
  if (typeof value === 'boolean') return value === true;
  if (value === undefined || value === null) return false;
  return String(value).toLowerCase() === 'true';
}

/**
 * HTMLエスケープ関数（Utilities.htmlEncodeの代替）
 * @param {string} text - エスケープするテキスト
 * @returns {string} エスケープされたテキスト
 */
function htmlEncode(text) {
  if (!text) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * HtmlOutputに安全にX-Frame-Optionsヘッダーを設定するユーティリティ
 * @param {HtmlOutput} htmlOutput - 設定対象のHtmlOutput
 * @returns {HtmlOutput} 設定後のHtmlOutput
 */
function safeSetXFrameOptionsDeny(htmlOutput) {
  try {
    if (htmlOutput && typeof htmlOutput.setXFrameOptionsMode === 'function' &&
        HtmlService && HtmlService.XFrameOptionsMode &&
        HtmlService.XFrameOptionsMode.DENY) {
      htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DENY);
    }
  } catch (e) {
    warnLog('Failed to set XFrameOptionsMode:', e.message);
  }
  return htmlOutput;
}

// getSecurityHeaders function removed - not used in current implementation

// 安定性を重視してconstを使用
const ULTRA_CONFIG = {
  EXECUTION_LIMITS: {
    MAX_TIME: 330000, // 5.5分（安全マージン）
    BATCH_SIZE: 100,
    API_RATE_LIMIT: 90 // 100秒間隔での制限
  },

  CACHE_STRATEGY: {
    L1_TTL: 300,     // Level 1: 5分
    L2_TTL: 3600,    // Level 2: 1時間
    L3_TTL: 21600    // Level 3: 6時間（最大）
  }
};

/**
 * 簡素化されたエラーハンドリング関数群
 */

// ログ出力の最適化
function log(level, message, details) {
  try {
    if (typeof globalProfiler !== 'undefined') {
      globalProfiler.start('logging');
    }

    if (level === 'info' && !DEBUG) {
      return;
    }

    switch (level) {
      case 'error':
        logError(message, 'debugLog', ERROR_SEVERITY.LOW, ERROR_CATEGORIES.SYSTEM, { details });
        break;
      case 'warn':
        warnLog(message, details || '');
        break;
      default:
        debugLog(message, details || '');
    }

    if (typeof globalProfiler !== 'undefined') {
      globalProfiler.end('logging');
    }
  } catch (e) {
    // ログ出力自体が失敗した場合は無視
  }
}

function debugLog() {
  if (!DEBUG) return;
  try {
    debugLog.apply(null, arguments);
  } catch (e) {
    // ignore logging errors
  }
}

/**
 * デプロイされたWebアプリのドメイン情報と現在のユーザーのドメイン情報を取得
 * AdminPanel.html と Login.html から共通で呼び出される
 */
function getDeployUserDomainInfo() {
  try {
    const activeUserEmail = getCurrentUserEmail();
    const currentDomain = getEmailDomain(activeUserEmail);

    // 統一されたURL取得システムを使用（開発URL除去機能付き）
    const webAppUrl = getWebAppUrlCached();
    let deployDomain = ''; // 個人アカウント/グローバルアクセスの場合、デフォルトで空

    if (webAppUrl) {
      // Google WorkspaceアカウントのドメインをURLから抽出
      const domainMatch =
        webAppUrl.match(/\/a\/([a-zA-Z0-9\-.]+)\/macros\//) ||
        webAppUrl.match(/\/a\/macros\/([a-zA-Z0-9\-.]+)\//);
      if (domainMatch && domainMatch[1]) {
        deployDomain = domainMatch[1];
      }
      // ドメインが抽出されなかった場合、deployDomainは空のままとなり、個人アカウント/グローバルアクセスを示す
    }

    // 現在のユーザーのドメインと抽出された/デフォルトのデプロイドメインを比較
    // deployDomainが空の場合、特定のドメインが強制されていないため、一致とみなす（グローバルアクセス）
    const isDomainMatch = (currentDomain === deployDomain) || (deployDomain === '');

    debugLog('Domain info:', {
      currentDomain: currentDomain,
      deployDomain: deployDomain,
      isDomainMatch: isDomainMatch,
      webAppUrl: webAppUrl
    });

    return {
      currentDomain: currentDomain,
      deployDomain: deployDomain,
      isDomainMatch: isDomainMatch,
      webAppUrl: webAppUrl
    };
  } catch (e) {
    logError(e, 'getDeployUserDomainInfo', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.SYSTEM);
    return {
      currentDomain: '不明',
      deployDomain: '不明',
      isDomainMatch: false,
      error: e.message
    };
  }
}
// PerformanceOptimizer.gsでglobalProfilerが既に定義されているため、
// フォールバックの宣言は不要

/**
 * ユーザーがデータベースに登録済みかを確認するヘルパー関数
 * @param {string} email 確認するユーザーのメールアドレス
 * @param {string} spreadsheetId データベースのスプレッドシートID
 * @return {boolean} 登録されていればtrue
 */

/**
 * システムの初期セットアップが完了しているかを確認するヘルパー関数
 * @returns {boolean} セットアップが完了していればtrue
 */
function isSystemSetup() {
  const props = PropertiesService.getScriptProperties();
  const dbSpreadsheetId = props.getProperty('DATABASE_SPREADSHEET_ID');
  const adminEmail = props.getProperty('ADMIN_EMAIL');
  const serviceAccountCreds = props.getProperty('SERVICE_ACCOUNT_CREDS');
  return !!dbSpreadsheetId && !!adminEmail && !!serviceAccountCreds;
}
/**
 * Get Google Client ID for fallback authentication
 * @return {Object} Object containing client ID
 */
function getGoogleClientId() {
  try {
    debugLog('Getting GOOGLE_CLIENT_ID from script properties...');
    const properties = PropertiesService.getScriptProperties();
    const clientId = properties.getProperty('GOOGLE_CLIENT_ID');

    debugLog('GOOGLE_CLIENT_ID retrieved:', clientId ? 'Found' : 'Not found');

    if (!clientId) {
      warnLog('GOOGLE_CLIENT_ID not found in script properties');

      // Try to get all properties to see what's available
      const allProperties = properties.getProperties();
      debugLog('Available properties:', Object.keys(allProperties));

      return {
        clientId: '',
        error: 'GOOGLE_CLIENT_ID not found in script properties',
        setupInstructions: 'Please set GOOGLE_CLIENT_ID in Google Apps Script project settings under Properties > Script Properties'
      };
    }

    return { status: 'success', message: 'Google Client IDを取得しました', data: { clientId: clientId } };
  } catch (error) {
    logError(error, 'getGoogleClientId', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.SYSTEM);
    return { status: 'error', message: 'Google Client IDの取得に失敗しました: ' + error.toString(), data: { clientId: '' } };
  }
}

/**
 * システム設定の詳細チェック
 * @return {Object} システム設定の詳細情報
 */
function checkSystemConfiguration() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const allProperties = properties.getProperties();

    const requiredProperties = [
      'GOOGLE_CLIENT_ID',
      'DATABASE_SPREADSHEET_ID',
      'ADMIN_EMAIL',
      'SERVICE_ACCOUNT_CREDS'
    ];

    const configStatus = {};
    const missingProperties = [];

    requiredProperties.forEach(function(prop) {
      const value = allProperties[prop];
      configStatus[prop] = {
        exists: !!value,
        hasValue: !!(value && value.trim()),
        length: value ? value.length : 0
      };

      if (!value || !value.trim()) {
        missingProperties.push(prop);
      }
    });

    return {
      isFullyConfigured: missingProperties.length === 0,
      configStatus: configStatus,
      missingProperties: missingProperties,
      availableProperties: Object.keys(allProperties),
      setupComplete: isSystemSetup()
    };
  } catch (error) {
    logError(error, 'checkSystemConfiguration', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.SYSTEM);
    return {
      isFullyConfigured: false,
      error: error.toString()
    };
  }
}

/**
 * Retrieves the administrator domain for the login page with domain match status.
 * @returns {{adminDomain: string, isDomainMatch: boolean}|{error: string}} Domain info or error message.
 */
function getSystemDomainInfo() {
  try {
    const props = PropertiesService.getScriptProperties();
    const adminEmail = props.getProperty(SCRIPT_PROPS_KEYS.ADMIN_EMAIL);
    if (!adminEmail) {
      throw new Error('システム管理者が設定されていません。');
    }

    const adminDomain = adminEmail.split('@')[1];

    // 現在のユーザーのドメイン一致状況を取得
    const domainInfo = getDeployUserDomainInfo();
    const isDomainMatch = domainInfo.isDomainMatch !== undefined ? domainInfo.isDomainMatch : false;

    return {
      adminDomain: adminDomain,
      isDomainMatch: isDomainMatch,
      currentDomain: domainInfo.currentDomain || '不明',
      deployDomain: domainInfo.deployDomain || adminDomain
    };
  } catch (e) {
    logError(e, 'getSystemDomainInfo', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.SYSTEM);
    return { error: e.message };
  }
}

/**
 * Webアプリケーションのメインエントリーポイント
 * @param {Object} e - URLパラメータを含むイベントオブジェクト
 * @returns {HtmlOutput} 表示するHTMLコンテンツ
 */
function doGet(e) {
  try {
    // Initialize request processing
    const initResult = initializeRequestProcessing();
    if (initResult) return initResult;

    // Parse and validate request parameters
    const params = parseRequestParams(e);
    
    // Validate user authentication
    const authResult = validateUserAuthentication();
    if (authResult) return authResult;

    // Handle app setup page requests
    if (params.setupParam === 'true') {
      return showAppSetupPage(params.userId);
    }

    // Route request based on mode
    return routeRequestByMode(params);

  } catch (error) {
    logError(error, 'doGet', ERROR_SEVERITY.CRITICAL, ERROR_CATEGORIES.SYSTEM);
    return showErrorPage('致命的なエラー', 'アプリケーションの処理中に予期せぬエラーが発生しました。', error);
  }
}

/**
 * Initialize request processing with system checks
 * @returns {HtmlOutput|null} Early return result or null to continue
 */
function initializeRequestProcessing() {
  // Clear execution-level cache for new request
  clearAllExecutionCache();

  // Check system setup (highest priority)
  if (!isSystemSetup()) {
    return showSetupPage();
  }

  // Check application access permissions
  const accessCheck = checkApplicationAccess();
  if (!accessCheck.hasAccess) {
    infoLog('アプリケーションアクセス拒否:', accessCheck.accessReason);
    return showAccessRestrictedPage(accessCheck);
  }

  return null; // Continue processing
}

/**
 * Validate user authentication
 * @returns {HtmlOutput|null} Early return result or null to continue
 */
function validateUserAuthentication() {
  const userEmail = getCurrentUserEmail();
  if (!userEmail) {
    return showLoginPage();
  }
  return null; // Continue processing
}

/**
 * Route request based on mode parameter
 * @param {Object} params - Request parameters
 * @returns {HtmlOutput} Appropriate page response
 */
function routeRequestByMode(params) {
  // Handle no mode parameter
  if (!params || !params.mode) {
    return handleDefaultRoute();
  }

  // Route based on mode
  switch (params.mode) {
    case 'login':
      return handleLoginMode();
    case 'appSetup':
      return handleAppSetupMode();
    case 'admin':
      return handleAdminMode(params);
    case 'view':
      return handleViewMode(params);
    default:
      return handleUnknownMode(params);
  }
}

/**
 * Handle default route (no mode parameter)
 * @returns {HtmlOutput} Appropriate page response
 */
function handleDefaultRoute() {
  debugLog('No mode parameter, checking previous admin session');

  const activeUserEmail = Session.getActiveUser().getEmail();
  if (!activeUserEmail) {
    return showLoginPage();
  }

  const userProperties = PropertiesService.getUserProperties();
  const lastAdminUserId = userProperties.getProperty('lastAdminUserId');

  if (lastAdminUserId && verifyAdminAccess(lastAdminUserId)) {
    debugLog('Found previous admin session, redirecting to admin panel:', lastAdminUserId);
    const userInfo = findUserById(lastAdminUserId);
    return renderAdminPanel(userInfo, 'admin');
  }

  // Clear invalid admin session
  if (lastAdminUserId) {
    userProperties.deleteProperty('lastAdminUserId');
  }

  debugLog('No previous admin session, showing login page');
  return showLoginPage();
}

/**
 * Handle login mode
 * @returns {HtmlOutput} Login page
 */
function handleLoginMode() {
  debugLog('Login mode requested, showing login page');
  return showLoginPage();
}

/**
 * Handle app setup mode
 * @returns {HtmlOutput} App setup page or error page
 */
function handleAppSetupMode() {
  debugLog('AppSetup mode requested');

  const userProperties = PropertiesService.getUserProperties();
  const lastAdminUserId = userProperties.getProperty('lastAdminUserId');

  if (!lastAdminUserId) {
    debugLog('No admin session found, redirecting to login');
    return showErrorPage('認証が必要です', 'アプリ設定にアクセスするにはログインが必要です。');
  }

  if (!verifyAdminAccess(lastAdminUserId)) {
    warnLog('Admin access denied for userId:', lastAdminUserId);
    return showErrorPage('アクセス拒否', 'アプリ設定にアクセスする権限がありません。');
  }

  debugLog('Showing app setup page for userId:', lastAdminUserId);
  return showAppSetupPage(lastAdminUserId);
}

/**
 * Handle admin mode
 * @param {Object} params - Request parameters
 * @returns {HtmlOutput} Admin panel or error page
 */
function handleAdminMode(params) {
  if (!params.userId) {
    return showErrorPage('不正なリクエスト', 'ユーザーIDが指定されていません。');
  }

  if (!verifyAdminAccess(params.userId)) {
    return showErrorPage('アクセス拒否', 'この管理パネルにアクセスする権限がありません。');
  }

  // Save admin session state
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('lastAdminUserId', params.userId);
  debugLog('Saved admin session state:', params.userId);

  const userInfo = findUserById(params.userId);
  return renderAdminPanel(userInfo, 'admin');
}

/**
 * Handle view mode
 * @param {Object} params - Request parameters
 * @returns {HtmlOutput} Answer board or unpublished page
 */
function handleViewMode(params) {
  if (!params.userId) {
    return showErrorPage('不正なリクエスト', 'ユーザーIDが指定されていません。');
  }

  // Get user info with cache bypass for accurate publication status
  const userInfo = findUserById(params.userId, {
    useExecutionCache: false,
    forceRefresh: true
  });

  if (!userInfo) {
    return showErrorPage('エラー', '指定されたユーザーが見つかりません。');
  }

  return processViewRequest(userInfo, params);
}

/**
 * Process view request with publication status checks
 * @param {Object} userInfo - User information
 * @param {Object} params - Request parameters
 * @returns {HtmlOutput} Answer board or unpublished page
 */
function processViewRequest(userInfo, params) {
  // Parse config safely
  let config = {};
  try {
    config = getConfigJSON(userInfo);
  } catch (e) {
    warnLog('Config JSON parse error during publication check:', e.message);
  }

  // Check for auto-stop and handle accordingly
  const wasAutoStopped = checkAndHandleAutoStop(config, userInfo);
  if (wasAutoStopped) {
    infoLog('🔄 自動停止が実行されました - 非公開ページに誘導します');
  }

  // Check if currently published
  const isCurrentlyPublished = !!(config.appPublished === true &&
    config.publishedSpreadsheetId &&
    config.publishedSheetName &&
    typeof config.publishedSheetName === 'string' &&
    config.publishedSheetName.trim() !== '');

  debugLog('🔍 Publication status check:', {
    appPublished: config.appPublished,
    hasSpreadsheetId: !!config.publishedSpreadsheetId,
    hasSheetName: !!config.publishedSheetName,
    isCurrentlyPublished: isCurrentlyPublished
  });

  // Redirect to unpublished page if not published
  if (!isCurrentlyPublished) {
    infoLog('🚫 Board is unpublished, redirecting to Unpublished page immediately');
    debugLog('🔍 UserInfo for unpublished page:', {
      userId: userInfo.userId,
      adminEmail: userInfo.adminEmail,
      spreadsheetId: userInfo.spreadsheetId
    });

    try {
      return renderUnpublishedPage(userInfo, params);
    } catch (unpublishedError) {
      logError(unpublishedError, 'renderUnpublishedPage', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.SYSTEM);
      return renderMinimalUnpublishedPage(userInfo);
    }
  }

  return renderAnswerBoard(userInfo, params);
}

/**
 * Handle unknown mode
 * @param {Object} params - Request parameters
 * @returns {HtmlOutput} Appropriate page response
 */
function handleUnknownMode(params) {
  warnLog('Unknown mode received:', params.mode);
  debugLog('Available modes: login, appSetup, admin, view');

  // If valid userId with admin access, redirect to admin panel
  if (params.userId && verifyAdminAccess(params.userId)) {
    debugLog('Redirecting unknown mode to admin panel for valid user:', params.userId);
    const userInfo = findUserById(params.userId);
    return renderAdminPanel(userInfo, 'admin');
  }

  // Otherwise redirect to login
  debugLog('Redirecting unknown mode to login page');
  return showLoginPage();
}

/**
 * 管理者ページのルートを処理
 * @param {Object} userInfo - ユーザー情報
 * @param {Object} params - リクエストパラメータ
 * @param {string} userEmail - 現在のユーザーのメールアドレス
 * @returns {HtmlOutput}
 */
function handleAdminRoute(userInfo, params, userEmail) {
  // この関数が呼ばれる時点でuserInfoはnullではないことが保証されている

  // セキュリティチェック: アクセスしようとしているuserIdが自分のものでなければ、自分の管理画面にリダイレクト
  if (params.userId && params.userId !== userInfo.userId) {
    warnLog(`不正アクセス試行: ${userEmail} が userId ${params.userId} にアクセスしようとしました。`);
    const correctUrl = buildUserAdminUrl(userInfo.userId);
    return redirectToUrl(correctUrl);
  }

  // 強化されたセキュリティ検証: 指定されたIDの登録メールアドレスと現在ログイン中のGoogleアカウントが一致するかを検証
  if (params.userId) {
    const isVerified = verifyAdminAccess(params.userId);
    if (!isVerified) {
      warnLog(`セキュリティ検証失敗: userId ${params.userId} への不正アクセス試行をブロックしました。`);
      const correctUrl = buildUserAdminUrl(userInfo.userId);
      return redirectToUrl(correctUrl);
    }
    debugLog(`✅ セキュリティ検証成功: userId ${params.userId} への正当なアクセスを確認しました。`);
  }

  return renderAdminPanel(userInfo, params.mode);
}
/**
 * 統合されたユーザー情報取得関数 (Phase 2: 統合完了)
 * キャッシュ戦略とセキュリティチェックを統合
 * @param {string|Object} identifier - メールアドレス、ユーザーID、または設定オブジェクト
 * @param {string} [type] - 'email' | 'userId' | null (auto-detect)
 * @param {Object} [options] - キャッシュオプション
 * @returns {Object|null} ユーザー情報オブジェクト、見つからない場合はnull
 */
function getOrFetchUserInfo(identifier, type = null, options = {}) {
  // オプションのデフォルト値
  const opts = {
    ttl: options.ttl || 300, // 5分キャッシュ
    enableSecurityCheck: options.enableSecurityCheck !== false, // デフォルトで有効
    currentUserEmail: options.currentUserEmail || null,
    useExecutionCache: options.useExecutionCache || false,
    ...options
  };

  // 引数の正規化
  let email = null;
  let userId = null;

  if (typeof identifier === 'object' && identifier !== null) {
    // オブジェクト形式の場合（後方互換性）
    email = identifier.email;
    userId = identifier.userId;
  } else if (typeof identifier === 'string') {
    // 文字列の場合、typeに基づいて判定
    if (type === 'email' || (!type && identifier.includes('@'))) {
      email = identifier;
    } else {
      userId = identifier;
    }
  }

  // キャッシュキーの生成
  const cacheKey = `unified_user_info_${userId || email}`;

  // 実行レベルキャッシュの確認（オプション）
  if (opts.useExecutionCache && _executionUserInfoCache &&
      _executionUserInfoCache.userId === userId) {
    return _executionUserInfoCache.userInfo;
  }

  // 統合キャッシュマネージャーを使用（キャッシュ miss 時は自動でデータベースから取得）
  let userInfo = null;

  try {
    userInfo = cacheManager.get(cacheKey, () => {
      // キャッシュに存在しない場合はデータベースから取得する
      // 通常フローのためエラーレベルでは記録しない
      debugLog('cache miss - fetching from database');

      const props = PropertiesService.getScriptProperties();
      if (!props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID)) {
        logError('DATABASE_SPREADSHEET_ID not set', 'getUnifiedCache', ERROR_SEVERITY.CRITICAL, ERROR_CATEGORIES.SYSTEM);
        return null;
      }

      let dbUserInfo = null;
      if (userId) {
        dbUserInfo = findUserById(userId);
        // セキュリティチェック: 取得した情報のemailが現在のユーザーと一致するか確認
        if (opts.enableSecurityCheck && dbUserInfo && opts.currentUserEmail &&
            dbUserInfo.adminEmail !== opts.currentUserEmail) {
          warnLog('セキュリティチェック失敗: 他人の情報へのアクセス試行');
          return null;
        }
      } else if (email) {
        dbUserInfo = findUserByEmail(email);
      }

      return dbUserInfo;
    }, {
      ttl: opts.ttl || 300,
      enableMemoization: opts.enableMemoization || false
    });

    // 実行レベルキャッシュにも保存（オプション）
    if (userInfo && opts.useExecutionCache && (userId || userInfo.userId)) {
      _executionUserInfoCache = { userId: userId || userInfo.userId, userInfo };
      debugLog('✅ 実行レベルキャッシュに保存:', userId || userInfo.userId);
    }

  } catch (cacheError) {
    logError(cacheError, 'getUnifiedCache', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.CACHE);
    // フォールバック: 直接データベースから取得
    if (userId) {
      userInfo = findUserById(userId);
    } else if (email) {
      userInfo = findUserByEmail(email);
    }
  }

  return userInfo;
}
/**
 * ログインページを表示
 * @returns {HtmlOutput}
 */
function showLoginPage() {
  const template = HtmlService.createTemplateFromFile('LoginPage');
  const htmlOutput = template.evaluate()
    .setTitle('StudyQuest - ログイン');

  // XFrameOptionsMode を安全に設定
  try {
    if (HtmlService && HtmlService.XFrameOptionsMode && HtmlService.XFrameOptionsMode.ALLOWALL) {
      htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  } catch (e) {
    warnLog('XFrameOptionsMode設定エラー:', e.message);
  }

  return htmlOutput;
}

/**
 * 初回セットアップページを表示
 * @returns {HtmlOutput}
 */
function showSetupPage() {
  const template = HtmlService.createTemplateFromFile('SetupPage');
  const htmlOutput = template.evaluate()
    .setTitle('StudyQuest - 初回セットアップ');

  // XFrameOptionsMode を安全に設定
  try {
    if (HtmlService && HtmlService.XFrameOptionsMode && HtmlService.XFrameOptionsMode.DENY) {
      htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DENY);
    }
  } catch (e) {
    warnLog('XFrameOptionsMode設定エラー:', e.message);
  }

  return htmlOutput;
}

/**
 * アプリ設定ページを表示
 * @returns {HtmlOutput}
 */
function showAppSetupPage(userId) {
    // システム管理者権限チェック
    try {
      debugLog('showAppSetupPage: Checking deploy user permissions...');
      const currentUserEmail = getCurrentUserEmail();
      debugLog('showAppSetupPage: Current user email:', currentUserEmail);
      const deployUserCheckResult = isDeployUser();
      debugLog('showAppSetupPage: isDeployUser() result:', deployUserCheckResult);

      if (!deployUserCheckResult) {
        warnLog('Unauthorized access attempt to app setup page:', currentUserEmail);
        return showErrorPage('アクセス権限がありません', 'この機能にアクセスする権限がありません。システム管理者にお問い合わせください。');
      }
    } catch (error) {
      logError(error, 'checkDeployUserPermissions', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.AUTHORIZATION);
      return showErrorPage('認証エラー', '権限確認中にエラーが発生しました。');
    }

    const appSetupTemplate = HtmlService.createTemplateFromFile('AppSetupPage');
    appSetupTemplate.userId = userId;
    const htmlOutput = appSetupTemplate.evaluate()
      .setTitle('アプリ設定 - StudyQuest');

    // XFrameOptionsMode を安全に設定
    try {
      if (HtmlService && HtmlService.XFrameOptionsMode && HtmlService.XFrameOptionsMode.DENY) {
        htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DENY);
      }
    } catch (e) {
      warnLog('XFrameOptionsMode設定エラー:', e.message);
    }

    return htmlOutput;
}

/**
 * 最後にアクセスした管理ユーザーIDを取得
 * リダイレクト時の戻り先決定に使用
 * @returns {string|null} 管理ユーザーID、存在しない場合はnull
 */
function getLastAdminUserId() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const lastAdminUserId = userProperties.getProperty('lastAdminUserId');

    // ユーザーIDが存在し、かつ有効な管理者権限を持つかチェック
    if (lastAdminUserId && verifyAdminAccess(lastAdminUserId)) {
      debugLog('Found valid admin user ID:', lastAdminUserId);
      return lastAdminUserId;
    } else {
      debugLog('No valid admin user ID found');
      return null;
    }
  } catch (error) {
    logError(error, 'getLastAdminUserId', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.DATABASE);
    return null;
  }
}

/**
 * アプリ設定ページのURLを取得（フロントエンドから呼び出し用）
 * @returns {string} アプリ設定ページのURL
 */
function getAppSetupUrl() {
  try {
    // システム管理者権限チェック
    debugLog('getAppSetupUrl: Checking deploy user permissions...');
    const currentUserEmail = Session.getActiveUser().getEmail();
    debugLog('getAppSetupUrl: Current user email:', currentUserEmail);
    const deployUserCheckResult = isDeployUser();
    debugLog('getAppSetupUrl: isDeployUser() result:', deployUserCheckResult);

    if (!deployUserCheckResult) {
      warnLog('Unauthorized access attempt to get app setup URL:', currentUserEmail);
      throw new Error('アクセス権限がありません');
    }

    // WebアプリのベースURLを取得
    const baseUrl = ScriptApp.getService().getUrl();
    if (!baseUrl) {
      throw new Error('WebアプリのURLを取得できませんでした');
    }

    // アプリ設定ページのURLを生成
    const appSetupUrl = baseUrl + '?mode=appSetup';
    debugLog('getAppSetupUrl: Generated URL:', appSetupUrl);

    return appSetupUrl;
  } catch (error) {
    logError(error, 'getAppSetupUrl', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.SYSTEM);
    throw new Error('アプリ設定URLの取得に失敗しました: ' + error.message);
  }
}

/**
 * エラーページを表示
 * @param {string} title - エラータイトル
 * @param {string} message - エラーメッセージ
 * @param {Error} [error] - (オプション) エラーオブジェクト
 * @returns {HtmlOutput}
 */
function showErrorPage(title, message, error) {
  const template = HtmlService.createTemplateFromFile('ErrorBoundary');
  template.title = title;
  template.message = message;
  template.mode = 'admin'; // エラーテンプレートが依存するmode変数にデフォルト値を提供
  if (DEBUG && error) {
    template.debugInfo = error.stack;
  } else {
    template.debugInfo = '';
  }
  const htmlOutput = template.evaluate()
    .setTitle(`エラー - ${title}`);

  // XFrameOptionsMode を安全に設定
  try {
    if (HtmlService && HtmlService.XFrameOptionsMode && HtmlService.XFrameOptionsMode.DENY) {
      htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DENY);
    }
  } catch (e) {
    warnLog('XFrameOptionsMode設定エラー:', e.message);
  }

  return htmlOutput;
}

/**
 * アクセス制限ページを表示
 * @param {object} accessCheck - アクセスチェック結果
 * @returns {HtmlOutput}
 */
function showAccessRestrictedPage(accessCheck) {
  try {
    const template = HtmlService.createTemplateFromFile('AccessRestricted');
    template.include = include;
    template.accessCheck = accessCheck;
    template.isSystemAdmin = accessCheck.isSystemAdmin || false;
    template.userEmail = accessCheck.userEmail || '';
    template.timestamp = new Date().toISOString();

    const htmlOutput = template.evaluate()
      .setTitle('アクセス制限 - StudyQuest');

    // XFrameOptionsMode を安全に設定
    try {
      if (HtmlService && HtmlService.XFrameOptionsMode && HtmlService.XFrameOptionsMode.DENY) {
        htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DENY);
      }
    } catch (e) {
      warnLog('XFrameOptionsMode設定エラー:', e.message);
    }

    return htmlOutput;
  } catch (error) {
    logError(error, 'showAccessRestrictedPage', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.SYSTEM);
    // フォールバック: シンプルなHTMLを返す
    return HtmlService.createHtmlOutput(`
      <html>
        <head><title>アクセス制限</title></head>
        <body style="font-family: Arial, sans-serif; text-align: center; padding: 50px;">
          <h1>アクセスが制限されています</h1>
          <p>このアプリケーションは現在利用できません。</p>
          <p>管理者にお問い合わせください。</p>
        </body>
      </html>
    `).setTitle('アクセス制限');
  }
}

/**
 * ユーザー専用の一意の管理パネルURLを構築
 * @param {string} userId ユーザーID
 * @return {string}
 */
function buildUserAdminUrl(userId) {
  const baseUrl = getWebAppUrl();
  return `${baseUrl}?mode=admin&userId=${encodeURIComponent(userId)}`;
}

/**
 * 標準化されたページURL生成のヘルパー関数群
 */
const URLBuilder = {
  /**
   * ログインページのURLを生成
   * @returns {string} ログインページURL
   */
  login: function() {
    const baseUrl = getWebAppUrl();
    return `${baseUrl}?mode=login`;
  },

  /**
   * 管理パネルのURLを生成
   * @param {string} userId - ユーザーID
   * @returns {string} 管理パネルURL
   */
  admin: function(userId) {
    const baseUrl = getWebAppUrl();
    return `${baseUrl}?mode=admin&userId=${encodeURIComponent(userId)}`;
  },

  /**
   * アプリ設定ページのURLを生成
   * @returns {string} アプリ設定ページURL
   */
  appSetup: function() {
    const baseUrl = getWebAppUrl();
    return `${baseUrl}?mode=appSetup`;
  },

  /**
   * 回答ボードのURLを生成
   * @param {string} userId - ユーザーID
   * @returns {string} 回答ボードURL
   */
  view: function(userId) {
    const baseUrl = getWebAppUrl();
    return `${baseUrl}?mode=view&userId=${encodeURIComponent(userId)}`;
  },

  /**
   * パラメータ付きURLを安全に生成
   * @param {string} mode - モード ('login', 'admin', 'view', 'appSetup')
   * @param {Object} params - 追加パラメータ
   * @returns {string} 生成されたURL
   */
  build: function(mode, params = {}) {
    const baseUrl = getWebAppUrl();
    const url = new URL(baseUrl);
    url.searchParams.set('mode', mode);

    Object.keys(params).forEach(key => {
      if (params[key] !== null && params[key] !== undefined) {
        url.searchParams.set(key, params[key]);
      }
    });

    return url.toString();
  }
};

/**
 * 指定されたURLへリダイレクトするサーバーサイド関数
 * @param {string} url - リダイレクト先のURL
 * @returns {HtmlOutput} リダイレクトを実行するHTML出力
 */
function redirectToUrl(url) {
  // XSS攻撃を防ぐため、URLをサニタイズ
  const sanitizedUrl = sanitizeRedirectUrl(url);
  return HtmlService.createHtmlOutput().setContent(`<script>window.top.location.href = '${sanitizedUrl}';</script>`);
}
/**
 * セキュアなリダイレクトHTMLを作成 (シンプル版)
 * @param {string} targetUrl リダイレクト先URL
 * @param {string} message 表示メッセージ
 * @return {HtmlOutput}
 */
function createSecureRedirect(targetUrl, message) {
  // URL検証とサニタイゼーション
  const sanitizedUrl = sanitizeRedirectUrl(targetUrl);

  debugLog('createSecureRedirect - Original URL:', targetUrl);
  debugLog('createSecureRedirect - Sanitized URL:', sanitizedUrl);

  // ユーザーアクティベーション必須のHTMLアンカー方式（サンドボックス制限準拠）
  const userActionRedirectHtml = `
    <!DOCTYPE html>
    <html>
    <head>
      <title>${message || 'アクセス確認'}</title>
      <meta charset="utf-8">
      <meta name="viewport" content="width=device-width, initial-scale=1">
      <style>
        body {
          font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
          background: linear-gradient(135deg, #1e3a8a 0%, #7c3aed 100%);
          min-height: 100vh;
          margin: 0;
          display: flex;
          align-items: center;
          justify-content: center;
          padding: 20px;
        }
        .container {
          background: rgba(31, 41, 55, 0.95);
          border-radius: 16px;
          padding: 40px;
          max-width: 500px;
          width: 100%;
          text-align: center;
          box-shadow: 0 25px 50px rgba(0, 0, 0, 0.5);
          border: 1px solid rgba(75, 85, 99, 0.3);
        }
        .icon { font-size: 64px; margin-bottom: 20px; }
        .title {
          color: #10b981;
          font-size: 24px;
          font-weight: bold;
          margin-bottom: 16px;
        }
        .subtitle {
          color: #d1d5db;
          margin-bottom: 32px;
          line-height: 1.5;
        }
        .main-button {
          display: inline-block;
          background: linear-gradient(135deg, #10b981 0%, #3b82f6 100%);
          color: white;
          font-weight: bold;
          padding: 16px 32px;
          border-radius: 12px;
          text-decoration: none;
          font-size: 18px;
          transition: all 0.3s ease;
          margin-bottom: 24px;
          box-shadow: 0 4px 15px rgba(16, 185, 129, 0.4);
        }
        .main-button:hover {
          transform: translateY(-2px);
          box-shadow: 0 8px 25px rgba(16, 185, 129, 0.6);
        }
        .url-info {
          background: rgba(17, 24, 39, 0.8);
          border-radius: 8px;
          padding: 16px;
          margin: 20px 0;
          border: 1px solid rgba(75, 85, 99, 0.5);
        }
        .url-text {
          color: #60a5fa;
          font-family: 'Courier New', monospace;
          font-size: 12px;
          word-break: break-all;
          line-height: 1.4;
        }
        .note {
          color: #9ca3af;
          font-size: 14px;
          margin-top: 20px;
          line-height: 1.4;
        }
      </style>
    </head>
    <body>
      <div class="container">
        <div class="icon">🔐</div>
        <h1 class="title">${message || 'アクセス確認'}</h1>
        <p class="subtitle">セキュリティのため、下のボタンをクリックして続行してください</p>

        <a href="${sanitizedUrl}" target="_top" class="main-button">
          🚀 続行する
        </a>

        <div class="url-info">
          <div class="url-text">${sanitizedUrl}</div>
        </div>

        <div class="note">
          ✓ このリンクは安全です<br>
          ✓ Google Apps Script公式のセキュリティガイドラインに準拠
        </div>
      </div>
    </body>
    </html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(userActionRedirectHtml);

  // XFrameOptionsMode を安全に設定
  try {
    if (HtmlService && HtmlService.XFrameOptionsMode && HtmlService.XFrameOptionsMode.ALLOWALL) {
      htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  } catch (e) {
    warnLog('XFrameOptionsMode設定エラー:', e.message);
  }

  return htmlOutput;
}

/**
 * リダイレクト用URLの検証とサニタイゼーション
 * @param {string} url 検証対象のURL
 * @return {string} サニタイズされたURL
 */
function sanitizeRedirectUrl(url) {
  if (!url) {
    return getWebAppUrlCached();
  }

  try {
    let cleanUrl = String(url).trim();

    // 複数レベルのクォート除去（JSON文字列化による多重クォートに対応）
    let previousUrl = '';
    while (cleanUrl !== previousUrl) {
      previousUrl = cleanUrl;

      // 先頭と末尾のクォートを除去
      if ((cleanUrl.startsWith('"') && cleanUrl.endsWith('"')) ||
          (cleanUrl.startsWith("'") && cleanUrl.endsWith("'"))) {
        cleanUrl = cleanUrl.slice(1, -1);
      }

      // エスケープされたクォートを除去
      cleanUrl = cleanUrl.replace(/\\"/g, '"').replace(/\\'/g, "'");

      // URL内に埋め込まれた別のURLを検出
      const embeddedUrlMatch = cleanUrl.match(/https?:\/\/[^\s<>"']+/);
      if (embeddedUrlMatch && embeddedUrlMatch[0] !== cleanUrl) {
        debugLog('Extracting embedded URL:', embeddedUrlMatch[0]);
        cleanUrl = embeddedUrlMatch[0];
      }
    }

    // 基本的なURL形式チェック
    if (!cleanUrl.match(/^https?:\/\/[^\s<>"']+$/)) {
      warnLog('Invalid URL format after sanitization:', cleanUrl);
      return getWebAppUrlCached();
    }

    // 開発モードURLのチェック（googleusercontent.comは有効なデプロイURLも含むため調整）
    if (cleanUrl.includes('userCodeAppPanel')) {
      warnLog('Development URL detected in redirect, using fallback:', cleanUrl);
      return getWebAppUrlCached();
    }

    // 最終的な URL 妥当性チェック（googleusercontent.comも有効URLとして認識）
    const isValidUrl = cleanUrl.includes('script.google.com') ||
                     cleanUrl.includes('googleusercontent.com') ||
                     cleanUrl.includes('localhost');

    if (!isValidUrl) {
      warnLog('Suspicious URL detected:', cleanUrl);
      return getWebAppUrlCached();
    }

    return cleanUrl;
  } catch (e) {
    logError(e, 'urlSanitization', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.SYSTEM);
    return getWebAppUrlCached();
  }
}

/**
 * 正しいWeb App URLを取得 (url.gsのgetWebAppUrlCachedを使用)
 * @return {string}
 */
function getWebAppUrl() {
  try {
    // url.gsの統一されたURL取得関数を使用
    return getWebAppUrlCached();
  } catch (error) {
    logError(error, 'getWebAppUrl', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.SYSTEM);
    // 緊急時のフォールバックURL
    return getFallbackUrl();
  }
}
/**
 * doGet のリクエストパラメータを解析
 * @param {Object} e Event object
 * @return {{mode:string,userId:string|null,setupParam:string|null,spreadsheetId:string|null,sheetName:string|null,isDirectPageAccess:boolean}}
 */
function parseRequestParams(e) {
  // 引数の安全性チェック
  if (!e || typeof e !== 'object') {
    debugLog('parseRequestParams: 無効なeventオブジェクト');
    return { mode: null, userId: null, setupParam: null, spreadsheetId: null, sheetName: null, isDirectPageAccess: false };
  }

  const p = e.parameter || {};
  const mode = p.mode || null;
  const userId = p.userId || null;
  const setupParam = p.setup || null;
  const spreadsheetId = p.spreadsheetId || null;
  const sheetName = p.sheetName || null;
  const isDirectPageAccess = !!(userId && mode === 'view');

  // デバッグログを追加
  debugLog('parseRequestParams - Received parameters:', JSON.stringify(p));
  debugLog('parseRequestParams - Parsed mode:', mode, 'setupParam:', setupParam);

  return { mode, userId, setupParam, spreadsheetId, sheetName, isDirectPageAccess };
}

/**
 * 管理パネルを表示
 * @param {Object} userInfo ユーザー情報
 * @param {string} mode 表示モード
 * @return {HtmlOutput} HTMLコンテンツ
 */
function renderAdminPanel(userInfo, mode) {
  // ガード節: userInfoが存在しない場合はエラーページを表示して処理を中断
  if (!userInfo) {
    logError('renderAdminPanelにuserInfoがnullで渡されました', 'renderAdminPanel', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.SYSTEM);
    return showErrorPage('エラー', 'ユーザー情報の読み込みに失敗したため、管理パネルを表示できません。');
  }

  const adminTemplate = HtmlService.createTemplateFromFile('AdminPanel');
  adminTemplate.include = include;
  adminTemplate.userInfo = userInfo;
  adminTemplate.userId = userInfo.userId;
  adminTemplate.mode = mode || 'admin'; // 安全のためのフォールバック
  adminTemplate.displayMode = 'named';
  adminTemplate.showAdminFeatures = true;
  const deployUserResult = isDeployUser();
  debugLog('renderAdminPanel - isDeployUser() result:', deployUserResult);
  debugLog('renderAdminPanel - current user email:', getCurrentUserEmail());
  adminTemplate.isDeployUser = deployUserResult;
  adminTemplate.DEBUG_MODE = shouldEnableDebugMode();


  const htmlOutput = adminTemplate.evaluate()
    .setTitle('StudyQuest - 管理パネル')
    .setSandboxMode(HtmlService.SandboxMode.NATIVE);

  // XFrameOptionsMode を安全に設定
  try {
    if (HtmlService && HtmlService.XFrameOptionsMode && HtmlService.XFrameOptionsMode.ALLOWALL) {
      htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  } catch (e) {
    warnLog('XFrameOptionsMode設定エラー:', e.message);
  }

  return htmlOutput;
}

/**
 * 回答ボードまたは未公開ページを表示
 * @param {Object} userInfo ユーザー情報
 * @param {Object} params リクエストパラメータ
 * @return {HtmlOutput} HTMLコンテンツ
 */
/**
 * 非公開ボード専用のレンダリング関数
 * ErrorBoundaryを回避して確実にUnpublished.htmlを表示
 * @param {Object} userInfo - ユーザー情報
 * @param {Object} params - URLパラメータ
 * @returns {HtmlOutput} Unpublished.htmlテンプレート
 */
function renderUnpublishedPage(userInfo, params) {
  try {
    debugLog('🚫 renderUnpublishedPage: Rendering unpublished page for userId:', userInfo.userId);

    const template = HtmlService.createTemplateFromFile('Unpublished');
    template.include = include;

    // 基本情報の設定（安全なデフォルト値付き）
    template.userId = userInfo.userId || '';
    template.spreadsheetId = userInfo.spreadsheetId || '';
    template.ownerName = userInfo.adminEmail || 'システム管理者';
    template.isOwner = true; // 非公開ページは所有者のみアクセス可能
    template.adminEmail = userInfo.adminEmail || '';
    template.cacheTimestamp = Date.now();

    // 安全な変数設定
    template.include = include;

    // URL生成（エラー耐性を持たせる）
    let appUrls;
    try {
      appUrls = generateUserUrls(userInfo.userId);
      if (!appUrls || appUrls.status === 'error') {
        throw new Error('URL生成に失敗しました');
      }
    } catch (urlError) {
      warnLog('URL生成エラー、フォールバック値を使用:', urlError);
      // フォールバック: 基本的なURL構造
      const baseUrl = getWebAppUrlCached() || getFallbackUrl();
      appUrls = {
        adminUrl: `${baseUrl}?mode=admin&userId=${encodeURIComponent(userInfo.userId)}`,
        viewUrl: `${baseUrl}?mode=view&userId=${encodeURIComponent(userInfo.userId)}`,
        status: 'fallback'
      };
    }

    template.adminPanelUrl = appUrls.adminUrl || '';
    template.boardUrl = appUrls.viewUrl || '';

    debugLog('✅ renderUnpublishedPage: Template setup completed');

    // キャッシュを無効化して確実なリダイレクトを保証
    const htmlOutput = template.evaluate()
      .setTitle('StudyQuest - 準備中')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .addMetaTag('cache-control', 'no-cache, no-store, must-revalidate')
      .addMetaTag('pragma', 'no-cache')
      .addMetaTag('expires', '0');

    try {
      if (HtmlService && HtmlService.XFrameOptionsMode &&
          HtmlService.XFrameOptionsMode.ALLOWALL) {
        htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      }
    } catch (e) {
      warnLog('XFrameOptionsMode設定エラー:', e.message);
    }

    return htmlOutput;

  } catch (error) {
    logError(error, 'renderUnpublishedPage', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.SYSTEM);
    // フォールバック: 基本的なエラーページ
    return showErrorPage('準備中', 'ボードの準備が完了していません。管理者にお問い合わせください。');
  }
}

/**
 * 最小限の非公開ページレンダリング（フォールバック用）
 * テンプレートエラーを回避してHTMLを直接生成
 * @param {Object} userInfo - ユーザー情報
 * @returns {HtmlOutput} 最小限のHTML
 */
function renderMinimalUnpublishedPage(userInfo) {
  try {
    debugLog('🚫 renderMinimalUnpublishedPage: Creating minimal unpublished page');

    const userId = userInfo.userId || '';
    const adminEmail = userInfo.adminEmail || '';

    // 直接HTMLを生成（テンプレートを使わない）
    const htmlContent = `
      <!DOCTYPE html>
      <html lang="ja">
      <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1">
          <title>StudyQuest - 準備中</title>
          <style>
              body { font-family: -apple-system, BlinkMacSystemFont, sans-serif; margin: 0; padding: 20px; background: #1a1a1a; color: white; text-align: center; }
              .container { max-width: 600px; margin: 50px auto; padding: 40px 20px; background: #2a2a2a; border-radius: 12px; }
              .status { font-size: 24px; margin-bottom: 20px; color: #fbbf24; }
              .message { font-size: 16px; margin-bottom: 30px; color: #9ca3af; }
              .admin-btn { display: inline-block; padding: 12px 24px; background: #3b82f6; color: white; text-decoration: none; border-radius: 8px; margin: 10px; }
              .admin-btn:hover { background: #2563eb; }
          </style>
      </head>
      <body>
          <div class="container">
              <div class="status">⏳ 回答ボード準備中</div>
              <div class="message">現在、回答ボードは非公開になっています</div>
              <a href="?mode=admin&userId=${encodeURIComponent(userId)}" class="admin-btn">管理パネルを開く</a>
              <div style="margin-top: 20px; font-size: 12px; color: #6b7280;">管理者: ${adminEmail}</div>
          </div>
      </body>
      </html>
    `;

    return HtmlService.createHtmlOutput(htmlContent)
      .setTitle('StudyQuest - 準備中')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .addMetaTag('cache-control', 'no-cache, no-store, must-revalidate');

  } catch (error) {
    logError(error, 'renderMinimalUnpublishedPage', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.SYSTEM);
    // 最終フォールバック
    return HtmlService.createHtmlOutput('<h1>準備中</h1><p>回答ボードは現在準備中です。</p>')
      .setTitle('準備中');
  }
}

function renderAnswerBoard(userInfo, params) {
  try {
    let config = {};
    try {
      config = getConfigJSON(userInfo);
    } catch (e) {
      warnLog('Invalid configJson:', e.message);
    }
  // publishedSheetNameの型安全性確保（'true'問題の修正）
  let safePublishedSheetName = '';
  if (config.publishedSheetName) {
    if (typeof config.publishedSheetName === 'string') {
      safePublishedSheetName = config.publishedSheetName;
    } else {
      logValidationError('publishedSheetName', config.publishedSheetName, 'string_type', `不正な型: ${typeof config.publishedSheetName}`);
      warnLog('🔧 main.gs: publishedSheetNameを空文字にリセットしました');
      safePublishedSheetName = '';
    }
  }

  // 強化されたパブリケーション状態検証（キャッシュバスティング対応）
  const isPublished = !!(config.appPublished && config.publishedSpreadsheetId && safePublishedSheetName);

  // リアルタイム検証: 非公開状態の場合は確実に検出
  const isCurrentlyPublished = isPublished &&
    config.appPublished === true &&
    config.publishedSpreadsheetId &&
    safePublishedSheetName;

  const sheetConfigKey = 'sheet_' + (safePublishedSheetName || params.sheetName);
  const sheetConfig = config[sheetConfigKey] || {};

  // この関数は公開ボード専用（非公開判定は呼び出し前に完了）
  debugLog('✅ renderAnswerBoard: Rendering published board for userId:', userInfo.userId);

  const template = HtmlService.createTemplateFromFile('Page');
  template.include = include;

  try {
      if (userInfo.spreadsheetId) {
        try { addServiceAccountToSpreadsheet(userInfo.spreadsheetId); } catch (err) { warnLog('アクセス権設定警告:', err.message); }
      }
      template.userId = userInfo.userId;
      template.spreadsheetId = userInfo.spreadsheetId;
      template.ownerName = userInfo.adminEmail;
      template.sheetName = escapeJavaScript(safePublishedSheetName || params.sheetName);
      template.DEBUG_MODE = shouldEnableDebugMode();
      // setupStatus未完了時の安全なopinionHeader取得
      const setupStatus = config.setupStatus || 'pending';
      let rawOpinionHeader;

      if (setupStatus === 'pending') {
        rawOpinionHeader = 'セットアップ中...';
      } else {
        rawOpinionHeader = sheetConfig.opinionHeader || safePublishedSheetName || 'お題';
      }
      template.opinionHeader = escapeJavaScript(rawOpinionHeader);
      template.cacheTimestamp = Date.now();
      template.displayMode = config.displayMode || 'anonymous';
      template.showCounts = config.showCounts !== undefined ? config.showCounts : false;
      template.showScoreSort = template.showCounts;
      const currentUserEmail = getCurrentUserEmail();
      const isOwner = currentUserEmail === userInfo.adminEmail;
      template.showAdminFeatures = isOwner;
      template.isAdminUser = isOwner;
    } catch (e) {
      template.opinionHeader = escapeJavaScript('お題の読込エラー');
      template.cacheTimestamp = Date.now();
      template.userId = userInfo.userId;
      template.spreadsheetId = userInfo.spreadsheetId;
      template.ownerName = userInfo.adminEmail;
      template.sheetName = escapeJavaScript(params.sheetName);
      template.displayMode = 'anonymous';
      template.showCounts = false;
      template.showScoreSort = false;
      const currentUserEmail = getCurrentUserEmail();
      const isOwner = currentUserEmail === userInfo.adminEmail;
      template.showAdminFeatures = isOwner;
      template.isAdminUser = isOwner;
    }

  // 公開ボード: 通常のキャッシュ設定
  return template.evaluate()
    .setTitle('StudyQuest -みんなの回答ボード-')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');

  } catch (error) {
    logError(
      error,
      'renderAnswerBoard',
      ERROR_SEVERITY.HIGH,
      ERROR_CATEGORIES.SYSTEM,
      {
        userId: userInfo.userId,
        spreadsheetId: userInfo.spreadsheetId,
        sheetName: safePublishedSheetName || params.sheetName,
      }
    );
    // フォールバック: 基本的なエラーページ
    return showErrorPage('エラー', 'ボードの表示でエラーが発生しました。管理者にお問い合わせください。');
  }
}

/**
 * クライアントサイドからのパブリケーション状態チェック
 * キャッシュバスティング対応のため、リアルタイムで状態を確認
 * @param {string} userId - ユーザーID
 * @returns {Object} パブリケーション状態情報
 */
function checkCurrentPublicationStatus(userId) {
  try {
    debugLog('🔍 checkCurrentPublicationStatus called for userId:', userId);

    if (!userId) {
      warnLog('userId is required for publication status check');
      return { error: 'userId is required', isPublished: false };
    }

    // ユーザー情報を強制的に最新状態で取得（キャッシュバイパス）
    const userInfo = findUserById(userId, {
      useExecutionCache: false,
      forceRefresh: true
    });

    if (!userInfo) {
      warnLog('User not found for publication status check:', userId);
      return { error: 'User not found', isPublished: false };
    }

    // 設定情報を解析
    let config = {};
    try {
      config = getConfigJSON(userInfo);
    } catch (e) {
      warnLog('Config JSON parse error during publication status check:', e.message);
      return { error: 'Config parse error', isPublished: false };
    }

    // 現在のパブリケーション状態を厳密にチェック
    const isCurrentlyPublished = !!(
      config.appPublished === true &&
      config.publishedSpreadsheetId &&
      config.publishedSheetName &&
      typeof config.publishedSheetName === 'string' &&
      config.publishedSheetName.trim() !== ''
    );

    debugLog('📊 Publication status check result:', {
      userId: userId,
      appPublished: config.appPublished,
      hasSpreadsheetId: !!config.publishedSpreadsheetId,
      hasSheetName: !!config.publishedSheetName,
      isCurrentlyPublished: isCurrentlyPublished,
      timestamp: new Date().toISOString()
    });

    return {
      isPublished: isCurrentlyPublished,
      publishedSheetName: config.publishedSheetName || null,
      publishedSpreadsheetId: config.publishedSpreadsheetId || null,
      lastChecked: new Date().toISOString()
    };

  } catch (error) {
    logError(error, 'checkCurrentPublicationStatus', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.SYSTEM);
    return {
      error: error.message,
      isPublished: false,
      lastChecked: new Date().toISOString()
    };
  }
}
/**
 * JavaScript エスケープのテスト関数
 */

/**
 * パフォーマンス監視エンドポイント（簡易版）
 */

/**
 * escapeJavaScript関数のテスト
 */

/**
 * Base64エンコード/デコードのテスト
 */
