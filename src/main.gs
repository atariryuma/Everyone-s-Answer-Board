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
    return HtmlService.createHtmlOutputFromFile(path).getContent();
  } catch (error) {
    logError(
      error,
      'includeFile',
      UNIFIED_CONSTANTS.ERROR.SEVERITY.HIGH,
      UNIFIED_CONSTANTS.ERROR.CATEGORIES.SYSTEM,
      { filePath: path }
    );
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
  ADMIN_EMAIL: 'ADMIN_EMAIL',
};

// constants.gsで定義済みのERROR_SEVERITY, ERROR_CATEGORIESを使用

// DB_SHEET_CONFIG is defined in constants.gs

// LOG_SHEET_CONFIG is defined in constants.gs

// 履歴管理の定数
// MAX_HISTORY_ITEMS is defined in constants.gs

// 実行中のユーザー情報キャッシュ（パフォーマンス最適化用）
let _executionUserInfoCache = null;

// COLUMN_HEADERS is defined in constants.gs

// 表示モード定数
// DISPLAY_MODES is defined in constants.gs

// リアクション関連定数
// REACTION_KEYS is defined in constants.gs

// スコア計算設定
const SCORING_CONFIG = {
  LIKE_MULTIPLIER_FACTOR: 0.1,
  RANDOM_SCORE_FACTOR: 0.01,
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
    ULog.debug('自動停止チェック: 無効またはデータ不足', {
      autoStopEnabled: config.autoStopEnabled,
      hasScheduledEndAt: !!config.scheduledEndAt,
    });
    return false;
  }

  const scheduledEndTime = new Date(config.scheduledEndAt);
  const now = new Date();

  ULog.debug('自動停止チェック:', {
    scheduledEndAt: config.scheduledEndAt,
    now: now.toISOString(),
    isOverdue: now >= scheduledEndTime,
  });

  // 期限切れチェック
  if (now >= scheduledEndTime) {
    ULog.warn('期限切れ検出 - 自動停止を実行します');

    // 自動停止前に履歴を保存
    try {
      saveHistoryOnAutoStop(config, userInfo);
    } catch (historyError) {
      logError(
        historyError,
        'autoStopHistorySave',
        MAIN_MAIN_UNIFIED_CONSTANTS.ERROR.SEVERITY.MEDIUM,
        MAIN_MAIN_UNIFIED_CONSTANTS.ERROR.CATEGORIES.DATABASE
      );
      // 履歴保存エラーは処理を継続
    }

    // 自動停止実行
    config.appPublished = false;
    config.autoStoppedAt = now.toISOString();
    config.autoStopReason = 'scheduled_timeout';

    try {
      // データベース更新
      updateUser(userInfo.userId, {
        configJson: JSON.stringify(config),
      });

      ULog.info(`自動停止実行完了: ${userInfo.adminEmail} (期限: ${config.scheduledEndAt})`, {}, ULog.CATEGORIES.WORKFLOW);
      return true; // 自動停止実行済み
    } catch (error) {
      logError(
        error,
        'autoStopProcess',
        MAIN_MAIN_UNIFIED_CONSTANTS.ERROR.SEVERITY.HIGH,
        MAIN_MAIN_UNIFIED_CONSTANTS.ERROR.CATEGORIES.SYSTEM
      );
      return false;
    }
  }

  ULog.debug('まだ期限内です');
  return false; // まだ期限内
}

/**
 * 自動停止時の履歴保存関数
 * @param {Object} config - ユーザーのconfig情報
 * @param {Object} userInfo - ユーザー情報
 */
function saveHistoryOnAutoStop(config, userInfo) {
  ULog.debug('自動停止時履歴保存開始', {}, ULog.CATEGORIES.SYSTEM);

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
    endReason: 'auto_timeout',
  };

  // サーバーサイドでの履歴保存（スプレッドシート）
  try {
    saveHistoryToSheet(historyItem, userInfo);
    ULog.info('自動停止履歴保存完了', { questionText: historyItem.questionText }, ULog.CATEGORIES.SYSTEM);
  } catch (error) {
    ULog.error('serverSideHistorySave failed', { error: error.toString() }, ULog.CATEGORIES.DATABASE);
  }
}

/**
 * configからメイン質問文を取得
 * @param {Object} config - config情報
 * @param {Object} userInfo - ユーザー情報
 * @return {string} 質問文
 */
function getQuestionTextFromConfig(config, userInfo) {
  // 1. sheet固有設定から取得（guessedConfig優先）
  if (config.publishedSheetName) {
    const sheetConfigKey = `sheet_${config.publishedSheetName}`;
    const sheetConfig = config[sheetConfigKey];
    if (sheetConfig) {
      // guessedConfig内のopinionHeaderを優先
      if (sheetConfig.guessedConfig && sheetConfig.guessedConfig.opinionHeader) {
        return sheetConfig.guessedConfig.opinionHeader;
      }
      // フォールバック: 直接のopinionHeader
      if (sheetConfig.opinionHeader) {
        return sheetConfig.opinionHeader;
      }
    }
  }

  // 2. グローバル設定から取得
  if (config.opinionHeader) {
    return config.opinionHeader;
  }

  // 3. カスタムフォーム情報から取得
  if (userInfo.customFormInfo) {
    try {
      const customInfo =
        typeof userInfo.customFormInfo === 'string'
          ? JSON.parse(userInfo.customFormInfo)
          : userInfo.customFormInfo;
      if (customInfo.mainQuestion) {
        return customInfo.mainQuestion;
      }
    } catch (e) {
      ULog.warn('customFormInfo パースエラー', e, ULog.CATEGORIES.SYSTEM);
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

    const sheet = openSpreadsheetOptimized(userInfo.spreadsheetId).getSheetByName(
      config.publishedSheetName
    );
    if (!sheet) {
      return 0;
    }

    const lastRow = sheet.getLastRow();
    return Math.max(0, lastRow - 1); // ヘッダー行を除外
  } catch (error) {
    ULog.warn('回答数取得エラー', error, ULog.CATEGORIES.DATABASE);
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
  ULog.debug('サーバーサイド履歴保存開始', { questionText: historyItem.questionText }, ULog.CATEGORIES.DATABASE);

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
    let configJson;
    try {
      configJson = JSON.parse(existingUser.configJson || '{}');
    } catch (parseError) {
      ULog.warn('configJson解析エラー、新規作成します', { error: parseError.message }, ULog.CATEGORIES.SYSTEM);
      configJson = {};
    }

    // 履歴配列を取得または初期化
    if (!Array.isArray(configJson.historyArray)) {
      configJson.historyArray = [];
    }

    // 軽量化: 大きな構造・長文を除去し、短いインデックス情報のみ保存
    var question = historyItem.questionText || '（問題文未設定）';
    var questionShort = String(question).substring(0, 140);
    var sheetName = String(historyItem.sheetName || '').substring(0, 80);
    var nowIso = new Date().toISOString();

    // 表示モードを短い表現に正規化
    var displayModeText =
      historyItem.displayMode || (historyItem.config && historyItem.config.displayMode) || '';
    if (displayModeText === true || displayModeText === 'named') displayModeText = '通常表示';
    if (displayModeText === false || displayModeText === 'anonymous') displayModeText = '匿名表示';
    if (!displayModeText) displayModeText = '通常表示';

    // ステータス正規化
    var ended = !!historyItem.endTime;
    var statusText = ended ? 'ended' : historyItem.status || 'published';
    var isActive = historyItem.isActive !== undefined ? !!historyItem.isActive : !ended;

    // 可能ならIDベースに置換
    var spreadsheetId = (userInfo && userInfo.spreadsheetId) || historyItem.spreadsheetId || '';
    var formId = null;
    try {
      if (historyItem.formUrl && typeof extractFormIdFromUrl === 'function') {
        formId = extractFormIdFromUrl(historyItem.formUrl);
      }
    } catch (e) {}

    // 新しい履歴アイテム（軽量）
    const serverHistoryItem = {
      id: historyItem.id || 'server_' + Date.now(),
      timestamp: nowIso,
      createdAt: historyItem.createdAt || historyItem.timestamp || nowIso,
      publishedAt: historyItem.publishedAt || nowIso,
      endTime: historyItem.endTime || null,
      scheduledEndTime: historyItem.scheduledEndTime || null,
      questionText: questionShort,
      sheetName: sheetName,
      displayMode: displayModeText,
      status: statusText,
      isActive: isActive,
      spreadsheetId: spreadsheetId,
      formId: formId,
      answerCount: historyItem.answerCount || 0,
      reactionCount: historyItem.reactionCount || 0,
      setupType: historyItem.setupType || 'custom',
    };

    // 重複抑制（直近項目と同一ならスキップ）
    var skipAdd = false;
    if (configJson.historyArray.length > 0) {
      var last = configJson.historyArray[0] || {};
      if (
        last &&
        last.sheetName === serverHistoryItem.sheetName &&
        (last.questionText || '') === serverHistoryItem.questionText &&
        (last.publishedAt || '') === serverHistoryItem.publishedAt
      ) {
        skipAdd = true;
      }
    }
    if (!skipAdd) {
      configJson.historyArray.unshift(serverHistoryItem);
    }

    // 最大履歴件数まで保持
    // 件数制御をMAX_HISTORY_ITEMSと同期（50件まで）
    var SERVER_MAX_HISTORY = MAX_HISTORY_ITEMS; // 50件に拡大
    if (configJson.historyArray.length > SERVER_MAX_HISTORY) {
      configJson.historyArray.splice(SERVER_MAX_HISTORY);
    }

    // configJsonを更新
    configJson.lastModified = new Date().toISOString();

    // サイズ制御: JSONが長すぎる場合は末尾から削除して調整
    var serialized = JSON.stringify(configJson);
    var safetyLimit = 15000; // サイズ制限を緩和（15KBまで）
    var guard = 0;
    var deletedCount = 0;

    while (serialized.length > safetyLimit && configJson.historyArray.length > 10 && guard < 100) {
      configJson.historyArray.pop();
      serialized = JSON.stringify(configJson);
      guard++;
      deletedCount++;
    }

    // 削除が発生した場合はログに記録
    if (deletedCount > 0) {
      ULog.warn('履歴サイズ制限により履歴を削減しました', {
        deletedCount,
        remainingCount: configJson.historyArray.length,
        finalSize: serialized.length,
      });
    }

    // データベースに保存
    const updateResult = updateUser(userInfo.userId, {
      configJson: serialized,
      lastAccessedAt: new Date().toISOString(),
    });

    if (updateResult.success) {
      ULog.info('サーバーサイド履歴保存完了', {
        userId: userInfo.userId,
        questionText: serverHistoryItem.questionText,
        historyCount: configJson.historyArray.length,
      }, ULog.CATEGORIES.DATABASE);
    } else {
      throw new Error('データベース更新に失敗: ' + updateResult.message);
    }
  } catch (error) {
    ULog.error('serverSideHistorySave failed', { error: error.toString() }, ULog.CATEGORIES.DATABASE);
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
      success: true,
      message: '履歴がサーバーに保存されました',
      timestamp: new Date().toISOString(),
    };
  } catch (error) {
    logError(
      error,
      'saveHistoryToSheetAPI',
      MAIN_MAIN_UNIFIED_CONSTANTS.ERROR.SEVERITY.MEDIUM,
      MAIN_MAIN_UNIFIED_CONSTANTS.ERROR.CATEGORIES.DATABASE
    );
    return {
      success: false,
      message: error.message,
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
      configJson = JSON.parse(userInfo.configJson || '{}');
    } catch (parseError) {
      ULog.warn('configJson解析エラー', { error: parseError.message }, ULog.CATEGORIES.SYSTEM);
      configJson = {};
    }

    const historyArray = Array.isArray(configJson.historyArray) ? configJson.historyArray : [];

    return {
      success: true,
      historyArray: historyArray,
      count: historyArray.length,
      lastModified: configJson.lastModified || null,
    };
  } catch (error) {
    logError(
      error,
      'getHistoryFromServerAPI',
      MAIN_MAIN_UNIFIED_CONSTANTS.ERROR.SEVERITY.MEDIUM,
      MAIN_MAIN_UNIFIED_CONSTANTS.ERROR.CATEGORIES.DATABASE
    );
    return {
      success: false,
      message: error.message,
      historyArray: [],
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
      configJson = JSON.parse(userInfo.configJson || '{}');
    } catch (parseError) {
      ULog.warn('configJson解析エラー、新規作成します', { error: parseError.message }, ULog.CATEGORIES.SYSTEM);
      configJson = {};
    }

    // 履歴配列をクリア
    configJson.historyArray = [];
    configJson.lastModified = new Date().toISOString();

    // データベースに保存
    const updateResult = updateUser(requestUserId, {
      configJson: JSON.stringify(configJson),
      lastAccessedAt: new Date().toISOString(),
    });

    if (updateResult.success) {
      ULog.info('✅ サーバーサイド履歴クリア完了:', requestUserId);
      return {
        success: true,
        message: 'サーバーサイドの履歴をクリアしました',
        timestamp: new Date().toISOString(),
      };
    } else {
      throw new Error('データベース更新に失敗: ' + updateResult.message);
    }
  } catch (error) {
    logError(
      error,
      'clearHistoryFromServerAPI',
      MAIN_MAIN_UNIFIED_CONSTANTS.ERROR.SEVERITY.MEDIUM,
      MAIN_MAIN_UNIFIED_CONSTANTS.ERROR.CATEGORIES.DATABASE
    );
    return {
      success: false,
      message: error.message,
    };
  }
}

// EMAIL_REGEX is defined in constants.gs
// DEBUG is defined in constants.gs

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

// getSecurityHeaders function removed - not used in current implementation

// 安定性を重視してconstを使用
const ULTRA_CONFIG = {
  EXECUTION_LIMITS: {
    MAX_TIME: 330000, // 5.5分（安全マージン）
    BATCH_SIZE: 100,
    API_RATE_LIMIT: 90, // 100秒間隔での制限
  },

  CACHE_STRATEGY: {
    L1_TTL: 300, // Level 1: 5分
    L2_TTL: 3600, // Level 2: 1時間
    L3_TTL: 21600, // Level 3: 6時間（最大）
  },
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
        logError(
          message,
          'debugLog',
          MAIN_MAIN_UNIFIED_CONSTANTS.ERROR.SEVERITY.LOW,
          MAIN_MAIN_UNIFIED_CONSTANTS.ERROR.CATEGORIES.SYSTEM,
          { details }
        );
        break;
      case 'warn':
        ULog.warn(message, { details: details || '' }, ULog.CATEGORIES.SYSTEM);
        break;
      default:
        ULog.debug(message, { details: details || '' }, ULog.CATEGORIES.SYSTEM);
    }

    if (typeof globalProfiler !== 'undefined') {
      globalProfiler.end('logging');
    }
  } catch (e) {
    // ログ出力自体が失敗した場合は無視
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
    const webAppUrl = getWebAppUrl();
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
    const isDomainMatch = currentDomain === deployDomain || deployDomain === '';

    ULog.debug('Domain info:', {
      currentDomain: currentDomain,
      deployDomain: deployDomain,
      isDomainMatch: isDomainMatch,
      webAppUrl: webAppUrl,
    });

    return {
      currentDomain: currentDomain,
      deployDomain: deployDomain,
      isDomainMatch: isDomainMatch,
      webAppUrl: webAppUrl,
    };
  } catch (e) {
    logError(
      e,
      'getDeployUserDomainInfo',
      MAIN_MAIN_UNIFIED_CONSTANTS.ERROR.SEVERITY.MEDIUM,
      MAIN_MAIN_UNIFIED_CONSTANTS.ERROR.CATEGORIES.SYSTEM
    );
    return {
      currentDomain: '不明',
      deployDomain: '不明',
      isDomainMatch: false,
      error: e.message,
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
    ULog.debug('Getting GOOGLE_CLIENT_ID from script properties...');
    const properties = PropertiesService.getScriptProperties();
    const clientId = properties.getProperty('GOOGLE_CLIENT_ID');

    ULog.debug('GOOGLE_CLIENT_ID retrieved:', clientId ? 'Found' : 'Not found');

    if (!clientId) {
      ULog.warn('GOOGLE_CLIENT_ID not found in script properties');

      // Try to get all properties to see what's available
      const allProperties = properties.getProperties();
      ULog.debug('Available properties:', Object.keys(allProperties));

      return {
        clientId: '',
        error: 'GOOGLE_CLIENT_ID not found in script properties',
        setupInstructions:
          'Please set GOOGLE_CLIENT_ID in Google Apps Script project settings under Properties > Script Properties',
      };
    }

    return createSuccessResponse({ clientId: clientId }, 'Google Client IDを取得しました');
  } catch (error) {
    logError(
      error,
      'getGoogleClientId',
      MAIN_MAIN_UNIFIED_CONSTANTS.ERROR.SEVERITY.HIGH,
      MAIN_MAIN_UNIFIED_CONSTANTS.ERROR.CATEGORIES.SYSTEM
    );
    return createErrorResponse(error.toString(), 'Google Client IDの取得に失敗しました', {
      clientId: '',
    });
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
      'SERVICE_ACCOUNT_CREDS',
    ];

    const configStatus = {};
    const missingProperties = [];

    requiredProperties.forEach(function (prop) {
      const value = allProperties[prop];
      configStatus[prop] = {
        exists: !!value,
        hasValue: !!(value && value.trim()),
        length: value ? value.length : 0,
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
      setupComplete: isSystemSetup(),
    };
  } catch (error) {
    ULog.error('checkSystemConfiguration failed', { error: error.toString() }, ULog.CATEGORIES.SYSTEM);
    return {
      isFullyConfigured: false,
      error: error.toString(),
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
      deployDomain: domainInfo.deployDomain || adminDomain,
    };
  } catch (e) {
    ULog.error('getSystemDomainInfo failed', { error: e.message }, ULog.CATEGORIES.SYSTEM);
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
    ULog.error('doGet failed', { error: error.toString() }, ULog.CATEGORIES.SYSTEM);
    return showErrorPage(
      '致命的なエラー',
      'アプリケーションの処理中に予期せぬエラーが発生しました。',
      error
    );
  }
}

/**
 * システムコンポーネントの依存関係をチェック
 * @returns {Object} チェック結果
 */
function validateSystemDependencies() {
  const errors = [];

  try {
    // 軽量化されたシステム依存関係チェック（パフォーマンス最適化）

    // PropertiesService の基本的な存在確認のみ
    try {
      const props = PropertiesService.getScriptProperties();
      if (!props || typeof props.getProperty !== 'function') {
        errors.push('PropertiesService が利用できません');
      }
      // 実際のアクセステストは削除（パフォーマンス改善）
    } catch (propsError) {
      errors.push(`PropertiesService エラー: ${propsError.message}`);
    }

    // resilientExecutor の基本存在確認のみ
    try {
      if (typeof resilientExecutor === 'undefined') {
        errors.push('resilientExecutor が定義されていません');
      }
      // 詳細機能テストは削除（パフォーマンス改善）
    } catch (executorError) {
      errors.push(`resilientExecutor エラー: ${executorError.message}`);
    }

    // secretManager の詳細診断は削除（パフォーマンス改善）
    // 必要に応じて後で実行

    // CacheService テスト
    try {
      CacheService.getScriptCache().get('_DEPENDENCY_TEST_KEY');
    } catch (cacheError) {
      errors.push(`CacheService エラー: ${cacheError.message}`);
    }

    // Utilities テスト
    try {
      if (typeof Utilities === 'undefined' || typeof Utilities.getUuid !== 'function') {
        errors.push('Utilities サービスが利用できません');
      }
    } catch (utilsError) {
      errors.push(`Utilities エラー: ${utilsError.message}`);
    }
  } catch (generalError) {
    errors.push(`システムチェック中の一般エラー: ${generalError.message}`);
  }

  return {
    success: errors.length === 0,
    errors: errors,
    timestamp: new Date().toISOString(),
  };
}

/**
 * Initialize request processing with system checks
 * @returns {HtmlOutput|null} Early return result or null to continue
 */
function initializeRequestProcessing() {
  // Clear execution-level cache for new request
  clearAllExecutionCache();

  // システムコンポーネントの依存関係チェック
  const dependencyCheck = validateSystemDependencies();
  if (!dependencyCheck.success) {
    ULog.error('システムの依存関係チェックに失敗', { errors: dependencyCheck.errors }, ULog.CATEGORIES.SYSTEM);
    return showErrorPage(
      'システム初期化エラー',
      'システムの初期化に失敗しました。管理者にご連絡ください。\n\n' +
        `エラー詳細: ${dependencyCheck.errors.join(', ')}`
    );
  }

  // Check system setup (highest priority)
  if (!isSystemSetup()) {
    return showSetupPage();
  }

  // Check application access permissions
  const accessCheck = checkApplicationAccess();
  if (!accessCheck.hasAccess) {
    ULog.info('アプリケーションアクセス拒否', { reason: accessCheck.accessReason }, ULog.CATEGORIES.AUTH);
    return showErrorPage(
      'アクセスが制限されています',
      'アプリケーションは現在利用できません。管理者にお問い合わせください。'
    );
  }

  return null; // Continue processing
}

/**
 * Validate user authentication
 * @returns {HtmlOutput|null} Early return result or null to continue
 */
function validateUserAuthentication() {
  try {
    // 統一Validationシステムを使用（利用可能な場合）
    if (typeof UnifiedValidation !== 'undefined') {
      const userEmail = getCurrentUserEmail();
      const result = UnifiedValidation.validate('authentication', 'basic', { 
        userId: userEmail,
        userEmail: userEmail
      });
      
      if (!result.success) {
        ULog.debug('validateUserAuthentication: Unified validation failed, showing login page', {
          result: result
        }, ULog.CATEGORIES.AUTH);
        return showLoginPage();
      }
      
      ULog.debug('validateUserAuthentication: Unified validation successful', {
        passRate: result.summary.passRate
      }, ULog.CATEGORIES.AUTH);
      return null; // Continue processing
    }
    
    // フォールバック: 既存の実装
    ULog.debug('validateUserAuthentication: Starting authentication check', {}, ULog.CATEGORIES.AUTH);
    const userEmail = getCurrentUserEmail();
    ULog.debug('validateUserAuthentication: userEmail from getCurrentUserEmail', { userEmail }, ULog.CATEGORIES.AUTH);
    if (!userEmail) {
      ULog.debug('validateUserAuthentication: userEmail is empty, showing login page', {}, ULog.CATEGORIES.AUTH);
      return showLoginPage();
    }
    ULog.debug('validateUserAuthentication: userEmail is present, continuing processing', {}, ULog.CATEGORIES.AUTH);
    return null; // Continue processing
  } catch (error) {
    ULog.error('validateUserAuthentication: Validation failed with error', {
      error: error.message
    }, ULog.CATEGORIES.AUTH);
    return showLoginPage();
  }
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
  ULog.debug('No mode parameter, checking previous admin session', {}, ULog.CATEGORIES.AUTH);

  const activeUserEmail = getCurrentUserEmail();
  if (!activeUserEmail) {
    return showLoginPage();
  }

  const userProperties = PropertiesService.getUserProperties();
  const lastAdminUserId = userProperties.getProperty('lastAdminUserId');

  if (lastAdminUserId && verifyAdminAccess(lastAdminUserId)) {
    ULog.debug('Found previous admin session, redirecting to admin panel', { lastAdminUserId }, ULog.CATEGORIES.AUTH);
    ULog.debug('Calling findUserById with lastAdminUserId', { lastAdminUserId }, ULog.CATEGORIES.AUTH);
    const userInfo = findUserById(lastAdminUserId);
    return renderAdminPanel(userInfo, 'admin');
  }

  // Clear invalid admin session
  if (lastAdminUserId) {
    userProperties.deleteProperty('lastAdminUserId');
  }

  ULog.debug('No previous admin session, showing login page', {}, ULog.CATEGORIES.AUTH);
  return showLoginPage();
}

/**
 * Handle login mode
 * @returns {HtmlOutput} Login page
 */
function handleLoginMode() {
  ULog.debug('Login mode requested, showing login page', {}, ULog.CATEGORIES.AUTH);
  return showLoginPage();
}

/**
 * Handle app setup mode
 * @returns {HtmlOutput} App setup page or error page
 */
function handleAppSetupMode() {
  ULog.debug('AppSetup mode requested', {}, ULog.CATEGORIES.SYSTEM);

  const userProperties = PropertiesService.getUserProperties();
  const lastAdminUserId = userProperties.getProperty('lastAdminUserId');

  if (!lastAdminUserId) {
    ULog.debug('No admin session found, redirecting to login', {}, ULog.CATEGORIES.AUTH);
    return showErrorPage('認証が必要です', 'アプリ設定にアクセスするにはログインが必要です。');
  }

  if (!verifyAdminAccess(lastAdminUserId)) {
    ULog.warn('Admin access denied for userId', { lastAdminUserId }, ULog.CATEGORIES.AUTH);
    return showErrorPage('アクセス拒否', 'アプリ設定にアクセスする権限がありません。');
  }

  ULog.debug('Showing app setup page for userId', { lastAdminUserId }, ULog.CATEGORIES.SYSTEM);
  return showAppSetupPage(lastAdminUserId);
}

/**
 * Handle admin mode
 * @param {Object} params - Request parameters
 * @returns {HtmlOutput} Admin panel or error page
 */
function handleAdminMode(params) {
  const requestStartTime = Date.now();

  if (!params.userId) {
    return showErrorPage('不正なリクエスト', 'ユーザーIDが指定されていません。');
  }

  // システム状態診断
  const systemDiagnostics = {
    requestTime: new Date().toISOString(),
    userId: params.userId,
    userEmail: getCurrentUserEmail(),
    cacheStatus: {},
    databaseConnectivity: 'unknown',
    performanceMetrics: {},
  };

  try {
    // キャッシュ状態確認
    try {
      const scriptCache = CacheService.getScriptCache();
      systemDiagnostics.cacheStatus.scriptCache = 'available';
      systemDiagnostics.cacheStatus.executionCache = 'available';
    } catch (cacheError) {
      systemDiagnostics.cacheStatus.error = cacheError.message;
    }

    // データベース接続性テスト
    try {
      const dbId = getSecureDatabaseId();
      systemDiagnostics.databaseConnectivity = dbId ? 'connected' : 'disconnected';
    } catch (dbError) {
      systemDiagnostics.databaseConnectivity = 'error: ' + dbError.message;
    }

    ULog.info('handleAdminMode: システム診断完了', systemDiagnostics, ULog.CATEGORIES.SYSTEM);
  } catch (diagError) {
    ULog.warn('handleAdminMode: システム診断でエラー', { error: diagError.message }, ULog.CATEGORIES.SYSTEM);
  }

  // 管理者権限確認（詳細ログ付き）
  ULog.debug('handleAdminMode: 統合管理者権限確認開始', {
    userId: params.userId,
    timestamp: new Date().toISOString(),
    systemStatus: systemDiagnostics,
  }, ULog.CATEGORIES.AUTH);

  const authStartTime = Date.now();
  const adminAccessResult = verifyAdminAccess(params.userId);
  const authDuration = Date.now() - authStartTime;

  systemDiagnostics.performanceMetrics.authDuration = authDuration + 'ms';

  if (!adminAccessResult) {
    const totalRequestTime = Date.now() - requestStartTime;
    systemDiagnostics.performanceMetrics.totalRequestTime = totalRequestTime + 'ms';

    ULog.error('handleAdminMode: 管理者権限確認失敗', {
      userId: params.userId,
      currentUser: getCurrentUserEmail(),
      authDuration: authDuration + 'ms',
      totalTime: totalRequestTime + 'ms',
      systemDiagnostics: systemDiagnostics,
      timestamp: new Date().toISOString(),
    }, ULog.CATEGORIES.AUTH);

    // 詳細な診断情報付きエラーページ
    let propertiesDiagnostics = 'unknown';
    try {
      const userProps = PropertiesService.getUserProperties();
      const scriptProps = PropertiesService.getScriptProperties();

      const userPropsData = userProps.getProperties();
      const allScriptProps = scriptProps.getProperties();
      const scriptPropsKeys = Object.keys(allScriptProps).filter((k) => k.startsWith('newUser_'));

      // より詳細な診断情報
      const newUserDetails = scriptPropsKeys.map((key) => {
        try {
          const data = JSON.parse(allScriptProps[key]);
          const timeDiff = Date.now() - parseInt(data.createdTime);
          return {
            key: key,
            email: data.email,
            userId: data.userId,
            ageMinutes: Math.floor(timeDiff / 60000),
          };
        } catch (e) {
          return { key: key, error: 'parse_failed' };
        }
      });

      propertiesDiagnostics = {
        userProperties: Object.keys(userPropsData).length,
        scriptProperties: scriptPropsKeys.length,
        recentUsers: scriptPropsKeys.slice(0, 3), // 最新3件のキー
        newUserDetails: newUserDetails.slice(0, 5), // 詳細情報（最新5件）
        currentUser: params.userId,
        currentEmail: getCurrentUserEmail(),
      };
    } catch (propError) {
      propertiesDiagnostics = 'error: ' + propError.message;
    }

    const diagnosticInfo = [
      `ユーザーID: ${params.userId}`,
      `現在のメール: ${getCurrentUserEmail()}`,
      `認証時間: ${authDuration}ms`,
      `総処理時間: ${totalRequestTime}ms`,
      `データベース接続: ${systemDiagnostics.databaseConnectivity}`,
      `プロパティ状態: ${JSON.stringify(propertiesDiagnostics)}`,
      `時刻: ${new Date().toLocaleString('ja-JP')}`,
    ].join('\n');

    return showErrorPage(
      'アクセス拒否',
      'アカウントが一時的に無効化されています。\n\n' +
        '対処法:\n' +
        '• 新規登録から1-2分お待ちください\n' +
        '• ブラウザを更新してお試しください\n' +
        '• 問題が続く場合は管理者にお問い合わせください\n\n' +
        '詳細診断情報:\n' +
        diagnosticInfo
    );
  }

  const totalRequestTime = Date.now() - requestStartTime;
  systemDiagnostics.performanceMetrics.totalRequestTime = totalRequestTime + 'ms';

  ULog.info('handleAdminMode: 統合管理者権限確認成功', {
    userId: params.userId,
    authDuration: authDuration + 'ms',
    totalTime: totalRequestTime + 'ms',
    systemDiagnostics: systemDiagnostics,
  }, ULog.CATEGORIES.AUTH);

  // Save admin session state
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('lastAdminUserId', params.userId);
  userProperties.setProperty('lastSuccessfulAdminAccess', Date.now().toString());
  ULog.debug('Saved enhanced admin session state', { userId: params.userId }, ULog.CATEGORIES.AUTH);

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

  // Get user info optimized for viewer access with enhanced caching
  const userInfo = findUserByIdForViewer(params.userId);

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
    config = JSON.parse(userInfo.configJson || '{}');
  } catch (e) {
    ULog.warn('Config JSON parse error during publication check:', e.message);
  }

  // Check for auto-stop and handle accordingly
  const wasAutoStopped = checkAndHandleAutoStop(config, userInfo);
  if (wasAutoStopped) {
    ULog.info('🔄 自動停止が実行されました - 非公開ページに誘導します');
  }

  // Check if currently published
  const isCurrentlyPublished = !!(
    config.appPublished === true &&
    config.publishedSpreadsheetId &&
    config.publishedSheetName &&
    typeof config.publishedSheetName === 'string' &&
    config.publishedSheetName.trim() !== ''
  );

  ULog.debug('🔍 Publication status check:', {
    appPublished: config.appPublished,
    hasSpreadsheetId: !!config.publishedSpreadsheetId,
    hasSheetName: !!config.publishedSheetName,
    isCurrentlyPublished: isCurrentlyPublished,
  });

  // Redirect to unpublished page if not published
  if (!isCurrentlyPublished) {
    ULog.info('🚫 Board is unpublished, redirecting to Unpublished page');
    return renderUnpublishedPage(userInfo, params);
  }

  return renderAnswerBoard(userInfo, params);
}

/**
 * Handle unknown mode
 * @param {Object} params - Request parameters
 * @returns {HtmlOutput} Appropriate page response
 */
function handleUnknownMode(params) {
  ULog.warn('Unknown mode received:', params.mode);
  ULog.debug('Available modes: login, appSetup, admin, view');

  // If valid userId with admin access, redirect to admin panel
  if (params.userId && verifyAdminAccess(params.userId)) {
    ULog.debug('Redirecting unknown mode to admin panel for valid user:', params.userId);
    const userInfo = findUserById(params.userId);
    return renderAdminPanel(userInfo, 'admin');
  }

  // Otherwise redirect to login
  ULog.debug('Redirecting unknown mode to login page');
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
    ULog.warn(
      `不正アクセス試行: ${userEmail} が userId ${params.userId} にアクセスしようとしました。`
    );
    const correctUrl = buildUserAdminUrl(userInfo.userId);
    return redirectToUrl(correctUrl);
  }

  // 強化されたセキュリティ検証: 指定されたIDの登録メールアドレスと現在ログイン中のGoogleアカウントが一致するかを検証
  if (params.userId) {
    const isVerified = verifyAdminAccess(params.userId);
    if (!isVerified) {
      ULog.warn(
        `セキュリティ検証失敗: userId ${params.userId} への不正アクセス試行をブロックしました。`
      );
      const correctUrl = buildUserAdminUrl(userInfo.userId);
      return redirectToUrl(correctUrl);
    }
    ULog.debug(`✅ セキュリティ検証成功: userId ${params.userId} への正当なアクセスを確認しました。`);
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
// getOrFetchUserInfo() は unifiedUserManager.gs に統合済み
/**
 * ログインページを表示
 * @returns {HtmlOutput}
 */
function showLoginPage() {
  const template = HtmlService.createTemplateFromFile('LoginPage');
  const htmlOutput = template.evaluate().setTitle('StudyQuest - ログイン');

  // XFrameOptionsMode を安全に設定
  try {
    if (HtmlService && HtmlService.XFrameOptionsMode && HtmlService.XFrameOptionsMode.ALLOWALL) {
      htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  } catch (e) {
    ULog.warn('XFrameOptionsMode設定エラー:', e.message);
  }

  return htmlOutput;
}

/**
 * 初回セットアップページを表示
 * @returns {HtmlOutput}
 */
function showSetupPage() {
  const template = HtmlService.createTemplateFromFile('SetupPage');
  const htmlOutput = template.evaluate().setTitle('StudyQuest - 初回セットアップ');

  // XFrameOptionsMode を安全に設定
  try {
    if (HtmlService && HtmlService.XFrameOptionsMode && HtmlService.XFrameOptionsMode.DENY) {
      htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DENY);
    }
  } catch (e) {
    ULog.warn('XFrameOptionsMode設定エラー:', e.message);
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
    ULog.debug('showAppSetupPage: Checking deploy user permissions...');
    const currentUserEmail = getCurrentUserEmail();
    ULog.debug('showAppSetupPage: Current user email:', currentUserEmail);
    const deployUserCheckResult = isDeployUser();
    ULog.debug('showAppSetupPage: isDeployUser() result:', deployUserCheckResult);

    if (!deployUserCheckResult) {
      ULog.warn('Unauthorized access attempt to app setup page:', currentUserEmail);
      return showErrorPage(
        'アクセス権限がありません',
        'この機能にアクセスする権限がありません。システム管理者にお問い合わせください。'
      );
    }
  } catch (error) {
    logError(
      error,
      'checkDeployUserPermissions',
      MAIN_MAIN_UNIFIED_CONSTANTS.ERROR.SEVERITY.HIGH,
      MAIN_MAIN_UNIFIED_CONSTANTS.ERROR.CATEGORIES.AUTHORIZATION
    );
    return showErrorPage('認証エラー', '権限確認中にエラーが発生しました。');
  }

  const appSetupTemplate = HtmlService.createTemplateFromFile('AppSetupPage');
  appSetupTemplate.userId = userId;
  const htmlOutput = appSetupTemplate.evaluate().setTitle('アプリ設定 - StudyQuest');

  // XFrameOptionsMode を安全に設定
  try {
    if (HtmlService && HtmlService.XFrameOptionsMode && HtmlService.XFrameOptionsMode.DENY) {
      htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DENY);
    }
  } catch (e) {
    ULog.warn('XFrameOptionsMode設定エラー:', e.message);
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
      ULog.debug('Found valid admin user ID:', lastAdminUserId);
      return lastAdminUserId;
    } else {
      ULog.debug('No valid admin user ID found');
      return null;
    }
  } catch (error) {
    logError(
      error,
      'getLastAdminUserId',
      MAIN_MAIN_UNIFIED_CONSTANTS.ERROR.SEVERITY.MEDIUM,
      MAIN_MAIN_UNIFIED_CONSTANTS.ERROR.CATEGORIES.DATABASE
    );
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
    ULog.debug('getAppSetupUrl: Checking deploy user permissions...');
    const currentUserEmail = getCurrentUserEmail();
    ULog.debug('getAppSetupUrl: Current user email:', currentUserEmail);
    const deployUserCheckResult = isDeployUser();
    ULog.debug('getAppSetupUrl: isDeployUser() result:', deployUserCheckResult);

    if (!deployUserCheckResult) {
      ULog.warn('Unauthorized access attempt to get app setup URL:', currentUserEmail);
      throw new Error('アクセス権限がありません');
    }

    // WebアプリのベースURLを取得
    const baseUrl = ScriptApp.getService().getUrl();
    if (!baseUrl) {
      throw new Error('WebアプリのURLを取得できませんでした');
    }

    // アプリ設定ページのURLを生成
    const appSetupUrl = baseUrl + '?mode=appSetup';
    ULog.debug('getAppSetupUrl: Generated URL:', appSetupUrl);

    return appSetupUrl;
  } catch (error) {
    logError(
      error,
      'getAppSetupUrl',
      MAIN_MAIN_UNIFIED_CONSTANTS.ERROR.SEVERITY.MEDIUM,
      MAIN_MAIN_UNIFIED_CONSTANTS.ERROR.CATEGORIES.SYSTEM
    );
    throw new Error('アプリ設定URLの取得に失敗しました: ' + error.message);
  }
}

// =================================================================
// ERROR HANDLING & CATEGORIZATION
// =================================================================

// エラータイプ定義
// ERROR_TYPES is defined in constants.gs

/**
 * エラータイプに基づいてメッセージを分類・整理する
 * @param {string} title - エラータイトル
 * @param {string} message - エラーメッセージ
 * @returns {Object} 分類されたエラー情報
 */
function categorizeError(title, message) {
  const titleLower = title.toLowerCase();
  const messageLower = message.toLowerCase();

  // エラータイプの判定
  let errorType = ERROR_TYPES.USER;
  if (titleLower.includes('致命的') || titleLower.includes('システム')) {
    errorType = ERROR_TYPES.CRITICAL;
  } else if (
    titleLower.includes('アクセス') ||
    titleLower.includes('認証') ||
    titleLower.includes('権限')
  ) {
    errorType = ERROR_TYPES.ACCESS;
  } else if (titleLower.includes('不正') || messageLower.includes('指定されていません')) {
    errorType = ERROR_TYPES.VALIDATION;
  } else if (messageLower.includes('ネットワーク') || messageLower.includes('接続')) {
    errorType = ERROR_TYPES.NETWORK;
  }

  return {
    type: errorType,
    icon: getErrorIcon(errorType),
    severity: getErrorSeverity(errorType),
  };
}

/**
 * エラータイプに対応するアイコンを取得
 */
function getErrorIcon(errorType) {
  const icons = {
    [ERROR_TYPES.CRITICAL]: '🔥',
    [ERROR_TYPES.ACCESS]: '🔒',
    [ERROR_TYPES.VALIDATION]: '⚠️',
    [ERROR_TYPES.NETWORK]: '🌐',
    [ERROR_TYPES.USER]: '❓',
  };
  return icons[errorType] || '⚠️';
}

// getErrorSeverity() は constants.gs に統合済み

/**
 * 長い診断情報を構造化して整理する
 * @param {string} diagnosticInfo - 診断情報文字列
 * @returns {Object} 構造化された診断情報
 */
function structureDiagnosticInfo(diagnosticInfo) {
  if (!diagnosticInfo) return null;

  const lines = diagnosticInfo.split('\n');
  const structured = {
    summary: [],
    technical: [],
    properties: null,
  };

  let currentSection = 'summary';

  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) continue;

    // プロパティ状態のJSON部分を検出
    if (trimmed.includes('プロパティ状態:')) {
      currentSection = 'properties';
      const jsonStart = line.indexOf('{');
      if (jsonStart !== -1) {
        try {
          const jsonStr = line.substring(jsonStart);
          structured.properties = JSON.parse(jsonStr);
        } catch (e) {
          structured.technical.push(line);
        }
      }
      continue;
    }

    // 基本情報と技術情報を分類
    if (
      trimmed.startsWith('ユーザーID:') ||
      trimmed.startsWith('現在のメール:') ||
      trimmed.startsWith('認証時間:') ||
      trimmed.startsWith('時刻:')
    ) {
      structured.summary.push(trimmed);
    } else {
      structured.technical.push(trimmed);
    }
  }

  return structured;
}

/**
 * エラーページを表示する関数（改善版）
 * @param {string} title - エラータイトル
 * @param {string} message - エラーメッセージ
 * @param {Error|string} error - エラーオブジェクトまたは診断情報
 * @returns {HtmlOutput} エラーページのHTML出力
 */
function showErrorPage(title, message, error) {
  const template = HtmlService.createTemplateFromFile('ErrorBoundary');

  // エラー分類
  const errorInfo = categorizeError(title, message);

  // 基本情報設定
  template.title = title;
  template.errorType = errorInfo.type;
  template.errorIcon = errorInfo.icon;
  template.errorSeverity = errorInfo.severity;
  template.mode = 'admin';

  // メッセージを構造化
  if (message && message.includes('詳細診断情報:')) {
    const parts = message.split('詳細診断情報:');
    template.message = parts[0].trim();
    template.diagnosticInfo = structureDiagnosticInfo(parts[1]);
  } else {
    template.message = message;
    template.diagnosticInfo = null;
  }

  // 現在のユーザーがデータベースに登録されているかチェック
  let isRegisteredUser = false;
  let currentUserEmail = '';
  try {
    currentUserEmail = getCurrentUserEmail();
    if (currentUserEmail) {
      const userInfo = findUserByEmail(currentUserEmail);
      isRegisteredUser = !!userInfo;
    }
  } catch (e) {
    ULog.warn('⚠️ showErrorPage: ユーザー登録状態の確認でエラー:', e);
  }

  template.isRegisteredUser = isRegisteredUser;
  template.userEmail = currentUserEmail;

  // デバッグ情報設定
  if (DEBUG && error) {
    if (typeof error === 'string') {
      template.debugInfo = error;
    } else if (error.stack) {
      template.debugInfo = error.stack;
    } else {
      template.debugInfo = error.toString();
    }
  } else {
    template.debugInfo = '';
  }

  const htmlOutput = template.evaluate().setTitle(`エラー - ${title}`);

  // XFrameOptionsMode を安全に設定
  try {
    if (HtmlService && HtmlService.XFrameOptionsMode && HtmlService.XFrameOptionsMode.DENY) {
      htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DENY);
    }
  } catch (e) {
    ULog.warn('XFrameOptionsMode設定エラー:', e.message);
  }

  return htmlOutput;
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
  login: function () {
    const baseUrl = getWebAppUrl();
    return `${baseUrl}?mode=login`;
  },

  /**
   * 管理パネルのURLを生成
   * @param {string} userId - ユーザーID
   * @returns {string} 管理パネルURL
   */
  admin: function (userId) {
    const baseUrl = getWebAppUrl();
    return `${baseUrl}?mode=admin&userId=${encodeURIComponent(userId)}`;
  },

  /**
   * アプリ設定ページのURLを生成
   * @returns {string} アプリ設定ページURL
   */
  appSetup: function () {
    const baseUrl = getWebAppUrl();
    return `${baseUrl}?mode=appSetup`;
  },

  /**
   * 回答ボードのURLを生成
   * @param {string} userId - ユーザーID
   * @returns {string} 回答ボードURL
   */
  view: function (userId) {
    const baseUrl = getWebAppUrl();
    return `${baseUrl}?mode=view&userId=${encodeURIComponent(userId)}`;
  },

  /**
   * パラメータ付きURLを安全に生成
   * @param {string} mode - モード ('login', 'admin', 'view', 'appSetup')
   * @param {Object} params - 追加パラメータ
   * @returns {string} 生成されたURL
   */
  build: function (mode, params = {}) {
    const baseUrl = getWebAppUrl();
    const url = new URL(baseUrl);
    url.searchParams.set('mode', mode);

    Object.keys(params).forEach((key) => {
      if (params[key] !== null && params[key] !== undefined) {
        url.searchParams.set(key, params[key]);
      }
    });

    return url.toString();
  },
};

/**
 * 指定されたURLへリダイレクトするサーバーサイド関数
 * @param {string} url - リダイレクト先のURL
 * @returns {HtmlOutput} リダイレクトを実行するHTML出力
 */
function redirectToUrl(url) {
  // XSS攻撃を防ぐため、URLをサニタイズ
  const sanitizedUrl = sanitizeRedirectUrl(url);
  return HtmlService.createHtmlOutput().setContent(
    `<script>window.top.location.href = '${sanitizedUrl}';</script>`
  );
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

  ULog.debug('createSecureRedirect - Original URL:', targetUrl);
  ULog.debug('createSecureRedirect - Sanitized URL:', sanitizedUrl);

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
        @keyframes pulse {
          0% { transform: scale(1); }
          50% { transform: scale(1.05); }
          100% { transform: scale(1); }
        }
      </style>
    </head>
    <body>
      <div class="container">
        <div class="icon">🔐</div>
        <h1 class="title">${message || 'アクセス確認'}</h1>
        <p class="subtitle">セキュリティのため、下のボタンをクリックして続行してください</p>

        <a href="${sanitizedUrl}" target="_top" class="main-button" onclick="handleSecureRedirect(event, '${sanitizedUrl}')">
          🚀 続行する
        </a>

        <div class="url-info">
          <div class="url-text">${sanitizedUrl}</div>
        </div>

        <div class="note">
          ✓ このリンクは安全です<br>
          ✓ Google Apps Script公式のセキュリティガイドラインに準拠<br>
          ✓ iframe制約回避のため新しいタブで開きます
        </div>
        
        <script>
          function handleSecureRedirect(event, url) {
            try {
              // iframe内かどうかを判定
              const isInFrame = (window !== window.top);
              
              if (isInFrame) {
                // iframe内の場合は親ウィンドウで開く
                event.preventDefault();
                ULog.info('🔄 iframe内からの遷移を検出、parent window で開きます');
                window.top.location.href = url;
              } else {
                // 通常の場合はそのまま遷移
                ULog.info('🚀 通常の遷移を実行します');
                // target="_top" が有効になります
              }
            } catch (error) {
              ULog.error('リダイレクト処理エラー:', error);
              // エラーの場合はフォールバック
              window.location.href = url;
            }
          }
          
          // 自動遷移を無効化（X-Frame-Options制約のためユーザーアクション必須）
          // ユーザーに明確なアクションを要求することで、より確実な遷移を実現
          ULog.info('ℹ️ 自動遷移は無効です。ユーザーによるボタンクリックが必要です。');
          
          // 代わりに、5秒後にボタンを強調表示
          setTimeout(function() {
            const mainButton = document.querySelector('.main-button');
            if (mainButton) {
              mainButton.style.animation = 'pulse 1s infinite';
              mainButton.style.boxShadow = '0 0 20px rgba(16, 185, 129, 0.5)';
              ULog.info('✨ ボタンを強調表示しました');
            }
          }, 3000);
        </script>
      </div>
    </body>
    </html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(userActionRedirectHtml);

  // XFrameOptionsMode を安全に設定（セキュアリダイレクト用）
  try {
    if (HtmlService && HtmlService.XFrameOptionsMode && HtmlService.XFrameOptionsMode.ALLOWALL) {
      htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      ULog.debug('✅ Secure Redirect XFrameOptionsMode.ALLOWALL設定完了');
    } else {
      ULog.warn('⚠️ HtmlService.XFrameOptionsMode.ALLOWALLが利用できません');
    }
  } catch (e) {
    ULog.error('[ERROR]', '❌ Secure Redirect XFrameOptionsMode設定エラー:', e.message);
    // フォールバック: 従来の方法で設定を試行
    try {
      htmlOutput.setXFrameOptionsMode('ALLOWALL');
      ULog.info('💡 フォールバック方法でSecure Redirect XFrameOptionsMode設定完了');
    } catch (fallbackError) {
      ULog.error('[ERROR]', '❌ フォールバック方法も失敗:', fallbackError.message);
    }
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
    return getWebAppUrl();
  }

  try {
    let cleanUrl = String(url).trim();

    // 複数レベルのクォート除去（JSON文字列化による多重クォートに対応）
    let previousUrl = '';
    while (cleanUrl !== previousUrl) {
      previousUrl = cleanUrl;

      // 先頭と末尾のクォートを除去
      if (
        (cleanUrl.startsWith('"') && cleanUrl.endsWith('"')) ||
        (cleanUrl.startsWith("'") && cleanUrl.endsWith("'"))
      ) {
        cleanUrl = cleanUrl.slice(1, -1);
      }

      // エスケープされたクォートを除去
      cleanUrl = cleanUrl.replace(/\\"/g, '"').replace(/\\'/g, "'");

      // URL内に埋め込まれた別のURLを検出
      const embeddedUrlMatch = cleanUrl.match(/https?:\/\/[^\s<>"']+/);
      if (embeddedUrlMatch && embeddedUrlMatch[0] !== cleanUrl) {
        ULog.debug('Extracting embedded URL:', embeddedUrlMatch[0]);
        cleanUrl = embeddedUrlMatch[0];
      }
    }

    // 基本的なURL形式チェック
    if (!cleanUrl.match(/^https?:\/\/[^\s<>"']+$/)) {
      ULog.warn('Invalid URL format after sanitization:', cleanUrl);
      return getWebAppUrl();
    }

    // 開発モードURLのチェック（googleusercontent.comは有効なデプロイURLも含むため調整）
    if (cleanUrl.includes('userCodeAppPanel')) {
      ULog.warn('Development URL detected in redirect, using fallback:', cleanUrl);
      return getWebAppUrl();
    }

    // 最終的な URL 妥当性チェック（googleusercontent.comも有効URLとして認識）
    const isValidUrl =
      cleanUrl.includes('script.google.com') ||
      cleanUrl.includes('googleusercontent.com') ||
      cleanUrl.includes('localhost');

    if (!isValidUrl) {
      ULog.warn('Suspicious URL detected:', cleanUrl);
      return getWebAppUrl();
    }

    return cleanUrl;
  } catch (e) {
    logError(
      e,
      'urlSanitization',
      MAIN_MAIN_UNIFIED_CONSTANTS.ERROR.SEVERITY.HIGH,
      MAIN_MAIN_UNIFIED_CONSTANTS.ERROR.CATEGORIES.SYSTEM
    );
    return getWebAppUrl();
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
    ULog.debug('parseRequestParams: 無効なeventオブジェクト');
    return {
      mode: null,
      userId: null,
      setupParam: null,
      spreadsheetId: null,
      sheetName: null,
      isDirectPageAccess: false,
    };
  }

  const p = e.parameter || {};
  const mode = p.mode || null;
  const userId = p.userId || null;
  const setupParam = p.setup || null;
  const spreadsheetId = p.spreadsheetId || null;
  const sheetName = p.sheetName || null;
  const isDirectPageAccess = !!(userId && mode === 'view');

  // デバッグログを追加
  ULog.debug('parseRequestParams - Received parameters:', JSON.stringify(p));
  ULog.debug('parseRequestParams - Parsed mode:', mode, 'setupParam:', setupParam);

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
    logError(
      'renderAdminPanelにuserInfoがnullで渡されました',
      'renderAdminPanel',
      MAIN_MAIN_UNIFIED_CONSTANTS.ERROR.SEVERITY.HIGH,
      MAIN_MAIN_UNIFIED_CONSTANTS.ERROR.CATEGORIES.SYSTEM
    );
    return showErrorPage(
      'エラー',
      'ユーザー情報の読み込みに失敗したため、管理パネルを表示できません。'
    );
  }

  const adminTemplate = HtmlService.createTemplateFromFile('AdminPanel');
  adminTemplate.include = include;
  adminTemplate.userInfo = userInfo;
  adminTemplate.userId = userInfo.userId;
  adminTemplate.mode = mode || 'admin'; // 安全のためのフォールバック
  adminTemplate.displayMode = 'named';
  adminTemplate.showAdminFeatures = true;
  const deployUserResult = isDeployUser();
  const currentUserEmail = getCurrentUserEmail();
  const adminEmail = PropertiesService.getScriptProperties().getProperty(
    SCRIPT_PROPS_KEYS.ADMIN_EMAIL
  );

  ULog.debug('renderAdminPanel - isDeployUser() result:', deployUserResult);
  ULog.debug('renderAdminPanel - current user email:', currentUserEmail);
  ULog.debug('renderAdminPanel - ADMIN_EMAIL property:', adminEmail);
  ULog.debug('renderAdminPanel - emails match:', adminEmail === currentUserEmail);
  adminTemplate.isDeployUser = deployUserResult;
  adminTemplate.DEBUG_MODE = shouldEnableDebugMode();

  const htmlOutput = adminTemplate
    .evaluate()
    .setTitle('StudyQuest - 管理パネル')
    .setSandboxMode(HtmlService.SandboxMode.NATIVE);

  // XFrameOptionsMode を安全に設定（iframe embedding許可）
  try {
    if (HtmlService && HtmlService.XFrameOptionsMode && HtmlService.XFrameOptionsMode.ALLOWALL) {
      htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      ULog.debug('✅ Admin Panel XFrameOptionsMode.ALLOWALL設定完了 - iframe embedding許可');
    } else {
      ULog.warn('⚠️ HtmlService.XFrameOptionsMode.ALLOWALLが利用できません');
    }
  } catch (e) {
    ULog.error('[ERROR]', '❌ Admin Panel XFrameOptionsMode設定エラー:', e.message);
    // フォールバック: 従来の方法で設定を試行
    try {
      htmlOutput.setXFrameOptionsMode('ALLOWALL');
      ULog.info('💡 フォールバック方法でAdmin Panel XFrameOptionsMode設定完了');
    } catch (fallbackError) {
      ULog.error('[ERROR]', '❌ フォールバック方法も失敗:', fallbackError.message);
    }
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
    ULog.debug('🚫 renderUnpublishedPage: Rendering unpublished page for userId:', userInfo.userId);

    let template;
    try {
      template = HtmlService.createTemplateFromFile('Unpublished');
      ULog.debug('✅ renderUnpublishedPage: Template created successfully');
    } catch (templateError) {
      ULog.error('❌ renderUnpublishedPage: Template creation failed:', templateError);
      throw new Error('Unpublished.htmlテンプレートの読み込みに失敗: ' + templateError.message);
    }

    template.include = include;

    // 基本情報の設定（安全なデフォルト値付き）
    template.userId = userInfo.userId || '';
    template.spreadsheetId = userInfo.spreadsheetId || '';
    template.ownerName = userInfo.adminEmail || 'システム管理者';
    template.isOwner = getCurrentUserEmail() === userInfo.adminEmail; // 現在のユーザーがボードの所有者であるかを確認
    template.adminEmail = userInfo.adminEmail || '';
    template.cacheTimestamp = Date.now();

    // 安全な変数設定
    template.include = include;

    // URL生成（エラー耐性を持たせる）
    let appUrls;
    try {
      appUrls = generateUserUrls(userInfo.userId);
      if (!appUrls || !appUrls.success) {
        throw new Error('URL生成に失敗しました');
      }
    } catch (urlError) {
      ULog.warn('URL生成エラー、フォールバック値を使用:', urlError);
      // フォールバック: 基本的なURL構造
      const baseUrl = getWebAppUrl();
      appUrls = {
        adminUrl: `${baseUrl}?mode=admin&userId=${encodeURIComponent(userInfo.userId)}`,
        viewUrl: `${baseUrl}?mode=view&userId=${encodeURIComponent(userInfo.userId)}`,
        status: 'fallback',
      };
    }

    template.adminPanelUrl = appUrls.adminUrl || '';
    template.boardUrl = appUrls.viewUrl || '';

    ULog.debug('✅ renderUnpublishedPage: Template setup completed');

    // キャッシュを無効化して確実なリダイレクトを保証
    const htmlOutput = template.evaluate().setTitle('StudyQuest - 準備中');

    // addMetaTagを安全に追加
    try {
      htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
    } catch (e) {
      ULog.warn('⚠️ addMetaTag(viewport) failed:', e.message);
    }

    try {
      htmlOutput.addMetaTag('cache-control', 'no-cache, no-store, must-revalidate');
    } catch (e) {
      ULog.warn('⚠️ addMetaTag(cache-control) failed:', e.message);
    }

    try {
      if (HtmlService && HtmlService.XFrameOptionsMode && HtmlService.XFrameOptionsMode.ALLOWALL) {
        htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      }
    } catch (e) {
      ULog.warn('XFrameOptionsMode設定エラー:', e.message);
    }

    return htmlOutput;
  } catch (error) {
    logError(
      error,
      'renderUnpublishedPage',
      MAIN_MAIN_UNIFIED_CONSTANTS.ERROR.SEVERITY.HIGH,
      MAIN_MAIN_UNIFIED_CONSTANTS.ERROR.CATEGORIES.SYSTEM,
      {
        userId: userInfo ? userInfo.userId : 'null',
        hasUserInfo: !!userInfo,
        errorMessage: error.message,
        errorStack: error.stack,
      }
    );
    ULog.error('🚨 renderUnpublishedPage error details:', {
      error: error,
      userInfo: userInfo,
      userId: userInfo ? userInfo.userId : 'N/A',
      adminEmail: userInfo ? userInfo.adminEmail : 'N/A',
    });
    // フォールバック: ErrorBoundary.htmlを回避して確実にUnpublishedページを表示
    return renderMinimalUnpublishedPage(userInfo);
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
    ULog.debug('🚫 renderMinimalUnpublishedPage: Creating minimal unpublished page');

    // 安全にuserInfoを処理
    if (!userInfo) {
      ULog.warn('⚠️ renderMinimalUnpublishedPage: userInfo is null/undefined');
      userInfo = { userId: '', adminEmail: '' };
    }

    const userId = userInfo.userId && typeof userInfo.userId === 'string' ? userInfo.userId : '';
    const adminEmail =
      userInfo.adminEmail && typeof userInfo.adminEmail === 'string' ? userInfo.adminEmail : '';

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

    const htmlOutput = HtmlService.createHtmlOutput(htmlContent).setTitle('StudyQuest - 準備中');

    // addMetaTagを安全に追加
    try {
      htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
    } catch (e) {
      ULog.warn('⚠️ renderMinimalUnpublishedPage addMetaTag(viewport) failed:', e.message);
    }

    try {
      htmlOutput.addMetaTag('cache-control', 'no-cache, no-store, must-revalidate');
    } catch (e) {
      ULog.warn('⚠️ renderMinimalUnpublishedPage addMetaTag(cache-control) failed:', e.message);
    }

    return htmlOutput;
  } catch (error) {
    logError(
      error,
      'renderMinimalUnpublishedPage',
      MAIN_MAIN_UNIFIED_CONSTANTS.ERROR.SEVERITY.HIGH,
      MAIN_MAIN_UNIFIED_CONSTANTS.ERROR.CATEGORIES.SYSTEM,
      {
        userId: userInfo ? userInfo.userId : 'null',
        hasUserInfo: !!userInfo,
        errorMessage: error.message,
        errorStack: error.stack,
      }
    );
    ULog.error('🚨 renderMinimalUnpublishedPage error details:', {
      error: error,
      userInfo: userInfo,
      userId: userInfo ? userInfo.userId : 'N/A',
      adminEmail: userInfo ? userInfo.adminEmail : 'N/A',
    });
    // 最終フォールバック: 管理者向け機能付き
    const userId = userInfo && userInfo.userId ? userInfo.userId : '';
    const adminEmail = userInfo && userInfo.adminEmail ? userInfo.adminEmail : '';

    const finalFallbackHtml = `
      <!DOCTYPE html>
      <html lang="ja">
      <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1">
          <title>StudyQuest - 準備中</title>
          <style>
              * { margin: 0; padding: 0; box-sizing: border-box; }
              body {
                  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Noto Sans JP', sans-serif;
                  background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                  min-height: 100vh;
                  color: #ffffff;
                  line-height: 1.6;
                  overflow-x: hidden;
              }
              .background-overlay {
                  position: fixed;
                  top: 0;
                  left: 0;
                  width: 100%;
                  height: 100%;
                  background: 
                      radial-gradient(circle at 20% 80%, rgba(120, 119, 198, 0.3) 0%, transparent 50%),
                      radial-gradient(circle at 80% 20%, rgba(255, 183, 77, 0.3) 0%, transparent 50%),
                      radial-gradient(circle at 40% 40%, rgba(120, 119, 198, 0.2) 0%, transparent 50%);
                  z-index: -1;
              }
              .main-container {
                  min-height: 100vh;
                  display: flex;
                  align-items: center;
                  justify-content: center;
                  padding: 20px;
                  position: relative;
              }
              .content-card {
                  background: rgba(255, 255, 255, 0.95);
                  backdrop-filter: blur(20px);
                  border-radius: 24px;
                  box-shadow: 
                      0 25px 50px -12px rgba(0, 0, 0, 0.25),
                      0 0 0 1px rgba(255, 255, 255, 0.1);
                  max-width: 600px;
                  width: 100%;
                  padding: 40px;
                  text-align: center;
                  color: #374151;
                  position: relative;
                  overflow: hidden;
              }
              .content-card::before {
                  content: '';
                  position: absolute;
                  top: 0;
                  left: 0;
                  right: 0;
                  height: 4px;
                  background: linear-gradient(90deg, #667eea, #764ba2, #667eea);
                  background-size: 200% 100%;
                  animation: shimmer 3s ease-in-out infinite;
              }
              @keyframes shimmer {
                  0%, 100% { background-position: 200% 0; }
                  50% { background-position: -200% 0; }
              }
              .status-icon {
                  width: 80px;
                  height: 80px;
                  margin: 0 auto 24px;
                  background: linear-gradient(135deg, #fbbf24, #f59e0b);
                  border-radius: 50%;
                  display: flex;
                  align-items: center;
                  justify-content: center;
                  font-size: 36px;
                  animation: pulse 2s infinite;
                  box-shadow: 0 10px 30px rgba(251, 191, 36, 0.3);
              }
              @keyframes pulse {
                  0%, 100% { transform: scale(1); }
                  50% { transform: scale(1.05); }
              }
              .status-title {
                  font-size: 32px;
                  font-weight: 700;
                  color: #1f2937;
                  margin-bottom: 16px;
                  background: linear-gradient(135deg, #667eea, #764ba2);
                  -webkit-background-clip: text;
                  -webkit-text-fill-color: transparent;
                  background-clip: text;
              }
              .status-message {
                  font-size: 18px;
                  color: #6b7280;
                  margin-bottom: 40px;
                  line-height: 1.7;
              }
              .button-group {
                  display: flex;
                  flex-wrap: wrap;
                  gap: 16px;
                  justify-content: center;
                  margin-bottom: 32px;
              }
              .btn {
                  display: inline-flex;
                  align-items: center;
                  gap: 8px;
                  padding: 14px 28px;
                  border: none;
                  border-radius: 12px;
                  font-size: 16px;
                  font-weight: 600;
                  text-decoration: none;
                  cursor: pointer;
                  transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
                  position: relative;
                  overflow: hidden;
              }
              .btn::before {
                  content: '';
                  position: absolute;
                  top: 0;
                  left: -100%;
                  width: 100%;
                  height: 100%;
                  background: linear-gradient(90deg, transparent, rgba(255,255,255,0.2), transparent);
                  transition: left 0.5s;
              }
              .btn:hover::before {
                  left: 100%;
              }
              .btn-primary {
                  background: linear-gradient(135deg, #10b981, #059669);
                  color: white;
                  box-shadow: 0 4px 14px rgba(16, 185, 129, 0.3);
              }
              .btn-primary:hover {
                  transform: translateY(-2px);
                  box-shadow: 0 8px 25px rgba(16, 185, 129, 0.4);
              }
              .btn-secondary {
                  background: linear-gradient(135deg, #3b82f6, #2563eb);
                  color: white;
                  box-shadow: 0 4px 14px rgba(59, 130, 246, 0.3);
              }
              .btn-secondary:hover {
                  transform: translateY(-2px);
                  box-shadow: 0 8px 25px rgba(59, 130, 246, 0.4);
              }
              .btn-tertiary {
                  background: linear-gradient(135deg, #6b7280, #4b5563);
                  color: white;
                  box-shadow: 0 4px 14px rgba(107, 114, 128, 0.3);
              }
              .btn-tertiary:hover {
                  transform: translateY(-2px);
                  box-shadow: 0 8px 25px rgba(107, 114, 128, 0.4);
              }
              .info-section {
                  background: linear-gradient(135deg, rgba(99, 102, 241, 0.1), rgba(139, 92, 246, 0.1));
                  border: 1px solid rgba(99, 102, 241, 0.2);
                  border-radius: 16px;
                  padding: 24px;
                  margin-bottom: 24px;
              }
              .info-title {
                  font-size: 14px;
                  font-weight: 600;
                  color: #6366f1;
                  margin-bottom: 8px;
              }
              .info-detail {
                  font-size: 14px;
                  color: #6b7280;
              }
              .error-notice {
                  background: linear-gradient(135deg, rgba(239, 68, 68, 0.1), rgba(220, 38, 38, 0.1));
                  border: 1px solid rgba(239, 68, 68, 0.3);
                  border-radius: 16px;
                  padding: 20px;
                  margin-top: 24px;
              }
              .error-notice-title {
                  font-size: 16px;
                  font-weight: 600;
                  color: #dc2626;
                  margin-bottom: 8px;
                  display: flex;
                  align-items: center;
                  gap: 8px;
              }
              .error-notice-text {
                  font-size: 14px;
                  color: #7f1d1d;
                  line-height: 1.6;
              }
              @media (max-width: 640px) {
                  .content-card { padding: 24px; margin: 16px; }
                  .status-title { font-size: 24px; }
                  .status-message { font-size: 16px; }
                  .button-group { flex-direction: column; }
                  .btn { width: 100%; justify-content: center; }
              }
          </style>
      </head>
      <body>
          <div class="background-overlay"></div>
          <div class="main-container">
              <div class="content-card">
                  <div class="status-icon">⏳</div>
                  <h1 class="status-title">回答ボード準備中</h1>
                  <p class="status-message">
                      現在、回答ボードは非公開になっています。<br>
                      管理者として以下の操作が可能です。
                  </p>
                  
                  <div class="button-group">
                      <button onclick="republishBoard()" class="btn btn-primary">
                          🔄 回答ボードを再公開
                      </button>
                      <a href="?mode=admin&userId=${encodeURIComponent(userId)}" class="btn btn-secondary">
                          ⚙️ 管理パネルを開く
                      </a>
                      <button onclick="location.reload()" class="btn btn-tertiary">
                          🔄 ページを更新
                      </button>
                  </div>
                  
                  <div class="info-section">
                      <div class="info-title">管理者情報</div>
                      <div class="info-detail">
                          管理者: ${adminEmail || 'システム管理者'}<br>
                          ユーザーID: ${userId || '不明'}
                      </div>
                  </div>
                  
                  <div class="error-notice">
                      <div class="error-notice-title">
                          ⚠️ システム情報
                      </div>
                      <div class="error-notice-text">
                          テンプレートの読み込みでエラーが発生したため、基本機能のみ表示しています。すべての管理機能は正常に動作します。
                      </div>
                  </div>
              </div>
          </div>
          
          <script>
              function republishBoard() {
                  if (!confirm('回答ボードを再公開しますか？')) return;
                  
                  const button = event.target;
                  button.disabled = true;
                  button.textContent = '再公開中...';
                  
                  try {
                      google.script.run
                          .withSuccessHandler((result) => {
                              alert('再公開が完了しました！ページを更新します。');
                              setTimeout(() => {
                                  window.location.href = '?mode=view&userId=${encodeURIComponent(userId)}&_cb=' + Date.now();
                              }, 1000);
                          })
                          .withFailureHandler((error) => {
                              alert('再公開に失敗しました: ' + error.message);
                              button.disabled = false;
                              button.textContent = '🔄 回答ボードを再公開';
                          })
                          .republishBoard('${userId}');
                  } catch (error) {
                      alert('エラーが発生しました: ' + error.message);
                      button.disabled = false;
                      button.textContent = '🔄 回答ボードを再公開';
                  }
              }
          </script>
      </body>
      </html>
    `;

    const finalHtmlOutput =
      HtmlService.createHtmlOutput(finalFallbackHtml).setTitle('StudyQuest - 準備中');

    // addMetaTagを安全に追加
    try {
      finalHtmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
    } catch (e) {
      ULog.warn('⚠️ Final fallback addMetaTag(viewport) failed:', e.message);
    }

    try {
      finalHtmlOutput.addMetaTag('cache-control', 'no-cache, no-store, must-revalidate');
    } catch (e) {
      ULog.warn('⚠️ Final fallback addMetaTag(cache-control) failed:', e.message);
    }

    return finalHtmlOutput;
  }
}

function renderAnswerBoard(userInfo, params) {
  try {
    let config = {};
    try {
      config = JSON.parse(userInfo.configJson || '{}');
    } catch (e) {
      ULog.warn('Invalid configJson:', e.message);
    }
    // publishedSheetNameの型安全性確保（'true'問題の修正）
    let safePublishedSheetName = '';
    if (config.publishedSheetName) {
      if (typeof config.publishedSheetName === 'string') {
        safePublishedSheetName = config.publishedSheetName;
      } else {
        logValidationError(
          'publishedSheetName',
          config.publishedSheetName,
          'string_type',
          `不正な型: ${typeof config.publishedSheetName}`
        );
        ULog.warn('🔧 main.gs: publishedSheetNameを空文字にリセットしました');
        safePublishedSheetName = '';
      }
    }

    // 強化されたパブリケーション状態検証（キャッシュバスティング対応）
    const isPublished = !!(
      config.appPublished &&
      config.publishedSpreadsheetId &&
      safePublishedSheetName
    );

    // リアルタイム検証: 非公開状態の場合は確実に検出
    const isCurrentlyPublished =
      isPublished &&
      config.appPublished === true &&
      config.publishedSpreadsheetId &&
      safePublishedSheetName;

    const sheetConfigKey = 'sheet_' + (safePublishedSheetName || params.sheetName);
    const sheetConfig = config[sheetConfigKey] || {};

    // この関数は公開ボード専用（非公開判定は呼び出し前に完了）
    ULog.debug('✅ renderAnswerBoard: Rendering published board for userId:', userInfo.userId);

    const template = HtmlService.createTemplateFromFile('Page');
    template.include = include;

    try {
      if (userInfo.spreadsheetId) {
        try {
          addServiceAccountToSpreadsheet(userInfo.spreadsheetId);
        } catch (err) {
          ULog.warn('アクセス権設定警告:', err.message);
        }
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
    return template
      .evaluate()
      .setTitle('StudyQuest -みんなの回答ボード-')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (error) {
    logError(
      error,
      'renderAnswerBoard',
      MAIN_MAIN_UNIFIED_CONSTANTS.ERROR.SEVERITY.HIGH,
      MAIN_MAIN_UNIFIED_CONSTANTS.ERROR.CATEGORIES.SYSTEM,
      {
        userId: userInfo.userId,
        spreadsheetId: userInfo.spreadsheetId,
        sheetName: safePublishedSheetName || params.sheetName,
      }
    );
    // フォールバック: 基本的なエラーページ
    return showErrorPage(
      'エラー',
      'ボードの表示でエラーが発生しました。管理者にお問い合わせください。'
    );
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
    ULog.debug('🔍 checkCurrentPublicationStatus called for userId:', userId);

    if (!userId) {
      ULog.warn('userId is required for publication status check');
      return { error: 'userId is required', isPublished: false };
    }

    // ユーザー情報を強制的に最新状態で取得（キャッシュバイパス）
    const userInfo = findUserById(userId, {
      useExecutionCache: false,
      forceRefresh: true,
    });

    if (!userInfo) {
      ULog.warn('User not found for publication status check:', userId);
      return { error: 'User not found', isPublished: false };
    }

    // 設定情報を解析
    let config = {};
    try {
      config = JSON.parse(userInfo.configJson || '{}');
    } catch (e) {
      ULog.warn('Config JSON parse error during publication status check:', e.message);
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

    ULog.debug('📊 Publication status check result:', {
      userId: userId,
      appPublished: config.appPublished,
      hasSpreadsheetId: !!config.publishedSpreadsheetId,
      hasSheetName: !!config.publishedSheetName,
      isCurrentlyPublished: isCurrentlyPublished,
      timestamp: new Date().toISOString(),
    });

    return {
      isPublished: isCurrentlyPublished,
      publishedSheetName: config.publishedSheetName || null,
      publishedSpreadsheetId: config.publishedSpreadsheetId || null,
      lastChecked: new Date().toISOString(),
    };
  } catch (error) {
    logError(
      error,
      'checkCurrentPublicationStatus',
      MAIN_MAIN_UNIFIED_CONSTANTS.ERROR.SEVERITY.MEDIUM,
      MAIN_MAIN_UNIFIED_CONSTANTS.ERROR.CATEGORIES.SYSTEM
    );
    return {
      error: error.message,
      isPublished: false,
      lastChecked: new Date().toISOString(),
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

// =================================================================
// DEBUG_MODE & USER ACCESS CONTROL API
// =================================================================

/**
 * 現在のDEBUG_MODE状態を取得
 * @returns {Object} DEBUG_MODEの状態情報
 */
function getDebugModeStatus() {
  try {
    const debugMode = PropertiesService.getScriptProperties().getProperty('DEBUG_MODE') === 'true';

    return {
      success: true,
      debugMode: debugMode,
      message: debugMode ? 'デバッグモードが有効です' : 'デバッグモードが無効です',
      lastModified:
        PropertiesService.getScriptProperties().getProperty('DEBUG_MODE_LAST_MODIFIED') ||
        'unknown',
    };
  } catch (error) {
    ULog.error('[ERROR]', 'getDebugModeStatus error:', error.message);
    return {
      success: false,
      message: 'DEBUG_MODE状態の取得に失敗しました: ' + error.message,
    };
  }
}

/**
 * DEBUG_MODEの状態を切り替える（システム管理者のみ）
 * @param {boolean} enable - デバッグモードを有効にするかどうか
 * @returns {Object} 操作結果
 */
function toggleDebugMode(enable) {
  try {
    // システム管理者権限チェック
    if (!isSystemAdmin()) {
      throw new Error('システム管理者権限が必要です');
    }

    const props = PropertiesService.getScriptProperties();
    const newValue = enable ? 'true' : 'false';
    const currentValue = props.getProperty('DEBUG_MODE');

    // 変更があるかチェック
    if (currentValue === newValue) {
      return {
        success: true,
        debugMode: enable,
        message: `DEBUG_MODEは既に${enable ? '有効' : '無効'}です`,
        changed: false,
      };
    }

    // DEBUG_MODE設定を更新
    props.setProperties({
      DEBUG_MODE: newValue,
      DEBUG_MODE_LAST_MODIFIED: new Date().toISOString(),
    });

    ULog.info('DEBUG_MODE changed:', {
      from: currentValue || 'undefined',
      to: newValue,
      by: getCurrentUserEmail(),
    });

    return {
      success: true,
      debugMode: enable,
      message: `DEBUG_MODEを${enable ? '有効' : '無効'}にしました`,
      changed: true,
      timestamp: new Date().toISOString(),
    };
  } catch (error) {
    ULog.error('[ERROR]', 'toggleDebugMode error:', error.message);
    return {
      success: false,
      message: 'DEBUG_MODE切り替えに失敗しました: ' + error.message,
    };
  }
}

/**
 * 個別ユーザー自身のアクセス状態を取得
 * @returns {Object} 現在のアクセス状態
 */
function getUserActiveStatus() {
  try {
    const currentUser = findUserById(getUserId());
    if (!currentUser || !currentUser.userId) {
      throw new Error('ユーザー情報が取得できません');
    }

    // データベースから現在のユーザー情報を取得
    const userData = fetchUserFromDatabase(currentUser.userId);
    if (!userData) {
      throw new Error('ユーザーデータが見つかりません');
    }

    const isActive = userData.isActive !== false; // デフォルトはtrue

    return {
      success: true,
      isActive: isActive,
      userId: currentUser.userId,
      email: currentUser.adminEmail,
      timestamp: new Date().toISOString(),
    };
  } catch (error) {
    ULog.error('[ERROR]', 'getUserActiveStatus error:', error.message);
    return {
      success: false,
      error: 'アクセス状態の取得に失敗しました: ' + error.message,
      isActive: true, // フォールバック
    };
  }
}

/**
 * 個別ユーザー自身のアクセス状態を更新
 * @param {string} targetUserId - 対象ユーザーのID（現在のユーザーと一致する必要がある）
 * @param {boolean} isActive - 新しいアクティブ状態
 * @returns {Object} 操作結果
 */
function updateSelfActiveStatus(targetUserId, isActive) {
  return updateUserActiveStatusCore(targetUserId, isActive, {
    callerType: 'self',
    clearCache: true,
  });
}

/**
 * 個別ユーザーのisActiveステータスを更新（管理者用）
 * @param {string} userId - 対象ユーザーのID
 * @param {boolean} isActive - 新しいアクティブ状態
 * @returns {Object} 操作結果
 */
function updateUserActiveStatus(userId, isActive) {
  return updateUserActiveStatusCore(userId, isActive, {
    callerType: 'admin',
    clearCache: true,
  });
}

/**
 * アクティブステータス更新のコア統一関数
 * @param {string} targetUserId - 対象ユーザーID
 * @param {boolean} isActive - 新しいアクティブ状態
 * @param {Object} options - オプション設定
 * @param {string} options.callerType - 呼び出し元タイプ ('self', 'admin', 'api', 'ui')
 * @param {boolean} options.skipPermissionCheck - 権限チェックスキップ
 * @param {boolean} options.clearCache - キャッシュクリア実行
 * @returns {Object} 操作結果
 */
function updateUserActiveStatusCore(targetUserId, isActive, options = {}) {
  try {
    const { callerType = 'api', skipPermissionCheck = false, clearCache = true } = options;

    // 入力値検証
    if (!targetUserId || typeof targetUserId !== 'string') {
      throw new Error('有効なユーザーIDが必要です');
    }

    if (typeof isActive !== 'boolean') {
      throw new Error('isActiveは真偽値である必要があります');
    }

    // 対象ユーザー情報を取得
    const userInfo = findUserById(targetUserId);
    if (!userInfo) {
      throw new Error('指定されたユーザーが見つかりません');
    }

    // 権限チェック（スキップ指定がない場合）
    if (!skipPermissionCheck) {
      switch (callerType) {
        case 'self':
          const currentUser = findUserById(getUserId());
          if (!currentUser || currentUser.userId !== targetUserId) {
            throw new Error('自分自身のアクセス状態のみ変更できます');
          }
          break;

        case 'admin':
        case 'ui':
          if (!isSystemAdmin()) {
            throw new Error('システム管理者権限が必要です');
          }
          break;

        case 'api':
          const activeUserEmail = getCurrentUserEmail();
          if (!activeUserEmail) {
            throw new Error('ユーザーが認証されていません');
          }
          const apiUser = findUserByEmail(activeUserEmail);
          if (!apiUser || !isTrue(apiUser.isActive)) {
            throw new Error('この操作を実行する権限がありません');
          }
          break;

        default:
          throw new Error('不正な呼び出し元タイプです');
      }
    }

    // 現在の状態と比較
    const currentActive = Boolean(userInfo.isActive);
    if (currentActive === isActive) {
      return {
        success: true,
        userId: targetUserId,
        email: userInfo.adminEmail,
        isActive: isActive,
        message: `ユーザー ${userInfo.adminEmail || targetUserId} は既に${isActive ? 'アクティブ' : '非アクティブ'}です`,
        changed: false,
        timestamp: new Date().toISOString(),
      };
    }

    // データベース更新（統一化）
    const updateResult = updateUserDatabaseField(targetUserId, 'isActive', isActive);

    if (!updateResult.success) {
      throw new Error(updateResult.error || 'データベース更新に失敗しました');
    }

    // キャッシュクリア
    if (clearCache) {
      clearUserCache(targetUserId);
    }

    // ログ出力
    const logData = {
      userId: targetUserId,
      email: userInfo.adminEmail,
      from: currentActive,
      to: isActive,
      callerType: callerType,
      by: getCurrentUserEmail() || 'system',
    };

    ULog.info('User active status changed:', logData);

    return {
      success: true,
      userId: targetUserId,
      email: userInfo.adminEmail,
      isActive: isActive,
      message: `ユーザー ${userInfo.adminEmail || targetUserId} を${isActive ? 'アクティブ' : '非アクティブ'}にしました`,
      changed: true,
      timestamp: new Date().toISOString(),
    };
  } catch (error) {
    ULog.error('[ERROR]', 'updateUserActiveStatusCore error:', error.message);
    return {
      success: false,
      message: 'アクティブステータス更新に失敗しました: ' + error.message,
      error: error.message,
    };
  }
}

/**
 * 複数ユーザーのisActiveステータスを一括更新
 * @param {string[]} userIds - 対象ユーザーIDの配列
 * @param {boolean} isActive - 新しいアクティブ状態
 * @returns {Object} 操作結果
 */
function bulkUpdateUserActiveStatus(userIds, isActive) {
  try {
    // システム管理者権限チェック
    if (!isSystemAdmin()) {
      throw new Error('システム管理者権限が必要です');
    }

    if (!Array.isArray(userIds) || userIds.length === 0) {
      throw new Error('有効なユーザーID配列が必要です');
    }

    const results = [];
    let successCount = 0;
    let errorCount = 0;

    // 各ユーザーを個別に更新
    for (const userId of userIds) {
      try {
        const result = updateUserActiveStatus(userId, isActive);
        results.push(result);
        if (result.success) {
          successCount++;
        } else {
          errorCount++;
        }
      } catch (error) {
        results.push({
          success: false,
          userId: userId,
          message: error.message,
        });
        errorCount++;
      }
    }

    ULog.info('Bulk user active status update:', {
      totalUsers: userIds.length,
      successCount: successCount,
      errorCount: errorCount,
      isActive: isActive,
      by: getCurrentUserEmail(),
    });

    return {
      status: errorCount === 0 ? 'success' : 'partial',
      results: results,
      summary: {
        total: userIds.length,
        success: successCount,
        errors: errorCount,
      },
      message: `${successCount}人のユーザーを${isActive ? 'アクティブ' : '非アクティブ'}にしました${errorCount > 0 ? ` (${errorCount}件のエラー)` : ''}`,
      timestamp: new Date().toISOString(),
    };
  } catch (error) {
    return UError(
      error,
      'bulkUpdateUserActiveStatus',
      UNIFIED_CONSTANTS.ERROR.SEVERITY.HIGH,
      UNIFIED_CONSTANTS.ERROR.CATEGORIES.SYSTEM
    );
  }
}

/**
 * 全ユーザーのisActiveステータスを一括更新
 * @param {boolean} isActive - 新しいアクティブ状態
 * @returns {Object} 操作結果
 */
function bulkUpdateAllUsersActiveStatus(isActive) {
  try {
    // システム管理者権限チェック
    if (!isSystemAdmin()) {
      throw new Error('システム管理者権限が必要です');
    }

    // 全ユーザーリストを取得
    const allUsers = getAllUsers();
    if (!allUsers || allUsers.length === 0) {
      return {
        success: true,
        message: '更新対象のユーザーがいません',
        summary: { total: 0, success: 0, errors: 0 },
      };
    }

    // 全ユーザーのIDを抽出
    const userIds = allUsers.map((user) => user.userId).filter((id) => id);

    // 一括更新を実行
    return bulkUpdateUserActiveStatus(userIds, isActive);
  } catch (error) {
    return UError(
      error,
      'bulkUpdateAllUsersActiveStatus',
      UNIFIED_CONSTANTS.ERROR.SEVERITY.HIGH,
      UNIFIED_CONSTANTS.ERROR.CATEGORIES.SYSTEM
    );
  }
}

/**
 * 【Phase 8整合性修正】フロントエンド互換性のためのclearCache関数
 * フロントエンドで使用されているclearCache呼び出しに対応
 */
function clearCache() {
  try {
    UUtilities.logger.info('Cache', 'フロントエンドからのキャッシュクリア要求');
    
    // PropertiesService キャッシュのクリア
    PropertiesService.getScriptProperties().deleteAll();
    
    // CacheService キャッシュのクリア (キャッシュは自動的に期限切れになります)
    try {
      // 明示的なキャッシュクリアは不要（キャッシュは10分で自動期限切れ）
      console.log('[Cache] キャッシュは自動的に期限切れになります');
    } catch (error) {
      console.log('[Cache] キャッシュクリア処理をスキップ:', error.message);
    }
    
    UUtilities.logger.info('Cache', 'すべてのキャッシュをクリアしました');
    return UUtilities.generatorFactory.response.success(null, 'キャッシュを正常にクリアしました');
    
  } catch (error) {
    UUtilities.logger.error('Cache', 'キャッシュクリア中にエラーが発生', error.message);
    return UUtilities.generatorFactory.response.error(error.message, 'キャッシュクリアに失敗しました');
  }
}

/**
 * 【Phase 8整合性修正】クライアントエラー報告
 * ErrorBoundaryから呼び出されるエラー報告機能
 */
function reportClientError(errorInfo) {
  try {
    UUtilities.logger.error('Client', 'フロントエンドエラー報告', {
      message: errorInfo.message || 'Unknown error',
      stack: errorInfo.stack || 'No stack trace',
      componentStack: errorInfo.componentStack || 'No component stack',
      timestamp: new Date().toISOString(),
      userAgent: errorInfo.userAgent || 'Unknown'
    });
    
    // エラーをコンソールに記録
    console.error('[CLIENT ERROR]', JSON.stringify(errorInfo, null, 2));
    
    return UUtilities.generatorFactory.response.success(null, 'クライアントエラーを記録しました');
    
  } catch (error) {
    UUtilities.logger.error('Client', 'クライアントエラー報告中にエラーが発生', error.message);
    return UUtilities.generatorFactory.response.error(error.message, 'エラー報告に失敗しました');
  }
}
