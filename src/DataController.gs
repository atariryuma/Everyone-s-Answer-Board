/**
 * @fileoverview DataController - データ操作・表示関数の集約
 *
 * 🎯 責任範囲:
 * - メインページのデータ取得・表示
 * - リアクション・ハイライト操作
 * - データの公開・取得
 * - ユーザー管理機能
 *
 * 📝 main.gsから移動されたデータ操作関数群
 */

/* global ConfigService, DataService, getCurrentEmail, createErrorResponse, dsGetUserSheetData, findUserByEmail, findUserById, openSpreadsheet, updateUser, getUserSpreadsheetData, getUserConfig */

// ===========================================
// 📊 メインページデータAPI
// ===========================================

/**
 * メインページ用データの取得
 * Page.html から呼び出される
 *
 * @param {Object} request - リクエストパラメータ
 * @returns {Object} データ取得結果
 */

function handleGetData(request) {
  try {
    // 🎯 Zero-dependency: 直接Session APIでユーザー取得
    const email = getCurrentEmail();
    if (!email) {
      return {
        success: false,
        message: 'ユーザー情報が見つかりません'
      };
    }

    // 🔧 Zero-Dependency統一: 直接findUserByEmail使用
    const user = findUserByEmail(email);
    if (!user) {
      return {
        success: false,
        message: 'ユーザーが登録されていません'
      };
    }

    // 直接DataServiceに依存せず、安定APIで取得
    const data = dsGetUserSheetData(user.userId, request.options || {});
    return {
      success: true,
      data
    };

  } catch (error) {
    console.error('DataController.handleGetData エラー:', error.message);
    return {
      success: false,
      message: error.message
    };
  }
}


/**
 * データの更新
 * Page.html から呼び出される
 *
 * @param {Object} request - 更新リクエスト
 * @returns {Object} 更新結果
 */
function handleRefreshData(request) {
  try {
    // 🎯 Zero-dependency: 直接Session APIでユーザー取得
    const email = getCurrentEmail();
    if (!email) {
      return {
        success: false,
        message: 'ユーザー情報が見つかりません'
      };
    }

    // 🔧 Zero-Dependency統一: 直接findUserByEmail使用
    const user = findUserByEmail(email);
    if (!user) {
      return {
        success: false,
        message: 'ユーザーが登録されていません'
      };
    }

    // DataServiceに依存せず、直に最新データを再取得して返却
    const data = dsGetUserSheetData(user.userId, request.options || {});
    return { success: true, data };

  } catch (error) {
    console.error('DataController.handleRefreshData エラー:', error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

// ===========================================
// 📊 データ公開・取得API
// ===========================================

function getRecentSubmissions(userId, limit = 10) {
  try {
    if (!userId) {
      return {
        success: false,
        message: 'ユーザーIDが指定されていません'
      };
    }

    // GAS-Native: 直接Session APIでユーザー取得
    const email = Session.getActiveUser().getEmail();
    if (!email) {
      return {
        success: false,
        message: 'セッションが無効です'
      };
    }

    // 🔧 統一API使用: 直接getUserConfigを使用してConfigService依存を除去
    const configResult = getUserConfig(userId);
    const config = configResult.success ? configResult.config : null;
    if (!config || !config.isPublished) {
      return {
        success: false,
        message: 'このデータは公開されていません'
      };
    }

    const data = dsGetUserSheetData(userId, { limit, includeTimestamp: true });
    return {
      success: true,
      data,
      isPublic: true
    };

  } catch (error) {
    console.error('DataController.getPublishedSheetData エラー:', error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * ヘッダーインデックスの取得
 * 管理者向け機能
 *
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} ヘッダー情報
 */
function getHeaderIndices(spreadsheetId, sheetName) {
  try {
    if (!spreadsheetId || !sheetName) {
      return {
        success: false,
        message: 'スプレッドシートIDとシート名が必要です'
      };
    }

    const dataAccess = openSpreadsheet(spreadsheetId);
    const {spreadsheet} = dataAccess;
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      return {
        success: false,
        message: 'シートが見つかりません'
      };
    }

    const [headers] = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
    const indices = {};

    headers.forEach((header, index) => {
      if (header) {
        indices[header] = index;
      }
    });

    return {
      success: true,
      headers,
      indices,
      totalColumns: sheet.getLastColumn()
    };

  } catch (error) {
    console.error('DataController.getHeaderIndices エラー:', error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

// ===========================================
// 📊 ユーザー管理API（管理者向け）
// ===========================================

/**
 * 全ユーザー一覧の取得（管理者向け）
 * AdminPanel.js.html から呼び出される
 *
 * @param {Object} options - 取得オプション
 * @returns {Object} ユーザー一覧
 */


/**
 * スプレッドシートURL追加
 * 管理者向け機能
 *
 * @param {string} url - スプレッドシートURL
 * @returns {Object} 追加結果
 */
function addSpreadsheetUrl(url) {
  try {
    if (!url) {
      return {
        success: false,
        message: 'URLが指定されていません'
      };
    }

    // URL形式の簡易検証
    if (!url.includes('docs.google.com/spreadsheets')) {
      return {
        success: false,
        message: 'Googleスプレッドシートの有効なURLではありません'
      };
    }

    // スプレッドシートIDを抽出
    const match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (!match) {
      return {
        success: false,
        message: 'スプレッドシートIDが取得できません'
      };
    }

    const [, spreadsheetId] = match;

    // アクセステスト
    try {
      const dataAccess = openSpreadsheet(spreadsheetId);
    const {spreadsheet} = dataAccess;
      const name = spreadsheet.getName();

      return {
        success: true,
        message: 'スプレッドシートを確認しました',
        spreadsheet: {
          id: spreadsheetId,
          name,
          url
        }
      };
    } catch (accessError) {
      return {
        success: false,
        message: accessError && accessError.message ? `スプレッドシートにアクセスできません: ${accessError.message}` : 'スプレッドシートにアクセスできません: 詳細不明'
      };
    }

  } catch (error) {
    console.error('DataController.addSpreadsheetUrl エラー:', error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

// ===========================================
// 📊 API Gateway互換関数（main.gsから呼び出し用）
// ===========================================





/**
 * ボードデータ更新（API Gateway互換）
 */
function refreshBoardData(userId, options = {}) {
  try {
    if (!userId) {
      return { success: false, message: 'ユーザーIDが必要です' };
    }
    const data = dsGetUserSheetData(userId, options || {});
    return { success: true, data };
  } catch (err) {
    return { success: false, message: err.message || '更新エラー' };
  }
}

// ===========================================
// 📊 GAS Best Practices - Flat Function Structure
// ===========================================

/**
 * Converted from Object.freeze format to flat functions following GAS best practices
 * - Converted methods to function declarations
 * - Fixed internal this.handleXxx() references to direct function calls
 * - Maintained all existing functionality and comments
 * - Added function name prefixes to avoid conflicts (e.g., addDataReaction)
 */
