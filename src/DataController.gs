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

/* global ConfigService, DataService, getCurrentEmail, createErrorResponse, getUserSheetData, findUserByEmail, findUserById, findUserBySpreadsheetId, openSpreadsheet, updateUser, getUserSpreadsheetData, getUserConfig */

// ===========================================
// 📊 メインページデータAPI
// ===========================================


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

    // GAS-Native: getCurrentEmail関数でユーザー取得（統一パターン）
    const email = getCurrentEmail();
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

    const data = getUserSheetData(userId, { limit, includeTimestamp: true });
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

    // 🔧 CLAUDE.md準拠: Context-aware service account usage
    // ✅ **Cross-user**: Only use service account for accessing other user's spreadsheets
    // ✅ **Self-access**: Use normal permissions for own spreadsheets
    const currentEmail = getCurrentEmail();

    // CLAUDE.md準拠: spreadsheetIdから所有者を特定して直接比較
    const targetUser = findUserBySpreadsheetId(spreadsheetId);
    const isSelfAccess = targetUser && targetUser.userEmail === currentEmail;
    const useServiceAccount = !isSelfAccess;

    console.log(`getHeaderIndices: ${useServiceAccount ? 'Cross-user service account' : 'Self-access normal permissions'} for spreadsheet`);
    const dataAccess = openSpreadsheet(spreadsheetId, { useServiceAccount });
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
      // 🔧 CLAUDE.md準拠: Context-aware service account usage for URL verification
      // ✅ **Cross-user**: Only use service account for accessing other user's spreadsheets
      // ✅ **Self-access**: Use normal permissions for own spreadsheets
      const currentEmail = getCurrentEmail();

      // CLAUDE.md準拠: spreadsheetIdから所有者を特定して直接比較
      const targetUser = findUserBySpreadsheetId(spreadsheetId);
      const isSelfAccess = targetUser && targetUser.userEmail === currentEmail;
      const useServiceAccount = !isSelfAccess;

      console.log(`verifySpreadsheetUrl: ${useServiceAccount ? 'Cross-user service account' : 'Self-access normal permissions'} for spreadsheet verification`);
      const dataAccess = openSpreadsheet(spreadsheetId, { useServiceAccount });
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
