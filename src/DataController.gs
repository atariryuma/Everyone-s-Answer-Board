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

/* global UserService, ConfigService, DataService, DB */

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
/**
 * DataController専用：直接Session APIでメール取得
 */
function getCurrentEmailDirectDC() {
  try {
    // Method 1: Session.getActiveUser()
    let email = Session.getActiveUser().getEmail();
    if (email) {
      return email;
    }

    // Method 2: Session.getEffectiveUser()
    email = Session.getEffectiveUser().getEmail();
    if (email) {
      return email;
    }

    console.warn('getCurrentEmailDirectDC: No email available from Session API');
    return null;
  } catch (error) {
    console.error('getCurrentEmailDirectDC:', error.message);
    return null;
  }
}

function handleGetData(request) {
  try {
    // 🎯 Zero-dependency: 直接Session APIでユーザー取得
    const email = getCurrentEmailDirectDC();
    if (!email) {
      return {
        success: false,
        message: 'ユーザー情報が見つかりません'
      };
    }

    // 🎯 Zero-dependency: 直接DBからユーザー情報取得
    const user = DB.findUserByEmail(email);
    if (!user) {
      return {
        success: false,
        message: 'ユーザーが登録されていません'
      };
    }

    // DataService.getSheetDataは既にzero-dependency実装済み
    const data = DataService.getSheetData(user.userId, request.options || {});
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
 * リアクションの追加
 * Page.html から呼び出される
 *
 * @param {Object} request - リアクション情報
 * @returns {Object} 追加結果
 */
function handleAddReaction(request) {
  try {
    // 🎯 Zero-dependency: 直接Session APIでユーザー取得
    const email = getCurrentEmailDirectDC();
    if (!email) {
      return {
        success: false,
        message: 'ユーザー情報が見つかりません'
      };
    }

    // 🎯 Zero-dependency: 直接DBからユーザー情報取得
    const user = DB.findUserByEmail(email);
    if (!user) {
      return {
        success: false,
        message: 'ユーザーが登録されていません'
      };
    }

    // DataService.addReactionは既にzero-dependency実装済み
    const result = DataService.addReaction(
      user.userId,
      request.rowId,
      request.reactionType
    );

    return result;

  } catch (error) {
    console.error('DataController.handleAddReaction エラー:', error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * ハイライトの切り替え
 * Page.html から呼び出される
 *
 * @param {Object} request - ハイライト情報
 * @returns {Object} 切り替え結果
 */
function handleToggleHighlight(request) {
  try {
    const userInfo = UserService.getCurrentUserInfo();
    if (!userInfo) {
      return {
        success: false,
        message: 'ユーザー情報が見つかりません'
      };
    }

    const result = DataService.toggleHighlight(
      userInfo.userId,
      request.rowId
    );

    return result;

  } catch (error) {
    console.error('DataController.handleToggleHighlight エラー:', error.message);
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
    const userInfo = UserService.getCurrentUserInfo();
    if (!userInfo) {
      return {
        success: false,
        message: 'ユーザー情報が見つかりません'
      };
    }

    const result = DataService.refreshBoardData(userInfo.userId, request.options || {});
    return result;

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

/**
 * 公開されたシートデータの取得
 * 外部からのアクセス用
 *
 * @param {string} userId - ユーザーID
 * @param {Object} options - 取得オプション
 * @returns {Object} 公開データ
 */
function getPublishedSheetData(userId, options = {}) {
  try {
    if (!userId) {
      return {
        success: false,
        message: 'ユーザーIDが指定されていません'
      };
    }

    // セキュリティチェック: 公開されているかの確認
    const config = ConfigService.getUserConfig(userId);
    if (!config || !config.appPublished) {
      return {
        success: false,
        message: 'このデータは公開されていません'
      };
    }

    const data = DataService.getSheetData(userId, options);
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

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
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
function getAllUsersForAdminForUI(_options = {}) {
  try {
    // 管理者権限チェック
    const currentEmail = UserService.getCurrentEmail();
    if (!UserService.isSystemAdmin(currentEmail)) {
      return {
        success: false,
        message: '管理者権限が必要です'
      };
    }

    // データベースから全ユーザーを取得
    const users = DB.getAllUsers();
    const processedUsers = users.map(user => ({
      userId: user.userId,
      userEmail: user.userEmail,
      createdAt: user.createdAt,
      lastLogin: user.lastLogin,
      hasConfig: !!user.configJSON
    }));

    return {
      success: true,
      users: processedUsers,
      total: processedUsers.length,
      timestamp: new Date().toISOString()
    };

  } catch (e) {
    console.error('DataController.getAllUsersForAdminForUI エラー:', e.message);
    return {
      success: false,
      message: e.message
    };
  }
}

/**
 * ユーザーアカウント削除（管理者向け）
 * AdminPanel.js.html から呼び出される
 *
 * @param {string} targetUserId - 削除対象のユーザーID
 * @returns {Object} 削除結果
 */
function deleteUserAccountByAdminForUI(targetUserId) {
  try {
    // 管理者権限チェック
    const currentEmail = UserService.getCurrentEmail();
    if (!UserService.isSystemAdmin(currentEmail)) {
      return {
        success: false,
        message: '管理者権限が必要です'
      };
    }

    if (!targetUserId) {
      return {
        success: false,
        message: 'ユーザーIDが指定されていません'
      };
    }

    // 対象ユーザーの存在確認
    const targetUser = DB.findUserById(targetUserId);
    if (!targetUser) {
      return {
        success: false,
        message: 'ユーザーが見つかりません'
      };
    }

    // 削除実行
    const result = DB.deleteUser(targetUserId);

    if (result.success) {
      console.log('管理者によるユーザー削除:', {
        admin: currentEmail,
        deletedUser: targetUser.userEmail,
        deletedUserId: targetUserId,
        timestamp: new Date().toISOString()
      });

      return {
        success: true,
        message: 'ユーザーアカウントを削除しました',
        deletedUser: {
          userId: targetUserId,
          email: targetUser.userEmail
        }
      };
    } else {
      return {
        success: false,
        message: result.message || '削除に失敗しました'
      };
    }

  } catch (error) {
    console.error('DataController.deleteUserAccountByAdminForUI エラー:', error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

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
      const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
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
        message: `スプレッドシートにアクセスできません: ${accessError.message}`
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
 * リアクション追加（API Gateway互換）
 */
function addDataReaction(userId, rowId, reactionType) {
  return handleAddReaction({
    userId,
    rowId,
    reactionType
  });
}

/**
 * ハイライト切り替え（API Gateway互換）
 */
function toggleDataHighlight(userId, rowId) {
  return handleToggleHighlight({
    userId,
    rowId
  });
}

/**
 * ボードデータ更新（API Gateway互換）
 */
function refreshBoardData(userId, options = {}) {
  return handleRefreshData({
    userId,
    options
  });
}

// ===========================================
// 📊 GAS Best Practices - Flat Function Structure
// ===========================================

/**
 * Converted from Object.freeze format to flat functions following GAS best practices
 * - Removed Object.freeze wrapper
 * - Converted methods to function declarations
 * - Fixed internal this.handleXxx() references to direct function calls
 * - Maintained all existing functionality and comments
 * - Added function name prefixes to avoid conflicts (e.g., addDataReaction)
 */
