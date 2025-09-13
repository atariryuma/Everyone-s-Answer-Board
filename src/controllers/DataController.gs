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

/* global UserService, ConfigService, DataService, SecurityService, DB, SpreadsheetApp, ScriptApp */

/**
 * DataController - データ操作用コントローラー
 * メインページとデータ管理機能を集約
 */
const DataController = Object.freeze({

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
  handleGetData(request) {
    try {
      const userInfo = UserService.getCurrentUserInfo();
      if (!userInfo) {
        return {
          success: false,
          message: 'ユーザー情報が見つかりません'
        };
      }

      const data = DataService.getSheetData(userInfo.userId, request.options || {});
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
  },

  /**
   * リアクションの追加
   * Page.html から呼び出される
   *
   * @param {Object} request - リアクション情報
   * @returns {Object} 追加結果
   */
  handleAddReaction(request) {
    try {
      const userInfo = UserService.getCurrentUserInfo();
      if (!userInfo) {
        return {
          success: false,
          message: 'ユーザー情報が見つかりません'
        };
      }

      const result = DataService.addReaction(
        userInfo.userId,
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
  },

  /**
   * ハイライトの切り替え
   * Page.html から呼び出される
   *
   * @param {Object} request - ハイライト情報
   * @returns {Object} 切り替え結果
   */
  handleToggleHighlight(request) {
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
  },

  /**
   * データの更新
   * Page.html から呼び出される
   *
   * @param {Object} request - 更新リクエスト
   * @returns {Object} 更新結果
   */
  handleRefreshData(request) {
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
  },

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
  getPublishedSheetData(userId, options = {}) {
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
  },

  /**
   * ヘッダーインデックスの取得
   * 管理者向け機能
   *
   * @param {string} spreadsheetId - スプレッドシートID
   * @param {string} sheetName - シート名
   * @returns {Object} ヘッダー情報
   */
  getHeaderIndices(spreadsheetId, sheetName) {
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

      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
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
  },

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
  getAllUsersForAdminForUI(options = {}) {
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
  },

  /**
   * ユーザーアカウント削除（管理者向け）
   * AdminPanel.js.html から呼び出される
   *
   * @param {string} targetUserId - 削除対象のユーザーID
   * @returns {Object} 削除結果
   */
  deleteUserAccountByAdminForUI(targetUserId) {
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
  },

  /**
   * スプレッドシートURL追加
   * 管理者向け機能
   *
   * @param {string} url - スプレッドシートURL
   * @returns {Object} 追加結果
   */
  addSpreadsheetUrl(url) {
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

      const spreadsheetId = match[1];

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

});

// ===========================================
// 📊 グローバル関数エクスポート（GAS互換性のため）
// ===========================================

/**
 * データ操作用API関数を個別にエクスポート
 * HTMLファイルからの google.script.run 呼び出しに対応
 */

function handleGetData(request) {
  return DataController.handleGetData(request);
}

function handleAddReaction(request) {
  return DataController.handleAddReaction(request);
}

function handleToggleHighlight(request) {
  return DataController.handleToggleHighlight(request);
}

function handleRefreshData(request) {
  return DataController.handleRefreshData(request);
}

function getPublishedSheetData(userId, options = {}) {
  return DataController.getPublishedSheetData(userId, options);
}

function getHeaderIndices(spreadsheetId, sheetName) {
  return DataController.getHeaderIndices(spreadsheetId, sheetName);
}

function getAllUsersForAdminForUI(options = {}) {
  return DataController.getAllUsersForAdminForUI(options);
}

function deleteUserAccountByAdminForUI(targetUserId) {
  return DataController.deleteUserAccountByAdminForUI(targetUserId);
}

function addSpreadsheetUrl(url) {
  return DataController.addSpreadsheetUrl(url);
}