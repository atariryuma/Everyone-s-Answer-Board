/**
 * @fileoverview AdminController - 管理パネル用API関数の集約
 *
 * 🎯 責任範囲:
 * - 管理パネル固有のAPI関数
 * - フロントエンドからの呼び出し処理
 * - サービス層への委譲
 *
 * 📝 main.gsから移動されたAPI関数群
 */

/* global UserService, ConfigService, DataService, SecurityService, DB */

/**
 * AdminController - 管理パネル用コントローラー
 * フロントエンド（AdminPanel.js.html）からの呼び出しを処理
 */
const AdminController = Object.freeze({

  // ===========================================
  // 📊 設定管理API
  // ===========================================

  /**
   * 現在の設定を取得
   * AdminPanel.js.html から呼び出される
   *
   * @returns {Object} 設定情報
   */
  getConfig() {
    try {
      // ユーザー情報の取得
      let userInfo = UserService.getCurrentUserInfo();
      let userId = userInfo && userInfo.userId;

      if (!userId) {
        const email = UserService.getCurrentEmail();
        if (!email) {
          return { success: false, message: 'ユーザー情報が見つかりません' };
        }

        // DB から直接ユーザー検索を試行
        try {
          const dbUser = DB.findUserByEmail(email);
          if (dbUser && dbUser.userId) {
            userId = dbUser.userId;
            userInfo = userInfo || { userId, userEmail: email };
          }
        } catch (e) {
          // ignore and fallback to auto-create
        }

        // Auto-create user if still missing
        if (!userId) {
          try {
            const created = UserService.createUser(email);
            userId = created && created.userId;
            userInfo = created || { userId, userEmail: email };
          } catch (createErr) {
            return { success: false, message: 'ユーザー作成に失敗しました' };
          }
        }
      }

      const config = ConfigService.getUserConfig(userId);
      return {
        success: true,
        config: config || {}
      };
    } catch (error) {
      console.error('AdminController.getConfig エラー:', error.message);
      return {
        success: false,
        message: error.message
      };
    }
  },

  // ===========================================
  // 📊 データソース管理API（DataServiceに委譲）
  // ===========================================

  /**
   * スプレッドシート一覧を取得
   * @returns {Object} スプレッドシート一覧
   */
  getSpreadsheetList() {
    return DataService.getSpreadsheetList();
  },

  /**
   * シート一覧を取得
   * @param {string} spreadsheetId - スプレッドシートID
   * @returns {Object} シート一覧
   */
  getSheetList(spreadsheetId) {
    return DataService.getSheetList(spreadsheetId);
  },

  /**
   * 列を分析
   * @param {string} spreadsheetId - スプレッドシートID
   * @param {string} sheetName - シート名
   * @returns {Object} 列分析結果
   */
  analyzeColumns(spreadsheetId, sheetName) {
    return DataService.analyzeColumns(spreadsheetId, sheetName);
  },

  // ===========================================
  // 📊 設定・公開管理API（ConfigServiceに委譲）
  // ===========================================

  /**
   * 設定の下書き保存
   * @param {Object} config - 保存する設定
   * @returns {Object} 保存結果
   */
  saveDraftConfiguration(config) {
    return ConfigService.saveDraftConfiguration(config);
  },

  /**
   * アプリケーションの公開
   * @param {Object} publishConfig - 公開設定
   * @returns {Object} 公開結果
   */
  publishApplication(publishConfig) {
    return ConfigService.publishApplication(publishConfig);
  },

  // ===========================================
  // 📊 ユーティリティAPI
  // ===========================================

  /**
   * スプレッドシートへのアクセス権限を検証
   * AdminPanel.js.html から呼び出される
   *
   * @param {string} spreadsheetId - スプレッドシートID
   * @returns {Object} 検証結果
   */
  validateAccess(spreadsheetId) {
    try {
      if (!spreadsheetId) {
        return {
          success: false,
          message: 'スプレッドシートIDが指定されていません'
        };
      }

      // アクセステスト
      const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      const name = spreadsheet.getName();
      const sheets = spreadsheet.getSheets().map(sheet => ({
        name: sheet.getName(),
        index: sheet.getIndex()
      }));

      return {
        success: true,
        name,
        sheets,
        message: 'アクセス権限が確認できました'
      };

    } catch (error) {
      console.error('AdminController.validateAccess エラー:', error.message);

      let userMessage = 'スプレッドシートにアクセスできません。';
      if (error.message.includes('Permission') || error.message.includes('権限')) {
        userMessage = 'スプレッドシートへのアクセス権限がありません。編集権限を確認してください。';
      } else if (error.message.includes('not found') || error.message.includes('見つかりません')) {
        userMessage = 'スプレッドシートが見つかりません。URLが正しいか確認してください。';
      }

      return {
        success: false,
        message: userMessage,
        error: error.message
      };
    }
  },

  /**
   * システム管理者権限の確認
   * AdminPanel.js.html から呼び出される
   *
   * @returns {boolean} 管理者権限の有無
   */
  checkIsSystemAdmin() {
    try {
      const email = UserService.getCurrentEmail();
      if (!email) {
        return false;
      }

      return UserService.isSystemAdmin(email);
    } catch (error) {
      console.error('AdminController.checkIsSystemAdmin エラー:', error.message);
      return false;
    }
  },

  /**
   * 現在のボード情報とURLを取得
   * AdminPanel.js.html から呼び出される
   *
   * @returns {Object} ボード情報
   */
  getCurrentBoardInfoAndUrls() {
    try {
      // ユーザー情報の取得
      let userInfo = UserService.getCurrentUserInfo();
      let userId = userInfo && userInfo.userId;

      if (!userId) {
        const email = UserService.getCurrentEmail();
        if (!email) {
          return {
            isActive: false,
            error: 'ユーザー情報が見つかりません',
            appPublished: false
          };
        }

        // DB から直接ユーザー検索
        try {
          const dbUser = DB.findUserByEmail(email);
          if (dbUser && dbUser.userId) {
            userId = dbUser.userId;
            userInfo = { userId, userEmail: email };
          } else {
            return {
              isActive: false,
              error: 'ユーザーが見つかりません',
              appPublished: false
            };
          }
        } catch (e) {
          return {
            isActive: false,
            error: 'ユーザー情報の処理に失敗しました',
            appPublished: false
          };
        }
      }

      const config = ConfigService.getUserConfig(userId);
      if (!config || !config.appPublished) {
        return {
          isActive: false,
          appPublished: false,
          questionText: 'アプリケーションが公開されていません'
        };
      }

      // WebAppのベースURL取得
      const baseUrl = ScriptApp.getService().getUrl();
      const viewUrl = `${baseUrl}?userId=${userId}`;

      return {
        isActive: true,
        appPublished: true,
        questionText: config.questionText || config.boardTitle || 'Everyone\'s Answer Board',
        urls: {
          view: viewUrl,
          admin: `${baseUrl}?mode=admin&userId=${userId}`
        },
        lastUpdated: config.publishedAt || config.lastModified || new Date().toISOString()
      };

    } catch (error) {
      console.error('AdminController.getCurrentBoardInfoAndUrls エラー:', error.message);
      return {
        isActive: false,
        appPublished: false,
        error: error.message
      };
    }
  }

});

// ===========================================
// 📊 グローバル関数エクスポート（GAS互換性のため）
// ===========================================

/**
 * 管理パネル用API関数を個別にエクスポート
 * AdminPanel.js.html からの google.script.run 呼び出しに対応
 */

function getConfig() {
  return AdminController.getConfig();
}

function getSpreadsheetList() {
  return AdminController.getSpreadsheetList();
}

function getSheetList(spreadsheetId) {
  return AdminController.getSheetList(spreadsheetId);
}

function analyzeColumns(spreadsheetId, sheetName) {
  return AdminController.analyzeColumns(spreadsheetId, sheetName);
}

function saveDraftConfiguration(config) {
  return AdminController.saveDraftConfiguration(config);
}

function publishApplication(publishConfig) {
  return AdminController.publishApplication(publishConfig);
}

function validateAccess(spreadsheetId) {
  return AdminController.validateAccess(spreadsheetId);
}

function checkIsSystemAdmin() {
  return AdminController.checkIsSystemAdmin();
}

function getCurrentBoardInfoAndUrls() {
  return AdminController.getCurrentBoardInfoAndUrls();
}