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

/* global UserService, ConfigService, DataService, SecurityService, DB, UnifiedLogger */

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
      let email = userInfo && userInfo.userEmail;

      if (!userId) {
        email = email || UserService.getCurrentEmail();
        if (!email) {
          return { success: false, message: 'ユーザー情報が見つかりません' };
        }

        // UserService経由でユーザー検索を試行
        try {
          const foundUser = UserService.findUserByEmail(email);
          if (foundUser && foundUser.userId) {
            userId = foundUser.userId;
            userInfo = userInfo || { userId, userEmail: email };
          }
        } catch (e) {
          // ignore and fallback to auto-create
        }

        // Auto-create user if still missing
        if (!userId) {
          try {
            const created = UserService.createUser(email);
            // UserService.createUser() の戻り値構造をチェック
            const actualUser = created && created.value ? created.value : created;
            if (actualUser && actualUser.userId) {
              userId = actualUser.userId;
              userInfo = actualUser;
            } else {
              console.error('AdminController.getConfig: ユーザー作成に失敗 - userIdが見つかりません', {
                created,
                actualUser,
                hasValue: created && !!created.value,
                hasUserId: actualUser && !!actualUser.userId
              });
              return { success: false, message: 'ユーザー作成に失敗しました' };
            }
          } catch (createErr) {
            console.error('AdminController.getConfig: ユーザー作成エラー', createErr);
            return { success: false, message: 'ユーザー作成に失敗しました' };
          }
        }
      }

      // Final validation before calling getUserConfig
      if (!userId) {
        console.error('AdminController.getConfig: userId が未定義です', { userInfo, email });
        return { success: false, message: 'ユーザーIDが取得できませんでした' };
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
    // 直接DataServiceに委譲（冗長なログを除去）
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
    return SecurityService.validateSpreadsheetAccess(spreadsheetId);
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

        // UserService経由でユーザー検索
        try {
          const foundUser = UserService.findUserByEmail(email);
          if (foundUser && foundUser.userId) {
            userId = foundUser.userId;
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
  },

  /**
   * フォーム情報を取得
   * AdminPanel.js.html から呼び出される
   *
   * @param {string} spreadsheetId - スプレッドシートID
   * @param {string} sheetName - シート名
   * @returns {Object} フォーム情報
   */
  getFormInfo(spreadsheetId, sheetName) {
    try {
      return ConfigService.getFormInfo(spreadsheetId, sheetName);
    } catch (error) {
      console.error('AdminController.getFormInfo エラー:', error.message);
      return {
        success: false,
        error: error.message
      };
    }
  },

  /**
   * 現在の公開状態を確認
   * AdminPanel.js.html から呼び出される
   *
   * @returns {Object} 公開状態情報
   */
  checkCurrentPublicationStatus() {
    try {
      // ユーザー情報の取得
      const userInfo = UserService.getCurrentUserInfo();
      const userId = userInfo && userInfo.userId;

      if (!userId) {
        return {
          success: false,
          published: false,
          error: 'ユーザー情報が見つかりません'
        };
      }

      // 設定情報を取得
      const config = ConfigService.getUserConfig(userId);
      if (!config) {
        return {
          success: false,
          published: false,
          error: '設定情報が見つかりません'
        };
      }

      return {
        success: true,
        published: config.appPublished === true,
        publishedAt: config.publishedAt || null,
        lastModified: config.lastModified || null,
        hasDataSource: !!(config.spreadsheetId && config.sheetName)
      };

    } catch (error) {
      console.error('AdminController.checkCurrentPublicationStatus エラー:', error.message);
      return {
        success: false,
        published: false,
        error: error.message
      };
    }
  }

});

// ===========================================
// 📊 重複削除完了 - グローバル関数エクスポート削除
// ===========================================

/**
 * 注意: google.script.run は直接 AdminController.methodName() を呼び出してください
 * 例: google.script.run.withSuccessHandler(resolve).AdminController.getConfig()
 *
 * グローバル関数の重複を削除し、適切なオブジェクト指向アプローチを採用
 */