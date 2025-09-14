/**
 * @fileoverview FrontendController - フロントエンド用API関数の集約
 *
 * 🎯 責任範囲:
 * - フロントエンドHTMLからの呼び出し処理
 * - 基本的なユーザー情報取得
 * - 認証・ログイン関連機能
 * - システムユーティリティ関数
 *
 * 📝 main.gsから移動されたフロントエンドAPI関数群
 */

/* global UserService */

/**
 * FrontendController - フロントエンド用コントローラー
 * HTMLファイル（login.js.html, SetupPage.html等）からの呼び出しを処理
 */
const FrontendController = Object.freeze({

  // ===========================================
  // 📊 基本ユーザー情報API
  // ===========================================

  /**
   * 現在のユーザー情報を取得
   * login.js.html, SetupPage.html, AdminPanel.js.html から呼び出される
   *
   * @param {string} [kind='email'] - 取得する情報の種類（'email' or 'full'）
   * @returns {Object|string|null} 統一されたレスポンス形式
   */
  getUser(kind = 'email') {
    try {
      const userEmail = UserService.getCurrentEmail();

      if (!userEmail) {
        return kind === 'email' ? '' : { success: false, message: 'ユーザー情報が取得できません' };
      }

      // 後方互換性重視: kind==='email' の場合は純粋な文字列を返す
      if (kind === 'email') {
        return String(userEmail);
      }

      // 統一オブジェクト形式（'full' など）
      const userInfo = UserService.getCurrentUserInfo();
      return {
        success: true,
        email: userEmail,
        userId: userInfo?.userId || null,
        isActive: userInfo?.isActive || false,
        hasConfig: !!userInfo?.config
      };
    } catch (error) {
      console.error('FrontendController.getUser エラー:', error.message);
      return kind === 'email' ? '' : { success: false, message: error.message };
    }
  },

  /**
   * WebアプリケーションのURLを取得
   * 複数のHTMLファイルから呼び出される基本機能
   *
   * @returns {string} WebアプリのURL
   */
  getWebAppUrl() {
    try {
      return ScriptApp.getService().getUrl();
    } catch (error) {
      console.error('FrontendController.getWebAppUrl エラー:', error.message);
      return '';
    }
  },

  // ===========================================
  // 📊 認証・ログイン関連API
  // ===========================================

  /**
   * ログインアクションを処理
   * login.js.html から呼び出される
   *
   * @returns {Object} ログイン処理結果
   */
  processLoginAction() {
    try {
      const userEmail = UserService.getCurrentEmail();
      if (!userEmail) {
        return {
          success: false,
          message: 'ユーザー情報が取得できません',
          needsAuth: true
        };
      }

      // ユーザー情報を取得または作成
      let userInfo = UserService.getCurrentUserInfo();
      if (!userInfo) {
        userInfo = UserService.createUser(userEmail);
        // createUserの戻り値構造を正規化
        if (userInfo && userInfo.value) {
          userInfo = userInfo.value;
        }
      }

      // 管理パネル用URLを構築（userId必須）
      const baseUrl = this.getWebAppUrl();
      const userId = userInfo?.userId;

      if (!userId) {
        return {
          success: false,
          message: 'ユーザーIDの取得に失敗しました',
          error: 'USER_ID_MISSING'
        };
      }

      const adminUrl = `${baseUrl}?mode=admin&userId=${userId}`;

      return {
        success: true,
        userInfo,
        redirectUrl: baseUrl,
        adminUrl,
        // 後方互換性のための追加プロパティ
        appUrl: baseUrl,
        url: adminUrl
      };

    } catch (error) {
      console.error('FrontendController.processLoginAction エラー:', error.message);
      return {
        success: false,
        message: error.message
      };
    }
  },

  /**
   * 認証状態を確認
   * login.js.html, SetupPage.html から呼び出される
   *
   * @returns {Object} 認証状態
   */
  verifyUserAuthentication() {
    try {
      const userEmail = UserService.getCurrentEmail();
      if (!userEmail) {
        return {
          isAuthenticated: false,
          message: '認証が必要です'
        };
      }

      const userInfo = UserService.getCurrentUserInfo();
      return {
        isAuthenticated: true,
        userEmail,
        userInfo,
        hasConfig: !!userInfo?.config
      };

    } catch (error) {
      console.error('FrontendController.verifyUserAuthentication エラー:', error.message);
      return {
        isAuthenticated: false,
        message: error.message
      };
    }
  },

  /**
   * 認証情報のリセット
   * login.js.html から呼び出される
   *
   * @returns {Object} リセット結果
   */
  resetAuth() {
    try {
      console.log('FrontendController.resetAuth: 認証リセット開始');

      // ユーザーキャッシュクリア
      UserService.clearUserCache();

      // セッション関連の情報をクリア
      const props = PropertiesService.getScriptProperties();

      // 一時的な認証情報をクリア
      const authKeys = ['temp_auth_token', 'last_login_attempt', 'auth_retry_count'];
      authKeys.forEach(key => {
        props.deleteProperty(key);
      });

      console.log('FrontendController.resetAuth: 認証リセット完了');

      return {
        success: true,
        message: '認証情報がリセットされました'
      };

    } catch (error) {
      console.error('FrontendController.resetAuth エラー:', error.message);
      return {
        success: false,
        message: `認証リセットに失敗しました: ${error.message}`
      };
    }
  },

  /**
   * ログイン状態を取得
   * login.js.html から呼び出される
   *
   * @returns {Object} ログイン状態
   */
  getLoginStatus() {
    try {
      const userEmail = UserService.getCurrentEmail();
      if (!userEmail) {
        return {
          isLoggedIn: false,
          user: null
        };
      }

      const userInfo = UserService.getCurrentUserInfo();
      return {
        isLoggedIn: true,
        user: {
          email: userEmail,
          userId: userInfo?.userId,
          hasSetup: !!userInfo?.config?.setupComplete
        }
      };

    } catch (error) {
      console.error('FrontendController.getLoginStatus エラー:', error.message);
      return {
        isLoggedIn: false,
        user: null,
        error: error.message
      };
    }
  },

  // ===========================================
  // 📊 エラー報告・デバッグAPI
  // ===========================================

  /**
   * クライアントエラーを報告
   * ErrorBoundary.html から呼び出される
   *
   * @param {Object} errorInfo - エラー情報
   * @returns {Object} 報告結果
   */
  reportClientError(errorInfo) {
    try {
      console.error('クライアントエラー報告:', errorInfo);

      // エラーログを記録（将来的にはSecurityServiceや専用のログサービスに委譲）
      const logEntry = {
        timestamp: new Date().toISOString(),
        type: 'client_error',
        userEmail: UserService.getCurrentEmail() || 'unknown',
        errorInfo
      };

      // コンソールにログ出力（将来的には永続化）
      console.log('Error Log Entry:', JSON.stringify(logEntry));

      return {
        success: true,
        message: 'エラーが報告されました'
      };
    } catch (error) {
      console.error('FrontendController.reportClientError エラー:', error.message);
      return {
        success: false,
        message: error.message
      };
    }
  },

  /**
   * 強制ログアウトとリダイレクトのテスト
   * ErrorBoundary.html から呼び出される
   *
   * @returns {Object} テスト結果
   */
  testForceLogoutRedirect() {
    try {
      console.log('強制ログアウトテストが実行されました');

      return {
        success: true,
        message: 'ログアウトテスト完了',
        redirectUrl: `${this.getWebAppUrl()}?mode=login`
      };
    } catch (error) {
      console.error('FrontendController.testForceLogoutRedirect エラー:', error.message);
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
 * 重複削除完了 - グローバル関数エクスポート削除
 * 使用方法: google.script.run.FrontendController.methodName()
 *
 * 適切なオブジェクト指向アプローチを採用し、グローバル関数の重複を回避
 */
