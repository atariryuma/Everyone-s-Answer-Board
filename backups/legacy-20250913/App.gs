/**
 * アプリケーション名前空間システム
 * 2024年GASベストプラクティス準拠の初期化システム
 */

const App = {
  // システム状態
  initialized: false,

  // サービスインスタンス
  _config: null,
  _access: null,

  /**
   * システム初期化
   * 依存関係を適切に管理した初期化処理
   */
  init() {
    if (this.initialized) return;

    try {
      // ConfigurationManagerを最初に初期化
      this._config = new ConfigurationManager();

      // AccessControllerを初期化（ConfigurationManagerを注入）
      this._access = new AccessController(this._config);

      this.initialized = true;
      console.log('✅ App システム初期化完了');
    } catch (error) {
      console.error('❌ App システム初期化エラー:', error);
      throw error;
    }
  },

  /**
   * ConfigurationManagerインスタンスを取得
   * @return {ConfigurationManager} 設定管理インスタンス
   */
  getConfig() {
    if (!this.initialized) this.init();
    return this._config;
  },

  /**
   * AccessControllerインスタンスを取得
   * @return {AccessController} アクセス制御インスタンス
   */
  getAccess() {
    if (!this.initialized) this.init();
    return this._access;
  },

  /**
   * システムリセット（テスト用）
   * テスト環境でのみ使用
   */
  reset() {
    this.initialized = false;
    this._config = null;
    this._access = null;
  },

  /**
   * システム状態確認
   * @return {Object} システム状態
   */
  getStatus() {
    return {
      initialized: this.initialized,
      hasConfig: !!this._config,
      hasAccess: !!this._access,
      timestamp: new Date().toISOString(),
    };
  },
};
