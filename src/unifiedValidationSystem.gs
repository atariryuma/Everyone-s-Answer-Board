/**
 * 統一Validationシステム
 * 全ての検証機能を統合管理するクラス
 */

/**
 * 統一Validationシステムのメインクラス
 */
class UnifiedValidationSystem {
  constructor() {
    this.validationCategories = {
      AUTH: 'authentication',
      CONFIG: 'configuration',
      DATA: 'data_integrity',
      SYSTEM: 'system_health',
      WORKFLOW: 'workflow'
    };

    this.validationLevels = {
      BASIC: 'basic',
      STANDARD: 'standard',
      COMPREHENSIVE: 'comprehensive'
    };

    this.results = new Map();
  }

  /**
   * メイン検証実行メソッド
   * @param {string} category - 検証カテゴリ
   * @param {string} level - 検証レベル
   * @param {Object} options - 検証オプション
   * @returns {Object} 統一検証結果
   */
  validate(category, level = this.validationLevels.STANDARD, options = {}) {
    const validationId = `${category}_${level}_${Date.now()}`;
    const startTime = new Date();

    try {
      let result;

      switch (category) {
        case this.validationCategories.AUTH:
          result = this._validateAuthentication(level, options);
          break;
        case this.validationCategories.CONFIG:
          result = this._validateConfiguration(level, options);
          break;
        case this.validationCategories.DATA:
          result = this._validateDataIntegrity(level, options);
          break;
        case this.validationCategories.SYSTEM:
          result = this._validateSystemHealth(level, options);
          break;
        case this.validationCategories.WORKFLOW:
          result = this._validateWorkflow(level, options);
          break;
        default:
          throw new Error(`未対応の検証カテゴリ: ${category}`);
      }

      const finalResult = this._formatValidationResult(
        validationId,
        category,
        level,
        result,
        startTime,
        options
      );

      this.results.set(validationId, finalResult);
      return finalResult;

    } catch (error) {
      return this._handleValidationError(validationId, category, level, error, startTime, options);
    }
  }

  /**
   * 認証系検証
   */
  _validateAuthentication(level, options) {
    const results = {
      tests: {},
      summary: { passed: 0, total: 0, critical: 0 }
    };

    // 基本認証チェック
    if (level === this.validationLevels.BASIC || level === this.validationLevels.STANDARD || level === this.validationLevels.COMPREHENSIVE) {
      results.tests.userAuth = this._checkUserAuthentication(options.userId);
      results.tests.sessionValid = this._checkSessionValidity(options.userId);
    }

    // 標準認証チェック
    if (level === this.validationLevels.STANDARD || level === this.validationLevels.COMPREHENSIVE) {
      results.tests.userAccess = this._checkUserAccess(options.userId);
      results.tests.adminCheck = this._checkAdminPermissions(options.userId);
    }

    // 包括認証チェック
    if (level === this.validationLevels.COMPREHENSIVE) {
      results.tests.authHealth = this._checkAuthenticationHealth();
      results.tests.securityCheck = this._performAuthSecurityCheck(options.userId);
    }

    return this._calculateSummary(results);
  }

  /**
   * 設定系検証
   */
  _validateConfiguration(level, options) {
    const results = {
      tests: {},
      summary: { passed: 0, total: 0, critical: 0 }
    };

    // 基本設定チェック
    if (level === this.validationLevels.BASIC || level === this.validationLevels.STANDARD || level === this.validationLevels.COMPREHENSIVE) {
      results.tests.configJson = this._checkConfigJson(options.config);
      results.tests.configState = this._checkConfigState(options.config, options.userInfo);
    }

    // 標準設定チェック
    if (level === this.validationLevels.STANDARD || level === this.validationLevels.COMPREHENSIVE) {
      results.tests.systemConfig = this._checkSystemConfiguration();
      results.tests.dependencies = this._checkSystemDependencies();
    }

    // 包括設定チェック
    if (level === this.validationLevels.COMPREHENSIVE) {
      results.tests.serviceAccount = this._checkServiceAccountConfiguration();
      results.tests.applicationAccess = this._checkApplicationAccess();
    }

    return this._calculateSummary(results);
  }

  /**
   * データ整合性検証
   */
  _validateDataIntegrity(level, options) {
    const results = {
      tests: {},
      summary: { passed: 0, total: 0, critical: 0 }
    };

    // 基本データチェック
    if (level === this.validationLevels.BASIC || level === this.validationLevels.STANDARD || level === this.validationLevels.COMPREHENSIVE) {
      results.tests.headerIntegrity = this._checkHeaderIntegrity(options.userId);
      results.tests.requiredHeaders = this._checkRequiredHeaders(options.indices);
    }

    // 標準データチェック
    if (level === this.validationLevels.STANDARD || level === this.validationLevels.COMPREHENSIVE) {
      results.tests.duplicates = this._checkForDuplicates(options.headers, options.userRows);
      results.tests.missingFields = this._checkMissingRequiredFields(options.headers, options.userRows);
      results.tests.invalidFormats = this._checkInvalidDataFormats(options.headers, options.userRows);
    }

    // 包括データチェック
    if (level === this.validationLevels.COMPREHENSIVE) {
      results.tests.orphanedData = this._checkOrphanedData(options.headers, options.userRows);
      results.tests.dataIntegrity = this._performDataIntegrityCheck(options);
      results.tests.userScopedKey = this._checkUserScopedKey(options.key, options.expectedUserId);
    }

    return this._calculateSummary(results);
  }

  /**
   * システムヘルス検証
   */
  _validateSystemHealth(level, options) {
    const results = {
      tests: {},
      summary: { passed: 0, total: 0, critical: 0 }
    };

    // 基本ヘルスチェック
    if (level === this.validationLevels.BASIC || level === this.validationLevels.STANDARD || level === this.validationLevels.COMPREHENSIVE) {
      results.tests.basicHealth = this._performBasicHealthCheck();
      results.tests.cacheStatus = this._checkCacheStatus(options.userId);
    }

    // 標準ヘルスチェック
    if (level === this.validationLevels.STANDARD || level === this.validationLevels.COMPREHENSIVE) {
      results.tests.databaseHealth = this._checkDatabaseHealth();
      results.tests.performanceCheck = this._performPerformanceCheck();
    }

    // 包括ヘルスチェック
    if (level === this.validationLevels.COMPREHENSIVE) {
      results.tests.securityHealth = this._performSecurityCheck();
      results.tests.multiTenantHealth = this._performMultiTenantHealthCheck();
      results.tests.serviceAccountHealth = this._performServiceAccountHealthCheck();
      results.tests.userManagerHealth = this._checkUnifiedUserManagerHealth();
      results.tests.batchProcessorHealth = this._performUnifiedBatchHealthCheck();
    }

    return this._calculateSummary(results);
  }

  /**
   * ワークフロー検証
   */
  _validateWorkflow(level, options) {
    const results = {
      tests: {},
      summary: { passed: 0, total: 0, critical: 0 }
    };

    // 基本ワークフロー
    if (level === this.validationLevels.BASIC || level === this.validationLevels.STANDARD || level === this.validationLevels.COMPREHENSIVE) {
      results.tests.workflowGAS = this._validateWorkflowGAS(options.data);
      results.tests.publicationStatus = this._checkCurrentPublicationStatus(options.userId);
    }

    // 標準ワークフロー
    if (level === this.validationLevels.STANDARD || level === this.validationLevels.COMPREHENSIVE) {
      results.tests.workflowWithSheet = this._validateWorkflowWithSheet();
      results.tests.autoStop = this._checkAndHandleAutoStop(options.config, options.userInfo);
      results.tests.formUpdates = this._checkIfNewOrUpdatedForm(options.requestUserId, options.spreadsheetId, options.sheetName);
    }

    // 包括ワークフロー
    if (level === this.validationLevels.COMPREHENSIVE) {
      results.tests.comprehensiveWorkflow = this._comprehensiveWorkflowValidation();
    }

    return this._calculateSummary(results);
  }

  // ============================================================================
  // 個別検証メソッド（既存関数を統合）
  // ============================================================================

  /**
   * ユーザー認証チェック
   */
  _checkUserAuthentication(userId) {
    try {
      // validateUserAuthentication() の統合実装
      if (!userId) {
        return { passed: false, message: 'ユーザーIDが未定義', critical: true };
      }

      const userEmail = Session.getActiveUser().getEmail();
      if (!userEmail) {
        return { passed: false, message: 'ユーザーのEmailが取得できません', critical: true };
      }

      return { passed: true, message: 'ユーザー認証成功', userEmail, critical: false };
    } catch (error) {
      return { passed: false, message: `認証エラー: ${error.message}`, critical: true };
    }
  }

  /**
   * セッション有効性チェック
   */
  _checkSessionValidity(userId) {
    try {
      const loginStatus = checkLoginStatus(userId); // 既存関数利用
      return {
        passed: loginStatus.isValid,
        message: loginStatus.message || (loginStatus.isValid ? 'セッション有効' : 'セッション無効'),
        critical: !loginStatus.isValid
      };
    } catch (error) {
      return { passed: false, message: `セッションチェックエラー: ${error.message}`, critical: true };
    }
  }

  /**
   * ユーザーアクセスチェック
   */
  _checkUserAccess(userId) {
    try {
      const result = verifyUserAccess(userId); // 既存関数利用
      return {
        passed: result.hasAccess || false,
        message: result.message || 'アクセス権確認完了',
        critical: !result.hasAccess
      };
    } catch (error) {
      return { passed: false, message: `アクセス権チェックエラー: ${error.message}`, critical: true };
    }
  }

  /**
   * 管理者権限チェック
   */
  _checkAdminPermissions(userId) {
    try {
      const isAdmin = checkAdmin(userId); // 既存関数利用
      return {
        passed: true, // チェック自体は成功
        message: `管理者権限: ${isAdmin ? '有効' : '無効'}`,
        isAdmin,
        critical: false
      };
    } catch (error) {
      return { passed: false, message: `管理者権限チェックエラー: ${error.message}`, critical: false };
    }
  }

  /**
   * 認証ヘルスチェック
   */
  _checkAuthenticationHealth() {
    try {
      const result = checkAuthenticationHealth(); // 既存関数利用
      return {
        passed: result.status === 'healthy',
        message: result.message || '認証ヘルスチェック完了',
        details: result,
        critical: result.status !== 'healthy'
      };
    } catch (error) {
      return { passed: false, message: `認証ヘルスチェックエラー: ${error.message}`, critical: true };
    }
  }

  /**
   * 認証セキュリティチェック
   */
  _performAuthSecurityCheck(userId) {
    try {
      const result = performComprehensiveSecurityHealthCheck(); // 既存関数利用
      return {
        passed: result.overallStatus === 'HEALTHY',
        message: result.summary || 'セキュリティチェック完了',
        details: result,
        critical: result.overallStatus === 'CRITICAL'
      };
    } catch (error) {
      return { passed: false, message: `セキュリティチェックエラー: ${error.message}`, critical: true };
    }
  }

  /**
   * ConfigJSON検証
   */
  _checkConfigJson(config) {
    try {
      const result = validateConfigJson(config); // 既存関数利用
      return {
        passed: result.isValid,
        message: result.isValid ? 'Config JSON有効' : `Config JSON無効: ${result.errors.join(', ')}`,
        errors: result.errors,
        critical: !result.isValid
      };
    } catch (error) {
      return { passed: false, message: `Config JSONチェックエラー: ${error.message}`, critical: true };
    }
  }

  /**
   * Config状態検証
   */
  _checkConfigState(config, userInfo) {
    try {
      const result = validateConfigJsonState(config, userInfo); // 既存関数利用
      return {
        passed: result.isValid,
        message: result.message || 'Config状態チェック完了',
        critical: !result.isValid
      };
    } catch (error) {
      return { passed: false, message: `Config状態チェックエラー: ${error.message}`, critical: true };
    }
  }

  /**
   * システム設定チェック
   */
  _checkSystemConfiguration() {
    try {
      const result = checkSystemConfiguration(); // 既存関数利用
      return {
        passed: result.success || result.isValid,
        message: result.message || 'システム設定チェック完了',
        details: result,
        critical: !result.success && !result.isValid
      };
    } catch (error) {
      return { passed: false, message: `システム設定チェックエラー: ${error.message}`, critical: true };
    }
  }

  /**
   * システム依存関係チェック
   */
  _checkSystemDependencies() {
    try {
      const result = validateSystemDependencies(); // 既存関数利用
      return {
        passed: result.allDependenciesValid,
        message: result.message || 'システム依存関係チェック完了',
        details: result,
        critical: !result.allDependenciesValid
      };
    } catch (error) {
      return { passed: false, message: `システム依存関係チェックエラー: ${error.message}`, critical: true };
    }
  }

  /**
   * サービスアカウント設定チェック
   */
  _checkServiceAccountConfiguration() {
    try {
      const result = checkServiceAccountConfiguration(); // 既存関数利用
      return {
        passed: result.isValid,
        message: result.message || 'サービスアカウント設定チェック完了',
        details: result,
        critical: !result.isValid
      };
    } catch (error) {
      return { passed: false, message: `サービスアカウント設定チェックエラー: ${error.message}`, critical: true };
    }
  }

  /**
   * アプリケーションアクセスチェック
   */
  _checkApplicationAccess() {
    try {
      const result = checkApplicationAccess(); // 既存関数利用
      return {
        passed: result.hasAccess,
        message: result.message || 'アプリケーションアクセスチェック完了',
        critical: !result.hasAccess
      };
    } catch (error) {
      return { passed: false, message: `アプリケーションアクセスチェックエラー: ${error.message}`, critical: true };
    }
  }

  /**
   * ヘッダー整合性チェック
   */
  _checkHeaderIntegrity(userId) {
    try {
      const result = validateHeaderIntegrity(userId); // 既存関数利用
      return {
        passed: result.isValid,
        message: result.message || 'ヘッダー整合性チェック完了',
        details: result,
        critical: !result.isValid
      };
    } catch (error) {
      return { passed: false, message: `ヘッダー整合性チェックエラー: ${error.message}`, critical: true };
    }
  }

  /**
   * 必須ヘッダーチェック
   */
  _checkRequiredHeaders(indices) {
    try {
      const result = validateRequiredHeaders(indices); // 既存関数利用
      return {
        passed: result.isValid,
        message: result.message || '必須ヘッダーチェック完了',
        critical: !result.isValid
      };
    } catch (error) {
      return { passed: false, message: `必須ヘッダーチェックエラー: ${error.message}`, critical: true };
    }
  }

  /**
   * 重複データチェック
   */
  _checkForDuplicates(headers, userRows) {
    try {
      const result = checkForDuplicates(headers, userRows); // 既存関数利用
      return {
        passed: result.duplicates.length === 0,
        message: result.duplicates.length === 0 ? '重複なし' : `重複データ検出: ${result.duplicates.length}件`,
        duplicates: result.duplicates,
        critical: result.duplicates.length > 10 // 10件以上は重要
      };
    } catch (error) {
      return { passed: false, message: `重複チェックエラー: ${error.message}`, critical: false };
    }
  }

  /**
   * 必須フィールド欠損チェック
   */
  _checkMissingRequiredFields(headers, userRows) {
    try {
      const result = checkMissingRequiredFields(headers, userRows); // 既存関数利用
      return {
        passed: result.missing.length === 0,
        message: result.missing.length === 0 ? '必須フィールドOK' : `必須フィールド欠損: ${result.missing.length}件`,
        missing: result.missing,
        critical: result.missing.length > 0
      };
    } catch (error) {
      return { passed: false, message: `必須フィールドチェックエラー: ${error.message}`, critical: true };
    }
  }

  /**
   * 無効データ形式チェック
   */
  _checkInvalidDataFormats(headers, userRows) {
    try {
      const result = checkInvalidDataFormats(headers, userRows); // 既存関数利用
      return {
        passed: result.invalid.length === 0,
        message: result.invalid.length === 0 ? 'データ形式OK' : `無効データ形式: ${result.invalid.length}件`,
        invalid: result.invalid,
        critical: result.invalid.length > 5 // 5件以上は重要
      };
    } catch (error) {
      return { passed: false, message: `データ形式チェックエラー: ${error.message}`, critical: false };
    }
  }

  /**
   * 孤立データチェック
   */
  _checkOrphanedData(headers, userRows) {
    try {
      const result = checkOrphanedData(headers, userRows); // 既存関数利用
      return {
        passed: result.orphaned.length === 0,
        message: result.orphaned.length === 0 ? '孤立データなし' : `孤立データ検出: ${result.orphaned.length}件`,
        orphaned: result.orphaned,
        critical: false
      };
    } catch (error) {
      return { passed: false, message: `孤立データチェックエラー: ${error.message}`, critical: false };
    }
  }

  /**
   * データ整合性チェック
   */
  _performDataIntegrityCheck(options) {
    try {
      const result = performDataIntegrityCheck(options); // 既存関数利用
      return {
        passed: result.overallStatus === 'HEALTHY',
        message: result.summary || 'データ整合性チェック完了',
        details: result,
        critical: result.overallStatus === 'CRITICAL'
      };
    } catch (error) {
      return { passed: false, message: `データ整合性チェックエラー: ${error.message}`, critical: true };
    }
  }

  /**
   * ユーザースコープキー検証
   */
  _checkUserScopedKey(key, expectedUserId) {
    try {
      const result = validateUserScopedKey(key, expectedUserId); // 既存関数利用
      return {
        passed: result.isValid,
        message: result.message || 'ユーザースコープキー検証完了',
        critical: !result.isValid
      };
    } catch (error) {
      return { passed: false, message: `ユーザースコープキー検証エラー: ${error.message}`, critical: true };
    }
  }

  /**
   * 基本ヘルスチェック
   */
  _performBasicHealthCheck() {
    try {
      const result = performHealthCheck(); // 既存関数利用
      return {
        passed: result.overallStatus === 'HEALTHY',
        message: result.summary || '基本ヘルスチェック完了',
        details: result,
        critical: result.overallStatus === 'CRITICAL'
      };
    } catch (error) {
      return { passed: false, message: `基本ヘルスチェックエラー: ${error.message}`, critical: true };
    }
  }

  /**
   * キャッシュ状態チェック
   */
  _checkCacheStatus(userId) {
    try {
      const result = checkCacheStatus(userId); // 既存関数利用
      return {
        passed: result.status === 'healthy',
        message: result.message || 'キャッシュ状態チェック完了',
        details: result,
        critical: result.status === 'critical'
      };
    } catch (error) {
      return { passed: false, message: `キャッシュ状態チェックエラー: ${error.message}`, critical: false };
    }
  }

  /**
   * データベースヘルスチェック
   */
  _checkDatabaseHealth() {
    try {
      const result = checkDatabaseHealth(); // 既存関数利用
      return {
        passed: result.status === 'healthy',
        message: result.message || 'データベースヘルスチェック完了',
        details: result,
        critical: result.status !== 'healthy'
      };
    } catch (error) {
      return { passed: false, message: `データベースヘルスチェックエラー: ${error.message}`, critical: true };
    }
  }

  /**
   * パフォーマンスチェック
   */
  _performPerformanceCheck() {
    try {
      const result = performPerformanceCheck(); // 既存関数利用
      return {
        passed: result.overallStatus === 'HEALTHY',
        message: result.summary || 'パフォーマンスチェック完了',
        details: result,
        critical: result.overallStatus === 'CRITICAL'
      };
    } catch (error) {
      return { passed: false, message: `パフォーマンスチェックエラー: ${error.message}`, critical: false };
    }
  }

  /**
   * セキュリティチェック
   */
  _performSecurityCheck() {
    try {
      const result = performSecurityCheck(); // 既存関数利用
      return {
        passed: result.overallStatus === 'HEALTHY',
        message: result.summary || 'セキュリティチェック完了',
        details: result,
        critical: result.overallStatus === 'CRITICAL'
      };
    } catch (error) {
      return { passed: false, message: `セキュリティチェックエラー: ${error.message}`, critical: true };
    }
  }

  /**
   * マルチテナントヘルスチェック
   */
  _performMultiTenantHealthCheck() {
    try {
      const result = performMultiTenantHealthCheck(); // 既存関数利用
      return {
        passed: result.overallStatus === 'HEALTHY',
        message: result.summary || 'マルチテナントヘルスチェック完了',
        details: result,
        critical: result.overallStatus === 'CRITICAL'
      };
    } catch (error) {
      return { passed: false, message: `マルチテナントヘルスチェックエラー: ${error.message}`, critical: false };
    }
  }

  /**
   * サービスアカウントヘルスチェック
   */
  _performServiceAccountHealthCheck() {
    try {
      const result = performServiceAccountHealthCheck(); // 既存関数利用
      return {
        passed: result.status === 'healthy',
        message: result.message || 'サービスアカウントヘルスチェック完了',
        details: result,
        critical: result.status !== 'healthy'
      };
    } catch (error) {
      return { passed: false, message: `サービスアカウントヘルスチェックエラー: ${error.message}`, critical: true };
    }
  }

  /**
   * ユーザーマネージャーヘルスチェック
   */
  _checkUnifiedUserManagerHealth() {
    try {
      const result = checkUnifiedUserManagerHealth(); // 既存関数利用
      return {
        passed: result.overallStatus === 'HEALTHY',
        message: result.summary || 'ユーザーマネージャーヘルスチェック完了',
        details: result,
        critical: result.overallStatus === 'CRITICAL'
      };
    } catch (error) {
      return { passed: false, message: `ユーザーマネージャーヘルスチェックエラー: ${error.message}`, critical: false };
    }
  }

  /**
   * バッチプロセッサーヘルスチェック
   */
  _performUnifiedBatchHealthCheck() {
    try {
      const result = performUnifiedBatchHealthCheck(); // 既存関数利用
      return {
        passed: result.overallStatus === 'HEALTHY',
        message: result.summary || 'バッチプロセッサーヘルスチェック完了',
        details: result,
        critical: result.overallStatus === 'CRITICAL'
      };
    } catch (error) {
      return { passed: false, message: `バッチプロセッサーヘルスチェックエラー: ${error.message}`, critical: false };
    }
  }

  /**
   * ワークフローGAS検証
   */
  _validateWorkflowGAS(data) {
    try {
      const result = validateWorkflowGAS(data); // 既存関数利用
      return {
        passed: result.valid,
        message: result.message,
        details: result,
        critical: !result.valid
      };
    } catch (error) {
      return { passed: false, message: `ワークフローGAS検証エラー: ${error.message}`, critical: true };
    }
  }

  /**
   * 公開状態チェック
   */
  _checkCurrentPublicationStatus(userId) {
    try {
      const result = checkCurrentPublicationStatus(userId); // 既存関数利用
      return {
        passed: true, // 取得自体は成功
        message: `公開状態: ${result.isPublished ? '公開中' : '非公開'}`,
        isPublished: result.isPublished,
        critical: false
      };
    } catch (error) {
      return { passed: false, message: `公開状態チェックエラー: ${error.message}`, critical: false };
    }
  }

  /**
   * スプレッドシート付きワークフロー検証
   */
  _validateWorkflowWithSheet() {
    try {
      const result = validateWorkflowWithSheet(); // 既存関数利用
      return {
        passed: result.valid,
        message: result.message,
        details: result,
        critical: !result.valid
      };
    } catch (error) {
      return { passed: false, message: `スプレッドシート付きワークフロー検証エラー: ${error.message}`, critical: true };
    }
  }

  /**
   * 自動停止チェック
   */
  _checkAndHandleAutoStop(config, userInfo) {
    try {
      const result = checkAndHandleAutoStop(config, userInfo); // 既存関数利用
      return {
        passed: true, // チェック自体は成功
        message: result ? '自動停止実行' : '自動停止なし',
        autoStopExecuted: result,
        critical: false
      };
    } catch (error) {
      return { passed: false, message: `自動停止チェックエラー: ${error.message}`, critical: false };
    }
  }

  /**
   * フォーム更新チェック
   */
  _checkIfNewOrUpdatedForm(requestUserId, spreadsheetId, sheetName) {
    try {
      const result = checkIfNewOrUpdatedForm(requestUserId, spreadsheetId, sheetName); // 既存関数利用
      return {
        passed: true, // チェック自体は成功
        message: result.isNewOrUpdated ? 'フォーム更新あり' : 'フォーム更新なし',
        isNewOrUpdated: result.isNewOrUpdated,
        critical: false
      };
    } catch (error) {
      return { passed: false, message: `フォーム更新チェックエラー: ${error.message}`, critical: false };
    }
  }

  /**
   * 包括ワークフロー検証
   */
  _comprehensiveWorkflowValidation() {
    try {
      const result = comprehensiveWorkflowValidation(); // 既存関数利用
      return {
        passed: result.overall.success,
        message: `包括ワークフロー検証: ${result.overall.passed}/${result.overall.total} 通過`,
        details: result,
        critical: !result.overall.success
      };
    } catch (error) {
      return { passed: false, message: `包括ワークフロー検証エラー: ${error.message}`, critical: true };
    }
  }

  // ============================================================================
  // ユーティリティメソッド
  // ============================================================================

  /**
   * 検証結果のサマリー計算
   */
  _calculateSummary(results) {
    let passed = 0;
    let total = 0;
    let critical = 0;

    for (const [testName, testResult] of Object.entries(results.tests)) {
      total++;
      if (testResult.passed) passed++;
      if (testResult.critical) critical++;
    }

    results.summary = { passed, total, critical };
    results.summary.success = passed === total;
    results.summary.hasCritical = critical > 0;

    return results;
  }

  /**
   * 検証結果の標準化
   */
  _formatValidationResult(validationId, category, level, result, startTime, options) {
    const endTime = new Date();
    const duration = endTime.getTime() - startTime.getTime();

    return {
      validationId,
      category,
      level,
      timestamp: endTime.toISOString(),
      duration: `${duration}ms`,
      success: result.summary.success,
      hasCritical: result.summary.hasCritical,
      summary: {
        passed: result.summary.passed,
        total: result.summary.total,
        critical: result.summary.critical,
        passRate: result.summary.total > 0 ? Math.round((result.summary.passed / result.summary.total) * 100) : 0
      },
      tests: result.tests,
      options: this._sanitizeOptions(options),
      metadata: {
        executedAt: startTime.toISOString(),
        completedAt: endTime.toISOString(),
        version: '1.0.0'
      }
    };
  }

  /**
   * 検証エラーハンドリング
   */
  _handleValidationError(validationId, category, level, error, startTime, options) {
    const endTime = new Date();
    const duration = endTime.getTime() - startTime.getTime();

    ULog('ERROR', 'UnifiedValidationSystem', 'Validation execution failed', {
      validationId,
      category,
      level,
      error: error.message,
      duration: `${duration}ms`
    });

    return {
      validationId,
      category,
      level,
      timestamp: endTime.toISOString(),
      duration: `${duration}ms`,
      success: false,
      hasCritical: true,
      summary: {
        passed: 0,
        total: 1,
        critical: 1,
        passRate: 0
      },
      error: {
        message: error.message,
        stack: error.stack
      },
      tests: {
        validationExecution: {
          passed: false,
          message: `検証実行エラー: ${error.message}`,
          critical: true
        }
      },
      options: this._sanitizeOptions(options),
      metadata: {
        executedAt: startTime.toISOString(),
        completedAt: endTime.toISOString(),
        version: '1.0.0'
      }
    };
  }

  /**
   * オプションの秘匿情報除去
   */
  _sanitizeOptions(options) {
    const sanitized = { ...options };
    
    // パスワード、トークン等の機密情報を除去
    const sensitiveKeys = ['password', 'token', 'secret', 'key', 'auth'];
    
    for (const key of Object.keys(sanitized)) {
      if (sensitiveKeys.some(sensitive => key.toLowerCase().includes(sensitive))) {
        sanitized[key] = '[REDACTED]';
      }
    }

    return sanitized;
  }

  /**
   * 検証結果の取得
   */
  getValidationResult(validationId) {
    return this.results.get(validationId);
  }

  /**
   * 全検証結果の取得
   */
  getAllValidationResults() {
    return Array.from(this.results.values());
  }

  /**
   * 検証結果のクリア
   */
  clearValidationResults(olderThanHours = 24) {
    const cutoffTime = new Date(Date.now() - olderThanHours * 60 * 60 * 1000);
    
    for (const [validationId, result] of this.results) {
      const resultTime = new Date(result.timestamp);
      if (resultTime < cutoffTime) {
        this.results.delete(validationId);
      }
    }
  }
}

// ============================================================================
// グローバル関数（既存関数との互換性維持）
// ============================================================================

/**
 * 統一Validationシステムのインスタンス
 */
const UnifiedValidation = new UnifiedValidationSystem();

/**
 * 簡易検証実行（既存関数のリプレース用）
 */
function runValidation(category, level = 'standard', options = {}) {
  return UnifiedValidation.validate(category, level, options);
}

/**
 * 認証系検証（既存関数統合）
 */
function validateAuthentication(userId, level = 'standard') {
  return UnifiedValidation.validate('authentication', level, { userId });
}

/**
 * 設定系検証（既存関数統合）
 */
function validateConfiguration(config, userInfo, level = 'standard') {
  return UnifiedValidation.validate('configuration', level, { config, userInfo });
}

/**
 * データ整合性検証（既存関数統合）
 */
function validateDataIntegrity(userId, headers, userRows, level = 'standard') {
  return UnifiedValidation.validate('data_integrity', level, {
    userId,
    headers,
    userRows
  });
}

/**
 * システムヘルス検証（既存関数統合）
 */
function validateSystemHealth(userId, level = 'standard') {
  return UnifiedValidation.validate('system_health', level, { userId });
}

/**
 * ワークフロー検証（既存関数統合）
 */
function validateWorkflow(userId, data, config, userInfo, level = 'standard') {
  return UnifiedValidation.validate('workflow', level, {
    userId,
    data,
    config,
    userInfo
  });
}

/**
 * 包括検証（全カテゴリ実行）
 */
function comprehensiveValidation(options = {}) {
  const results = {};
  const categories = ['authentication', 'configuration', 'data_integrity', 'system_health', 'workflow'];
  
  for (const category of categories) {
    try {
      results[category] = UnifiedValidation.validate(category, 'comprehensive', options);
    } catch (error) {
      results[category] = {
        success: false,
        error: error.message,
        category
      };
    }
  }

  return {
    timestamp: new Date().toISOString(),
    overallSuccess: Object.values(results).every(r => r.success),
    categories: results,
    summary: {
      total: categories.length,
      passed: Object.values(results).filter(r => r.success).length,
      critical: Object.values(results).filter(r => r.hasCritical).length
    }
  };
}