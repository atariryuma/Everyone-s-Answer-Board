/**
 * @fileoverview 統合セキュリティマネージャー - 全セキュリティシステムを統合
 * auth.gs、multiTenantSecurity.gs、securityHealthCheck.gs の機能を統合
 * 既存の全ての関数をそのまま保持し、互換性を完全維持
 */

// =============================================================================
// SECTION 1: 認証管理システム（元auth.gs）
// =============================================================================

// 認証管理のための定数
const AUTH_CACHE_KEY = 'SA_TOKEN_CACHE';
const TOKEN_EXPIRY_BUFFER = 300; // 5分のバッファ

/**
 * キャッシュされたサービスアカウントトークンを取得
 * @returns {string} アクセストークン
 */
function getServiceAccountTokenCached() {
  return cacheManager.get(AUTH_CACHE_KEY, generateNewServiceAccountToken, {
    ttl: 3500,
    enableMemoization: true,
  }); // メモ化対応でトークン取得を高速化
}

/**
 * 新しいJWTトークンを生成
 * @returns {string} アクセストークン
 */
function generateNewServiceAccountToken() {
  // 統一秘密情報管理システムで安全に取得
  const serviceAccountCreds = getSecureServiceAccountCreds();

  const privateKey = serviceAccountCreds.private_key.replace(/\n/g, '\n'); // 改行文字を正規化
  const clientEmail = serviceAccountCreds.client_email;
  const tokenUrl = 'https://www.googleapis.com/oauth2/v4/token';

  const now = Math.floor(Date.now() / 1000);
  const expiresAt = now + 3600; // 1時間後

  // JWT生成
  const jwtHeader = { alg: 'RS256', typ: 'JWT' };
  const jwtClaimSet = {
    iss: clientEmail,
    scope: 'https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive',
    aud: tokenUrl,
    exp: expiresAt,
    iat: now,
  };

  const encodedHeader = Utilities.base64EncodeWebSafe(JSON.stringify(jwtHeader));
  const encodedClaimSet = Utilities.base64EncodeWebSafe(JSON.stringify(jwtClaimSet));
  const signatureInput = encodedHeader + '.' + encodedClaimSet;
  const signature = Utilities.computeRsaSha256Signature(signatureInput, privateKey);
  const encodedSignature = Utilities.base64EncodeWebSafe(signature);
  const jwt = signatureInput + '.' + encodedSignature;

  // トークンリクエスト
  const response = resilientUrlFetch(tokenUrl, {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: {
      grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
      assertion: jwt,
    },
    muteHttpExceptions: true,
  });

  // レスポンスオブジェクト検証（resilientUrlFetchで既に検証済みだが、念のため）
  if (!response || typeof response.getResponseCode !== 'function') {
    throw new Error('サービスアカウント認証: 無効なレスポンスオブジェクトが返されました');
  }

  const responseCode = response.getResponseCode();
  if (responseCode !== 200) {
    const responseText = response.getContentText();
    console.error('[ERROR]', 'Token request failed. Status:', responseCode);
    console.error('[ERROR]', 'Response:', responseText);

    // より詳細なエラーメッセージ
    let errorMessage = 'サービスアカウントトークンの取得に失敗しました。';
    if (responseCode === 400) {
      errorMessage += ' 認証情報またはJWTの形式に問題があります。';
    } else if (responseCode === 401) {
      errorMessage += ' 認証情報が無効です。サービスアカウントキーを確認してください。';
    } else if (responseCode === 403) {
      errorMessage += ' 権限が不足しています。サービスアカウントの権限を確認してください。';
    } else {
      errorMessage += ` Status: ${responseCode}`;
    }

    throw new Error(errorMessage);
  }

  var responseData = JSON.parse(response.getContentText());
  if (!responseData.access_token) {
    console.error('[ERROR]', 'No access token in response:', responseData);
    throw new Error(
      'アクセストークンが返されませんでした。サービスアカウント設定を確認してください。'
    );
  }

  console.log('Service account token generated successfully for:', clientEmail);
  return responseData.access_token;
}

/**
 * トークンキャッシュをクリア
 */
function clearServiceAccountTokenCache() {
  cacheManager.remove(AUTH_CACHE_KEY);
  console.log('トークンキャッシュをクリアしました');
}

/**
 * 設定されているサービスアカウントのメールアドレスを取得
 * @returns {string} サービスアカウントのメールアドレス
 */
function getServiceAccountEmail() {
  try {
    const serviceAccountCreds = getSecureServiceAccountCreds();
    return serviceAccountCreds.client_email || 'メールアドレス不明';
  } catch (error) {
    console.warn('サービスアカウントメール取得エラー:', error.message);
    return 'メールアドレス取得エラー';
  }
}

/**
 * シンプル・確実な管理者権限検証（3重チェック）
 * メールアドレス + ユーザーID + アクティブ状態の照合
 * @param {string} userId - 検証するユーザーのID
 * @returns {boolean} 管理者権限がある場合は true、そうでなければ false
 */
function verifyAdminAccess(userId) {
  try {
    // 基本的な引数チェック
    if (!userId || typeof userId !== 'string' || userId.trim() === '') {
      console.warn('verifyAdminAccess: 無効なuserIdが渡されました:', userId);
      return false;
    }

    // 現在のログインユーザーのメールアドレスを取得
    const activeUserEmail = User.email();
    if (!activeUserEmail) {
      console.warn('verifyAdminAccess: アクティブユーザーのメールアドレスが取得できませんでした');
      return false;
    }

    console.log('verifyAdminAccess: 認証開始', { userId, activeUserEmail });

    // データベースからユーザー情報を取得（統一検索関数使用）
    let userFromDb = null;

    // まずは通常のキャッシュ付き検索を試行
    userFromDb = DB.findUserById(userId);

    // 見つからない場合は強制フレッシュで再試行
    if (!userFromDb) {
      console.log('verifyAdminAccess: 強制フレッシュで再検索中...');
      userFromDb = fetchUserFromDatabase('userId', userId, { forceFresh: true });
    }

    // ユーザーが見つからない場合は認証失敗
    if (!userFromDb) {
      console.warn('verifyAdminAccess: ユーザーが見つかりません:', {
        userId,
        activeUserEmail,
      });
      return false;
    }

    // 3重チェック実行
    // 1. メールアドレス照合
    const dbEmail = String(userFromDb.adminEmail || '')
      .toLowerCase()
      .trim();
    const currentEmail = String(activeUserEmail).toLowerCase().trim();
    const isEmailMatched = dbEmail && currentEmail && dbEmail === currentEmail;

    // 2. ユーザーID照合（念のため）
    const isUserIdMatched = String(userFromDb.userId) === String(userId);

    // 3. アクティブ状態確認
    const isActive = Boolean(userFromDb.isActive);

    console.log('verifyAdminAccess: 3重チェック結果:', {
      isEmailMatched,
      isUserIdMatched,
      isActive,
      dbEmail,
      currentEmail,
    });

    // 3つの条件すべてが満たされた場合のみ認証成功
    if (isEmailMatched && isUserIdMatched && isActive) {
      console.log('✅ verifyAdminAccess: 認証成功', { userId, email: activeUserEmail });
      return true;
    } else {
      console.warn('❌ verifyAdminAccess: 認証失敗', {
        userId,
        activeUserEmail,
        failures: {
          email: !isEmailMatched,
          userId: !isUserIdMatched,
          active: !isActive,
        },
      });
      return false;
    }
  } catch (error) {
    console.error('[ERROR]', '❌ verifyAdminAccess: 認証処理エラー:', error.message);
    return false;
  }
}

/**
 * ユーザーの最終アクセス時刻のみを更新（設定は保護）
 * @param {string} userId - 更新対象のユーザーID
 */
function updateUserLastAccess(userId) {
  try {
    if (!userId) {
      console.warn('updateUserLastAccess: userIdが指定されていません');
      return;
    }

    const now = new Date().toISOString();
    console.log('最終アクセス時刻を更新:', userId, now);

    // lastAccessedAtフィールドのみを更新（他の設定は保護）
    updateUser(userId, { lastAccessedAt: now });
  } catch (error) {
    console.error('[ERROR]', 'updateUserLastAccess エラー:', error.message);
  }
}

/**
 * configJsonからsetupStatusを安全に取得
 * @param {string} configJsonString - JSONエンコードされた設定文字列
 * @returns {string} setupStatus ('pending', 'completed', 'error')
 */
function getSetupStatusFromConfig(configJsonString) {
  try {
    if (!configJsonString || configJsonString.trim() === '' || configJsonString === '{}') {
      return 'pending'; // 空の場合はセットアップ未完了とみなす
    }

    const config = JSON.parse(configJsonString);

    // setupStatusが明示的に設定されている場合はそれを使用
    if (config.setupStatus) {
      return config.setupStatus;
    }

    // setupStatusがない場合、他のフィールドから推測（循環参照回避）
    // Note: この推測ロジックは循環参照を避けるため、formUrlベースに変更
    if (config.formCreated === true && config.formUrl && config.formUrl.trim()) {
      return 'completed';
    }

    return 'pending';
  } catch (error) {
    console.warn('getSetupStatusFromConfig JSON解析エラー:', error.message);
    return 'pending'; // エラー時はセットアップ未完了とみなす
  }
}

// =============================================================================
// SECTION 2: マルチテナントデータ分離システム（元multiTenantSecurity.gs）
// =============================================================================

/**
 * マルチテナントセキュリティ管理クラス
 * ユーザー間のデータ分離を強制し、テナント境界を厳密に管理
 */
class MultiTenantSecurityManager {
  constructor() {
    this.tenantBoundaryEnforcement = true;
    this.strictDataIsolation = true;
    this.auditTrailEnabled = true;
  }

  /**
   * ユーザーIDベースのテナント境界検証
   * @param {string} requestUserId - リクエスト元ユーザーID
   * @param {string} targetUserId - 対象データのユーザーID
   * @param {string} operation - 実行しようとする操作
   * @returns {boolean} アクセス許可の可否
   */
  validateTenantBoundary(requestUserId, targetUserId, operation) {
    if (!requestUserId || !targetUserId) {
      this.logSecurityViolation('MISSING_USER_ID', { requestUserId, targetUserId, operation });
      return false;
    }

    // 同一テナント（ユーザー）内のアクセスは許可
    if (requestUserId === targetUserId) {
      this.logDataAccess('TENANT_BOUNDARY_VALID', { requestUserId, operation });
      return true;
    }

    // 管理者権限チェック（必要に応じて実装）
    const hasAdminAccess = this.checkAdminAccess(requestUserId, targetUserId, operation);
    if (hasAdminAccess) {
      this.logDataAccess('ADMIN_ACCESS_GRANTED', { requestUserId, targetUserId, operation });
      return true;
    }

    // テナント境界違反
    this.logSecurityViolation('TENANT_BOUNDARY_VIOLATION', {
      requestUserId,
      targetUserId,
      operation,
    });

    return false;
  }

  /**
   * データアクセスパターンの検証
   * @param {string} userId - ユーザーID
   * @param {string} dataType - データタイプ
   * @param {string} operation - 操作タイプ
   * @returns {boolean} アクセス許可の可否
   */
  validateDataAccessPattern(userId, dataType, operation) {
    const allowedPatterns = {
      user_data: ['read', 'write', 'delete'],
      user_config: ['read', 'write'],
      user_cache: ['read', 'write', 'delete'],
      shared_config: ['read'], // 共有設定は読み取りのみ
    };

    if (!allowedPatterns[dataType] || !allowedPatterns[dataType].includes(operation)) {
      this.logSecurityViolation('INVALID_ACCESS_PATTERN', { userId, dataType, operation });
      return false;
    }

    return true;
  }

  /**
   * キャッシュキーのテナント分離強制
   * @param {string} baseKey - ベースキー
   * @param {string} userId - ユーザーID
   * @param {string} context - コンテキスト情報
   * @returns {string} テナント分離されたキー
   */
  generateSecureKey(baseKey, userId, context = '') {
    if (!userId) {
      throw new Error('SECURITY_ERROR: ユーザーIDが必要です');
    }

    // ユーザーIDの暗号化ハッシュを使用してテナント分離を強化
    const userHash = this.generateUserHash(userId);
    const timestamp = Math.floor(Date.now() / 3600000); // 1時間単位

    let secureKey = `MT_${userHash}_${baseKey}`;
    if (context) {
      secureKey += `_${context}`;
    }

    // タイムベースの検証情報を追加
    secureKey += `_T${timestamp}`;

    return secureKey;
  }

  /**
   * セキュアなユーザーハッシュ生成
   * @param {string} userId - ユーザーID
   * @returns {string} ハッシュ値
   */
  generateUserHash(userId) {
    if (!userId) return null;

    // MD5ハッシュを使用してユーザーIDを匿名化
    const digest = Utilities.computeDigest(
      Utilities.DigestAlgorithm.MD5,
      userId + '_TENANT_SALT_2024',
      Utilities.Charset.UTF_8
    );

    return digest
      .map(function (byte) {
        return (byte + 256).toString(16).slice(-2);
      })
      .join('')
      .substring(0, 16); // 16文字に短縮
  }

  /**
   * 管理者権限チェック
   * @param {string} requestUserId - リクエスト元ユーザーID
   * @param {string} targetUserId - 対象ユーザーID
   * @param {string} operation - 操作
   * @returns {boolean} 管理者権限の有無
   */
  checkAdminAccess(requestUserId, targetUserId, operation) {
    // 管理者権限は現在は実装しない（必要に応じて実装）
    // 実装する場合は、管理者ロールの検証ロジックをここに追加
    return false;
  }

  /**
   * セキュリティ違反の記録
   * @param {string} violationType - 違反タイプ
   * @param {object} details - 詳細情報
   */
  logSecurityViolation(violationType, details) {
    if (!this.auditTrailEnabled) return;

    const logEntry = {
      timestamp: new Date().toISOString(),
      type: 'SECURITY_VIOLATION',
      violation: violationType,
      details: details,
      userAgent: User.email(),
    };

    console.error('[ERROR]', `🚨 セキュリティ違反: ${violationType}`, JSON.stringify(logEntry));

    // 重大な違反の場合は追加の通知を送信する可能性
    if (violationType === 'TENANT_BOUNDARY_VIOLATION') {
      this.handleCriticalSecurityViolation(logEntry);
    }
  }

  /**
   * データアクセスの記録
   * @param {string} accessType - アクセスタイプ
   * @param {object} details - 詳細情報
   */
  logDataAccess(accessType, details) {
    if (!this.auditTrailEnabled) return;

    const logEntry = {
      timestamp: new Date().toISOString(),
      type: 'DATA_ACCESS',
      access: accessType,
      details: details,
    };

    console.log(`🔒 データアクセス: ${accessType}`, JSON.stringify(logEntry));
  }

  /**
   * 重大なセキュリティ違反の処理
   * @param {object} logEntry - ログエントリ
   */
  handleCriticalSecurityViolation(logEntry) {
    // 重大な違反の場合の追加処理
    // 例：管理者への通知、一時的なアクセス制限など
    console.warn('🚨 重大なセキュリティ違反が発生しました', logEntry);

    // フューチャー実装: 通知システムとの連携
    // this.sendSecurityAlert(logEntry);
  }

  /**
   * テナント分離検証の統計情報取得
   * @returns {object} 統計情報
   */
  getSecurityStats() {
    return {
      tenantBoundaryEnforcement: this.tenantBoundaryEnforcement,
      strictDataIsolation: this.strictDataIsolation,
      auditTrailEnabled: this.auditTrailEnabled,
      timestamp: new Date().toISOString(),
    };
  }
}

/**
 * グローバルなマルチテナントセキュリティマネージャーインスタンス
 */
const multiTenantSecurity = new MultiTenantSecurityManager();

/**
 * テナント分離を強制するキャッシュ操作ラッパー
 * @param {string} operation - 操作タイプ（'get', 'set', 'delete'）
 * @param {string} baseKey - ベースキー
 * @param {string} userId - ユーザーID
 * @param {any} value - 設定値（set操作時）
 * @param {object} options - オプション
 * @returns {any} 結果
 */
function secureMultiTenantCacheOperation(operation, baseKey, userId, value = null, options = {}) {
  // テナント境界検証
  const currentUserId = User.email(); // 簡易実装
  if (!multiTenantSecurity.validateTenantBoundary(currentUserId, userId, `cache_${operation}`)) {
    throw new Error(`SECURITY_ERROR: テナント境界違反 - ${operation} operation denied`);
  }

  // データアクセスパターン検証
  if (!multiTenantSecurity.validateDataAccessPattern(userId, 'user_cache', operation)) {
    throw new Error(`SECURITY_ERROR: 不正なデータアクセスパターン - ${operation} denied`);
  }

  // セキュアキー生成
  const secureKey = multiTenantSecurity.generateSecureKey(baseKey, userId, options.context);

  // 回復力のある実行で操作実行
  try {
    switch (operation) {
      case 'get':
        return resilientCacheOperation(
          () => CacheService.getUserCache().get(secureKey),
          `SecureCache_Get_${baseKey}`,
          () => null // フォールバック
        );

      case 'set':
        const ttl = options.ttl || 300; // 5分デフォルト
        return resilientCacheOperation(
          () => CacheService.getUserCache().put(secureKey, JSON.stringify(value), ttl),
          `SecureCache_Set_${baseKey}`
        );

      case 'delete':
        return resilientCacheOperation(
          () => CacheService.getUserCache().remove(secureKey),
          `SecureCache_Delete_${baseKey}`
        );

      default:
        throw new Error(`SECURITY_ERROR: 不正な操作タイプ - ${operation}`);
    }
  } catch (error) {
    multiTenantSecurity.logSecurityViolation('CACHE_OPERATION_FAILED', {
      operation,
      baseKey,
      userId,
      error: error.message,
    });
    throw error;
  }
}

/**
 * セキュアなユーザー情報取得
 * @param {string} userId - ユーザーID
 * @param {object} options - オプション
 * @returns {object|null} ユーザー情報
 */
/**
 * @deprecated 統合実装：unifiedUserManager.getSecureUserInfo() を使用してください
 */
// getSecureUserInfo - 重複削除済み
// → unifiedUserManager.gsの実装を使用してください

/**
 * セキュアなユーザー設定情報取得
 * @param {string} userId - ユーザーID
 * @param {string} configKey - 設定キー
 * @param {object} options - オプション
 * @returns {any} 設定値
 */
// getSecureUserConfig - 重複削除済み
// → unifiedUserManager.gsの実装を使用してください

/**
 * マルチテナントセキュリティの健全性チェック
 * @returns {object} 健全性チェック結果
 */
// performMultiTenantHealthCheck - 削除済み（過度に複雑なため不要）
function performMultiTenantHealthCheck() {
  const results = {
    timestamp: new Date().toISOString(),
    securityManagerStatus: 'OK',
    tenantIsolationTest: 'OK',
    cacheSecurityTest: 'OK',
    issues: [],
  };

  try {
    // セキュリティマネージャーの状態確認
    const stats = multiTenantSecurity.getSecurityStats();
    if (!stats.tenantBoundaryEnforcement) {
      results.issues.push('テナント境界強制が無効化されています');
      results.securityManagerStatus = 'WARNING';
    }

    // テナント分離テスト
    const testUserId1 = 'test_user_1@example.com';
    const testUserId2 = 'test_user_2@example.com';

    const isolation1 = multiTenantSecurity.validateTenantBoundary(testUserId1, testUserId2, 'test');
    if (isolation1) {
      results.issues.push('テナント分離テストが失敗 - 異なるテナント間のアクセスが許可されました');
      results.tenantIsolationTest = 'FAILED';
    }

    const isolation2 = multiTenantSecurity.validateTenantBoundary(testUserId1, testUserId1, 'test');
    if (!isolation2) {
      results.issues.push('テナント分離テストが失敗 - 同一テナント内のアクセスが拒否されました');
      results.tenantIsolationTest = 'FAILED';
    }

    // キャッシュセキュリティテスト
    const secureKey1 = multiTenantSecurity.generateSecureKey('test', testUserId1);
    const secureKey2 = multiTenantSecurity.generateSecureKey('test', testUserId2);

    if (secureKey1 === secureKey2) {
      results.issues.push(
        'セキュアキー生成テストが失敗 - 異なるユーザーで同じキーが生成されました'
      );
      results.cacheSecurityTest = 'FAILED';
    }
  } catch (error) {
    results.issues.push(`健全性チェック中にエラー: ${error.message}`);
    results.securityManagerStatus = 'ERROR';
  }

  return results;
}

// =============================================================================
// SECTION 3: セキュリティヘルスチェック統合システム（元securityHealthCheck.gs）
// =============================================================================

/**
 * 統合セキュリティヘルスチェック実行
 * @returns {Promise<object>} 包括的なヘルスチェック結果
 */
function performComprehensiveSecurityHealthCheck() {
  const startTime = Date.now();

  const healthCheckResult = {
    timestamp: new Date().toISOString(),
    executionTime: 0,
    overallStatus: 'UNKNOWN',
    components: {},
    criticalIssues: [],
    warnings: [],
    recommendations: [],
  };

  try {
    console.log('🔒 統合セキュリティヘルスチェック開始');

    // 1. 統一秘密情報管理システム
    try {
      const secretManagerHealth = unifiedSecretManager.performHealthCheck();
      healthCheckResult.components.secretManager = secretManagerHealth;

      if (secretManagerHealth.issues.length > 0) {
        healthCheckResult.warnings.push(
          `秘密情報管理: ${secretManagerHealth.issues.length}件の問題`
        );
      }
    } catch (error) {
      healthCheckResult.components.secretManager = {
        status: 'ERROR',
        error: error.message,
      };
      healthCheckResult.criticalIssues.push(`秘密情報管理システムエラー: ${error.message}`);
    }

    // 2. マルチテナントセキュリティ
    try {
      const multiTenantHealth = performMultiTenantHealthCheck();
      healthCheckResult.components.multiTenantSecurity = multiTenantHealth;

      if (multiTenantHealth.issues.length > 0) {
        healthCheckResult.criticalIssues.push(
          `マルチテナントセキュリティ: ${multiTenantHealth.issues.length}件の問題`
        );
      }
    } catch (error) {
      healthCheckResult.components.multiTenantSecurity = {
        status: 'ERROR',
        error: error.message,
      };
      healthCheckResult.criticalIssues.push(`マルチテナントセキュリティエラー: ${error.message}`);
    }

    // 3. 統一バッチ処理システム
    try {
      const batchProcessorHealth = performUnifiedBatchHealthCheck();
      healthCheckResult.components.batchProcessor = batchProcessorHealth;

      if (batchProcessorHealth.issues.length > 0) {
        healthCheckResult.warnings.push(
          `バッチ処理: ${batchProcessorHealth.issues.length}件の問題`
        );
      }
    } catch (error) {
      healthCheckResult.components.batchProcessor = {
        status: 'ERROR',
        error: error.message,
      };
      healthCheckResult.warnings.push(`バッチ処理システムエラー: ${error.message}`);
    }

    // 4. 回復力のある実行機構
    try {
      const resilientExecutorStats = resilientExecutor.getStats();
      healthCheckResult.components.resilientExecutor = {
        status: 'OK',
        stats: resilientExecutorStats,
        circuitBreakerState: resilientExecutorStats.circuitBreakerState,
      };

      if (resilientExecutorStats.circuitBreakerState === 'OPEN') {
        healthCheckResult.criticalIssues.push('回復力のある実行機構: 回路ブレーカーがOPEN状態');
      }

      const successRate = parseFloat(resilientExecutorStats.successRate.replace('%', ''));
      if (successRate < 95) {
        healthCheckResult.warnings.push(
          `回復力のある実行機構: 成功率が低下 (${resilientExecutorStats.successRate})`
        );
      }
    } catch (error) {
      healthCheckResult.components.resilientExecutor = {
        status: 'ERROR',
        error: error.message,
      };
      healthCheckResult.warnings.push(`回復力のある実行機構エラー: ${error.message}`);
    }

    // 5. 認証システム
    try {
      const authHealth = checkAuthenticationHealth();
      healthCheckResult.components.authentication = authHealth;

      if (!authHealth.serviceAccountValid) {
        healthCheckResult.criticalIssues.push('認証システム: サービスアカウント認証情報が無効');
      }
    } catch (error) {
      healthCheckResult.components.authentication = {
        status: 'ERROR',
        error: error.message,
      };
      healthCheckResult.criticalIssues.push(`認証システムエラー: ${error.message}`);
    }

    // 6. データベース接続性
    try {
      const dbHealth = checkDatabaseHealth();
      healthCheckResult.components.database = dbHealth;

      if (!dbHealth.accessible) {
        healthCheckResult.criticalIssues.push('データベース: 接続不可');
      }
    } catch (error) {
      healthCheckResult.components.database = {
        status: 'ERROR',
        error: error.message,
      };
      healthCheckResult.criticalIssues.push(`データベース接続エラー: ${error.message}`);
    }

    // 総合ステータス判定
    if (healthCheckResult.criticalIssues.length > 0) {
      healthCheckResult.overallStatus = 'CRITICAL';
      healthCheckResult.recommendations.push('重要な問題があります。即座に対処してください。');
    } else if (healthCheckResult.warnings.length > 0) {
      healthCheckResult.overallStatus = 'WARNING';
      healthCheckResult.recommendations.push('警告があります。確認と対処を推奨します。');
    } else {
      healthCheckResult.overallStatus = 'HEALTHY';
      healthCheckResult.recommendations.push(
        '全てのセキュリティコンポーネントが正常に動作しています。'
      );
    }

    // 実行時間を記録
    healthCheckResult.executionTime = Date.now() - startTime;

    // 結果をログ出力
    if (healthCheckResult.overallStatus === 'CRITICAL') {
      console.error(
        '[ERROR]',
        '🚨 統合セキュリティヘルスチェック: 重要な問題が検出されました',
        healthCheckResult
      );
    } else if (healthCheckResult.overallStatus === 'WARNING') {
      console.warn('⚠️ 統合セキュリティヘルスチェック: 警告があります', healthCheckResult);
    } else {
      console.log('✅ 統合セキュリティヘルスチェック: 正常', healthCheckResult);
    }

    return healthCheckResult;
  } catch (error) {
    healthCheckResult.overallStatus = 'ERROR';
    healthCheckResult.criticalIssues.push(`ヘルスチェック実行エラー: ${error.message}`);
    healthCheckResult.executionTime = Date.now() - startTime;

    console.error('[ERROR]', '❌ 統合セキュリティヘルスチェック実行エラー:', error);
    return healthCheckResult;
  }
}

/**
 * 認証システムのヘルスチェック
 * @private
 */
function checkAuthenticationHealth() {
  const authHealth = {
    status: 'UNKNOWN',
    serviceAccountValid: false,
    tokenGenerationWorking: false,
    lastTokenGeneration: null,
    issues: [],
  };

  try {
    // サービスアカウント認証情報の検証
    const serviceAccountCreds = getSecureServiceAccountCreds();
    if (
      serviceAccountCreds &&
      serviceAccountCreds.client_email &&
      serviceAccountCreds.private_key
    ) {
      authHealth.serviceAccountValid = true;
    } else {
      authHealth.issues.push('サービスアカウント認証情報が不完全');
    }

    // トークン生成テスト
    try {
      const testToken = getServiceAccountTokenCached();
      if (testToken && testToken.length > 0) {
        authHealth.tokenGenerationWorking = true;
        authHealth.lastTokenGeneration = new Date().toISOString();
      }
    } catch (tokenError) {
      authHealth.issues.push(`トークン生成エラー: ${tokenError.message}`);
    }

    authHealth.status = authHealth.issues.length === 0 ? 'OK' : 'WARNING';
  } catch (error) {
    authHealth.status = 'ERROR';
    authHealth.issues.push(`認証ヘルスチェックエラー: ${error.message}`);
  }

  return authHealth;
}

/**
 * データベース接続性のヘルスチェック
 * @private
 */
function checkDatabaseHealth() {
  const dbHealth = {
    status: 'UNKNOWN',
    accessible: false,
    databaseId: null,
    sheetsServiceWorking: false,
    lastAccess: null,
    issues: [],
  };

  try {
    // データベースID取得
    const dbId = getSecureDatabaseId();
    if (dbId) {
      dbHealth.databaseId = dbId.substring(0, 10) + '...';
    } else {
      dbHealth.issues.push('データベースIDが設定されていません');
      dbHealth.status = 'ERROR';
      return dbHealth;
    }

    // Sheets Service接続テスト
    try {
      const service = getSheetsServiceCached();
      if (service) {
        dbHealth.sheetsServiceWorking = true;

        // 軽量なデータベースアクセステスト
        const testResponse = resilientExecutor.execute(
          async () => {
            return resilientUrlFetch(`${service.baseUrl}/${encodeURIComponent(dbId)}`, {
              method: 'GET',
              headers: {
                Authorization: `Bearer ${service.accessToken}`,
              },
            });
          },
          {
            name: 'DatabaseHealthCheck',
            idempotent: true,
          }
        );

        if (testResponse && testResponse.getResponseCode() === 200) {
          dbHealth.accessible = true;
          dbHealth.lastAccess = new Date().toISOString();
        } else {
          dbHealth.issues.push(
            `データベースアクセステスト失敗: ${testResponse ? testResponse.getResponseCode() : 'No response'}`
          );
        }
      }
    } catch (serviceError) {
      dbHealth.issues.push(`Sheets Service エラー: ${serviceError.message}`);
    }

    dbHealth.status = dbHealth.accessible ? 'OK' : 'ERROR';
  } catch (error) {
    dbHealth.status = 'ERROR';
    dbHealth.issues.push(`データベースヘルスチェックエラー: ${error.message}`);
  }

  return dbHealth;
}

/**
 * セキュリティメトリクスのサマリー取得
 * @returns {object} セキュリティメトリクス
 */
function getSecurityMetrics() {
  const metrics = {
    timestamp: new Date().toISOString(),
    resilientExecutor: null,
    secretManager: null,
    multiTenantSecurity: null,
    batchProcessor: null,
  };

  try {
    // 回復力のある実行機構のメトリクス
    if (typeof resilientExecutor !== 'undefined') {
      metrics.resilientExecutor = resilientExecutor.getStats();
    }

    // 統一バッチ処理のメトリクス
    if (typeof unifiedBatchProcessor !== 'undefined') {
      metrics.batchProcessor = unifiedBatchProcessor.getMetrics();
    }

    // マルチテナントセキュリティの統計
    if (typeof multiTenantSecurity !== 'undefined') {
      metrics.multiTenantSecurity = multiTenantSecurity.getSecurityStats();
    }

    // 秘密情報管理の監査ログサマリー
    if (typeof unifiedSecretManager !== 'undefined') {
      const auditLog = unifiedSecretManager.getAuditLog();
      metrics.secretManager = {
        auditLogEntries: auditLog.length,
        lastAccess: auditLog.length > 0 ? auditLog[auditLog.length - 1].timestamp : null,
      };
    }
  } catch (error) {
    console.warn('セキュリティメトリクス取得エラー:', error.message);
  }

  return metrics;
}

/**
 * 定期実行される関数（トリガーから呼び出し）
 */
function runScheduledSecurityHealthCheck() {
  try {
    const healthResult = performComprehensiveSecurityHealthCheck();

    // 重要な問題がある場合は管理者に通知
    if (healthResult.overallStatus === 'CRITICAL') {
      // フューチャー実装: 管理者への緊急通知
      console.error(
        '[ERROR]',
        '🚨 緊急: セキュリティヘルスチェックで重要な問題を検出',
        healthResult
      );
    }

    // 結果をログに記録
    console.log('🔒 定期セキュリティヘルスチェック完了', {
      status: healthResult.overallStatus,
      criticalIssues: healthResult.criticalIssues.length,
      warnings: healthResult.warnings.length,
      executionTime: healthResult.executionTime,
    });
  } catch (error) {
    console.error('[ERROR]', '定期セキュリティヘルスチェック実行エラー:', error.message);
  }
}

/**
 * サービスアカウントをスプレッドシートに編集者として共有
 * @param {string} spreadsheetId - スプレッドシートID
 */
function addServiceAccountToSpreadsheet(spreadsheetId) {
  try {
    const serviceAccountEmail = getServiceAccountEmail();
    if (serviceAccountEmail === 'メールアドレス取得エラー') {
      console.warn(
        'サービスアカウントのメールアドレスが取得できないため、スプレッドシートの共有をスキップします。'
      );
      return;
    }

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const permissions = spreadsheet.getEditors();
    const isAlreadyEditor = permissions.some((editor) => editor.getEmail() === serviceAccountEmail);

    if (!isAlreadyEditor) {
      spreadsheet.addEditor(serviceAccountEmail);
      console.log(
        `✅ サービスアカウント (${serviceAccountEmail}) をスプレッドシート (${spreadsheetId}) に編集者として追加しました。`
      );
    } else {
      console.log(
        `サービスアカウント (${serviceAccountEmail}) は既にスプレッドシート (${spreadsheetId}) の編集者です。`
      );
    }
  } catch (error) {
    console.error(
      '[ERROR]',
      `サービスアカウントをスプレッドシート (${spreadsheetId}) に共有中にエラーが発生しました: ${error.message}`
    );
    throw new Error(
      `サービスアカウントをスプレッドシートに共有できませんでした。手動で ${serviceAccountEmail} を編集者として追加してください。`
    );
  }
}

/**
 * スプレッドシートをサービスアカウントと共有
 * @param {string} spreadsheetId - スプレッドシートID
 */
function shareSpreadsheetWithServiceAccount(spreadsheetId) {
  try {
    const serviceAccountEmail = getServiceAccountEmail();
    if (serviceAccountEmail === 'メールアドレス取得エラー') {
      console.warn(
        'サービスアカウントのメールアドレスが取得できないため、スプレッドシートの共有をスキップします。'
      );
      return;
    }

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const permissions = spreadsheet.getEditors();
    const isAlreadyEditor = permissions.some((editor) => editor.getEmail() === serviceAccountEmail);

    if (!isAlreadyEditor) {
      spreadsheet.addEditor(serviceAccountEmail);
      console.log(
        `✅ サービスアカウント (${serviceAccountEmail}) をスプレッドシート (${spreadsheetId}) に編集者として追加しました。`
      );
    } else {
      console.log(
        `サービスアカウント (${serviceAccountEmail}) は既にスプレッドシート (${spreadsheetId}) の編集者です。`
      );
    }
  } catch (error) {
    console.error(
      '[ERROR]',
      `サービスアカウントをスプレッドシート (${spreadsheetId}) に共有中にエラーが発生しました: ${error.message}`
    );
    throw new Error(
      `サービスアカウントをスプレッドシートに共有できませんでした。手動で ${serviceAccountEmail} を編集者として追加してください。`
    );
  }
}

/**
 * ユーザーのアクセス権限を検証
 * @param {string} requestUserId - リクエスト元のユーザーID
 */
// verifyUserAccess は Core.gs の詳細実装にリダイレクト
// 統一APIアーキテクチャの一部として、最も完全な実装を使用

// invalidateUserCache() は unifiedCacheManager.gs に統合済み

/**
 * クリティカルな更新後にキャッシュを同期
 * @param {string} userId - ユーザーID
 * @param {string} email - ユーザーのメールアドレス
 * @param {string} oldSpreadsheetId - 変更前のスプレッドシートID
 * @param {string} newSpreadsheetId - 変更後のスプレッドシートID
 */
// synchronizeCacheAfterCriticalUpdate() は unifiedCacheManager.gs に統合済み

// clearDatabaseCache() は unifiedCacheManager.gs に統合済み

/**
 * ユーザー固有のキーを構築
 * @param {string} baseKey - ベースキー
 * @param {string} userId - ユーザーID
 * @param {string} sheetName - シート名 (オプション)
 * @returns {string} ユーザー固有のキー
 */
function buildUserScopedKey(baseKey, userId, sheetName = '') {
  if (!userId) {
    throw new Error('ユーザーIDが指定されていません');
  }
  let key = `${baseKey}_${userId}`;
  if (sheetName) {
    key += `_${sheetName}`;
  }
  return key;
}

// この重複関数は削除されました。L117のgetServiceAccountEmail()を使用してください。

/**
 * サービスアカウント認証情報を安全に取得
 * @returns {object} サービスアカウント認証情報
 */
function getSecureServiceAccountCreds() {
  try {
    const credsJson = PropertiesService.getScriptProperties().getProperty(
      SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS
    );
    if (!credsJson) {
      throw new Error('サービスアカウント認証情報が設定されていません。');
    }
    return JSON.parse(credsJson);
  } catch (e) {
    console.error('[ERROR]', 'getSecureServiceAccountCreds エラー: ' + e.message);
    throw new Error('サービスアカウント認証情報の取得に失敗しました。');
  }
}

/**
 * データベーススプレッドシートIDを安全に取得
 * @returns {string} データベーススプレッドシートID
 */
function getSecureDatabaseId() {
  try {
    const dbId = PropertiesService.getScriptProperties().getProperty(
      SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID
    );
    if (!dbId) {
      throw new Error('データベーススプレッドシートIDが設定されていません。');
    }
    return dbId;
  } catch (e) {
    console.error('[ERROR]', 'getSecureDatabaseId エラー: ' + e.message);
    throw new Error('データベーススプレッドシートIDの取得に失敗しました。');
  }
}

/**
 * ユーザーのメールアドレスからユーザーIDを取得
 * @param {string} email - ユーザーのメールアドレス
 * @returns {string} ユーザーID
 */
// getUserIdFromEmail() は unifiedUserManager.gs に統合済み

// getEmailFromUserId(), checkLoginStatus(), updateLoginStatus() は unifiedUserManager.gs に統合済み
