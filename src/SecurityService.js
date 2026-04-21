/**
 * @fileoverview SecurityService - 統一セキュリティサービス
 *
 * 責任範囲:
 * - 認証・認可管理
 * - 入力検証・サニタイズ
 * - セキュリティ監査・ログ
 * - Service Account管理
 */

/* global openSpreadsheet, getCurrentEmail, getCachedProperty, hasCoreSystemProps */

/**
 * 安全にメールアドレスからドメインを抽出する
 * @param {string} email - メールアドレス
 * @returns {string|null} ドメイン（抽出失敗時はnull）
 */
function extractDomainSafely(email) {
  if (!email || typeof email !== 'string') {
    return null;
  }

  try {
    const parsed = email.match(/^[^@]+@([^@]+)$/);
    return parsed ? parsed[1].toLowerCase() : null;
  } catch (error) {
    console.warn('extractDomainSafely: Failed to extract domain:', error.message);
    return null;
  }
}

/**
 * システムセットアップ完了状態を判定
 * @returns {boolean} セットアップ完了状態
 */
function isSystemSetupComplete() {
  try {
    if (typeof hasCoreSystemProps === 'function') {
      return hasCoreSystemProps();
    }

    const props = PropertiesService.getScriptProperties();
    const hasAdmin = !!props.getProperty('ADMIN_EMAIL');
    const hasDb = !!props.getProperty('DATABASE_SPREADSHEET_ID');
    const hasCreds = !!props.getProperty('SERVICE_ACCOUNT_CREDS');
    return hasAdmin && hasDb && hasCreds;
  } catch (error) {
    console.warn('isSystemSetupComplete: Failed to evaluate setup status:', error.message);
    return false;
  }
}

/**
 * ドメイン制限を適用すべきか判定
 * @returns {boolean} ドメイン制限を適用する場合 true
 */
function shouldEnforceDomainRestrictions() {
  return isSystemSetupComplete();
}

/**
 * メールアドレスのドメインアクセスを検証
 * @param {string|null} email - 検証対象メール
 * @param {Object} options - オプション
 * @returns {Object} 検証結果
 */
function validateDomainAccess(email, options = {}) {
  const allowIfAdminUnconfigured = options.allowIfAdminUnconfigured !== false;
  const allowIfEmailMissing = options.allowIfEmailMissing === true;

  if (!email || typeof email !== 'string' || !email.trim()) {
    return {
      allowed: allowIfEmailMissing,
      message: allowIfEmailMissing ? 'No email provided, access allowed by configuration' : 'Authentication required',
      reason: 'missing_email'
    };
  }

  const normalizedEmail = email.trim();
  const userDomain = extractDomainSafely(normalizedEmail);
  if (!userDomain) {
    return {
      allowed: false,
      message: 'Invalid email format',
      reason: 'invalid_email_format'
    };
  }

  const adminEmail = getCachedProperty('ADMIN_EMAIL');
  const adminDomain = adminEmail ? extractDomainSafely(adminEmail) : null;
  if (!adminDomain) {
    return {
      allowed: allowIfAdminUnconfigured,
      message: allowIfAdminUnconfigured ? 'Admin domain not configured, access allowed' : 'Admin domain not configured',
      reason: 'admin_domain_unconfigured'
    };
  }

  const allowed = userDomain === adminDomain;
  return {
    allowed,
    userDomain,
    adminDomain,
    message: allowed ? 'Domain access allowed' : 'Domain mismatch',
    reason: allowed ? 'domain_match' : 'domain_mismatch'
  };
}

/**
 * Deploy user domain information retrieval
 * @returns {Object} Domain information and validation result
 */
function getDeployUserDomainInfo() {
  try {
    const email = getCurrentEmail();
    const accessResult = validateDomainAccess(email, {
      allowIfAdminUnconfigured: true,
      allowIfEmailMissing: false
    });

    return {
      success: accessResult.success,
      domain: accessResult.userDomain,
      userEmail: email,
      userDomain: accessResult.userDomain,
      adminDomain: accessResult.adminDomain,
      isValidDomain: accessResult.allowed,
      reason: accessResult.reason,
      message: accessResult.message,
      timestamp: new Date().toISOString()
    };
  } catch (error) {
    console.error('SecurityService.getDeployUserDomainInfo ERROR:', error.message);
    return {
      success: false,
      message: error.message,
      domain: null,
      isValidDomain: false
    };
  }
}
