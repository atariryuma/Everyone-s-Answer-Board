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
      success: false,
      allowed: allowIfEmailMissing,
      message: allowIfEmailMissing ? 'No email provided, access allowed by configuration' : 'Authentication required',
      reason: 'missing_email',
      userDomain: null,
      adminDomain: null
    };
  }

  const normalizedEmail = email.trim();
  const userDomain = extractDomainSafely(normalizedEmail);
  if (!userDomain) {
    return {
      success: false,
      allowed: false,
      message: 'Invalid email format',
      reason: 'invalid_email_format',
      userDomain: null,
      adminDomain: null
    };
  }

  const adminEmail = getCachedProperty('ADMIN_EMAIL');
  const adminDomain = adminEmail ? extractDomainSafely(adminEmail) : null;
  if (!adminDomain) {
    return {
      success: true,
      allowed: allowIfAdminUnconfigured,
      message: allowIfAdminUnconfigured ? 'Admin domain not configured, access allowed' : 'Admin domain not configured',
      reason: 'admin_domain_unconfigured',
      userDomain,
      adminDomain: null
    };
  }

  const allowed = userDomain === adminDomain;
  return {
    success: true,
    allowed,
    message: allowed ? 'Domain access allowed' : 'Domain mismatch',
    reason: allowed ? 'domain_match' : 'domain_mismatch',
    userDomain,
    adminDomain
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
/**
 * セキュリティイベントログ
 * @param {Object} event - セキュリティイベント
 */
function logSecurityEvent(event) {
    try {
      // ✅ SECURITY FIX: getActiveUser() のみ使用（getEffectiveUser() は権限昇格リスクあり）
      const sessionEmail = Session.getActiveUser().getEmail();
      const logEntry = {
        timestamp: new Date().toISOString(),
        sessionEmail,
        eventType: event.type || 'unknown',
        severity: event.severity || 'info',
        description: event.description || '',
        metadata: event.metadata || {},
        userAgent: event.userAgent || '',
        ipAddress: event.ipAddress || ''
      };

      switch (event.severity) {
        case 'critical':
        case 'high':
          console.error('SecurityEvent:', logEntry);
          break;
        case 'medium':
          console.warn('SecurityEvent:', logEntry);
          break;
        default:
      }

      if (event.severity === 'critical' || event.severity === 'high') {
        persistSecurityLog(logEntry);
      }

    } catch (error) {
      console.error('SecurityService.logSecurityEvent: ログ記録エラー', error.message);
    }
}

/**
 * セキュリティログ永続化
 * @param {Object} logEntry - ログエントリ
 */
function persistSecurityLog(logEntry) {
  try {
    const props = PropertiesService.getScriptProperties();
    const allProps = props.getProperties();

    const MAX_LOG_ENTRIES = 100;
    const MAX_TOTAL_STORAGE_CHARS = 450000;
    const MAX_LOG_ENTRY_CHARS = 8000;

    const logKeys = Object.keys(allProps)
      .filter((key) => key.startsWith('security_log_'))
      .sort((a, b) => {
        const ta = parseInt(a.split('_')[2], 10) || 0;
        const tb = parseInt(b.split('_')[2], 10) || 0;
        return ta - tb;
      });

    let totalStorageChars = 0;
    Object.entries(allProps).forEach(([key, value]) => {
      totalStorageChars += key.length + String(value || '').length;
    });

    const serializedEntryRaw = JSON.stringify(logEntry || {});
    const serializedEntry = serializedEntryRaw.length > MAX_LOG_ENTRY_CHARS
      ? serializedEntryRaw.substring(0, MAX_LOG_ENTRY_CHARS)
      : serializedEntryRaw;

    const keysToDelete = [];
    const mutableLogKeys = [...logKeys];
    while (
      mutableLogKeys.length >= MAX_LOG_ENTRIES ||
      totalStorageChars + serializedEntry.length > MAX_TOTAL_STORAGE_CHARS
    ) {
      const oldestKey = mutableLogKeys.shift();
      if (!oldestKey) break;
      keysToDelete.push(oldestKey);
      totalStorageChars -= oldestKey.length + String(allProps[oldestKey] || '').length;
    }

    keysToDelete.forEach((key) => {
      try {
        props.deleteProperty(key);
      } catch (deleteError) {
        console.warn(`SecurityService: Failed to delete log ${key}:`, deleteError.message);
      }
    });

    if (totalStorageChars + serializedEntry.length > MAX_TOTAL_STORAGE_CHARS) {
      console.warn('SecurityService.persistSecurityLog: Skip due to storage quota safety guard');
      return;
    }

    const logKey = `security_log_${Date.now()}`;
    props.setProperty(logKey, serializedEntry);
  } catch (error) {
    console.error('SecurityService.persistSecurityLog: エラー', error.message);
  }
}




/**
 * スプレッドシートアクセス権限検証
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {Object} 検証結果
 */
function validateSpreadsheetAccess(spreadsheetId) {
    const started = Date.now();
    try {

      if (!spreadsheetId) {
        const errorResponse = {
          success: false,
          message: 'スプレッドシートIDが指定されていません',
          sheets: [],
          executionTime: `${Date.now() - started}ms`
        };
        return errorResponse;
      }

      let spreadsheet;
      try {
        const { spreadsheet: spreadsheetFromData } = openSpreadsheet(spreadsheetId, { useServiceAccount: false });
        spreadsheet = spreadsheetFromData;
      } catch (openError) {
        const errorResponse = {
          success: false,
          message: (openError && openError.message) ? `スプレッドシートへのアクセスに失敗: ${openError.message}` : 'スプレッドシートへのアクセスに失敗: 詳細不明',
          sheets: [],
          error: openError.message,
          executionTime: `${Date.now() - started}ms`
        };
        return errorResponse;
      }

      let name, sheets;
      try {
        name = spreadsheet.getName();

        sheets = spreadsheet.getSheets().map(sheet => ({
          name: sheet.getName(),
          index: sheet.getIndex()
        }));
      } catch (metaError) {
        const errorResponse = {
          success: false,
          message: (metaError && metaError.message) ? `スプレッドシート情報の取得に失敗: ${metaError.message}` : 'スプレッドシート情報の取得に失敗: 詳細不明',
          sheets: [],
          error: metaError.message,
          executionTime: `${Date.now() - started}ms`
        };
        return errorResponse;
      }

      const result = {
        success: true,
        name,
        sheets,
        message: 'アクセス権限が確認できました',
        executionTime: `${Date.now() - started}ms`
      };


      return result;

    } catch (error) {
      const errorResponse = {
        success: false,
        message: (error && error.message) ? `予期しないエラー: ${error.message}` : '予期しないエラー: 詳細不明',
        sheets: [],
        error: error.message,
        executionTime: `${Date.now() - started}ms`
      };


      console.error('SecurityService.validateSpreadsheetAccess 予期しないエラー:', {
        error: error.message,
        stack: error.stack,
        spreadsheetId: spreadsheetId ? `${spreadsheetId.substring(0, 10)}...` : 'null'
      });

      let userMessage = 'スプレッドシートにアクセスできません。';
      if (error.message.includes('Permission') || error.message.includes('権限')) {
        userMessage = 'スプレッドシートへのアクセス権限がありません。編集権限を確認してください。';
      } else if (error.message.includes('not found') || error.message.includes('見つかりません')) {
        userMessage = 'スプレッドシートが見つかりません。URLが正しいか確認してください。';
      } else if (error.message.includes('timeout') || error.message.includes('タイムアウト')) {
        userMessage = 'スプレッドシートへのアクセスがタイムアウトしました。しばらく時間をおいて再試行してください。';
      }

      errorResponse.message = userMessage;

      return errorResponse;
    }
}

/**
 * 古いセキュリティログのクリーンアップ
 * PropertiesServiceで統一管理（最新100件まで保持）
 */
function cleanupOldSecurityLogs() {
  try {
    const props = PropertiesService.getScriptProperties();
    const allProps = props.getProperties();

    const logKeys = Object.keys(allProps).filter(key => key.startsWith('security_log_'));

    if (logKeys.length > 100) {
      const sortedKeys = logKeys.sort((a, b) => {
        const timestampA = parseInt(a.split('_')[2], 10);
        const timestampB = parseInt(b.split('_')[2], 10);
        return timestampA - timestampB;
      });

      const keysToDelete = sortedKeys.slice(0, -100);
      keysToDelete.forEach(key => {
        try {
          props.deleteProperty(key);
        } catch (deleteError) {
          console.warn(`SecurityService: Failed to delete log ${key}:`, deleteError.message);
        }
      });

    }
  } catch (error) {
    console.warn('SecurityService.cleanupOldSecurityLogs: Cleanup failed:', error.message);
  }
}
