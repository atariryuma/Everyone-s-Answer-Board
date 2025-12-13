/**
 * @fileoverview SecurityService - 統一セキュリティサービス
 *
 * 責任範囲:
 * - 認証・認可管理
 * - 入力検証・サニタイズ
 * - セキュリティ監査・ログ
 * - Service Account管理
 */

/* global validateEmail, validateText, validateUrl, getUnifiedAccessLevel, findUserByEmail, findUserById, openSpreadsheet, updateUser, URL, getCurrentEmail, getCachedProperty */



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
 * Deploy user domain information retrieval
 * @returns {Object} Domain information and validation result
 */
function getDeployUserDomainInfo() {
  try {
    const email = getCurrentEmail();

    if (!email || typeof email !== 'string' || email.trim() === '') {
      console.error('Authentication failed - invalid email:', typeof email, email);
      return {
        success: false,
        message: 'Authentication required - invalid email',
        domain: null,
        isValidDomain: false
      };
    }

    const domain = extractDomainSafely(email);
    if (!domain) {
      console.warn('getDeployUserDomainInfo: Invalid email format:', email);
      return {
        success: false,
        message: 'Invalid email format',
        domain: null,
        isValidDomain: false
      };
    }

    const adminEmail = getCachedProperty('ADMIN_EMAIL');
    const adminDomain = adminEmail ? extractDomainSafely(adminEmail) : null;
    const isValidDomain = adminDomain ? domain === adminDomain : true;

    return {
      success: true,
      domain,
      userEmail: email,
      userDomain: domain,
      adminDomain,
      isValidDomain,
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
 * ユーザーデータ総合検証
 * @param {Object} userData - ユーザーデータ
 * @returns {Object} 検証結果
 */
function validateUserData(userData) {
    const result = {
      isValid: true,
      errors: [],
      sanitizedData: {},
      securityWarnings: []
    };

    try {
      if (!userData || typeof userData !== 'object') {
        result.errors.push('無効なデータ形式');
        result.isValid = false;
        return result;
      }

      if (userData.email) {
        const emailValidation = validateEmail(userData.email);
        if (!emailValidation.isValid) {
          result.errors.push('無効なメールアドレス');
          result.isValid = false;
        } else {
          result.sanitizedData.email = emailValidation.sanitized;
        }
      }

      ['answer', 'reason', 'name', 'className'].forEach(field => {
        if (userData[field]) {
          const textValidation = validateText(userData[field]);
          if (!textValidation.isValid) {
            if (field && textValidation && textValidation.error) {
              result.errors.push(`${field}: ${textValidation.error}`);
            } else {
              result.errors.push('Validation error: 詳細不明');
            }
            result.isValid = false;
          } else {
            result.sanitizedData[field] = textValidation.sanitized;
            if (textValidation.hasSecurityRisk) {
              result.securityWarnings.push(`${field}: セキュリティリスクを検出`);
            }
          }
        }
      });

      if (userData.url) {
        const urlValidation = validateUrl(userData.url);
        if (!urlValidation.isValid) {
          result.errors.push('無効なURL');
          result.isValid = false;
        } else {
          result.sanitizedData.url = urlValidation.sanitized;
        }
      }

      const dataSize = JSON.stringify(result.sanitizedData).length;
      if (dataSize > 32000) {
        result.errors.push('データサイズが制限を超えています');
        result.isValid = false;
      }

    } catch (error) {
      if (error && error.message) {
        result.errors.push(`検証エラー: ${error.message}`);
      } else {
        result.errors.push('検証エラー: 詳細不明');
      }
      result.isValid = false;
    }

    return result;
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
      const logKey = `security_log_${Date.now()}`;
      
      props.setProperty(logKey, JSON.stringify(logEntry));
      
      cleanupOldSecurityLogs();
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

