/**
 * @fileoverview SecurityService - 統一セキュリティサービス
 *
 * 責任範囲:
 * - 認証・認可管理
 * - 入力検証・サニタイズ
 * - セキュリティ監査・ログ
 * - Service Account管理
 */

/* global validateEmail, validateText, validateUrl, getUnifiedAccessLevel, findUserByEmail, findUserById, openSpreadsheet, updateUser, URL, getCurrentEmail */


// 認証・セッション管理

/**
 * Deploy user domain information retrieval
 * @returns {Object} Domain information and validation result
 */
function getDeployUserDomainInfo() {
  try {
    const email = getCurrentEmail();

    // Enhanced type validation for email
    if (!email || typeof email !== 'string' || email.trim() === '') {
      console.error('Authentication failed - invalid email:', typeof email, email);
      return {
        success: false,
        message: 'Authentication required - invalid email',
        domain: null,
        isValidDomain: false
      };
    }

    const domain = email.includes('@') ? email.split('@')[1] : 'unknown';
    const adminEmail = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL');
    const adminDomain = adminEmail ? adminEmail.split('@')[1] : null;
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



// 入力検証・サニタイズ

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

      // メールアドレス検証
      if (userData.email) {
        const emailValidation = validateEmail(userData.email);
        if (!emailValidation.isValid) {
          result.errors.push('無効なメールアドレス');
          result.isValid = false;
        } else {
          result.sanitizedData.email = emailValidation.sanitized;
        }
      }

      // テキストフィールド検証・サニタイズ
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

      // URL検証
      if (userData.url) {
        const urlValidation = validateUrl(userData.url);
        if (!urlValidation.isValid) {
          result.errors.push('無効なURL');
          result.isValid = false;
        } else {
          result.sanitizedData.url = urlValidation.sanitized;
        }
      }

      // ファイルサイズ制限チェック
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

// 🔒 アクセス制御・権限管理

/**
 * ユーザー権限確認
 * @param {string} userId - ユーザーID
 * @param {string} requiredLevel - 必要権限レベル
 * @returns {Object} 権限確認結果
 */
function checkSecurityUserPermission(userId, requiredLevel = 'authenticated_user') {
    try {
      const session = { email: Session.getActiveUser().getEmail() };
      const currentEmail = session.email;
      if (!currentEmail) {
        return {
          hasPermission: false,
          reason: 'セッションが無効です',
          currentLevel: 'none'
        };
      }

      // ユーザーIDが不明でも、認証済みであれば authenticated_user 要件は満たす
      if (!userId && requiredLevel === 'authenticated_user') {
        return {
          hasPermission: true,
          currentLevel: 'authenticated_user',
          requiredLevel,
          userEmail: currentEmail,
          timestamp: new Date().toISOString()
        };
      }

      // 統一されたアクセスレベル取得（email-based access control）
      const accessLevel = getUnifiedAccessLevel(currentEmail, userId);
      const hasPermission = compareSecurityAccessLevels(accessLevel, requiredLevel);

      return {
        hasPermission,
        currentLevel: accessLevel,
        requiredLevel,
        userEmail: currentEmail,
        timestamp: new Date().toISOString()
      };
    } catch (error) {
      console.error('SecurityService.checkUserPermission: エラー', error.message);
      return {
        hasPermission: false,
        reason: error.message,
        currentLevel: 'none'
      };
    }
}

/**
 * アクセスレベル比較
 * @param {string} currentLevel - 現在のレベル
 * @param {string} requiredLevel - 必要レベル
 * @returns {boolean} 権限があるかどうか
 */
function compareSecurityAccessLevels(currentLevel, requiredLevel) {
    const levelHierarchy = {
      'none': 0,
      'guest': 1,
      'authenticated_user': 2,
      'editor': 3,
      'administrator': 4
    };

    const currentScore = levelHierarchy[currentLevel] || 0;
    const requiredScore = levelHierarchy[requiredLevel] || 0;

    return currentScore >= requiredScore;
}

// セキュリティ監査・ログ

/**
 * セキュリティイベントログ
 * @param {Object} event - セキュリティイベント
 */
function logSecurityEvent(event) {
    try {
      const logEntry = {
        timestamp: new Date().toISOString(),
        sessionEmail: { email: Session.getActiveUser().getEmail() }.email,
        effectiveEmail: Session.getEffectiveUser().getEmail(),
        eventType: event.type || 'unknown',
        severity: event.severity || 'info',
        description: event.description || '',
        metadata: event.metadata || {},
        userAgent: event.userAgent || '',
        ipAddress: event.ipAddress || ''
      };

      // 重要度によってログレベル調整
      switch (event.severity) {
        case 'critical':
        case 'high':
          console.error('SecurityEvent:', logEntry);
          break;
        case 'medium':
          console.warn('SecurityEvent:', logEntry);
          break;
        default:
          console.info('SecurityEvent:', logEntry);
      }

      // 重要なログのみ永続化（PropertiesServiceで統一）
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
      
      // 古いログの削除（最新100件まで保持）
      cleanupOldSecurityLogs();
    } catch (error) {
      console.error('SecurityService.persistSecurityLog: エラー', error.message);
    }
}


// ユーティリティ・診断


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

      // アクセステスト - 段階的にチェック
      let spreadsheet;
      try {
        // Try normal permissions first, fall back to service account if needed
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

      // 名前とシート情報を取得
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

      // より具体的なエラーメッセージ
      let userMessage = 'スプレッドシートにアクセスできません。';
      if (error.message.includes('Permission') || error.message.includes('権限')) {
        userMessage = 'スプレッドシートへのアクセス権限がありません。編集権限を確認してください。';
      } else if (error.message.includes('not found') || error.message.includes('見つかりません')) {
        userMessage = 'スプレッドシートが見つかりません。URLが正しいか確認してください。';
      } else if (error.message.includes('timeout') || error.message.includes('タイムアウト')) {
        userMessage = 'スプレッドシートへのアクセスがタイムアウトしました。しばらく時間をおいて再試行してください。';
      }

      errorResponse.message = userMessage;

      // 絶対にnullを返さない
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

    // セキュリティログのキーを抽出
    const logKeys = Object.keys(allProps).filter(key => key.startsWith('security_log_'));

    if (logKeys.length > 100) {
      // タイムスタンプでソートして古いものから削除
      const sortedKeys = logKeys.sort((a, b) => {
        const timestampA = parseInt(a.split('_')[2], 10);
        const timestampB = parseInt(b.split('_')[2], 10);
        return timestampA - timestampB;
      });

      // 古いログを削除（最新100件を残す）
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

