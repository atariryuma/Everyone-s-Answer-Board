/**
 * @fileoverview SecurityService - 統一セキュリティサービス
 *
 * 🎯 責任範囲:
 * - 認証・認可管理
 * - 入力検証・サニタイズ
 * - セキュリティ監査・ログ
 * - Service Account管理
 *
 * 🔄 GAS Best Practices準拠:
 * - フラット関数構造 (Object.freeze削除)
 * - 直接的な関数エクスポート
 * - 単一責任原則の維持
 */

/* global ServiceFactory, validateEmail, validateUrl, getUserAccessLevel, Data */

/**
 * ServiceFactory統合初期化
 * SecurityService用Zero-Dependency実装
 * @returns {boolean} 初期化成功可否
 */
function initSecurityServiceZero() {
  return ServiceFactory.getUtils().initService('SecurityService');
}

// ===========================================
// 🔑 認証・セッション管理
// ===========================================

// Service account token generation function removed - was using user OAuth token impersonation

/**
 * トークン形式検証（セキュリティ強化）
 * @param {string} token - 検証対象トークン
 * @returns {boolean} 有効かどうか
 */
function validateTokenFormat(token) {
    if (!token || typeof token !== 'string') {
      return false;
    }

    // 基本的な形式チェック
    if (token.length < 20 || token.length > 4000) {
      return false;
    }

    // OAuth 2.0トークンの一般的な形式チェック
    if (!/^[A-Za-z0-9._-]+$/.test(token)) {
      return false;
    }

    // 明らかに無効な値の除外
    const invalidTokens = ['undefined', 'null', 'error', 'expired'];
    if (invalidTokens.includes(token.toLowerCase())) {
      return false;
    }

    return true;
}

/**
 * セッション状態検証
 * @returns {Object} セッション検証結果
 */
function validateSession() {
    try {
      const email = Session.getActiveUser().getEmail();
      const effectiveEmail = Session.getEffectiveUser().getEmail();

      return {
        isValid: !!email,
        userEmail: email,
        effectiveEmail,
        isImpersonated: email !== effectiveEmail,
        timestamp: new Date().toISOString()
      };
    } catch (error) {
      console.error('SecurityService.validateSession: エラー', error.message);
      return {
        isValid: false,
        error: error.message,
        timestamp: new Date().toISOString()
      };
    }
}

// ===========================================
// 🛡️ 入力検証・サニタイズ
// ===========================================

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
          const textValidation = validateSecureText(userData[field]);
          if (!textValidation.isValid) {
            result.errors.push(`${field}: ${textValidation.error}`);
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
      result.errors.push(`検証エラー: ${error.message}`);
      result.isValid = false;
    }

    return result;
}

/**
 * メールアドレス検証・サニタイズ
 * @param {string} email - メールアドレス
 * @returns {Object} 検証結果
 */

/**
 * テキスト検証・サニタイズ
 * @param {string} text - テキスト
 * @returns {Object} 検証結果
 */
function validateSecureText(text) {
    try {
      if (typeof text !== 'string') {
        return { isValid: false, error: 'テキストが必要です' };
      }

      const originalLength = text.length;
      let sanitized = text;
      let hasSecurityRisk = false;

      // 危険なパターンチェック
      const dangerousPatterns = [
        /<script/i,
        /javascript:/i,
        /on\w+\s*=/i,
        /<iframe/i,
        /<object/i,
        /<embed/i
      ];

      dangerousPatterns.forEach(pattern => {
        if (pattern.test(sanitized)) {
          hasSecurityRisk = true;
          sanitized = sanitized.replace(pattern, '[REMOVED]');
        }
      });

      // HTMLエンティティエスケープ
      sanitized = sanitized
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#x27;');

      // 長さ制限（8KB）
      if (sanitized.length > 8192) {
        sanitized = `${sanitized.substring(0, 8192)  }...`;
      }

      // 改行正規化
      sanitized = sanitized.replace(/\r\n/g, '\n').replace(/\r/g, '\n');

      return {
        isValid: true,
        sanitized,
        hasSecurityRisk,
        originalLength,
        sanitizedLength: sanitized.length,
        wasTruncated: sanitized.endsWith('...')
      };
    } catch (error) {
      return { isValid: false, error: error.message };
    }
}

/**
 * URL検証・サニタイズ
 * @param {string} url - URL
 * @returns {Object} 検証結果
 */

// ===========================================
// 🔒 アクセス制御・権限管理
// ===========================================

/**
 * ユーザー権限確認
 * @param {string} userId - ユーザーID
 * @param {string} requiredLevel - 必要権限レベル
 * @returns {Object} 権限確認結果
 */
function checkSecurityUserPermission(userId, requiredLevel = 'authenticated_user') {
    try {
      const currentEmail = Session.getActiveUser().getEmail();
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

      // UserServiceから権限レベル取得
      const accessLevel = getUserAccessLevel(userId);
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
      'system_admin': 3,
      'owner': 4
    };

    const currentScore = levelHierarchy[currentLevel] || 0;
    const requiredScore = levelHierarchy[requiredLevel] || 0;

    return currentScore >= requiredScore;
}

// ===========================================
// 📊 セキュリティ監査・ログ
// ===========================================

/**
 * セキュリティイベントログ
 * @param {Object} event - セキュリティイベント
 */
function logSecurityEvent(event) {
    try {
      const logEntry = {
        timestamp: new Date().toISOString(),
        sessionEmail: Session.getActiveUser().getEmail(),
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
          console.error('🚨 SecurityEvent:', logEntry);
          break;
        case 'medium':
          console.warn('⚠️ SecurityEvent:', logEntry);
          break;
        default:
          console.info('ℹ️ SecurityEvent:', logEntry);
      }

      // 永続化が必要な場合はPropertiesServiceを使用
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
      if (!initSecurityServiceZero()) {
        console.error('validateEmailAccess: ServiceFactory not available');
        return false;
      }
      const props = ServiceFactory.getProperties();
      const logKey = `security_log_${Date.now()}`;
      
      props.setProperty(logKey, JSON.stringify(logEntry));
      
      // 古いログの削除（最新100件まで保持）
      cleanupOldSecurityLogs();
    } catch (error) {
      console.error('SecurityService.persistSecurityLog: エラー', error.message);
    }
}

/**
 * 古いセキュリティログ削除
 */
function cleanupOldSecurityLogs() {
    try {
      if (!initSecurityServiceZero()) {
        console.error('validateEmailAccess: ServiceFactory not available');
        return false;
      }
      const props = ServiceFactory.getProperties();
      const allProps = props.getProperties();
      
      const securityLogs = Object.keys(allProps)
        .filter(key => key.startsWith('security_log_'))
        .sort()
        .reverse();

      // 100件を超える古いログを削除
      if (securityLogs.length > 100) {
        const logsToDelete = securityLogs.slice(100);
        logsToDelete.forEach(key => props.deleteProperty(key));
      }
    } catch (error) {
      console.error('SecurityService.cleanupOldSecurityLogs: エラー', error.message);
    }
}

// ===========================================
// 🔧 ユーティリティ・診断
// ===========================================


/**
 * スプレッドシートアクセス権限検証
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {Object} 検証結果
 */
function validateSpreadsheetAccess(spreadsheetId) {
    const started = Date.now();
    try {
      console.log('SecurityService', {
        operation: 'validateSpreadsheetAccess',
        phase: 'start',
        spreadsheetId: spreadsheetId ? `${spreadsheetId.substring(0, 10)  }...` : 'null'
      });

      if (!spreadsheetId) {
        const errorResponse = {
          success: false,
          message: 'スプレッドシートIDが指定されていません',
          sheets: [],
          executionTime: `${Date.now() - started}ms`
        };
        console.error('SecurityService', {
          operation: 'validateSpreadsheetAccess',
          error: 'Missing spreadsheetId',
          executionTime: errorResponse.executionTime
        });
        return errorResponse;
      }

      // アクセステスト - 段階的にチェック
      let spreadsheet;
      try {
        console.log('SecurityService', { operation: 'Data.open', phase: 'start' });
        const { spreadsheet: spreadsheetFromData } = Data.open(spreadsheetId);
        spreadsheet = spreadsheetFromData;
        console.log('SecurityService', { operation: 'Data.open', phase: 'success' });
      } catch (openError) {
        const errorResponse = {
          success: false,
          message: `スプレッドシートへのアクセスに失敗: ${openError.message}`,
          sheets: [],
          error: openError.message,
          executionTime: `${Date.now() - started}ms`
        };
        console.error('SecurityService', {
          operation: 'SpreadsheetApp.openById',
          error: openError.message,
          executionTime: errorResponse.executionTime
        });
        return errorResponse;
      }

      // 名前とシート情報を取得
      let name, sheets;
      try {
        console.log('SecurityService', { operation: 'spreadsheet.getName', phase: 'start' });
        name = spreadsheet.getName();

        console.log('SecurityService', { operation: 'spreadsheet.getSheets', phase: 'start' });
        sheets = spreadsheet.getSheets().map(sheet => ({
          name: sheet.getName(),
          index: sheet.getIndex()
        }));
        console.log('SecurityService', {
          operation: 'spreadsheet metadata',
          phase: 'success',
          sheetsCount: sheets.length
        });
      } catch (metaError) {
        const errorResponse = {
          success: false,
          message: `スプレッドシート情報の取得に失敗: ${metaError.message}`,
          sheets: [],
          error: metaError.message,
          executionTime: `${Date.now() - started}ms`
        };
        console.error('SecurityService', {
          operation: 'spreadsheet metadata',
          error: metaError.message,
          executionTime: errorResponse.executionTime
        });
        return errorResponse;
      }

      const result = {
        success: true,
        name,
        sheets,
        message: 'アクセス権限が確認できました',
        executionTime: `${Date.now() - started}ms`
      };

      console.log('SecurityService', {
        operation: 'validateSpreadsheetAccess',
        spreadsheetName: name,
        sheetsCount: sheets.length,
        executionTime: result.executionTime
      });

      return result;

    } catch (error) {
      const errorResponse = {
        success: false,
        message: `予期しないエラー: ${error.message}`,
        sheets: [],
        error: error.message,
        executionTime: `${Date.now() - started}ms`
      };

      console.error('SecurityService', {
        operation: 'validateSpreadsheetAccess',
        error: error.message,
        stack: error.stack,
        executionTime: errorResponse.executionTime
      });

      console.error('SecurityService.validateSpreadsheetAccess 予期しないエラー:', {
        error: error.message,
        stack: error.stack,
        spreadsheetId: spreadsheetId ? `${spreadsheetId.substring(0, 10)  }...` : 'null'
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
 * セキュリティ設定推奨事項
 * @returns {Array} 推奨事項リスト
 */
function getSecurityRecommendations() {
    const recommendations = [];

    try {
      // 基本的なセキュリティチェック
      const session = validateSession();
      if (!session) {
        recommendations.push({
          priority: 'high',
          category: 'authentication',
          message: 'セッションの再認証が必要です',
          action: 'ログインを確認してください'
        });
      }

      // トークンの有効性確認
      // Service account token access removed - impersonation detected
      const token = null; // Placeholder for removed impersonation function
      if (!token) {
        recommendations.push({
          priority: 'medium',
          category: 'authorization',
          message: 'Service Accountトークンの更新が必要です',
          action: 'トークンを再生成してください'
        });
      }

      // デフォルト推奨事項
      recommendations.push({
        priority: 'low',
        category: 'general',
        message: '定期的なセキュリティ監査を実施してください',
        action: '月次でのセキュリティレビューを設定'
      });

    } catch (error) {
      console.error('SecurityService.getSecurityRecommendations: エラー', error.message);
      recommendations.push({
        priority: 'high',
        category: 'error',
        message: 'セキュリティ診断でエラーが発生しました',
        action: 'システム管理者に連絡してください'
      });
    }

    return recommendations;
}
