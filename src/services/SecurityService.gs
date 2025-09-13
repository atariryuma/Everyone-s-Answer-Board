/**
 * @fileoverview SecurityService - 統一セキュリティサービス
 * 
 * 🎯 責任範囲:
 * - 認証・認可管理
 * - 入力検証・サニタイズ
 * - セキュリティ監査・ログ
 * - Service Account管理
 * 
 * 🔄 置き換え対象:
 * - auth.gs（認証機能）
 * - security.gs（セキュリティ機能）
 * - 分散しているバリデーション機能
 */

/* global PROPS_KEYS, AppCacheService, UserService, CONSTANTS, DB */

/**
 * SecurityService - 統一セキュリティサービス
 * ゼロトラスト原則に基づく多層防御システム
 */
const SecurityService = Object.freeze({

  // ===========================================
  // 🔑 認証・セッション管理
  // ===========================================

  /**
   * Service Accountトークン取得（セキュリティ強化版）
   * @returns {string|null} アクセストークン
   */
  getServiceAccountToken() {
    const cacheKey = 'SA_TOKEN_CACHE';
    
    try {
      // キャッシュから取得試行
      const cached = CacheService.getScriptCache().get(cacheKey);
      if (cached) {
        // セキュリティ：トークン有効性の簡易検証
        if (this.validateTokenFormat(cached)) {
          return cached;
        } else {
          // 無効なトークンをクリア
          CacheService.getScriptCache().remove(cacheKey);
        }
      }

      // 新規トークン生成
      const newToken = this.generateServiceAccountToken();
      if (!newToken) {
        console.error('SecurityService.getServiceAccountToken: トークン生成失敗（詳細はログなし）');
        return null;
      }

      // セキュリティ：トークン形式検証
      if (!this.validateTokenFormat(newToken)) {
        console.error('SecurityService.getServiceAccountToken: 生成されたトークンが無効な形式');
        return null;
      }

      // 1時間キャッシュ（短時間で自動失効）
      CacheService.getScriptCache().put(cacheKey, newToken, 3600);
      
      console.info('SecurityService.getServiceAccountToken: トークン生成・キャッシュ完了（詳細省略）');
      return newToken;
    } catch (error) {
      // セキュリティ：エラー詳細をログに出力しない
      console.error('SecurityService.getServiceAccountToken: トークン取得処理エラー');
      return null;
    }
  },

  /**
   * Service Accountトークン生成（セキュリティ強化版）
   * @returns {string|null} 生成されたトークン
   */
  generateServiceAccountToken() {
    try {
      // GAS標準のOAuthトークン取得
      const token = ScriptApp.getOAuthToken();
      if (!token) {
        throw new Error('OAuthトークンの取得に失敗');
      }

      // セキュリティ：生成されたトークンの最小検証
      if (token.length < 20) {
        throw new Error('生成されたトークンが短すぎます');
      }

      return token;
    } catch (error) {
      // セキュリティ：具体的なエラー内容をログに出力しない
      console.error('SecurityService.generateServiceAccountToken: トークン生成処理エラー');
      return null;
    }
  },

  /**
   * トークン形式検証（セキュリティ強化）
   * @param {string} token - 検証対象トークン
   * @returns {boolean} 有効かどうか
   */
  validateTokenFormat(token) {
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
  },

  /**
   * セッション状態検証
   * @returns {Object} セッション検証結果
   */
  validateSession() {
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
  },

  // ===========================================
  // 🛡️ 入力検証・サニタイズ
  // ===========================================

  /**
   * ユーザーデータ総合検証
   * @param {Object} userData - ユーザーデータ
   * @returns {Object} 検証結果
   */
  validateUserData(userData) {
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
        const emailValidation = this.validateEmail(userData.email);
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
          const textValidation = this.validateAndSanitizeText(userData[field]);
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
        const urlValidation = this.validateUrl(userData.url);
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
  },

  /**
   * メールアドレス検証・サニタイズ
   * @param {string} email - メールアドレス
   * @returns {Object} 検証結果
   */
  validateEmail(email) {
    try {
      if (!email || typeof email !== 'string') {
        return { isValid: false, error: 'メールアドレスが必要です' };
      }

      // 基本形式チェック
      const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
      if (!emailRegex.test(email)) {
        return { isValid: false, error: '無効なメールアドレス形式' };
      }

      // 危険文字除去
      const sanitized = email.toLowerCase().trim();

      // 長さ制限
      if (sanitized.length > 254) {
        return { isValid: false, error: 'メールアドレスが長すぎます' };
      }

      return {
        isValid: true,
        sanitized,
        originalLength: email.length,
        sanitizedLength: sanitized.length
      };
    } catch (error) {
      return { isValid: false, error: error.message };
    }
  },

  /**
   * テキスト検証・サニタイズ
   * @param {string} text - テキスト
   * @returns {Object} 検証結果
   */
  validateAndSanitizeText(text) {
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
  },

  /**
   * URL検証・サニタイズ
   * @param {string} url - URL
   * @returns {Object} 検証結果
   */
  validateUrl(url) {
    try {
      if (!url || typeof url !== 'string') {
        return { isValid: false, error: 'URLが必要です' };
      }

      // 許可されたプロトコル
      const allowedProtocols = ['https:', 'http:'];
      
      // 許可されたドメイン（Google関連のみ）
      const allowedDomains = [
        'docs.google.com',
        'forms.gle',
        'script.google.com',
        'drive.google.com'
      ];

      // Basic URL validation for GAS environment
      if (!url.startsWith('http://') && !url.startsWith('https://')) {
        return { isValid: false, error: '無効なURL形式' };
      }
      
      // Extract protocol and hostname for validation
      let protocol, hostname;
      try {
        const urlMatch = url.match(/^(https?):\/\/([^/]+)/);
        if (!urlMatch) {
          return { isValid: false, error: '無効なURL形式' };
        }
        protocol = `${urlMatch[1]  }:`; // 'http:' or 'https:'
        hostname = urlMatch[2];
      } catch (error) {
        return { isValid: false, error: '無効なURL形式' };
      }

      // プロトコルチェック
      if (!allowedProtocols.includes(protocol)) {
        return { isValid: false, error: '許可されていないプロトコル' };
      }

      // ドメインチェック
      const isAllowedDomain = allowedDomains.some(domain => 
        hostname === domain || hostname.endsWith(`.${domain}`)
      );

      if (!isAllowedDomain) {
        return { isValid: false, error: '許可されていないドメイン' };
      }

      return {
        isValid: true,
        sanitized: url.trim(),
        protocol,
        hostname
      };
    } catch (error) {
      return { isValid: false, error: error.message };
    }
  },

  // ===========================================
  // 🔒 アクセス制御・権限管理
  // ===========================================

  /**
   * ユーザー権限確認
   * @param {string} userId - ユーザーID
   * @param {string} requiredLevel - 必要権限レベル
   * @returns {Object} 権限確認結果
   */
  checkUserPermission(userId, requiredLevel = 'authenticated_user') {
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
      const accessLevel = UserService.getAccessLevel(userId);
      const hasPermission = this.compareAccessLevels(accessLevel, requiredLevel);

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
  },

  /**
   * アクセスレベル比較
   * @param {string} currentLevel - 現在のレベル
   * @param {string} requiredLevel - 必要レベル
   * @returns {boolean} 権限があるかどうか
   */
  compareAccessLevels(currentLevel, requiredLevel) {
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
  },

  // ===========================================
  // 📊 セキュリティ監査・ログ
  // ===========================================

  /**
   * セキュリティイベントログ
   * @param {Object} event - セキュリティイベント
   */
  logSecurityEvent(event) {
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
        this.persistSecurityLog(logEntry);
      }

    } catch (error) {
      console.error('SecurityService.logSecurityEvent: ログ記録エラー', error.message);
    }
  },

  /**
   * セキュリティログ永続化
   * @param {Object} logEntry - ログエントリ
   */
  persistSecurityLog(logEntry) {
    try {
      const props = PropertiesService.getScriptProperties();
      const logKey = `security_log_${Date.now()}`;
      
      props.setProperty(logKey, JSON.stringify(logEntry));
      
      // 古いログの削除（最新100件まで保持）
      this.cleanupOldSecurityLogs();
    } catch (error) {
      console.error('SecurityService.persistSecurityLog: エラー', error.message);
    }
  },

  /**
   * 古いセキュリティログ削除
   */
  cleanupOldSecurityLogs() {
    try {
      const props = PropertiesService.getScriptProperties();
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
  },

  // ===========================================
  // 🔧 ユーティリティ・診断
  // ===========================================

  /**
   * セキュリティ状態診断
   * @returns {Object} 診断結果
   */
  diagnose() {
    const results = {
      service: 'SecurityService',
      timestamp: new Date().toISOString(),
      checks: [],
      securityScore: 0
    };

    try {
      // セッション状態確認
      const sessionValidation = this.validateSession();
      results.checks.push({
        name: 'Session Validation',
        status: sessionValidation.isValid ? '✅' : '❌',
        details: sessionValidation.isValid 
          ? `Active user: ${sessionValidation.userEmail}`
          : sessionValidation.error
      });

      // Service Accountトークン確認
      const token = this.getServiceAccountToken();
      results.checks.push({
        name: 'Service Account Token',
        status: token ? '✅' : '❌',
        details: token ? 'Token available' : 'Token generation failed'
      });

      // プロパティサービス確認
      try {
        PropertiesService.getScriptProperties().getProperties();
        results.checks.push({
          name: 'Properties Service',
          status: '✅',
          details: 'Property access available'
        });
      } catch (propError) {
        results.checks.push({
          name: 'Properties Service',
          status: '❌',
          details: propError.message
        });
      }

      // キャッシュサービス確認
      try {
        CacheService.getScriptCache().get('test');
        results.checks.push({
          name: 'Cache Service',
          status: '✅',
          details: 'Cache access available'
        });
      } catch (cacheError) {
        results.checks.push({
          name: 'Cache Service',
          status: '❌',
          details: cacheError.message
        });
      }

      // セキュリティスコア計算
      const passedChecks = results.checks.filter(check => check.status === '✅').length;
      results.securityScore = Math.round((passedChecks / results.checks.length) * 100);
      results.overall = results.securityScore >= 80 ? '✅' : '⚠️';

    } catch (error) {
      results.checks.push({
        name: 'Service Diagnosis',
        status: '❌',
        details: error.message
      });
      results.overall = '❌';
      results.securityScore = 0;
    }

    return results;
  },

  /**
   * スプレッドシートアクセス権限検証
   * @param {string} spreadsheetId - スプレッドシートID
   * @returns {Object} 検証結果
   */
  validateSpreadsheetAccess(spreadsheetId) {
    try {
      if (!spreadsheetId) {
        return {
          success: false,
          message: 'スプレッドシートIDが指定されていません'
        };
      }

      // アクセステスト
      const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      const name = spreadsheet.getName();
      const sheets = spreadsheet.getSheets().map(sheet => ({
        name: sheet.getName(),
        index: sheet.getIndex()
      }));

      return {
        success: true,
        name,
        sheets,
        message: 'アクセス権限が確認できました'
      };

    } catch (error) {
      console.error('SecurityService.validateSpreadsheetAccess エラー:', error.message);

      let userMessage = 'スプレッドシートにアクセスできません。';
      if (error.message.includes('Permission') || error.message.includes('権限')) {
        userMessage = 'スプレッドシートへのアクセス権限がありません。編集権限を確認してください。';
      } else if (error.message.includes('not found') || error.message.includes('見つかりません')) {
        userMessage = 'スプレッドシートが見つかりません。URLが正しいか確認してください。';
      }

      return {
        success: false,
        message: userMessage,
        error: error.message
      };
    }
  },

  /**
   * セキュリティ設定推奨事項
   * @returns {Array} 推奨事項リスト
   */
  getSecurityRecommendations() {
    const recommendations = [];

    try {
      // 基本的なセキュリティチェック
      const session = this.validateSession();
      if (!session.isValid) {
        recommendations.push({
          priority: 'high',
          category: 'authentication',
          message: 'セッションの再認証が必要です',
          action: 'ログインを確認してください'
        });
      }

      // トークンの有効性確認
      const token = this.getServiceAccountToken();
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

});
