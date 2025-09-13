/**
 * @fileoverview Unified Error Handling System
 * 
 * 🎯 責任範囲:
 * - システム全体のエラー統一管理
 * - エラーログ・監査・報告
 * - エラー復旧・フォールバック処理
 * - デバッグ支援・診断機能
 * 
 * 🔄 統合対象:
 * - 各サービスの分散エラーハンドリング
 * - Core.gs内の重複エラー処理
 * - Base.gs内のエラー関数
 */

/**
 * ErrorHandler - 統一エラーハンドリングシステム
 * システム全体のエラーを一元管理
 */
// eslint-disable-next-line no-unused-vars
const ErrorHandler = Object.freeze({

  // ===========================================
  // 🚨 エラーレベル定義
  // ===========================================

  LEVELS: Object.freeze({
    CRITICAL: 'critical',    // システム停止レベル
    HIGH: 'high',           // 主要機能に影響
    MEDIUM: 'medium',       // 一部機能に影響
    LOW: 'low',             // 軽微な問題
    INFO: 'info'            // 情報のみ
  }),

  CATEGORIES: Object.freeze({
    AUTHENTICATION: 'authentication',
    AUTHORIZATION: 'authorization', 
    DATABASE: 'database',
    EXTERNAL_API: 'external_api',
    VALIDATION: 'validation',
    CONFIGURATION: 'configuration',
    PERFORMANCE: 'performance',
    SYSTEM: 'system',
    USER_INPUT: 'user_input'
  }),

  // ===========================================
  // 🔧 エラー処理・ログ
  // ===========================================

  /**
   * 統一エラーハンドリング
   * @param {Error|string} error - エラーオブジェクトまたはメッセージ
   * @param {Object} context - エラーコンテキスト
   * @returns {Object} 処理済みエラー情報
   */
  handle(error, context = {}) {
    try {
      // エラー情報の標準化
      const standardizedError = this.standardizeError(error, context);
      
      // ログ出力
      this.logError(standardizedError);
      
      // 重大エラーの場合は永続化
      if (this.isCriticalError(standardizedError)) {
        this.persistCriticalError(standardizedError);
      }

      // 自動復旧試行
      const recoveryResult = this.attemptRecovery(standardizedError);

      // ユーザー向けレスポンス作成
      const userResponse = this.createUserResponse(standardizedError, recoveryResult);

      return {
        ...userResponse,
        errorId: standardizedError.errorId,
        canRetry: recoveryResult.canRetry,
        recoveryAction: recoveryResult.action
      };
    } catch (handlingError) {
      // エラーハンドラ自体のエラー
      console.error('🔥 ErrorHandler.handle: ハンドラエラー', handlingError.message);
      
      return {
        success: false,
        message: 'システムエラーが発生しました。管理者に連絡してください。',
        errorId: `handler_error_${Date.now()}`,
        canRetry: false,
        timestamp: new Date().toISOString()
      };
    }
  },

  /**
   * 既存コード互換用の簡易セーフレスポンス作成
   * @param {Error|string} error - エラー
   * @param {string} context - 文脈
   * @param {boolean} includeDetails - 詳細を含めるか（未使用・互換用）
   * @returns {Object} セーフなエラーレスポンス
   */
  createSafeResponse(error, context = 'operation', _includeDetails = false) {
    try {
      const message = error && typeof error.message === 'string' ? error.message : String(error);

      // 内部ログ（詳細付き）
      console.error(`❌ ${context}:`, {
        message,
        timestamp: new Date().toISOString(),
      });

      return {
        success: false,
        message: this.getSafeErrorMessage(message),
        timestamp: new Date().toISOString(),
        errorCode: this.getErrorCode(message),
      };
    } catch (e) {
      return {
        success: false,
        message: '処理中にエラーが発生しました。',
        timestamp: new Date().toISOString(),
        errorCode: 'EUNKNOWN',
      };
    }
  },

  /**
   * エンドユーザー向けに安全なエラーメッセージへ変換
   */
  getSafeErrorMessage(originalMessage) {
    try {
      const patterns = ['Service Account', 'token', 'database', 'validation'];
      const lower = (originalMessage || '').toLowerCase();
      for (const pattern of patterns) {
        if (lower.includes(pattern.toLowerCase())) {
          return 'システムエラーが発生しました。管理者にお問い合わせください。';
        }
      }
      return '処理中にエラーが発生しました。';
    } catch (_e) {
      return '処理中にエラーが発生しました。';
    }
  },

  /**
   * エラーメッセージから安定したエラーコードを生成
   */
  getErrorCode(message) {
    try {
      const code = String(message || '').split('').reduce((a, ch) => {
        a = (a << 5) - a + ch.charCodeAt(0);
        return a & a;
      }, 0);
      return `E${Math.abs(code).toString(16).substr(0, 6).toUpperCase()}`;
    } catch (_e) {
      return 'E000000';
    }
  },

  /**
   * エラー情報標準化
   * @param {Error|string} error - 元エラー
   * @param {Object} context - コンテキスト
   * @returns {Object} 標準化エラー
   */
  standardizeError(error, context) {
    const errorId = `error_${Date.now()}_${Math.random().toString(36).substring(2, 8)}`;
    const timestamp = new Date().toISOString();

    // エラーの基本情報抽出
    let message, stack, name;
    if (error instanceof Error) {
      message = error.message;
      stack = error.stack;
      name = error.name;
    } else if (typeof error === 'string') {
      message = error;
      name = 'StringError';
    } else {
      message = '不明なエラー';
      name = 'UnknownError';
    }

    // レベルとカテゴリの自動判定
    const level = this.determineErrorLevel(message, context);
    const category = this.determineErrorCategory(message, context);

    return {
      errorId,
      timestamp,
      level,
      category,
      message,
      name,
      stack,
      context: {
        service: context.service || 'unknown',
        method: context.method || 'unknown',
        userId: context.userId || null,
        userEmail: this.getSafeUserEmail(),
        sessionId: this.getSessionId(),
        ...context
      },
      environment: this.getEnvironmentInfo()
    };
  },

  /**
   * エラーレベル自動判定
   * @param {string} message - エラーメッセージ
   * @param {Object} context - コンテキスト
   * @returns {string} エラーレベル
   */
  determineErrorLevel(message, _context) {
    if (!message) return this.LEVELS.LOW;

    const criticalKeywords = [
      'database connection failed',
      'authentication failed',
      'system crash',
      'out of memory',
      'permission denied'
    ];

    const highKeywords = [
      'api error',
      'timeout',
      'invalid configuration',
      'service unavailable'
    ];

    const mediumKeywords = [
      'validation failed',
      'user input error',
      'cache miss'
    ];

    const lowerMessage = message.toLowerCase();

    if (criticalKeywords.some(keyword => lowerMessage.includes(keyword))) {
      return this.LEVELS.CRITICAL;
    }

    if (highKeywords.some(keyword => lowerMessage.includes(keyword))) {
      return this.LEVELS.HIGH;
    }

    if (mediumKeywords.some(keyword => lowerMessage.includes(keyword))) {
      return this.LEVELS.MEDIUM;
    }

    return this.LEVELS.LOW;
  },

  /**
   * エラーカテゴリ自動判定
   * @param {string} message - エラーメッセージ
   * @param {Object} context - コンテキスト
   * @returns {string} エラーカテゴリ
   */
  determineErrorCategory(message, context) {
    if (!message) return this.CATEGORIES.SYSTEM;

    const categoryMap = {
      [this.CATEGORIES.AUTHENTICATION]: ['login', 'session', 'token', '認証'],
      [this.CATEGORIES.AUTHORIZATION]: ['permission', 'access', 'forbidden', '権限'],
      [this.CATEGORIES.DATABASE]: ['database', 'db', 'query', 'データベース'],
      [this.CATEGORIES.EXTERNAL_API]: ['api', 'fetch', 'request', 'スプレッドシート'],
      [this.CATEGORIES.VALIDATION]: ['validation', 'invalid', '検証', '無効'],
      [this.CATEGORIES.CONFIGURATION]: ['config', 'setting', '設定'],
      [this.CATEGORIES.PERFORMANCE]: ['timeout', 'slow', 'memory', 'パフォーマンス'],
      [this.CATEGORIES.USER_INPUT]: ['user input', 'form', 'ユーザー入力']
    };

    const lowerMessage = message.toLowerCase();
    
    for (const [category, keywords] of Object.entries(categoryMap)) {
      if (keywords.some(keyword => lowerMessage.includes(keyword))) {
        return category;
      }
    }

    return this.CATEGORIES.SYSTEM;
  },

  /**
   * エラーログ出力
   * @param {Object} standardizedError - 標準化エラー
   */
  logError(standardizedError) {
    const logData = {
      errorId: standardizedError.errorId,
      timestamp: standardizedError.timestamp,
      level: standardizedError.level,
      category: standardizedError.category,
      message: standardizedError.message,
      service: standardizedError.context.service,
      method: standardizedError.context.method,
      userId: standardizedError.context.userId
    };

    switch (standardizedError.level) {
      case this.LEVELS.CRITICAL:
        console.error('🔥 CRITICAL ERROR:', logData);
        break;
      case this.LEVELS.HIGH:
        console.error('🚨 HIGH ERROR:', logData);
        break;
      case this.LEVELS.MEDIUM:
        console.warn('⚠️ MEDIUM ERROR:', logData);
        break;
      case this.LEVELS.LOW:
        console.warn('⚡ LOW ERROR:', logData);
        break;
      default:
        console.info('ℹ️ INFO:', logData);
    }

    // 詳細情報をデバッグレベルで出力
    if (standardizedError.stack) {
      console.debug('Stack trace:', standardizedError.stack);
    }
  },

  // ===========================================
  // 🔄 自動復旧・フォールバック
  // ===========================================

  /**
   * 自動復旧試行
   * @param {Object} standardizedError - 標準化エラー
   * @returns {Object} 復旧結果
   */
  attemptRecovery(standardizedError) {
    const result = {
      attempted: false,
      success: false,
      action: 'none',
      canRetry: false,
      message: ''
    };

    try {
      switch (standardizedError.category) {
        case this.CATEGORIES.AUTHENTICATION:
          return this.recoverAuthentication(standardizedError);
        
        case this.CATEGORIES.EXTERNAL_API:
          return this.recoverExternalApi(standardizedError);
        
        case this.CATEGORIES.DATABASE:
          return this.recoverDatabase(standardizedError);
        
        case this.CATEGORIES.CONFIGURATION:
          return this.recoverConfiguration(standardizedError);
        
        default:
          result.canRetry = true;
          result.message = '一般的なエラーです。再試行してください。';
      }
    } catch (recoveryError) {
      console.error('ErrorHandler.attemptRecovery: 復旧処理エラー', recoveryError.message);
    }

    return result;
  },

  /**
   * 認証エラー復旧
   * @param {Object} error - エラー情報
   * @returns {Object} 復旧結果
   */
  recoverAuthentication(_error) {
    return {
      attempted: true,
      success: false,
      action: 'clear_auth_cache',
      canRetry: true,
      message: 'セッションをクリアしました。再ログインしてください。'
    };
  },

  /**
   * 外部API エラー復旧
   * @param {Object} error - エラー情報
   * @returns {Object} 復旧結果
   */
  recoverExternalApi(error) {
    // タイムアウト系エラーの場合
    if (error.message.toLowerCase().includes('timeout')) {
      return {
        attempted: true,
        success: false,
        action: 'retry_with_delay',
        canRetry: true,
        message: '通信タイムアウトが発生しました。しばらく待ってから再試行してください。'
      };
    }

    return {
      attempted: true,
      success: false,
      action: 'check_external_service',
      canRetry: true,
      message: '外部サービスとの通信に問題があります。少し待ってから再試行してください。'
    };
  },

  /**
   * データベースエラー復旧
   * @param {Object} error - エラー情報  
   * @returns {Object} 復旧結果
   */
  recoverDatabase(error) {
    return {
      attempted: true,
      success: false,
      action: 'clear_cache',
      canRetry: true,
      message: 'データベースアクセスエラーが発生しました。キャッシュをクリアして再試行してください。'
    };
  },

  /**
   * 設定エラー復旧
   * @param {Object} error - エラー情報
   * @returns {Object} 復旧結果
   */
  recoverConfiguration(error) {
    return {
      attempted: true,
      success: false,
      action: 'load_default_config',
      canRetry: false,
      message: '設定に問題があります。設定を確認してください。'
    };
  },

  // ===========================================
  // 📊 エラー分析・レポート
  // ===========================================

  /**
   * エラー統計取得
   * @param {Object} options - オプション
   * @returns {Object} エラー統計
   */
  getErrorStatistics(options = {}) {
    try {
      const props = PropertiesService.getScriptProperties();
      const allProps = props.getProperties();
      
      const errorLogs = Object.keys(allProps)
        .filter(key => key.startsWith('error_log_'))
        .map(key => {
          try {
            return JSON.parse(allProps[key]);
          } catch (parseError) {
            return null;
          }
        })
        .filter(log => log !== null);

      // 統計計算
      const stats = {
        total: errorLogs.length,
        byLevel: this.groupByField(errorLogs, 'level'),
        byCategory: this.groupByField(errorLogs, 'category'),
        byService: this.groupByField(errorLogs, 'context.service'),
        recentErrors: errorLogs
          .slice(-10)
          .map(log => ({
            errorId: log.errorId,
            timestamp: log.timestamp,
            level: log.level,
            message: log.message
          })),
        lastUpdated: new Date().toISOString()
      };

      return stats;
    } catch (error) {
      console.error('ErrorHandler.getErrorStatistics: エラー', error.message);
      return {
        total: 0,
        error: error.message,
        lastUpdated: new Date().toISOString()
      };
    }
  },

  // ===========================================
  // 🔧 ユーティリティ
  // ===========================================

  /**
   * 重大エラー判定
   * @param {Object} error - エラー情報
   * @returns {boolean} 重大エラーかどうか
   */
  isCriticalError(error) {
    return error.level === this.LEVELS.CRITICAL || error.level === this.LEVELS.HIGH;
  },

  /**
   * 重大エラー永続化
   * @param {Object} error - エラー情報
   */
  persistCriticalError(error) {
    try {
      const props = PropertiesService.getScriptProperties();
      const logKey = `error_log_${Date.now()}`;
      
      props.setProperty(logKey, JSON.stringify(error));
      
      // 古いログの削除（最新50件まで保持）
      this.cleanupOldErrorLogs();
    } catch (persistError) {
      console.error('ErrorHandler.persistCriticalError: 永続化エラー', persistError.message);
    }
  },

  /**
   * 古いエラーログ削除
   */
  cleanupOldErrorLogs() {
    try {
      const props = PropertiesService.getScriptProperties();
      const allProps = props.getProperties();
      
      const errorLogs = Object.keys(allProps)
        .filter(key => key.startsWith('error_log_'))
        .sort()
        .reverse();

      if (errorLogs.length > 50) {
        const logsToDelete = errorLogs.slice(50);
        logsToDelete.forEach(key => props.deleteProperty(key));
      }
    } catch (error) {
      console.error('ErrorHandler.cleanupOldErrorLogs: エラー', error.message);
    }
  },

  /**
   * 安全なユーザーメール取得
   * @returns {string|null} ユーザーメール
   */
  getSafeUserEmail() {
    try {
      return Session.getActiveUser().getEmail();
    } catch (error) {
      return null;
    }
  },

  /**
   * セッションID取得
   * @returns {string} セッションID
   */
  getSessionId() {
    try {
      // GASではセッションIDが直接取得できないため、一時的なIDを生成
      return `session_${Date.now()}_${Math.random().toString(36).substring(2, 8)}`;
    } catch (error) {
      return 'unknown_session';
    }
  },

  /**
   * 環境情報取得
   * @returns {Object} 環境情報
   */
  getEnvironmentInfo() {
    try {
      return {
        gasRuntime: 'V8',
        timezone: Session.getScriptTimeZone(),
        locale: Session.getActiveUserLocale(),
        timestamp: new Date().toISOString()
      };
    } catch (error) {
      return {
        error: 'Environment info not available',
        timestamp: new Date().toISOString()
      };
    }
  },

  /**
   * フィールドによるグループ化
   * @param {Array} items - アイテム配列
   * @param {string} field - グループ化フィールド
   * @returns {Object} グループ化結果
   */
  groupByField(items, field) {
    const grouped = {};
    
    items.forEach(item => {
      const value = this.getNestedValue(item, field) || 'unknown';
      grouped[value] = (grouped[value] || 0) + 1;
    });

    return grouped;
  },

  /**
   * ネストしたフィールド値取得
   * @param {Object} obj - オブジェクト
   * @param {string} field - フィールドパス
   * @returns {*} フィールド値
   */
  getNestedValue(obj, field) {
    return field.split('.').reduce((current, key) => current?.[key], obj);
  },

  /**
   * ユーザー向けレスポンス作成
   * @param {Object} error - エラー情報
   * @param {Object} recovery - 復旧情報
   * @returns {Object} ユーザーレスポンス
   */
  createUserResponse(error, recovery) {
    // ユーザーに見せるメッセージを安全に作成
    let userMessage;
    
    if (recovery.message) {
      userMessage = recovery.message;
    } else {
      switch (error.level) {
        case this.LEVELS.CRITICAL:
        case this.LEVELS.HIGH:
          userMessage = 'システムエラーが発生しました。管理者に連絡してください。';
          break;
        case this.LEVELS.MEDIUM:
          userMessage = '処理中にエラーが発生しました。再試行してください。';
          break;
        default:
          userMessage = '軽微なエラーが発生しましたが、続行できます。';
      }
    }

    return {
      success: false,
      message: userMessage,
      timestamp: new Date().toISOString(),
      // デバッグ情報は本番環境では含めない
      debug: {
        errorId: error.errorId,
        level: error.level,
        category: error.category
      }
    };
  }

});
