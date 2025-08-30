/**
 * @fileoverview 監視・保守性向上機能
 * 構造化ログ、健全性チェック、システム診断機能を提供
 */

/**
 * 構造化ログマネージャー
 */
class StructuredLogger {
  constructor() {
    this.logLevels = {
      DEBUG: 0,
      INFO: 1,
      WARN: 2,
      ERROR: 3,
      FATAL: 4
    };
    this.currentLevel = this.logLevels.INFO;
    this.sessionId = Utilities.getUuid();
    this.startTime = Date.now();
  }
  
  /**
   * 構造化ログを出力
   */
  log(level, message, metadata = {}) {
    if (this.logLevels[level] < this.currentLevel) {
      return; // ログレベルが低い場合はスキップ
    }
    
    const logEntry = {
      timestamp: new Date().toISOString(),
      level: level,
      sessionId: this.sessionId,
      message: message,
      metadata: metadata,
      uptime: Date.now() - this.startTime
    };
    
    // レベルに応じた出力先を選択
    switch (level) {
      case 'DEBUG':
        console.log(`[DEBUG] ${JSON.stringify(logEntry)}`);
        break;
      case 'INFO':
        console.log(`[INFO] ${message}`, metadata);
        break;
      case 'WARN':
        console.warn(`[WARN] ${message}`, metadata);
        break;
      case 'ERROR':
      case 'FATAL':
        console.error(`[${level}] ${message}`, metadata);
        break;
    }
    
    // 重要なログは永続化（オプション）
    if (level === 'ERROR' || level === 'FATAL') {
      this._persistLog(logEntry);
    }
  }
  
  /**
   * 重要なログを永続化
   */
  _persistLog(logEntry) {
    try {
      // PropertiesServiceに最新のエラーログを保存
      const props = PropertiesService.getScriptProperties();
      const errorLogs = JSON.parse(props.getProperty('ERROR_LOGS') || '[]');
      
      // 最新20件のエラーログを保持
      errorLogs.push(logEntry);
      if (errorLogs.length > 20) {
        errorLogs.shift();
      }
      
      props.setProperty('ERROR_LOGS', JSON.stringify(errorLogs));
    } catch (error) {
      console.error('ログ永続化エラー:', error.message);
    }
  }
  
  /**
   * パフォーマンス測定開始
   */
  startTimer(operationName) {
    return {
      name: operationName,
      startTime: Date.now(),
      end: () => {
        const duration = Date.now() - this.startTime;
        this.log('INFO', `Performance: ${operationName}`, { duration: `${duration}ms` });
        return duration;
      }
    };
  }
  
  // 便利メソッド
  debug(message, metadata = {}) { this.log('DEBUG', message, metadata); }
  info(message, metadata = {}) { this.log('INFO', message, metadata); }
  warn(message, metadata = {}) { this.log('WARN', message, metadata); }
  error(message, metadata = {}) { this.log('ERROR', message, metadata); }
  fatal(message, metadata = {}) { this.log('FATAL', message, metadata); }
}

// グローバルロガーインスタンス
const monitoringLogger = new StructuredLogger();

/**
 * システム健全性チェック機能
 */
class SystemHealthChecker {
  constructor() {
    this.checks = new Map();
    this.lastCheckTime = null;
    this.healthStatus = 'unknown';
  }
  
  /**
   * ヘルスチェック項目を登録
   */
  registerCheck(name, checkFunction, options = {}) {
    this.checks.set(name, {
      fn: checkFunction,
      timeout: options.timeout || 5000,
      critical: options.critical || false,
      description: options.description || ''
    });
  }
  
  /**
   * 全てのヘルスチェックを実行
   */
  async runAllChecks() {
    monitoringLogger.info('システムヘルスチェック開始');
    const results = new Map();
    let overallStatus = 'healthy';
    
    for (const [name, check] of this.checks) {
      try {
        const timer = logger.startTimer(`HealthCheck: ${name}`);
        const result = await this._runSingleCheck(name, check);
        timer.end();
        
        results.set(name, result);
        
        if (result.status === 'unhealthy' && check.critical) {
          overallStatus = 'critical';
        } else if (result.status === 'unhealthy' && overallStatus === 'healthy') {
          overallStatus = 'degraded';
        }
        
      } catch (error) {
        monitoringLogger.error(`ヘルスチェック実行エラー: ${name}`, { error: error.message });
        results.set(name, {
          status: 'error',
          message: error.message,
          timestamp: new Date().toISOString()
        });
        
        if (check.critical) {
          overallStatus = 'critical';
        }
      }
    }
    
    this.healthStatus = overallStatus;
    this.lastCheckTime = new Date().toISOString();
    
    const healthReport = {
      overall: overallStatus,
      timestamp: this.lastCheckTime,
      checks: Object.fromEntries(results),
      summary: this._generateSummary(results)
    };
    
    monitoringLogger.info('システムヘルスチェック完了', { status: overallStatus });
    return healthReport;
  }
  
  /**
   * 単一のヘルスチェックを実行
   */
  async _runSingleCheck(name, check) {
    return new Promise((resolve, reject) => {
      const timeout = setTimeout(() => {
        reject(new Error(`ヘルスチェックタイムアウト: ${name}`));
      }, check.timeout);
      
      try {
        const result = check.fn();
        clearTimeout(timeout);
        resolve({
          status: result ? 'healthy' : 'unhealthy',
          message: result === true ? 'OK' : (result || 'Check failed'),
          timestamp: new Date().toISOString()
        });
      } catch (error) {
        clearTimeout(timeout);
        reject(error);
      }
    });
  }
  
  /**
   * ヘルスチェック結果のサマリーを生成
   */
  _generateSummary(results) {
    const total = results.size;
    let healthy = 0;
    let unhealthy = 0;
    let errors = 0;
    
    for (const [name, result] of results) {
      switch (result.status) {
        case 'healthy':
          healthy++;
          break;
        case 'unhealthy':
          unhealthy++;
          break;
        case 'error':
          errors++;
          break;
      }
    }
    
    return {
      total,
      healthy,
      unhealthy,
      errors,
      healthPercentage: Math.round((healthy / total) * 100)
    };
  }
}

// グローバルヘルスチェッカー
const healthChecker = new SystemHealthChecker();

/**
 * システム診断機能
 */
function performSystemDiagnostics() {
  monitoringLogger.info('システム診断開始');
  const diagnostics = {
    timestamp: new Date().toISOString(),
    environment: {},
    performance: {},
    security: {},
    dependencies: {}
  };
  
  try {
    // 環境情報
    diagnostics.environment = {
      gasVersion: 'V8 Runtime',
      timeZone: Session.getScriptTimeZone(),
      locale: Session.getActiveLocale(),
      userEmail: Session.getActiveUser().getEmail()
    };
    
    // パフォーマンス情報
    diagnostics.performance = {
      cacheHealth: cacheManager?.getHealth() || 'unavailable',
      memoryStats: globalContextManager?.activeContexts?.size || 0,
      lastExecutionTime: Date.now() - executionStartTime
    };
    
    // セキュリティ情報
    diagnostics.security = {
      authenticationEnabled: true,
      domainRestriction: getDeployUserDomainInfo(),
      userVerificationActive: typeof verifyUserAccess === 'function'
    };
    
    // 依存関係チェック
    diagnostics.dependencies = {
      sheetsApi: typeof getSheetsServiceCached === 'function',
      database: typeof findUserById === 'function',
      cache: typeof cacheManager !== 'undefined',
      monitoring: typeof logger !== 'undefined'
    };
    
    monitoringLogger.info('システム診断完了');
    return diagnostics;
    
  } catch (error) {
    monitoringLogger.error('システム診断エラー', { error: error.message });
    diagnostics.error = error.message;
    return diagnostics;
  }
}

/**
 * 基本的なヘルスチェック項目を初期化
 */
function initializeHealthChecks() {
  // データベース接続チェック
  healthChecker.registerCheck('database', () => {
    try {
      const props = PropertiesService.getScriptProperties();
      const dbId = props.getProperty('DATABASE_SPREADSHEET_ID');
      return !!dbId;
    } catch (error) {
      return false;
    }
  }, { critical: true, description: 'データベース接続確認' });
  
  // Sheets API サービスチェック
  healthChecker.registerCheck('sheetsApi', () => {
    try {
      const service = getSheetsServiceCached();
      return !!(service && service.baseUrl && service.accessToken);
    } catch (error) {
      return false;
    }
  }, { critical: true, description: 'Sheets API サービス状態' });
  
  // キャッシュシステムチェック
  healthChecker.registerCheck('cache', () => {
    try {
      if (!cacheManager) return false;
      const health = cacheManager.getHealth();
      return health.status === 'ok';
    } catch (error) {
      return false;
    }
  }, { critical: false, description: 'キャッシュシステム健全性' });
  
  // 認証システムチェック
  healthChecker.registerCheck('authentication', () => {
    try {
      const email = Session.getActiveUser().getEmail();
      return !!email;
    } catch (error) {
      return false;
    }
  }, { critical: true, description: '認証システム動作確認' });
  
  // メモリ使用量チェック
  healthChecker.registerCheck('memory', () => {
    try {
      const contextCount = globalContextManager.activeContexts.size;
      return contextCount < globalContextManager.maxConcurrentContexts;
    } catch (error) {
      return false;
    }
  }, { critical: false, description: 'メモリ使用量確認' });
  
  monitoringLogger.info('ヘルスチェック項目初期化完了');
}

/**
 * 監視ダッシュボード用データ取得
 */
function getMonitoringDashboard() {
  try {
    const timer = logger.startTimer('MonitoringDashboard');
    
    // 各種統計情報を収集
    const dashboard = {
      timestamp: new Date().toISOString(),
      system: performSystemDiagnostics(),
      health: healthChecker.lastCheckTime ? {
        status: healthChecker.healthStatus,
        lastCheck: healthChecker.lastCheckTime
      } : null,
      cache: cacheManager?.getHealth() || null,
      memory: globalContextManager ? {
        activeContexts: globalContextManager.activeContexts.size,
        maxContexts: globalContextManager.maxConcurrentContexts
      } : null,
      errors: getRecentErrors(),
      uptime: Date.now() - executionStartTime
    };
    
    timer.end();
    return dashboard;
    
  } catch (error) {
    monitoringLogger.error('監視ダッシュボードデータ取得エラー', { error: error.message });
    return {
      timestamp: new Date().toISOString(),
      error: error.message
    };
  }
}

/**
 * 最近のエラーログを取得
 */
function getRecentErrors() {
  try {
    const props = PropertiesService.getScriptProperties();
    const errorLogs = JSON.parse(props.getProperty('ERROR_LOGS') || '[]');
    return errorLogs.slice(-10); // 最新10件
  } catch (error) {
    monitoringLogger.warn('エラーログ取得失敗', { error: error.message });
    return [];
  }
}

/**
 * システム健全性の定期チェックを実行
 */
function runScheduledHealthCheck() {
  return healthChecker.runAllChecks()
    .then(result => {
      monitoringLogger.info('定期ヘルスチェック完了', { status: result.overall });
      return result;
    })
    .catch(error => {
      monitoringLogger.error('定期ヘルスチェックエラー', { error: error.message });
      return {
        overall: 'error',
        timestamp: new Date().toISOString(),
        error: error.message
      };
    });
}

// 初期化
try {
  initializeHealthChecks();
  monitoringLogger.info('監視システム初期化完了');
} catch (error) {
  console.error('監視システム初期化エラー:', error.message);
}