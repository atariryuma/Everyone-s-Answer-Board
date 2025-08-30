/**
 * @fileoverview 監視・保守性向上機能
 * 構造化ログ、健全性チェック、システム診断機能を提供
 */

/**
 * 構造化ログマネージャー
 */
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
    console.log('システムヘルスチェック開始');
    const results = new Map();
    let overallStatus = 'healthy';
    
    for (const [name, check] of this.checks) {
      try {
        const startTime = Date.now();
        const result = await this._runSingleCheck(name, check);
        const duration = Date.now() - startTime;
        console.log(`HealthCheck: ${name} - ${duration}ms`);
        
        results.set(name, result);
        
        if (result.status === 'unhealthy' && check.critical) {
          overallStatus = 'critical';
        } else if (result.status === 'unhealthy' && overallStatus === 'healthy') {
          overallStatus = 'degraded';
        }
        
      } catch (error) {
        console.error(`ヘルスチェック実行エラー: ${name}`, { error: error.message });
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
    
    console.log('システムヘルスチェック完了', { status: overallStatus });
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
  console.log('システム診断開始');
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
      userEmail: User.email()
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
      domainRestriction: Deploy.domain(),
      userVerificationActive: typeof verifyUserAccess === 'function'
    };
    
    // 依存関係チェック
    diagnostics.dependencies = {
      sheetsApi: typeof getSheetsServiceCached === 'function',
      database: typeof findUserById === 'function',
      cache: typeof cacheManager !== 'undefined',
      monitoring: typeof logger !== 'undefined'
    };
    
    console.log('システム診断完了');
    return diagnostics;
    
  } catch (error) {
    console.error('システム診断エラー', { error: error.message });
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
      const email = User.email();
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
  
  console.log('ヘルスチェック項目初期化完了');
}

/**
 * 監視ダッシュボード用データ取得
 */
function getMonitoringDashboard() {
  try {
    const startTime = Date.now();
    
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
    
    const duration = Date.now() - startTime;
    console.log(`MonitoringDashboard - ${duration}ms`);
    return dashboard;
    
  } catch (error) {
    console.error('監視ダッシュボードデータ取得エラー', { error: error.message });
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
    console.warn('エラーログ取得失敗', { error: error.message });
    return [];
  }
}

/**
 * システム健全性の定期チェックを実行
 */
function runScheduledHealthCheck() {
  return healthChecker.runAllChecks()
    .then(result => {
      console.log('定期ヘルスチェック完了', { status: result.overall });
      return result;
    })
    .catch(error => {
      console.error('定期ヘルスチェックエラー', { error: error.message });
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
  console.log('監視システム初期化完了');
} catch (error) {
  console.error('監視システム初期化エラー:', error.message);
}