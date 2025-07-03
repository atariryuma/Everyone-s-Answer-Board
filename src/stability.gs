/**
 * @fileoverview 安定性強化システム - 2024年最新技術
 * エラーハンドリング、フォールバック、復旧機能を包括的に実装
 */

/**
 * 回復力のあるエラーハンドリングシステム
 */
class StabilityEnhancer {
  
  /**
   * サーキットブレーカーパターン実装
   */
  static createCircuitBreaker(operation, options = {}) {
    const {
      failureThreshold = 5,
      resetTimeoutMs = 60000,
      monitorWindowMs = 300000 // 5分
    } = options;
    
    let state = 'CLOSED'; // CLOSED, OPEN, HALF_OPEN
    let failureCount = 0;
    let lastFailureTime = 0;
    let successCount = 0;
    
    return {
      execute: async (...args) => {
        const now = Date.now();
        
        // 監視ウィンドウ外の失敗をリセット
        if (now - lastFailureTime > monitorWindowMs) {
          failureCount = 0;
          state = 'CLOSED';
        }
        
        // OPEN状態のチェック
        if (state === 'OPEN') {
          if (now - lastFailureTime < resetTimeoutMs) {
            throw new Error('Circuit breaker is OPEN');
          } else {
            state = 'HALF_OPEN';
            successCount = 0;
          }
        }
        
        try {
          const result = await operation(...args);
          
          // 成功時の処理
          if (state === 'HALF_OPEN') {
            successCount++;
            if (successCount >= 3) {
              state = 'CLOSED';
              failureCount = 0;
            }
          }
          
          return result;
        } catch (error) {
          failureCount++;
          lastFailureTime = now;
          
          if (failureCount >= failureThreshold) {
            state = 'OPEN';
          }
          
          throw error;
        }
      },
      
      getState: () => ({ state, failureCount, lastFailureTime })
    };
  }
  
  /**
   * 自動復旧機能付き操作実行
   */
  static resilientExecute(operation, options = {}) {
    const {
      maxRetries = 3,
      baseDelayMs = 1000,
      maxDelayMs = 30000,
      backoffMultiplier = 2,
      jitterMs = 100,
      retryCondition = (error) => true
    } = options;
    
    return new Promise((resolve, reject) => {
      let attempt = 0;
      
      const tryOperation = async () => {
        try {
          const result = await operation();
          resolve(result);
        } catch (error) {
          attempt++;
          
          if (attempt > maxRetries || !retryCondition(error)) {
            reject(new Error(`Operation failed after ${attempt} attempts: ${error.message}`));
            return;
          }
          
          // 指数バックオフ + ジッター
          const delay = Math.min(
            baseDelayMs * Math.pow(backoffMultiplier, attempt - 1) + 
            Math.random() * jitterMs,
            maxDelayMs
          );
          
          console.warn(`Attempt ${attempt} failed, retrying in ${delay}ms: ${error.message}`);
          
          setTimeout(tryOperation, delay);
        }
      };
      
      tryOperation();
    });
  }
  
  /**
   * 健康状態監視システム
   */
  static createHealthMonitor() {
    const metrics = {
      totalRequests: 0,
      successfulRequests: 0,
      failedRequests: 0,
      averageResponseTime: 0,
      lastError: null,
      lastSuccessTime: 0,
      systemLoad: 0
    };
    
    return {
      recordSuccess: (responseTime) => {
        metrics.totalRequests++;
        metrics.successfulRequests++;
        metrics.lastSuccessTime = Date.now();
        
        // 移動平均で応答時間更新
        metrics.averageResponseTime = 
          (metrics.averageResponseTime * (metrics.successfulRequests - 1) + responseTime) / 
          metrics.successfulRequests;
      },
      
      recordFailure: (error) => {
        metrics.totalRequests++;
        metrics.failedRequests++;
        metrics.lastError = {
          message: error.message,
          timestamp: Date.now(),
          stack: error.stack
        };
      },
      
      getHealth: () => {
        const successRate = metrics.totalRequests > 0 ? 
          (metrics.successfulRequests / metrics.totalRequests) : 0;
        
        const timeSinceLastSuccess = Date.now() - metrics.lastSuccessTime;
        
        let status = 'HEALTHY';
        if (successRate < 0.9) status = 'DEGRADED';
        if (successRate < 0.5 || timeSinceLastSuccess > 300000) status = 'UNHEALTHY';
        
        return {
          status,
          successRate,
          ...metrics,
          timeSinceLastSuccess
        };
      },
      
      reset: () => {
        Object.keys(metrics).forEach(key => {
          if (typeof metrics[key] === 'number') {
            metrics[key] = 0;
          } else {
            metrics[key] = null;
          }
        });
      }
    };
  }
  
  /**
   * フェイルセーフ データ アクセス
   */
  static createFailsafeDataAccess(primarySource, fallbackSources = []) {
    return async (key, defaultValue = null) => {
      const sources = [primarySource, ...fallbackSources];
      
      for (let i = 0; i < sources.length; i++) {
        try {
          const result = await sources[i](key);
          if (result !== null && result !== undefined) {
            return result;
          }
        } catch (error) {
          console.warn(`Data source ${i} failed for key ${key}:`, error.message);
          if (i === sources.length - 1) {
            console.error('All data sources failed, returning default value');
          }
        }
      }
      
      return defaultValue;
    };
  }
  
  /**
   * リソース制限監視
   */
  static createResourceMonitor() {
    const limits = {
      maxExecutionTime: 330000, // 5.5分
      maxMemoryUsage: 0.8, // 80%（概算）
      maxApiCalls: 900 // 15分あたり
    };
    
    let executionStart = Date.now();
    let apiCallCount = 0;
    let lastApiCallReset = Date.now();
    
    return {
      checkExecutionTime: () => {
        const elapsed = Date.now() - executionStart;
        return {
          elapsed,
          remaining: limits.maxExecutionTime - elapsed,
          withinLimit: elapsed < limits.maxExecutionTime
        };
      },
      
      recordApiCall: () => {
        const now = Date.now();
        
        // 15分ごとにAPIコールカウントリセット
        if (now - lastApiCallReset > 900000) {
          apiCallCount = 0;
          lastApiCallReset = now;
        }
        
        apiCallCount++;
        
        return {
          count: apiCallCount,
          remaining: limits.maxApiCalls - apiCallCount,
          withinLimit: apiCallCount < limits.maxApiCalls
        };
      },
      
      shouldContinue: () => {
        const timeCheck = this.checkExecutionTime();
        const apiCheck = this.recordApiCall();
        
        return timeCheck.withinLimit && apiCheck.withinLimit;
      },
      
      reset: () => {
        executionStart = Date.now();
        apiCallCount = 0;
        lastApiCallReset = Date.now();
      }
    };
  }
}

/**
 * データ整合性チェッカー
 */
class DataIntegrityChecker {
  
  /**
   * データバリデーション
   */
  static validateData(data, schema) {
    const errors = [];
    
    // 必須フィールドチェック
    if (schema.required) {
      schema.required.forEach(field => {
        if (!(field in data) || data[field] === null || data[field] === undefined) {
          errors.push(`Required field missing: ${field}`);
        }
      });
    }
    
    // 型チェック
    if (schema.types) {
      Object.keys(schema.types).forEach(field => {
        if (field in data) {
          const expectedType = schema.types[field];
          const actualType = typeof data[field];
          
          if (actualType !== expectedType) {
            errors.push(`Type mismatch for ${field}: expected ${expectedType}, got ${actualType}`);
          }
        }
      });
    }
    
    // カスタムバリデーション
    if (schema.validators) {
      Object.keys(schema.validators).forEach(field => {
        if (field in data) {
          const validator = schema.validators[field];
          if (!validator(data[field])) {
            errors.push(`Validation failed for field: ${field}`);
          }
        }
      });
    }
    
    return {
      isValid: errors.length === 0,
      errors: errors
    };
  }
  
  /**
   * データ修復機能
   */
  static repairData(data, repairRules) {
    const repaired = { ...data };
    
    Object.keys(repairRules).forEach(field => {
      const rule = repairRules[field];
      
      if (!(field in repaired) || repaired[field] === null || repaired[field] === undefined) {
        if (rule.default !== undefined) {
          repaired[field] = rule.default;
        }
      }
      
      if (rule.transform && field in repaired) {
        try {
          repaired[field] = rule.transform(repaired[field]);
        } catch (error) {
          console.warn(`Transformation failed for ${field}:`, error.message);
        }
      }
    });
    
    return repaired;
  }
}

/**
 * 緊急シャットダウン機能
 */
class EmergencyShutdown {
  
  static initiate(reason, cleanupFunction = null) {
    console.error(`Emergency shutdown initiated: ${reason}`);
    
    try {
      // クリーンアップ処理
      if (cleanupFunction) {
        cleanupFunction();
      }
      
      // 進行中の処理を停止
      this._stopAllProcesses();
      
      // 状態を保存
      this._saveEmergencyState(reason);
      
    } catch (error) {
      console.error('Error during emergency shutdown:', error.message);
    }
    
    throw new Error(`Emergency shutdown: ${reason}`);
  }
  
  static _stopAllProcesses() {
    // アクティブな処理を停止
    // GASでは制限があるため、フラグベースの制御
    PropertiesService.getScriptProperties().setProperty('EMERGENCY_STOP', 'true');
  }
  
  static _saveEmergencyState(reason) {
    const emergencyData = {
      reason: reason,
      timestamp: Date.now(),
      executionId: Utilities.getUuid()
    };
    
    try {
      PropertiesService.getScriptProperties().setProperty(
        'EMERGENCY_STATE', 
        JSON.stringify(emergencyData)
      );
    } catch (error) {
      console.error('Failed to save emergency state:', error.message);
    }
  }
  
  static checkEmergencyStop() {
    const stopFlag = PropertiesService.getScriptProperties().getProperty('EMERGENCY_STOP');
    if (stopFlag === 'true') {
      throw new Error('Execution stopped by emergency shutdown');
    }
  }
  
  static clearEmergencyState() {
    const props = PropertiesService.getScriptProperties();
    props.deleteProperty('EMERGENCY_STOP');
    props.deleteProperty('EMERGENCY_STATE');
  }
}

/**
 * 自動復旧サービス
 */
class AutoRecoveryService {
  
  /**
   * システム復旧チェック
   */
  static performRecoveryCheck() {
    const results = {
      timestamp: Date.now(),
      checks: {},
      overallStatus: 'HEALTHY',
      recommendations: []
    };
    
    // キャッシュ健康状態チェック
    try {
      // キャッシュの健全性をチェック
  const cacheHealth = cacheManager.getHealth();
      results.checks.cache = cacheHealth;
      
      if (cacheHealth.healthScore < 70) {
        results.recommendations.push('Consider clearing expired cache items');
      }
    } catch (error) {
      results.checks.cache = { error: error.message };
      results.overallStatus = 'DEGRADED';
    }
    
    // プロパティサービス容量チェック
    try {
      const props = PropertiesService.getScriptProperties().getProperties();
      const propCount = Object.keys(props).length;
      results.checks.properties = { count: propCount };
      
      if (propCount > 450) { // 500制限の90%
        results.recommendations.push('Properties approaching limit, consider cleanup');
        results.overallStatus = 'WARNING';
      }
    } catch (error) {
      results.checks.properties = { error: error.message };
      results.overallStatus = 'DEGRADED';
    }
    
    // 緊急停止状態チェック
    try {
      const emergencyState = PropertiesService.getScriptProperties().getProperty('EMERGENCY_STATE');
      if (emergencyState) {
        results.checks.emergency = JSON.parse(emergencyState);
        results.overallStatus = 'EMERGENCY';
        results.recommendations.push('Clear emergency state if issue is resolved');
      }
    } catch (error) {
      // 緊急状態データなし（正常）
    }
    
    return results;
  }
  
  /**
   * 自動修復実行
   */
  static performAutoRepair() {
    const repairLog = [];
    
    try {
      // 期限切れキャッシュクリーンアップ
      cacheManager.clearExpired();
      repairLog.push('Cache cleanup completed');
      
      // 古い緊急状態クリア（24時間以上前）
      const emergencyState = PropertiesService.getScriptProperties().getProperty('EMERGENCY_STATE');
      if (emergencyState) {
        const data = JSON.parse(emergencyState);
        if (Date.now() - data.timestamp > 86400000) { // 24時間
          EmergencyShutdown.clearEmergencyState();
          repairLog.push('Cleared old emergency state');
        }
      }
      
    } catch (error) {
      repairLog.push(`Repair error: ${error.message}`);
    }
    
    return repairLog;
  }
}