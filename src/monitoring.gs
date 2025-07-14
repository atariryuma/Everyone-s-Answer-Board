/**
 * @fileoverview リアルタイム監視機能 - パフォーマンス監視の自動化
 */

/**
 * 監視設定
 */
const MONITORING_CONFIG = {
  ENABLED: true,
  ALERT_THRESHOLDS: {
    ERROR_RATE: 10,           // エラー率（%）
    RESPONSE_TIME: 5000,      // レスポンス時間（ms）
    CACHE_HIT_RATE: 50,       // キャッシュヒット率（%）
    API_USAGE_RATE: 80        // API使用率（%）
  },
  COLLECTION_INTERVAL: 60000, // データ収集間隔（ms）
  RETENTION_HOURS: 24,        // データ保持時間
  MAX_LOG_ENTRIES: 1000       // 最大ログエントリ数
};

/**
 * パフォーマンス監視マネージャー
 */
class PerformanceMonitor {
  constructor() {
    this.metrics = {
      requests: 0,
      errors: 0,
      totalResponseTime: 0,
      apiCalls: 0,
      cacheHits: 0,
      cacheMisses: 0,
      lastReset: Date.now()
    };
    
    this.alerts = [];
    this.isMonitoring = false;
  }

  /**
   * 監視を開始
   */
  start() {
    if (!MONITORING_CONFIG.ENABLED || this.isMonitoring) {
      return;
    }
    
    this.isMonitoring = true;
    console.log('📊 パフォーマンス監視を開始しました');
    
    // 定期的なメトリクス収集を開始
    this.scheduleMetricsCollection();
  }

  /**
   * 監視を停止
   */
  stop() {
    this.isMonitoring = false;
    console.log('📊 パフォーマンス監視を停止しました');
  }

  /**
   * リクエストを記録
   * @param {string} operation - 操作名
   * @param {number} responseTime - レスポンス時間（ms）
   * @param {boolean} success - 成功フラグ
   */
  recordRequest(operation, responseTime, success = true) {
    if (!MONITORING_CONFIG.ENABLED) return;
    
    this.metrics.requests++;
    this.metrics.totalResponseTime += responseTime;
    
    if (!success) {
      this.metrics.errors++;
    }
    
    // アラートチェック
    this.checkAlerts(operation, responseTime, success);
    
    // ログ記録
    this.logMetric({
      timestamp: Date.now(),
      operation: operation,
      responseTime: responseTime,
      success: success,
      type: 'request'
    });
  }

  /**
   * API呼び出しを記録
   * @param {string} apiType - API種別
   * @param {number} responseTime - レスポンス時間
   */
  recordApiCall(apiType, responseTime) {
    if (!MONITORING_CONFIG.ENABLED) return;
    
    this.metrics.apiCalls++;
    
    this.logMetric({
      timestamp: Date.now(),
      apiType: apiType,
      responseTime: responseTime,
      type: 'api_call'
    });
  }

  /**
   * キャッシュアクセスを記録
   * @param {boolean} hit - ヒットフラグ
   */
  recordCacheAccess(hit) {
    if (!MONITORING_CONFIG.ENABLED) return;
    
    if (hit) {
      this.metrics.cacheHits++;
    } else {
      this.metrics.cacheMisses++;
    }
  }

  /**
   * 現在のメトリクスを取得
   * @returns {object} メトリクス情報
   */
  getCurrentMetrics() {
    const uptime = Date.now() - this.metrics.lastReset;
    const errorRate = this.metrics.requests > 0 ? 
      (this.metrics.errors / this.metrics.requests * 100).toFixed(2) : 0;
    const avgResponseTime = this.metrics.requests > 0 ? 
      (this.metrics.totalResponseTime / this.metrics.requests).toFixed(0) : 0;
    const cacheHitRate = (this.metrics.cacheHits + this.metrics.cacheMisses) > 0 ? 
      (this.metrics.cacheHits / (this.metrics.cacheHits + this.metrics.cacheMisses) * 100).toFixed(2) : 0;

    return {
      uptime: uptime,
      requests: this.metrics.requests,
      errors: this.metrics.errors,
      errorRate: errorRate + '%',
      avgResponseTime: avgResponseTime + 'ms',
      apiCalls: this.metrics.apiCalls,
      cacheHitRate: cacheHitRate + '%',
      status: this.getSystemStatus(),
      alerts: this.alerts.length,
      lastUpdate: new Date().toISOString()
    };
  }

  /**
   * システム状態を判定
   * @returns {string} 状態
   */
  getSystemStatus() {
    const errorRate = this.metrics.requests > 0 ? 
      (this.metrics.errors / this.metrics.requests * 100) : 0;
    const avgResponseTime = this.metrics.requests > 0 ? 
      (this.metrics.totalResponseTime / this.metrics.requests) : 0;

    if (errorRate > MONITORING_CONFIG.ALERT_THRESHOLDS.ERROR_RATE) {
      return 'critical';
    }
    
    if (avgResponseTime > MONITORING_CONFIG.ALERT_THRESHOLDS.RESPONSE_TIME) {
      return 'warning';
    }
    
    if (errorRate > 5 || avgResponseTime > 3000) {
      return 'degraded';
    }
    
    return 'healthy';
  }

  /**
   * アラートをチェック
   * @private
   */
  checkAlerts(operation, responseTime, success) {
    const now = Date.now();
    
    // レスポンス時間アラート
    if (responseTime > MONITORING_CONFIG.ALERT_THRESHOLDS.RESPONSE_TIME) {
      this.addAlert({
        type: 'SLOW_RESPONSE',
        message: `操作 "${operation}" のレスポンス時間が ${responseTime}ms と閾値を超えました`,
        timestamp: now,
        severity: 'warning',
        operation: operation,
        value: responseTime
      });
    }
    
    // エラー率アラート
    const errorRate = this.metrics.requests > 0 ? 
      (this.metrics.errors / this.metrics.requests * 100) : 0;
    
    if (errorRate > MONITORING_CONFIG.ALERT_THRESHOLDS.ERROR_RATE) {
      this.addAlert({
        type: 'HIGH_ERROR_RATE',
        message: `エラー率が ${errorRate.toFixed(2)}% と閾値を超えました`,
        timestamp: now,
        severity: 'critical',
        value: errorRate
      });
    }
  }

  /**
   * アラートを追加
   * @private
   */
  addAlert(alert) {
    // 重複アラートを防ぐ（直近5分以内の同種アラートは無視）
    const recentSimilarAlert = this.alerts.find(a => 
      a.type === alert.type && 
      (Date.now() - a.timestamp) < 300000 // 5分
    );
    
    if (recentSimilarAlert) {
      return;
    }
    
    this.alerts.unshift(alert);
    
    // アラート数制限
    if (this.alerts.length > 50) {
      this.alerts = this.alerts.slice(0, 50);
    }
    
    console.warn(`🚨 [Alert] ${alert.type}: ${alert.message}`);
    
    // 重要なアラートは即座にログ記録
    if (alert.severity === 'critical') {
      this.logSecurityEvent({
        type: 'PERFORMANCE_ALERT',
        severity: alert.severity,
        details: alert,
        timestamp: Date.now()
      });
    }
  }

  /**
   * メトリクスログを記録
   * @private
   */
  logMetric(metric) {
    try {
      const props = PropertiesService.getScriptProperties();
      const logsKey = 'performance_logs';
      
      let logs = [];
      try {
        const existingLogs = props.getProperty(logsKey);
        if (existingLogs) {
          logs = JSON.parse(existingLogs);
        }
      } catch (parseError) {
        console.warn('既存ログの読み込みに失敗:', parseError.message);
      }
      
      logs.unshift(metric);
      
      // ログ数制限
      if (logs.length > MONITORING_CONFIG.MAX_LOG_ENTRIES) {
        logs = logs.slice(0, MONITORING_CONFIG.MAX_LOG_ENTRIES);
      }
      
      // 古いログを削除（保持期間を超えたもの）
      const cutoffTime = Date.now() - (MONITORING_CONFIG.RETENTION_HOURS * 60 * 60 * 1000);
      logs = logs.filter(log => log.timestamp > cutoffTime);
      
      props.setProperty(logsKey, JSON.stringify(logs));
      
    } catch (error) {
      console.error('メトリクスログの記録に失敗:', error.message);
    }
  }

  /**
   * セキュリティイベントをログ記録
   * @private
   */
  logSecurityEvent(event) {
    try {
      const props = PropertiesService.getScriptProperties();
      const securityLogsKey = 'security_logs';
      
      let logs = [];
      try {
        const existingLogs = props.getProperty(securityLogsKey);
        if (existingLogs) {
          logs = JSON.parse(existingLogs);
        }
      } catch (parseError) {
        console.warn('セキュリティログの読み込みに失敗:', parseError.message);
      }
      
      logs.unshift(event);
      
      // ログ数制限
      if (logs.length > 500) {
        logs = logs.slice(0, 500);
      }
      
      props.setProperty(securityLogsKey, JSON.stringify(logs));
      
      console.log(`🔒 [Security] ${event.type}: ${JSON.stringify(event.details)}`);
      
    } catch (error) {
      console.error('セキュリティログの記録に失敗:', error.message);
    }
  }

  /**
   * 定期的なメトリクス収集をスケジュール
   * @private
   */
  scheduleMetricsCollection() {
    // GAS環境では真の定期実行は制限されるため、
    // 代わりに主要な操作時にメトリクスを記録する
    
    try {
      // システム状態を記録
      this.logMetric({
        timestamp: Date.now(),
        type: 'system_status',
        metrics: this.getCurrentMetrics()
      });
      
    } catch (error) {
      console.error('定期メトリクス収集エラー:', error.message);
    }
  }

  /**
   * パフォーマンスレポートを生成
   * @returns {string} レポート
   */
  generateReport() {
    const metrics = this.getCurrentMetrics();
    const alerts = this.alerts.slice(0, 10); // 最新10件のアラート
    
    let report = '=== パフォーマンス監視レポート ===\n';
    report += `生成日時: ${new Date().toLocaleString()}\n`;
    report += `システム状態: ${this.getStatusIcon(metrics.status)} ${metrics.status}\n\n`;
    
    report += '📊 メトリクス:\n';
    report += `  リクエスト数: ${metrics.requests}\n`;
    report += `  エラー数: ${metrics.errors}\n`;
    report += `  エラー率: ${metrics.errorRate}\n`;
    report += `  平均レスポンス時間: ${metrics.avgResponseTime}\n`;
    report += `  API呼び出し数: ${metrics.apiCalls}\n`;
    report += `  キャッシュヒット率: ${metrics.cacheHitRate}\n`;
    report += `  稼働時間: ${Math.floor(metrics.uptime / 60000)}分\n\n`;
    
    if (alerts.length > 0) {
      report += '🚨 最近のアラート:\n';
      alerts.forEach(alert => {
        const timeAgo = Math.floor((Date.now() - alert.timestamp) / 60000);
        report += `  [${timeAgo}分前] ${alert.type}: ${alert.message}\n`;
      });
    } else {
      report += '✅ アラートはありません\n';
    }
    
    return report;
  }

  /**
   * 状態アイコンを取得
   * @private
   */
  getStatusIcon(status) {
    switch (status) {
      case 'healthy': return '✅';
      case 'degraded': return '⚠️';
      case 'warning': return '🟡';
      case 'critical': return '🔴';
      default: return '❓';
    }
  }

  /**
   * メトリクスをリセット
   */
  resetMetrics() {
    this.metrics = {
      requests: 0,
      errors: 0,
      totalResponseTime: 0,
      apiCalls: 0,
      cacheHits: 0,
      cacheMisses: 0,
      lastReset: Date.now()
    };
    
    this.alerts = [];
    console.log('📊 メトリクスをリセットしました');
  }
}

/**
 * セキュリティ監査ログマネージャー
 */
class SecurityAuditLogger {
  constructor() {
    this.eventTypes = {
      USER_ACCESS: 'user_access',
      DATA_MODIFICATION: 'data_modification',
      AUTHENTICATION: 'authentication',
      PERMISSION_CHANGE: 'permission_change',
      SUSPICIOUS_ACTIVITY: 'suspicious_activity',
      SYSTEM_EVENT: 'system_event'
    };
  }

  /**
   * アクセスイベントをログ記録
   * @param {string} userId - ユーザーID
   * @param {string} action - アクション
   * @param {object} details - 詳細情報
   */
  logAccess(userId, action, details = {}) {
    this.logEvent(this.eventTypes.USER_ACCESS, {
      userId: userId,
      action: action,
      userAgent: details.userAgent || 'unknown',
      ipAddress: details.ipAddress || 'unknown',
      sessionId: details.sessionId || Session.getTemporaryActiveUserKey(),
      ...details
    });
  }

  /**
   * データ変更イベントをログ記録
   * @param {string} userId - ユーザーID
   * @param {string} dataType - データ種別
   * @param {string} operation - 操作（create, update, delete）
   * @param {object} details - 詳細情報
   */
  logDataModification(userId, dataType, operation, details = {}) {
    this.logEvent(this.eventTypes.DATA_MODIFICATION, {
      userId: userId,
      dataType: dataType,
      operation: operation,
      targetId: details.targetId,
      changes: details.changes || {},
      ...details
    });
  }

  /**
   * 認証イベントをログ記録
   * @param {string} userId - ユーザーID
   * @param {string} authResult - 認証結果（success, failure, logout）
   * @param {object} details - 詳細情報
   */
  logAuthentication(userId, authResult, details = {}) {
    this.logEvent(this.eventTypes.AUTHENTICATION, {
      userId: userId,
      result: authResult,
      method: details.method || 'google_oauth',
      ...details
    });
  }

  /**
   * 権限変更イベントをログ記録
   * @param {string} executorId - 実行者ID
   * @param {string} targetUserId - 対象ユーザーID
   * @param {string} permissionChange - 権限変更内容
   * @param {object} details - 詳細情報
   */
  logPermissionChange(executorId, targetUserId, permissionChange, details = {}) {
    this.logEvent(this.eventTypes.PERMISSION_CHANGE, {
      executorId: executorId,
      targetUserId: targetUserId,
      change: permissionChange,
      reason: details.reason || '',
      ...details
    });
  }

  /**
   * 不審な活動をログ記録
   * @param {string} userId - ユーザーID
   * @param {string} suspiciousActivity - 不審な活動の内容
   * @param {object} details - 詳細情報
   */
  logSuspiciousActivity(userId, suspiciousActivity, details = {}) {
    this.logEvent(this.eventTypes.SUSPICIOUS_ACTIVITY, {
      userId: userId,
      activity: suspiciousActivity,
      riskLevel: details.riskLevel || 'medium',
      automaticAction: details.automaticAction || 'none',
      ...details
    });
  }

  /**
   * セキュリティイベントをログ記録
   * @param {string} eventType - イベント種別
   * @param {object} eventData - イベントデータ
   */
  logEvent(eventType, eventData) {
    try {
      const event = {
        timestamp: Date.now(),
        eventType: eventType,
        data: eventData,
        userEmail: Session.getActiveUser().getEmail(),
        id: Utilities.getUuid()
      };
      
      const props = PropertiesService.getScriptProperties();
      const auditLogsKey = 'security_audit_logs';
      
      let logs = [];
      try {
        const existingLogs = props.getProperty(auditLogsKey);
        if (existingLogs) {
          logs = JSON.parse(existingLogs);
        }
      } catch (parseError) {
        console.warn('監査ログの読み込みに失敗:', parseError.message);
      }
      
      logs.unshift(event);
      
      // ログ数制限（セキュリティログは多めに保持）
      if (logs.length > 2000) {
        logs = logs.slice(0, 2000);
      }
      
      // 古いログを削除（90日以上）
      const cutoffTime = Date.now() - (90 * 24 * 60 * 60 * 1000);
      logs = logs.filter(log => log.timestamp > cutoffTime);
      
      props.setProperty(auditLogsKey, JSON.stringify(logs));
      
      console.log(`🔒 [SecurityAudit] ${eventType}: ${JSON.stringify(eventData, null, 2)}`);
      
    } catch (error) {
      console.error('セキュリティ監査ログの記録に失敗:', error.message);
    }
  }

  /**
   * 監査ログを取得
   * @param {object} filters - フィルター条件
   * @returns {array} 監査ログ
   */
  getAuditLogs(filters = {}) {
    try {
      const props = PropertiesService.getScriptProperties();
      const auditLogsKey = 'security_audit_logs';
      
      let logs = [];
      const existingLogs = props.getProperty(auditLogsKey);
      if (existingLogs) {
        logs = JSON.parse(existingLogs);
      }
      
      // フィルター適用
      if (filters.eventType) {
        logs = logs.filter(log => log.eventType === filters.eventType);
      }
      
      if (filters.userId) {
        logs = logs.filter(log => 
          log.data.userId === filters.userId || 
          log.data.executorId === filters.userId ||
          log.data.targetUserId === filters.userId
        );
      }
      
      if (filters.startTime) {
        logs = logs.filter(log => log.timestamp >= filters.startTime);
      }
      
      if (filters.endTime) {
        logs = logs.filter(log => log.timestamp <= filters.endTime);
      }
      
      // 最大1000件に制限
      if (logs.length > 1000) {
        logs = logs.slice(0, 1000);
      }
      
      return logs;
      
    } catch (error) {
      console.error('監査ログの取得に失敗:', error.message);
      return [];
    }
  }

  /**
   * セキュリティレポートを生成
   * @param {number} hours - 過去何時間のデータを対象にするか
   * @returns {string} レポート
   */
  generateSecurityReport(hours = 24) {
    try {
      const startTime = Date.now() - (hours * 60 * 60 * 1000);
      const logs = this.getAuditLogs({ startTime: startTime });
      
      let report = `=== セキュリティ監査レポート（過去${hours}時間） ===\n`;
      report += `生成日時: ${new Date().toLocaleString()}\n`;
      report += `対象イベント数: ${logs.length}\n\n`;
      
      // イベント種別別の集計
      const eventCounts = {};
      logs.forEach(log => {
        eventCounts[log.eventType] = (eventCounts[log.eventType] || 0) + 1;
      });
      
      report += '📊 イベント種別別集計:\n';
      for (const [eventType, count] of Object.entries(eventCounts)) {
        report += `  ${eventType}: ${count}件\n`;
      }
      report += '\n';
      
      // 不審な活動
      const suspiciousLogs = logs.filter(log => log.eventType === this.eventTypes.SUSPICIOUS_ACTIVITY);
      if (suspiciousLogs.length > 0) {
        report += '🚨 不審な活動:\n';
        suspiciousLogs.slice(0, 10).forEach(log => {
          const timeAgo = Math.floor((Date.now() - log.timestamp) / 60000);
          report += `  [${timeAgo}分前] ${log.data.activity} (User: ${log.data.userId})\n`;
        });
        report += '\n';
      }
      
      // 認証失敗
      const authFailures = logs.filter(log => 
        log.eventType === this.eventTypes.AUTHENTICATION && 
        log.data.result === 'failure'
      );
      if (authFailures.length > 0) {
        report += '❌ 認証失敗:\n';
        authFailures.slice(0, 5).forEach(log => {
          const timeAgo = Math.floor((Date.now() - log.timestamp) / 60000);
          report += `  [${timeAgo}分前] ${log.data.userId}\n`;
        });
        report += '\n';
      }
      
      return report;
      
    } catch (error) {
      return `セキュリティレポートの生成に失敗しました: ${error.message}`;
    }
  }
}

// シングルトンインスタンス
const performanceMonitor = new PerformanceMonitor();
const securityAuditLogger = new SecurityAuditLogger();

// 監視開始
performanceMonitor.start();