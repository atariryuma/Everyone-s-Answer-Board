/**
 * @fileoverview システム統合管理 - 全最適化コンポーネントの統合制御
 * 安定性・高速化・マルチテナントセキュリティの統一管理
 */

/**
 * 統合システム管理クラス
 * 全ての最適化コンポーネントのライフサイクルを統一管理
 */
class SystemIntegrationManager {
  constructor() {
    this.initialized = false;
    this.components = {
      resilientExecutor: { status: 'NOT_INITIALIZED', instance: null },
      unifiedBatchProcessor: { status: 'NOT_INITIALIZED', instance: null },
      multiTenantSecurity: { status: 'NOT_INITIALIZED', instance: null },
      unifiedSecretManager: { status: 'NOT_INITIALIZED', instance: null },
      cacheManager: { status: 'NOT_INITIALIZED', instance: null }
    };

    this.initializationOrder = [
      'unifiedSecretManager',
      'resilientExecutor', 
      'multiTenantSecurity',
      'cacheManager',
      'unifiedBatchProcessor'
    ];

    this.systemMetrics = {
      startupTime: null,
      lastHealthCheck: null,
      totalRequests: 0,
      errorRate: 0,
      averageResponseTime: 0
    };
  }

  /**
   * システム全体の初期化
   * @param {object} options - 初期化オプション
   * @returns {Promise<object>} 初期化結果
   */
  async initializeSystem(options = {}) {
    const startTime = Date.now();
    const initResult = {
      success: false,
      timestamp: new Date().toISOString(),
      initializationTime: 0,
      componentsInitialized: [],
      componentsFailedToInitialize: [],
      warnings: [],
      errors: []
    };

    try {
      infoLog('🚀 統合システム初期化開始');

      // 依存関係の順序で各コンポーネントを初期化
      for (const componentName of this.initializationOrder) {
        try {
          this.initializeComponent(componentName, options);
          initResult.componentsInitialized.push(componentName);
          infoLog(`✅ ${componentName} 初期化完了`);
        } catch (error) {
          initResult.componentsFailedToInitialize.push({
            component: componentName,
            error: error.message
          });
          initResult.errors.push(`${componentName} 初期化失敗: ${error.message}`);
          warnLog(`❌ ${componentName} 初期化失敗:`, error.message);

          // 重要なコンポーネントが失敗した場合は初期化を中断
          if (this.isCriticalComponent(componentName)) {
            throw new Error(`重要コンポーネント ${componentName} の初期化に失敗`);
          }
        }
      }

      // 初期ヘルスチェック実行
      try {
        const healthResult = this.performInitialHealthCheck();
        initResult.initialHealthCheck = healthResult;
        
        if (healthResult.overallStatus === 'CRITICAL') {
          initResult.warnings.push('初期ヘルスチェックで重要な問題を検出');
        }
      } catch (healthError) {
        initResult.warnings.push(`初期ヘルスチェック失敗: ${healthError.message}`);
      }

      // システム統計の初期化
      this.systemMetrics.startupTime = new Date().toISOString();
      this.initialized = true;

      initResult.success = initResult.componentsFailedToInitialize.length === 0;
      initResult.initializationTime = Date.now() - startTime;

      // 定期ヘルスチェックのスケジュール
      if (options.enablePeriodicHealthCheck !== false) {
        this.schedulePeriodicTasks();
      }

      infoLog(`🎉 統合システム初期化完了 (${initResult.initializationTime}ms)`, {
        success: initResult.success,
        initialized: initResult.componentsInitialized.length,
        failed: initResult.componentsFailedToInitialize.length
      });

      return initResult;

    } catch (error) {
      initResult.success = false;
      initResult.errors.push(`システム初期化エラー: ${error.message}`);
      initResult.initializationTime = Date.now() - startTime;
      
      errorLog('❌ 統合システム初期化エラー:', error);
      return initResult;
    }
  }

  /**
   * 個別コンポーネントの初期化
   * @private
   */
  async initializeComponent(componentName, options) {
    switch (componentName) {
      case 'unifiedSecretManager':
        if (typeof unifiedSecretManager !== 'undefined') {
          // 秘密情報管理の健全性チェック
          const healthCheck = unifiedSecretManager.performHealthCheck();
          if (healthCheck.criticalSecretsStatus === 'ERROR') {
            throw new Error('重要な秘密情報が見つかりません');
          }
          this.components.unifiedSecretManager.status = 'INITIALIZED';
          this.components.unifiedSecretManager.instance = unifiedSecretManager;
        }
        break;

      case 'resilientExecutor':
        if (typeof resilientExecutor !== 'undefined') {
          // 統計をリセットして初期化
          resilientExecutor.resetStats();
          this.components.resilientExecutor.status = 'INITIALIZED';
          this.components.resilientExecutor.instance = resilientExecutor;
        }
        break;

      case 'multiTenantSecurity':
        if (typeof multiTenantSecurity !== 'undefined') {
          // マルチテナントセキュリティの初期検証
          const testResult = multiTenantSecurity.validateTenantBoundary(
            'test@example.com', 
            'test@example.com', 
            'init_test'
          );
          if (!testResult) {
            throw new Error('マルチテナントセキュリティの初期検証に失敗');
          }
          this.components.multiTenantSecurity.status = 'INITIALIZED';
          this.components.multiTenantSecurity.instance = multiTenantSecurity;
        }
        break;

      case 'cacheManager':
        if (typeof cacheManager !== 'undefined') {
          // キャッシュマネージャーの初期化確認
          try {
            cacheManager.get('init_test', () => 'test_value');
            this.components.cacheManager.status = 'INITIALIZED';
            this.components.cacheManager.instance = cacheManager;
          } catch (error) {
            throw new Error(`キャッシュマネージャー初期化失敗: ${error.message}`);
          }
        }
        break;

      case 'unifiedBatchProcessor':
        if (typeof unifiedBatchProcessor !== 'undefined') {
          // バッチ処理システムの統計リセット
          unifiedBatchProcessor.resetMetrics();
          this.components.unifiedBatchProcessor.status = 'INITIALIZED';
          this.components.unifiedBatchProcessor.instance = unifiedBatchProcessor;
        }
        break;

      default:
        throw new Error(`未知のコンポーネント: ${componentName}`);
    }
  }

  /**
   * 重要コンポーネントの判定
   * @private
   */
  isCriticalComponent(componentName) {
    const criticalComponents = [
      'unifiedSecretManager',
      'resilientExecutor'
    ];
    return criticalComponents.includes(componentName);
  }

  /**
   * 初期ヘルスチェック実行
   * @private
   */
  async performInitialHealthCheck() {
    try {
      // セキュリティヘルスチェックが利用可能な場合は実行
      if (typeof performComprehensiveSecurityHealthCheck !== 'undefined') {
        return performComprehensiveSecurityHealthCheck();
      } else {
        // 基本的なヘルスチェック
        return this.performBasicHealthCheck();
      }
    } catch (error) {
      warnLog('ヘルスチェック実行エラー:', error.message);
      return {
        overallStatus: 'ERROR',
        error: error.message,
        timestamp: new Date().toISOString()
      };
    }
  }

  /**
   * 基本ヘルスチェック
   * @private
   */
  async performBasicHealthCheck() {
    const healthResult = {
      overallStatus: 'UNKNOWN',
      timestamp: new Date().toISOString(),
      components: {},
      issues: []
    };

    let healthyComponents = 0;
    let totalComponents = 0;

    for (const [name, component] of Object.entries(this.components)) {
      totalComponents++;
      if (component.status === 'INITIALIZED') {
        healthyComponents++;
        healthResult.components[name] = { status: 'OK' };
      } else {
        healthResult.components[name] = { status: 'ERROR', error: 'Not initialized' };
        healthResult.issues.push(`${name} が初期化されていません`);
      }
    }

    if (healthyComponents === totalComponents) {
      healthResult.overallStatus = 'HEALTHY';
    } else if (healthyComponents > totalComponents / 2) {
      healthResult.overallStatus = 'WARNING';
    } else {
      healthResult.overallStatus = 'CRITICAL';
    }

    return healthResult;
  }

  /**
   * 定期タスクのスケジュール
   * @private
   */
  schedulePeriodicTasks() {
    try {
      // 定期セキュリティヘルスチェックをスケジュール（1時間間隔）
      if (typeof scheduleSecurityHealthCheck !== 'undefined') {
        scheduleSecurityHealthCheck(60);
      }

      // システムメトリクス更新タスクをスケジュール
      this.scheduleMetricsUpdate();

      infoLog('📅 定期タスクをスケジュールしました');
    } catch (error) {
      warnLog('定期タスクスケジュールエラー:', error.message);
    }
  }

  /**
   * システムメトリクス更新タスクのスケジュール
   * @private
   */
  scheduleMetricsUpdate() {
    try {
      const triggers = ScriptApp.getProjectTriggers();
      
      // 既存のメトリクス更新トリガーを削除
      triggers.forEach(trigger => {
        if (trigger.getHandlerFunction() === 'updateSystemMetrics') {
          ScriptApp.deleteTrigger(trigger);
        }
      });

      // 新しいメトリクス更新トリガーを作成（15分間隔）
      ScriptApp.newTrigger('updateSystemMetrics')
        .timeBased()
        .everyMinutes(15)
        .create();

    } catch (error) {
      warnLog('メトリクス更新タスクスケジュールエラー:', error.message);
    }
  }

  /**
   * システム統計の取得
   * @returns {object} システム統計
   */
  getSystemMetrics() {
    const metrics = {
      ...this.systemMetrics,
      components: {},
      timestamp: new Date().toISOString()
    };

    // 各コンポーネントのメトリクスを収集
    for (const [name, component] of Object.entries(this.components)) {
      if (component.status === 'INITIALIZED' && component.instance) {
        try {
          switch (name) {
            case 'resilientExecutor':
              metrics.components.resilientExecutor = component.instance.getStats();
              break;
            case 'unifiedBatchProcessor':
              metrics.components.unifiedBatchProcessor = component.instance.getMetrics();
              break;
            case 'multiTenantSecurity':
              metrics.components.multiTenantSecurity = component.instance.getSecurityStats();
              break;
          }
        } catch (error) {
          warnLog(`${name} メトリクス取得エラー:`, error.message);
        }
      }
    }

    return metrics;
  }

  /**
   * システムの優雅なシャットダウン
   * @returns {Promise<boolean>} シャットダウン成功フラグ
   */
  async shutdown() {
    try {
      infoLog('🔄 システムシャットダウン開始');

      // 各コンポーネントのクリーンアップ
      for (const [name, component] of Object.entries(this.components)) {
        if (component.status === 'INITIALIZED') {
          try {
            switch (name) {
              case 'resilientExecutor':
                // 進行中の操作の完了を待機
                break;
              case 'unifiedSecretManager':
                // キャッシュクリア
                component.instance.clearSecretCache();
                break;
            }
            component.status = 'SHUTDOWN';
          } catch (error) {
            warnLog(`${name} シャットダウンエラー:`, error.message);
          }
        }
      }

      this.initialized = false;
      infoLog('✅ システムシャットダウン完了');
      return true;

    } catch (error) {
      errorLog('システムシャットダウンエラー:', error);
      return false;
    }
  }

  /**
   * システム状態の確認
   * @returns {object} システム状態
   */
  getSystemStatus() {
    return {
      initialized: this.initialized,
      components: Object.fromEntries(
        Object.entries(this.components).map(([name, component]) => [
          name, 
          { status: component.status }
        ])
      ),
      metrics: this.systemMetrics,
      timestamp: new Date().toISOString()
    };
  }
}

/**
 * グローバルなシステム統合管理インスタンス
 */
const systemIntegrationManager = new SystemIntegrationManager();

/**
 * システム初期化のエントリーポイント
 * @param {object} options - 初期化オプション
 * @returns {Promise<object>} 初期化結果
 */
function initializeOptimizedSystem(options = {}) {
  return systemIntegrationManager.initializeSystem(options);
}

/**
 * システム統計更新（トリガーから呼び出される関数）
 */
function updateSystemMetrics() {
  try {
    if (!systemIntegrationManager.initialized) {
      return;
    }

    const metrics = systemIntegrationManager.getSystemMetrics();
    systemIntegrationManager.systemMetrics.lastHealthCheck = new Date().toISOString();

    // メトリクスの集計
    let totalOperations = 0;
    let totalErrors = 0;

    if (metrics.components.resilientExecutor) {
      totalOperations += metrics.components.resilientExecutor.totalExecutions || 0;
      totalErrors += metrics.components.resilientExecutor.failedExecutions || 0;
    }

    if (metrics.components.unifiedBatchProcessor) {
      totalOperations += metrics.components.unifiedBatchProcessor.totalOperations || 0;
      totalErrors += metrics.components.unifiedBatchProcessor.failedOperations || 0;
    }

    // エラー率の計算
    systemIntegrationManager.systemMetrics.totalRequests = totalOperations;
    systemIntegrationManager.systemMetrics.errorRate = totalOperations > 0 ? 
      ((totalErrors / totalOperations) * 100).toFixed(2) + '%' : '0%';

    debugLog('📊 システムメトリクス更新完了', {
      totalRequests: systemIntegrationManager.systemMetrics.totalRequests,
      errorRate: systemIntegrationManager.systemMetrics.errorRate
    });

  } catch (error) {
    warnLog('システムメトリクス更新エラー:', error.message);
  }
}

/**
 * 手動でのシステム診断実行
 * @returns {Promise<object>} 診断結果
 */
function diagnoseOptimizedSystem() {
  const diagnostics = {
    timestamp: new Date().toISOString(),
    systemStatus: systemIntegrationManager.getSystemStatus(),
    systemMetrics: systemIntegrationManager.getSystemMetrics(),
    healthCheck: null,
    recommendations: []
  };

  try {
    // 包括的ヘルスチェック実行
    if (typeof performComprehensiveSecurityHealthCheck !== 'undefined') {
      diagnostics.healthCheck = performComprehensiveSecurityHealthCheck();
    }

    // 推奨事項の生成
    if (diagnostics.healthCheck) {
      if (diagnostics.healthCheck.overallStatus === 'CRITICAL') {
        diagnostics.recommendations.push('緊急: 重要な問題があります。システム管理者にお問い合わせください。');
      } else if (diagnostics.healthCheck.overallStatus === 'WARNING') {
        diagnostics.recommendations.push('警告: いくつかの問題が検出されました。定期メンテナンスを検討してください。');
      } else {
        diagnostics.recommendations.push('システムは正常に動作しています。');
      }
    }

    // パフォーマンス推奨事項
    const errorRate = parseFloat(diagnostics.systemMetrics.errorRate?.replace('%', '') || '0');
    if (errorRate > 5) {
      diagnostics.recommendations.push('エラー率が高めです。ログを確認してください。');
    } else if (errorRate > 1) {
      diagnostics.recommendations.push('エラー率が少し高めです。監視を継続してください。');
    }

    infoLog('🔍 システム診断完了', diagnostics);
    return diagnostics;

  } catch (error) {
    diagnostics.error = error.message;
    errorLog('システム診断エラー:', error);
    return diagnostics;
  }
}