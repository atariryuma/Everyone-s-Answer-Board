/**
 * @fileoverview セキュリティヘルスチェック統合システム
 * 全セキュリティコンポーネントの健全性を統一的に監視
 */

/**
 * 統合セキュリティヘルスチェック実行
 * @returns {Promise<object>} 包括的なヘルスチェック結果
 */
function performComprehensiveSecurityHealthCheck() {
  const startTime = Date.now();
  
  const healthCheckResult = {
    timestamp: new Date().toISOString(),
    executionTime: 0,
    overallStatus: 'UNKNOWN',
    components: {},
    criticalIssues: [],
    warnings: [],
    recommendations: []
  };

  try {
    infoLog('🔒 統合セキュリティヘルスチェック開始');

    // 1. 統一秘密情報管理システム
    try {
      const secretManagerHealth = unifiedSecretManager.performHealthCheck();
      healthCheckResult.components.secretManager = secretManagerHealth;
      
      if (secretManagerHealth.issues.length > 0) {
        healthCheckResult.warnings.push(`秘密情報管理: ${secretManagerHealth.issues.length}件の問題`);
      }
    } catch (error) {
      healthCheckResult.components.secretManager = { 
        status: 'ERROR', 
        error: error.message 
      };
      healthCheckResult.criticalIssues.push(`秘密情報管理システムエラー: ${error.message}`);
    }

    // 2. マルチテナントセキュリティ
    try {
      const multiTenantHealth = performMultiTenantHealthCheck();
      healthCheckResult.components.multiTenantSecurity = multiTenantHealth;
      
      if (multiTenantHealth.issues.length > 0) {
        healthCheckResult.criticalIssues.push(`マルチテナントセキュリティ: ${multiTenantHealth.issues.length}件の問題`);
      }
    } catch (error) {
      healthCheckResult.components.multiTenantSecurity = { 
        status: 'ERROR', 
        error: error.message 
      };
      healthCheckResult.criticalIssues.push(`マルチテナントセキュリティエラー: ${error.message}`);
    }

    // 3. 統一バッチ処理システム
    try {
      const batchProcessorHealth = performUnifiedBatchHealthCheck();
      healthCheckResult.components.batchProcessor = batchProcessorHealth;
      
      if (batchProcessorHealth.issues.length > 0) {
        healthCheckResult.warnings.push(`バッチ処理: ${batchProcessorHealth.issues.length}件の問題`);
      }
    } catch (error) {
      healthCheckResult.components.batchProcessor = { 
        status: 'ERROR', 
        error: error.message 
      };
      healthCheckResult.warnings.push(`バッチ処理システムエラー: ${error.message}`);
    }

    // 4. 回復力のある実行機構
    try {
      const resilientExecutorStats = resilientExecutor.getStats();
      healthCheckResult.components.resilientExecutor = {
        status: 'OK',
        stats: resilientExecutorStats,
        circuitBreakerState: resilientExecutorStats.circuitBreakerState
      };

      if (resilientExecutorStats.circuitBreakerState === 'OPEN') {
        healthCheckResult.criticalIssues.push('回復力のある実行機構: 回路ブレーカーがOPEN状態');
      }

      const successRate = parseFloat(resilientExecutorStats.successRate.replace('%', ''));
      if (successRate < 95) {
        healthCheckResult.warnings.push(`回復力のある実行機構: 成功率が低下 (${resilientExecutorStats.successRate})`);
      }
    } catch (error) {
      healthCheckResult.components.resilientExecutor = { 
        status: 'ERROR', 
        error: error.message 
      };
      healthCheckResult.warnings.push(`回復力のある実行機構エラー: ${error.message}`);
    }

    // 5. 認証システム
    try {
      const authHealth = checkAuthenticationHealth();
      healthCheckResult.components.authentication = authHealth;
      
      if (!authHealth.serviceAccountValid) {
        healthCheckResult.criticalIssues.push('認証システム: サービスアカウント認証情報が無効');
      }
    } catch (error) {
      healthCheckResult.components.authentication = { 
        status: 'ERROR', 
        error: error.message 
      };
      healthCheckResult.criticalIssues.push(`認証システムエラー: ${error.message}`);
    }

    // 6. データベース接続性
    try {
      const dbHealth = checkDatabaseHealth();
      healthCheckResult.components.database = dbHealth;
      
      if (!dbHealth.accessible) {
        healthCheckResult.criticalIssues.push('データベース: 接続不可');
      }
    } catch (error) {
      healthCheckResult.components.database = { 
        status: 'ERROR', 
        error: error.message 
      };
      healthCheckResult.criticalIssues.push(`データベース接続エラー: ${error.message}`);
    }

    // 総合ステータス判定
    if (healthCheckResult.criticalIssues.length > 0) {
      healthCheckResult.overallStatus = 'CRITICAL';
      healthCheckResult.recommendations.push('重要な問題があります。即座に対処してください。');
    } else if (healthCheckResult.warnings.length > 0) {
      healthCheckResult.overallStatus = 'WARNING';
      healthCheckResult.recommendations.push('警告があります。確認と対処を推奨します。');
    } else {
      healthCheckResult.overallStatus = 'HEALTHY';
      healthCheckResult.recommendations.push('全てのセキュリティコンポーネントが正常に動作しています。');
    }

    // 実行時間を記録
    healthCheckResult.executionTime = Date.now() - startTime;

    // 結果をログ出力
    if (healthCheckResult.overallStatus === 'CRITICAL') {
      errorLog('🚨 統合セキュリティヘルスチェック: 重要な問題が検出されました', healthCheckResult);
    } else if (healthCheckResult.overallStatus === 'WARNING') {
      warnLog('⚠️ 統合セキュリティヘルスチェック: 警告があります', healthCheckResult);
    } else {
      infoLog('✅ 統合セキュリティヘルスチェック: 正常', healthCheckResult);
    }

    return healthCheckResult;

  } catch (error) {
    healthCheckResult.overallStatus = 'ERROR';
    healthCheckResult.criticalIssues.push(`ヘルスチェック実行エラー: ${error.message}`);
    healthCheckResult.executionTime = Date.now() - startTime;
    
    errorLog('❌ 統合セキュリティヘルスチェック実行エラー:', error);
    return healthCheckResult;
  }
}

/**
 * 認証システムのヘルスチェック
 * @private
 */
function checkAuthenticationHealth() {
  const authHealth = {
    status: 'UNKNOWN',
    serviceAccountValid: false,
    tokenGenerationWorking: false,
    lastTokenGeneration: null,
    issues: []
  };

  try {
    // サービスアカウント認証情報の検証
    const serviceAccountCreds = getSecureServiceAccountCreds();
    if (serviceAccountCreds && serviceAccountCreds.client_email && serviceAccountCreds.private_key) {
      authHealth.serviceAccountValid = true;
    } else {
      authHealth.issues.push('サービスアカウント認証情報が不完全');
    }

    // トークン生成テスト
    try {
      const testToken = getServiceAccountTokenCached();
      if (testToken && testToken.length > 0) {
        authHealth.tokenGenerationWorking = true;
        authHealth.lastTokenGeneration = new Date().toISOString();
      }
    } catch (tokenError) {
      authHealth.issues.push(`トークン生成エラー: ${tokenError.message}`);
    }

    authHealth.status = authHealth.issues.length === 0 ? 'OK' : 'WARNING';

  } catch (error) {
    authHealth.status = 'ERROR';
    authHealth.issues.push(`認証ヘルスチェックエラー: ${error.message}`);
  }

  return authHealth;
}

/**
 * データベース接続性のヘルスチェック
 * @private
 */
function checkDatabaseHealth() {
  const dbHealth = {
    status: 'UNKNOWN',
    accessible: false,
    databaseId: null,
    sheetsServiceWorking: false,
    lastAccess: null,
    issues: []
  };

  try {
    // データベースID取得
    const dbId = getSecureDatabaseId();
    if (dbId) {
      dbHealth.databaseId = dbId.substring(0, 10) + '...';
    } else {
      dbHealth.issues.push('データベースIDが設定されていません');
      dbHealth.status = 'ERROR';
      return dbHealth;
    }

    // Sheets Service接続テスト
    try {
      const service = getSheetsServiceCached();
      if (service) {
        dbHealth.sheetsServiceWorking = true;

        // 軽量なデータベースアクセステスト
        const testResponse = resilientExecutor.execute(
          async () => {
            return resilientUrlFetch(
              `${service.baseUrl}/${encodeURIComponent(dbId)}`,
              {
                method: 'GET',
                headers: {
                  'Authorization': `Bearer ${service.accessToken}`
                }
              }
            );
          },
          {
            name: 'DatabaseHealthCheck',
            idempotent: true
          }
        );

        if (testResponse && testResponse.getResponseCode() === 200) {
          dbHealth.accessible = true;
          dbHealth.lastAccess = new Date().toISOString();
        } else {
          dbHealth.issues.push(`データベースアクセステスト失敗: ${testResponse ? testResponse.getResponseCode() : 'No response'}`);
        }
      }
    } catch (serviceError) {
      dbHealth.issues.push(`Sheets Service エラー: ${serviceError.message}`);
    }

    dbHealth.status = dbHealth.accessible ? 'OK' : 'ERROR';

  } catch (error) {
    dbHealth.status = 'ERROR';
    dbHealth.issues.push(`データベースヘルスチェックエラー: ${error.message}`);
  }

  return dbHealth;
}

/**
 * セキュリティメトリクスのサマリー取得
 * @returns {object} セキュリティメトリクス
 */
function getSecurityMetrics() {
  const metrics = {
    timestamp: new Date().toISOString(),
    resilientExecutor: null,
    secretManager: null,
    multiTenantSecurity: null,
    batchProcessor: null
  };

  try {
    // 回復力のある実行機構のメトリクス
    if (typeof resilientExecutor !== 'undefined') {
      metrics.resilientExecutor = resilientExecutor.getStats();
    }

    // 統一バッチ処理のメトリクス
    if (typeof unifiedBatchProcessor !== 'undefined') {
      metrics.batchProcessor = unifiedBatchProcessor.getMetrics();
    }

    // マルチテナントセキュリティの統計
    if (typeof multiTenantSecurity !== 'undefined') {
      metrics.multiTenantSecurity = multiTenantSecurity.getSecurityStats();
    }

    // 秘密情報管理の監査ログサマリー
    if (typeof unifiedSecretManager !== 'undefined') {
      const auditLog = unifiedSecretManager.getAuditLog();
      metrics.secretManager = {
        auditLogEntries: auditLog.length,
        lastAccess: auditLog.length > 0 ? auditLog[auditLog.length - 1].timestamp : null
      };
    }

  } catch (error) {
    warnLog('セキュリティメトリクス取得エラー:', error.message);
  }

  return metrics;
}

/**
 * 定期的なセキュリティヘルスチェックの設定
 * @param {number} intervalMinutes - チェック間隔（分）
 */
function scheduleSecurityHealthCheck(intervalMinutes = 60) {
  try {
    // GAS環境では時間ベースのトリガーを使用
    const triggers = ScriptApp.getProjectTriggers();
    
    // 既存のヘルスチェック トリガーを削除
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'runScheduledSecurityHealthCheck') {
        ScriptApp.deleteTrigger(trigger);
      }
    });

    // 新しいトリガーを作成
    ScriptApp.newTrigger('runScheduledSecurityHealthCheck')
      .timeBased()
      .everyMinutes(intervalMinutes)
      .create();

    infoLog(`📅 定期セキュリティヘルスチェックを${intervalMinutes}分間隔で設定しました`);

  } catch (error) {
    errorLog('定期セキュリティヘルスチェック設定エラー:', error.message);
  }
}

/**
 * 定期実行される関数（トリガーから呼び出し）
 */
function runScheduledSecurityHealthCheck() {
  try {
    const healthResult = performComprehensiveSecurityHealthCheck();
    
    // 重要な問題がある場合は管理者に通知
    if (healthResult.overallStatus === 'CRITICAL') {
      // フューチャー実装: 管理者への緊急通知
      errorLog('🚨 緊急: セキュリティヘルスチェックで重要な問題を検出', healthResult);
    }

    // 結果をログに記録
    infoLog('🔒 定期セキュリティヘルスチェック完了', {
      status: healthResult.overallStatus,
      criticalIssues: healthResult.criticalIssues.length,
      warnings: healthResult.warnings.length,
      executionTime: healthResult.executionTime
    });

  } catch (error) {
    errorLog('定期セキュリティヘルスチェック実行エラー:', error.message);
  }
}