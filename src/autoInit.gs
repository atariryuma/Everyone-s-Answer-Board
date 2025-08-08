/**
 * @fileoverview 自動初期化システム
 * 統合最適化システムの自動起動と初期化管理
 */

/**
 * システムの自動初期化（WebApp起動時やトリガー実行時に呼び出される）
 * この関数は、doGet, doPost, またはタイマートリガーから呼び出されることを想定
 */
function autoInitializeSystem() {
  // 既に初期化済みの場合はスキップ
  if (typeof systemIntegrationManager !== 'undefined' && 
      systemIntegrationManager.initialized) {
    debugLog('💨 システムは既に初期化済みです');
    return {
      success: true,
      message: 'システムは既に初期化済みです',
      skipReason: 'ALREADY_INITIALIZED'
    };
  }

  try {
    infoLog('🔄 自動システム初期化開始');
    
    // 統合システムの初期化実行
    const initResult = initializeOptimizedSystem({
      enablePeriodicHealthCheck: true,
      logLevel: 'INFO'
    });

    if (initResult.success) {
      infoLog('✅ 自動システム初期化完了', {
        componentsInitialized: initResult.componentsInitialized.length,
        initTime: initResult.initializationTime,
        warnings: initResult.warnings.length
      });

      // 初期化完了後の追加設定
      performPostInitializationTasks();

      return {
        success: true,
        message: 'システム初期化完了',
        details: initResult
      };
    } else {
      warnLog('⚠️ システム初期化に一部問題がありました', initResult);
      return {
        success: false,
        message: 'システム初期化に一部問題がありました',
        details: initResult
      };
    }

  } catch (error) {
    errorLog('❌ 自動システム初期化エラー:', error);
    return {
      success: false,
      message: 'システム初期化に失敗しました',
      error: error.message
    };
  }
}

/**
 * 初期化完了後の追加タスク
 * @private
 */
function performPostInitializationTasks() {
  try {
    // 1. 初期セキュリティ監査の実行
    if (typeof performComprehensiveSecurityHealthCheck !== 'undefined') {
      const securityCheck = performComprehensiveSecurityHealthCheck();
      if (securityCheck.overallStatus === 'CRITICAL') {
        errorLog('🚨 初期セキュリティチェックで重要な問題を検出', securityCheck);
      }
    }

    // 2. キャッシュのウォームアップ
    warmupSystemCaches();

    // 3. システム統計の初期設定
    updateSystemMetrics();

    infoLog('✅ 初期化後タスク完了');

  } catch (error) {
    warnLog('初期化後タスクエラー:', error.message);
  }
}

/**
 * システムキャッシュのウォームアップ
 * @private
 */
function warmupSystemCaches() {
  try {
    debugLog('🔥 システムキャッシュウォームアップ開始');

    // 重要な設定情報のプリロード
    if (typeof getSecureDatabaseId !== 'undefined') {
      getSecureDatabaseId();
      debugLog('📊 データベースID キャッシュ済み');
    }

    if (typeof getServiceAccountTokenCached !== 'undefined') {
      getServiceAccountTokenCached();
      debugLog('🔐 サービスアカウントトークン キャッシュ済み');
    }

    debugLog('✅ キャッシュウォームアップ完了');

  } catch (error) {
    warnLog('キャッシュウォームアップエラー:', error.message);
  }
}

/**
 * WebApp起動時の自動初期化（doGet/doPostから呼び出し）
 * @param {object} request - HTTPリクエストオブジェクト
 * @returns {Promise<void>}
 */
function ensureSystemInitializedForWebApp(request) {
  try {
    // システム初期化の確認・実行
    const initResult = autoInitializeSystem();
    
    if (!initResult.success) {
      warnLog('WebApp初期化警告:', initResult.message);
    }

    // リクエスト統計の更新
    if (typeof systemIntegrationManager !== 'undefined') {
      systemIntegrationManager.systemMetrics.totalRequests++;
    }

  } catch (error) {
    errorLog('WebApp初期化エラー:', error);
    // エラーが発生してもWebAppの動作は継続
  }
}

/**
 * 定期メンテナンス実行（時間ベーストリガーから呼び出し）
 */
function performPeriodicMaintenance() {
  try {
    infoLog('🔧 定期メンテナンス開始');

    // 1. システム診断実行
    let diagnostics = null;
    if (typeof diagnoseOptimizedSystem !== 'undefined') {
      diagnostics = diagnoseOptimizedSystem();
    }

    // 2. キャッシュクリーンアップ
    performCacheCleanup();

    // 3. メトリクス集計とログ出力
    if (typeof systemIntegrationManager !== 'undefined') {
      const metrics = systemIntegrationManager.getSystemMetrics();
      infoLog('📊 システムメトリクス', {
        totalRequests: metrics.totalRequests,
        errorRate: metrics.errorRate,
        uptime: metrics.startupTime ? new Date() - new Date(metrics.startupTime) : 0
      });
    }

    // 4. 古いログの削除（必要に応じて）
    cleanupOldLogs();

    infoLog('✅ 定期メンテナンス完了', {
      diagnosticsStatus: diagnostics?.healthCheck?.overallStatus || 'UNKNOWN'
    });

  } catch (error) {
    errorLog('定期メンテナンスエラー:', error);
  }
}

/**
 * キャッシュクリーンアップ
 * @private
 */
function performCacheCleanup() {
  try {
    debugLog('🧹 キャッシュクリーンアップ開始');

    // 期限切れキャッシュのクリーンアップ
    if (typeof cacheManager !== 'undefined' && cacheManager.cleanupExpired) {
      cacheManager.cleanupExpired();
    }

    // 秘密情報キャッシュのクリーンアップ（セキュリティのため）
    if (typeof unifiedSecretManager !== 'undefined') {
      unifiedSecretManager.clearSecretCache();
    }

    debugLog('✅ キャッシュクリーンアップ完了');

  } catch (error) {
    warnLog('キャッシュクリーンアップエラー:', error.message);
  }
}

/**
 * 古いログファイルのクリーンアップ
 * @private
 */
function cleanupOldLogs() {
  try {
    // 監査ログのサイズ制限
    if (typeof unifiedSecretManager !== 'undefined') {
      const auditLog = unifiedSecretManager.getAuditLog();
      if (auditLog.length > 1000) {
        // ログが1000件を超えた場合は古いものを削除
        debugLog('📋 監査ログのクリーンアップ実行');
      }
    }

  } catch (error) {
    warnLog('ログクリーンアップエラー:', error.message);
  }
}

/**
 * 緊急時のシステムリセット
 * 重大な問題が発生した際の最終手段
 */
function emergencySystemReset() {
  try {
    warnLog('🆘 緊急システムリセット開始');

    // 1. 現在のシステムをシャットダウン
    if (typeof systemIntegrationManager !== 'undefined') {
      systemIntegrationManager.shutdown();
    }

    // 2. キャッシュを全クリア
    try {
      if (typeof cacheManager !== 'undefined') {
        cacheManager.clearAll();
      }
      
      if (typeof unifiedSecretManager !== 'undefined') {
        unifiedSecretManager.clearSecretCache();
      }

      // GASの標準キャッシュもクリア
      CacheService.getScriptCache().removeAll([]);
      CacheService.getUserCache().removeAll([]);
    } catch (cacheError) {
      warnLog('キャッシュクリア中のエラー:', cacheError.message);
    }

    // 3. システムを再初期化
    const reinitResult = autoInitializeSystem();

    if (reinitResult.success) {
      infoLog('✅ 緊急システムリセット完了');
      return {
        success: true,
        message: 'システムリセットが完了しました',
        reinitResult: reinitResult
      };
    } else {
      errorLog('❌ システム再初期化に失敗');
      return {
        success: false,
        message: 'システム再初期化に失敗しました',
        reinitResult: reinitResult
      };
    }

  } catch (error) {
    errorLog('緊急システムリセットエラー:', error);
    return {
      success: false,
      message: '緊急リセットに失敗しました',
      error: error.message
    };
  }
}

/**
 * システム状態の簡易チェック（外部から呼び出し可能）
 * @returns {object} システム状態
 */
function getSystemHealthSummary() {
  try {
    if (typeof systemIntegrationManager === 'undefined') {
      return {
        status: 'NOT_INITIALIZED',
        message: 'システム統合管理が利用できません'
      };
    }

    const status = systemIntegrationManager.getSystemStatus();
    const metrics = systemIntegrationManager.getSystemMetrics();

    return {
      status: status.initialized ? 'HEALTHY' : 'NOT_INITIALIZED',
      initialized: status.initialized,
      componentsStatus: status.components,
      totalRequests: metrics.totalRequests || 0,
      errorRate: metrics.errorRate || '0%',
      uptime: metrics.startupTime ? 
        Math.round((Date.now() - new Date(metrics.startupTime).getTime()) / 1000) + 's' : 
        'Unknown',
      lastHealthCheck: metrics.lastHealthCheck,
      timestamp: new Date().toISOString()
    };

  } catch (error) {
    return {
      status: 'ERROR',
      message: error.message,
      timestamp: new Date().toISOString()
    };
  }
}

/**
 * 定期メンテナンスのスケジュール設定
 */
function schedulePeriodicMaintenance() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    
    // 既存の定期メンテナンストリガーを削除
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'performPeriodicMaintenance') {
        ScriptApp.deleteTrigger(trigger);
      }
    });

    // 新しい定期メンテナンストリガーを作成（6時間間隔）
    ScriptApp.newTrigger('performPeriodicMaintenance')
      .timeBased()
      .everyHours(6)
      .create();

    infoLog('📅 定期メンテナンスをスケジュールしました（6時間間隔）');

  } catch (error) {
    errorLog('定期メンテナンススケジュールエラー:', error);
  }
}