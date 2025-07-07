/**
 * @fileoverview パフォーマンス監視システム
 * 最適化効果の測定とレポート機能
 */

/**
 * パフォーマンス監視システム
 * 関数の実行時間、キャッシュヒット率、エラー率を監視
 */
var PerformanceMonitor = {
  
  /**
   * パフォーマンスメトリクスを取得
   */
  getMetrics: function() {
    try {
      var props = PropertiesService.getScriptProperties();
      var metricsData = props.getProperty('PERFORMANCE_METRICS');
      
      if (!metricsData) {
        return this.initializeMetrics();
      }
      
      return JSON.parse(metricsData);
    } catch (e) {
      console.error('メトリクス取得エラー: ' + e.message);
      return this.initializeMetrics();
    }
  },
  
  /**
   * メトリクスの初期化
   */
  initializeMetrics: function() {
    var initialMetrics = {
      functionCalls: {},
      cacheStats: {
        hits: 0,
        misses: 0,
        hitRate: 0
      },
      errorStats: {
        total: 0,
        byFunction: {}
      },
      systemStats: {
        averageResponseTime: 0,
        totalExecutionTime: 0,
        callCount: 0
      },
      lastUpdated: new Date().toISOString()
    };
    
    this.saveMetrics(initialMetrics);
    return initialMetrics;
  },
  
  /**
   * 関数実行の記録
   */
  recordFunctionCall: function(functionName, executionTime, success, cacheHit) {
    try {
      var metrics = this.getMetrics();
      
      // 関数呼び出し統計
      if (!metrics.functionCalls[functionName]) {
        metrics.functionCalls[functionName] = {
          totalCalls: 0,
          totalTime: 0,
          averageTime: 0,
          errors: 0,
          successRate: 100
        };
      }
      
      var funcStats = metrics.functionCalls[functionName];
      funcStats.totalCalls++;
      funcStats.totalTime += executionTime;
      funcStats.averageTime = funcStats.totalTime / funcStats.totalCalls;
      
      if (!success) {
        funcStats.errors++;
        metrics.errorStats.total++;
        if (!metrics.errorStats.byFunction[functionName]) {
          metrics.errorStats.byFunction[functionName] = 0;
        }
        metrics.errorStats.byFunction[functionName]++;
      }
      
      funcStats.successRate = ((funcStats.totalCalls - funcStats.errors) / funcStats.totalCalls) * 100;
      
      // キャッシュ統計
      if (cacheHit !== undefined) {
        if (cacheHit) {
          metrics.cacheStats.hits++;
        } else {
          metrics.cacheStats.misses++;
        }
        var totalCacheRequests = metrics.cacheStats.hits + metrics.cacheStats.misses;
        metrics.cacheStats.hitRate = (metrics.cacheStats.hits / totalCacheRequests) * 100;
      }
      
      // システム統計
      metrics.systemStats.callCount++;
      metrics.systemStats.totalExecutionTime += executionTime;
      metrics.systemStats.averageResponseTime = metrics.systemStats.totalExecutionTime / metrics.systemStats.callCount;
      metrics.lastUpdated = new Date().toISOString();
      
      this.saveMetrics(metrics);
      
    } catch (e) {
      console.error('メトリクス記録エラー: ' + e.message);
    }
  },
  
  /**
   * メトリクスの保存
   */
  saveMetrics: function(metrics) {
    try {
      var props = PropertiesService.getScriptProperties();
      props.setProperty('PERFORMANCE_METRICS', JSON.stringify(metrics));
    } catch (e) {
      console.error('メトリクス保存エラー: ' + e.message);
    }
  },
  
  /**
   * パフォーマンスレポートの生成
   */
  generateReport: function() {
    var metrics = this.getMetrics();
    var report = {
      summary: {
        totalCalls: metrics.systemStats.callCount,
        averageResponseTime: Math.round(metrics.systemStats.averageResponseTime * 100) / 100,
        cacheHitRate: Math.round(metrics.cacheStats.hitRate * 100) / 100,
        overallErrorRate: Math.round((metrics.errorStats.total / metrics.systemStats.callCount) * 100 * 100) / 100,
        lastUpdated: metrics.lastUpdated
      },
      topFunctions: this.getTopFunctions(metrics.functionCalls),
      slowestFunctions: this.getSlowestFunctions(metrics.functionCalls),
      errorProneFunctions: this.getErrorProneFunctions(metrics.functionCalls),
      cacheEffectiveness: this.analyzeCacheEffectiveness(metrics.cacheStats),
      recommendations: this.generateRecommendations(metrics)
    };
    
    return report;
  },
  
  /**
   * 最も呼び出される関数トップ5
   */
  getTopFunctions: function(functionCalls) {
    return Object.keys(functionCalls)
      .map(function(name) {
        return {
          name: name,
          calls: functionCalls[name].totalCalls,
          averageTime: Math.round(functionCalls[name].averageTime * 100) / 100
        };
      })
      .sort(function(a, b) { return b.calls - a.calls; })
      .slice(0, 5);
  },
  
  /**
   * 最も遅い関数トップ5
   */
  getSlowestFunctions: function(functionCalls) {
    return Object.keys(functionCalls)
      .map(function(name) {
        return {
          name: name,
          averageTime: Math.round(functionCalls[name].averageTime * 100) / 100,
          calls: functionCalls[name].totalCalls
        };
      })
      .sort(function(a, b) { return b.averageTime - a.averageTime; })
      .slice(0, 5);
  },
  
  /**
   * エラー率の高い関数
   */
  getErrorProneFunctions: function(functionCalls) {
    return Object.keys(functionCalls)
      .map(function(name) {
        var stats = functionCalls[name];
        return {
          name: name,
          errorRate: Math.round((100 - stats.successRate) * 100) / 100,
          totalErrors: stats.errors,
          totalCalls: stats.totalCalls
        };
      })
      .filter(function(func) { return func.errorRate > 0; })
      .sort(function(a, b) { return b.errorRate - a.errorRate; })
      .slice(0, 5);
  },
  
  /**
   * キャッシュ効果の分析
   */
  analyzeCacheEffectiveness: function(cacheStats) {
    var effectiveness = 'unknown';
    var hitRate = cacheStats.hitRate;
    
    if (hitRate >= 80) {
      effectiveness = 'excellent';
    } else if (hitRate >= 60) {
      effectiveness = 'good';
    } else if (hitRate >= 40) {
      effectiveness = 'fair';
    } else if (hitRate >= 20) {
      effectiveness = 'poor';
    } else {
      effectiveness = 'very poor';
    }
    
    return {
      effectiveness: effectiveness,
      hitRate: Math.round(hitRate * 100) / 100,
      totalHits: cacheStats.hits,
      totalMisses: cacheStats.misses,
      recommendation: this.getCacheRecommendation(effectiveness)
    };
  },
  
  /**
   * キャッシュ改善の推奨事項
   */
  getCacheRecommendation: function(effectiveness) {
    var recommendations = {
      'excellent': 'キャッシュは非常に効果的です。現在の設定を維持してください。',
      'good': 'キャッシュは良好に機能しています。さらなる最適化の余地があります。',
      'fair': 'キャッシュ戦略の見直しを検討してください。TTL設定や対象データの確認が必要です。',
      'poor': 'キャッシュ効率が低いです。キャッシュ対象の選定とTTL設定を見直してください。',
      'very poor': 'キャッシュがほとんど機能していません。実装の見直しが必要です。'
    };
    
    return recommendations[effectiveness] || 'キャッシュ効果を確認できません。';
  },
  
  /**
   * 最適化の推奨事項を生成
   */
  generateRecommendations: function(metrics) {
    var recommendations = [];
    
    // 平均応答時間に基づく推奨
    if (metrics.systemStats.averageResponseTime > 5000) {
      recommendations.push('平均応答時間が5秒を超えています。重い処理のバッチ化や非同期処理を検討してください。');
    } else if (metrics.systemStats.averageResponseTime > 2000) {
      recommendations.push('応答時間が2秒を超えています。パフォーマンス最適化を検討してください。');
    }
    
    // エラー率に基づく推奨
    var errorRate = (metrics.errorStats.total / metrics.systemStats.callCount) * 100;
    if (errorRate > 10) {
      recommendations.push('エラー率が10%を超えています。エラーハンドリングの強化が必要です。');
    } else if (errorRate > 5) {
      recommendations.push('エラー率が5%を超えています。安定性の向上を検討してください。');
    }
    
    // キャッシュヒット率に基づく推奨
    if (metrics.cacheStats.hitRate < 40) {
      recommendations.push('キャッシュヒット率が低いです。キャッシュ戦略の見直しを推奨します。');
    }
    
    if (recommendations.length === 0) {
      recommendations.push('現在のパフォーマンスは良好です。継続的な監視を行ってください。');
    }
    
    return recommendations;
  },
  
  /**
   * メトリクスのリセット
   */
  resetMetrics: function() {
    var props = PropertiesService.getScriptProperties();
    props.deleteProperty('PERFORMANCE_METRICS');
    return this.initializeMetrics();
  }
};

/**
 * パフォーマンス測定用のラッパー関数
 */
function measurePerformance(functionName, func, context) {
  var startTime = Date.now();
  var success = true;
  var result;
  
  try {
    if (context) {
      result = func.call(context);
    } else {
      result = func();
    }
  } catch (e) {
    success = false;
    result = e;
  }
  
  var executionTime = Date.now() - startTime;
  PerformanceMonitor.recordFunctionCall(functionName, executionTime, success);
  
  if (!success) {
    throw result;
  }
  
  return result;
}

/**
 * キャッシュヒット情報付きパフォーマンス測定
 */
function measurePerformanceWithCache(functionName, func, context, cacheHit) {
  var startTime = Date.now();
  var success = true;
  var result;
  
  try {
    if (context) {
      result = func.call(context);
    } else {
      result = func();
    }
  } catch (e) {
    success = false;
    result = e;
  }
  
  var executionTime = Date.now() - startTime;
  PerformanceMonitor.recordFunctionCall(functionName, executionTime, success, cacheHit);
  
  if (!success) {
    throw result;
  }
  
  return result;
}

/**
 * パフォーマンスレポートの取得（UI向け）
 */
function getPerformanceReport() {
  try {
    return PerformanceMonitor.generateReport();
  } catch (e) {
    console.error('パフォーマンスレポート取得エラー: ' + e.message);
    return {
      error: e.message,
      summary: {
        totalCalls: 0,
        averageResponseTime: 0,
        cacheHitRate: 0,
        overallErrorRate: 0,
        lastUpdated: new Date().toISOString()
      }
    };
  }
}

/**
 * パフォーマンスメトリクスのリセット（管理者向け）
 */
function resetPerformanceMetrics() {
  try {
    // 管理者確認
    if (!checkAdmin()) {
      throw new Error('管理者権限が必要です');
    }
    
    PerformanceMonitor.resetMetrics();
    return { status: 'success', message: 'パフォーマンスメトリクスをリセットしました' };
  } catch (e) {
    console.error('メトリクスリセットエラー: ' + e.message);
    return { status: 'error', message: e.message };
  }
}