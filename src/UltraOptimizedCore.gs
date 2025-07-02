/**
 * @fileoverview 超最適化コアシステム - 2024年最新技術の結集
 * V8ランタイム、最新パフォーマンス技術、安定性強化を統合
 */

// 安定性を重視してvarを使用
var ULTRA_CONFIG = {
  EXECUTION_LIMITS: {
    MAX_TIME: 330000, // 5.5分（安全マージン）
    BATCH_SIZE: 100,
    API_RATE_LIMIT: 90 // 100秒間隔での制限
  },
  
  CACHE_STRATEGY: {
    L1_TTL: 300,     // Level 1: 5分
    L2_TTL: 3600,    // Level 2: 1時間  
    L3_TTL: 21600    // Level 3: 6時間（最大）
  }
};

/**
 * 簡素化されたエラーハンドリング関数群
 */

// ログ出力の最適化
function logOptimized(level, message, details) {
  try {
    if (typeof globalProfiler !== 'undefined') {
      globalProfiler.start('logging');
    }
    
    switch (level) {
      case 'error':
        console.error(message, details || '');
        break;
      case 'warn':
        console.warn(message, details || '');
        break;
      default:
        console.log(message, details || '');
    }
    
    if (typeof globalProfiler !== 'undefined') {
      globalProfiler.end('logging');
    }
  } catch (e) {
    // ログ出力自体が失敗した場合は無視
  }
}

function performSimpleCleanup() {
  try {
    // 期限切れPropertiesServiceエントリの削除
    var props = PropertiesService.getScriptProperties();
    var allProps = props.getProperties();
    var now = Date.now();
    
    Object.keys(allProps).forEach(function(key) {
      if (key.startsWith('CACHE_')) {
        try {
          var data = JSON.parse(allProps[key]);
          if (data.expiresAt && data.expiresAt < now) {
            props.deleteProperty(key);
          }
        } catch (e) {
          // 無効なデータは削除
          props.deleteProperty(key);
        }
      }
    });
  } catch (e) {
    console.warn('Cleanup failed:', e.message);
  }
}

// PerformanceOptimizer.gsでglobalProfilerが既に定義されているため、
// フォールバックの宣言は不要

/**
 * 公開エントリーポイント（安定版）
 * 既存の動作するdoGet関数をベースに最小限の最適化を適用
 */
function doGet(e) {
  try {
    var userId = e.parameter.userId;
    var mode = e.parameter.mode;
    var setup = e.parameter.setup;
    
    // セットアップページの表示
    if (setup === 'true') {
      return HtmlService.createTemplateFromFile('SetupPage')
        .evaluate()
        .setTitle('StudyQuest - サービスアカウント セットアップ');
    }
    
    if (!userId) {
      return HtmlService.createTemplateFromFile('Registration')
        .evaluate()
        .setTitle('新規登録');
    }

    var userInfo = findUserByIdOptimized(userId);
    if (!userInfo) {
      return HtmlService.createHtmlOutput('無効なユーザーIDです。');
    }
    
    // ユーザーの最終アクセス日時を更新（非同期）
    try {
      updateUserOptimized(userId, { lastAccessedAt: new Date().toISOString() });
    } catch (e) {
      console.error('最終アクセス日時の更新に失敗: ' + e.message);
    }

    // ユーザー情報をプロパティに保存（リアクション機能で使用）
    PropertiesService.getUserProperties().setProperty('CURRENT_USER_ID', userId);

    if (mode === 'admin') {
      var template = HtmlService.createTemplateFromFile('AdminPanel');
      template.userInfo = userInfo;
      template.userId = userId;
      return template.evaluate().setTitle('管理パネル - みんなの回答ボード');
    } else {
      var template = HtmlService.createTemplateFromFile('Page');
      template.userInfo = userInfo;
      template.userId = userId;
      return template.evaluate().setTitle('みんなの回答ボード');
    }
    
  } catch (error) {
    console.error('doGet error:', error.message);
    console.error('Stack trace:', error.stack);
    
    // エラー詳細をユーザーに表示（デバッグ用）
    return HtmlService.createHtmlOutput(
      '<h2>システムエラーが発生しました</h2>' +
      '<p>エラー詳細: ' + error.message + '</p>' +
      '<p>しばらく待ってから再度お試しください。</p>'
    );
  }
}

/**
 * パフォーマンス監視エンドポイント（簡易版）
 */
function getPerformanceMetrics() {
  try {
    var metrics = {};
    
    // プロファイラー情報
    if (typeof globalProfiler !== 'undefined') {
      metrics.profiler = globalProfiler.getReport();
    }
    
    // キャッシュ健康状態
    if (typeof AdvancedCacheManager !== 'undefined') {
      metrics.cache = AdvancedCacheManager.getHealth();
    }
    
    // 基本システム情報
    metrics.system = {
      timestamp: Date.now(),
      status: 'running'
    };
    
    return metrics;
  } catch (e) {
    return {
      error: e.message,
      timestamp: Date.now()
    };
  }
}