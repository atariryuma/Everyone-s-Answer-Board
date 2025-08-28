/**
 * Validation関数統合移行プラン
 * 既存関数から統一システムへの移行を管理
 */

/**
 * 移行対象関数のマッピング
 * 重複・類似関数を統一システムで統合
 */
const VALIDATION_MIGRATION_MAP = {
  // 基本検証系の統合
  validateWorkflowGAS: {
    newFunction: 'validateWorkflow',
    category: 'workflow',
    level: 'basic',
    deprecated: false, // 既存関数は保持（内部で統一システムを利用）
    notes: 'ワークフロー検証の基本機能'
  },
  
  validateWorkflowWithSheet: {
    newFunction: 'validateWorkflow',
    category: 'workflow', 
    level: 'standard',
    deprecated: false,
    notes: 'スプレッドシート統合ワークフロー検証'
  },

  comprehensiveWorkflowValidation: {
    newFunction: 'validateWorkflow',
    category: 'workflow',
    level: 'comprehensive', 
    deprecated: false,
    notes: '包括的ワークフロー検証'
  },

  // Config検証系の統合
  validateConfigJson: {
    newFunction: 'validateConfiguration',
    category: 'configuration',
    level: 'basic',
    deprecated: false,
    notes: 'ConfigJSON構造検証'
  },

  parseAndValidateConfigJson: {
    newFunction: 'validateConfiguration',
    category: 'configuration', 
    level: 'standard',
    deprecated: true, // 重複のため非推奨
    replacedBy: 'validateConfigJson + validateConfigJsonState',
    notes: 'parseとvalidateが混在していたため分離'
  },

  validateConfigJsonState: {
    newFunction: 'validateConfiguration',
    category: 'configuration',
    level: 'standard',
    deprecated: false,
    notes: 'Config状態と整合性検証'
  },

  // Header検証系の統合
  validateRequiredHeaders: {
    newFunction: 'validateDataIntegrity', 
    category: 'data_integrity',
    level: 'basic',
    deprecated: false,
    notes: '必須ヘッダー存在確認'
  },

  validateHeaderIntegrity: {
    newFunction: 'validateDataIntegrity',
    category: 'data_integrity', 
    level: 'standard',
    deprecated: false,
    notes: 'ヘッダー整合性検証'
  },

  // 認証系検証の統合
  validateUserAuthentication: {
    newFunction: 'validateAuthentication',
    category: 'authentication',
    level: 'basic',
    deprecated: false,
    notes: '基本ユーザー認証検証'
  },

  verifyUserAuthentication: {
    newFunction: 'validateAuthentication',
    category: 'authentication',
    level: 'standard', 
    deprecated: true, // 重複
    replacedBy: 'validateUserAuthentication',
    notes: '機能重複のため統合'
  },

  verifyUserAccess: {
    newFunction: 'validateAuthentication',
    category: 'authentication',
    level: 'standard',
    deprecated: false,
    notes: 'ユーザーアクセス権検証'
  },

  // システムチェック系の統合
  checkSystemConfiguration: {
    newFunction: 'validateConfiguration',
    category: 'configuration',
    level: 'standard', 
    deprecated: false,
    notes: 'システム設定検証'
  },

  validateSystemDependencies: {
    newFunction: 'validateConfiguration',
    category: 'configuration',
    level: 'comprehensive',
    deprecated: false,
    notes: 'システム依存関係検証'
  },

  checkAndHandleAutoStop: {
    newFunction: 'validateWorkflow',
    category: 'workflow',
    level: 'standard',
    deprecated: false,
    notes: '自動停止条件チェック（副作用あり）'
  },

  checkCurrentPublicationStatus: {
    newFunction: 'validateWorkflow', 
    category: 'workflow',
    level: 'basic',
    deprecated: false,
    notes: '公開状態確認'
  },

  checkApplicationAccess: {
    newFunction: 'validateAuthentication',
    category: 'authentication',
    level: 'standard',
    deprecated: false,
    notes: 'アプリケーションアクセス権確認'
  },

  checkAdmin: {
    newFunction: 'validateAuthentication',
    category: 'authentication', 
    level: 'basic',
    deprecated: false,
    notes: '管理者権限確認'
  },

  checkIfNewOrUpdatedForm: {
    newFunction: 'validateWorkflow',
    category: 'workflow',
    level: 'standard',
    deprecated: false,
    notes: 'フォーム更新状態確認'
  },

  // データ整合性検証系の統合
  performDataIntegrityCheck: {
    newFunction: 'validateDataIntegrity',
    category: 'data_integrity',
    level: 'comprehensive',
    deprecated: false,
    notes: '包括的データ整合性検証'
  },

  checkForDuplicates: {
    newFunction: 'validateDataIntegrity',
    category: 'data_integrity',
    level: 'standard', 
    deprecated: false,
    notes: '重複データ検出'
  },

  checkMissingRequiredFields: {
    newFunction: 'validateDataIntegrity',
    category: 'data_integrity',
    level: 'standard',
    deprecated: false,
    notes: '必須フィールド欠損チェック'
  },

  checkInvalidDataFormats: {
    newFunction: 'validateDataIntegrity',
    category: 'data_integrity',
    level: 'standard',
    deprecated: false, 
    notes: 'データ形式検証'
  },

  checkOrphanedData: {
    newFunction: 'validateDataIntegrity',
    category: 'data_integrity',
    level: 'comprehensive',
    deprecated: false,
    notes: '孤立データ検出'
  },

  validateUserScopedKey: {
    newFunction: 'validateDataIntegrity',
    category: 'data_integrity',
    level: 'basic',
    deprecated: false,
    notes: 'ユーザースコープキー検証'
  },

  // ヘルスチェック・監視系の統合
  performHealthCheck: {
    newFunction: 'validateSystemHealth',
    category: 'system_health', 
    level: 'basic',
    deprecated: false,
    notes: '基本ヘルスチェック'
  },

  performPerformanceCheck: {
    newFunction: 'validateSystemHealth',
    category: 'system_health',
    level: 'standard',
    deprecated: false,
    notes: 'パフォーマンス監視'
  },

  performSecurityCheck: {
    newFunction: 'validateSystemHealth',
    category: 'system_health',
    level: 'comprehensive',
    deprecated: false, 
    notes: 'セキュリティ状態検証'
  },

  checkAuthenticationHealth: {
    newFunction: 'validateAuthentication',
    category: 'authentication',
    level: 'comprehensive',
    deprecated: false,
    notes: '認証システムヘルス確認'
  },

  checkDatabaseHealth: {
    newFunction: 'validateSystemHealth', 
    category: 'system_health',
    level: 'standard',
    deprecated: false,
    notes: 'データベース接続・状態確認'
  },

  performGlobalMemoryCleanup: {
    newFunction: 'validateSystemHealth',
    category: 'system_health',
    level: 'comprehensive',
    deprecated: false,
    notes: 'メモリクリーンアップ（副作用あり）'
  }
};

/**
 * 非推奨関数の警告メッセージ生成
 */
function getDeprecationWarning(functionName) {
  const migration = VALIDATION_MIGRATION_MAP[functionName];
  
  if (!migration) {
    return null;
  }

  if (!migration.deprecated) {
    return null;
  }

  return {
    message: `⚠️  関数 '${functionName}' は非推奨です`,
    replacement: migration.replacedBy || migration.newFunction,
    category: migration.category,
    level: migration.level,
    notes: migration.notes
  };
}

/**
 * 移行統計の生成
 */
function getValidationMigrationStats() {
  const stats = {
    total: Object.keys(VALIDATION_MIGRATION_MAP).length,
    deprecated: 0,
    byCategory: {},
    byLevel: {}
  };

  for (const [funcName, migration] of Object.entries(VALIDATION_MIGRATION_MAP)) {
    if (migration.deprecated) {
      stats.deprecated++;
    }

    // カテゴリ別集計
    if (!stats.byCategory[migration.category]) {
      stats.byCategory[migration.category] = { total: 0, deprecated: 0 };
    }
    stats.byCategory[migration.category].total++;
    if (migration.deprecated) {
      stats.byCategory[migration.category].deprecated++;
    }

    // レベル別集計
    if (!stats.byLevel[migration.level]) {
      stats.byLevel[migration.level] = { total: 0, deprecated: 0 };
    }
    stats.byLevel[migration.level].total++;
    if (migration.deprecated) {
      stats.byLevel[migration.level].deprecated++;
    }
  }

  return stats;
}

/**
 * 統一システム移行推奨パターンの提供
 */
function getRecommendedMigrationPattern(oldFunctionName, ...args) {
  const migration = VALIDATION_MIGRATION_MAP[oldFunctionName];
  
  if (!migration) {
    return {
      error: `関数 '${oldFunctionName}' の移行情報が見つかりません`
    };
  }

  // 非推奨警告
  if (migration.deprecated) {
    ULog('WARN', 'ValidationMigration', `非推奨関数の使用`, {
      oldFunction: oldFunctionName,
      replacement: migration.replacedBy || migration.newFunction,
      notes: migration.notes
    });
  }

  // 推奨される新しい呼び出しパターン
  const recommendedCall = {
    method: 'UnifiedValidation.validate',
    category: migration.category,
    level: migration.level,
    example: `UnifiedValidation.validate('${migration.category}', '${migration.level}', options)`,
    wrappedFunction: migration.newFunction
  };

  // 引数の最適化提案
  const optimizedOptions = generateOptimizedOptions(migration.category, args);

  return {
    migration,
    recommendedCall,
    optimizedOptions,
    deprecationWarning: migration.deprecated ? getDeprecationWarning(oldFunctionName) : null
  };
}

/**
 * カテゴリ別の最適化されたオプション生成
 */
function generateOptimizedOptions(category, args) {
  const baseOptions = {};

  switch (category) {
    case 'authentication':
      if (args[0]) baseOptions.userId = args[0];
      break;
      
    case 'configuration':
      if (args[0]) baseOptions.config = args[0];
      if (args[1]) baseOptions.userInfo = args[1];
      break;
      
    case 'data_integrity':
      if (args[0]) baseOptions.userId = args[0];
      if (args[1]) baseOptions.headers = args[1]; 
      if (args[2]) baseOptions.userRows = args[2];
      if (args[3]) baseOptions.indices = args[3];
      break;
      
    case 'system_health':
      if (args[0]) baseOptions.userId = args[0];
      break;
      
    case 'workflow':
      if (args[0]) baseOptions.userId = args[0];
      if (args[1]) baseOptions.data = args[1];
      if (args[2]) baseOptions.config = args[2];
      if (args[3]) baseOptions.userInfo = args[3];
      break;
  }

  return baseOptions;
}

/**
 * レガシー関数ラッパー生成器
 * 既存関数のAPIを維持しながら内部で統一システムを使用
 */
function createLegacyWrapper(oldFunctionName, category, level, argumentMapper) {
  return function(...args) {
    // 非推奨警告（必要に応じて）
    const migration = VALIDATION_MIGRATION_MAP[oldFunctionName];
    if (migration && migration.deprecated) {
      ULog('WARN', 'LegacyWrapper', `非推奨関数の使用: ${oldFunctionName}`, {
        replacement: migration.replacedBy || migration.newFunction
      });
    }

    // 引数の最適化
    const options = argumentMapper ? argumentMapper(...args) : generateOptimizedOptions(category, args);

    // 統一システムを使用して検証実行
    const result = UnifiedValidation.validate(category, level, options);

    // レガシー形式に変換して返す
    return convertToLegacyFormat(oldFunctionName, result);
  };
}

/**
 * 統一システム結果をレガシー形式に変換
 */
function convertToLegacyFormat(oldFunctionName, unifiedResult) {
  // 関数ごとの個別変換ロジック
  switch (oldFunctionName) {
    case 'validateConfigJson':
      return {
        isValid: unifiedResult.success,
        errors: unifiedResult.success ? [] : extractErrorMessages(unifiedResult)
      };
      
    case 'validateUserAuthentication':
      return {
        hasAccess: unifiedResult.success,
        message: unifiedResult.summary ? `${unifiedResult.summary.passed}/${unifiedResult.summary.total} 通過` : 'チェック完了'
      };
      
    case 'performHealthCheck':
      return {
        overallStatus: unifiedResult.success ? 'HEALTHY' : (unifiedResult.hasCritical ? 'CRITICAL' : 'WARNING'),
        summary: unifiedResult.summary ? `通過率: ${unifiedResult.summary.passRate}%` : 'ヘルスチェック完了',
        details: unifiedResult.tests
      };
      
    default:
      // 汎用形式
      return {
        success: unifiedResult.success,
        valid: unifiedResult.success,
        isValid: unifiedResult.success,
        message: unifiedResult.summary ? `検証完了: ${unifiedResult.summary.passed}/${unifiedResult.summary.total} 通過` : '検証完了',
        result: unifiedResult
      };
  }
}

/**
 * 統一結果からエラーメッセージを抽出
 */
function extractErrorMessages(unifiedResult) {
  const errors = [];
  
  if (unifiedResult.tests) {
    for (const [testName, testResult] of Object.entries(unifiedResult.tests)) {
      if (!testResult.passed) {
        errors.push(`${testName}: ${testResult.message}`);
      }
    }
  }

  return errors.length > 0 ? errors : ['検証エラーが発生しました'];
}

/**
 * 移行レポート生成
 */
function generateMigrationReport() {
  const stats = getValidationMigrationStats();
  const timestamp = new Date().toISOString();
  
  return {
    timestamp,
    summary: {
      totalFunctions: stats.total,
      deprecatedFunctions: stats.deprecated,
      activeFunction: stats.total - stats.deprecated,
      migrationProgress: Math.round(((stats.total - stats.deprecated) / stats.total) * 100)
    },
    categoryBreakdown: stats.byCategory,
    levelBreakdown: stats.byLevel,
    deprecatedFunctions: Object.entries(VALIDATION_MIGRATION_MAP)
      .filter(([name, migration]) => migration.deprecated)
      .map(([name, migration]) => ({
        name,
        replacement: migration.replacedBy || migration.newFunction,
        notes: migration.notes
      })),
    recommendations: [
      '非推奨関数の使用を新しい統一システムに移行してください',
      'comprehensive レベルでの定期的な全体検証を実施してください',
      '検証結果のログ監視を設定してください'
    ]
  };
}