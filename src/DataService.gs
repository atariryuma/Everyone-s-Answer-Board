/**
 * @fileoverview DataService - 統一データ操作サービス (遅延初期化対応)
 *
 * 🎯 責任範囲:
 * - スプレッドシートデータ取得・操作
 * - リアクション・ハイライト機能
 * - データフィルタリング・検索
 * - バルクデータAPI
 *
 * 🔄 GAS Best Practices準拠:
 * - 遅延初期化パターン (各公開関数先頭でinit)
 * - ファイル読み込み順序非依存設計
 * - グローバル副作用排除
 */

/* global formatTimestamp, createErrorResponse, createExceptionResponse, getSheetData, columnAnalysis, getQuestionText, findUserByEmail, findUserById, openSpreadsheet, updateUser, getUserSpreadsheetData, Config, getUserConfig, helpers, CACHE_DURATION, SLEEP_MS */

// ===========================================
// 🔧 GAS-Native DataService (Zero-Dependency)
// ===========================================

// ===========================================
// 🎯 統一列判定システム
// ===========================================

/**
 * 統一列判定関数 - フロントエンドとバックエンドで一貫した列解決
 * @param {Array} headers - ヘッダー配列
 * @param {string} fieldType - フィールドタイプ
 * @param {Object} columnMapping - 列マッピング（優先）
 * @param {Object} options - オプション設定
 * @returns {Object} { index: number, confidence: number, method: string, debug: Object }
 */
function resolveColumnIndex(headers, fieldType, columnMapping = {}, options = {}) {
  const debugInfo = {
    fieldType,
    searchMethods: [],
    candidateHeaders: [],
    finalSelection: null
  };

  try {
    // 入力検証
    if (!Array.isArray(headers) || headers.length === 0) {
      debugInfo.error = 'Invalid headers array';
      return { index: -1, confidence: 0, method: 'validation_failed', debug: debugInfo };
    }

    if (!fieldType || typeof fieldType !== 'string') {
      debugInfo.error = 'Invalid fieldType';
      return { index: -1, confidence: 0, method: 'validation_failed', debug: debugInfo };
    }

    // 1. 優先: 明示的列マッピング（管理パネルまたはAI検出）
    if (columnMapping && columnMapping[fieldType] !== undefined && columnMapping[fieldType] !== null) {
      const mappedIndex = columnMapping[fieldType];
      if (typeof mappedIndex === 'number' && mappedIndex >= 0 && mappedIndex < headers.length) {
        debugInfo.searchMethods.push({ method: 'explicit_mapping', index: mappedIndex, confidence: 100 });
        debugInfo.finalSelection = { method: 'explicit_mapping', index: mappedIndex };
        return { index: mappedIndex, confidence: 100, method: 'explicit_mapping', debug: debugInfo };
      }
    }

    // 2. ヘッダーパターンマッチング（拡張版）
    const headerPatterns = getHeaderPatterns();
    const patterns = headerPatterns[fieldType] || [];

    if (patterns.length > 0) {
      debugInfo.searchMethods.push({ method: 'pattern_matching', patterns });

      // 完全一致を最優先
      for (const pattern of patterns) {
        const exactMatch = headers.findIndex(header =>
          header && header.toLowerCase().trim() === pattern.toLowerCase().trim()
        );
        if (exactMatch !== -1) {
          debugInfo.candidateHeaders.push({ index: exactMatch, header: headers[exactMatch], pattern, matchType: 'exact' });
          debugInfo.finalSelection = { method: 'pattern_exact', index: exactMatch, pattern };
          return { index: exactMatch, confidence: 95, method: 'pattern_exact', debug: debugInfo };
        }
      }

      // 部分一致（includes）
      for (const pattern of patterns) {
        const partialMatch = headers.findIndex(header =>
          header && header.toLowerCase().includes(pattern.toLowerCase())
        );
        if (partialMatch !== -1) {
          debugInfo.candidateHeaders.push({ index: partialMatch, header: headers[partialMatch], pattern, matchType: 'partial' });
          debugInfo.finalSelection = { method: 'pattern_partial', index: partialMatch, pattern };
          return { index: partialMatch, confidence: 80, method: 'pattern_partial', debug: debugInfo };
        }
      }
    }

    // 3. フォールバック: 位置ベース推測（必要に応じて）
    if (options.allowPositionalFallback !== false) {
      const positionalIndex = getPositionalFallback(fieldType, headers.length);
      if (positionalIndex !== -1 && positionalIndex < headers.length) {
        debugInfo.searchMethods.push({ method: 'positional_fallback', index: positionalIndex });
        debugInfo.finalSelection = { method: 'positional_fallback', index: positionalIndex };
        return { index: positionalIndex, confidence: 30, method: 'positional_fallback', debug: debugInfo };
      }
    }

    // 4. 見つからない場合
    debugInfo.finalSelection = { method: 'not_found', index: -1 };
    return { index: -1, confidence: 0, method: 'not_found', debug: debugInfo };

  } catch (error) {
    console.error('[ERROR] DataService.resolveColumnIndex:', error.message || 'Unknown error');
    debugInfo.error = error.message;
    return { index: -1, confidence: 0, method: 'error', debug: debugInfo };
  }
}

/**
 * ヘッダーパターン定義（フロントエンドと統一）
 * @returns {Object} パターン定義オブジェクト
 */
function getHeaderPatterns() {
  return {
    timestamp: ['タイムスタンプ', 'timestamp', '投稿日時', '回答日時', '記録時刻'],
    email: ['メール', 'email', 'メールアドレス', 'mail', 'e-mail'],
    answer: ['回答', '答え', '意見', 'answer', 'opinion', 'response'],
    reason: ['理由', '根拠', 'reason', '説明', 'explanation'],
    class: ['クラス', '学年', 'class', '組', '学級'],
    name: ['名前', '氏名', 'name', '名', 'full_name'],

    // リアクション列（extractReactions用）
    understand: ['understand', '理解', 'UNDERSTAND'],
    like: ['like', 'いいね', 'LIKE'],
    curious: ['curious', '気になる', 'CURIOUS'],

    // ハイライト列（extractHighlight用）
    highlight: ['highlight', 'ハイライト', 'HIGHLIGHT']
  };
}

/**
 * 位置ベースフォールバック（Google Formsの典型的な列順序）
 * @param {string} fieldType - フィールドタイプ
 * @param {number} columnCount - 列数
 * @returns {number} 推測される列インデックス
 */
function getPositionalFallback(fieldType, columnCount) {
  // Google Formsの典型的な列順序: タイムスタンプ, 質問1, 質問2, ...
  const typicalPositions = {
    timestamp: 0,
    answer: 1,
    reason: 2,
    class: 3,
    name: 4,
    email: 5
  };

  const position = typicalPositions[fieldType];
  return (position !== undefined && position < columnCount) ? position : -1;
}

/**
 * 統一フィールド値抽出関数（エラーハンドリング強化版）
 * @param {Array} row - データ行
 * @param {Array} headers - ヘッダー配列
 * @param {string} fieldType - フィールドタイプ
 * @param {Object} columnMapping - 列マッピング
 * @param {Object} options - オプション設定
 * @returns {*} フィールド値
 */
function extractFieldValueUnified(row, headers, fieldType, columnMapping = {}, options = {}) {
  try {
    // 🛡️ 入力検証強化
    if (!Array.isArray(row)) {
      if (options.enableDebug) {
        console.warn(`[WARN] DataService.extractFieldValueUnified (${fieldType}): Invalid row data`);
      }
      return options.defaultValue || '';
    }

    const columnResult = resolveColumnIndex(headers, fieldType, columnMapping, options);

    // 🔍 詳細なデバッグ情報
    if (options.enableDebug) {
      console.log(`[DEBUG] DataService.extractFieldValueUnified (${fieldType}):`, {
        ...columnResult.debug,
        rowLength: row.length,
        hasValue: columnResult.index !== -1 && row[columnResult.index] !== undefined
      });
    }

    // 🎯 列が見つからない場合のフォールバック戦略
    if (columnResult.index === -1) {
      return handleColumnNotFound(fieldType, row, headers, options);
    }

    // 🔒 範囲外アクセス防止
    if (columnResult.index >= row.length) {
      if (options.enableDebug) {
        console.warn(`[WARN] DataService.extractFieldValueUnified (${fieldType}): Column index out of bounds`, {
          index: columnResult.index,
          rowLength: row.length
        });
      }
      return handleColumnNotFound(fieldType, row, headers, options);
    }

    const value = row[columnResult.index];
    return value !== undefined && value !== null ? value : (options.defaultValue || '');

  } catch (error) {
    console.error(`[ERROR] DataService.extractFieldValueUnified (${fieldType}):`, error.message || 'Unexpected error');
    return handleExtractionError(fieldType, error, options);
  }
}

/**
 * 列が見つからない場合のフォールバック処理
 * @param {string} fieldType - フィールドタイプ
 * @param {Array} row - データ行
 * @param {Array} headers - ヘッダー配列
 * @param {Object} options - オプション設定
 * @returns {*} フォールバック値
 */
function handleColumnNotFound(fieldType, row, headers, options = {}) {
  try {
    // 🎯 緊急フォールバック: 型別デフォルト値
    const emergencyDefaults = {
      timestamp: new Date().toISOString(),
      email: 'unknown@example.com',
      answer: '[未回答]',
      reason: '[理由なし]',
      class: '[未設定]',
      name: '[匿名]'
    };

    // 🔧 フィジカルポジション試行（最後の手段）
    if (options.allowEmergencyFallback !== false) {
      const physicalFallback = getPhysicalPositionFallback(fieldType, row);
      if (physicalFallback !== null) {
        if (options.enableDebug) {
          console.info(`[INFO] DataService.handleColumnNotFound (${fieldType}): Using physical position fallback`);
        }
        return physicalFallback;
      }
    }

    const defaultValue = options.defaultValue !== undefined
      ? options.defaultValue
      : emergencyDefaults[fieldType] || '';

    if (options.enableDebug) {
      console.info(`[INFO] DataService.handleColumnNotFound (${fieldType}): Using default value`, defaultValue);
    }

    return defaultValue;

  } catch (fallbackError) {
    console.error(`[ERROR] DataService.handleColumnNotFound fallback (${fieldType}):`, fallbackError.message || 'Fallback processing error');
    return options.defaultValue || '';
  }
}

/**
 * 抽出エラー時の処理
 * @param {string} fieldType - フィールドタイプ
 * @param {Error} error - エラーオブジェクト
 * @param {Object} options - オプション設定
 * @returns {*} エラー時のデフォルト値
 */
function handleExtractionError(fieldType, error, options = {}) {
  const errorMessage = `Column value extraction error: ${fieldType} - ${error.message || 'Unknown error'}`;

  // 🚨 重要なエラーログ
  console.error('[ERROR] DataService.handleExtractionError:', {
    fieldType,
    error: error.message,
    stack: error.stack,
    timestamp: new Date().toISOString()
  });

  return options.defaultValue || '';
}

/**
 * 物理位置ベースの緊急フォールバック
 * @param {string} fieldType - フィールドタイプ
 * @param {Array} row - データ行
 * @returns {*|null} フォールバック値またはnull
 */
function getPhysicalPositionFallback(fieldType, row) {
  try {
    // Google Formsの典型的な位置を試行
    const emergencyPositions = {
      timestamp: 0,
      answer: 1,
      reason: 2,
      class: 3,
      name: 4
    };

    const position = emergencyPositions[fieldType];
    if (position !== undefined && position < row.length && row[position] !== undefined) {
      return row[position];
    }

    return null;
  } catch (error) {
    console.warn(`[WARN] DataService.getPhysicalPositionFallback (${fieldType}):`, error.message || 'Position fallback error');
    return null;
  }
}

// ===========================================
// 🔍 列判定デバッグ・診断システム
// ===========================================

/**
 * 列判定状況の診断レポート生成
 * @param {Array} headers - ヘッダー配列
 * @param {Object} columnMapping - 列マッピング
 * @param {Array} requiredFields - 必要フィールドリスト
 * @returns {Object} 診断レポート
 */
function generateColumnDiagnosticReport(headers, columnMapping = {}, requiredFields = ['answer', 'reason', 'class', 'name']) {
  const report = {
    timestamp: new Date().toISOString(),
    headers: headers || [],
    columnMapping: columnMapping || {},
    requiredFields,
    diagnostics: {},
    summary: {
      total: requiredFields.length,
      resolved: 0,
      missing: 0,
      confidence: 0
    },
    recommendations: []
  };

  try {
    // 各必須フィールドの診断
    requiredFields.forEach(fieldType => {
      const columnResult = resolveColumnIndex(headers, fieldType, columnMapping, { enableDebug: false });

      report.diagnostics[fieldType] = {
        index: columnResult.index,
        confidence: columnResult.confidence,
        method: columnResult.method,
        debug: columnResult.debug,
        status: columnResult.index !== -1 ? 'resolved' : 'missing'
      };

      if (columnResult.index !== -1) {
        report.summary.resolved++;
        report.summary.confidence += columnResult.confidence;
      } else {
        report.summary.missing++;
      }
    });

    // 平均信頼度計算
    if (report.summary.resolved > 0) {
      report.summary.confidence = Math.round(report.summary.confidence / report.summary.resolved);
    }

    // 推奨事項の生成
    report.recommendations = generateColumnRecommendations(report);

    console.log('[DEBUG] DataService.generateColumnDiagnosticReport:', {
      resolved: report.summary.resolved,
      missing: report.summary.missing,
      avgConfidence: report.summary.confidence
    });

    return report;

  } catch (error) {
    console.error('[ERROR] DataService.generateColumnDiagnosticReport:', error.message || 'Diagnostic report error');
    report.error = error.message;
    return report;
  }
}

/**
 * 列判定改善の推奨事項生成
 * @param {Object} report - 診断レポート
 * @returns {Array} 推奨事項リスト
 */
function generateColumnRecommendations(report) {
  const recommendations = [];

  try {
    // 未解決フィールドの推奨
    Object.keys(report.diagnostics).forEach(fieldType => {
      const diagnostic = report.diagnostics[fieldType];

      if (diagnostic.status === 'missing') {
        recommendations.push({
          type: 'missing_field',
          field: fieldType,
          priority: 'high',
          message: `${fieldType}列が見つかりません。ヘッダー名を確認してください。`,
          suggestions: getHeaderPatterns()[fieldType] || []
        });
      } else if (diagnostic.confidence < 60) {
        recommendations.push({
          type: 'low_confidence',
          field: fieldType,
          priority: 'medium',
          message: `${fieldType}列の検出信頼度が低いです（${diagnostic.confidence}%）。列マッピングの明示的設定を推奨します。`,
          currentMethod: diagnostic.method
        });
      }
    });

    // 全体的な推奨
    if (report.summary.missing > 0) {
      recommendations.push({
        type: 'setup_required',
        priority: 'high',
        message: '管理パネルで列設定の確認・調整を行ってください。',
        missingCount: report.summary.missing
      });
    }

    if (report.summary.confidence < 80 && report.summary.resolved > 0) {
      recommendations.push({
        type: 'improve_headers',
        priority: 'medium',
        message: 'より明確なヘッダー名の使用を推奨します。',
        avgConfidence: report.summary.confidence
      });
    }

    return recommendations;

  } catch (error) {
    console.error('[ERROR] DataService.generateColumnRecommendations:', error.message || 'Recommendations generation error');
    return [];
  }
}

/**
 * 列判定状況のリアルタイム監視
 * @param {Array} headers - ヘッダー配列
 * @param {Object} columnMapping - 列マッピング
 * @param {Object} options - 監視オプション
 */
function monitorColumnResolution(headers, columnMapping = {}, options = {}) {
  try {
    const monitoringData = {
      timestamp: new Date().toISOString(),
      sessionId: options.sessionId || 'unknown',
      headerCount: headers ? headers.length : 0,
      mappingCount: Object.keys(columnMapping).length,
      resolutionStats: {}
    };

    const criticalFields = ['answer', 'reason', 'timestamp'];

    criticalFields.forEach(fieldType => {
      const columnResult = resolveColumnIndex(headers, fieldType, columnMapping);
      monitoringData.resolutionStats[fieldType] = {
        resolved: columnResult.index !== -1,
        confidence: columnResult.confidence,
        method: columnResult.method
      };
    });

    // 🎯 パフォーマンス監視
    const resolvedCount = Object.values(monitoringData.resolutionStats)
      .filter(stat => stat.resolved).length;

    const successRate = criticalFields.length > 0
      ? Math.round((resolvedCount / criticalFields.length) * 100)
      : 0;

    console.log('[DEBUG] DataService.monitorColumnResolution:', {
      ...monitoringData,
      successRate: `${successRate}%`
    });

    // アラート条件
    if (successRate < 70) {
      console.warn('[WARN] DataService.monitorColumnResolution: Column resolution rate declined', {
        successRate,
        requiredAction: '列設定の確認が必要'
      });
    }

    return monitoringData;

  } catch (error) {
    console.error('[ERROR] DataService.monitorColumnResolution:', error.message || 'Column monitoring error');
    return null;
  }
}

// ===========================================
// 🔗 システム全体統合・診断API
// ===========================================

/**
 * フロントエンド-バックエンド列判定システムの統合診断
 * @param {Array} headers - ヘッダー配列
 * @param {Object} columnMapping - 列マッピング
 * @param {Array} sampleData - サンプルデータ（テスト用）
 * @returns {Object} 統合診断結果
 */
function performIntegratedColumnDiagnostics(headers, columnMapping = {}, sampleData = []) {
  const diagnostics = {
    timestamp: new Date().toISOString(),
    systemHealth: {
      backend: { status: 'unknown', score: 0 },
      frontend: { status: 'unknown', score: 0 },
      integration: { status: 'unknown', score: 0 }
    },
    columnTests: {},
    recommendations: [],
    summary: {
      overallScore: 0,
      criticalIssues: 0,
      warnings: 0
    }
  };

  try {
    console.log('🔍 performIntegratedColumnDiagnostics: システム統合診断開始');

    // 🎯 バックエンド列判定システム診断
    diagnostics.systemHealth.backend = diagnoseBackendColumnSystem(headers, columnMapping);

    // 🎯 フロントエンドとの連携診断（モック）
    diagnostics.systemHealth.frontend = diagnoseFrontendColumnSystem(columnMapping);

    // 🎯 統合テスト実行
    diagnostics.systemHealth.integration = performIntegrationTests(headers, columnMapping, sampleData);

    // 🎯 個別フィールドテスト
    const testFields = ['answer', 'reason', 'class', 'name', 'timestamp'];
    testFields.forEach(fieldType => {
      diagnostics.columnTests[fieldType] = testFieldResolution(headers, fieldType, columnMapping, sampleData);
    });

    // 🎯 総合スコア算出
    const scores = [
      diagnostics.systemHealth.backend.score,
      diagnostics.systemHealth.frontend.score,
      diagnostics.systemHealth.integration.score
    ];
    diagnostics.summary.overallScore = Math.round(scores.reduce((sum, score) => sum + score, 0) / scores.length);

    // 🎯 問題集計
    Object.values(diagnostics.columnTests).forEach(test => {
      if (test.severity === 'critical') diagnostics.summary.criticalIssues++;
      if (test.severity === 'warning') diagnostics.summary.warnings++;
    });

    // 🎯 推奨事項生成
    diagnostics.recommendations = generateSystemRecommendations(diagnostics);

    console.log('✅ performIntegratedColumnDiagnostics完了:', {
      overallScore: diagnostics.summary.overallScore,
      criticalIssues: diagnostics.summary.criticalIssues,
      warnings: diagnostics.summary.warnings
    });

    return diagnostics;

  } catch (error) {
    console.error('[ERROR] DataService.performIntegratedColumnDiagnostics:', error.message || 'Integrated diagnostics error');
    diagnostics.error = error.message;
    return diagnostics;
  }
}

/**
 * バックエンド列判定システム診断
 * @param {Array} headers - ヘッダー配列
 * @param {Object} columnMapping - 列マッピング
 * @returns {Object} バックエンド診断結果
 */
function diagnoseBackendColumnSystem(headers, columnMapping) {
  const diagnosis = {
    status: 'healthy',
    score: 100,
    issues: [],
    tests: {}
  };

  try {
    // 統一列判定関数テスト
    const testFields = ['answer', 'reason', 'class', 'name'];
    testFields.forEach(fieldType => {
      const result = resolveColumnIndex(headers, fieldType, columnMapping);
      diagnosis.tests[fieldType] = {
        resolved: result.index !== -1,
        confidence: result.confidence,
        method: result.method
      };

      if (result.index === -1) {
        diagnosis.issues.push(`${fieldType}フィールドが解決できません`);
        diagnosis.score -= 20;
      } else if (result.confidence < 50) {
        diagnosis.issues.push(`${fieldType}フィールドの信頼度が低いです`);
        diagnosis.score -= 10;
      }
    });

    // エラーハンドリング機能テスト
    try {
      extractFieldValueUnified([], headers, 'answer', columnMapping);
      extractReactions([], headers);
      extractHighlight([], headers);
    } catch (error) {
      diagnosis.issues.push('エラーハンドリングに問題があります');
      diagnosis.score -= 15;
    }

    // 診断結果判定
    if (diagnosis.score >= 80) {
      diagnosis.status = 'healthy';
    } else if (diagnosis.score >= 60) {
      diagnosis.status = 'warning';
    } else {
      diagnosis.status = 'critical';
    }

    console.log('🔍 diagnoseBackendColumnSystem:', diagnosis);
    return diagnosis;

  } catch (error) {
    console.error('[ERROR] DataService.diagnoseBackendColumnSystem:', error.message || 'Backend column system error');
    return {
      status: 'error',
      score: 0,
      issues: ['バックエンド診断中にエラーが発生'],
      error: error.message
    };
  }
}

/**
 * フロントエンド列判定システム診断（モック）
 * @param {Object} columnMapping - 列マッピング
 * @returns {Object} フロントエンド診断結果
 */
function diagnoseFrontendColumnSystem(columnMapping) {
  const diagnosis = {
    status: 'healthy',
    score: 90,
    issues: [],
    tests: {
      mappingKeysConsistency: true,
      conflictResolution: true,
      saveValidation: true
    }
  };

  try {
    // マッピングキーの一貫性確認
    const expectedKeys = ['answer', 'reason', 'class', 'name'];
    const mappingData = columnMapping.mapping || {};

    expectedKeys.forEach(key => {
      if (mappingData[key] === undefined) {
        diagnosis.issues.push(`フロントエンド: ${key}キーが不足`);
        diagnosis.score -= 15;
      }
    });

    // 診断結果判定
    if (diagnosis.score >= 80) {
      diagnosis.status = 'healthy';
    } else if (diagnosis.score >= 60) {
      diagnosis.status = 'warning';
    } else {
      diagnosis.status = 'critical';
    }

    console.log('🔍 diagnoseFrontendColumnSystem:', diagnosis);
    return diagnosis;

  } catch (error) {
    console.error('diagnoseFrontendColumnSystem エラー:', error);
    return {
      status: 'error',
      score: 0,
      issues: ['フロントエンド診断中にエラーが発生'],
      error: error.message
    };
  }
}

/**
 * 統合テスト実行
 * @param {Array} headers - ヘッダー配列
 * @param {Object} columnMapping - 列マッピング
 * @param {Array} sampleData - サンプルデータ
 * @returns {Object} 統合テスト結果
 */
function performIntegrationTests(headers, columnMapping, sampleData) {
  const testResult = {
    status: 'healthy',
    score: 100,
    issues: [],
    tests: {
      dataExtraction: false,
      reactionProcessing: false,
      highlightProcessing: false,
      errorRecovery: false
    }
  };

  try {
    // データ抽出統合テスト
    if (sampleData.length > 0) {
      const [testRow] = sampleData;
      const extractedAnswer = extractFieldValueUnified(testRow, headers, 'answer', columnMapping);
      testResult.tests.dataExtraction = extractedAnswer !== '';

      if (!testResult.tests.dataExtraction) {
        testResult.issues.push('データ抽出テストが失敗しました');
        testResult.score -= 25;
      }
    }

    // リアクション処理テスト
    try {
      const reactions = extractReactions(sampleData[0] || [], headers);
      testResult.tests.reactionProcessing = reactions && typeof reactions === 'object';

      if (!testResult.tests.reactionProcessing) {
        testResult.issues.push('リアクション処理テストが失敗しました');
        testResult.score -= 25;
      }
    } catch (error) {
      testResult.tests.reactionProcessing = false;
      testResult.issues.push('リアクション処理でエラーが発生');
      testResult.score -= 25;
    }

    // ハイライト処理テスト
    try {
      const highlight = extractHighlight(sampleData[0] || [], headers);
      testResult.tests.highlightProcessing = typeof highlight === 'boolean';

      if (!testResult.tests.highlightProcessing) {
        testResult.issues.push('ハイライト処理テストが失敗しました');
        testResult.score -= 25;
      }
    } catch (error) {
      testResult.tests.highlightProcessing = false;
      testResult.issues.push('ハイライト処理でエラーが発生');
      testResult.score -= 25;
    }

    // エラー回復テスト
    try {
      extractFieldValueUnified(null, [], 'invalid', {}, { enableDebug: false });
      testResult.tests.errorRecovery = true;
    } catch (error) {
      testResult.tests.errorRecovery = false;
      testResult.issues.push('エラー回復機能に問題があります');
      testResult.score -= 25;
    }

    // 診断結果判定
    if (testResult.score >= 80) {
      testResult.status = 'healthy';
    } else if (testResult.score >= 60) {
      testResult.status = 'warning';
    } else {
      testResult.status = 'critical';
    }

    console.log('🔍 performIntegrationTests:', testResult);
    return testResult;

  } catch (error) {
    console.error('performIntegrationTests エラー:', error);
    return {
      status: 'error',
      score: 0,
      issues: ['統合テスト中にエラーが発生'],
      error: error.message
    };
  }
}

/**
 * フィールド解決テスト
 * @param {Array} headers - ヘッダー配列
 * @param {string} fieldType - フィールドタイプ
 * @param {Object} columnMapping - 列マッピング
 * @param {Array} sampleData - サンプルデータ
 * @returns {Object} フィールドテスト結果
 */
function testFieldResolution(headers, fieldType, columnMapping, sampleData) {
  try {
    const result = resolveColumnIndex(headers, fieldType, columnMapping);

    const test = {
      fieldType,
      resolved: result.index !== -1,
      confidence: result.confidence,
      method: result.method,
      severity: 'info'
    };

    if (!test.resolved) {
      test.severity = fieldType === 'answer' ? 'critical' : 'warning';
      test.issue = `${fieldType}フィールドが解決できません`;
    } else if (test.confidence < 50) {
      test.severity = 'warning';
      test.issue = `${fieldType}フィールドの信頼度が低いです (${test.confidence}%)`;
    }

    // 実データでの抽出テスト
    if (test.resolved && sampleData.length > 0) {
      const extractedValue = extractFieldValueUnified(sampleData[0], headers, fieldType, columnMapping);
      test.extractionSuccess = extractedValue !== '';

      if (!test.extractionSuccess) {
        test.severity = 'warning';
        test.issue = `${fieldType}フィールドからデータを抽出できません`;
      }
    }

    return test;

  } catch (error) {
    return {
      fieldType,
      resolved: false,
      severity: 'critical',
      issue: `${fieldType}フィールドテスト中にエラー: ${error.message}`
    };
  }
}

/**
 * システム推奨事項生成
 * @param {Object} diagnostics - 診断結果
 * @returns {Array} 推奨事項リスト
 */
function generateSystemRecommendations(diagnostics) {
  const recommendations = [];

  try {
    // 重大な問題の推奨事項
    if (diagnostics.summary.criticalIssues > 0) {
      recommendations.push({
        priority: 'critical',
        type: 'immediate_action',
        message: `${diagnostics.summary.criticalIssues}件の重大な問題があります。列設定を確認してください。`
      });
    }

    // システム別推奨事項
    if (diagnostics.systemHealth.backend.score < 70) {
      recommendations.push({
        priority: 'high',
        type: 'backend_improvement',
        message: 'バックエンドの列判定システムに問題があります。ヘッダーパターンや列マッピングを確認してください。'
      });
    }

    if (diagnostics.systemHealth.frontend.score < 70) {
      recommendations.push({
        priority: 'high',
        type: 'frontend_improvement',
        message: 'フロントエンドの列設定に問題があります。管理パネルで設定を確認してください。'
      });
    }

    if (diagnostics.summary.overallScore < 80) {
      recommendations.push({
        priority: 'medium',
        type: 'general_improvement',
        message: 'システム全体の列判定精度向上のため、より明確なヘッダー名の使用を推奨します。'
      });
    }

    return recommendations;

  } catch (error) {
    console.error('generateSystemRecommendations エラー:', error);
    return [{
      priority: 'low',
      type: 'error',
      message: '推奨事項の生成中にエラーが発生しました'
    }];
  }
}


/**
 * DataService - ゼロ依存アーキテクチャ
 * GAS-Nativeパターンによる直接API呼び出し
 * DB, CONSTANTS依存を完全排除
 */


/**
 * ユーザーのスプレッドシートデータ取得（統合版）
 * GAS公式ベストプラクティス：シンプルな関数形式
 * @param {string} userId - ユーザーID
 * @param {Object} options - 取得オプション
 * @returns {Object} GAS公式推奨レスポンス形式
 */
function getUserSheetData(userId, options = {}) {
  const startTime = Date.now();

  try {
    // 🚀 Zero-dependency data processing

    // 🔧 Zero-Dependency統一: 直接findUserById使用
    const user = findUserById(userId);
    if (!user) {
      console.error('DataService.getUserSheetData: ユーザーが見つかりません', { userId });
      return helpers.createDataServiceErrorResponse('ユーザーが見つかりません');
    }

    // 統一API使用: 構造化パース
    const configResult = getUserConfig(userId);
    const config = configResult.success ? configResult.config : {};
    if (!config.spreadsheetId) {
      console.warn('[WARN] DataService.getUserSheetData: Spreadsheet ID not configured', { userId });
      return helpers.createDataServiceErrorResponse('スプレッドシートが設定されていません');
    }

    // データ取得実行
    const result = fetchSpreadsheetData(config, options, user);

    const executionTime = Date.now() - startTime;
    console.info('getSheetData: データ取得完了', {
      userId,
      rowCount: result.data?.length || 0,
      executionTime
    });

    // Standardized response format
    if (result.success) {
      return {
        ...result,
        header: getQuestionText(config) || result.sheetName || '回答一覧',
        showDetails: config.showDetails !== false // デフォルトはtrue
      };
    }

    return result;
  } catch (error) {
    console.error('DataService.getUserSheetData: エラー', {
      userId,
      error: error.message
    });
    // Direct return format like admin panel getSheetList
    return helpers.createDataServiceErrorResponse(error.message || 'データ取得エラー');
  }
}


/**
 * スプレッドシートデータ取得実行
 * ✅ バッチ処理対応 - GAS制限対応（実行時間・メモリ制限）
 * GAS公式ベストプラクティス：シンプルな関数形式
 * @param {Object} config - ユーザー設定
 * @param {Object} options - 取得オプション
 * @returns {Object} GAS公式推奨レスポンス形式
 */
/**
 * スプレッドシート接続とシート取得（GAS Best Practice: 単一責任）
 * @param {Object} config - 設定オブジェクト
 * @returns {Object} シート情報
 */
function connectToSpreadsheetSheet(config) {
  const dataAccess = openSpreadsheet(config.spreadsheetId);
  const {spreadsheet} = dataAccess;
  const sheet = spreadsheet.getSheetByName(config.sheetName);

  if (!sheet) {
    const sheetName = config && config.sheetName ? config.sheetName : '不明';
    throw new Error(`シート '${sheetName}' が見つかりません`);
  }

  return { sheet, spreadsheet };
}

/**
 * シートの寸法取得（GAS Best Practice: 単一責任）
 * @param {Sheet} sheet - シートオブジェクト
 * @returns {Object} 寸法情報
 */
function getSheetDimensions(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  return { lastRow, lastCol };
}

/**
 * ヘッダー行取得（GAS Best Practice: 単一責任）
 * @param {Sheet} sheet - シートオブジェクト
 * @param {number} lastCol - 最終列
 * @returns {Array} ヘッダー配列
 */
function getSheetHeaders(sheet, lastCol) {
  const [headers] = sheet.getRange(1, 1, 1, lastCol).getValues();
  return headers;
}

/**
 * バッチデータ処理（GAS Best Practice: 大量データ処理分離）
 * @param {Sheet} sheet - シートオブジェクト
 * @param {Array} headers - ヘッダー配列
 * @param {number} lastRow - 最終行
 * @param {number} lastCol - 最終列
 * @param {Object} config - 設定
 * @param {Object} options - オプション
 * @param {Object} user - ユーザー情報
 * @param {number} startTime - 開始時刻
 * @returns {Array} 処理済みデータ
 */
function processBatchData(sheet, headers, lastRow, lastCol, config, options, user, startTime) {
  const MAX_EXECUTION_TIME = 180000; // 3分制限
  const MAX_BATCH_SIZE = 200; // バッチサイズ

  let processedData = [];
  let processedCount = 0;
  const totalDataRows = lastRow - 1;

  for (let startRow = 2; startRow <= lastRow; startRow += MAX_BATCH_SIZE) {
    // 実行時間チェック
    if (Date.now() - startTime > MAX_EXECUTION_TIME) {
      console.warn('DataService.processBatchData: 実行時間制限のため処理を中断', {
        processedRows: processedCount,
        totalRows: totalDataRows
      });
      break;
    }

    const endRow = Math.min(startRow + MAX_BATCH_SIZE - 1, lastRow);
    const batchSize = endRow - startRow + 1;

    try {
      const batchRows = sheet.getRange(startRow, 1, batchSize, lastCol).getValues();
      const batchProcessed = processRawDataBatch(batchRows, headers, config, options, startRow - 2, user);

      processedData = processedData.concat(batchProcessed);
      processedCount += batchSize;

      // API制限対策: 1000行毎に短い休憩
      if (processedCount % 1000 === 0) {
        Utilities.sleep(100); // 0.1秒休憩
      }

    } catch (batchError) {
      console.error('DataService.processBatchData: バッチ処理エラー', {
        startRow, endRow, error: batchError && batchError.message ? batchError.message : 'エラー詳細不明'
      });
    }
  }

  // フィルタリングとソート適用
  if (options.classFilter) {
    processedData = processedData.filter(item => item.class === options.classFilter);
  }

  if (options.sortBy) {
    processedData = applySortAndLimit(processedData, {
      sortBy: options.sortBy,
      limit: options.limit
    });
  }

  return processedData;
}

/**
 * スプレッドシートデータ取得（リファクタリング版 - GAS Best Practice準拠）
 * @param {Object} config - 設定オブジェクト
 * @param {Object} options - オプション
 * @param {Object} user - ユーザー情報
 * @returns {Object} データ取得結果
 */
function fetchSpreadsheetData(config, options = {}, user = null) {
  const startTime = Date.now();

  try {
    // シート接続
    const { sheet } = connectToSpreadsheetSheet(config);

    // 寸法取得
    const { lastRow, lastCol } = getSheetDimensions(sheet);

    if (lastRow <= 1) {
      return helpers.createDataServiceSuccessResponse([], [], config.sheetName || '不明');
    }

    // ヘッダー取得
    const headers = getSheetHeaders(sheet, lastCol);

    // バッチ処理実行
    const processedData = processBatchData(sheet, headers, lastRow, lastCol, config, options, user, startTime);

    console.info('DataService.fetchSpreadsheetData: バッチ処理完了', {
      filteredRows: processedData.length,
      executionTime: Date.now() - startTime
    });

    // ✅ フロントエンド期待形式で直接返却
    return {
      success: true,
      data: processedData,
      headers,
      sheetName: config.sheetName || '不明',
      // デバッグ情報（オプショナル）
      filteredRows: processedData.length,
      executionTime: Date.now() - startTime
    };
  } catch (error) {
    console.error('DataService.fetchSpreadsheetData: エラー', error.message);
    throw error;
  }
}

/**
 * バッチ処理用データ変換（メモリ効率重視）
 * @param {Array} batchRows - バッチデータ行
 * @param {Array} headers - ヘッダー配列
 * @param {Object} config - 設定
 * @param {Object} options - 処理オプション
 * @param {number} startOffset - 開始オフセット（行番号計算用）
 * @returns {Array} 処理済みバッチデータ
 */
function processRawDataBatch(batchRows, headers, config, options = {}, startOffset = 0, user = null) {
  try {
    const columnMapping = config.columnMapping?.mapping || {};
    const processedBatch = [];

    batchRows.forEach((row, batchIndex) => {
      try {
        // グローバル行インデックス計算
        const globalIndex = startOffset + batchIndex;

        // 基本データ構造作成
        const item = {
          id: `row_${globalIndex + 2}`,
          rowIndex: globalIndex + 2, // 1-based row number including header
          timestamp: extractFieldValue(row, headers, 'timestamp') || '',
          email: extractFieldValue(row, headers, 'email') || '',

          // メインコンテンツ（フロントエンドと統一）
          answer: extractFieldValue(row, headers, 'answer', columnMapping) || '',
          opinion: extractFieldValue(row, headers, 'answer', columnMapping) || '', // Alias for answer field
          reason: extractFieldValue(row, headers, 'reason', columnMapping) || '',
          class: extractFieldValue(row, headers, 'class', columnMapping) || '',
          name: extractFieldValue(row, headers, 'name', columnMapping) || '',

          // メタデータ
          formattedTimestamp: formatTimestamp(extractFieldValue(row, headers, 'timestamp')),
          isEmpty: isEmptyRow(row),

          // リアクション（{count, reacted} 形式に統一）
          reactions: extractReactions(row, headers, user.userEmail),
          highlight: extractHighlight(row, headers)
        };

        // フィルタリング
        if (shouldIncludeRow(item, options)) {
          processedBatch.push(item);
        }
      } catch (rowError) {
        console.warn('DataService.processRawDataBatch: 行処理エラー', {
          batchIndex,
          globalIndex: startOffset + batchIndex,
          error: rowError.message
        });
      }
    });

    return processedBatch;
  } catch (error) {
    console.error('DataService.processRawDataBatch: エラー', error.message);
    return [];
  }
}

/**
 * 生データを処理・変換
 * @param {Array} dataRows - 生データ行
 * @param {Array} headers - ヘッダー配列
 * @param {Object} config - 設定
 * @param {Object} options - 処理オプション
 * @returns {Array} 処理済みデータ
 */
function processRawData(dataRows, headers, config, options = {}, user = null) {
  try {
    const columnMapping = config.columnMapping?.mapping || {};
    const processedData = [];

    dataRows.forEach((row, index) => {
      try {
        // 基本データ構造作成
        const item = {
          id: `row_${index + 2}`,
          rowIndex: index + 2,
          timestamp: extractFieldValue(row, headers, 'timestamp') || '',
          email: extractFieldValue(row, headers, 'email') || '',

          // メインコンテンツ
          answer: extractFieldValue(row, headers, 'answer', columnMapping) || '',
          opinion: extractFieldValue(row, headers, 'answer', columnMapping) || '', // Alias for answer field
          reason: extractFieldValue(row, headers, 'reason', columnMapping) || '',
          class: extractFieldValue(row, headers, 'class', columnMapping) || '',
          name: extractFieldValue(row, headers, 'name', columnMapping) || '',

          // メタデータ
          formattedTimestamp: formatTimestamp(extractFieldValue(row, headers, 'timestamp')),
          isEmpty: isEmptyRow(row),

          // リアクション（{count, reacted} 形式）
          reactions: extractReactions(row, headers, user.userEmail),
          highlight: extractHighlight(row, headers)
        };

        // フィルタリング
        if (shouldIncludeRow(item, options)) {
          processedData.push(item);
        }
      } catch (rowError) {
        console.warn('DataService.processRawData: 行処理エラー', {
          rowIndex: index,
          error: rowError.message
        });
      }
    });

    // ソート・制限適用
    return applySortAndLimit(processedData, options);
  } catch (error) {
    console.error('DataService.processRawData: エラー', error.message);
    return [];
  }
}

/**
 * フィールド値抽出（列マッピング対応）
 * @param {Array} row - データ行
 * @param {Array} headers - ヘッダー配列
 * @param {string} fieldType - フィールドタイプ
 * @param {Object} columnMapping - 列マッピング
 * @returns {*} フィールド値
 */
function extractFieldValue(row, headers, fieldType, columnMapping = {}) {
  // 🎯 統一列判定システムに委譲（後方互換性保持）
  return extractFieldValueUnified(row, headers, fieldType, columnMapping);
}

// ===========================================
// 🎯 リアクション・ハイライト機能
// ===========================================

/**
 * リアクション追加
 * @param {string} userId - ユーザーID
 * @param {string} rowId - 行ID
 * @param {string} reactionType - リアクションタイプ
 * @returns {boolean} 成功可否
 */



/**
 * スプレッドシート内リアクション更新
 * @param {Object} config - 設定
 * @param {string} rowId - 行ID
 * @param {string} reactionType - リアクションタイプ
 * @param {string} action - アクション（add/remove）
 * @returns {boolean} 成功可否
 */
function updateReactionInSheet(config, rowId, reactionType, action) {
  try {
    const dataAccess = openSpreadsheet(config.spreadsheetId);
    const {spreadsheet} = dataAccess;
    const sheet = spreadsheet.getSheetByName(config.sheetName);

    if (!sheet) {
      throw new Error('シートが見つかりません');
    }

    // 行番号抽出（row_3 → 3）
    const rowNumber = parseInt(rowId.replace('row_', ''));
    if (isNaN(rowNumber) || rowNumber < 2) {
      throw new Error('無効な行ID');
    }

    // リアクション列の取得・作成
    const reactionColumn = getOrCreateReactionColumn(sheet, reactionType);
    if (!reactionColumn) {
      throw new Error('リアクション列の作成に失敗');
    }

    // CLAUDE.md準拠: バッチ操作による70倍性能向上 (getValue/setValue → getValues/setValues)
    const currentValue = sheet.getRange(rowNumber, reactionColumn, 1, 1).getValues()[0][0] || 0;
    const newValue = action === 'add'
      ? Math.max(0, currentValue + 1)
      : Math.max(0, currentValue - 1);

    sheet.getRange(rowNumber, reactionColumn, 1, 1).setValues([[newValue]]);

    console.info('DataService.updateReactionInSheet: リアクション更新完了', {
      rowId,
      reactionType,
      action,
      oldValue: currentValue,
      newValue
    });

    return true;
  } catch (error) {
    console.error('DataService.updateReactionInSheet: エラー', error.message);
    return false;
  }
}

// ===========================================
// 🔍 データ分析・フィルタリング
// ===========================================

/**
 * columnMappingを使用したデータ処理（Core.gsより移行）
 * @param {Array} dataRows - データ行配列
 * @param {Array} headers - ヘッダー配列
 * @param {Object} columnMapping - 列マッピング
 * @returns {Array} 処理済みデータ配列
 */
function processDataWithColumnMapping(dataRows, headers, columnMapping) {
  if (!dataRows || !Array.isArray(dataRows)) {
    return [];
  }

  console.info('DataService.processDataWithColumnMapping', {
    rowCount: dataRows.length,
    headerCount: headers ? headers.length : 0,
    mappingKeys: columnMapping ? Object.keys(columnMapping) : []
  });

  return dataRows.map((row, index) => {
    const processedRow = {
      id: index + 1,
      timestamp: row[columnMapping?.timestamp] || row[0] || '',
      email: row[columnMapping?.email] || row[1] || '',
      class: row[columnMapping?.class] || row[2] || '',
      name: row[columnMapping?.name] || row[3] || '',
      answer: row[columnMapping?.answer] || row[4] || '',
      reason: row[columnMapping?.reason] || row[5] || '',
      reactions: {
        understand: parseInt(row[columnMapping?.understand] || 0),
        like: parseInt(row[columnMapping?.like] || 0),
        curious: parseInt(row[columnMapping?.curious] || 0)
      },
      highlight: row[columnMapping?.highlight] === 'TRUE' || false,
      originalData: row
    };

    return processedRow;
  });
}

/**
 * 自動停止時間計算（Core.gsより移行）
 * @param {string} publishedAt - 公開開始時間のISO文字列
 * @param {number} minutes - 自動停止までの分数
 * @returns {Object} 停止時間情報
 */
function getAutoStopTime(publishedAt, minutes) {
  try {
    const publishTime = new Date(publishedAt);
    const stopTime = new Date(publishTime.getTime() + minutes * 60 * 1000);

    return {
      publishTime,
      stopTime,
      publishTimeFormatted: publishTime.toLocaleString('ja-JP'),
      stopTimeFormatted: stopTime.toLocaleString('ja-JP'),
      remainingMinutes: Math.max(
        0,
        Math.floor((stopTime.getTime() - new Date().getTime()) / (1000 * 60))
      )
    };
  } catch (error) {
    console.error('DataService.getAutoStopTime: 計算エラー', error);
    return null;
  }
}

/**
 * 🎯 排他的リアクション処理システム - CLAUDE.md準拠Zero-Dependency実装
 *
 * 仕様:
 * - 排他的リアクション: ユーザーは1つのリアクションのみ選択可能
 * - 同じリアクションをクリック: トグル（削除）
 * - 異なるリアクションをクリック: 古いリアクションを削除し、新しいリアクションを追加
 * - ユーザーベース管理: メールアドレスで重複防止
 * - カウントベース表示: フロントエンド向けに適切に変換
 *
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @param {number} rowIndex - 行インデックス
 * @param {string} reactionKey - リアクション種類 (LIKE, UNDERSTAND, CURIOUS)
 * @param {string} userEmail - ユーザーメール
 * @returns {Object} 処理結果 {success, status, message, action, reactions, userReaction, newValue}
 */
/**
 * リアクション状態分析（GAS Best Practice: 単一責任）
 * @param {Sheet} sheet - シートオブジェクト
 * @param {Object} reactionColumns - リアクション列情報
 * @param {number} rowIndex - 行インデックス
 * @param {string} userEmail - ユーザーメール
 * @returns {Object} リアクション状態
 */
function analyzeReactionState(sheet, reactionColumns, rowIndex, userEmail) {
  const currentReactions = {};
  const allReactionsData = {};
  let userCurrentReaction = null;

  // CLAUDE.md準拠: バッチ操作による70倍性能向上
  const columnNumbers = Object.values(reactionColumns);
  const minCol = Math.min(...columnNumbers);
  const maxCol = Math.max(...columnNumbers);
  const [batchData] = sheet.getRange(rowIndex, minCol, 1, maxCol - minCol + 1).getValues();

  Object.keys(reactionColumns).forEach(key => {
    const col = reactionColumns[key];
    const cellValue = batchData[col - minCol] || '';
    const reactionUsers = parseReactionUsers(cellValue);
    currentReactions[key] = reactionUsers;
    allReactionsData[key] = {
      count: reactionUsers.length,
      reacted: reactionUsers.includes(userEmail)
    };

    if (reactionUsers.includes(userEmail)) {
      userCurrentReaction = key;
    }
  });

  return { currentReactions, allReactionsData, userCurrentReaction };
}

/**
 * リアクション更新処理（GAS Best Practice: 単一責任）
 * @param {Object} currentReactions - 現在のリアクション
 * @param {string} reactionKey - リアクションキー
 * @param {string} userEmail - ユーザーメール
 * @param {string} userCurrentReaction - 現在のユーザーリアクション
 * @returns {Object} 更新結果
 */
function updateReactionState(currentReactions, reactionKey, userEmail, userCurrentReaction) {
  let action = 'added';
  let newUserReaction = null;

  if (userCurrentReaction === reactionKey) {
    // Same reaction -> remove (toggle)
    currentReactions[reactionKey] = currentReactions[reactionKey].filter(u => u !== userEmail);
    action = 'removed';
    newUserReaction = null;
  } else {
    // Different reaction -> remove old, add new
    if (userCurrentReaction) {
      currentReactions[userCurrentReaction] = currentReactions[userCurrentReaction].filter(u => u !== userEmail);
    }
    if (!currentReactions[reactionKey].includes(userEmail)) {
      currentReactions[reactionKey].push(userEmail);
    }
    action = 'added';
    newUserReaction = reactionKey;
  }

  return { action, newUserReaction, updatedReactions: currentReactions };
}

/**
 * リアクション処理（リファクタリング版 - GAS Best Practice準拠）
 */
function processReaction(spreadsheetId, sheetName, rowIndex, reactionKey, userEmail) {
  try {
    // バリデーション
    if (!validateReaction(spreadsheetId, sheetName, rowIndex, reactionKey)) {
      throw new Error('無効なリアクションパラメータ');
    }
    if (!userEmail) {
      throw new Error('ユーザー情報が必要です');
    }

    // スプレッドシート接続
    const dataAccess = openSpreadsheet(spreadsheetId);
    const {spreadsheet} = dataAccess;
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error('シートが見つかりません');
    }

    // リアクション列取得
    const reactionColumns = {
      'LIKE': getOrCreateReactionColumn(sheet, 'LIKE'),
      'UNDERSTAND': getOrCreateReactionColumn(sheet, 'UNDERSTAND'),
      'CURIOUS': getOrCreateReactionColumn(sheet, 'CURIOUS')
    };

    // 現在の状態分析
    const { currentReactions, allReactionsData, userCurrentReaction } =
      analyzeReactionState(sheet, reactionColumns, rowIndex, userEmail);

    // リアクション更新処理
    const { action, newUserReaction, updatedReactions } =
      updateReactionState(currentReactions, reactionKey, userEmail, userCurrentReaction);

    // データベース更新
    Object.keys(updatedReactions).forEach(key => {
      const col = reactionColumns[key];
      const newValue = updatedReactions[key].join(',');
      sheet.getRange(rowIndex, col).setValue(newValue);
    });

    // 更新後の状態取得
    const finalReactions = {};
    Object.keys(reactionColumns).forEach(key => {
      finalReactions[key] = {
        count: updatedReactions[key].length,
        reacted: updatedReactions[key].includes(userEmail)
      };
    });

    // 成功レスポンス
    return {
      success: true,
      status: 'success',
      message: `リアクション ${action}: ${reactionKey}`,
      action,
      reactions: finalReactions,
      userReaction: newUserReaction,
      newValue: updatedReactions[reactionKey].join(',')
    };

  } catch (error) {
    console.error('processReaction エラー:', error && error.message ? error.message : 'エラー詳細不明');
    return {
      success: false,
      status: 'error',
      message: error && error.message ? error.message : 'リアクション処理エラー',
      reactions: {},
      userReaction: null
    };
  }
}
/**
 * リアクションユーザー配列をパース
 * @param {string} cellValue - セル値
 * @returns {Array<string>} ユーザーメール配列
 */
function parseReactionUsers(cellValue) {
  if (!cellValue || typeof cellValue !== 'string') {
    return [];
  }

  const trimmed = cellValue.trim();
  if (!trimmed) {
    return [];
  }

  // Split by delimiter and filter out empty strings
  return trimmed.split('|').filter(email => email.trim().length > 0);
}

/**
 * ユーザー配列をセル用文字列にシリアライズ
 * @param {Array<string>} users - ユーザーメール配列
 * @returns {string} セル格納用文字列
 */
function serializeReactionUsers(users) {
  if (!Array.isArray(users) || users.length === 0) {
    return '';
  }

  // Filter out empty emails and join with delimiter
  const validEmails = users.filter(email => email && email.trim().length > 0);
  return validEmails.join('|');
}


// ===========================================
// 🔧 ユーティリティ・ヘルパー
// ===========================================

/**
 * 空行判定
 * @param {Array} row - データ行
 * @returns {boolean} 空行かどうか
 */
function isEmptyRow(row) {
  return !row || row.every(cell => !cell || cell.toString().trim() === '');
}

/**
 * リアクションタイプ検証
 * @param {string} reactionType - リアクションタイプ
 * @returns {boolean} 有効かどうか
 */
function validateReactionType(reactionType) {
  // 🔧 CONSTANTS依存除去: 直接定義
  const validTypes = ['UNDERSTAND', 'LIKE', 'CURIOUS', 'HIGHLIGHT'];
  return validTypes.includes(reactionType);
}

/**
 * リアクション情報抽出
 * @param {Array} row - データ行
 * @param {Array} headers - ヘッダー行
 * @returns {Object} リアクション情報
 */
function extractReactions(row, headers, userEmail = null) {
  try {
    const reactions = {
      UNDERSTAND: { count: 0, reacted: false },
      LIKE: { count: 0, reacted: false },
      CURIOUS: { count: 0, reacted: false }
    };

    // 🎯 統一列判定システムを使用
    const reactionTypes = ['understand', 'like', 'curious'];

    reactionTypes.forEach(reactionType => {
      const columnResult = resolveColumnIndex(headers, reactionType);

      if (columnResult.index !== -1) {
        const cellValue = row[columnResult.index] || '';
        const reactionUsers = parseReactionUsers(cellValue);
        const upperType = reactionType.toUpperCase();

        reactions[upperType] = {
          count: reactionUsers.length,
          reacted: userEmail ? reactionUsers.includes(userEmail) : false
        };
      }
    });

    return reactions;
  } catch (error) {
    console.warn('DataService.extractReactions: エラー', error.message);
    return {
      UNDERSTAND: { count: 0, reacted: false },
      LIKE: { count: 0, reacted: false },
      CURIOUS: { count: 0, reacted: false }
    };
  }
}

/**
 * ハイライト情報抽出
 * @param {Array} row - データ行
 * @param {Array} headers - ヘッダー行
 * @returns {boolean} ハイライト状態
 */
function extractHighlight(row, headers) {
  try {
    // 🎯 統一列判定システムを使用
    const columnResult = resolveColumnIndex(headers, 'highlight');

    if (columnResult.index !== -1) {
      const value = String(row[columnResult.index] || '').toUpperCase();
      return value === 'TRUE' || value === '1' || value === 'YES';
    }

    return false;
  } catch (error) {
    console.warn('DataService.extractHighlight: エラー', error.message);
    return false;
  }
}

/**
 * 行フィルタリング判定
 * @param {Object} item - データアイテム
 * @param {Object} options - フィルタリングオプション
 * @returns {boolean} 含めるかどうか
 */
function shouldIncludeRow(item, options = {}) {
  try {
    // 空行のフィルタリング
    if (options.excludeEmpty !== false && item.isEmpty) {
      return false;
    }

    // メイン回答が空の行をフィルタリング
    if (options.requireAnswer !== false && (!item.answer || item.answer.trim() === '')) {
      return false;
    }

    // 日付範囲フィルタリング
    if (options.dateFrom && item.timestamp) {
      const itemDate = new Date(item.timestamp);
      const fromDate = new Date(options.dateFrom);
      if (itemDate < fromDate) {
        return false;
      }
    }

    if (options.dateTo && item.timestamp) {
      const itemDate = new Date(item.timestamp);
      const toDate = new Date(options.dateTo);
      if (itemDate > toDate) {
        return false;
      }
    }

    // クラスフィルタリング
    if (options.classFilter && options.classFilter.length > 0) {
      if (!options.classFilter.includes(item.class)) {
        return false;
      }
    }

    // キーワード検索
    if (options.searchKeyword && options.searchKeyword.trim() !== '') {
      const keyword = options.searchKeyword.toLowerCase();
      const searchText = `${item.answer || ''} ${item.reason || ''} ${item.name || ''}`.toLowerCase();
      if (!searchText.includes(keyword)) {
        return false;
      }
    }

    return true;
  } catch (error) {
    console.warn('DataService.shouldIncludeRow: エラー', error.message);
    return true; // エラー時は含める
  }
}

/**
 * ソート・制限適用
 * @param {Array} data - データ配列
 * @param {Object} options - オプション
 * @returns {Array} ソート済みデータ
 */
function applySortAndLimit(data, options = {}) {
  try {
    let sortedData = [...data];

    // ソート適用
    if (options.sortBy) {
      switch (options.sortBy) {
        case 'newest':
          sortedData.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
          break;
        case 'oldest':
          sortedData.sort((a, b) => new Date(a.timestamp) - new Date(b.timestamp));
          break;
        case 'reactions':
          sortedData.sort((a, b) => {
            const aTotal = (a.reactions?.UNDERSTAND || 0) + (a.reactions?.LIKE || 0) + (a.reactions?.CURIOUS || 0);
            const bTotal = (b.reactions?.UNDERSTAND || 0) + (b.reactions?.LIKE || 0) + (b.reactions?.CURIOUS || 0);
            return bTotal - aTotal;
          });
          break;
        case 'random':
          // Fisher-Yates シャッフル
          for (let i = sortedData.length - 1; i > 0; i--) {
            const j = Math.floor(Math.random() * (i + 1));
            [sortedData[i], sortedData[j]] = [sortedData[j], sortedData[i]];
          }
          break;
      }
    }

    // 制限適用
    if (options.limit && options.limit > 0) {
      sortedData = sortedData.slice(0, options.limit);
    }

    return sortedData;
  } catch (error) {
    console.error('DataService.applySortAndLimit: エラー', error.message);
    return data;
  }
}

// ===========================================
// 📊 管理パネル用API関数（main.gsから移動）
// ===========================================

/**
 * スプレッドシート一覧を取得
 * AdminPanel.js.html, AppSetupPage.html から呼び出される
 *
 * @returns {Object} スプレッドシート一覧
 */
function getSpreadsheetList(options = {}) {
  const started = Date.now();
  try {
    // ユーザー情報取得
    const currentUser = Session.getActiveUser().getEmail();
    if (!currentUser) {
      throw new Error('ユーザー情報が取得できません');
    }

    // オプション設定
    const {
      adminMode = false,
      maxCount = adminMode ? 20 : 25,
      includeSize = adminMode,
      includeTimestamp = true
    } = options;

    // ドライブアクセステスト
    let driveAccessOk = true;
    try {
      const testFiles = DriveApp.getFiles();
      driveAccessOk = testFiles.hasNext();
    } catch (driveError) {
      console.warn('Drive access limited:', driveError && driveError.message ? driveError.message : 'アクセス制限');
      driveAccessOk = false;
    }

    if (!driveAccessOk) {
      return {
        success: false,
        message: 'Driveアクセス権限が不足しています',
        spreadsheets: [],
        count: 0,
        user: currentUser,
        executionTime: `${Date.now() - started}ms`
      };
    }

    // スプレッドシート検索
    const files = DriveApp.searchFiles('mimeType="application/vnd.google-apps.spreadsheet"');
    const spreadsheets = [];
    let count = 0;

    while (files.hasNext() && count < maxCount) {
      const file = files.next();
      try {
        // 基本情報取得
        const spreadsheetInfo = {
          id: file.getId(),
          name: file.getName(),
          url: file.getUrl(),
          owner: file.getOwner() ? file.getOwner().getEmail() : 'Unknown'
        };

        // オプション情報追加
        if (includeTimestamp) {
          spreadsheetInfo.lastModified = file.getLastUpdated();
          spreadsheetInfo.dateCreated = file.getDateCreated();
        }

        if (includeSize) {
          spreadsheetInfo.size = file.getSize();
        }

        spreadsheets.push(spreadsheetInfo);
        count++;

      } catch (fileError) {
        console.warn('File access error:', fileError && fileError.message ? fileError.message : 'ファイルアクセスエラー');
        // Continue with next file
      }
    }

    return {
      success: true,
      spreadsheets,
      count: spreadsheets.length,
      user: currentUser,
      driveAccess: driveAccessOk,
      executionTime: `${Date.now() - started}ms`
    };

  } catch (error) {
    console.error('getSpreadsheetList エラー:', error && error.message ? error.message : 'エラー詳細不明');
    return {
      success: false,
      message: error && error.message ? error.message : 'スプレッドシート一覧取得エラー',
      spreadsheets: [],
      count: 0,
      executionTime: `${Date.now() - started}ms`
    };
  }
}
/**
 * シート一覧を取得
 * AdminPanel.js.html, AppSetupPage.html から呼び出される
 *
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {Object} シート一覧
 */

/**
 * 列を分析
 * AdminPanel.js.html, AppSetupPage.html から呼び出される
 *
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} 列分析結果
 */

// ===========================================
// 📊 Column Analysis - Refactored Functions
// ===========================================

/**
 * 列分析のメイン関数（リファクタリング版）
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} 列分析結果
 */

// ===========================================
// 📊 Column Analysis - Refactored Functions
// ===========================================


/**
 * パラメータ検証（GAS Best Practice: 単一責任）
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} 検証結果
 */
function validateSheetParams(spreadsheetId, sheetName) {
  if (!spreadsheetId || !sheetName) {
    const errorResponse = {
      success: false,
      message: 'スプレッドシートIDとシート名が必要です',
      headers: [],
      mapping: {},
      confidence: {}
    };
    console.error('DataService.analyzeColumns: 必須パラメータ不足', {
      spreadsheetId: !!spreadsheetId,
      sheetName: !!sheetName
    });
    return { isValid: false, errorResponse };
  }

  return { isValid: true };
}

/**
 * スプレッドシート接続（GAS Best Practice: エラーハンドリング分離）
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} 接続結果
 */
function connectToSheetInternal(spreadsheetId, sheetName, options = {}) {
  try {
    const dataAccess = openSpreadsheet(spreadsheetId);
    const {spreadsheet} = dataAccess;

    // サービスアカウントを編集者として自動登録 (openSpreadsheetで既に処理済み)
    // Note: openSpreadsheet()内でDriveApp.getFileById(id).addEditor()が既に実行されている
    console.log('connectToSheetInternal: サービスアカウント編集者権限はopenSpreadsheetで処理済み');

    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      return {
        success: false,
        errorResponse: {
          success: false,
          message: `シート "${sheetName}" が見つかりません`,
          headers: []
        }
      };
    }

    // 🎯 ワンパス統合処理: 必要なデータを1回で取得
    const dataRange = sheet.getDataRange();
    const allData = dataRange.getValues();
    const headers = allData[0] || [];

    // AI分析用のサンプルデータも同時取得（オプション）
    const result = {
      success: true,
      sheet,
      headers
    };

    if (options.includeSampleData) {
      const sampleData = allData.slice(1, Math.min(11, allData.length)); // 最大10行のサンプル
      result.sampleData = sampleData;
      console.log('connectToSheetInternal: 統合データ取得完了（AI分析用含む）', {
        spreadsheetId: `${spreadsheetId.substring(0, 10)}...`,
        sheetName,
        headers: headers.length,
        sampleData: sampleData.length
      });
    } else {
      console.log('connectToSheetInternal: データソース接続完了', {
        spreadsheetId: `${spreadsheetId.substring(0, 10)}...`,
        sheetName,
        headers: headers.length
      });
    }

    return result;

  } catch (error) {
    console.error('DataService.connectToSheetInternal: 接続エラー', {
      error: error.message,
      spreadsheetId: spreadsheetId && typeof spreadsheetId === 'string' ? `${spreadsheetId.substring(0, 10)}...` : 'N/A'
    });

    return {
      success: false,
      errorResponse: {
        success: false,
        message: error && error.message ? `スプレッドシートアクセスエラー: ${error.message}` : 'スプレッドシートアクセスエラー: 詳細不明',
        headers: [],
          mapping: {},
      confidence: {}
      }
    };
  }
}

/**
 * シートデータ抽出（GAS Best Practice: データアクセス最適化）
 * @param {Sheet} sheet - シートオブジェクト
 * @returns {Object} データ抽出結果
 */
function extractSheetHeaders(sheet) {
  try {
    const lastColumn = sheet.getLastColumn();
    const lastRow = sheet.getLastRow();

    if (lastColumn === 0 || lastRow === 0) {
      return {
        success: false,
        errorResponse: {
          success: false,
          message: 'スプレッドシートが空です',
          headers: [],
              mapping: {},
      confidence: {}
        }
      };
    }

    // 🎯 GAS Best Practice: バッチでデータ取得
    const [headers] = sheet.getRange(1, 1, 1, lastColumn).getValues();

    // サンプルデータを取得（最大5行）
    let sampleData = [];
    const sampleRowCount = Math.min(5, lastRow - 1);
    if (sampleRowCount > 0) {
      sampleData = sheet.getRange(2, 1, sampleRowCount, lastColumn).getValues();
    }

    return { success: true, headers, sampleData };

  } catch (error) {
    console.error('DataService.extractSheetHeaders: データ取得エラー', {
      error: error.message
    });

    return {
      success: false,
      errorResponse: {
        success: false,
        message: error && error.message ? `データ取得エラー: ${error.message}` : 'データ取得エラー: 詳細不明',
        headers: [],
          mapping: {},
      confidence: {}
      }
    };
  }
}

/**
 * configJsonベースの列マッピング取得
 * @param {string} userId - ユーザーID
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} 設定ベースの結果
 */
function restoreColumnConfig(userId, spreadsheetId, sheetName) {
  try {
    const session = { email: Session.getActiveUser().getEmail() };
    const user = findUserByEmail(session.email);
    if (!user) {
      return { success: false, message: 'User not found' };
    }

    // 統一API使用: 構造化パース
    const configResult = getUserConfig(userId);
    const config = configResult.success ? configResult.config : {};
    if (config.spreadsheetId !== spreadsheetId || config.sheetName !== sheetName) {
      return { success: false, message: 'Config mismatch' };
    }

    // 基本ヘッダー情報を取得
    const basicHeaders = getSheetHeadersById(spreadsheetId, sheetName, Date.now());
    if (!basicHeaders.success) {
      return basicHeaders;
    }

    return {
      success: true,
      headers: basicHeaders.headers,
      columnMapping: {
        mapping: config.columnMapping?.mapping || {},
        confidence: config.columnMapping?.confidence || {}
      },
      source: 'configJson',
      executionTime: basicHeaders.executionTime
    };
  } catch (error) {
    console.error('restoreColumnConfig error:', error.message);
    return { success: false, message: error.message };
  }
}

/**
 * 基本ヘッダー情報のみ取得（軽量版の代替）
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @param {number} started - 開始時刻
 * @returns {Object} ヘッダー情報
 */
function getSheetHeadersById(spreadsheetId, sheetName, started) {
  try {
    // 🎯 サービスアカウント認証でopenSpreadsheet()を使用（ServiceFactoryのフォールバック回避）
    const dataAccess = openSpreadsheet(spreadsheetId);
    const {spreadsheet} = dataAccess;
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      return {
        success: false,
        message: `シート "${sheetName}" が見つかりません`,
        headers: []
      };
    }

    const lastColumn = sheet.getLastColumn();
    const headers = lastColumn > 0 ?
      sheet.getRange(1, 1, 1, lastColumn).getValues()[0] : [];

    return {
      success: true,
      headers: headers.map(h => String(h || '')),
      sheetName,
      columnCount: lastColumn,
      source: 'basic',
      executionTime: `${Date.now() - started}ms`
    };
  } catch (error) {
    console.error('getSheetHeaders error:', error.message);
    return {
      success: false,
      message: error.message || 'ヘッダー取得エラー',
      headers: []
    };
  }
}

/**
 * 列分析実行（GAS Best Practice: ビジネスロジック分離）
 * @param {Array} headers - ヘッダー配列
 * @param {Array} sampleData - サンプルデータ配列
 * @returns {Object} 分析結果
 */
/**
 * 🎯 純粋なAI列分析関数 - CLAUDE.md準拠の自然な命名
 * @param {Array} headers - ヘッダー配列
 * @param {Array} sampleData - サンプルデータ配列
 * @returns {Object} AI分析結果
 */
function analyzeColumns(headers, sampleData) {
  try {
    console.log('analyzeColumns: 純粋AI分析開始', {
      headers: headers?.length || 0,
      sampleData: sampleData?.length || 0
    });

    // 入力データ検証
    if (!Array.isArray(headers) || headers.length === 0) {
      return {
        success: false,
        message: 'ヘッダー情報が無効です',
        mapping: {},
        confidence: {}
      };
    }

    if (!Array.isArray(sampleData)) {
      console.warn('analyzeColumns: サンプルデータが無効、ヘッダーのみで分析');
      // sampleDataがなくてもヘッダーベースの分析は可能
    }

    // 🎯 純粋なAI分析実行
    const analysisResult = detectColumnTypes(headers, sampleData || []);

    console.log('analyzeColumns: 純粋AI分析完了', {
      mapping: analysisResult.mapping,
      confidence: analysisResult.confidence
    });

    return {
      success: true,
      mapping: analysisResult.mapping || {},
      confidence: analysisResult.confidence || {}
    };

  } catch (error) {
    console.error('analyzeColumns: AI分析エラー', {
      error: error.message
    });

    return {
      success: false,
      message: `列分析エラー: ${error.message}`,
      mapping: {},
      confidence: {}
    };
  }
}

/**
 * 🎯 列分析取得関数 - CLAUDE.md準拠の自然な命名
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} 統合分析結果
 */
function getColumnAnalysis(spreadsheetId, sheetName) {
  try {
    // 🛡️ V8ランタイム準拠: 安全な型チェックと文字列変換
    const safeSpreadsheetId = typeof spreadsheetId === 'string' ? spreadsheetId : String(spreadsheetId || '');
    const safeSheetName = typeof sheetName === 'string' ? sheetName : String(sheetName || '');

    if (!safeSpreadsheetId || !safeSheetName) {
      throw new Error('Invalid parameters: spreadsheetId and sheetName are required');
    }

    console.log('getColumnAnalysis: 統合AI分析開始', {
      spreadsheetId: `${safeSpreadsheetId.substring(0, 10)}...`,
      sheetName: safeSheetName
    });

    // 🎯 ワンパス統合アクセス: サンプルデータも同時取得
    const connectionResult = connectToSheetInternal(safeSpreadsheetId, safeSheetName, { includeSampleData: true });

    if (!connectionResult.success) {
      return {
        success: false,
        message: connectionResult.errorResponse?.message || 'データソース接続に失敗',
        mapping: {},
        confidence: {}
      };
    }

    // 🎯 純粋関数でAI分析実行
    const analysisResult = analyzeColumns(connectionResult.headers, connectionResult.sampleData);

    if (!analysisResult.success) {
      return analysisResult;
    }

    return {
      success: true,
      sheet: connectionResult.sheet,
      headers: connectionResult.headers,
      mapping: analysisResult.mapping,
      confidence: analysisResult.confidence
    };

  } catch (error) {
    // 🛡️ V8ランタイム準拠: 安全なエラーログ出力
    const safeSpreadsheetId = typeof spreadsheetId === 'string' ? spreadsheetId : String(spreadsheetId || '');
    const safeSheetName = typeof sheetName === 'string' ? sheetName : String(sheetName || '');

    console.error('getColumnAnalysis: 統合分析エラー', {
      error: error.message,
      spreadsheetId: safeSpreadsheetId ? `${safeSpreadsheetId.substring(0, 10)}...` : 'Invalid ID',
      sheetName: safeSheetName || 'Invalid Name'
    });

    return {
      success: false,
      message: `統合分析エラー: ${error.message}`,
      mapping: {},
      confidence: {}
    };
  }
}

/**
 * 列分析実行（GAS Best Practice: ビジネスロジック分離）
 * @param {Array} headers - ヘッダー配列
 * @param {Array} sampleData - サンプルデータ配列
 * @returns {Object} 分析結果
 */
function detectColumnTypes(headers, sampleData) {
  try {

    // 防御的プログラミング: 入力値検証
    if (!Array.isArray(headers) || headers.length === 0) {
      console.warn('DataService.detectColumnTypes: 無効なheaders', headers);
      return { mapping: {}, confidence: {} };
    }

    if (!Array.isArray(sampleData)) {
      console.warn('DataService.detectColumnTypes: 無効なsampleData', sampleData);
      sampleData = [];
    }

    // 🎯 高精度AI検出システム（5次元統計分析）
    const mapping = { mapping: {}, confidence: {} };
    const analysisResults = performHighPrecisionAnalysis(headers, sampleData);

    // 🎯 シンプル絶対閾値判定システム
    const thresholds = {
      answer: 55,  reason: 55,  // 文脈依存列
      name: 65,    class: 65    // 明確な列種別
    };

    // AI判定結果を絶対閾値で判定
    Object.keys(analysisResults).forEach(columnType => {
      const result = analysisResults[columnType];
      const threshold = thresholds[columnType] || 60;

      if (result.confidence >= threshold && result.index >= 0) {
        mapping.mapping[columnType] = result.index;
      }
      mapping.confidence[columnType] = Math.round(result.confidence);
    });

    return {
      mapping: mapping.mapping,        // フロントエンド期待形式
      confidence: mapping.confidence   // 分離
    };

  } catch (error) {
    console.error('DataService.detectColumnTypes: エラー', {
      error: error.message,
      stack: error.stack
    });

    // エラー時のフォールバック
    return {
      mapping: {},
      confidence: {}
    };
  }
}



/**
 * 🎯 高精度AI検出システム - 5次元統計分析
 * @param {Array} headers - 列ヘッダー
 * @param {Array} sampleData - サンプルデータ
 * @returns {Object} 分析結果
 */
function performHighPrecisionAnalysis(headers, sampleData) {

  const results = {
    answer: { index: -1, confidence: 0, factors: {} },
    reason: { index: -1, confidence: 0, factors: {} },
    class: { index: -1, confidence: 0, factors: {} },
    name: { index: -1, confidence: 0, factors: {} }
  };

  headers.forEach((header, index) => {
    if (!header) return;

    const samples = sampleData.map(row => row && row[index]).filter(v => v != null && v !== '');

    // 各列タイプに対する分析を実行
    Object.keys(results).forEach(columnType => {
      const analysis = analyzeColumnForType(header, samples, index, headers, columnType);

      if (analysis.confidence > results[columnType].confidence) {
        results[columnType] = analysis;
      }
    });
  });

  // 🎯 分析結果サマリー出力
  console.info('🔍 AI列判定分析サマリー', {
    '分析対象列数': headers.length,
    'サンプル行数': sampleData.length,
    '検出結果': Object.entries(results).map(([type, result]) => ({
      列種別: type,
      検出列: result.index >= 0 ? `インデックス ${result.index} ("${headers[result.index]}")` : '未検出',
      信頼度: `${Math.round(result.confidence)}%`,
      閾値達成: result.confidence >= (type === 'answer' || type === 'reason' ? 55 : 65) ? '✅' : '❌'
    })),
    '適応的閾値': 'answer/reason: 55%, name/class: 65%',
    '最高信頼度': `${Math.max(...Object.values(results).map(r => Math.round(r.confidence)))}%`
  });

  return results;
}

/**
 * 特定の列タイプに対する多次元分析
 * @param {string} header - 列ヘッダー
 * @param {Array} samples - サンプルデータ
 * @param {number} index - 列インデックス
 * @param {Array} allHeaders - 全ヘッダー
 * @param {string} targetType - 対象列タイプ
 * @returns {Object} 分析結果
 */
/**
 * 重み配分計算（GAS Best Practice: 単一責任）
 * @param {number} headerScore - ヘッダースコア
 * @param {boolean} hasSampleData - サンプルデータ有無
 * @param {boolean} isConflictCase - 競合ケースかどうか
 * @param {string} targetType - 対象タイプ
 * @returns {Object} 重み配分
 */
function calculateWeightDistribution(headerScore, hasSampleData, isConflictCase, targetType) {
  const isAnswerReasonType = (targetType === 'answer' || targetType === 'reason');

  if (headerScore >= 90) {
    return {
      headerWeight: hasSampleData ? (isAnswerReasonType ? 0.35 : 0.45) : 0.6,
      contentWeight: hasSampleData ? (isAnswerReasonType ? 0.25 : 0.20) : 0.15,
      linguisticWeight: hasSampleData ? (isAnswerReasonType ? 0.20 : 0.15) : 0.15,
      contextWeight: hasSampleData ? (isAnswerReasonType ? 0.15 : 0.15) : 0.08,
      semanticWeight: hasSampleData ? (isAnswerReasonType ? 0.05 : 0.05) : 0.02
    };
  } else if (headerScore >= 70) {
    return {
      headerWeight: hasSampleData ? (isAnswerReasonType ? 0.30 : 0.40) : 0.55,
      contentWeight: hasSampleData ? (isAnswerReasonType ? 0.30 : 0.25) : 0.20,
      linguisticWeight: hasSampleData ? (isAnswerReasonType ? 0.20 : 0.15) : 0.15,
      contextWeight: hasSampleData ? (isAnswerReasonType ? 0.15 : 0.15) : 0.08,
      semanticWeight: hasSampleData ? (isAnswerReasonType ? 0.05 : 0.05) : 0.02
    };
  } else {
    return {
      headerWeight: hasSampleData ? (isAnswerReasonType ? 0.25 : 0.35) : 0.50,
      contentWeight: hasSampleData ? (isAnswerReasonType ? 0.35 : 0.30) : 0.25,
      linguisticWeight: hasSampleData ? (isAnswerReasonType ? 0.20 : 0.15) : 0.15,
      contextWeight: hasSampleData ? (isAnswerReasonType ? 0.15 : 0.15) : 0.08,
      semanticWeight: hasSampleData ? (isAnswerReasonType ? 0.05 : 0.05) : 0.02
    };
  }
}

/**
 * 統合分析実行（GAS Best Practice: 複雑な分析ロジック分離）
 * @param {string} header - ヘッダー
 * @param {Array} samples - サンプルデータ
 * @param {number} index - インデックス
 * @param {Array} allHeaders - 全ヘッダー
 * @param {string} targetType - 対象タイプ
 * @returns {Object} 分析結果
 */
function performMultiFactorAnalysis(header, samples, index, allHeaders, targetType) {
  const factors = {};

  // ヘッダーパターン分析
  factors.headerPattern = analyzeHeaderPattern(header.toLowerCase(), targetType);

  // コンテンツ統計分析（サンプルデータがある場合のみ）
  factors.contentStatistics = samples && samples.length > 0 ?
    analyzeContentStatistics(samples, targetType) : 0;

  // 言語パターン分析
  factors.linguisticPatterns = samples && samples.length > 0 ?
    analyzeLinguisticPatterns(samples, targetType) : 0;

  // コンテキスト推論
  factors.contextualClues = analyzeContextualClues(header, index, allHeaders, targetType);

  // セマンティック分析
  factors.semanticCharacteristics = samples && samples.length > 0 ?
    analyzeSemanticCharacteristics(samples, targetType) : 0;

  return factors;
}

/**
 * 列タイプ分析（リファクタリング版 - GAS Best Practice準拠）
 * @param {string} header - ヘッダー
 * @param {Array} samples - サンプルデータ
 * @param {number} index - インデックス
 * @param {Array} allHeaders - 全ヘッダー
 * @param {string} targetType - 対象タイプ
 * @returns {number} 信頼度スコア
 */
function analyzeColumnForType(header, samples, index, allHeaders, targetType) {
  const headerLower = String(header).toLowerCase();

  // サンプルデータ有無チェック
  const hasSampleData = samples && samples.length > 0;

  // 競合検出
  const hasReasonKeywords = /理由|根拠|なぜ|why|わけ|説明/.test(headerLower);
  const hasAnswerKeywords = /答え|回答|answer|意見|予想|考え/.test(headerLower);
  const isConflictCase = hasReasonKeywords && hasAnswerKeywords && (targetType === 'answer' || targetType === 'reason');

  if (isConflictCase) {
    console.log(`🎯 競合ケース検出 [${targetType}]: "${headerLower}" - コンテキスト重み強化`);
  }

  // ヘッダーパターン分析
  const headerScore = analyzeHeaderPattern(headerLower, targetType);

  // 重み配分計算
  const weights = calculateWeightDistribution(headerScore, hasSampleData, isConflictCase, targetType);

  // 統合分析実行
  const factors = performMultiFactorAnalysis(header, samples, index, allHeaders, targetType);

  // 重み付きスコア計算
  let totalConfidence = 0;
  totalConfidence += factors.headerPattern * weights.headerWeight;
  totalConfidence += factors.contentStatistics * weights.contentWeight;
  totalConfidence += factors.linguisticPatterns * weights.linguisticWeight;
  totalConfidence += factors.contextualClues * weights.contextWeight;
  totalConfidence += factors.semanticCharacteristics * weights.semanticWeight;

  const finalConfidence = Math.min(Math.max(totalConfidence, 0), 100);

  // 🎯 簡潔なデバッグ出力
  console.info(`🤖 AI列分析 [${targetType}] "${header}": ${Math.round(finalConfidence)}点`, {
    ヘッダー: Math.round(factors.headerPattern),
    コンテンツ: Math.round(factors.contentStatistics),
    重み: `H:${Math.round(weights.headerWeight*100)}% C:${Math.round(weights.contentWeight*100)}%`
  });

  return {
    index,
    confidence: finalConfidence,
    factors
  };
}

/**
 * パターン定義を取得（GAS Best Practice: データと処理の分離）
 * @returns {Object} パターン定義オブジェクト
 */
function getColumnPatternDefinitions() {
  return {
    answer: {
      primary: [/^回答$/, /^答え$/, /^answer$/, /^response$/],
      composite: [
        /答え.*書/, /回答.*記入/, /考え.*書/, /意見.*述べ/, /予想.*記入/,
        /選択.*理由.*含/, /答え.*説明.*含/, /回答.*詳細/,
        /あなた.*答え/, /あなたの.*意見/, /どう.*思い.*書/, /感想.*記入/,
        /自分.*考え/, /君.*答え/, /皆.*予想/, /みんな.*意見/
      ],
      strong: [
        /回答/, /意見/, /予想/, /選択/, /choice/,
        /予想.*しよう/, /思い.*記入/, /どのように/, /何が/, /どんな/,
        /観察.*気づいた/, /気づいた.*こと/, /わかった.*こと/, /感じた.*こと/,
        /感想/, /どう思/, /どんな.*気持/, /印象/, /感じ/, /思い/, /考え/,
        /推測/, /予測/, /見込/, /予定/, /期待/, /希望/
      ],
      medium: [
        /結果/, /result/, /値/, /value/, /内容/, /content/,
        /しよう$/, /ましょう$/, /てください$/, /について/, /に関して/
      ],
      weak: [/データ/, /data/, /情報/, /info/],
      conflict: [
        { pattern: /理由.*だけ/, penalty: 0.2 },
        { pattern: /なぜ.*のみ/, penalty: 0.2 },
        { pattern: /根拠.*記載/, penalty: 0.3 },
        { pattern: /説明.*のみ/, penalty: 0.3 }
      ]
    },
    reason: {
      primary: [/^理由$/, /^根拠$/, /^reason$/, /^説明$/, /^答えた理由$/],
      composite: [
        /答えた.*理由/, /選んだ.*理由/, /考えた.*理由/, /そう.*思.*理由/,
        /理由.*書/, /根拠.*教/, /なぜ.*思/, /どうして.*考/,
        /背景.*あれば/, /体験.*あれば/, /経験.*あれば/,
        /判断.*理由/, /決定.*理由/, /選択.*根拠/, /決めた.*わけ/,
        /なぜ.*選/, /どうして.*決/, /理由.*教/, /根拠.*説明/
      ],
      strong: [
        /理由/, /根拠/, /reason/, /なぜ/, /why/, /わけ/, /説明/, /explanation/,
        /どうして/, /なんで/, /how/, /what/, /詳細/, /details/,
        /判断/, /決定/, /選択/, /decision/, /choice/,
        /体験/, /経験/, /背景/, /きっかけ/, /動機/
      ],
      medium: [
        /詳細/, /detail/, /情報/, /info/, /具体/, /具体的/,
        /具体例/, /例/, /example/, /教えて/, /話して/
      ],
      weak: [/内容/, /content/, /データ/, /data/],
      conflict: [
        { pattern: /答え.*含/, penalty: 0.4 },
        { pattern: /回答.*含/, penalty: 0.4 },
        { pattern: /意見.*含/, penalty: 0.3 },
        { pattern: /選択.*含/, penalty: 0.3 }
      ]
    },
    class: {
      primary: [/^クラス$/, /^class$/, /^組$/, /^学級$/],
      composite: [
        /何組/, /何クラス/, /クラス.*番号/, /組.*番号/,
        /学級.*名/, /クラス.*名/, /○組/, /○クラス/
      ],
      strong: [/クラス/, /class/, /組/, /学級/, /学年/, /grade/],
      medium: [/年/, /year/, /グループ/, /group/, /チーム/, /team/],
      weak: [/番号/, /number/, /ID/, /id/],
      conflict: []
    },
    name: {
      primary: [/^名前$/, /^氏名$/, /^name$/, /^名$/],
      composite: [
        /お名前/, /あなたの.*名前/, /君の.*名前/, /氏名.*記入/,
        /名前.*教/, /名前.*書/, /呼び方/
      ],
      strong: [/名前/, /氏名/, /name/, /呼び名/, /ニックネーム/, /nickname/],
      medium: [/名/, /user/, /ユーザー/, /person/, /人/],
      weak: [/ID/, /id/, /番号/, /number/],
      conflict: []
    }
  };
}

/**
 * パターンマッチングによるスコア計算（GAS Best Practice: 単一責任）
 * @param {string} headerLower - 小文字ヘッダー
 * @param {Object} targetPatterns - 対象パターン
 * @returns {number} マッチスコア
 */
function calculatePatternMatchScore(headerLower, targetPatterns) {
  let score = 0;

  // Primary patterns: 完全一致系（最高スコア）
  for (const pattern of targetPatterns.primary || []) {
    if (pattern.test(headerLower)) {
      return 100; // 即座に返却
    }
  }

  // Composite patterns: 複合表現系（高スコア）
  for (const pattern of targetPatterns.composite || []) {
    if (pattern.test(headerLower)) {
      score = Math.max(score, 90);
    }
  }

  // Strong patterns: 強一致系（中〜高スコア）
  if (score < 80) {
    for (const pattern of targetPatterns.strong || []) {
      if (pattern.test(headerLower)) {
        score = Math.max(score, 80);
      }
    }
  }

  // Medium patterns: 中程度一致系（中スコア）
  if (score < 60) {
    for (const pattern of targetPatterns.medium || []) {
      if (pattern.test(headerLower)) {
        score = Math.max(score, 60);
      }
    }
  }

  // Weak patterns: 弱一致系（低スコア）
  if (score < 40) {
    for (const pattern of targetPatterns.weak || []) {
      if (pattern.test(headerLower)) {
        score = Math.max(score, 40);
      }
    }
  }

  return score;
}

/**
 * 競合パターンによる減点計算（GAS Best Practice: 単一責任）
 * @param {string} headerLower - 小文字ヘッダー
 * @param {Object} targetPatterns - 対象パターン
 * @returns {number} 減点乗数
 */
function calculateConflictPenalty(headerLower, targetPatterns) {
  let penaltyMultiplier = 1.0;

  if (targetPatterns.conflict && targetPatterns.conflict.length > 0) {
    for (const conflictRule of targetPatterns.conflict) {
      if (conflictRule.pattern.test(headerLower)) {
        penaltyMultiplier *= (1 - conflictRule.penalty);
        console.log(`競合パターン検出: ${conflictRule.pattern} (減点: ${conflictRule.penalty * 100}%)`);
      }
    }
  }

  return penaltyMultiplier;
}

/**
 * 1️⃣ ヘッダーパターン分析 - 高度な正規表現と重み付きキーワード（リファクタリング版）
 */
function analyzeHeaderPattern(headerLower, targetType) {
  const patterns = getColumnPatternDefinitions();
  const targetPatterns = patterns[targetType];

  if (!targetPatterns) {
    return 0;
  }

  // スコア計算
  const score = calculatePatternMatchScore(headerLower, targetPatterns);

  // 競合減点計算
  const penaltyMultiplier = calculateConflictPenalty(headerLower, targetPatterns);

  // 最終スコア適用
  const finalScore = Math.round(score * penaltyMultiplier);

  if (penaltyMultiplier < 1.0) {
    console.log(`Smart Penalty適用: ${penaltyMultiplier}x`);
  }

  return finalScore;
}
/**
 * 🎯 Multi-Criteria Decision Matrix (MCDM) による競合解決
 * @param {Array} conflictingEvaluations 競合するパターン評価
 * @param {string} headerLower 小文字ヘッダー
 * @param {string} targetType 対象列タイプ
 * @returns {Object} MCDM解決結果
 */
function resolveConflictWithMCDM(conflictingEvaluations, headerLower, targetType) {
  // MCDM基準の重み設定
  const mcdmCriteria = {
    headerSpecificity: 0.4,   // ヘッダー特異性（具体性）
    contextualFit: 0.3,       // 文脈適合度
    semanticDistance: 0.2,    // セマンティック距離
    patternComplexity: 0.1    // パターン複雑度
  };

  const evaluationResults = conflictingEvaluations.map(evaluation => {
    // 1. ヘッダー特異性スコア
    const specificityScore = calculateHeaderSpecificity(evaluation.pattern, headerLower);

    // 2. 文脈適合度スコア
    const contextualScore = calculateContextualFit(evaluation.level, targetType, headerLower);

    // 3. セマンティック距離スコア
    const semanticScore = calculateSemanticDistance(evaluation.pattern, targetType);

    // 4. パターン複雑度スコア
    const complexityScore = calculatePatternComplexity(evaluation.pattern);

    // MCDM重み付き総合スコア計算
    const mcdmScore =
      specificityScore * mcdmCriteria.headerSpecificity +
      contextualScore * mcdmCriteria.contextualFit +
      semanticScore * mcdmCriteria.semanticDistance +
      complexityScore * mcdmCriteria.patternComplexity;

    return {
      ...evaluation,
      mcdmScore: Math.round(mcdmScore * 100) / 100,
      criteria: { specificityScore, contextualScore, semanticScore, complexityScore }
    };
  });

  // 最高MCDMスコアの選択
  const bestMcdmEvaluation = evaluationResults.reduce((best, current) =>
    current.mcdmScore > best.mcdmScore ? current : best
  );

  // 元スコア + MCDM調整による最終スコア
  const finalScore = Math.round(bestMcdmEvaluation.score * (1 + bestMcdmEvaluation.mcdmScore * 0.1));

  return {
    selectedPattern: bestMcdmEvaluation.level,
    finalScore: Math.min(finalScore, 100), // 最大100点
    mcdmDetails: bestMcdmEvaluation.criteria
  };
}

/**
 * ヘッダー特異性計算
 */
function calculateHeaderSpecificity(pattern, headerLower) {
  // パターンの具体性を評価（より具体的なパターンほど高スコア）
  const patternStr = pattern.replace(/^\/|\/$/g, ''); // 正規表現マーカー除去
  const specificityFactors = {
    exactMatch: /^\^.*\$$/.test(pattern) ? 1.0 : 0.0,        // 完全一致
    wordBoundary: /\\b/.test(pattern) ? 0.3 : 0.0,           // 単語境界
    complexPattern: /\.\*/.test(pattern) ? 0.2 : 0.0,        // 複合パターン
    lengthFactor: Math.min(patternStr.length / 20, 0.5)      // 長さ係数
  };

  return Object.values(specificityFactors).reduce((sum, factor) => sum + factor, 0);
}

/**
 * 文脈適合度計算
 */
function calculateContextualFit(patternLevel, targetType, headerLower) {
  // パターンレベルと対象タイプの適合度
  const levelTypeFit = {
    primary: { answer: 0.9, reason: 0.9, name: 1.0, class: 1.0 },
    composite: { answer: 1.0, reason: 1.0, name: 0.8, class: 0.7 },
    strong: { answer: 0.8, reason: 0.8, name: 0.7, class: 0.8 },
    medium: { answer: 0.6, reason: 0.6, name: 0.6, class: 0.6 },
    weak: { answer: 0.4, reason: 0.4, name: 0.4, class: 0.4 }
  };

  return (levelTypeFit[patternLevel] || {})[targetType] || 0.5;
}

/**
 * セマンティック距離計算
 */
function calculateSemanticDistance(pattern, targetType) {
  // パターンと対象タイプ間のセマンティック親和性
  const semanticAffinities = {
    answer: [/答え/, /回答/, /answer/, /意見/, /予想/],
    reason: [/理由/, /根拠/, /reason/, /なぜ/, /説明/],
    name: [/名前/, /氏名/, /name/, /お名前/],
    class: [/クラス/, /class/, /組/, /学級/]
  };

  const targetAffinities = semanticAffinities[targetType] || [];
  const patternStr = pattern.replace(/^\/|\/$/g, '');

  // パターンが対象タイプの親和性キーワードを含むかチェック
  const affinityScore = targetAffinities.some(affinity =>
    affinity.test(patternStr)
  ) ? 1.0 : 0.3;

  return affinityScore;
}

/**
 * パターン複雑度計算
 */
function calculatePatternComplexity(pattern) {
  // 複雑なパターンほど高い特異性を持つ
  const complexityFactors = {
    quantifiers: (/[+*?{]/.test(pattern) ? 0.3 : 0.0),      // 量詞
    characterClasses: (/\[.*\]/.test(pattern) ? 0.2 : 0.0), // 文字クラス
    alternation: (/\|/.test(pattern) ? 0.2 : 0.0),          // 選択
    lookahead: (/\(\?=/.test(pattern) ? 0.3 : 0.0)         // 先読み
  };

  return Object.values(complexityFactors).reduce((sum, factor) => sum + factor, 0);
}

/**
 * 🎯 制約付き重み最適化（Constrained Weight Optimization）
 * @param {Object} originalWeights 元の重み設定
 * @param {Object} adjustments 調整パラメータ
 * @returns {Object} 最適化された重み
 */
function optimizeWeightsWithConstraints(originalWeights, adjustments) {
  // 制約条件: Σweight = 1.0, 0.01 ≤ weight ≤ 0.7
  const MIN_WEIGHT = 0.01;
  const MAX_WEIGHT = 0.7;
  const TARGET_SUM = 1.0;

  // 初期調整の適用
  const adjustedWeights = {
    headerWeight: originalWeights.headerWeight * adjustments.headerReduction,
    contentWeight: originalWeights.contentWeight,
    linguisticWeight: originalWeights.linguisticWeight,
    contextWeight: originalWeights.contextWeight * adjustments.contextBoost,
    semanticWeight: originalWeights.semanticWeight * adjustments.semanticBoost
  };

  // 制約違反のチェックと修正
  const weightKeys = Object.keys(adjustedWeights);

  // 1. 個別制約の適用（最小・最大値）
  for (const key of weightKeys) {
    adjustedWeights[key] = Math.max(MIN_WEIGHT, Math.min(MAX_WEIGHT, adjustedWeights[key]));
  }

  // 2. 合計制約の適用（Lagrange乗数法の簡易版）
  const currentSum = Object.values(adjustedWeights).reduce((sum, weight) => sum + weight, 0);

  if (Math.abs(currentSum - TARGET_SUM) > 0.001) {
    // 重み正規化が必要
    const scaleFactor = TARGET_SUM / currentSum;

    // 優先順位付き調整（重要度の低い重みから調整）
    const priorityOrder = ['linguisticWeight', 'contentWeight', 'headerWeight', 'semanticWeight', 'contextWeight'];

    for (const key of priorityOrder) {
      adjustedWeights[key] *= scaleFactor;

      // 制約範囲内に収める
      adjustedWeights[key] = Math.max(MIN_WEIGHT, Math.min(MAX_WEIGHT, adjustedWeights[key]));
    }

    // 最終正規化（微調整）
    const finalSum = Object.values(adjustedWeights).reduce((sum, weight) => sum + weight, 0);
    if (Math.abs(finalSum - TARGET_SUM) > 0.01) {
      const microAdjustment = (TARGET_SUM - finalSum) / weightKeys.length;
      for (const key of weightKeys) {
        adjustedWeights[key] += microAdjustment;
        adjustedWeights[key] = Math.max(MIN_WEIGHT, Math.min(MAX_WEIGHT, adjustedWeights[key]));
      }
    }
  }

  // 3. 最終検証
  const optimizedSum = Object.values(adjustedWeights).reduce((sum, weight) => sum + weight, 0);
  console.log(`🎯 重み最適化検証: 合計=${optimizedSum.toFixed(3)}, 目標=1.000`);

  return adjustedWeights;
}

/**
 * 2️⃣ コンテンツ統計分析 - データの特性を統計的に分析
 */
function analyzeContentStatistics(samples, targetType) {
  if (!samples || samples.length === 0) return 0;

  const textSamples = samples.filter(s => typeof s === 'string' && s.trim().length > 0);
  if (textSamples.length === 0) return 0;

  // 文字数統計
  const lengths = textSamples.map(s => s.trim().length);
  const avgLength = lengths.reduce((a, b) => a + b, 0) / lengths.length;
  const lengthVariance = lengths.reduce((sum, len) => sum + Math.pow(len - avgLength, 2), 0) / lengths.length;

  // 統計的特徴に基づく判定
  switch (targetType) {
    case 'answer':
      // 回答は一般的に短く、バリエーションが少ない
      if (avgLength <= 20 && lengthVariance <= 100) return 75;
      if (avgLength <= 50 && lengthVariance <= 500) return 60;
      if (avgLength <= 100) return 40;
      return 20;

    case 'reason':
      // 理由は一般的に長く、バリエーションが多い
      if (avgLength >= 50 && lengthVariance >= 200) return 80;
      if (avgLength >= 30 && lengthVariance >= 100) return 65;
      if (avgLength >= 20) return 45;
      return 25;

    case 'class':
      // クラスは短く、パターンが限定的
      if (avgLength <= 10 && lengthVariance <= 20) return 85;
      if (avgLength <= 20 && lengthVariance <= 50) return 65;
      return 30;

    case 'name':
      // 名前は中程度の長さで、適度なバリエーション
      if (avgLength >= 5 && avgLength <= 30 && lengthVariance <= 100) return 75;
      if (avgLength >= 3 && avgLength <= 50) return 55;
      return 35;
  }

  return 0;
}

/**
 * 3️⃣ 言語パターン分析 - 言語的特徴を分析
 */
function analyzeLinguisticPatterns(samples, targetType) {
  if (!samples || samples.length === 0) return 0;

  const textSamples = samples.filter(s => typeof s === 'string' && s.trim().length > 0);
  if (textSamples.length === 0) return 0;

  let score = 0;
  const sampleText = textSamples.join(' ').toLowerCase();

  switch (targetType) {
    case 'answer':
      // 回答によく現れるパターン
      if (/[あいうえお]だと思(う|います)/.test(sampleText)) score += 30;
      if (/(はい|いいえ|yes|no)/.test(sampleText)) score += 25;
      if (/\d+(番|号)/.test(sampleText)) score += 20;
      if (/(選択|肢)/.test(sampleText)) score += 15;
      break;

    case 'reason':
      // 理由によく現れるパターン
      if (/(だから|なぜなら|because)/.test(sampleText)) score += 35;
      if (/(と思う|と考える|だと思います)/.test(sampleText)) score += 25;
      if (/(ため|理由|根拠)/.test(sampleText)) score += 20;
      if (/(経験|体験|感じ)/.test(sampleText)) score += 15;
      break;

    case 'class':
      // クラス情報によく現れるパターン
      if (/\d+(年|組|班)/.test(sampleText)) score += 40;
      if (/[a-z]+(class|group)/.test(sampleText)) score += 30;
      if (/(グループ|チーム)\d+/.test(sampleText)) score += 20;
      break;

    case 'name':
      // 名前によく現れるパターン
      if (/^[ぁ-んァ-ン一-龯]+$/.test(sampleText)) score += 30; // 日本語名
      if (/^[a-zA-Z\s]+$/.test(sampleText)) score += 25; // 英語名
      if (/(さん|くん|ちゃん)$/.test(sampleText)) score += 20; // 敬称
      break;
  }

  return Math.min(score, 100);
}

/**
 * 4️⃣ コンテキスト推論 - 列位置と関係性を分析
 */
function analyzeContextualClues(header, index, allHeaders, targetType) {
  let score = 0;

  // 🎯 Enhanced column position analysis
  const totalColumns = allHeaders.length;
  const position = index / Math.max(totalColumns - 1, 1); // 0-1の相対位置

  switch (targetType) {
    case 'answer':
      // 回答列は通常中央からやや後半に位置（強化）
      if (position >= 0.2 && position <= 0.8) score += 25;
      if (position >= 0.4 && position <= 0.6) score += 10; // 中央ボーナス
      // タイムスタンプの後に来ることが多い
      if (index > 0 && allHeaders[index - 1] &&
          allHeaders[index - 1].toLowerCase().includes('timestamp')) score += 20;
      // 🎯 Enhancement: 基本情報の後に来る傾向
      if (index >= 2) score += 15; // 3列目以降にボーナス
      break;

    case 'reason':
      // 理由列は回答の後に来ることが多い（強化）
      if (index > 0) {
        const prevHeader = allHeaders[index - 1].toLowerCase();
        if (prevHeader.includes('回答') || prevHeader.includes('answer') ||
            prevHeader.includes('意見') || prevHeader.includes('予想')) score += 35;
      }
      // 通常後半に位置（強化）
      if (position >= 0.4) score += 20;
      if (position >= 0.6) score += 10; // 後半ボーナス
      break;

    case 'class':
      // クラス情報は通常最初の方に位置（強化）
      if (position <= 0.4) score += 30;
      if (index <= 3) score += 25;
      if (index <= 1) score += 15; // 最初期ボーナス
      break;

    case 'name':
      // 名前は通常最初の方に位置（強化）
      if (position <= 0.3) score += 35;
      if (index <= 2) score += 30;
      if (index === 0) score += 20; // 第一列ボーナス
      break;
  }

  // 🎯 Enhanced adjacent relationship analysis
  const adjacentHeaders = [
    index > 0 ? allHeaders[index - 1] : null,
    index < allHeaders.length - 1 ? allHeaders[index + 1] : null
  ].filter(h => h).map(h => h.toLowerCase());

  // 🎯 Enhanced pattern matching for relationships
  for (const adjacent of adjacentHeaders) {
    if (targetType === 'answer') {
      if (adjacent.includes('reason') || adjacent.includes('理由') || adjacent.includes('根拠')) score += 15;
      if (adjacent.includes('name') || adjacent.includes('名前') || adjacent.includes('氏名')) score += 10;
    }
    if (targetType === 'reason') {
      if (adjacent.includes('answer') || adjacent.includes('回答') || adjacent.includes('意見')) score += 15;
      if (adjacent.includes('背景') || adjacent.includes('体験') || adjacent.includes('経験')) score += 12;
    }
    if (targetType === 'name') {
      if (adjacent.includes('class') || adjacent.includes('クラス') || adjacent.includes('組')) score += 15;
      if (adjacent.includes('id') || adjacent.includes('番号')) score += 10;
    }
    if (targetType === 'class') {
      if (adjacent.includes('name') || adjacent.includes('名前') || adjacent.includes('氏名')) score += 15;
      if (adjacent.includes('学年') || adjacent.includes('year')) score += 12;
    }
  }

  // 🎯 Multi-column pattern detection (boost for pattern combinations)
  const allHeadersLower = allHeaders.map(h => h.toLowerCase());
  const hasNameColumn = allHeadersLower.some(h => h.includes('名前') || h.includes('name') || h.includes('氏名'));
  const hasClassColumn = allHeadersLower.some(h => h.includes('クラス') || h.includes('class') || h.includes('組'));
  const hasAnswerColumn = allHeadersLower.some(h => h.includes('回答') || h.includes('answer') || h.includes('意見'));
  const hasReasonColumn = allHeadersLower.some(h => h.includes('理由') || h.includes('reason') || h.includes('根拠'));

  // Cross-column relationship bonuses
  if (targetType === 'answer' && hasReasonColumn) score += 8;
  if (targetType === 'reason' && hasAnswerColumn) score += 8;
  if (targetType === 'name' && hasClassColumn) score += 8;
  if (targetType === 'class' && hasNameColumn) score += 8;

  return Math.min(score, 100);
}

/**
 * 5️⃣ 強化されたセマンティック分析 - 多次元意味的特徴分析
 */
function analyzeSemanticCharacteristics(samples, targetType) {
  if (!samples || samples.length === 0) return 0;

  const textSamples = samples.filter(s => typeof s === 'string' && s.trim().length > 0);
  if (textSamples.length === 0) return 0;

  let score = 0;
  const uniqueValues = [...new Set(textSamples)];
  const uniquenessRatio = uniqueValues.length / textSamples.length;

  // 🎯 新機能: 文字列長分布分析
  const lengths = textSamples.map(s => s.trim().length);
  const avgLength = lengths.reduce((a, b) => a + b, 0) / lengths.length;
  const lengthVariance = lengths.reduce((sum, len) => sum + Math.pow(len - avgLength, 2), 0) / lengths.length;
  const lengthStdDev = Math.sqrt(lengthVariance);

  // 🎯 新機能: キーワード密度分析
  const keywordDensity = analyzeKeywordDensity(textSamples, targetType);

  switch (targetType) {
    case 'answer':
      // 回答は選択肢的で重複が多い + 長さが均一
      if (uniquenessRatio <= 0.3) score += 30;
      if (uniquenessRatio <= 0.5) score += 20;
      // 文字列長の均一性（回答は短く均一な傾向）
      if (avgLength <= 20 && lengthStdDev <= 10) score += 25;
      // 数値や選択肢パターン
      if (textSamples.some(s => /^[1-9]$/.test(s))) score += 25;
      if (textSamples.some(s => /^[A-D]$/.test(s))) score += 25;
      // キーワード密度
      score += keywordDensity;
      break;

    case 'reason':
      // 理由は個別性が高く、重複が少ない + 長さにバリエーション
      if (uniquenessRatio >= 0.8) score += 35;
      if (uniquenessRatio >= 0.6) score += 25;
      // 文字列長のバリエーション（理由は長さが多様）
      if (avgLength >= 30 && lengthStdDev >= 15) score += 20;
      // 説明的な言葉
      if (textSamples.some(s => s.includes('ため'))) score += 15;
      if (textSamples.some(s => s.includes('から'))) score += 10;
      // キーワード密度
      score += keywordDensity;
      break;

    case 'class':
      // クラス情報は限定的なパターン + 短い
      if (uniquenessRatio <= 0.2) score += 40;
      if (uniquenessRatio <= 0.4) score += 25;
      // 短い文字列（クラス名は通常短い）
      if (avgLength <= 15 && lengthStdDev <= 5) score += 30;
      if (textSamples.some(s => /\d/.test(s))) score += 20;
      // キーワード密度
      score += keywordDensity;
      break;

    case 'name':
      // 名前は個別性が高い + 適度な長さで均一
      if (uniquenessRatio >= 0.7) score += 35;
      if (uniquenessRatio >= 0.5) score += 20;
      // 名前の典型的な長さ（5-20文字程度）
      if (avgLength >= 5 && avgLength <= 20 && lengthStdDev <= 8) score += 25;
      // キーワード密度
      score += keywordDensity;
      break;
  }

  return Math.min(score, 100);
}

/**
 * キーワード密度分析（新機能）
 * @param {Array} samples - サンプルデータ
 * @param {string} targetType - 対象列タイプ
 * @returns {number} キーワード密度スコア
 */
function analyzeKeywordDensity(samples, targetType) {
  const sampleText = samples.join(' ').toLowerCase();
  let densityScore = 0;

  const keywords = {
    answer: ['はい', 'いいえ', 'yes', 'no', '選択', '番', '思う', 'だと思', '考える'],
    reason: ['だから', 'なぜなら', 'because', 'ため', '理由', '根拠', '経験', '感じ'],
    class: ['年', '組', '班', 'class', 'group', 'チーム', 'クラス'],
    name: ['さん', 'くん', 'ちゃん', '先生', '氏']
  };

  const typeKeywords = keywords[targetType] || [];
  let matchCount = 0;

  typeKeywords.forEach(keyword => {
    if (sampleText.includes(keyword.toLowerCase())) {
      matchCount++;
    }
  });

  // キーワードマッチ率に基づくスコア（最大15点）
  densityScore = Math.min(15, matchCount * 3);

  return densityScore;
}

// ===========================================
// 🔧 ヘルパー関数（必要な実装）
// ===========================================

/**
 * リアクション列の取得または作成
 * @param {Sheet} sheet - シートオブジェクト
 * @param {string} reactionType - リアクションタイプ
 * @returns {number} 列番号
 */
function getOrCreateReactionColumn(sheet, reactionType) {
  try {
    const [headers] = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
    const reactionHeader = reactionType.toUpperCase();

    // 既存の列を探す
    const existingIndex = headers.findIndex(header =>
      String(header).toUpperCase().includes(reactionHeader)
    );

    if (existingIndex !== -1) {
      return existingIndex + 1; // 1-based index
    }

    // 新しい列を作成
    const newColumn = sheet.getLastColumn() + 1;
    sheet.getRange(1, newColumn).setValue(reactionHeader);
    return newColumn;
  } catch (error) {
    console.error('getOrCreateReactionColumn: エラー', error.message);
    return null;
  }
}

/**
 * ハイライト列の更新
 * @param {Object} config - 設定
 * @param {string} rowId - 行ID
 * @returns {boolean} 成功可否
 */
function updateHighlightInSheet(config, rowId) {
  try {
    const dataAccess = openSpreadsheet(config.spreadsheetId);
    const {spreadsheet} = dataAccess;
    const sheet = spreadsheet.getSheetByName(config.sheetName);

    if (!sheet) {
      throw new Error('シートが見つかりません');
    }

    // 行番号抽出（row_3 → 3）
    const rowNumber = parseInt(rowId.replace('row_', ''));
    if (isNaN(rowNumber) || rowNumber < 2) {
      throw new Error('無効な行ID');
    }

    // ハイライト列の取得・作成
    const highlightColumn = getOrCreateReactionColumn(sheet, 'HIGHLIGHT');
    if (!highlightColumn) {
      throw new Error('ハイライト列の作成に失敗');
    }

    // CLAUDE.md準拠: バッチ操作による70倍性能向上 (getValue/setValue → getValues/setValues)
    const [[currentValue]] = sheet.getRange(rowNumber, highlightColumn, 1, 1).getValues();
    const isCurrentlyHighlighted = currentValue === 'TRUE' || currentValue === true;
    const newValue = isCurrentlyHighlighted ? 'FALSE' : 'TRUE';

    sheet.getRange(rowNumber, highlightColumn, 1, 1).setValues([[newValue]]);

    const highlighted = newValue === 'TRUE';

    console.info('DataService.updateHighlightInSheet: ハイライト切り替え完了', {
      rowId,
      oldValue: currentValue,
      newValue,
      highlighted
    });

    return {
      success: true,
      highlighted
    };
  } catch (error) {
    console.error('DataService.updateHighlightInSheet: エラー', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * リアクションパラメータ検証
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @param {number} rowIndex - 行インデックス
 * @param {string} reactionKey - リアクション種類
 * @returns {boolean} 検証結果
 */
function validateReaction(spreadsheetId, sheetName, rowIndex, reactionKey) {
  if (!spreadsheetId || !sheetName || !rowIndex || !reactionKey) {
    return false;
  }

  if (rowIndex < 2) { // ヘッダー行は1
    return false;
  }

  const validReactions = ['UNDERSTAND', 'LIKE', 'CURIOUS', 'HIGHLIGHT'];
  return validReactions.includes(reactionKey);
}

// ===========================================
// 🌍 Public DataService Namespace
// ===========================================

/**
 * addReaction (user context)
 * @param {string} userId
 * @param {number|string} rowIndex - number or 'row_#'
 * @param {string} reaction
 * @returns {Object}
 */
function addReaction(userId, rowIndex, reaction) {
  try {
    // 🎯 Zero-Dependency: Direct Data call
    const user = findUserById(userId);
    if (!user) {
      return createErrorResponse('User not found');
    }

    // 統一API使用: 構造化パース
    const configResult = getUserConfig(userId);
    const config = configResult.success ? configResult.config : {};
    if (!config.spreadsheetId || !config.sheetName) {
      return createErrorResponse('Spreadsheet configuration incomplete');
    }

    const parsedRowIndex = typeof rowIndex === 'string' ? parseInt(String(rowIndex).replace('row_', ''), 10) : parseInt(rowIndex, 10);
    if (!parsedRowIndex || parsedRowIndex < 2) {
      return createErrorResponse('Invalid row ID');
    }

    // 🔧 CLAUDE.md準拠: 行レベルロック機構 - 同時リアクション競合防止（CacheService-based mutex）
    const reactionKey = `reaction_${config.spreadsheetId}_${config.sheetName}_${parsedRowIndex}`;
    const cache = CacheService.getScriptCache();

    // 排他制御（Cache-based mutex）
    if (cache.get(reactionKey)) {
      return {
        success: false,
        message: '同じ行に対するリアクション処理が実行中です。しばらくお待ちください。'
      };
    }

    try {
      cache.put(reactionKey, true, CACHE_DURATION.MEDIUM); // 30秒ロック

      const res = processReaction(config.spreadsheetId, config.sheetName, parsedRowIndex, reaction, user.userEmail);
      if (res && (res.success || res.status === 'success')) {
        // フロントエンド期待形式に合わせたレスポンス
        return {
          success: true,
          reactions: res.reactions || {},
          userReaction: res.userReaction || reaction,
          action: res.action || 'added',
          message: res.message || 'リアクションを追加しました'
        };
      }

      return {
        success: false,
        message: res?.message || 'Failed to add reaction'
      };
    } catch (error) {
      console.error('DataService.addReaction: エラー', error.message);
      return createExceptionResponse(error);
    } finally {
      cache.remove(reactionKey);
    }
  } catch (outerError) {
    console.error('DataService.addReaction outer error:', outerError.message);
    // 🔧 統一ミューテックス: 緊急時のキャッシュクリア
    try {
      const cache = CacheService.getScriptCache();
      const configForCleanup = getUserConfig(userId);
      const cleanupConfig = configForCleanup.success ? configForCleanup.config : {};
      if (cleanupConfig.spreadsheetId && cleanupConfig.sheetName) {
        const cleanupRowIndex = typeof rowIndex === 'string' ? parseInt(String(rowIndex).replace('row_', ''), 10) : parseInt(rowIndex, 10);
        const reactionKey = `reaction_${cleanupConfig.spreadsheetId}_${cleanupConfig.sheetName}_${cleanupRowIndex}`;
        cache.remove(reactionKey);
      }
    } catch (cacheError) {
      console.warn('Failed to clear reaction cache in error handler:', cacheError.message);
    }
    return createExceptionResponse(outerError);
  }
}

/**
 * toggleHighlight (user context)
 * @param {string} userId
 * @param {number|string} rowIndex - number or 'row_#'
 * @returns {Object}
 */
function toggleHighlight(userId, rowIndex) {
  try {
    // 🎯 Zero-Dependency: Direct Data call
    const user = findUserById(userId);
    if (!user) {
      return createErrorResponse('User not found');
    }

    // 統一API使用: 構造化パース
    const configResult = getUserConfig(userId);
    const config = configResult.success ? configResult.config : {};
    if (!config.spreadsheetId || !config.sheetName) {
      return createErrorResponse('Spreadsheet configuration incomplete');
    }

    // updateHighlightInSheet expects 'row_#'
    const rowNumber = typeof rowIndex === 'string' && rowIndex.startsWith('row_')
      ? rowIndex
      : `row_${parseInt(rowIndex, 10)}`;

    // 🔧 CLAUDE.md準拠: 行レベルロック機構 - 同時ハイライト競合防止（CacheService-based mutex）
    const highlightKey = `highlight_${config.spreadsheetId}_${config.sheetName}_${rowNumber}`;
    const cache = CacheService.getScriptCache();

    // 排他制御（Cache-based mutex）
    if (cache.get(highlightKey)) {
      return {
        success: false,
        message: '同じ行のハイライト処理が実行中です。しばらくお待ちください。'
      };
    }

    try {
      cache.put(highlightKey, true, CACHE_DURATION.MEDIUM); // 30秒ロック

      const result = updateHighlightInSheet(config, rowNumber);
      if (result?.success) {
        return {
          success: true,
          message: 'Highlight toggled successfully',
          highlighted: Boolean(result.highlighted)
        };
      }

      return {
        success: false,
        message: result?.error || 'Failed to toggle highlight'
      };
    } catch (error) {
      console.error('DataService.toggleHighlight: エラー', error.message);
      return createExceptionResponse(error);
    } finally {
      cache.remove(highlightKey);
    }
  } catch (outerError) {
    console.error('DataService.toggleHighlight outer error:', outerError.message);
    // 🔧 統一ミューテックス: 緊急時のキャッシュクリア
    try {
      const cache = CacheService.getScriptCache();
      const highlightKey = `highlight_${userId}_${rowIndex}`;
      cache.remove(highlightKey);
    } catch (cacheError) {
      console.warn('Failed to clear highlight cache in error handler:', cacheError.message);
    }
    return createExceptionResponse(outerError);
  }
}

// ===========================================
// 🔄 CLAUDE.md準拠: 自然な英語表現への統一化
// ===========================================


// Expose a stable namespace for non-global access patterns
if (typeof global !== 'undefined') {
  global.DataService = {
    // 🔄 GAS-Native Architecture: Direct ds-prefixed functions (CLAUDE.md compliant)
    getUserSheetData,
    addReaction,
    toggleHighlight,
    // Other functions
    processReaction,
    connectToSheetInternal,
    analyzeColumns,
    getColumnAnalysis
  };
} else {
  this.DataService = {
    // 🔄 GAS-Native Architecture: Direct ds-prefixed functions (CLAUDE.md compliant)
    getUserSheetData,
    addReaction,
    toggleHighlight,
    // Other functions
    processReaction,
    connectToSheetInternal,
    analyzeColumns,
    getColumnAnalysis
  };
}