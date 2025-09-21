/**
 * @fileoverview ColumnMappingService - 統一列判定・マッピングサービス
 *
 * 🎯 責任範囲:
 * - 列インデックス解決（マッピング・パターン・位置ベース）
 * - ヘッダーパターン判定
 * - フィールド値抽出（統一API）
 * - 列診断・レポート生成
 * - システム健全性チェック
 *
 * 🔄 CLAUDE.md Best Practices準拠:
 * - Zero-Dependency Architecture（完全独立）
 * - GAS-Native Pattern（直接API）
 * - 単一責任原則（列マッピング専用）
 * - V8ランタイム最適化
 */

// ===========================================
// 🎯 統一列判定システム - CLAUDE.md準拠
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
    console.error('[ERROR] ColumnMappingService.resolveColumnIndex:', error.message || 'Unknown error');
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
        console.warn(`[WARN] ColumnMappingService.extractFieldValueUnified (${fieldType}): Invalid row data`);
      }
      return options.defaultValue || '';
    }

    const columnResult = resolveColumnIndex(headers, fieldType, columnMapping, options);

    // 🔍 詳細なデバッグ情報
    if (options.enableDebug) {
      console.log(`[DEBUG] ColumnMappingService.extractFieldValueUnified (${fieldType}):`, {
        columnIndex: columnResult.index,
        confidence: columnResult.confidence,
        method: columnResult.method,
        rowLength: row.length,
        headersLength: headers ? headers.length : 0
      });
    }

    // 列が見つからない場合の処理
    if (columnResult.index === -1) {
      return handleColumnNotFound(fieldType, row, headers, options);
    }

    // 範囲外チェック
    if (columnResult.index >= row.length) {
      if (options.enableDebug) {
        console.warn(`[WARN] ColumnMappingService.extractFieldValueUnified (${fieldType}): Column index ${columnResult.index} is out of range (row length: ${row.length})`);
      }
      return options.defaultValue || '';
    }

    // 🎯 値抽出実行
    const extractedValue = row[columnResult.index];

    // 🧹 値のクリーンアップ（V8最適化）
    if (extractedValue === null || extractedValue === undefined) {
      return options.defaultValue || '';
    }

    const cleanedValue = String(extractedValue).trim();
    return cleanedValue || (options.defaultValue || '');

  } catch (error) {
    return handleExtractionError(fieldType, error, options);
  }
}

/**
 * 列が見つからない場合の処理
 * @param {string} fieldType - フィールドタイプ
 * @param {Array} row - データ行
 * @param {Array} headers - ヘッダー配列
 * @param {Object} options - オプション設定
 * @returns {*} フォールバック値
 */
function handleColumnNotFound(fieldType, row, headers, options = {}) {
  if (options.enableDebug) {
    console.warn(`[WARN] ColumnMappingService.handleColumnNotFound: Column not found for ${fieldType}`, {
      availableHeaders: headers || [],
      rowLength: row.length
    });
  }

  // フィールドタイプ別のフォールバック戦略
  switch (fieldType) {
    case 'timestamp':
      // タイムスタンプは通常最初の列
      return row[0] || options.defaultValue || '';

    case 'answer':
    case 'opinion':
      // 回答フィールドが見つからない場合は物理位置フォールバック
      return getPhysicalPositionFallback('answer', row) || options.defaultValue || '';

    case 'email': {
      // メールアドレスパターンを検索
      const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
      const emailCandidate = row.find(cell =>
        cell && typeof cell === 'string' && emailPattern.test(cell)
      );
      return emailCandidate || options.defaultValue || '';
    }

    default:
      return options.defaultValue || '';
  }
}

/**
 * 抽出エラー処理
 * @param {string} fieldType - フィールドタイプ
 * @param {Error} error - エラーオブジェクト
 * @param {Object} options - オプション設定
 * @returns {*} エラー時のフォールバック値
 */
function handleExtractionError(fieldType, error, options = {}) {
  const errorMessage = error && error.message ? error.message : 'Unknown extraction error';
  console.error(`[ERROR] ColumnMappingService.extractFieldValueUnified (${fieldType}):`, errorMessage);

  // エラー発生時のフォールバック値
  return options.defaultValue || '';
}

/**
 * 物理位置フォールバック
 * @param {string} fieldType - フィールドタイプ
 * @param {Array} row - データ行
 * @returns {*} フォールバック値
 */
function getPhysicalPositionFallback(fieldType, row) {
  if (!Array.isArray(row) || row.length === 0) {
    return '';
  }

  // フィールドタイプ別の物理位置推測
  const positions = {
    timestamp: 0,
    answer: Math.min(1, row.length - 1),
    reason: Math.min(2, row.length - 1),
    class: Math.min(3, row.length - 1),
    name: Math.min(4, row.length - 1),
    email: Math.min(5, row.length - 1)
  };

  const position = positions[fieldType];
  return position !== undefined ? (row[position] || '') : '';
}

// ===========================================
// 🔍 列診断・レポート生成システム
// ===========================================

/**
 * 列診断レポート生成（統合版）
 * @param {Array} headers - ヘッダー配列
 * @param {Object} columnMapping - 列マッピング
 * @param {Array} requiredFields - 必須フィールド配列
 * @returns {Object} 診断レポート
 */
function generateColumnDiagnosticReport(headers, columnMapping = {}, requiredFields = ['answer', 'reason', 'class', 'name']) {
  const report = {
    timestamp: new Date().toISOString(),
    headers: headers || [],
    requiredFields,
    optionalFields: ['timestamp', 'email'],
    fieldAnalysis: {},
    issues: [],
    recommendations: [],
    summary: {
      totalFields: 0,
      resolvedFields: 0,
      unresolvedFields: 0,
      confidence: {
        high: 0,
        medium: 0,
        low: 0
      },
      overallScore: 0
    }
  };

  try {
    // 全フィールドを分析対象に
    const allFields = [...requiredFields, ...report.optionalFields];
    report.summary.totalFields = allFields.length;

    // フィールド別分析
    allFields.forEach(fieldType => {
      const columnResult = resolveColumnIndex(headers, fieldType, columnMapping);

      const analysis = {
        fieldType,
        resolved: columnResult.index !== -1,
        columnIndex: columnResult.index,
        confidence: columnResult.confidence,
        method: columnResult.method,
        header: columnResult.index !== -1 ? headers[columnResult.index] : null,
        required: requiredFields.includes(fieldType),
        issues: [],
        recommendations: []
      };

      // 解決状況の評価
      if (analysis.resolved) {
        report.summary.resolvedFields++;

        // 信頼度評価
        if (analysis.confidence >= 80) {
          report.summary.confidence.high++;
        } else if (analysis.confidence >= 50) {
          report.summary.confidence.medium++;
        } else {
          report.summary.confidence.low++;
          analysis.issues.push(`信頼度が低い (${analysis.confidence}%)`);
          analysis.recommendations.push('より明確なヘッダー名の使用を推奨');
        }
      } else {
        report.summary.unresolvedFields++;

        if (analysis.required) {
          analysis.issues.push('必須フィールドが見つかりません');
          analysis.recommendations.push('列マッピングを手動設定してください');
          report.issues.push(`必須フィールド '${fieldType}' が見つかりません`);
        } else {
          analysis.issues.push('オプションフィールドが見つかりません');
          analysis.recommendations.push('必要に応じて列マッピングを設定してください');
        }
      }

      report.fieldAnalysis[fieldType] = analysis;
    });

    // 全体スコア計算
    const resolvedWeight = (report.summary.resolvedFields / report.summary.totalFields) * 100;
    const confidenceWeight =
      (report.summary.confidence.high * 1.0 +
       report.summary.confidence.medium * 0.7 +
       report.summary.confidence.low * 0.4) / report.summary.totalFields * 100;

    report.summary.overallScore = Math.round((resolvedWeight + confidenceWeight) / 2);

    // 全体推奨事項
    report.recommendations = generateColumnRecommendations(report);

    return report;

  } catch (error) {
    console.error('generateColumnDiagnosticReport エラー:', error.message);
    report.issues.push(`診断中にエラーが発生しました: ${error.message}`);
    return report;
  }
}

/**
 * 列推奨事項生成
 * @param {Object} report - 診断レポート
 * @returns {Array} 推奨事項配列
 */
function generateColumnRecommendations(report) {
  const recommendations = [];

  try {
    // 未解決必須フィールドの推奨事項
    const unresolvedRequired = Object.values(report.fieldAnalysis)
      .filter(field => field.required && !field.resolved);

    if (unresolvedRequired.length > 0) {
      recommendations.push({
        priority: 'critical',
        type: 'missing_required_fields',
        message: `必須フィールド (${unresolvedRequired.map(f => f.fieldType).join(', ')}) が見つかりません。列マッピングを確認してください。`,
        fields: unresolvedRequired.map(f => f.fieldType)
      });
    }

    // 低信頼度フィールドの推奨事項
    const lowConfidenceFields = Object.values(report.fieldAnalysis)
      .filter(field => field.resolved && field.confidence < 70);

    if (lowConfidenceFields.length > 0) {
      recommendations.push({
        priority: 'medium',
        type: 'low_confidence_fields',
        message: `一部フィールドの判定信頼度が低いです。より明確なヘッダー名を使用することを推奨します。`,
        fields: lowConfidenceFields.map(f => ({ fieldType: f.fieldType, confidence: f.confidence }))
      });
    }

    // 全体スコアベースの推奨事項
    if (report.summary.overallScore < 60) {
      recommendations.push({
        priority: 'high',
        type: 'overall_improvement',
        message: `全体的な列判定精度が低いです (${report.summary.overallScore}%)。スプレッドシートのヘッダー構造を見直すことを推奨します。`
      });
    } else if (report.summary.overallScore < 80) {
      recommendations.push({
        priority: 'medium',
        type: 'partial_improvement',
        message: `列判定精度に改善の余地があります (${report.summary.overallScore}%)。一部のヘッダー名をより具体的にすることを推奨します。`
      });
    }

    return recommendations;

  } catch (error) {
    console.error('generateColumnRecommendations エラー:', error);
    return [{
      priority: 'low',
      type: 'error',
      message: '推奨事項の生成中にエラーが発生しました'
    }];
  }
}

/**
 * 列解決監視機能（パフォーマンス監視用）
 * @param {Array} headers - ヘッダー配列
 * @param {Object} columnMapping - 列マッピング
 * @param {Object} options - 監視オプション
 * @returns {Object} 監視結果
 */
function monitorColumnResolution(headers, columnMapping = {}, options = {}) {
  const startTime = Date.now();
  const monitoringResult = {
    timestamp: new Date().toISOString(),
    performance: {},
    resolutionStats: {},
    systemHealth: 'unknown'
  };

  try {
    const fieldsToTest = options.fields || ['answer', 'reason', 'class', 'name', 'timestamp', 'email'];
    const resolutionResults = [];

    // フィールド別解決テスト
    fieldsToTest.forEach(fieldType => {
      const fieldStartTime = Date.now();
      const result = resolveColumnIndex(headers, fieldType, columnMapping);
      const fieldEndTime = Date.now();

      resolutionResults.push({
        fieldType,
        resolved: result.index !== -1,
        confidence: result.confidence,
        method: result.method,
        executionTime: fieldEndTime - fieldStartTime
      });
    });

    // パフォーマンス統計
    const totalExecutionTime = Date.now() - startTime;
    monitoringResult.performance = {
      totalExecutionTime,
      averageFieldTime: totalExecutionTime / fieldsToTest.length,
      fieldResults: resolutionResults
    };

    // 解決統計
    const resolvedCount = resolutionResults.filter(r => r.resolved).length;
    const highConfidenceCount = resolutionResults.filter(r => r.confidence >= 80).length;

    monitoringResult.resolutionStats = {
      totalFields: fieldsToTest.length,
      resolvedFields: resolvedCount,
      unresolvedFields: fieldsToTest.length - resolvedCount,
      highConfidenceFields: highConfidenceCount,
      resolutionRate: (resolvedCount / fieldsToTest.length) * 100,
      averageConfidence: resolutionResults.reduce((sum, r) => sum + r.confidence, 0) / fieldsToTest.length
    };

    // システム健全性評価
    if (monitoringResult.resolutionStats.resolutionRate >= 90 && monitoringResult.resolutionStats.averageConfidence >= 80) {
      monitoringResult.systemHealth = 'excellent';
    } else if (monitoringResult.resolutionStats.resolutionRate >= 70) {
      monitoringResult.systemHealth = 'good';
    } else if (monitoringResult.resolutionStats.resolutionRate >= 50) {
      monitoringResult.systemHealth = 'fair';
    } else {
      monitoringResult.systemHealth = 'poor';
    }

    return monitoringResult;

  } catch (error) {
    console.error('monitorColumnResolution エラー:', error);
    monitoringResult.systemHealth = 'error';
    monitoringResult.error = error.message;
    return monitoringResult;
  }
}

// ===========================================
// 🔬 統合診断システム（高度）
// ===========================================

/**
 * 統合列診断（サンプルデータ付き高精度版）
 * @param {Array} headers - ヘッダー配列
 * @param {Object} columnMapping - 列マッピング
 * @param {Array} sampleData - サンプルデータ（オプション）
 * @returns {Object} 統合診断結果
 */
function performIntegratedColumnDiagnostics(headers, columnMapping = {}, sampleData = []) {
  const startTime = Date.now();
  const diagnostics = {
    timestamp: new Date().toISOString(),
    executionTime: 0,

    // 基本診断
    basicReport: generateColumnDiagnosticReport(headers, columnMapping),

    // システム健全性
    systemHealth: {
      backend: diagnoseBackendColumnSystem(headers, columnMapping),
      frontend: diagnoseFrontendColumnSystem(columnMapping),
      integration: null // 後で設定
    },

    // 統合テスト結果
    integrationTests: [],

    // 総合評価
    summary: {
      overallScore: 0,
      criticalIssues: 0,
      warnings: 0,
      recommendations: []
    }
  };

  try {
    // 統合テスト実行
    if (sampleData.length > 0) {
      diagnostics.integrationTests = performIntegrationTests(headers, columnMapping, sampleData);
    }

    // システム健全性統合評価
    diagnostics.systemHealth.integration = {
      backendFrontendSync: Math.abs(diagnostics.systemHealth.backend.score - diagnostics.systemHealth.frontend.score) <= 20,
      overallSystemScore: (diagnostics.systemHealth.backend.score + diagnostics.systemHealth.frontend.score) / 2
    };

    // 総合スコア計算
    const scores = [
      diagnostics.basicReport.summary.overallScore,
      diagnostics.systemHealth.backend.score,
      diagnostics.systemHealth.frontend.score
    ];
    diagnostics.summary.overallScore = Math.round(scores.reduce((sum, score) => sum + score, 0) / scores.length);

    // 重大問題・警告カウント
    diagnostics.summary.criticalIssues = diagnostics.basicReport.issues.length;
    diagnostics.summary.warnings = Object.values(diagnostics.basicReport.fieldAnalysis)
      .filter(field => field.issues.length > 0).length;

    // 統合推奨事項
    diagnostics.summary.recommendations = generateSystemRecommendations(diagnostics);

    diagnostics.executionTime = Date.now() - startTime;
    return diagnostics;

  } catch (error) {
    console.error('performIntegratedColumnDiagnostics エラー:', error);
    diagnostics.error = error.message;
    diagnostics.executionTime = Date.now() - startTime;
    return diagnostics;
  }
}

/**
 * バックエンド列システム診断
 * @param {Array} headers - ヘッダー配列
 * @param {Object} columnMapping - 列マッピング
 * @returns {Object} バックエンド診断結果
 */
function diagnoseBackendColumnSystem(headers, columnMapping) {
  const diagnosis = {
    system: 'backend',
    score: 0,
    issues: [],
    strengths: [],
    details: {}
  };

  try {
    // ヘッダーパターン解析
    const patterns = getHeaderPatterns();
    const recognizedHeaders = [];

    headers.forEach((header, index) => {
      if (header && typeof header === 'string') {
        const lowerHeader = header.toLowerCase();

        // パターンマッチング
        Object.entries(patterns).forEach(([fieldType, fieldPatterns]) => {
          const isMatch = fieldPatterns.some(pattern =>
            lowerHeader.includes(pattern.toLowerCase())
          );
          if (isMatch) {
            recognizedHeaders.push({ header, index, fieldType, matchedPattern: fieldPatterns[0] });
          }
        });
      }
    });

    diagnosis.details.recognizedHeaders = recognizedHeaders;
    diagnosis.details.totalHeaders = headers.length;
    diagnosis.details.recognitionRate = headers.length > 0 ? (recognizedHeaders.length / headers.length) * 100 : 0;

    // スコア計算
    let score = 0;

    // ヘッダー認識率
    score += diagnosis.details.recognitionRate * 0.4;

    // 必須フィールドの存在
    const essentialFields = ['answer', 'timestamp'];
    const foundEssential = essentialFields.filter(field =>
      recognizedHeaders.some(h => h.fieldType === field)
    );
    score += (foundEssential.length / essentialFields.length) * 40;

    // オプションフィールドの存在
    const optionalFields = ['reason', 'class', 'name', 'email'];
    const foundOptional = optionalFields.filter(field =>
      recognizedHeaders.some(h => h.fieldType === field)
    );
    score += (foundOptional.length / optionalFields.length) * 20;

    diagnosis.score = Math.round(score);

    // 強み・問題の特定
    if (diagnosis.details.recognitionRate >= 80) {
      diagnosis.strengths.push('ヘッダーパターン認識率が高い');
    } else if (diagnosis.details.recognitionRate < 50) {
      diagnosis.issues.push('ヘッダーパターン認識率が低い');
    }

    if (foundEssential.length === essentialFields.length) {
      diagnosis.strengths.push('必須フィールドがすべて認識されている');
    } else {
      diagnosis.issues.push(`必須フィールドが不足 (${foundEssential.length}/${essentialFields.length})`);
    }

    return diagnosis;

  } catch (error) {
    diagnosis.score = 0;
    diagnosis.issues.push(`バックエンド診断エラー: ${error.message}`);
    return diagnosis;
  }
}

/**
 * フロントエンド列システム診断
 * @param {Object} columnMapping - 列マッピング
 * @returns {Object} フロントエンド診断結果
 */
function diagnoseFrontendColumnSystem(columnMapping) {
  const diagnosis = {
    system: 'frontend',
    score: 0,
    issues: [],
    strengths: [],
    details: {}
  };

  try {
    // 列マッピングの分析
    const mappingAnalysis = {
      hasMappings: columnMapping && Object.keys(columnMapping).length > 0,
      mappingCount: columnMapping ? Object.keys(columnMapping).length : 0,
      validMappings: 0,
      invalidMappings: 0,
      mappingDetails: []
    };

    if (mappingAnalysis.hasMappings) {
      Object.entries(columnMapping).forEach(([fieldType, columnIndex]) => {
        const mappingDetail = {
          fieldType,
          columnIndex,
          isValid: typeof columnIndex === 'number' && columnIndex >= 0
        };

        if (mappingDetail.isValid) {
          mappingAnalysis.validMappings++;
        } else {
          mappingAnalysis.invalidMappings++;
        }

        mappingAnalysis.mappingDetails.push(mappingDetail);
      });
    }

    diagnosis.details = mappingAnalysis;

    // スコア計算
    let score = 0;

    if (mappingAnalysis.hasMappings) {
      // マッピング存在ボーナス
      score += 30;

      // 有効マッピング率
      const validRate = mappingAnalysis.validMappings / mappingAnalysis.mappingCount;
      score += validRate * 50;

      // 必須フィールドマッピング
      const essentialFields = ['answer'];
      const hasEssentialMappings = essentialFields.every(field =>
        columnMapping[field] !== undefined && typeof columnMapping[field] === 'number'
      );
      if (hasEssentialMappings) {
        score += 20;
      }
    } else {
      // マッピングなしの場合は低スコア
      diagnosis.issues.push('列マッピングが設定されていません');
    }

    diagnosis.score = Math.round(score);

    // 強み・問題の特定
    if (mappingAnalysis.hasMappings) {
      if (mappingAnalysis.invalidMappings === 0) {
        diagnosis.strengths.push('すべての列マッピングが有効');
      } else {
        diagnosis.issues.push(`無効な列マッピングがあります (${mappingAnalysis.invalidMappings}件)`);
      }

      if (mappingAnalysis.mappingCount >= 4) {
        diagnosis.strengths.push('十分な数の列マッピングが設定されている');
      }
    }

    return diagnosis;

  } catch (error) {
    diagnosis.score = 0;
    diagnosis.issues.push(`フロントエンド診断エラー: ${error.message}`);
    return diagnosis;
  }
}

/**
 * 統合テスト実行
 * @param {Array} headers - ヘッダー配列
 * @param {Object} columnMapping - 列マッピング
 * @param {Array} sampleData - サンプルデータ
 * @returns {Array} テスト結果配列
 */
function performIntegrationTests(headers, columnMapping, sampleData) {
  const tests = [];

  try {
    const fieldsToTest = ['answer', 'reason', 'class', 'name', 'timestamp', 'email'];

    fieldsToTest.forEach(fieldType => {
      const test = testFieldResolution(headers, fieldType, columnMapping, sampleData);
      tests.push(test);
    });

    return tests;

  } catch (error) {
    console.error('performIntegrationTests エラー:', error);
    return [{
      fieldType: 'error',
      resolved: false,
      severity: 'critical',
      issue: `統合テスト実行エラー: ${error.message}`
    }];
  }
}

/**
 * フィールド解決テスト
 * @param {Array} headers - ヘッダー配列
 * @param {string} fieldType - フィールドタイプ
 * @param {Object} columnMapping - 列マッピング
 * @param {Array} sampleData - サンプルデータ
 * @returns {Object} テスト結果
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