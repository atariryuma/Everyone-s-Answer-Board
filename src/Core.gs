/**
 * @fileoverview StudyQuest - Core Functions (最適化版)
 * 主要な業務ロジックとAPI エンドポイント
 */

// =================================================================
// エラーハンドリング
// =================================================================

// 統合定数使用（constants.gsから）
// 注意: constants.gsがロードされた後に実行される前提
// これらの定数は既にconstants.gsで定義されており、後方互換性のため存在

/**
 * 統一エラーハンドラークラス
 */
class UnifiedErrorHandler {
  constructor() {
    this.sessionId = Utilities.getUuid();
    this.errorCount = 0;
    this.startTime = Date.now();
  }

  /**
   * 構造化エラーログを生成・出力
   * @param {Error|string} error エラーオブジェクトまたはメッセージ
   * @param {string} context エラー発生箇所/関数名
   * @param {string} [severity=ERROR_SEVERITY.MEDIUM] エラーの重要度
   * @param {string} [category=ERROR_CATEGORIES.SYSTEM] エラーカテゴリ
   * @param {Object} [metadata={}] 追加メタデータ
   * @returns {Object} 構造化エラー情報
   */
  logError(error, context, severity = ERROR_SEVERITY.MEDIUM, category = ERROR_CATEGORIES.SYSTEM, metadata = {}) {
    this.errorCount++;

    const errorInfo = this._buildErrorInfo(error, context, severity, category, metadata);

    // 重要度に応じた出力方法
    switch (severity) {
      case ERROR_SEVERITY.CRITICAL:
        console.error('🚨 CRITICAL ERROR [' + context + ']:', JSON.stringify(errorInfo, null, 2));
        break;
      case ERROR_SEVERITY.HIGH:
        console.error('❌ HIGH SEVERITY [' + context + ']:', JSON.stringify(errorInfo, null, 2));
        break;
      case ERROR_SEVERITY.MEDIUM:
        console.warn('⚠️ MEDIUM SEVERITY [' + context + ']:', errorInfo.message, errorInfo.metadata);
        break;
      case ERROR_SEVERITY.LOW:
        console.log('ℹ️ LOW SEVERITY [' + context + ']:', errorInfo.message);
        break;
    }

    return errorInfo;
  }

  /**
   * データベース操作エラーの処理
   * @param {Error|string} error エラーオブジェクトまたはメッセージ
   * @param {string} operation データベース操作の種類
   * @param {Object} [operationDetails={}] 操作詳細
   */
  logDatabaseError(error, operation, operationDetails = {}) {
    try {
      const dbMetadata = {
        operation: operation,
        ...operationDetails,
        retryable: this._isRetryableError(error),
        timestamp: new Date().toISOString(),
      };

      return this.logError(
        error,
        'database.' + operation,
        ERROR_SEVERITY.MEDIUM,
        ERROR_CATEGORIES.DATABASE,
        dbMetadata,
      );
    } catch (loggingError) {
      // エラーログ処理でエラーが発生した場合は、基本的なconsole出力にフォールバック
      console.error('⚠️ LOGGING ERROR [database.' + operation + ']:', {
        originalError: error && error.message ? error.message : String(error),
        loggingError: loggingError.message,
        operationDetails,
        timestamp: new Date().toISOString()
      });
      
      // 元のエラー情報を簡易的に返す
      return {
        message: error && error.message ? error.message : String(error),
        context: 'database.' + operation,
        severity: ERROR_SEVERITY.MEDIUM,
        category: ERROR_CATEGORIES.DATABASE,
        metadata: operationDetails,
        timestamp: new Date().toISOString(),
        loggingFallback: true
      };
    }
  }

  /**
   * バリデーションエラーの処理
   * @param {string} field バリデーション対象フィールド
   * @param {string} value 入力値
   * @param {string} rule バリデーションルール
   * @param {string} message エラーメッセージ
   */
  logValidationError(field, value, rule, message) {
    const validationMetadata = {
      field: field,
      value: typeof value === 'string' ? value.substring(0, 100) : String(value).substring(0, 100),
      rule: rule,
      timestamp: new Date().toISOString(),
    };

    return this.logError(
      message,
      'validation.' + field,
      ERROR_SEVERITY.LOW,
      ERROR_CATEGORIES.VALIDATION,
      validationMetadata,
    );
  }

  /**
   * エラー情報オブジェクトを構築
   * @private
   */
  _buildErrorInfo(error, context, severity, category, metadata) {
    const baseInfo = {
      timestamp: new Date().toISOString(),
      sessionId: this.sessionId,
      errorId: Utilities.getUuid(),
      context: context,
      severity: severity,
      category: category,
      errorNumber: this.errorCount,
      uptime: Date.now() - this.startTime,
    };

    if (error instanceof Error) {
      return {
        ...baseInfo,
        message: error.message,
        name: error.name,
        stack: error.stack,
        metadata: metadata,
      };
    }

    return {
      ...baseInfo,
      message: String(error),
      metadata: metadata,
    };
  }

  /**
   * エラーがリトライ可能かどうかを判定
   * @private
   */
  _isRetryableError(error) {
    if (!error) return false;

    const retryablePatterns = [
      /timeout/i,
      /rate limit/i,
      /service unavailable/i,
      /temporary/i,
      /quota/i,
    ];

    const errorMessage = error.message || String(error);
    return retryablePatterns.some(pattern => pattern.test(errorMessage));
  }
}

// グローバルインスタンス
const globalErrorHandler = new UnifiedErrorHandler();

/**
 * 便利な関数群 - グローバルエラーハンドラーのラッパー
 */

/**
 * 一般的なエラーをログに記録
 */
function logError(error, context, severity, category, metadata) {
  return globalErrorHandler.logError(error, context, severity, category, metadata);
}

/**
 * データベースエラーをログに記録
 */
function logDatabaseError(error, operation, operationDetails) {
  return globalErrorHandler.logDatabaseError(error, operation, operationDetails);
}

/**
 * バリデーションエラーをログに記録
 */
function logValidationError(field, value, rule, message) {
  return globalErrorHandler.logValidationError(field, value, rule, message);
}

// デバッグログ関数の定義（テスト環境対応）
// debugLog関数はdebugConfig.gsで統一定義されていますが、テスト環境でのfallback定義
// debugLog は debugConfig.gs で統一制御されるため、重複定義を削除

if (typeof warnLog === 'undefined') {
  function warnLog(message, ...args) {
    console.warn('[WARN]', message, ...args);
  }
}

if (typeof errorLog === 'undefined') {
  function errorLog(message, ...args) {
    console.error('[ERROR]', message, ...args);
  }
}

if (typeof infoLog === 'undefined') {
  function infoLog(message, ...args) {
    console.log('[INFO]', message, ...args);
  }
}

// デフォルト質問文の共通定数
// 'あなたの考えや気づいたことを教えてください' はconstants.gsで統一管理

// 'そう考える理由や体験があれば教えてください（任意）' はconstants.gsで統一管理

/**
 * セットアップステップを判定する（サーバー側統一実装）
 * フロントエンドとバックエンドで共通のステップ判定ロジックを提供
 * @param {Object} userInfo - ユーザー情報
 * @param {Object} configJson - 設定JSON（オブジェクト形式）
 * @returns {number} setupStep (1-3)
 */
function getSetupStep(userInfo, configJson) {
  
  // 公開停止後の明示的なリセット判定：手動停止時はステップ1に戻す（最優先）
  if (configJson && configJson.appPublished === false && configJson.unpublishReason === 'manual_stop') {
    return 1;
  }
  
  // Step 1: データソース未設定
  if (!userInfo || !userInfo.spreadsheetId || userInfo.spreadsheetId.trim() === '') {
    return 1;
  }
  
  // configが存在しない場合は必ずStep 2
  if (!configJson || typeof configJson !== 'object') {
    return 2;
  }
  
  // Step 3: 公開中の判定（最優先）
  if (configJson.appPublished === true) {
    return 3;
  }
  
  // Step 2 vs Step 3: UI設定状態に基づく判定
  const hasRequiredSettings = (
    configJson.opinionHeader && configJson.opinionHeader.trim() !== '' &&  // 意見列が設定済み
    configJson.activeSheetName && configJson.activeSheetName.trim() !== '' // アクティブシートが選択済み
  );
  
  if (hasRequiredSettings) {
    return 3;
  }
  
  // デフォルト: ステップ2（設定継続中）
  return 2;
}

/**
 * 自動停止時間を計算する
 * @param {string} publishedAt - 公開開始時間のISO文字列
 * @param {number} minutes - 自動停止までの分数
 * @returns {object} 停止時間情報
 */
function getAutoStopTime(publishedAt, minutes) {
  try {
    const publishTime = new Date(publishedAt);
    const stopTime = new Date(publishTime.getTime() + (minutes * 60 * 1000));

    return {
      publishTime: publishTime,
      stopTime: stopTime,
      publishTimeFormatted: publishTime.toLocaleString('ja-JP'),
      stopTimeFormatted: stopTime.toLocaleString('ja-JP'),
      remainingMinutes: Math.max(0, Math.floor((stopTime.getTime() - new Date().getTime()) / (1000 * 60)))
    };
  } catch (error) {
    logError(error, 'calculateAutoStopTime', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.SYSTEM);
    return null;
  }
}

/**
 * アクティブなシートの公開を終了する（公開終了ボタン用）
 * @param {string} requestUserId - リクエスト元のユーザーID（オプション）
 * @returns {object} 公開終了結果
 */
function clearActiveSheet(requestUserId) {
  if (!requestUserId) {
    requestUserId = getUserId();
  }

  verifyUserAccess(requestUserId);
  
  // 統一ロック管理でアクティブシート終了処理を実行
  return executeWithStandardizedLock('WRITE_OPERATION', 'clearActiveSheet', () => {

    const userInfo = getConfigUserInfo(requestUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません。');
    }

    let configJson = {};
    try {
      configJson = JSON.parse(userInfo.configJson || '{}');
    } catch (parseError) {
      warnLog('ConfigJson parse error in stopPublishing:', parseError.message);
      configJson = {};
    }
    // 公開状態のクリア（データソースとシート選択は保持）
    configJson.publishedSheet = ''; // 後方互換性のため残す
    configJson.publishedSheetName = ''; // 正しいプロパティ名
    configJson.publishedSpreadsheetId = ''; // スプレッドシートIDもクリア
    configJson.appPublished = false; // 公開停止

    // データベースに保存
    updateUser(requestUserId, {
      configJson: JSON.stringify(configJson)
    });

    infoLog('clearActiveSheet完了: 公開を停止しました');

    return {
      success: true,
      message: '回答ボードの公開を終了しました',
      status: 'unpublished'
    };
  });
}

/**
 * セットアップステップを決定する統一関数（権威的実装）
 * @param {Object} userInfo - ユーザー情報
 * @param {Object} configJson - 設定JSON
 * @param {Object} options - オプション設定
 * @returns {number} セットアップステップ (1-3)
 */
function determineSetupStepUnified(userInfo, configJson, options = {}) {
  const debugMode = options.debugMode || false;

  // Step 1: データソース未設定
  if (!userInfo || !userInfo.spreadsheetId || userInfo.spreadsheetId.trim() === '') {
    return 1;
  }

  // configJsonが存在しない場合は必ずStep 2
  if (!configJson || typeof configJson !== 'object') {
    return 2;
  }

  const setupStatus = configJson.setupStatus || 'pending';
  const formCreated = !!configJson.formCreated;
  const hasFormUrl = !!(configJson.formUrl && configJson.formUrl.trim());

  // Step 2条件: 明示的な判定
  const isStep2 = (
    setupStatus === 'pending' ||      // 明示的な未完了状態
    setupStatus === 'reconfiguring' || // 再設定中
    setupStatus === 'error' ||        // エラー状態
    !formCreated ||                   // フォーム未作成
    !hasFormUrl                       // フォームURL未設定
  );

  if (isStep2) {
    return 2;
  }

  // Step 3: セットアップ完了（すべての条件をクリア）
  if (setupStatus === 'completed' && formCreated && hasFormUrl) {
    return 3;
  }

  // フォールバック: 不明な状態はStep 2として扱う
  return 2;
}

/**
 * configJsonの構造と型の妥当性を検証する
 * @param {Object} config - 検証対象のconfigオブジェクト
 * @returns {Object} 検証結果 {isValid: boolean, errors: string[]}
 */
function validateConfigJson(config) {
  const errors = [];

  if (!config || typeof config !== 'object') {
    return {
      isValid: false,
      errors: ['configが存在しないか、オブジェクトではありません']
    };
  }

  // 必須フィールドの型チェック
  const requiredFields = {
    setupStatus: 'string',
    formCreated: 'boolean',
    appPublished: 'boolean',
    publishedSheetName: 'string',
    publishedSpreadsheetId: 'string'
  };

  for (const [field, expectedType] of Object.entries(requiredFields)) {
    if (config[field] === undefined) {
      errors.push('必須フィールド \'' + field + '\' が未定義です');
    } else if (typeof config[field] !== expectedType) {
      errors.push('フィールド \'' + field + '\' の型が不正です。期待値: ' + expectedType + ', 実際の値: ' + typeof config[field]);
    }
  }

  // 文字列フィールドの特別な検証
  if (config.publishedSheetName === 'true') {
    errors.push("publishedSheetNameが不正な値 'true' になっています");
  }

  // setupStatusの値チェック
  const validSetupStatuses = ['pending', 'completed', 'error', 'reconfiguring'];
  if (config.setupStatus && !validSetupStatuses.includes(config.setupStatus)) {
    errors.push('setupStatusの値が不正です: ' + config.setupStatus);
  }

  return {
    isValid: errors.length === 0,
    errors: errors
  };
}

/**
 * configJsonを安全にパースし、必要に応じて修復する
 * @param {string} configJsonString - JSONエンコードされた設定文字列
 * @returns {Object} パース済みで検証されたconfig
 */
function parseAndValidateConfigJson(configJsonString) {
  let config = {};

  try {
    if (configJsonString && configJsonString.trim() !== '' && configJsonString !== '{}') {
      config = JSON.parse(configJsonString);
    }
  } catch (parseError) {
    warnLog('configJson解析エラー:', parseError.message);
    config = {};
  }

  // デフォルト値で不足フィールドを補完
  const defaultConfig = {
    setupStatus: 'pending',
    formCreated: false,
    formUrl: '',
    editFormUrl: '',
    appPublished: false,
    publishedSheetName: '',
    publishedSpreadsheetId: '',
    displayMode: 'anonymous',
    showCounts: false,
    sortOrder: 'newest',
    version: '1.0.0',
    lastModified: new Date().toISOString()
  };

  // 不足フィールドをデフォルト値で埋める
  const mergedConfig = { ...defaultConfig, ...config };

  // 不正な値の修正
  if (mergedConfig.publishedSheetName === 'true') {
    mergedConfig.publishedSheetName = '';
    warnLog('publishedSheetNameの不正値を修正しました');
  }

  const validation = validateConfigJson(mergedConfig);
  if (!validation.isValid) {
    warnLog('configJson検証警告:', validation.errors);
  }

  return mergedConfig;
}
/**
 * 統一された自動修復システム（循環参照を回避した安全な実装）
 * @param {Object} userInfo - ユーザー情報
 * @param {Object} configJson - 設定JSON
 * @param {string} userId - ユーザーID
 * @returns {Object} 修復結果 {updated: boolean, configJson: Object, changes: Array}
 */
function performAutoHealing(userInfo, configJson, userId) {
  const changes = [];
  let updated = false;
  const healedConfig = { ...configJson };

  try {
    // 修復ルール1: formUrlが存在するがformCreatedがfalseの場合
    if (healedConfig.formUrl && healedConfig.formUrl.trim() && !healedConfig.formCreated) {
      healedConfig.formCreated = true;
      changes.push('formCreated: false → true (formURL存在)');
      updated = true;
    }

    // 修復ルール2: formCreatedがtrueだがsetupStatusがcompletedでない場合
    if (healedConfig.formCreated && healedConfig.setupStatus !== 'completed') {
      healedConfig.setupStatus = 'completed';
      changes.push('setupStatus: ' + configJson.setupStatus + ' → completed (form作成済み)');
      updated = true;
    }

    // 修復ルール2b: setupStatusがcompletedだがformCreatedがfalseの場合
    if (!healedConfig.formCreated && healedConfig.setupStatus === 'completed') {
      healedConfig.setupStatus = 'pending';
      changes.push('setupStatus: completed → pending (formCreated=false)');
      updated = true;
    }

    // 修復ルール3: publishedSheetNameが存在するがappPublishedがfalseの場合
    // Note: これは公開状態の判定なので、より慎重に処理
    if (healedConfig.publishedSheetName &&
        healedConfig.publishedSheetName.trim() &&
        healedConfig.setupStatus === 'completed' &&
        !healedConfig.appPublished) {
      healedConfig.appPublished = true;
      changes.push('appPublished: false → true (公開シート名存在)');
      updated = true;
    }

    // 修復後の状態検証
    if (updated) {
      const validation = validateConfigJsonState(healedConfig, userInfo);
      if (!validation.isValid) {
        logError('Auto-healing後の状態が無効: ' + validation.errors, 'autoHealConfig', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.DATABASE);
        return { updated: false, configJson: configJson, changes: [] };
      }

      if (validation.warnings.length > 0) {
        warnLog('⚠️ Auto-healing後の警告:', validation.warnings);
      }
    }

    // データベース更新（変更があった場合のみ）
    if (updated && userId) {
      try {
        updateUser(userId, { configJson: JSON.stringify(healedConfig) });
      } catch (updateError) {
        logDatabaseError(updateError, 'autoHealConfigUpdate', { userId: user.userId });
        // DB更新失敗時は元の設定を返す
        return { updated: false, configJson: configJson, changes: [] };
      }
    }

    return { updated, configJson: healedConfig, changes };

  } catch (error) {
    logError(error, 'autoHealConfig', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.SYSTEM);
    return { updated: false, configJson: configJson, changes: [] };
  }
}

/**
 * ConfigJson状態検証・遷移管理システム
 * @param {Object} configJson - 検証対象の設定
 * @param {Object} userInfo - ユーザー情報
 * @returns {Object} 検証結果 {isValid: boolean, errors: Array, warnings: Array}
 */
function validateConfigJsonState(configJson, userInfo) {
  const errors = [];
  const warnings = [];

  if (!configJson || typeof configJson !== 'object') {
    return { isValid: false, errors: ['configJsonが無効です'], warnings: [] };
  }

  // 公開済み状態では厳格なvalidationをバイパス（安定性優先）
  const appPublished = !!configJson.appPublished;
  if (appPublished) {
    return { 
      isValid: true, 
      errors: [], 
      warnings: ['公開済み状態のためvalidationをバイパスしました'] 
    };
  }

  const setupStatus = configJson.setupStatus || 'pending';
  const formCreated = !!configJson.formCreated;
  const hasFormUrl = !!(configJson.formUrl && configJson.formUrl.trim());
  const hasPublishedSheet = !!(configJson.publishedSheetName && configJson.publishedSheetName.trim());

  // 検証ルール1: setupStatus = 'completed' だが必要な要素が不足
  if (setupStatus === 'completed') {
    if (!formCreated) {
      errors.push('setupStatus=completedですが、formCreated=falseです');
    }
    if (!hasFormUrl) {
      errors.push('setupStatus=completedですが、formUrlが未設定です');
    }
  }

  // 検証ルール2: formCreated = true だが関連要素が不足
  if (formCreated && !hasFormUrl) {
    errors.push('formCreated=trueですが、formUrlが未設定です');
  }

  // 検証ルール3: appPublished = true だが公開情報が不足
  if (appPublished) {
    if (!hasPublishedSheet) {
      errors.push('appPublished=trueですが、publishedSheetNameが未設定です');
    }
    if (setupStatus !== 'completed') {
      errors.push('appPublished=trueですが、setupStatus != completedです');
    }
  }

  // 検証ルール4: データソース検証
  if (!userInfo || !userInfo.spreadsheetId) {
    if (setupStatus === 'completed' || formCreated || appPublished) {
      errors.push('データソース未設定ですが、高度な設定が有効になっています');
    }
  }

  // 警告ルール: 推奨設定の確認
  if (setupStatus === 'completed' && formCreated && !appPublished && hasPublishedSheet) {
    warnings.push('公開準備完了していますが、appPublished=falseです');
  }

  return {
    isValid: errors.length === 0,
    errors,
    warnings
  };
}

/**
 * 安全な状態遷移を実行する関数
 * @param {Object} currentConfig - 現在の設定
 * @param {Object} newValues - 新しい値
 * @param {Object} userInfo - ユーザー情報
 * @returns {Object} 遷移結果 {success: boolean, configJson: Object, errors: Array}
 */
function safeStateTransition(currentConfig, newValues, userInfo) {
  const transitionConfig = { ...currentConfig, ...newValues };

  // 遷移前検証
  const validation = validateConfigJsonState(transitionConfig, userInfo);

  if (!validation.isValid) {
    return {
      success: false,
      configJson: currentConfig,
      errors: validation.errors
    };
  }

  // 遷移実行
  infoLog('✅ 状態遷移実行:', Object.keys(newValues).join(', '));
  if (validation.warnings.length > 0) {
    warnLog('⚠️ 状態遷移警告:', validation.warnings.join(', '));
  }

  return {
    success: true,
    configJson: transitionConfig,
    errors: []
  };
}
// =================================================================
// ユーザー情報キャッシュ（関数実行中の重複取得を防ぐ）
// =================================================================

let _executionSheetsServiceCache = null;
/**
 * 関数実行中のSheetsServiceキャッシュをクリア
 */
function clearExecutionSheetsServiceCache() {
  _executionSheetsServiceCache = null;
}

/**
 * 関数実行中のSheetsServiceを取得（キャッシュを使用）
 * @returns {Object} SheetsServiceオブジェクト
 */
function getCachedSheetsService() {
  if (_executionSheetsServiceCache === null) {
    _executionSheetsServiceCache = getSheetsService();
  } else {
  }
  return _executionSheetsServiceCache;
}

/**
 * ヘッダー整合性検証（リアルタイム検証用）
 * @param {string} userId - ユーザーID
 * @returns {Object} 検証結果
 */
function validateHeaderIntegrity(userId) {
  try {

    const userInfo = getOrFetchUserInfo(userId, 'userId');
    if (!userInfo || !userInfo.spreadsheetId) {
      return {
        success: false,
        error: 'User spreadsheet not found',
        timestamp: new Date().toISOString()
      };
    }

    const spreadsheetId = userInfo.spreadsheetId;
    const sheetName = userInfo.sheetName || 'EABDB';

    // 理由列のヘッダー検証を重点的に実施
    const indices = getHeaderIndices(spreadsheetId, sheetName);

    const validationResults = {
      success: true,
      timestamp: new Date().toISOString(),
      spreadsheetId: spreadsheetId,
      sheetName: sheetName,
      headerValidation: {
        reasonColumnIndex: indices[COLUMN_HEADERS.REASON],
        opinionColumnIndex: indices[COLUMN_HEADERS.OPINION],
        hasReasonColumn: indices[COLUMN_HEADERS.REASON] !== undefined,
        hasOpinionColumn: indices[COLUMN_HEADERS.OPINION] !== undefined
      },
      issues: []
    };

    // 理由列の必須チェック
    if (indices[COLUMN_HEADERS.REASON] === undefined) {
      validationResults.success = false;
      validationResults.issues.push('Reason column (理由) not found in headers');
    }

    // 回答列の必須チェック
    if (indices[COLUMN_HEADERS.OPINION] === undefined) {
      validationResults.success = false;
      validationResults.issues.push('Opinion column (回答) not found in headers');
    }

    // ログ出力
    if (validationResults.success) {
      infoLog('✅ Header integrity validation passed');
    } else {
      warnLog('⚠️ Header integrity validation failed:', validationResults.issues);
    }

    return validationResults;
  } catch (error) {
    logError(error, 'validateHeaderIntegrity', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.DATABASE);
    return {
      success: false,
      error: error.message,
      timestamp: new Date().toISOString()
    };
  }
}

/**
 * 実行開始時のキャッシュクリア（新しいリクエスト開始時に呼び出し）
 */
function clearAllExecutionCache() {
  clearExecutionUserInfoCache();
  clearExecutionSheetsServiceCache();
}
// =================================================================
// メインロジック
// =================================================================
/**
 * 意見ヘッダーを安全に取得する関数（テンプレート変数の問題を回避）
 * @param {string} userId - ユーザーID
 * @param {string} sheetName - シート名
 * @returns {string} 意見ヘッダー
 */
function getOpinionHeaderSafely(userId, sheetName) {
  try {
    const userInfo = getOrFetchUserInfo(userId, 'userId', {
      useExecutionCache: true,
      ttl: 300
    });
    if (!userInfo) {
      return 'お題';
    }

    let config = {};
    try {
      config = JSON.parse(userInfo.configJson || '{}');
    } catch (parseError) {
      warnLog('ConfigJson parse error in verifyUserAccessInternal:', parseError.message);
      config = {};
    }
    const sheetConfigKey = 'sheet_' + (config.publishedSheetName || sheetName);
    const sheetConfig = config[sheetConfigKey] || {};

    const opinionHeader = sheetConfig.opinionHeader || config.publishedSheetName || 'お題';

    return opinionHeader;
  } catch (e) {
    logError(e, 'getOpinionHeaderSafely', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.SYSTEM);
    return 'お題';
  }
}

/**
 * Register a new user or update an existing one.
 *
 * @param {string} adminEmail - Email address of the user performing the action.
 * @returns {Object} Registration result including URLs and flags.
 * @throws {Error} When authentication fails or database operations error.
 */
function registerNewUser(adminEmail) {
  const activeUserEmail = getCurrentUserEmail();
  if (!activeUserEmail) {
    throw new Error('認証エラー: アクティブユーザーの情報を取得できませんでした。');
  }
  if (adminEmail !== activeUserEmail) {
    throw new Error('認証エラー: 操作を実行しているユーザーとメールアドレスが一致しません。');
  }

  // ドメイン制限チェック
  const domainInfo = getDeployUserDomainInfo();
  if (domainInfo.deployDomain && domainInfo.deployDomain !== '' && !domainInfo.isDomainMatch) {
    throw new Error('ドメインアクセスが制限されています。許可されたドメイン: ' + domainInfo.deployDomain + ', 現在のドメイン: ' + domainInfo.currentDomain);
  }

  // 既存ユーザーチェック（1ユーザー1行の原則）
  const existingUser = findUserByEmail(adminEmail);
  let userId, appUrls;

  if (existingUser) {
    // 既存ユーザーの場合は最小限の更新のみ（設定は保護）
    userId = existingUser.userId;
    let existingConfig = {};
    try {
      existingConfig = JSON.parse(existingUser.configJson || '{}');
    } catch (parseError) {
      warnLog('ConfigJson parse error in createOrUpdateUser:', parseError.message);
      existingConfig = {};
    }

    // 最終アクセス時刻とアクティブ状態のみ更新（設定は保護）
    updateUser(userId, {
      lastAccessedAt: new Date().toISOString(),
      isActive: 'true'
      // 注意: configJsonは更新しない（既存の設定を保護）
    });

    // キャッシュを無効化して最新状態を反映
    invalidateUserCache(userId, adminEmail, existingUser.spreadsheetId, false);

    infoLog('✅ 既存ユーザーの最終アクセス時刻を更新しました（設定は保護）: ' + adminEmail);
    appUrls = generateUserUrls(userId);

    return {
      userId: userId,
      adminUrl: appUrls.adminUrl,
      viewUrl: appUrls.viewUrl,
      setupRequired: false, // 既存ユーザーはセットアップ完了済みと仮定
      message: 'おかえりなさい！管理パネルへリダイレクトします。',
      isExistingUser: true
    };
  }

  // 新規ユーザーの場合
  userId = Utilities.getUuid();

  /** @type {Object} 初期設定オブジェクト */
  const initialConfig = {
    // セットアップ管理
    setupStatus: 'pending',
    createdAt: new Date().toISOString(),

    // フォーム設定
    formCreated: false,
    formUrl: '',
    editFormUrl: '',

    // 公開設定
    appPublished: false,
    publishedSheetName: '',
    publishedSpreadsheetId: '',

    // 表示設定
    displayMode: 'anonymous',
    showCounts: false,
    sortOrder: 'newest',

    // メタデータ
    version: '1.0.0',
    lastModified: new Date().toISOString()
  };

  /** @type {Object} ユーザーデータオブジェクト */
  const userData = {
    userId: userId,
    adminEmail: adminEmail,
    spreadsheetId: '',
    spreadsheetUrl: '',
    createdAt: new Date().toISOString(),
    configJson: JSON.stringify(initialConfig),
    lastAccessedAt: new Date().toISOString(),
    isActive: 'true'
  };

  try {
    createUser(userData);
    infoLog('✅ データベースに新規ユーザーを登録しました: ' + adminEmail);

    // シンプルなキャッシュクリア
    try {
      CacheService.getScriptCache().remove('user_' + userId);
      CacheService.getScriptCache().remove('email_' + adminEmail);
    } catch (cacheError) {
      // キャッシュクリアの失敗は無視
    }

    // 原子的な作成→検証フロー（最大3秒待機）
    let verificationSuccess = false;
    let createdUser = null;
    const maxWaitTime = 3000; // 3秒に短縮
    const startTime = Date.now();
    let attemptCount = 0;

    // 簡素化された検証戦略（2段階のみ）
    const verificationStages = [
      { delay: 500, method: 'immediate' },      // 500ms後確認
      { delay: 1500, method: 'final' }          // 2秒後最終確認
    ];

    for (const stage of verificationStages) {
      if (Date.now() - startTime > maxWaitTime) break;
      
      Utilities.sleep(stage.delay);
      attemptCount++;
      
      try {
        // 複数の検証方法を試行
        const verificationMethods = [
          () => fetchUserFromDatabase('userId', userId, { forceFresh: true }),
          () => fetchUserFromDatabase('adminEmail', adminEmail, { forceFresh: true })
        ];

        for (const method of verificationMethods) {
          try {
            createdUser = method();
            if (createdUser && createdUser.userId === userId) {
              verificationSuccess = true;
              break;
            }
          } catch (methodError) {
          }
        }
        
        if (verificationSuccess) break;
        
      } catch (stageError) {
        warnLog('registerNewUser: 検証ステージでエラー:', {
          stage: stage.method,
          error: stageError.message,
          elapsed: Date.now() - startTime + 'ms'
        });
      }
    }

    // 継続検証を削除（時間短縮のため）
    // 2段階の検証で不十分な場合はエラーとして早期終了

    // 検証結果の判定
    if (!verificationSuccess || !createdUser) {
      const elapsedTime = Date.now() - startTime;
      errorLog('❌ registerNewUser: データベース同期検証失敗', {
        userId,
        email: adminEmail,
        attempts: attemptCount,
        elapsed: elapsedTime + 'ms',
        maxWaitTime: maxWaitTime + 'ms'
      });
      
      throw new Error('ユーザーデータの作成を確認できませんでした。' + attemptCount + '回の検証を実行しましたが、データベースの同期が完了していません。しばらく待ってから再試行してください。');
    }

    // 検証成功時の確実なキャッシュ設定
    try {
      CacheService.getScriptCache().put('user_' + userId, JSON.stringify(createdUser), 600); // 10分キャッシュ
      CacheService.getScriptCache().put('email_' + adminEmail, JSON.stringify(createdUser), 600);
    } catch (cacheError) {
      warnLog('registerNewUser: キャッシュ設定でエラー:', cacheError.message);
      // キャッシュ設定の失敗は登録成功を妨げない
    }
  } catch (e) {
    // 重複ユーザーエラーのシンプル処理
    /** @type {string} エラーメッセージテキスト */
    const messageText = (e && (e.message || e.toString())) || '';
    if (messageText.indexOf('既に登録されています') !== -1 || messageText.toLowerCase().indexOf('duplicate') !== -1) {
      warnLog('registerNewUser: メール重複検出、既存ユーザー検索中...', adminEmail);
      try {
        /** @type {Object|null} 既存ユーザー情報 */
        const existing = fetchUserFromDatabase('adminEmail', adminEmail, { forceFresh: true });
        if (existing && existing.userId) {
          appUrls = generateUserUrls(existing.userId);
          return {
            userId: existing.userId,
            adminUrl: appUrls.adminUrl,
            viewUrl: appUrls.viewUrl,
            setupRequired: !existing.spreadsheetId,
            message: '既に登録済みのため、管理パネルへ移動します',
            isExistingUser: true
          };
        }
      } catch (lookupError) {
        warnLog('registerNewUser: 既存ユーザー検索でエラー:', lookupError.message);
      }
    }

    errorLog('❌ registerNewUser: ユーザー登録失敗:', { error: e.message, userId, email: adminEmail });
    throw new Error('ユーザー登録に失敗しました。システム管理者に連絡してください。');
  }

  // 成功レスポンスを返す
  appUrls = generateUserUrls(userId);
  return {
    userId: userId,
    adminUrl: appUrls.adminUrl,
    viewUrl: appUrls.viewUrl,
    setupRequired: true,
    message: 'ユーザー登録が完了しました！次にクイックスタートでフォームを作成してください。',
    isExistingUser: false
  };
}

/**
 * リアクションを追加/削除する (マルチテナント対応版)
 * Page.htmlから呼び出される - フロントエンド期待形式に対応
 * @param {string} requestUserId - リクエスト元のユーザーID
 */
function addReaction(requestUserId, rowIndex, reactionKey, sheetName) {
  verifyUserAccess(requestUserId); // 内部でキャッシュクリア済み

  try {
    /** @type {string} リアクション実行ユーザーのメールアドレス */
    const reactingUserEmail = getCurrentUserEmail();
    /** @type {string} ボードオーナーのユーザーID */
    const ownerUserId = requestUserId; // requestUserId を使用

    // ボードオーナーの情報をDBから取得（キャッシュ利用）
    /** @type {Object|null} ボードオーナー情報 */
    const boardOwnerInfo = findUserById(ownerUserId);
    if (!boardOwnerInfo) {
      throw new Error('無効なボードです。');
    }

    /** @type {Object} リアクション処理結果 */
    const result = processReaction(
      boardOwnerInfo.spreadsheetId,
      sheetName,
      rowIndex,
      reactionKey,
      reactingUserEmail
    );

    // processReactionの戻り値から直接リアクション状態を取得（API呼び出し削減）
    if (result && result.status === 'success') {
      return {
        status: "ok",
        reactions: result.reactionStates || {}
      };
    } else {
      throw new Error(result.message || 'リアクションの処理に失敗しました');
    }
  } catch (e) {
    logError(e, 'addReaction', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.SYSTEM, { userId: requestUserId, rowIndex, reaction: reactionKey });
    return {
      status: "error",
      message: e.message
    };
  }
  // finally節のキャッシュクリア削除（verifyUserAccess内でクリア済みのため不要）
}

/**
 * バッチリアクション処理（既存addReaction機能を保持したまま追加）
 * 複数のリアクション操作を効率的に一括処理
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @param {Array} batchOperations - バッチ操作配列 [{rowIndex, reaction, timestamp}, ...]
 * @returns {object} バッチ処理結果
 */
function addReactionBatch(requestUserId, batchOperations) {
  verifyUserAccess(requestUserId);
  clearExecutionUserInfoCache();

  try {
    // 入力検証
    if (!Array.isArray(batchOperations) || batchOperations.length === 0) {
      throw new Error('バッチ操作が無効です');
    }

    // バッチサイズ制限（安全性のため）
    const MAX_BATCH_SIZE = 20;
    if (batchOperations.length > MAX_BATCH_SIZE) {
      throw new Error('バッチサイズが制限を超えています (最大' + MAX_BATCH_SIZE + '件)');
    }
    /** @type {string} リアクション実行ユーザーのメールアドレス */
    const reactingUserEmail = getCurrentUserEmail();
    /** @type {string} ボードオーナーのユーザーID */
    const ownerUserId = requestUserId;

    // ボードオーナーの情報をDBから取得（キャッシュ利用）
    /** @type {Object|null} ボードオーナー情報 */
    const boardOwnerInfo = findUserById(ownerUserId);
    if (!boardOwnerInfo) {
      throw new Error('無効なボードです。');
    }

    // バッチ処理結果を格納
    /** @type {Array} バッチ処理結果配列 */
    const batchResults = [];
    /** @type {Set<number>} 重複行の追跡用Set */
    const processedRows = new Set(); // 重複行の追跡

    // 既存のsheetNameを取得（最初の操作から）
    /** @type {string} シート名 */
    const sheetName = getCurrentSheetName(boardOwnerInfo.spreadsheetId);

    // バッチ操作を順次処理（既存のprocessReaction関数を再利用）
    for (let i = 0; i < batchOperations.length; i++) {
      /** @type {Object} 現在の操作オブジェクト */
      const operation = batchOperations[i];

      try {
        // 入力検証
        if (!operation.rowIndex || !operation.reaction) {
          warnLog('無効な操作をスキップ:', operation);
          continue;
        }

        // 既存のprocessReaction関数を使用（100%互換性保証）
        /** @type {Object} リアクション処理結果 */
        const result = processReaction(
          boardOwnerInfo.spreadsheetId,
          sheetName,
          operation.rowIndex,
          operation.reaction,
          reactingUserEmail
        );

        if (result && result.status === 'success') {
          // 更新後のリアクション情報を取得
          /** @type {Object} 更新されたリアクション情報 */
          const updatedReactions = getRowReactions(
            boardOwnerInfo.spreadsheetId,
            sheetName,
            operation.rowIndex,
            reactingUserEmail
          );

          batchResults.push({
            rowIndex: operation.rowIndex,
            reaction: operation.reaction,
            reactions: updatedReactions,
            status: 'success'
          });

          processedRows.add(operation.rowIndex);
        } else {
          warnLog('リアクション処理失敗:', operation, result.message);
          batchResults.push({
            rowIndex: operation.rowIndex,
            reaction: operation.reaction,
            status: 'error',
            message: result.message || 'リアクション処理失敗'
          });
        }

      } catch (operationError) {
        logError(operationError, 'batchOperation', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.SYSTEM, { operation, batchIndex: i });
        batchResults.push({
          rowIndex: operation.rowIndex,
          reaction: operation.reaction,
          status: 'error',
          message: operationError.message
        });
      }
    }

    // 成功した行の最新状態を収集
    /** @type {Array} 最終結果配列 */
    const finalResults = [];
    processedRows.forEach(function(rowIndex) {
      try {
        /** @type {Object} 最新のリアクション情報 */
        const latestReactions = getRowReactions(
          boardOwnerInfo.spreadsheetId,
          sheetName,
          rowIndex,
          reactingUserEmail
        );
        finalResults.push({
          rowIndex: rowIndex,
          reactions: latestReactions
        });
      } catch (error) {
        warnLog('最終状態取得エラー:', rowIndex, error.message);
      }
    });

    infoLog('✅ バッチリアクション処理完了:', {
      total: batchOperations.length,
      processed: processedRows.size,
      success: batchResults.filter(r => r.status === 'success').length
    });

    return {
      success: true,
      data: finalResults,
      processedCount: batchOperations.length,
      successCount: batchResults.filter(r => r.status === 'success').length,
      timestamp: new Date().toISOString(),
      details: batchResults // デバッグ用詳細情報
    };

  } catch (error) {
    logError(error, 'addReactionBatch', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.SYSTEM, { batchSize: operations.length });

    // バッチ処理失敗時は個別処理にフォールバック可能であることを示す
    return {
      success: false,
      error: error.message,
      fallbackToIndividual: true, // クライアント側が個別処理にフォールバック可能
      timestamp: new Date().toISOString()
    };

  } finally {
    // 実行終了時にユーザー情報キャッシュをクリア
    clearExecutionUserInfoCache();
  }
}

/**
 * 現在のシート名を取得するヘルパー関数
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {string} シート名
 */
function getCurrentSheetName(spreadsheetId) {
  try {
    /** @type {Object} スプレッドシートオブジェクト */
    const spreadsheet = openSpreadsheetOptimized(spreadsheetId);
    /** @type {Array} シート配列 */
    const sheets = spreadsheet.getSheets();

    // デフォルトでは最初のシートを使用
    if (sheets.length > 0) {
      return sheets[0].getName();
    }

    throw new Error('シートが見つかりません');
  } catch (error) {
    warnLog('シート名取得エラー:', error.message);
    return 'Sheet1'; // フォールバック
  }
}

// =================================================================
// データ取得関数
// =================================================================

/**
 * マルチテナント環境でのユーザーアクセス権限を検証します。
 * リクエストを投げたユーザー (Session.getActiveUser().getEmail()) が、
 * requestUserId のデータにアクセスする権限を持っているかを確認します。
 * 権限がない場合はエラーをスローします。
 * @param {string} requestUserId - アクセスを要求しているユーザーのID
 * @throws {Error} 認証エラーまたは権限エラー
 */
function verifyUserAccess(requestUserId) {
  // 型安全性強化: パラメータ検証（nullable対応）
  if (requestUserId != null) { // null と undefined をチェック
    if (typeof requestUserId !== 'string') {
      throw new Error('認証エラー: ユーザーIDは文字列である必要があります');
    }
    if (requestUserId.trim().length === 0) {
      throw new Error('認証エラー: ユーザーIDが空文字列です');
    }
    if (requestUserId.length > 255) {
      throw new Error('認証エラー: ユーザーIDが長すぎます（最大255文字）');
    }
  }

  clearExecutionUserInfoCache(); // キャッシュをクリアして最新のユーザー情報を取得

  // getCurrentUserEmail()を使用して安全にメールアドレスを取得（verifyUserAccessではemail比較が必要）
  const activeUserEmail = getCurrentUserEmail();
  if (!activeUserEmail) {
    throw new Error('認証エラー: アクティブユーザーの情報を取得できませんでした');
  }

  // requestUserIdがnull/undefinedの場合は新規ユーザーとして扱い、基本的なセッション検証のみ実行
  if (requestUserId == null) {
    return;
  }

  const requestedUserInfo = findUserById(requestUserId);

  if (!requestedUserInfo) {
    throw new Error('認証エラー: 指定されたユーザーID (' + requestUserId + ') が見つかりません。');
  }

  const freshUserInfo = fetchUserFromDatabase('userId', requestUserId, {
    retryCount: 0,
    enableDiagnostics: false,
    autoRepair: false,
    clearCache: true
  });
  if (!freshUserInfo) {
    throw new Error('認証エラー: 指定されたユーザーID (' + requestUserId + ') は無効です。');
  }

  // 管理者かどうかを確認
  if (activeUserEmail !== requestedUserInfo.adminEmail) {
    let config = {};
    try {
      config = JSON.parse(requestedUserInfo.configJson || '{}');
    } catch (parseError) {
      warnLog('ConfigJson parse error in getSheetDetails:', parseError.message);
      config = {};
    }
    if (config.appPublished === true) {
      return;
    }
    throw new Error('権限エラー: ' + activeUserEmail + ' はユーザーID ' + requestUserId + ' のデータにアクセスする権限がありません。');
  }
}

/**
 * 公開されたシートのデータを取得 (マルチテナント対応版)
 * Page.htmlから呼び出される - フロントエンド期待形式に対応
 * @param {string} requestUserId - リクエスト元のユーザーID
 */
function getPublishedSheetData(requestUserId, classFilter, sortOrder, adminMode, bypassCache) {
  verifyUserAccess(requestUserId);
  clearExecutionUserInfoCache(); // キャッシュをクリアして最新のユーザー情報を取得

  try {
    // アクティブボード識別子（シート切替時にキーが変わるように）
    /** @type {Object|null} キー作成用ユーザー情報 */
    const userInfoForKey = getOrFetchUserInfo(requestUserId, 'userId', { useExecutionCache: true, ttl: 120 });
    /** @type {Object} キー作成用設定 */
    let cfgForKey = {};
    try { cfgForKey = JSON.parse(userInfoForKey && userInfoForKey.configJson || '{}'); } catch (e) { cfgForKey = {}; }
    /** @type {string} アクティブなスプレッドシートID */
    const activeSsId = cfgForKey.publishedSpreadsheetId || 'none';
    /** @type {string} アクティブなシート名 */
    const activeSheet = cfgForKey.publishedSheetName || 'none';

    // キャッシュキー生成（アクティブなスプレッドシート/シート名を含める）
    /** @type {string} キャッシュリクエストキー */
    const requestKey = 'publishedData_' + requestUserId + '_' + activeSsId + '_' + activeSheet + '_' + classFilter + '_' + sortOrder + '_' + adminMode;

    // キャッシュバイパス時は直接実行
    if (bypassCache === true) {
      return executeGetPublishedSheetData(requestUserId, classFilter, sortOrder, adminMode);
    }

    return cacheManager.get(requestKey, () => {
      return executeGetPublishedSheetData(requestUserId, classFilter, sortOrder, adminMode);
    }, { ttl: 600 }); // 10分間キャッシュ
  } finally {
    // 実行終了時にユーザー情報キャッシュをクリア
    clearExecutionUserInfoCache();
  }
}

/**
 * 実際のデータ取得処理（キャッシュ制御から分離） (マルチテナント対応版)
 */
function executeGetPublishedSheetData(requestUserId, classFilter, sortOrder, adminMode) {
    try {
      /** @type {string} 現在のユーザーID */
    const currentUserId = requestUserId; // requestUserId を使用

      /** @type {Object} ユーザー情報 */
      const userInfo = getOrFetchUserInfo(currentUserId, 'userId', {
        useExecutionCache: true,
        ttl: 300
      });
      if (!userInfo) {
        throw new Error('ユーザー情報が見つかりません');
      }

      /** @type {Object} 設定JSONオブジェクト */
    let configJson = {};
      try {
        configJson = JSON.parse(userInfo.configJson || '{}');
      } catch (parseError) {
        warnLog('ConfigJson parse error in executeGetPublishedSheetData:', parseError.message);
        configJson = {};
      }

    // セットアップ状況を確認
    /** @type {string} セットアップ状態 */
    const setupStatus = configJson.setupStatus || 'pending';

    // 公開対象のスプレッドシートIDとシート名を取得
    /** @type {string} 公開されたスプレッドシートID */
    const publishedSpreadsheetId = configJson.publishedSpreadsheetId;
    /** @type {string} 公開されたシート名 */
    const publishedSheetName = configJson.publishedSheetName;

    if (!publishedSpreadsheetId || !publishedSheetName) {
      if (setupStatus === 'pending') {
        // セットアップ未完了の場合は適切なメッセージを返す
        return {
          status: 'setup_required',
          message: 'セットアップを完了してください。データ準備、シート・列設定、公開設定の順番で進めてください。',
          data: [],
          setupStatus: setupStatus
        };
      }
      throw new Error('公開対象のスプレッドシートまたはシートが設定されていません。');
    }

    // シート固有の設定を取得 (sheetKey is based only on sheet name)
    /** @type {string} シート設定キー */
    const sheetKey = 'sheet_' + publishedSheetName;
    /** @type {Object} シート固有の設定 */
    const sheetConfig = configJson[sheetKey] || {};

    // Check if current user is the board owner
    /** @type {boolean} オーナー権限フラグ */
    const isOwner = (configJson.ownerId === currentUserId);

    // データ取得
    /** @type {Object} シートデータ */
    const sheetData = getSheetData(currentUserId, publishedSheetName, classFilter, sortOrder, adminMode);

    // 診断: スプレッドシートとシートの存在確認
    try {
      if (sheetData.totalCount === 0) {
        /** @type {Object} スプレッドシートオブジェクト */
        const spreadsheet = openSpreadsheetOptimized(publishedSpreadsheetId);
        /** @type {Object} 指定されたシート */
        const sheet = spreadsheet.getSheetByName(publishedSheetName);
        if (sheet) {
          /** @type {number} 最終行番号 */
          const lastRow = sheet.getLastRow();
          /** @type {number} 最終列番号 */
    const lastColumn = sheet.getLastColumn();

          if (lastRow <= 1) {
          }
        }
      }
    } catch (diagnosisError) {
      warnLog('診断処理でエラー:', diagnosisError.message);
    }

    if (sheetData.status === 'error') {
      throw new Error(sheetData.message);
    }

    // Page.html期待形式に変換
    // 設定からヘッダー名を取得。setupStatus未完了時は安全なデフォルト値を使用。
    /** @type {string|undefined} メインヘッダー名 */
    let mainHeaderName;
    if (setupStatus === 'pending') {
      mainHeaderName = 'セットアップ中...';
    } else {
      mainHeaderName = sheetConfig.opinionHeader || COLUMN_HEADERS.OPINION;
    }

    // その他のヘッダーフィールドも安全に取得
    /** @type {string} 理由ヘッダー名 */
    let reasonHeaderName;
    /** @type {string} クラスヘッダー名 */
    let classHeaderName;
    /** @type {string} 名前ヘッダー名 */
    let nameHeaderName;
    if (setupStatus === 'pending') {
      reasonHeaderName = 'セットアップ中...';
      classHeaderName = 'セットアップ中...';
      nameHeaderName = 'セットアップ中...';
    } else {
      reasonHeaderName = sheetConfig.reasonHeader || COLUMN_HEADERS.REASON;
      classHeaderName = sheetConfig.classHeader !== undefined ? sheetConfig.classHeader : COLUMN_HEADERS.CLASS;
      nameHeaderName = sheetConfig.nameHeader !== undefined ? sheetConfig.nameHeader : COLUMN_HEADERS.NAME;
    }

    // ヘッダーインデックスマップを取得（キャッシュされた実際のマッピング）
    /** @type {Object} ヘッダーインデックス */
    const headerIndices = getHeaderIndices(publishedSpreadsheetId, publishedSheetName);

    // 動的列名のマッピング: 設定された名前と実際のヘッダーを照合
    /** @type {Object} マップされたヘッダーインデックス */
    const mappedIndices = mapConfigToActualHeaders({
      opinionHeader: mainHeaderName,
      reasonHeader: reasonHeaderName,
      classHeader: classHeaderName,
      nameHeader: nameHeaderName
    }, headerIndices);

    /** @type {Array} フォーマット済みデータ */
    const formattedData = formatSheetDataForFrontend(sheetData.data, mappedIndices, headerIndices, adminMode, isOwner, sheetData.displayMode);
    // ボードのタイトルを設定された質問文から取得（優先）、フォールバックで実際のヘッダー
    let headerTitle = publishedSheetName || '今日のお題';
    
    // 1. 設定された質問文（opinionHeader）を優先的に使用
    if (mainHeaderName && mainHeaderName.trim()) {
      headerTitle = mainHeaderName;
    } 
    // 2. フォールバック: 実際のスプレッドシートヘッダーから取得
    else if (mappedIndices.opinionHeader !== undefined) {
      for (var actualHeader in headerIndices) {
        if (headerIndices[actualHeader] === mappedIndices.opinionHeader) {
          headerTitle = actualHeader;
          break;
        }
      }
    }

    /** @type {string} 最終表示モード */
    const finalDisplayMode = (adminMode === true) ? DISPLAY_MODES.NAMED : (configJson.displayMode || DISPLAY_MODES.ANONYMOUS);

    /** @type {Object} レスポンス結果 */
    const result = {
      header: headerTitle,
      sheetName: publishedSheetName,
      showCounts: (adminMode === true) ? true : (configJson.showCounts === true),
      displayMode: finalDisplayMode,
      data: formattedData,
      rows: formattedData // 後方互換性のため
    };

    return result;

  } catch (e) {
    logError(e, 'getPublishedSheetData', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.DATABASE, { userId, publishedSpreadsheetId, publishedSheetName });
    return {
      status: 'error',
      message: 'データの取得に失敗しました: ' + e.message,
      data: [],
      rows: []
    };
  }
}

/**
 * 増分データ取得機能：指定された基準点以降の新しいデータのみを取得 (マルチテナント対応版)
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @param {string} classFilter - クラスフィルタ
 * @param {string} sortOrder - ソート順
 * @param {boolean} adminMode - 管理者モード
 * @param {number} sinceRowCount - この行数以降のデータを取得
 * @returns {object} 新しいデータのみを含む結果
 */
function getIncrementalSheetData(requestUserId, classFilter, sortOrder, adminMode, sinceRowCount) {
  verifyUserAccess(requestUserId);
  try {

    /** @type {string} 現在のユーザーID */
    const currentUserId = requestUserId; // requestUserId を使用

    /** @type {Object} ユーザー情報 */
    const userInfo = getOrFetchUserInfo(currentUserId, 'userId', {
      useExecutionCache: true,
      ttl: 300
    });
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    /** @type {Object} 設定JSONオブジェクト */
    let configJson = {};
    try {
      configJson = JSON.parse(userInfo.configJson || '{}');
    } catch (parseError) {
      warnLog('ConfigJson parse error in getSheetData:', parseError.message);
      configJson = {};
    }
    /** @type {string} セットアップ状態 */
    const setupStatus = configJson.setupStatus || 'pending';
    /** @type {string} 公開されたスプレッドシートID */
    const publishedSpreadsheetId = configJson.publishedSpreadsheetId;
    /** @type {string} 公開されたシート名 */
    const publishedSheetName = configJson.publishedSheetName;

    if (!publishedSpreadsheetId || !publishedSheetName) {
      if (setupStatus === 'pending') {
        // セットアップ未完了の場合は適切なメッセージを返す
        return {
          status: 'setup_required',
          message: 'セットアップを完了してください。',
          incrementalData: [],
          setupStatus: setupStatus
        };
      }
      throw new Error('公開対象のスプレッドシートまたはシートが設定されていません。');
    }

    // スプレッドシートとシートを取得
    /** @type {Object} スプレッドシートオブジェクト */
    const ss = openSpreadsheetOptimized(publishedSpreadsheetId);

      /** @type {Object} 指定されたシート */
      const sheet = ss.getSheetByName(publishedSheetName);

    if (!sheet) {
      throw new Error('指定されたシートが見つかりません: ' + publishedSheetName);
    }

    /** @type {number} スプレッドシートの最終行 */
    const lastRow = sheet.getLastRow(); // スプレッドシートの最終行
    /** @type {number} ヘッダー行は1行目と仮定 */
    const headerRow = 1; // ヘッダー行は1行目と仮定

    // 実際に読み込むべき開始行を計算 (sinceRowCountはデータ行数なので、+1してヘッダーを考慮)
    // sinceRowCountが0の場合、ヘッダーの次の行から読み込む
    /** @type {number} 読み取り開始行 */
    let startRowToRead = sinceRowCount + headerRow + 1;

    // 新しいデータがない場合
    if (lastRow < startRowToRead) {
      return {
        header: '', // 必要に応じて設定
        sheetName: publishedSheetName,
        showCounts: configJson.showCounts === true,
        displayMode: configJson.displayMode || DISPLAY_MODES.ANONYMOUS,
        data: [],
        rows: [],
        totalCount: lastRow - headerRow, // ヘッダーを除いたデータ総数
        newCount: 0,
        isIncremental: true
      };
    }

    // 読み込む行数
    /** @type {number} 読み取る行数 */
    const numRowsToRead = lastRow - startRowToRead + 1;

    // 必要なデータのみをスプレッドシートから直接取得
    // getRange(row, column, numRows, numColumns)
    // ここでは全列を取得すると仮定 (A列から最終列まで)
    /** @type {number} 最終列番号 */
    const lastColumn = sheet.getLastColumn();
    /** @type {Array} 生のスプレッドシートデータ */
    const rawNewData = sheet.getRange(startRowToRead, 1, numRowsToRead, lastColumn).getValues();
    // ヘッダーインデックスマップを取得（キャッシュされた実際のマッピング）
    /** @type {Object} ヘッダーインデックス */
    const headerIndices = getHeaderIndices(publishedSpreadsheetId, publishedSheetName);

    // 動的列名のマッピング: 設定された名前と実際のヘッダーを照合
    /** @type {Object} シート設定 */
    const sheetConfig = configJson['sheet_' + publishedSheetName] || {};
    /** @type {string} メインヘッダー名 */
    const mainHeaderName = sheetConfig.opinionHeader || COLUMN_HEADERS.OPINION;
    /** @type {string} 理由ヘッダー名 */
    const reasonHeaderName = sheetConfig.reasonHeader || COLUMN_HEADERS.REASON;
    /** @type {string} クラスヘッダー名 */
    const classHeaderName = sheetConfig.classHeader !== undefined ? sheetConfig.classHeader : COLUMN_HEADERS.CLASS;
    /** @type {string} 名前ヘッダー名 */
    const nameHeaderName = sheetConfig.nameHeader !== undefined ? sheetConfig.nameHeader : COLUMN_HEADERS.NAME;
    /** @type {Object} マップされたヘッダーインデックス */
    const mappedIndices = mapConfigToActualHeaders({
      opinionHeader: mainHeaderName,
      reasonHeader: reasonHeaderName,
      classHeader: classHeaderName,
      nameHeader: nameHeaderName
    }, headerIndices);

    // ユーザー情報と管理者モードの取得
    /** @type {boolean} オーナー権限フラグ */
    const isOwner = (configJson.ownerId === currentUserId);
    /** @type {string} 表示モード */
    const displayMode = configJson.displayMode || DISPLAY_MODES.ANONYMOUS;

    // 新しいデータを既存の処理パイプラインと同様に加工
    /** @type {Array} ヘッダー行の値 */
    const headers = sheet.getRange(headerRow, 1, 1, lastColumn).getValues()[0];
    /** @type {Map} 名簿マップ（未使用） */
    const rosterMap = buildRosterMap([]); // roster is not used
    /** @type {Array} 処理されたデータ */
    const processedData = rawNewData.map(function(row, idx) {
      return processRowData(row, headers, headerIndices, rosterMap, displayMode, startRowToRead + idx, isOwner);
    });

    // 取得した生データをPage.htmlが期待する形式にフォーマット
    /** @type {Array} フォーマット済みデータ */
    const formattedNewData = formatSheetDataForFrontend(processedData, mappedIndices, headerIndices, adminMode, isOwner, displayMode);

    infoLog('✅ 増分データ取得完了: %s件の新しいデータを返します', formattedNewData.length);

    return {
      header: '', // 必要に応じて設定
      sheetName: publishedSheetName,
      showCounts: false, // 必要に応じて設定
      displayMode: displayMode,
      data: formattedNewData,
      rows: formattedNewData, // 後方互換性のため
      totalCount: lastRow - headerRow, // ヘッダーを除いたデータ総数
      newCount: formattedNewData.length,
      isIncremental: true
    };
  } catch (e) {
    logError(e, 'getIncrementalData', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.DATABASE, { userId, timestamp });
    return {
      status: 'error',
      message: '増分データの取得に失敗しました: ' + e.message
    };
  }
}

/**
 * 利用可能なシート一覧を取得 (マルチテナント対応版)
 * Page.htmlから呼び出される - フロントエンド期待形式に対応
 * @param {string} requestUserId - リクエスト元のユーザーID
 */
function getAvailableSheets(requestUserId) {
  verifyUserAccess(requestUserId);
  try {
    /** @type {string} 現在のユーザーID */
    const currentUserId = requestUserId; // requestUserId を使用

    if (!currentUserId) {
      warnLog('getAvailableSheets: No current user ID set');
      return [];
    }

    /** @type {Array} シートリスト */
    const sheets = getSheetsList(currentUserId);

    if (!sheets || sheets.length === 0) {
      warnLog('getAvailableSheets: No sheets found for user:', currentUserId);
      return [];
    }

    // Page.html期待形式に変換: [{name: string, id: number}]
    return sheets.map(function(sheet) {
      return {
        name: sheet.name,
        id: sheet.id
      };
    });
  } catch (e) {
    logError(e, 'getAvailableSheets', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.DATABASE, { userId });
    return [];
  }
}

/**
 * 指定されたユーザーのスプレッドシートからシートのリストを取得します。
 * @param {string} userId - ユーザーID
 * @returns {Array<Object>} シートのリスト（例: [{name: 'Sheet1', id: 'sheetId1'}, ...]）
 */
function getSheetsList(userId) {
  try {
    /** @type {Object|null} ユーザー情報 */
    const userInfo = findUserById(userId);
    if (!userInfo || !userInfo.spreadsheetId) {
      warnLog('getSheetsList: User info or spreadsheetId not found for userId:', userId);
      return [];
    }

    /** @type {Object} スプレッドシートオブジェクト */
    const spreadsheet = openSpreadsheetOptimized(userInfo.spreadsheetId);
    /** @type {Array} シート配列 */
    const sheets = spreadsheet.getSheets();

    /** @type {Array} シートリスト */
    const sheetList = sheets.map(function(sheet) {
      return {
        name: sheet.getName(),
        id: sheet.getSheetId() // シートIDも必要に応じて取得
      };
    });

    infoLog('✅ getSheetsList: Found sheets for userId %s: %s', userId, JSON.stringify(sheetList));
    return sheetList;
  } catch (e) {
    logError(e, 'getSheetsList', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.DATABASE, { userId });
    return [];
  }
}

/**
 * ボードデータを再読み込み (マルチテナント対応版)
 * AdminPanel.htmlから呼び出される
 * @param {string} requestUserId - リクエスト元のユーザーID
 */
function refreshBoardData(requestUserId) {
  verifyUserAccess(requestUserId);
  try {
    /** @type {string} 現在のユーザーID */
    const currentUserId = requestUserId; // requestUserId を使用

    /** @type {Object|null} ユーザー情報 */
    const userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    // キャッシュをクリア
    invalidateUserCache(currentUserId, userInfo.adminEmail, userInfo.spreadsheetId, false);

    // 最新のステータスを取得
    return getAppConfig(requestUserId);
  } catch (e) {
    logError(e, 'refreshBoardData', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.SYSTEM, { userId });
    return { status: 'error', message: 'ボードデータの再読み込みに失敗しました: ' + e.message };
  }
}

/**
 * スプレッドシートの生データをフロントエンドが期待する形式にフォーマットするヘルパー関数
 * @param {Array<Object>} rawData - getSheetDataから返された生データ（originalData, reactionCountsなどを含む）
 * @param {Object} mappedIndices - 設定されたヘッダー名と実際の列インデックスのマッピング
 * @param {Object} headerIndices - 実際のヘッダー名と列インデックスのマッピング
 * @param {boolean} adminMode - 管理者モードかどうか
 * @param {boolean} isOwner - 現在のユーザーがボードのオーナーかどうか
 * @param {string} displayMode - 表示モード（'named' or 'anonymous'）
 * @returns {Array<Object>} フォーマットされたデータ
 */
function formatSheetDataForFrontend(rawData, mappedIndices, headerIndices, adminMode, isOwner, displayMode) {
  // 現在のユーザーメールを取得（リアクション状態判定用）
  /** @type {string} 現在のユーザーメールアドレス */
  const currentUserEmail = getCurrentUserEmail();

  return rawData.map(function(row, index) {
    /** @type {number|undefined} クラス列インデックス */
    const classIndex = mappedIndices.classHeader;
    /** @type {number|undefined} 意見列インデックス */
    const opinionIndex = mappedIndices.opinionHeader;
    /** @type {number|undefined} 理由列インデックス */
    const reasonIndex = mappedIndices.reasonHeader;
    /** @type {number|undefined} 名前列インデックス */
    const nameIndex = mappedIndices.nameHeader;
    /** @type {string} 名前の値 */
    let nameValue = '';
    /** @type {boolean} 名前表示フラグ */
    const shouldShowName = (adminMode === true || displayMode === DISPLAY_MODES.NAMED || isOwner);
    /** @type {boolean} 名前列インデックスの存在フラグ */
    const hasNameIndex = nameIndex !== undefined;
    /** @type {boolean} 元データ存在フラグ */
    const hasOriginalData = row.originalData && row.originalData.length > 0;

    if (shouldShowName && hasNameIndex && hasOriginalData) {
      nameValue = row.originalData[nameIndex] || '';
    }

    if (!nameValue && shouldShowName && hasOriginalData) {
      /** @type {number|undefined} メールインデックス */
      const emailIndex = headerIndices[COLUMN_HEADERS.EMAIL];
      if (emailIndex !== undefined && row.originalData[emailIndex]) {
        nameValue = row.originalData[emailIndex].split('@')[0];
      }
    }

    // リアクション状態を判定するヘルパー関数
    function checkReactionState(reactionKey) {
      /** @type {string} カラム名 */
      const columnName = COLUMN_HEADERS[reactionKey];
      /** @type {number|undefined} カラムインデックス */
      const columnIndex = headerIndices[columnName];
      /** @type {number} カウンター */
    let count = 0;
      /** @type {boolean} リアクション済みフラグ */
      let reacted = false;

      if (columnIndex !== undefined && row.originalData && row.originalData[columnIndex]) {
        /** @type {string} リアクション文字列 */
        const reactionString = row.originalData[columnIndex].toString();
        if (reactionString) {
          /** @type {Array} リアクション配列 */
          const reactions = parseReactionString(reactionString);
          count = reactions.length;
          reacted = reactions.indexOf(currentUserEmail) !== -1;
        }
      }

      // フォールバック: プリカウントされた値を使用
      if (count === 0) {
        if (reactionKey === 'UNDERSTAND') count = row.understandCount || 0;
        else if (reactionKey === 'LIKE') count = row.likeCount || 0;
        else if (reactionKey === 'CURIOUS') count = row.curiousCount || 0;
      }

      return { count: count, reacted: reacted };
    }

    // 理由列の値を取得
    /** @type {string} 理由の値 */
    let reasonValue = '';
    if (reasonIndex !== undefined && row.originalData && row.originalData[reasonIndex] !== undefined) {
      reasonValue = row.originalData[reasonIndex] || '';
    }

    // 意見と理由の取得（マッピングが利用できない場合はprocessedRowから取得）
    /** @type {string} 意見の値 */
    let opinionValue = '';
    /** @type {string} 最終理由値 */
    let finalReasonValue = reasonValue;

    if (opinionIndex !== undefined && row.originalData && row.originalData[opinionIndex]) {
      opinionValue = row.originalData[opinionIndex];
    } else if (row.opinion) {
      // フォールバック: processedRowから取得
      opinionValue = row.opinion;
    }

    // フォールバックは「理由列がマッピングできない場合」のみに限定する
    if ((reasonIndex === undefined || reasonIndex === null) && !finalReasonValue && row.reason) {
      finalReasonValue = row.reason;
    }

    return {
      rowIndex: row.rowNumber || (index + 2),
      name: nameValue,
      email: row.originalData && row.originalData[headerIndices[COLUMN_HEADERS.EMAIL]] ? row.originalData[headerIndices[COLUMN_HEADERS.EMAIL]] : (row.email || ''),
      class: (classIndex !== undefined && row.originalData && row.originalData[classIndex]) ? row.originalData[classIndex] : '',
      opinion: opinionValue,
      reason: finalReasonValue,
      reactions: {
        UNDERSTAND: checkReactionState('UNDERSTAND'),
        LIKE: checkReactionState('LIKE'),
        CURIOUS: checkReactionState('CURIOUS')
      },
      highlight: row.isHighlighted || false
    };
  });
}

/**
 * アプリ設定を取得（最適化版） (マルチテナント対応版)
 * @param {string} requestUserId - リクエスト元のユーザーID
 */
function getAppConfig(requestUserId) {
  verifyUserAccess(requestUserId);
  try {
    /** @type {string} 現在のユーザーID */
    const currentUserId = requestUserId;
    /** @type {Object|null} ユーザー情報 */
    const userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    /** @type {Object} 設定JSONオブジェクト */
    let configJson = {};
    try {
      configJson = JSON.parse(userInfo.configJson || '{}');
    } catch (parseError) {
      warnLog('ConfigJson parse error in getAppConfig:', parseError.message);
      configJson = {};
    }

    // --- 統一された自動修復システム ---
    const healingResult = performAutoHealing(userInfo, configJson, currentUserId);
    if (healingResult.updated) {
      configJson = healingResult.configJson;
    }

    /** @type {Array} シートリスト */
    const sheets = getSheetsList(currentUserId);
    /** @type {Object} アプリケーションURL群 */
    const appUrls = generateUserUrls(currentUserId);

    // 回答数を取得
    /** @type {number} 回答数 */
    let answerCount = 0;
    /** @type {number} 総リアクション数 */
    let totalReactions = 0;
    try {
      if (configJson.publishedSpreadsheetId && configJson.publishedSheetName) {
        /** @type {Object} レスポンスデータ */
        const responseData = getResponsesData(currentUserId, configJson.publishedSheetName);
        if (responseData.status === 'success') {
          answerCount = responseData.data.length;
          // リアクション数の概算計算（詳細実装は後回し）
          totalReactions = answerCount * 2; // 暫定値
        }
      }
    } catch (countError) {
      warnLog('回答数の取得に失敗: ' + countError.message);
    }

    // アクティブシートのフォームURLを優先（未設定時のみグローバルにフォールバック）
    let activeFormUrl = '';
    try {
      const activeSheet = (configJson && typeof configJson === 'object') ? (configJson.publishedSheetName || '') : '';
      if (activeSheet) {
        const sheetCfg = configJson['sheet_' + activeSheet];
        if (sheetCfg && typeof sheetCfg === 'object' && sheetCfg.formUrl) {
          activeFormUrl = sheetCfg.formUrl;
        }
      }
    } catch (e) {
      // フォームURLの解決失敗は致命的ではない
    }

    return {
      status: 'success',
      userId: currentUserId,
      adminEmail: userInfo.adminEmail,
      publishedSpreadsheetId: configJson.publishedSpreadsheetId || '',
      publishedSheetName: configJson.publishedSheetName || '',
      displayMode: configJson.displayMode || DISPLAY_MODES.ANONYMOUS,
      isPublished: configJson.appPublished || false,
      appPublished: configJson.appPublished || false, // AdminPanel.htmlで使用される
      availableSheets: sheets,
      allSheets: sheets, // AdminPanel.htmlで使用される
      spreadsheetUrl: userInfo.spreadsheetUrl,
      formUrl: activeFormUrl || configJson.formUrl || '',
      editFormUrl: configJson.editFormUrl || '',
      webAppUrl: appUrls.webAppUrl,
      adminUrl: appUrls.adminUrl,
      viewUrl: appUrls.viewUrl,
      activeSheetName: configJson.publishedSheetName || '',
      appUrls: appUrls,
      // AdminPanel.htmlが期待する表示設定プロパティ
      showNames: configJson.showNames || false,
      showCounts: configJson.showCounts === true,
      // データベース詳細情報
      userInfo: {
        userId: currentUserId,
        adminEmail: userInfo.adminEmail,
        spreadsheetId: userInfo.spreadsheetId || '',
        spreadsheetUrl: userInfo.spreadsheetUrl || '',
        createdAt: userInfo.createdAt || '',
        lastAccessedAt: userInfo.lastAccessedAt || '',
        isActive: userInfo.isActive || 'false',
        configJson: userInfo.configJson || '{}'
      },
      // 統計情報
      answerCount: answerCount,
      totalReactions: totalReactions,
      // システム状態
      systemStatus: {
        setupStatus: configJson.setupStatus || 'unknown',
        formCreated: configJson.formCreated || false,
        appPublished: configJson.appPublished || false,
        lastUpdated: new Date().toISOString()
      },
      // ユーザーのセットアップ段階を判定（統一化されたロジック）
      setupStep: determineSetupStepUnified(userInfo, configJson)
    };
  } catch (e) {
    logError(e, 'getAppSettings', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.SYSTEM);
    return {
      status: 'error',
      message: '設定の取得に失敗しました: ' + e.message
    };
  }
}

/**
 * シート設定を保存する（統合版：Batch/Optimized機能統合）
 * AdminPanel.htmlから呼び出される
 * @param {string} userId - ユーザーID
 * @param {string} spreadsheetId - 設定対象のスプレッドシートID
 * @param {string} sheetName - 設定対象のシート名
 * @param {object} config - 保存するシート固有の設定
 * @param {object} options - オプション設定
 * @param {object} options.sheetsService - 共有SheetsService（最適化用）
 * @param {object} options.userInfo - 事前取得済みユーザー情報（最適化用）
 * @param {boolean} options.batchMode - バッチモード（表示オプション同時更新）
 * @param {object} options.displayOptions - バッチモード時の表示オプション
 */
function saveSheetConfig(userId, spreadsheetId, sheetName, config, options = {}) {
  try {
    if (!spreadsheetId || typeof spreadsheetId !== 'string') {
      throw new Error('無効なspreadsheetIdです: ' + spreadsheetId);
    }
    if (!sheetName || typeof sheetName !== 'string') {
      throw new Error('無効なsheetNameです: ' + sheetName);
    }
    if (!config || typeof config !== 'object') {
      throw new Error('無効なconfigです: ' + config);
    }

    /** @type {string} 現在のユーザーID */
    const currentUserId = userId;

    // 最適化モード: 事前取得済みuserInfoを使用、なければ取得
    /** @type {Object} ユーザー情報 */
    const userInfo = options.userInfo || findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    /** @type {Object} 設定JSONオブジェクト */
    let configJson = {};
    try {
      configJson = JSON.parse(userInfo.configJson || '{}');
    } catch (parseError) {
      warnLog('ConfigJson parse error in saveSheetConfig:', parseError.message);
      configJson = {};
    }

    // シート設定を更新
    /** @type {string} シートキー */
    const sheetKey = 'sheet_' + sheetName;
    configJson[sheetKey] = {
      ...config,
      lastModified: new Date().toISOString()
    };

    // バッチモード: 表示オプションも同時更新
    if (options.batchMode && options.displayOptions) {
      configJson.publishedSheetName = sheetName;
      configJson.publishedSpreadsheetId = spreadsheetId;
      configJson.displayMode = options.displayOptions.showNames ? 'named' : 'anonymous';
      configJson.showCounts = options.displayOptions.showCounts;
      configJson.appPublished = true;
      configJson.lastModified = new Date().toISOString();
    }

    updateUser(currentUserId, { configJson: JSON.stringify(configJson) });
    infoLog('✅ シート設定を保存しました: %s', sheetKey);
    return { status: 'success', message: 'シート設定を保存しました。' };
  } catch (e) {
    logDatabaseError(e, 'saveSheetSettings', { userId });
    return { status: 'error', message: 'シート設定の保存に失敗しました: ' + e.message };
  }
}

/**
 * 表示するシートを切り替える（統合版：Optimized機能統合）
 * AdminPanel.htmlから呼び出される
 * @param {string} userId - ユーザーID
 * @param {string} spreadsheetId - 公開対象のスプレッドシートID
 * @param {string} sheetName - 公開対象のシート名
 * @param {object} options - オプション設定
 * @param {object} options.sheetsService - 共有SheetsService（最適化用）
 * @param {object} options.userInfo - 事前取得済みユーザー情報（最適化用）
 */
function switchToSheet(userId, spreadsheetId, sheetName, options = {}) {
  try {
    if (!spreadsheetId || typeof spreadsheetId !== 'string') {
      throw new Error('無効なspreadsheetIdです: ' + spreadsheetId);
    }
    if (!sheetName || typeof sheetName !== 'string') {
      throw new Error('無効なsheetNameです: ' + sheetName);
    }

    /** @type {string} 現在のユーザーID */
    const currentUserId = userId;

    // 最適化モード: 事前取得済みuserInfoを使用、なければ取得
    /** @type {Object} ユーザー情報 */
    const userInfo = options.userInfo || findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    /** @type {Object} 設定JSONオブジェクト */
    let configJson = {};
    try {
      configJson = JSON.parse(userInfo.configJson || '{}');
    } catch (parseError) {
      warnLog('ConfigJson parse error in switchToSheet:', parseError.message);
      configJson = {};
    }

    configJson.publishedSpreadsheetId = spreadsheetId;
    configJson.publishedSheetName = sheetName;
    configJson.appPublished = true; // シートを切り替えたら公開状態にする
    configJson.lastModified = new Date().toISOString();

    updateUser(currentUserId, { configJson: JSON.stringify(configJson) });
    infoLog('✅ 表示シートを切り替えました: %s - %s', spreadsheetId, sheetName);
    return { status: 'success', message: '表示シートを切り替えました。' };
  } catch (e) {
    logError(e, 'switchSheet', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.SYSTEM, { userId, sheetName });
    return { status: 'error', message: '表示シートの切り替えに失敗しました: ' + e.message };
  }
}

// =================================================================
// ヘルパー関数
// =================================================================

// include 関数は main.gs で定義されています
function getResponsesData(userId, sheetName) {
  /** @type {Object} ユーザー情報 */
  const userInfo = getOrFetchUserInfo(userId, 'userId', {
    useExecutionCache: true,
    ttl: 300
  });
  if (!userInfo) {
    return { status: 'error', message: 'ユーザー情報が見つかりません' };
  }

  try {
    /** @type {Object} Google Sheets APIサービス */
    const service = getSheetsServiceCached();
    /** @type {string} スプレッドシートID */
    const spreadsheetId = userInfo.spreadsheetId;
    /** @type {string} 範囲文字列 */
    const range = "'" + (sheetName || 'フォームの回答 1') + "'!A:Z";

    /** @type {Object} API応答データ */
    const response = batchGetSheetsData(service, spreadsheetId, [range]);
    /** @type {Array} スプレッドシート値の配列 */
    const values = response.valueRanges[0].values || [];

    if (values.length === 0) {
      return { status: 'success', data: [], headers: [] };
    }

    return {
      status: 'success',
      data: values.slice(1),
      headers: values[0]
    };
  } catch (e) {
    logError(e, 'getAnswerData', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.DATABASE, { userId });
    return { status: 'error', message: '回答データの取得に失敗しました: ' + e.message };
  }
}

// =================================================================
// HTML依存関数（UI連携）
// =================================================================

/**
 * 現在のユーザーのステータスを取得し、UIに表示するための情報を返します。(統一スキーマ対応版)
 * @param {string} [requestUserId] - リクエスト元のユーザーID (オプション、未指定時は自動取得)
 * @returns {object} 統一スキーマのユーザーステータス情報
 */
function getCurrentUserStatus(requestUserId = null) {
  try {
    const activeUserEmail = getCurrentUserEmail();

    // requestUserIdが未指定または無効な場合は、自動取得またはメールアドレスで検索
    let userInfo;
    if (requestUserId && requestUserId.trim() !== '') {
      verifyUserAccess(requestUserId);
      userInfo = findUserById(requestUserId);
    } else {
      // 自動取得を試行
      const autoUserId = getUserId();
      if (autoUserId) {
        userInfo = findUserById(autoUserId);
      } else {
        userInfo = findUserByEmail(activeUserEmail);
      }
    }

    if (!userInfo) {
      return { 
        status: 'error', 
        message: 'ユーザー情報が見つかりません。',
        data: null,
        userInfo: null,
        timestamp: new Date().toISOString()
      };
    }

    // 統一スキーマでの戻り値
    return {
      status: 'success',
      data: {
        userId: userInfo.userId,
        adminEmail: userInfo.adminEmail,
        isActive: userInfo.isActive,
        lastAccessedAt: userInfo.lastAccessedAt
      },
      userInfo: {
        userId: userInfo.userId,
        adminEmail: userInfo.adminEmail,
        isActive: userInfo.isActive,
        lastAccessedAt: userInfo.lastAccessedAt
      },
      message: null,
      timestamp: new Date().toISOString()
    };
  } catch (e) {
    logError(e, 'getCurrentUserStatus', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.SYSTEM, { requestUserId });
    return { 
      status: 'error', 
      message: 'ステータス取得に失敗しました: ' + e.message,
      data: null,
      userInfo: null,
      timestamp: new Date().toISOString()
    };
  }
}

/**
 * アクティブなフォーム情報を取得 (マルチテナント対応版)
 * Page.htmlから呼び出される（パラメータなし）
 * @param {string} requestUserId - リクエスト元のユーザーID
 */
function getActiveFormInfo(requestUserId) {
  verifyUserAccess(requestUserId);
  try {
    /** @type {string} 現在のユーザーID */
    const currentUserId = requestUserId; // requestUserId を使用

    /** @type {Object|null} ユーザー情報 */
    const userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    /** @type {Object} 設定JSONオブジェクト */
    let configJson = {};
    try {
      configJson = JSON.parse(userInfo.configJson || '{}');
    } catch (parseError) {
      warnLog('ConfigJson parse error in getActiveFormInfo:', parseError.message);
      configJson = {};
    }

    // フォーム回答数を取得
    /** @type {number} 回答数 */
    let answerCount = 0;
    try {
      if (configJson.publishedSpreadsheetId && configJson.publishedSheet) {
        /** @type {Object} レスポンスデータ */
        const responseData = getResponsesData(currentUserId, configJson.publishedSheet);
        if (responseData.status === 'success') {
          answerCount = responseData.data.length;
        }
      }
    } catch (countError) {
      warnLog('回答数の取得に失敗: ' + countError.message);
    }

    return {
      status: 'success',
      formUrl: configJson.formUrl || '',
      editFormUrl: configJson.editFormUrl || '',
      editUrl: configJson.editFormUrl || '',  // AdminPanel.htmlが期待するフィールド名
      formId: extractFormIdFromUrl(configJson.formUrl || configJson.editFormUrl || ''),
      spreadsheetUrl: userInfo.spreadsheetUrl || '',
      answerCount: answerCount,
      isFormActive: !!(configJson.formUrl && configJson.formCreated)
    };
  } catch (e) {
    logError(e, 'getActiveFormInfo', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.SYSTEM, { userId });
    return { status: 'error', message: 'フォーム情報の取得に失敗しました: ' + e.message };
  }
}

/**
 * ユーザーが管理者かどうかをチェックする (マルチテナント対応版)
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @returns {boolean} 管理者の場合はtrue、そうでない場合はfalse
 */
function checkAdmin(requestUserId) {
  verifyUserAccess(requestUserId);
  try {
    const userInfo = findUserById(requestUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    // getCurrentUserEmail() が requestUserId の adminEmail と一致するかどうかを verifyUserAccess で既にチェック済み
    // ここでは単に userInfo.adminEmail と getCurrentUserEmail() が一致するかを返す
    return getCurrentUserEmail() === userInfo.adminEmail;
  } catch (e) {
    errorLog('checkAdmin エラー: ' + e.message);
    return false;
  }
}

/**
 * 指定シートのデータ行数を取得します。
 * @param {string} spreadsheetId スプレッドシートID
 * @param {string} sheetName シート名
 * @param {string} classFilter クラスフィルタ
 * @returns {number} データ行数
 */
function countSheetRows(spreadsheetId, sheetName, classFilter) {
  const key = 'rowCount_' + spreadsheetId + '_' + sheetName + '_' + classFilter;
  return cacheManager.get(key, () => {
    const sheet = openSpreadsheetOptimized(spreadsheetId).getSheetByName(sheetName);
    if (!sheet) return 0;

    const lastRow = sheet.getLastRow();
    if (!classFilter || classFilter === 'すべて') {
      return Math.max(0, lastRow - 1);
    }

    const headerIndices = getHeaderIndices(spreadsheetId, sheetName);
    const classIndex = headerIndices[COLUMN_HEADERS.CLASS];
    if (classIndex === undefined) {
      return Math.max(0, lastRow - 1);
    }

    const values = sheet.getRange(2, classIndex + 1, lastRow - 1, 1).getValues();
    return values.reduce((cnt, row) => row[0] === classFilter ? cnt + 1 : cnt, 0);
  }, { ttl: 30, enableMemoization: true });
}

/**
 * データ数を取得する (マルチテナント対応版)
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @returns {number} データ数
 */
function getDataCount(requestUserId, classFilter, sortOrder, adminMode) {
  verifyUserAccess(requestUserId);
  try {
    const userInfo = findUserById(requestUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    let configJson = {};
    try {
      configJson = JSON.parse(userInfo.configJson || '{}');
    } catch (parseError) {
      warnLog('ConfigJson parse error in getUserIdByEmailAddress:', parseError.message);
      configJson = {};
    }

    if (!configJson.publishedSpreadsheetId || !configJson.publishedSheetName) {
      return {
        count: 0,
        lastUpdate: new Date().toISOString(),
        status: 'error',
        message: 'スプレッドシート設定なし'
      };
    }

    const count = countSheetRows(
      configJson.publishedSpreadsheetId,
      configJson.publishedSheetName,
      classFilter
    );
    return {
      count,
      lastUpdate: new Date().toISOString(),
      status: 'success'
    };
  } catch (e) {
    errorLog('getDataCount エラー: ' + e.message);
    return {
      count: 0,
      lastUpdate: new Date().toISOString(),
      status: 'error',
      message: e.message
    };
  }
}

/**
 * フォーム設定を更新 (マルチテナント対応版)
 * AdminPanel.htmlから呼び出される
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @param {string} title - 新しいタイトル
 * @param {string} description - 新しい説明
 */
function updateFormSettings(requestUserId, title, description) {
  verifyUserAccess(requestUserId);
  try {
    /** @type {string} 現在のユーザーID */
    const currentUserId = requestUserId; // requestUserId を使用

    /** @type {Object|null} ユーザー情報 */
    const userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    /** @type {Object} 設定JSONオブジェクト */
    let configJson = {};
    try {
      configJson = JSON.parse(userInfo.configJson || '{}');
    } catch (parseError) {
      warnLog('ConfigJson parse error in updateFormTitle:', parseError.message);
      configJson = {};
    }

    if (configJson.editFormUrl) {
      try {
        /** @type {string} フォームID */
        const formId = extractFormIdFromUrl(configJson.editFormUrl);
        /** @type {Object} Google Form オブジェクト */
        const form = FormApp.openById(formId);

        if (title) {
          form.setTitle(title);
        }
        if (description) {
          form.setDescription(description);
        }

        return {
          status: 'success',
          message: 'フォーム設定が更新されました'
        };
      } catch (formError) {
        errorLog('フォーム更新エラー: ' + formError.message);
        return { status: 'error', message: 'フォームの更新に失敗しました: ' + formError.message };
      }
    } else {
      return { status: 'error', message: 'フォームが見つかりません' };
    }
  } catch (e) {
    errorLog('updateFormSettings エラー: ' + e.message);
    return { status: 'error', message: 'フォーム設定の更新に失敗しました: ' + e.message };
  }
}

/**
 * システム設定を保存 (マルチテナント対応版)
 * AdminPanel.htmlから呼び出される
 */
function saveSystemConfig(requestUserId, config) {
  verifyUserAccess(requestUserId);
  try {
    /** @type {string} 現在のユーザーID */
    const currentUserId = requestUserId; // requestUserId を使用

    /** @type {Object|null} ユーザー情報 */
    const userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    /** @type {Object} 設定JSONオブジェクト */
    let configJson = {};
    try {
      configJson = JSON.parse(userInfo.configJson || '{}');
    } catch (parseError) {
      warnLog('ConfigJson parse error in saveSystemConfig:', parseError.message);
      configJson = {};
    }

    // システム設定を更新
    configJson.systemConfig = {
      cacheEnabled: config.cacheEnabled,
      autosaveInterval: config.autosaveInterval,
      logLevel: config.logLevel,
      updatedAt: new Date().toISOString()
    };

    updateUser(currentUserId, {
      configJson: JSON.stringify(configJson)
    });

    return {
      status: 'success',
      message: 'システム設定が保存されました'
    };
  } catch (e) {
    errorLog('saveSystemConfig エラー: ' + e.message);
    return { status: 'error', message: 'システム設定の保存に失敗しました: ' + e.message };
  }
}

/**
 * ハイライト状態の切り替え (マルチテナント対応版)
 * Page.htmlから呼び出される - フロントエンド期待形式に対応
 * @param {string} requestUserId - リクエスト元のユーザーID
 */
function toggleHighlight(requestUserId, rowIndex, sheetName) {
  verifyUserAccess(requestUserId);
  try {
    /** @type {string} 現在のユーザーID */
    const currentUserId = requestUserId; // requestUserId を使用

    /** @type {Object|null} ユーザー情報 */
    const userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    // 管理者権限チェック - 統一されたcheckAdmin関数を使用
    if (!checkAdmin(requestUserId)) {
      throw new Error('ハイライト機能は管理者のみ使用できます');
    }

    /** @type {Object} 処理結果 */
    const result = processHighlightToggle(
      userInfo.spreadsheetId,
      sheetName || 'フォームの回答 1',
      rowIndex
    );

    // Page.html期待形式に変換
    if (result && result.status === 'success') {
      return {
        status: "ok",
        highlight: result.highlighted || false
      };
    } else {
      throw new Error(result.message || 'ハイライト切り替えに失敗しました');
    }
  } catch (e) {
    errorLog('toggleHighlight エラー: ' + e.message);
    return {
      status: "error",
      message: e.message
    };
  }
}

/**
 * 利用可能なシート一覧を取得
 * Page.htmlから呼び出される - フロントエンド期待形式に対応
 */

/**
 * クイックスタートセットアップ（完全版） (マルチテナント対応版)
 * フォルダ作成、フォーム作成、スプレッドシート作成、ボード公開まで一括実行
 * @param {string} requestUserId - リクエスト元のユーザーID
 */
/**
 * クイックスタート用のファイル作成とフォルダ管理
 * @param {object} setupContext - セットアップコンテキスト
 * @returns {object} 作成されたファイル情報
 */
function createQuickStartFiles(setupContext) {
  /** @type {string} セットアップコンテキストからのユーザーメール */
  const userEmail = setupContext.userEmail;
  /** @type {string} セットアップコンテキストからのリクエストユーザーID */
  const requestUserId = setupContext.requestUserId;

  // ステップ1: ユーザー専用フォルダを作成
  /** @type {Object} ユーザーフォルダ */
    const folder = createUserFolder(userEmail);

  // ステップ2: Googleフォームとスプレッドシートを作成
  /** @type {Object} フォーム・スプレッドシート作成情報 */
    const formAndSsInfo = createUnifiedForm('study', userEmail, requestUserId);

  // 作成したファイルをフォルダに移動（改善版：冗長処理除去と安全な移動処理）
  if (folder) {
    /** @type {Object} 移動結果 */
    const moveResults = { form: false, spreadsheet: false };
    /** @type {Array} 移動エラー */
    const moveErrors = [];

    try {
      /** @type {Object} フォームファイル */
      const formFile = DriveApp.getFileById(formAndSsInfo.formId);
      /** @type {Object} スプレッドシートファイル */
      const ssFile = DriveApp.getFileById(formAndSsInfo.spreadsheetId);

      // フォームファイルの移動処理
      try {
        // 既にフォルダに存在するかチェック（重複移動を防止）
        /** @type {Object} フォーム親フォルダー */
        const formParents = formFile.getParents();
        /** @type {boolean} フォームが既にフォルダーに存在するか */
        let isFormAlreadyInFolder = false;

        while (formParents.hasNext()) {
          if (formParents.next().getId() === folder.getId()) {
            isFormAlreadyInFolder = true;
            break;
          }
        }

        if (!isFormAlreadyInFolder) {
          
          // 推奨メソッドmoveTo()を使用してファイル移動
          formFile.moveTo(folder);
          moveResults.form = true;
          infoLog('✅ フォームファイル移動完了: %s (ID: %s)', folder.getName(), formFile.getId());
        } else {
          moveResults.form = true;
        }
      } catch (formMoveError) {
        moveErrors.push('フォームファイル移動エラー: ' + formMoveError.message);
        errorLog('❌ フォームファイルの移動に失敗:', formMoveError.message);
        errorLog('フォームID: %s, フォルダID: %s', formAndSsInfo.formId, folder ? folder.getId() : 'なし');
      }

      // スプレッドシートファイルの移動処理
      try {
        // 既にフォルダに存在するかチェック（重複移動を防止）
        /** @type {Object} スプレッドシート親フォルダー */
        const ssParents = ssFile.getParents();
        /** @type {boolean} スプレッドシートが既にフォルダーに存在するか */
        let isSsAlreadyInFolder = false;

        while (ssParents.hasNext()) {
          if (ssParents.next().getId() === folder.getId()) {
            isSsAlreadyInFolder = true;
            break;
          }
        }

        if (!isSsAlreadyInFolder) {
          
          // 推奨メソッドmoveTo()を使用してファイル移動
          ssFile.moveTo(folder);
          moveResults.spreadsheet = true;
          infoLog('✅ スプレッドシートファイル移動完了: %s (ID: %s)', folder.getName(), ssFile.getId());
        } else {
          moveResults.spreadsheet = true;
        }
      } catch (ssMoveError) {
        moveErrors.push('スプレッドシートファイル移動エラー: ' + ssMoveError.message);
        errorLog('❌ スプレッドシートファイルの移動に失敗:', ssMoveError.message);
        errorLog('スプレッドシートID: %s, フォルダID: %s', formAndSsInfo.spreadsheetId, folder ? folder.getId() : 'なし');
      }

      // 移動結果のログ出力
      if (moveResults.form && moveResults.spreadsheet) {
        infoLog('✅ 全ファイルのフォルダ移動が完了: ' + folder.getName());
      } else {
        warnLog('⚠️ 一部のファイル移動に失敗しましたが、処理を継続します');
        if (moveErrors.length > 0) {
        }
      }

    } catch (generalError) {
      errorLog('❌ ファイル移動処理で予期しないエラー:', generalError.message);
      // ファイル移動失敗は致命的ではないため、処理は継続
    }
  }

  return {
    folder: folder,
    formAndSsInfo: formAndSsInfo,
    moveResults: moveResults || { form: false, spreadsheet: false }
  };
}

/**
 * クイックスタートのデータベース更新とキャッシュ管理
 * @param {object} setupContext - セットアップコンテキスト
 * @param {object} createdFiles - 作成されたファイル情報
 * @returns {object} 更新された設定オブジェクト
 */
function updateQuickStartDatabase(setupContext, createdFiles) {
  /** @type {string} セットアップコンテキストからのリクエストユーザーID */
  const requestUserId = setupContext.requestUserId;
  /** @type {Object} セットアップコンテキストからの設定JSON */
  const configJson = setupContext.configJson;
  /** @type {string} セットアップコンテキストからのユーザーメール */
  const userEmail = setupContext.userEmail;
  /** @type {Object} フォーム・スプレッドシート情報 */
  const formAndSsInfo = createdFiles.formAndSsInfo;
  /** @type {Object} 作成されたフォルダ */
  const folder = createdFiles.folder;
  // クイックスタート用の適切な初期設定を作成（guessedConfig形式）
  /** @type {string} シート設定キー */
  const sheetConfigKey = 'sheet_' + formAndSsInfo.sheetName;
  /** @type {Object} クイックスタートシート設定 */
  const quickStartSheetConfig = {
    // 実際の設定値（フロントエンドで使用される）
    opinionHeader: 'あなたの考えや気づいたことを教えてください',
    reasonHeader: 'そう考える理由や体験があれば教えてください（任意）',
    nameHeader: '名前',
    classHeader: 'クラス',
    timestampHeader: 'タイムスタンプ',
    guessedConfig: {
      opinionHeader: 'あなたの考えや気づいたことを教えてください',
      reasonHeader: 'そう考える理由や体験があれば教えてください（任意）',
      nameHeader: '名前',
      classHeader: 'クラス',
      timestampHeader: 'タイムスタンプ',
      // 基本的な列マッピング（QuickStart用デフォルト）
      opinionColumn: 1, // B列（0ベースなので1）
      nameColumn: 2,    // C列
      reasonColumn: 3,  // D列  
      classColumn: 4,   // E列
      timestampColumn: 0 // A列（タイムスタンプ）
    },
    formUrl: formAndSsInfo.viewFormUrl || formAndSsInfo.formUrl, // シート固有のフォームURL保存
    editFormUrl: formAndSsInfo.editFormUrl, // 編集用URL保存
    lastModified: new Date().toISOString(),
    flowType: 'quickstart' // 作成フロー識別用
  };
  // 型安全性確保: publishedSheetNameの明示的文字列変換
  /** @type {string} 安全なシート名 */
  let safeSheetName = formAndSsInfo.sheetName;
  if (typeof safeSheetName !== 'string') {
    errorLog('❌ quickStartSetup: formAndSsInfo.sheetNameが文字列ではありません:', typeof safeSheetName, safeSheetName);
    safeSheetName = String(safeSheetName); // 強制的に文字列化
  }
  if (safeSheetName === 'true' || safeSheetName === 'false') {
    errorLog('❌ quickStartSetup: 無効なシート名が検出されました:', safeSheetName);
    throw new Error('無効なシート名: ' + safeSheetName);
  }

  // 6時間自動停止機能の設定
  /** @type {string} 公開日時 */
  const publishedAt = new Date().toISOString();
  /** @type {number} 自動停止分数 */
  const autoStopMinutes = 360; // 6時間 = 360分
  /** @type {string} スケジュール終了日時 */
  const scheduledEndAt = new Date(Date.now() + (autoStopMinutes * 60 * 1000)).toISOString();

  /** @type {Object} 更新された設定 */
  const updatedConfig = {
    ...configJson,
    setupStatus: 'completed',
    formCreated: true,
    formUrl: formAndSsInfo.viewFormUrl || formAndSsInfo.formUrl,
    editFormUrl: formAndSsInfo.editFormUrl,
    publishedSpreadsheetId: formAndSsInfo.spreadsheetId,
    publishedSheetName: safeSheetName, // 型安全性が確保されたシート名
    appPublished: true,
    folderId: folder ? folder.getId() : '',
    folderUrl: folder ? folder.getUrl() : '',
    completedAt: new Date().toISOString(),
    // 6時間自動停止機能の設定
    publishedAt: publishedAt, // 公開開始時間
    autoStopEnabled: true, // 6時間自動停止フラグ
    autoStopMinutes: autoStopMinutes, // 6時間 = 360分
    scheduledEndAt: scheduledEndAt, // 予定終了日時
    lastPublishedAt: publishedAt, // 最後の公開日時
    totalPublishCount: (configJson.totalPublishCount || 0) + 1, // 累計公開回数
    autoStoppedAt: null, // 自動停止実行日時をリセット
    autoStopReason: null, // 自動停止理由をリセット
    [sheetConfigKey]: quickStartSheetConfig
  };

  // ユーザーデータベースを新しいセットアップ情報で完全に更新
  /** @type {Object} 更新データ */
  const updateData = {
    spreadsheetId: formAndSsInfo.spreadsheetId,
    spreadsheetUrl: formAndSsInfo.spreadsheetUrl,
    configJson: JSON.stringify(updatedConfig),
    lastAccessedAt: new Date().toISOString()
  };
  updateUser(requestUserId, updateData);

  infoLog('✅ ユーザーデータベース更新完了!');

  // 重要: 新しいセットアップ完了後に包括的キャッシュ同期を実行（二重保証）

  // updateUserで既にキャッシュ同期されているが、クイックスタートの場合は追加で確実性を高める
  synchronizeCacheAfterCriticalUpdate(requestUserId, userEmail, null, formAndSsInfo.spreadsheetId);

  // 最終検証: 更新が正確に反映されているかデータベースから直接確認

  // 少し待ってからデータベースを確認（更新の反映を待つ）
  Utilities.sleep(500);

  /** @type {Object} 検証ユーザー情報 */
  const verificationUserInfo = findUserByIdFresh(requestUserId);

  if (verificationUserInfo && verificationUserInfo.spreadsheetId === formAndSsInfo.spreadsheetId) {
    infoLog('✅ データベース更新検証成功: 新しいスプレッドシートID確認');
  } else {
    warnLog('⚠️ データベース更新検証で不一致が検出されましたが、処理を続行します');

    // 検証失敗でもQuick Startを続行する（フォーム作成は成功している）
  }

  return updatedConfig;
}

/**
 * QuickStart専用自動公開処理（堅牢性向上版）
 * @param {string} requestUserId - ユーザーID
 * @param {string} sheetName - 公開するシート名
 * @returns {object} 公開結果
 */
function performAutoPublish(requestUserId, sheetName) {
  try {
    
    // 入力値検証
    if (!requestUserId || typeof requestUserId !== 'string') {
      throw new Error('有効なユーザーIDが指定されていません');
    }
    
    if (!sheetName || typeof sheetName !== 'string' || sheetName.trim() === '') {
      throw new Error('有効なシート名が指定されていません');
    }

    const trimmedSheetName = sheetName.trim();
    
    // ユーザー情報の事前確認
    const userInfo = findUserById(requestUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    
    // スプレッドシートの存在確認
    if (!userInfo.spreadsheetId) {
      throw new Error('スプレッドシートIDが見つかりません');
    }

    // setActiveSheetを呼び出して自動公開を実行
    const publishResult = setActiveSheet(requestUserId, trimmedSheetName);
    
    // 公開結果の詳細検証
    if (publishResult && publishResult.success !== false) {
      infoLog('✅ QuickStart自動公開完了', {
        requestUserId,
        sheetName: trimmedSheetName,
        result: publishResult
      });
      
      return {
        success: true,
        published: true,
        sheetName: trimmedSheetName,
        message: '回答ボードが自動的に公開されました',
        details: publishResult,
        publishedAt: new Date().toISOString()
      };
    } else {
      throw new Error('setActiveSheetが失敗しました: ' + (publishResult?.message || 'unknown error'));
    }
    
  } catch (error) {
    errorLog('❌ QuickStart自動公開エラー', {
      requestUserId,
      sheetName,
      error: error.message,
      stack: error.stack
    });
    
    // 自動公開が失敗しても全体のQuickStartは成功とする（堅牢性）
    return {
      success: false,
      published: false,
      sheetName: sheetName,
      message: '自動公開に失敗しましたが、手動で公開できます',
      error: error.message,
      failedAt: new Date().toISOString(),
      // 手動公開用の情報を提供
      manualPublishInstructions: '管理パネルのシート選択から手動で公開してください'
    };
  }
}

/**
 * クイックスタート完了時のレスポンス生成
 * @param {object} setupContext - セットアップコンテキスト
 * @param {object} createdFiles - 作成されたファイル情報
 * @param {object} updatedConfig - 更新された設定
 * @param {object} publishResult - 公開結果（オプション）
 * @returns {object} 成功レスポンス
 */
function generateQuickStartResponse(setupContext, createdFiles, updatedConfig, publishResult) {
  /** @type {string} セットアップコンテキストからのリクエストユーザーID */
  const requestUserId = setupContext.requestUserId;
  /** @type {Object} フォーム・スプレッドシート情報 */
  const formAndSsInfo = createdFiles.formAndSsInfo;

  // 最終検証：新規作成されたファイルの確認

  // 公開結果の検証
  /** @type {boolean} 公開状態 */
  const isPublished = publishResult && publishResult.success && publishResult.published;
  /** @type {string} 公開メッセージ */
  const publishMessage = isPublished 
    ? '回答ボードが自動的に公開されました！' 
    : '回答ボードが作成されました。管理パネルから手動で公開してください。';

  infoLog('✅ クイックスタートセットアップ完了: ' + requestUserId);

  /** @type {Object} アプリケーションURL */
  const appUrls = generateUserUrls(requestUserId);
  
  // シート設定データの取得
  /** @type {string} シート設定キー */
  const sheetConfigKey = 'sheet_' + formAndSsInfo.sheetName;
  /** @type {Object} シート設定 */
  const sheetConfig = updatedConfig[sheetConfigKey] || {};
  
  // 拡張されたレスポンス情報
  // 統一スキーマでの戻り値構造
  /** @type {Object} レスポンスオブジェクト */
  const response = {
    status: 'success',
    message: 'クイックスタートが完了しました！' + publishMessage,
    data: {
      webAppUrl: appUrls.webAppUrl,
      adminUrl: appUrls.adminUrl,
      viewUrl: appUrls.viewUrl,
      setupUrl: appUrls.setupUrl,
      formUrl: updatedConfig.formUrl,
      editFormUrl: updatedConfig.editFormUrl,
      spreadsheetUrl: formAndSsInfo.spreadsheetUrl,
      folderUrl: updatedConfig.folderUrl,
      setupComplete: true,
      autoPublished: isPublished,
      publishResult: publishResult,
      sheetName: formAndSsInfo.sheetName,
      formId: formAndSsInfo.formId,
      spreadsheetId: formAndSsInfo.spreadsheetId,
      displayMode: updatedConfig.displayMode || 'anonymous',
      showCounts: updatedConfig.showCounts === true,
      config: sheetConfig.guessedConfig || {},
      publishedSheetName: formAndSsInfo.sheetName,
      completedSteps: [
        'ユーザー専用フォルダの作成',
        'Googleフォームとスプレッドシートの作成',
        'データベース更新とキャッシュ管理',
        isPublished ? '回答ボードの自動公開' : '回答ボード作成（手動公開待ち）',
        'キャッシュクリアと最終化',
        'セットアップ完了'
      ]
    },
    userInfo: null, // クイックスタート時は不要だが統一スキーマのため含める
    timestamp: new Date().toISOString()
  };
  
  return response;
}

/**
 * クイックスタートセットアップのコンテキストを初期化
 * @param {string} requestUserId - ユーザーID
 * @returns {object} セットアップコンテキスト
 */
function initializeQuickStartContext(requestUserId) {

  // ユーザー情報の取得
  /** @type {Object|null} ユーザー情報 */
    const userInfo = findUserById(requestUserId);
  if (!userInfo) {
    throw new Error('ユーザー情報が見つかりません');
  }

  /** @type {Object} 設定JSONオブジェクト */
    let configJson = {};
  try {
    configJson = JSON.parse(userInfo.configJson || '{}');
  } catch (parseError) {
    warnLog('ConfigJson parse error in initializeQuickStartContext:', parseError.message);
    configJson = {};
  }
  /** @type {string} ユーザーメールアドレス */
  const userEmail = userInfo.adminEmail;

  // クイックスタート繰り返し実行を許可
  // 既存のセットアップがある場合は完全にリセットして新しいセットアップで上書きする
  if (configJson.formCreated || userInfo.spreadsheetId) {

    // 既存のキャッシュを完全にクリアして、新しいセットアップが確実に反映されるようにする
    invalidateUserCache(requestUserId, userEmail, userInfo.spreadsheetId, true);

    // 既存の設定を初期化（重要な情報以外をリセット）
    configJson = {
      setupStatus: 'in_progress',
      createdAt: configJson.createdAt || new Date().toISOString(),
      formCreated: false,
      appPublished: false
    };

    // 強制的に新規作成を保証するためにspreadsheetIdをクリア

    // データベースレベルでもスプレッドシート情報をクリア（確実性を高める）
    updateUser(requestUserId, {
      spreadsheetId: '',
      spreadsheetUrl: '',
      configJson: JSON.stringify(configJson)
    });

    // userInfo オブジェクトもクリア  
    userInfo.spreadsheetId = '';
    userInfo.spreadsheetUrl = '';

    // 更新後のキャッシュを強制同期
    synchronizeCacheAfterCriticalUpdate(requestUserId, userEmail, '', null);
  } else {
  }

  return {
    requestUserId: requestUserId,
    userInfo: userInfo,
    configJson: configJson,
    userEmail: userEmail
  };
}

function quickStartSetup(requestUserId) {
  verifyUserAccess(requestUserId);
  try {
    
    // ステップ0: セットアップコンテキストを初期化
    /** @type {Object} セットアップコンテキスト */
    const setupContext = initializeQuickStartContext(requestUserId);
    /** @type {Object} セットアップコンテキストからの設定JSON */
  const configJson = setupContext.configJson;
    /** @type {string} セットアップコンテキストからのユーザーメール */
  const userEmail = setupContext.userEmail;
    /** @type {Object} セットアップコンテキストからのユーザー情報 */
    const userInfo = setupContext.userInfo;

    // ステップ1-2: ファイル作成とフォルダ管理を実行
    /** @type {Object} 作成されたファイル情報 */
    const createdFiles = createQuickStartFiles(setupContext);
    /** @type {Object} フォーム・スプレッドシート情報 */
  const formAndSsInfo = createdFiles.formAndSsInfo;
    /** @type {Object} 作成されたフォルダ */
  const folder = createdFiles.folder;

    // ステップ3: データベース更新とキャッシュ管理を実行
    /** @type {Object} 更新された設定 */
    const updatedConfig = updateQuickStartDatabase(setupContext, createdFiles);

    // ステップ4: 自動公開処理（重要な新機能）
    /** @type {Object} 自動公開結果 */
    const publishResult = performAutoPublish(requestUserId, formAndSsInfo.sheetName);

    // ステップ5: キャッシュクリアと最終化
    clearExecutionUserInfoCache();
    invalidateUserCache(requestUserId, userEmail, createdFiles.formAndSsInfo.spreadsheetId, true);

    // ステップ6: 最終レスポンス生成
    /** @type {Object} 最終レスポンス */
    const finalResponse = generateQuickStartResponse(setupContext, createdFiles, updatedConfig, publishResult);
    
    return finalResponse;

  } catch (e) {
    errorLog('❌ quickStartSetup エラー: ' + e.message);

    // エラー時はセットアップ状態をリセット
    try {
      /** @type {Object} 現在の設定 */
      let currentConfig = JSON.parse(userInfo.configJson || '{}');
      currentConfig.setupStatus = 'error';
      currentConfig.lastError = e.message;
      currentConfig.errorAt = new Date().toISOString();

      updateUser(requestUserId, {
        configJson: JSON.stringify(currentConfig)
      });
      invalidateUserCache(requestUserId, userEmail, null, false);
      clearExecutionUserInfoCache();
    } catch (updateError) {
      errorLog('エラー状態の更新に失敗: ' + updateError.message);
    }

    return {
      status: 'error',
      message: 'クイックスタートセットアップに失敗しました: ' + e.message,
      webAppUrl: '',
      adminUrl: '',
      viewUrl: '',
      setupUrl: ''
    };
  }
}

/**
 * ユーザー専用フォルダを作成 (マルチテナント対応版)
 */
function createUserFolder(userEmail) {
  try {
    /** @type {string} ルートフォルダ名 */
    const rootFolderName = "StudyQuest - ユーザーデータ";
    /** @type {string} ユーザーフォルダ名 */
    const userFolderName = "StudyQuest - " + userEmail + " - ファイル";

    // ルートフォルダを検索または作成
    /** @type {Object|undefined} ルートフォルダ */
    let rootFolder;
    /** @type {Object} フォルダイテレータ */
    const folders = DriveApp.getFoldersByName(rootFolderName);
    if (folders.hasNext()) {
      rootFolder = folders.next();
    } else {
      rootFolder = DriveApp.createFolder(rootFolderName);
    }

    // ユーザー専用フォルダを作成
    /** @type {Object} ユーザーフォルダイテレータ */
    const userFolders = rootFolder.getFoldersByName(userFolderName);
    if (userFolders.hasNext()) {
      return userFolders.next(); // 既存フォルダを返す
    } else {
      return rootFolder.createFolder(userFolderName);
    }

  } catch (e) {
    errorLog('createUserFolder エラー: ' + e.message);
    return null; // フォルダ作成に失敗してもnullを返して処理を継続
  }
}

/**
 * ハイライト切り替え処理
 */
function processHighlightToggle(spreadsheetId, sheetName, rowIndex) {
  try {
    /** @type {Object} Google Sheets APIサービス */
    const service = getSheetsServiceCached();
    /** @type {Object} ヘッダーインデックス */
    const headerIndices = getHeaderIndices(spreadsheetId, sheetName);
    /** @type {number} ハイライト列インデックス */
    const highlightColumnIndex = headerIndices[COLUMN_HEADERS.HIGHLIGHT];

    if (highlightColumnIndex === undefined) {
      throw new Error('ハイライト列が見つかりません');
    }

    // 現在の値を取得
    /** @type {string} ハイライト列のレンジ */
    const range = "'" + sheetName + "'!" + String.fromCharCode(65 + highlightColumnIndex) + rowIndex;
    /** @type {Object} API応答データ */
    const response = batchGetSheetsData(service, spreadsheetId, [range]);
    /** @type {boolean} ハイライト状態 */
    let isHighlighted = false;
    if (response && response.valueRanges && response.valueRanges[0] &&
        response.valueRanges[0].values && response.valueRanges[0].values[0] &&
        response.valueRanges[0].values[0][0]) {
      isHighlighted = response.valueRanges[0].values[0][0] === 'true';
    }

    // 値を切り替え
    /** @type {string} 新しい値 */
    const newValue = isHighlighted ? 'false' : 'true';
    updateSheetsData(service, spreadsheetId, range, [[newValue]]);

    // ハイライト更新後のキャッシュ無効化
    try {
      if (typeof cacheManager !== 'undefined' && typeof cacheManager.invalidateSheetData === 'function') {
        cacheManager.invalidateSheetData(spreadsheetId, sheetName);
      }
    } catch (cacheError) {
      warnLog('ハイライト後のキャッシュ無効化エラー:', cacheError.message);
    }

    return {
      status: 'success',
      highlighted: !isHighlighted,
      message: isHighlighted ? 'ハイライトを解除しました' : 'ハイライトしました'
    };
  } catch (e) {
    errorLog('ハイライト処理エラー: ' + e.message);
    return { status: 'error', message: e.message };
  }
}

// =================================================================
// 互換性関数（後方互換性のため）
// =================================================================

// getWebAppUrl function removed - now using the unified version from url.gs
function getHeaderIndices(spreadsheetId, sheetName) {

  /** @type {string} キャッシュキー */
  const cacheKey = 'hdr_' + spreadsheetId + '_' + sheetName;
  /** @type {Object} インデックス */
  let indices = getHeadersCached(spreadsheetId, sheetName);

  // 理由列が取得できていない場合のみキャッシュを無効化して再取得
  if (!indices || indices[COLUMN_HEADERS.REASON] === undefined) {
    cacheManager.remove(cacheKey);
    // 直接再取得は避け、getHeadersCachedの内部ロジックに委譲
    /** @type {Object} 更新されたインデックス */
    const refreshedIndices = getHeadersCached(spreadsheetId, sheetName);
    if (refreshedIndices && Object.keys(refreshedIndices).length > 0) {
      indices = refreshedIndices;
    }
  }

  return indices;
}

function getSheetColumns(userId, sheetId) {
  verifyUserAccess(userId);
  try {
    /** @type {Object|null} ユーザー情報 */
    const userInfo = findUserById(userId);
    if (!userInfo || !userInfo.spreadsheetId) {
      throw new Error('ユーザー情報またはスプレッドシートIDが見つかりません');
    }

    /** @type {Object} スプレッドシートオブジェクト */
    const spreadsheet = openSpreadsheetOptimized(userInfo.spreadsheetId);
    /** @type {Object} シートオブジェクト */
    const sheet = spreadsheet.getSheetById(sheetId);

    if (!sheet) {
      throw new Error('指定されたシートが見つかりません: ' + sheetId);
    }

    /** @type {number} 最終列番号 */
    const lastColumn = sheet.getLastColumn();
    if (lastColumn === 0) {
      return []; // 列がない場合
    }

    /** @type {Array} ヘッダー配列 */
    const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];

    // クライアントサイドが期待する形式に変換
    /** @type {Array} カラム配列 */
    const columns = headers.map(function(headerName) {
      return {
        id: headerName,
        name: headerName
      };
    });

    infoLog('✅ getSheetColumns: Found %s columns for sheetId %s', columns.length, sheetId);
    return columns;

  } catch (e) {
    errorLog('getSheetColumns エラー: ' + e.message);
    errorLog('Error details:', e.stack);
    throw new Error('列の取得に失敗しました: ' + e.message);
  }
}

/**
 * GoogleフォームURLからフォームIDを抽出
 */
function extractFormIdFromUrl(url) {
  if (!url) return '';

  try {
    // Regular expression to extract form ID from Google Forms URLs
    /** @type {Array|null} フォームIDマッチ */
    const formIdMatch = url.match(/\/forms\/d\/([a-zA-Z0-9-_]+)/);
    if (formIdMatch && formIdMatch[1]) {
      return formIdMatch[1];
    }

    // Alternative pattern for e/ URLs
    /** @type {Array|null} eフォームIDマッチ */
    const eFormIdMatch = url.match(/\/forms\/d\/e\/([a-zA-Z0-9-_]+)/);
    if (eFormIdMatch && eFormIdMatch[1]) {
      return eFormIdMatch[1];
    }

    return '';
  } catch (e) {
    warnLog('フォームID抽出エラー: ' + e.message);
    return '';
  }
}

// =================================================================
// リアクション処理関数
// =================================================================

/**
 * リアクション処理
 */
function processReaction(spreadsheetId, sheetName, rowIndex, reactionKey, reactingUserEmail) {
  try {
    // 統一ロック管理でリアクション処理を実行
    return executeWithStandardizedLock('WRITE_OPERATION', 'processReaction', () => {

      /** @type {Object} Google Sheets APIサービス */
    const service = getSheetsServiceCached();
      /** @type {Object} ヘッダーインデックス */
      const headerIndices = getHeaderIndices(spreadsheetId, sheetName);

      // すべてのリアクション列を取得してユーザーの重複リアクションをチェック
      /** @type {Array} 全リアクション範囲 */
      const allReactionRanges = [];
      /** @type {Object} 全リアクションカラム */
      const allReactionColumns = {};
      /** @type {number|null} ターゲットリアクションカラムインデックス */
      let targetReactionColumnIndex = null;

      // 全リアクション列の情報を準備
      REACTION_KEYS.forEach(function(key) {
        /** @type {string} カラム名 */
        const columnName = COLUMN_HEADERS[key];
        /** @type {number|undefined} カラムインデックス */
        const columnIndex = headerIndices[columnName];
        if (columnIndex !== undefined) {
          /** @type {string} 範囲文字列 */
          const range = "'" + sheetName + "'!" + String.fromCharCode(65 + columnIndex) + rowIndex;
          allReactionRanges.push(range);
          allReactionColumns[key] = {
            columnIndex: columnIndex,
            range: range
          };
          if (key === reactionKey) {
            targetReactionColumnIndex = columnIndex;
          }
        }
      });

      if (targetReactionColumnIndex === null) {
        throw new Error('対象リアクション列が見つかりません: ' + reactionKey);
      }

      // 全リアクション列の現在の値を一括取得
      /** @type {Object} APIレスポンス */
      const response = batchGetSheetsData(service, spreadsheetId, allReactionRanges);
      /** @type {Array} 更新データ */
      const updateData = [];
      /** @type {string|null} ユーザーアクション */
      let userAction = null;
      /** @type {number} ターゲットカウント */
      let targetCount = 0;

      // 各リアクション列を処理
      /** @type {number} 範囲インデックス */
      let rangeIndex = 0;
      REACTION_KEYS.forEach(function(key) {
        if (!allReactionColumns[key]) return;

        /** @type {string} 現在のリアクション文字列 */
        let currentReactionString = '';
        if (response && response.valueRanges && response.valueRanges[rangeIndex] &&
            response.valueRanges[rangeIndex].values && response.valueRanges[rangeIndex].values[0] &&
            response.valueRanges[rangeIndex].values[0][0]) {
          currentReactionString = response.valueRanges[rangeIndex].values[0][0];
        }

        /** @type {Array} 現在のリアクション */
        const currentReactions = parseReactionString(currentReactionString);
        /** @type {number} ユーザーインデックス */
        const userIndex = currentReactions.indexOf(reactingUserEmail);

        if (key === reactionKey) {
          // 対象リアクション列の処理
          if (userIndex >= 0) {
            // 既にリアクション済み → 削除（トグル）
            currentReactions.splice(userIndex, 1);
            userAction = 'removed';
          } else {
            // 未リアクション → 追加
            currentReactions.push(reactingUserEmail);
            userAction = 'added';
          }
          targetCount = currentReactions.length;
        } else {
          // 他のリアクション列からユーザーを削除（1人1リアクション制限）
          if (userIndex >= 0) {
            currentReactions.splice(userIndex, 1);
          }
        }

        // 更新データを準備
        /** @type {string} 更新されたリアクション文字列 */
        const updatedReactionString = currentReactions.join(', ');
        updateData.push({
          range: allReactionColumns[key].range,
          values: [[updatedReactionString]]
        });

        rangeIndex++;
      });

      // すべての更新を一括実行
      if (updateData.length > 0) {
        batchUpdateSheetsData(service, spreadsheetId, updateData);
      }

      // リアクション更新後のキャッシュ無効化
      try {
        if (typeof cacheManager !== 'undefined' && typeof cacheManager.invalidateSheetData === 'function') {
          cacheManager.invalidateSheetData(spreadsheetId, sheetName);
        }
      } catch (cacheError) {
        warnLog('リアクション後のキャッシュ無効化エラー:', cacheError.message);
      }

      // 更新後のリアクション状態を計算（追加のAPI呼び出しなし）
      /** @type {Object} リアクション状態 */
      const reactionStates = {};
      /** @type {number} 更新インデックス */
      let updateIndex = 0;
      REACTION_KEYS.forEach(function(key) {
        if (!allReactionColumns[key]) return;
        
        // 更新データから最終状態を取得
        /** @type {string} 更新されたリアクション文字列 */
        const updatedReactionString = updateData[updateIndex].values[0][0];
        /** @type {Array} 更新されたリアクション */
        const updatedReactions = parseReactionString(updatedReactionString);
        
        reactionStates[key] = {
          count: updatedReactions.length,
          reacted: updatedReactions.indexOf(reactingUserEmail) !== -1
        };
        
        updateIndex++;
      });

      return {
        status: 'success',
        message: 'リアクションを更新しました。',
        action: userAction,
        count: targetCount,
        reactionStates: reactionStates
      };
    });

  } catch (e) {
    errorLog('リアクション処理エラー: ' + e.message);
    return {
      status: 'error',
      message: 'リアクションの処理に失敗しました: ' + e.message
    };
  }
}

// =================================================================
// フォーム作成機能
// =================================================================

/**
 * 回答ボードの公開を停止する (マルチテナント対応版)
 * AdminPanel.htmlから呼び出される
 * @param {string} requestUserId - リクエスト元のユーザーID
 */
// NOTE: unpublishBoard関数はconfig.gsに実装済み

/**
 * フォーム作成ファクトリ
 * @param {Object} options - フォーム作成オプション
 * @param {string} options.userEmail - ユーザーのメールアドレス
 * @param {string} options.userId - ユーザーID
 * @param {string} options.questions - 質問タイプ（'default'など）
 * @param {string} options.formDescription - フォームの説明
 * @returns {Object} フォーム作成結果
 */
function createFormFactory(options) {
  try {
    /** @type {string} ユーザーメール */
    const userEmail = options.userEmail;
    /** @type {string} ユーザーID */
    const userId = options.userId;
    /** @type {string} フォーム説明 */
    const formDescription = options.formDescription || 'みんなの回答ボードへの投稿フォームです。';

    // タイムスタンプ生成（日時を含む）
    /** @type {Date} 現在日時 */
    const now = new Date();
    /** @type {string} 日時文字列 */
    const dateTimeString = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年MM月dd日 HH:mm');

    // フォームタイトル生成
    /** @type {string} フォームタイトル */
    const formTitle = options.formTitle || ('みんなの回答ボード ' + dateTimeString);

    // フォーム作成
    /** @type {Object} 新規作成されたGoogle Form */
    const form = FormApp.create(formTitle);
    form.setDescription(formDescription);
    
    // フォームを即座にユーザーフォルダに移動（作成直後）
    try {
      const userFolder = createUserFolder(userEmail);
      if (userFolder) {
        const formFile = DriveApp.getFileById(form.getId());
        formFile.moveTo(userFolder);
      }
    } catch (folderMoveError) {
      warnLog('⚠️ フォーム作成直後の移動に失敗（後で再移動されます）:', folderMoveError.message);
    }

    // 基本的な質問を追加
    addUnifiedQuestions(form, options.questions || 'default', options.customConfig || {});

    // メールアドレス収集を有効化（スプレッドシート連携前に設定）
    try {
      if (typeof form.setEmailCollectionType === 'function') {
        form.setEmailCollectionType(FormApp.EmailCollectionType.VERIFIED);
      } else {
        form.setCollectEmail(true);
      }
    } catch (emailError) {
      warnLog('Email collection setting failed:', emailError.message);
    }

    // スプレッドシート作成
    /** @type {Object} スプレッドシート結果 */
    const spreadsheetResult = createLinkedSpreadsheet(userEmail, form, dateTimeString);

    // リアクション関連列を追加
    addReactionColumnsToSpreadsheet(spreadsheetResult.spreadsheetId, spreadsheetResult.sheetName);

    return {
      formId: form.getId(),
      formUrl: form.getPublishedUrl(),
      viewFormUrl: form.getPublishedUrl(),
      editFormUrl: typeof form.getEditUrl === 'function' ? form.getEditUrl() : '',
      spreadsheetId: spreadsheetResult.spreadsheetId,
      spreadsheetUrl: spreadsheetResult.spreadsheetUrl,
      sheetName: spreadsheetResult.sheetName
    };

  } catch (error) {
    errorLog('createFormFactory エラー:', error.message);
    throw new Error('フォーム作成ファクトリでエラーが発生しました: ' + error.message);
  }
}

/**
 * 統一された質問をフォームに追加
 * @param {GoogleAppsScript.Forms.Form} form - フォームオブジェクト
 * @param {string} questionType - 質問タイプ
 * @param {Object} customConfig - カスタム設定
 */
function addUnifiedQuestions(form, questionType, customConfig) {
  try {
    /** @type {Object} 設定オブジェクト */
    const config = getQuestionConfig(questionType, customConfig);

    form.setCollectEmail(false);

    if (questionType === 'simple') {
      /** @type {Object} クラスアイテム */
      const classItem = form.addListItem();
      classItem.setTitle(config.classQuestion.title);
      classItem.setChoiceValues(config.classQuestion.choices);
      classItem.setRequired(true);

      /** @type {Object} 名前アイテム */
      const nameItem = form.addTextItem();
      nameItem.setTitle(config.nameQuestion.title);
      nameItem.setRequired(true);

      /** @type {Object} メインアイテム */
      const mainItem = form.addTextItem();
      mainItem.setTitle(config.mainQuestion.title);
      mainItem.setRequired(true);

      /** @type {Object} 理由アイテム */
      const reasonItem = form.addParagraphTextItem();
      reasonItem.setTitle(config.reasonQuestion.title);
      reasonItem.setHelpText(config.reasonQuestion.helpText);
      /** @type {Object} バリデーション */
      const validation = FormApp.createParagraphTextValidation()
        .requireTextLengthLessThanOrEqualTo(140)
        .build();
      reasonItem.setValidation(validation);
      reasonItem.setRequired(false);
    } else if (questionType === 'custom' && customConfig) {

      // クラス選択肢（有効な場合のみ）
      if (customConfig.enableClass && customConfig.classQuestion && customConfig.classQuestion.choices && customConfig.classQuestion.choices.length > 0) {
        /** @type {Object} クラスアイテム */
        const classItem = form.addListItem();
        classItem.setTitle('クラス');
        classItem.setChoiceValues(customConfig.classQuestion.choices);
        classItem.setRequired(true);
      }

      // 名前欄（常にオン）
      /** @type {Object} 名前アイテム */
      const nameItem = form.addTextItem();
      nameItem.setTitle('名前');
      nameItem.setRequired(false);

      // メイン質問
      /** @type {string} メイン質問タイトル */
      const mainQuestionTitle = customConfig.mainQuestion ? customConfig.mainQuestion.title : '今回のテーマについて、あなたの考えや意見を聞かせてください';
      /** @type {Object|undefined} メインアイテム */
      let mainItem;
      /** @type {string} 質問タイプ */
      const questionType = customConfig.mainQuestion ? customConfig.mainQuestion.type : 'text';

      switch(questionType) {
        case 'text':
          mainItem = form.addTextItem();
          break;
        case 'multiple':
          mainItem = form.addCheckboxItem();
          if (customConfig.mainQuestion && customConfig.mainQuestion.choices && customConfig.mainQuestion.choices.length > 0) {
            mainItem.setChoiceValues(customConfig.mainQuestion.choices);
          }
          if (customConfig.mainQuestion && customConfig.mainQuestion.includeOthers && typeof mainItem.showOtherOption === 'function') {
            mainItem.showOtherOption(true);
          }
          break;
        case 'choice':
          mainItem = form.addMultipleChoiceItem();
          if (customConfig.mainQuestion && customConfig.mainQuestion.choices && customConfig.mainQuestion.choices.length > 0) {
            mainItem.setChoiceValues(customConfig.mainQuestion.choices);
          }
          if (customConfig.mainQuestion && customConfig.mainQuestion.includeOthers && typeof mainItem.showOtherOption === 'function') {
            mainItem.showOtherOption(true);
          }
          break;
        default:
          mainItem = form.addParagraphTextItem();
      }
      mainItem.setTitle(mainQuestionTitle);
      mainItem.setRequired(true);

      // 理由欄（常にオン）
      /** @type {Object} 理由アイテム */
      const reasonItem = form.addParagraphTextItem();
      reasonItem.setTitle('そう考える理由や体験があれば教えてください。');
      reasonItem.setRequired(false);
    } else {
      /** @type {Object} クラスアイテム */
      const classItem = form.addTextItem();
      classItem.setTitle(config.classQuestion.title);
      classItem.setHelpText(config.classQuestion.helpText);
      classItem.setRequired(true);

      /** @type {Object} メインアイテム */
      const mainItem = form.addParagraphTextItem();
      mainItem.setTitle(config.mainQuestion.title);
      mainItem.setHelpText(config.mainQuestion.helpText);
      mainItem.setRequired(true);

      /** @type {Object} 理由アイテム */
      const reasonItem = form.addParagraphTextItem();
      reasonItem.setTitle(config.reasonQuestion.title);
      reasonItem.setHelpText(config.reasonQuestion.helpText);
      reasonItem.setRequired(false);

      /** @type {Object} 名前アイテム */
      const nameItem = form.addTextItem();
      nameItem.setTitle(config.nameQuestion.title);
      nameItem.setHelpText(config.nameQuestion.helpText);
      nameItem.setRequired(false);
    }
  } catch (error) {
    errorLog('addUnifiedQuestions エラー:', error.message);
    throw error;
  }
}

/**
 * 質問設定を取得
 * @param {string} questionType - 質問タイプ
 * @param {Object} customConfig - カスタム設定
 * @returns {Object} 質問設定
 */
function getQuestionConfig(questionType, customConfig) {
  // 統一されたテンプレート設定（simple のみ使用）
  /** @type {Object} 設定オブジェクト */
  const config = {
    classQuestion: {
      title: 'クラス',
      helpText: '',
      choices: ['クラス1', 'クラス2', 'クラス3', 'クラス4']
    },
    nameQuestion: {
      title: '名前',
      helpText: ''
    },
    mainQuestion: {
      title: 'あなたの考えや気づいたことを教えてください',
      helpText: '',
      choices: ['気づいたことがある。', '疑問に思うことがある。', 'もっと知りたいことがある。'],
      type: 'paragraph'
    },
    reasonQuestion: {
      title: 'そう考える理由や体験があれば教えてください（任意）',
      helpText: '',
      type: 'paragraph'
    }
  };

  // カスタム設定をマージ
  if (customConfig && typeof customConfig === 'object') {
    for (var key in customConfig) {
      if (config[key]) {
        Object.assign(config[key], customConfig[key]);
      }
    }
  }

  return config;
}

/**
 * 質問テンプレート取得用 API
 * @returns {ContentService.TextOutput} JSON形式の質問設定
 */
function doGetQuestionConfig() {
  try {
    // 現在の日時を取得してタイトルに含める
    const now = new Date();
    const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');

    const cfg = getQuestionConfig('simple');

    // タイトルにタイムスタンプを追加
    cfg.formTitle = 'フォーム作成 - ' + timestamp;

    return ContentService.createTextOutput(JSON.stringify(cfg)).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    errorLog('doGetQuestionConfig error:', error);
    return ContentService.createTextOutput(JSON.stringify({
      error: 'Failed to get question config',
      details: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * クラス選択肢をデータベースに保存
 */
function saveClassChoices(userId, classChoices) {
  try {
    /** @type {string} 現在のユーザーID */
    const currentUserId = userId;
    /** @type {Object|null} ユーザー情報 */
    const userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    /** @type {Object} 設定JSONオブジェクト */
    let configJson = {};
    try {
      configJson = JSON.parse(userInfo.configJson || '{}');
    } catch (parseError) {
      warnLog('ConfigJson parse error in saveClassChoices:', parseError.message);
      configJson = {};
    }
    configJson.savedClassChoices = classChoices;
    configJson.lastClassChoicesUpdate = new Date().toISOString();

    updateUser(currentUserId, {
      configJson: JSON.stringify(configJson)
    });

    return { status: 'success', message: 'クラス選択肢が保存されました' };
  } catch (error) {
    errorLog('クラス選択肢保存エラー:', error.message);
    return { status: 'error', message: 'クラス選択肢の保存に失敗しました: ' + error.message };
  }
}

/**
 * 保存されたクラス選択肢を取得
 */
function getSavedClassChoices(userId) {
  try {
    /** @type {string} 現在のユーザーID */
    const currentUserId = userId;
    /** @type {Object|null} ユーザー情報 */
    const userInfo = findUserById(currentUserId);
    if (!userInfo) {
      return { status: 'error', message: 'ユーザー情報が見つかりません' };
    }

    /** @type {Object} 設定JSONオブジェクト */
    let configJson = {};
    try {
      configJson = JSON.parse(userInfo.configJson || '{}');
    } catch (parseError) {
      warnLog('ConfigJson parse error in getSavedClassChoices:', parseError.message);
      configJson = {};
    }
    /** @type {Array} 保存されたクラス選択肢 */
    const savedClassChoices = configJson.savedClassChoices || ['クラス1', 'クラス2', 'クラス3', 'クラス4'];

    return {
      status: 'success',
      classChoices: savedClassChoices,
      lastUpdate: configJson.lastClassChoicesUpdate
    };
  } catch (error) {
    errorLog('クラス選択肢取得エラー:', error.message);
    return { status: 'error', message: 'クラス選択肢の取得に失敗しました: ' + error.message };
  }
}

/**
 * 統合フォーム作成関数（Phase 2: 最適化版）
 * プリセット設定を使用して重複コードを排除
 */
const FORM_PRESETS = {
  quickstart: {
    titlePrefix: 'みんなの回答ボード',
    questions: 'custom',
    description: 'このフォームは学校での学習や話し合いで使います。みんなで考えを共有して、お互いから学び合いましょう。\n\n【デジタル市民としてのお約束】\n• 相手を思いやる気持ちを持ち、建設的な意見を心がけましょう\n• 事実に基づいた正しい情報を共有しましょう\n• 多様な意見を尊重し、違いを学びの機会としましょう\n• 個人情報やプライバシーに関わる内容は書かないようにしましょう\n• 責任ある発言を心がけ、みんなが安心して参加できる環境を作りましょう\n\nあなたの意見や感想は、クラスメイトの学びを深める大切な資源です。',
    config: {
      mainQuestion: 'あなたの考えや気づいたことを教えてください',
      questionType: 'text',
      enableClass: false,
      includeOthers: false
    }
  },
  custom: {
    titlePrefix: 'みんなの回答ボード',
    questions: 'custom',
    description: 'このフォームは学校での学習や話し合いで使います。みんなで考えを共有して、お互いから学び合いましょう。\n\n【デジタル市民としてのお約束】\n• 相手を思いやる気持ちを持ち、建設的な意見を心がけましょう\n• 事実に基づいた正しい情報を共有しましょう\n• 多様な意見を尊重し、違いを学びの機会としましょう\n• 個人情報やプライバシーに関わる内容は書かないようにしましょう\n• 責任ある発言を心がけ、みんなが安心して参加できる環境を作りましょう\n\nあなたの意見や感想は、クラスメイトの学びを深める大切な資源です。',
    config: {} // Will be overridden by user input
  },
  study: {
    titlePrefix: 'みんなの回答ボード',
    questions: 'simple', // Default, can be overridden
    description: 'このフォームは学校での学習や話し合いで使います。みんなで考えを共有して、お互いから学び合いましょう。\n\n【デジタル市民としてのお約束】\n• 相手を思いやる気持ちを持ち、建設的な意見を心がけましょう\n• 事実に基づいた正しい情報を共有しましょう\n• 多様な意見を尊重し、違いを学びの機会としましょう\n• 個人情報やプライバシーに関わる内容は書かないようにしましょう\n• 責任ある発言を心がけ、みんなが安心して参加できる環境を作りましょう\n\nあなたの意見や感想は、クラスメイトの学びを深める大切な資源です。',
    config: {}
  }
};

/**
 * 統合フォーム作成関数
 * @param {string} presetType - 'quickstart' | 'custom' | 'study'
 * @param {string} userEmail - ユーザーメールアドレス
 * @param {string} userId - ユーザーID
 * @param {Object} overrides - プリセットを上書きする設定
 * @returns {Object} フォーム作成結果
 */
function createUnifiedForm(presetType, userEmail, userId, overrides = {}) {
  try {
    const preset = FORM_PRESETS[presetType];
    if (!preset) {
      throw new Error('未知のプリセットタイプ: ' + presetType);
    }

    const now = new Date();
    const dateTimeString = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年MM月dd日 HH:mm');

    // タイトル生成（上書き可能）
    const titlePrefix = overrides.titlePrefix || preset.titlePrefix;
    const formTitle = overrides.formTitle || (titlePrefix + ' ' + dateTimeString);

    // 設定をマージ（プリセット + ユーザー設定）
    const mergedConfig = { ...preset.config, ...overrides.customConfig };

    const factoryOptions = {
      userEmail: userEmail,
      userId: userId,
      formTitle: formTitle,
      questions: overrides.questions || preset.questions,
      formDescription: overrides.formDescription || preset.description,
      customConfig: mergedConfig
    };

    return createFormFactory(factoryOptions);
  } catch (error) {
    errorLog('createUnifiedForm Error (' + presetType + '):', error.message);
    throw new Error('フォーム作成に失敗しました (' + presetType + '): ' + error.message);
  }
}
/**
 * リンクされたスプレッドシートを作成
 * @param {string} userEmail - ユーザーメール
 * @param {GoogleAppsScript.Forms.Form} form - フォーム
 * @param {string} dateTimeString - 日付時刻文字列
 * @returns {Object} スプレッドシート情報
 */
function createLinkedSpreadsheet(userEmail, form, dateTimeString) {
  try {
    // スプレッドシート名を設定
    /** @type {string} スプレッドシート名 */
    const spreadsheetName = userEmail + ' - 回答データ - ' + dateTimeString;

    // 新しいスプレッドシートを作成
    /** @type {Object} スプレッドシートオブジェクト */
    const spreadsheetObj = SpreadsheetApp.create(spreadsheetName);
    /** @type {string} スプレッドシートID */
    const spreadsheetId = spreadsheetObj.getId();

    // スプレッドシートを即座にユーザーフォルダに移動（作成直後）
    try {
      const userFolder = createUserFolder(userEmail);
      if (userFolder) {
        const ssFile = DriveApp.getFileById(spreadsheetId);
        ssFile.moveTo(userFolder);
      }
    } catch (folderMoveError) {
      warnLog('⚠️ スプレッドシート作成直後の移動に失敗（後で再移動されます）:', folderMoveError.message);
    }

    // スプレッドシートの共有設定を同一ドメイン閲覧可能に設定
    try {
      /** @type {Object} ファイルオブジェクト */
      const file = DriveApp.getFileById(spreadsheetId);
      if (!file) {
        throw new Error('スプレッドシートが見つかりません: ' + spreadsheetId);
      }

      // 同一ドメインで閲覧可能に設定（教育機関対応）
      file.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.VIEW);

      // 作成者（現在のユーザー）は所有者として保持

    } catch (sharingError) {
      warnLog('共有設定の変更に失敗しましたが、処理を続行します: ' + sharingError.message);
    }

    // フォームの回答先としてスプレッドシートを設定
    form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheetId);

    // スプレッドシート名を再設定（フォーム設定後に名前が変わる場合があるため）
    spreadsheetObj.rename(spreadsheetName);

    // 重要: フォーム連携後のシート名を正確に取得（連携により変更される可能性があるため）
    Utilities.sleep(1000); // フォーム連携完了を待つ
    
    // スプレッドシートを再読み込みして最新のシート情報を取得
    /** @type {string} シート名 - スコープを外側に定義 */
    let sheetName = 'フォームの回答 1'; // デフォルト値
    
    try {
      spreadsheetObj = SpreadsheetApp.openById(spreadsheetId);
      /** @type {Array} シート配列 */
      const sheets = spreadsheetObj.getSheets();
      /** @type {string} 実際のシート名 */
      let actualSheetName = String(sheets[0].getName());
      
      // シート名が不正な値でないことを確認
      if (!actualSheetName || actualSheetName === 'true' || actualSheetName.trim() === '') {
        actualSheetName = 'フォームの回答 1'; // Google Formsのデフォルトシート名
        warnLog('不正なシート名が検出されました。デフォルトのシート名を使用します: ' + actualSheetName);
      }
      
      sheetName = actualSheetName;
      infoLog('✅ フォーム連携後の正確なシート名を取得:', sheetName);
      
    } catch (sheetNameError) {
      errorLog('❌ フォーム連携後のシート名取得エラー:', sheetNameError.message);
      // フォールバック: 標準的なシート名を使用（既に設定済み）
      warnLog('⚠️ シート名取得エラーのため、標準名を使用:', sheetName);
    }

    // サービスアカウントとスプレッドシートを共有（失敗しても処理継続）
    try {
      shareSpreadsheetWithServiceAccount(spreadsheetId);
    } catch (shareError) {
      warnLog('サービスアカウント共有エラー（処理継続）:', shareError.message);
      // 権限エラーの場合でも、スプレッドシート作成自体は成功とみなす
    }

    // スプレッドシートURL取得（リトライ機能付き）
    let spreadsheetUrl = '';
    try {
      spreadsheetUrl = spreadsheetObj.getUrl();
      if (!spreadsheetUrl) {
        throw new Error('URL取得結果が空です');
      }
    } catch (urlError) {
      warnLog('⚠️ 初回URL取得失敗、リトライします:', urlError.message);
      
      // リトライ処理（最大3回）
      for (let retry = 0; retry < 3; retry++) {
        try {
          Utilities.sleep(1000 * (retry + 1)); // 1秒、2秒、3秒の間隔
          spreadsheetObj = SpreadsheetApp.openById(spreadsheetId); // 再読み込み
          spreadsheetUrl = spreadsheetObj.getUrl();
          if (spreadsheetUrl) {
            infoLog('✅ スプレッドシートURL取得成功（' + (retry + 1) + '回目のリトライ）:', spreadsheetUrl);
            break;
          }
        } catch (retryError) {
          warnLog('❌ リトライ' + (retry + 1) + '回目失敗:', retryError.message);
        }
      }
      
      if (!spreadsheetUrl) {
        // 最後の手段: IDからURLを構築
        spreadsheetUrl = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/edit';
        warnLog('⚠️ URLを手動構築しました:', spreadsheetUrl);
      }
    }

    return {
      spreadsheetId: spreadsheetId,
      spreadsheetUrl: spreadsheetUrl,
      sheetName: sheetName
    };

  } catch (error) {
    errorLog('createLinkedSpreadsheet エラー:', error.message);
    throw new Error('スプレッドシート作成に失敗しました: ' + error.message);
  }
}

/**
 * スプレッドシートをサービスアカウントと共有
 * @param {string} spreadsheetId - スプレッドシートID
 */
function shareSpreadsheetWithServiceAccount(spreadsheetId) {
  try {
    /** @type {string} サービスアカウントメール */
    const serviceAccountEmail = getServiceAccountEmail();

    if (!serviceAccountEmail || serviceAccountEmail === 'サービスアカウント未設定' || serviceAccountEmail === 'サービスアカウント設定エラー') {
      throw new Error('サービスアカウントのメールアドレスが取得できません: ' + serviceAccountEmail);
    }
    // DriveAppを使用してスプレッドシートをサービスアカウントと共有
    try {
      /** @type {Object} ファイルオブジェクト */
      const file = DriveApp.getFileById(spreadsheetId);
      if (!file) {
        throw new Error('スプレッドシートが見つかりません: ' + spreadsheetId);
      }
      file.addEditor(serviceAccountEmail);
    } catch (driveError) {
      errorLog('DriveApp error:', driveError.message);
      throw new Error('Drive API操作に失敗しました: ' + driveError.message);
    }
  } catch (error) {
    errorLog('shareSpreadsheetWithServiceAccount エラー:', error.message);
    throw new Error('サービスアカウントとの共有に失敗しました: ' + error.message);
  }
}

/**
 * すべての既存スプレッドシートをサービスアカウントと共有
 * @returns {Object} 共有結果
 */
function shareAllSpreadsheetsWithServiceAccount() {
  try {

    /** @type {Array} 全ユーザー */
    const allUsers = getAllUsers();
    /** @type {Array} 結果配列 */
    const results = [];
    /** @type {number} 成功カウント */
    let successCount = 0;
    /** @type {number} エラーカウント */
    let errorCount = 0;

    for (let i = 0; i < allUsers.length; i++) {
      /** @type {Object} ユーザーオブジェクト */
      const user = allUsers[i];
      if (user.spreadsheetId && user.isActive === 'true') {
        try {
          shareSpreadsheetWithServiceAccount(user.spreadsheetId);
          results.push({
            userId: user.userId,
            adminEmail: user.adminEmail,
            spreadsheetId: user.spreadsheetId,
            status: 'success'
          });
          successCount++;
        } catch (shareError) {
          results.push({
            userId: user.userId,
            adminEmail: user.adminEmail,
            spreadsheetId: user.spreadsheetId,
            status: 'error',
            error: shareError.message
          });
          errorCount++;
          errorLog('共有失敗:', user.adminEmail, shareError.message);
        }
      }
    }
    return {
      status: 'completed',
      totalUsers: allUsers.length,
      successCount: successCount,
      errorCount: errorCount,
      results: results
    };

  } catch (error) {
    errorLog('shareAllSpreadsheetsWithServiceAccount エラー:', error.message);
    throw new Error('全スプレッドシート共有処理でエラーが発生しました: ' + error.message);
  }
}

/**
 * フォーム作成
 */

/**
 * サービスアカウントをスプレッドシートに追加
 */
function addServiceAccountToSpreadsheet(spreadsheetId) {
  try {
    /** @type {Object} プロパティ */
    const props = PropertiesService.getScriptProperties();
    /** @type {Object} サービスアカウント認証情報 */
    const serviceAccountCreds = JSON.parse(props.getProperty(SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS));
    /** @type {string} サービスアカウントメール */
    const serviceAccountEmail = serviceAccountCreds.client_email;

    /** @type {Object} スプレッドシートオブジェクト */
    const spreadsheet = openSpreadsheetOptimized(spreadsheetId);

    // サービスアカウントを編集者として追加
    if (serviceAccountEmail) {
      spreadsheet.addEditor(serviceAccountEmail);

      // セッション管理：サービスアカウントアクセス権限の記録
      try {
        const sessionData = {
          serviceAccountEmail: serviceAccountEmail,
          spreadsheetId: spreadsheetId,
          accessGranted: new Date().toISOString(),
          accessType: 'service_account_editor',
          securityLevel: 'domain_view'
        };
      } catch (sessionLogError) {
        warnLog('セッション記録でエラー:', sessionLogError.message);
      }
    }

    // 同一ドメインユーザーは共有設定により閲覧可能

  } catch (e) {
    errorLog('サービスアカウントの追加に失敗: ' + e.message);
    // エラーでも処理は継続
  }
}

/**
 * 既存ユーザーのスプレッドシートアクセス権限を修復
 * ユーザーがスプレッドシートにアクセスできない場合に使用
 * @param {string} userEmail - ユーザーのメールアドレス
 * @param {string} spreadsheetId - スプレッドシートID
 */
function repairUserSpreadsheetAccess(userEmail, spreadsheetId) {
  try {

    // DriveApp経由で共有設定を変更
    /** @type {Object|undefined} ファイル */
    let file;
    try {
      file = DriveApp.getFileById(spreadsheetId);
      if (!file) {
        throw new Error('スプレッドシートが見つかりません: ' + spreadsheetId);
      }
    } catch (driveError) {
      errorLog('DriveApp.getFileById error:', driveError.message);
      throw new Error('スプレッドシートへのアクセスに失敗しました: ' + driveError.message);
    }

    // ドメイン全体でアクセス可能に設定
    try {
      file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT);
    } catch (domainSharingError) {
      warnLog('ドメイン共有設定に失敗: ' + domainSharingError.message);

      // ドメイン共有に失敗した場合は個別にユーザーを追加
      try {
        file.addEditor(userEmail);
      } catch (individualError) {
        errorLog('個別ユーザー追加も失敗: ' + individualError.message);
      }
    }

    // SpreadsheetApp経由でも編集者として追加
    try {
      /** @type {Object} スプレッドシートオブジェクト */
    const spreadsheet = openSpreadsheetOptimized(spreadsheetId);
      spreadsheet.addEditor(userEmail);
    } catch (spreadsheetAddError) {
      warnLog('SpreadsheetApp経由の追加で警告: ' + spreadsheetAddError.message);
    }

    // サービスアカウントも確認
    /** @type {Object} プロパティ */
    const props = PropertiesService.getScriptProperties();
    /** @type {Object} サービスアカウント認証情報 */
    const serviceAccountCreds = JSON.parse(props.getProperty(SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS));
    /** @type {string} サービスアカウントメール */
    const serviceAccountEmail = serviceAccountCreds.client_email;

    if (serviceAccountEmail) {
      try {
        file.addEditor(serviceAccountEmail);
      } catch (serviceError) {
        warnLog('サービスアカウント追加で警告: ' + serviceError.message);
      }
    }

    return {
      success: true,
      message: 'スプレッドシートアクセス権限を修復しました。ドメイン全体でアクセス可能です。'
    };

  } catch (e) {
    errorLog('スプレッドシートアクセス権限の修復に失敗: ' + e.message);
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * 管理パネル専用：権限問題の緊急修復
 * @param {string} userEmail - ユーザーメール
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {object} 修復結果
 */
function addReactionColumnsToSpreadsheet(spreadsheetId, sheetName) {
  try {
    /** @type {Object} スプレッドシートオブジェクト */
    const spreadsheet = openSpreadsheetOptimized(spreadsheetId);
    /** @type {Object} シートオブジェクト */
    const sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.getSheets()[0];

    /** @type {Array} 追加ヘッダー */
    const additionalHeaders = [
      COLUMN_HEADERS.UNDERSTAND,
      COLUMN_HEADERS.LIKE,
      COLUMN_HEADERS.CURIOUS,
      COLUMN_HEADERS.HIGHLIGHT
    ];

    // 効率的にヘッダー情報を取得
    /** @type {Array} 現在のヘッダー */
    const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    /** @type {number} 開始カラム */
    const startCol = currentHeaders.length + 1;

    // バッチでヘッダーを追加
    sheet.getRange(1, startCol, 1, additionalHeaders.length).setValues([additionalHeaders]);

    // スタイリングを一括適用
    /** @type {Object} 全ヘッダー範囲 */
    const allHeadersRange = sheet.getRange(1, 1, 1, currentHeaders.length + additionalHeaders.length);
    allHeadersRange.setFontWeight('bold').setBackground('#E3F2FD');

    // 自動リサイズ（エラーが出ても続行）
    try {
      sheet.autoResizeColumns(1, allHeadersRange.getNumColumns());
    } catch (resizeError) {
      warnLog('Auto-resize failed:', resizeError.message);
    }

  }
  catch (e) {
    errorLog('リアクション列追加エラー: ' + e.message);
    // エラーでも処理は継続
  }
}

/**
 * スプレッドシートの共有設定をチェックする
 * @param {string} spreadsheetId - チェックするスプレッドシートのID
 * @returns {object} status ('success' or 'error') と message
 */
function getSheetData(userId, sheetName, classFilter, sortMode, adminMode) {
  // キャッシュキー生成（ユーザー、シート、フィルタ条件ごとに個別キャッシュ）
  /** @type {string} キャッシュキー */
  const cacheKey = 'sheetData_' + userId + '_' + sheetName + '_' + classFilter + '_' + sortMode;

  // 管理モードの場合はキャッシュをバイパス（最新データを取得）
  if (adminMode === true) {
    return executeGetSheetData(userId, sheetName, classFilter, sortMode);
  }

  return cacheManager.get(cacheKey, () => {
    return executeGetSheetData(userId, sheetName, classFilter, sortMode);
  }, { ttl: 180 }); // 3分間キャッシュ（短縮してデータ整合性向上）
}

/**
 * 実際のシートデータ取得処理（キャッシュ制御から分離）
 */
function executeGetSheetData(userId, sheetName, classFilter, sortMode) {
    try {
      /** @type {Object} ユーザー情報 */
      const userInfo = getOrFetchUserInfo(userId, 'userId', {
    useExecutionCache: true,
    ttl: 300
  });
      if (!userInfo) {
        throw new Error('ユーザー情報が見つかりません');
      }

      /** @type {string} スプレッドシートID */
    const spreadsheetId = userInfo.spreadsheetId;
      /** @type {Object} Google Sheets APIサービス */
    const service = getSheetsServiceCached();

      // フォーム回答データのみを取得（名簿機能は使用しない）
      /** @type {Array} 範囲配列 */
      const ranges = [sheetName + '!A:Z'];

      /** @type {Object} APIレスポンス */
      const responses = batchGetSheetsData(service, spreadsheetId, ranges);
      /** @type {Array} シートデータ */
      const sheetData = responses.valueRanges[0].values || [];

    // 名簿機能は使用せず、空の配列を設定
    /** @type {Array} 名簿データ */
    const rosterData = [];
    if (sheetData.length === 0) {
      return {
        status: 'success',
        data: [],
        headers: [],
        totalCount: 0
      };
    }

    /** @type {Array} ヘッダー */
    const headers = sheetData[0];
    /** @type {Array} データ行 */
    const dataRows = sheetData.slice(1);

    // ヘッダーインデックスを取得（キャッシュ利用）
    /** @type {Object} ヘッダーインデックス */
    const headerIndices = getHeaderIndices(spreadsheetId, sheetName);

    // 名簿マップを作成（キャッシュ利用）
    /** @type {Object} 名簿マップ */
    const rosterMap = buildRosterMap(rosterData);

    // 表示モードとシート固有設定を取得
    /** @type {Object} 設定JSONオブジェクト */
    let configJson = {};
    try {
      configJson = JSON.parse(userInfo.configJson || '{}');
    } catch (parseError) {
      warnLog('ConfigJson parse error in executeGetSheetData:', parseError.message);
      configJson = {};
    }
    /** @type {string} 表示モード */
    const displayMode = configJson.displayMode || DISPLAY_MODES.ANONYMOUS;

    // シート固有の設定を取得（最新のAI判定結果を反映）
    /** @type {string} シートキー */
    const sheetKey = 'sheet_' + sheetName;
    /** @type {Object} シート固有の設定 */
    const sheetConfig = configJson[sheetKey] || {};

    // AI判定結果またはguessedConfigがある場合、それを優先使用
    /** @type {Object} 有効なヘッダー設定 */
    const effectiveHeaderConfig = sheetConfig.guessedConfig || sheetConfig || {};

    // フォールバック設定: 設定が空の場合はデフォルト構造を提供
    if (!effectiveHeaderConfig.opinionHeader && headers.length > 1) {
      effectiveHeaderConfig.opinionHeader = headers[1]; // 通常、タイムスタンプの次が最初の質問
    }
    if (!effectiveHeaderConfig.reasonHeader && headers.length > 2) {
      effectiveHeaderConfig.reasonHeader = headers[2]; // 2番目の質問を理由として設定
    }


    // Check if current user is the board owner
    /** @type {boolean} オーナーかどうか */
    const isOwner = (configJson.ownerId === userId);

    // データを処理
    /** @type {Array} 処理されたデータ */
    const processedData = dataRows.map(function(row, index) {
      return processRowData(row, headers, headerIndices, rosterMap, displayMode, index + 2, isOwner);
    });

    // フィルタリング
    /** @type {Array} フィルターされたデータ */
    let filteredData = processedData;
    if (classFilter && classFilter !== 'すべて') {
      /** @type {number|undefined} クラスインデックス */
      const classIndex = headerIndices[COLUMN_HEADERS.CLASS];
      if (classIndex !== undefined) {
        filteredData = processedData.filter(function(row) {
          return row.originalData[classIndex] === classFilter;
        });
      }
    }

    // ソート適用
    /** @type {Array} ソートされたデータ */
    const sortedData = applySortMode(filteredData, sortMode || 'newest');

    // カスタムフォーム設定がある場合のヘッダー情報を優先
    /** @type {Array} 有効なヘッダー */
    const effectiveHeaders = headers;
    /** @type {string} メイン質問ヘッダー */
    let mainQuestionHeader = headers[0]; // デフォルトは最初の列をメイン質問とする

    // AI判定結果またはシート設定からメイン質問を特定
    if (effectiveHeaderConfig.opinionHeader) {
      /** @type {number} 意見インデックス */
      const opinionIndex = headers.indexOf(effectiveHeaderConfig.opinionHeader);
      if (opinionIndex !== -1) {
        mainQuestionHeader = effectiveHeaderConfig.opinionHeader;
      }
    }

    return {
      status: 'success',
      data: sortedData,
      headers: effectiveHeaders,
      header: mainQuestionHeader, // フロントエンドが期待するメインヘッダー
      totalCount: sortedData.length,
      displayMode: displayMode,
      sheetName: sheetName,
      showCounts: configJson.showCounts || false,
      // デバッグ情報
      _sheetConfig: sheetConfig,
      _effectiveHeaderConfig: effectiveHeaderConfig
    };

  } catch (e) {
    errorLog('シートデータ取得エラー: ' + e.message);
    errorLog('Error stack: ' + e.stack);

    // データ取得失敗時のフォールバック処理
    try {
      /** @type {Object|null} ユーザー情報 */
    const userInfo = findUserById(userId);
      /** @type {Object} 設定JSONオブジェクト */
    let configJson = {};
      try {
        configJson = JSON.parse(userInfo.configJson || '{}');
      } catch (parseError) {
        warnLog('ConfigJson parse error in executeGetSheetData fallback:', parseError.message);
        configJson = {};
      }
      /** @type {string} シートキー */
    const sheetKey = 'sheet_' + sheetName;
      /** @type {Object} シート固有の設定 */
    const sheetConfig = configJson[sheetKey] || {};
      /** @type {Object} 有効なヘッダー設定 */
    const effectiveHeaderConfig = sheetConfig.guessedConfig || sheetConfig || {};

      // 設定からメインヘッダーを復元
      /** @type {string} フォールバックヘッダー */
      const fallbackHeader = effectiveHeaderConfig.opinionHeader || sheetName;

      warnLog('🔄 データ取得失敗 - フォールバック情報で応答: header=' + fallbackHeader);

      return {
        status: 'success', // フロントエンドエラーを避けるためsuccessを返す
        data: [],
        headers: [],
        header: fallbackHeader,
        totalCount: 0,
        displayMode: configJson.displayMode || 'anonymous',
        sheetName: sheetName,
        showCounts: configJson.showCounts || false,
        _error: 'データ取得に失敗しました: ' + e.message,
        _fallbackUsed: true
      };
    } catch (fallbackError) {
      errorLog('フォールバック処理も失敗: ' + fallbackError.message);

      return {
        status: 'error',
        message: 'データの取得に失敗しました: ' + e.message,
        data: [],
        headers: [],
        header: sheetName,
        totalCount: 0
      };
    }
  }
}

/**
 * シート一覧取得
 */
function getSheetsList(userId) {
  try {
    /** @type {Object|null} ユーザー情報 */
    const userInfo = findUserById(userId);
    if (!userInfo) {
      warnLog('getSheetsList: User not found:', userId);
      return [];
    }

    if (!userInfo.spreadsheetId || userInfo.spreadsheetId.trim() === '') {
      warnLog('getSheetsList: No spreadsheet ID for user:', userId);
      // 空のシートリストを返すが、ユーザーにスプレッドシート設定を促すメッセージも考慮
      return [];
    }
    /** @type {Object} Google Sheets APIサービス */
    const service = getSheetsServiceCached();
    if (!service) {
      errorLog('❌ getSheetsList: Sheets service not initialized');
      return [];
    }

    infoLog('✅ getSheetsList: Service validated successfully');
    /** @type {Object|undefined} スプレッドシート */
    let spreadsheet;
    try {
      spreadsheet = getSpreadsheetsData(service, userInfo.spreadsheetId);
    } catch (accessError) {
      /** @type {Array|null} ステータスマッチ */
      const statusMatch = (accessError && accessError.message || '').match(/Sheets API error:\s*(\d+)/);
      /** @type {number|null} ステータスコード */
      const statusCode = statusMatch ? parseInt(statusMatch[1], 10) : null;

      if (statusCode === 403) {
        warnLog('getSheetsList: アクセスエラーを検出。サービスアカウント権限を修復中...', accessError.message);

        // サービスアカウントの権限修復を試行
        try {
          addServiceAccountToSpreadsheet(userInfo.spreadsheetId);

          // 少し待ってから再試行
          Utilities.sleep(1000);
          spreadsheet = getSpreadsheetsData(service, userInfo.spreadsheetId);

        } catch (repairError) {
          errorLog('getSheetsList: 権限修復に失敗:', repairError.message);

          // 最終手段：ユーザー権限での修復も試行
          try {
            /** @type {string} 現在のユーザーメールアドレス */
  const currentUserEmail = getCurrentUserEmail();
            if (currentUserEmail === userInfo.adminEmail) {
              repairUserSpreadsheetAccess(currentUserEmail, userInfo.spreadsheetId);
            }
          } catch (finalRepairError) {
            errorLog('getSheetsList: 最終修復も失敗:', finalRepairError.message);
          }

          return [];
        }
      } else if (statusCode >= 500 && statusCode < 600) {
        /** @type {number} 最大リトライ数 */
        const maxRetries = 2;
        for (let attempt = 0; attempt < maxRetries; attempt++) {
          Utilities.sleep(Math.pow(2, attempt) * 1000);
          try {
            spreadsheet = getSpreadsheetsData(service, userInfo.spreadsheetId);
            break;
          } catch (retryError) {
            if (attempt === maxRetries - 1) {
              throw retryError;
            }
          }
        }
      } else {
        throw accessError;
      }
    }

    if (!spreadsheet) {
      errorLog('getSheetsList: No spreadsheet data returned');
      return [];
    }

    if (!spreadsheet.sheets) {
      errorLog('getSheetsList: Spreadsheet data missing sheets property. Available properties:', Object.keys(spreadsheet));
      return [];
    }

    if (!Array.isArray(spreadsheet.sheets)) {
      errorLog('getSheetsList: sheets property is not an array:', typeof spreadsheet.sheets);
      return [];
    }

    /** @type {Array} シート配列 */
    const sheets = spreadsheet.sheets.map(function(sheet) {
      if (!sheet.properties) {
        warnLog('getSheetsList: Sheet missing properties:', sheet);
        return null;
      }
      return {
        name: sheet.properties.title,
        id: sheet.properties.sheetId
      };
    }).filter(function(sheet) { return sheet !== null; });

    return sheets;
  } catch (e) {
    errorLog('getSheetsList: シート一覧取得エラー:', e.message);
    errorLog('getSheetsList: Error details:', e.stack);
    errorLog('getSheetsList: Error for userId:', userId);
    return [];
  }
}

/**
 * 名簿マップを構築（名簿機能無効化のため空のマップを返す）
 * @param {array} rosterData - 名簿データ（使用されません）
 * @returns {object} 空の名簿マップ
 */
function buildRosterMap(rosterData) {
  // 名簿機能は使用しないため、常に空のマップを返す
  // 名前はフォーム入力から直接取得
  return {};
}

/**
 * 行データを処理（スコア計算、名前変換など）
 */
function processRowData(row, headers, headerIndices, rosterMap, displayMode, rowNumber, isOwner) {
  /** @type {Object} 処理された行 */
  const processedRow = {
    rowNumber: rowNumber,
    originalData: row,
    score: 0,
    likeCount: 0,
    understandCount: 0,
    curiousCount: 0,
    isHighlighted: false
  };

  // リアクションカウント計算
  REACTION_KEYS.forEach(function(reactionKey) {
    /** @type {string} カラム名 */
    const columnName = COLUMN_HEADERS[reactionKey];
    /** @type {number|undefined} カラムインデックス */
    const columnIndex = headerIndices[columnName];

    if (columnIndex !== undefined && row[columnIndex]) {
      /** @type {Array} リアクション配列 */
      const reactions = parseReactionString(row[columnIndex]);
      /** @type {number} カウント */
      const count = reactions.length;

      switch (reactionKey) {
        case 'LIKE':
          processedRow.likeCount = count;
          break;
        case 'UNDERSTAND':
          processedRow.understandCount = count;
          break;
        case 'CURIOUS':
          processedRow.curiousCount = count;
          break;
      }
    }
  });

  // ハイライト状態チェック
  /** @type {number|undefined} ハイライトインデックス */
  const highlightIndex = headerIndices[COLUMN_HEADERS.HIGHLIGHT];
  if (highlightIndex !== undefined && row[highlightIndex]) {
    processedRow.isHighlighted = row[highlightIndex].toString().toLowerCase() === 'true';
  }

  // スコア計算
  processedRow.score = calculateRowScore(processedRow);

  // 名前の表示処理（フォーム入力の名前を使用）
  /** @type {number|undefined} 名前インデックス */
  const nameIndex = headerIndices[COLUMN_HEADERS.NAME];
  if (nameIndex !== undefined && row[nameIndex] && (displayMode === DISPLAY_MODES.NAMED || isOwner)) {
    processedRow.displayName = row[nameIndex];
  } else if (displayMode === DISPLAY_MODES.NAMED || isOwner) {
    // 名前入力がない場合のフォールバック
    processedRow.displayName = '匿名';
  }

  // 重要: コンテンツフィールドを追加（opinion, reason）
  // これらのフィールドは formatSheetDataForFrontend で使用される
  // 一番最初の列を意見として設定（通常フォームの最初の質問）
  processedRow.opinion = row.length > 1 ? (row[1] || '') : '';

  // 2番目の列を理由として設定（フォーム回答の2番目の項目）
  processedRow.reason = row.length > 2 ? (row[2] || '') : '';

  // メールアドレス（フォーム回答者の識別用）
  /** @type {number|undefined} メールインデックス */
  const emailIndex = headerIndices[COLUMN_HEADERS.EMAIL];
  if (emailIndex !== undefined && row[emailIndex]) {
    processedRow.email = row[emailIndex];
  }

  // タイムスタンプ（フォーム回答時刻）
  /** @type {number|undefined} タイムスタンプインデックス */
  const timestampIndex = headerIndices[COLUMN_HEADERS.TIMESTAMP];
  if (timestampIndex !== undefined && row[timestampIndex]) {
    processedRow.timestamp = row[timestampIndex];
  }

  return processedRow;
}

/**
 * 行のスコアを計算
 */
function calculateRowScore(rowData) {
  /** @type {number} ベーススコア */
  const baseScore = 1.0;

  // いいね！による加算
  /** @type {number} いいねボーナス */
  const likeBonus = rowData.likeCount * SCORING_CONFIG.LIKE_MULTIPLIER_FACTOR;

  // その他のリアクションも軽微な加算
  /** @type {number} リアクションボーナス */
  const reactionBonus = (rowData.understandCount + rowData.curiousCount) * 0.01;

  // ハイライトによる大幅加算
  /** @type {number} ハイライトボーナス */
  const highlightBonus = rowData.isHighlighted ? 0.5 : 0;

  // ランダム要素（同じスコアの項目をランダムに並べるため）
  /** @type {number} ランダムファクター */
  const randomFactor = Math.random() * SCORING_CONFIG.RANDOM_SCORE_FACTOR;

  return baseScore + likeBonus + reactionBonus + highlightBonus + randomFactor;
}

/**
 * データにソートを適用
 */
function applySortMode(data, sortMode) {
  switch (sortMode) {
    case 'score':
      return data.sort(function(a, b) { return b.score - a.score; });
    case 'newest':
      return data.reverse();
    case 'oldest':
      return data; // 元の順序（古い順）
    case 'random':
      return shuffleArray(data.slice()); // コピーをシャッフル
    case 'likes':
      return data.sort(function(a, b) { return b.likeCount - a.likeCount; });
    default:
      return data;
  }
}

/**
 * 配列をシャッフル（Fisher-Yates shuffle）
 */
function shuffleArray(array) {
  for (var i = array.length - 1; i > 0; i--) {
    /** @type {number} ランダムインデックス */
    const j = Math.floor(Math.random() * (i + 1));
    /** @type {*} 一時変数 */
    const temp = array[i];
    array[i] = array[j];
    array[j] = temp;
  }
  return array;
}

/**
 * リアクション文字列をパース
 */
function parseReactionString(val) {
  if (!val) return [];
  return val.toString().split(',').map(function(s) { return s.trim(); }).filter(Boolean);
}

/**
 * ヘルパー関数：ヘッダー配列から指定した名前のインデックスを取得
 * COLUMN_HEADERSと統一された方式を使用
 */
function getHeaderIndex(headers, headerName) {
  if (!headers || !headerName) return -1;
  return headers.indexOf(headerName);
}

/**
 * COLUMN_HEADERSキーから適切なヘッダー名を取得
 * @param {string} columnKey - COLUMN_HEADERSのキー（例：'OPINION', 'CLASS'）
 * @returns {string} ヘッダー名
 */
function getColumnHeaderName(columnKey) {
  return COLUMN_HEADERS[columnKey] || '';
}

/**
 * 統一されたヘッダーインデックス取得関数
 * @param {array} headers - ヘッダー配列
 * @param {string} columnKey - COLUMN_HEADERSのキー
 * @returns {number} インデックス（見つからない場合は-1）
 */

/**
 * 設定された列名と実際のスプレッドシートヘッダーをマッピング
 * @param {Object} configHeaders - 設定されたヘッダー名
 * @param {Object} actualHeaderIndices - 実際のヘッダーインデックスマップ
 * @returns {Object} マッピングされたインデックス
 */
function mapConfigToActualHeaders(configHeaders, actualHeaderIndices) {
  /** @type {Object} マッピングされたインデックス */
  const mappedIndices = {};
  /** @type {Array} 利用可能ヘッダー */
  const availableHeaders = Object.keys(actualHeaderIndices);

  // 各設定ヘッダーでマッピングを試行
  for (var configKey in configHeaders) {
    /** @type {string} 設定ヘッダー名 */
    const configHeaderName = configHeaders[configKey];
    /** @type {number|undefined} マッピングインデックス */
    let mappedIndex = undefined;
    if (configHeaderName && actualHeaderIndices.hasOwnProperty(configHeaderName)) {
      // 完全一致
      mappedIndex = actualHeaderIndices[configHeaderName];
    } else if (configHeaderName) {
      // 部分一致で検索（大文字小文字を区別しない）
      /** @type {string} 正規化された設定名 */
      const normalizedConfigName = configHeaderName.toLowerCase().trim();

      for (var actualHeader in actualHeaderIndices) {
        /** @type {string} 正規化された実際ヘッダー */
        const normalizedActualHeader = actualHeader.toLowerCase().trim();
        if (normalizedActualHeader === normalizedConfigName) {
          mappedIndex = actualHeaderIndices[actualHeader];
          break;
        }
      }

      // 部分一致検索
      if (mappedIndex === undefined) {
        for (var actualHeader in actualHeaderIndices) {
          /** @type {string} 正規化された実際ヘッダー */
        const normalizedActualHeader = actualHeader.toLowerCase().trim();
          if (normalizedActualHeader.includes(normalizedConfigName) || normalizedConfigName.includes(normalizedActualHeader)) {
            mappedIndex = actualHeaderIndices[actualHeader];
            break;
          }
        }
      }
    }

    // opinionHeader（メイン質問）の特別処理：見つからない場合は最も長い質問様ヘッダーを使用
    if (mappedIndex === undefined && configKey === 'opinionHeader') {
      /** @type {Array} 標準ヘッダー */
      const standardHeaders = ['タイムスタンプ', 'メールアドレス', 'クラス', '名前', '理由', 'なるほど！', 'いいね！', 'もっと知りたい！', 'ハイライト'];
      /** @type {Array} 質問ヘッダー */
      const questionHeaders = [];

      for (var header in actualHeaderIndices) {
        /** @type {boolean} 標準ヘッダーかどうか */
        let isStandardHeader = false;
        for (var i = 0; i < standardHeaders.length; i++) {
          if (header.toLowerCase().includes(standardHeaders[i].toLowerCase()) ||
              standardHeaders[i].toLowerCase().includes(header.toLowerCase())) {
            isStandardHeader = true;
            break;
          }
        }

        if (!isStandardHeader && header.length > 10) { // 質問は通常長い
          questionHeaders.push({header: header, index: actualHeaderIndices[header]});
        }
      }

      if (questionHeaders.length > 0) {
        // 最も長いヘッダーを選択（通常メイン質問が最も長い）
        /** @type {string} 最長ヘッダー */
        const longestHeader = questionHeaders.reduce(function(prev, current) {
          return (prev.header.length > current.header.length) ? prev : current;
        });
        mappedIndex = longestHeader.index;
      }
    }

    // reasonHeader（理由列）の特別処理：見つからない場合は理由らしいヘッダーを自動検出
    if (mappedIndex === undefined && configKey === 'reasonHeader') {
      /** @type {Array} 理由キーワード */
      const reasonKeywords = ['理由', 'reason', 'なぜ', 'why', '根拠', 'わけ'];
      for (var header in actualHeaderIndices) {
        /** @type {string} 正規化されたヘッダー */
        const normalizedHeader = header.toLowerCase().trim();
        for (var k = 0; k < reasonKeywords.length; k++) {
          if (normalizedHeader.includes(reasonKeywords[k]) || reasonKeywords[k].includes(normalizedHeader)) {
            mappedIndex = actualHeaderIndices[header];
            break;
          }
        }
        if (mappedIndex !== undefined) break;
      }

      // より広範囲の検索（部分一致）
      if (mappedIndex === undefined) {
        for (var header in actualHeaderIndices) {
          /** @type {string} 正規化されたヘッダー */
        const normalizedHeader = header.toLowerCase().trim();
          if (normalizedHeader.indexOf('理由') !== -1 || normalizedHeader.indexOf('reason') !== -1) {
            mappedIndex = actualHeaderIndices[header];
            break;
          }
        }
      }
    }

    // フォールバック: 設定が見つからない場合はデフォルト位置を使用
    if (mappedIndex === undefined) {
      if (configKey === 'opinionHeader') {
        // 意見列が見つからない場合、1番目の列（タイムスタンプの次）を使用
        mappedIndex = 1;
      } else if (configKey === 'reasonHeader') {
        // 理由列が見つからない場合、2番目の列を使用
        mappedIndex = 2;
      }
    }

    mappedIndices[configKey] = mappedIndex;

    if (mappedIndex === undefined) {
    }
  }

  return mappedIndices;
}

/**
 * 特定の行のリアクション情報を取得
 */
function getRowReactions(spreadsheetId, sheetName, rowIndex, userEmail) {
  try {
    /** @type {Object} Google Sheets APIサービス */
    const service = getSheetsServiceCached();
    /** @type {Object} ヘッダーインデックス */
    const headerIndices = getHeaderIndices(spreadsheetId, sheetName);

    /** @type {Object} リアクションデータ */
    const reactionData = {
      UNDERSTAND: { count: 0, reacted: false },
      LIKE: { count: 0, reacted: false },
      CURIOUS: { count: 0, reacted: false }
    };

    // 全リアクション列の範囲を一括で構築
    /** @type {Array} 範囲配列 */
    const ranges = [];
    /** @type {Array} 有効キー */
    const validKeys = [];
    REACTION_KEYS.forEach(function(reactionKey) {
      /** @type {string} カラム名 */
      const columnName = COLUMN_HEADERS[reactionKey];
      /** @type {number|undefined} カラムインデックス */
      const columnIndex = headerIndices[columnName];
      
      if (columnIndex !== undefined) {
        /** @type {string} 範囲 */
        const range = sheetName + '!' + String.fromCharCode(65 + columnIndex) + rowIndex;
        ranges.push(range);
        validKeys.push(reactionKey);
      }
    });

    // 1回のAPI呼び出しで全リアクション列データを取得
    if (ranges.length > 0) {
      /** @type {Object} APIレスポンス */
      const response = batchGetSheetsData(service, spreadsheetId, ranges);
      
      if (response && response.valueRanges) {
        response.valueRanges.forEach(function(valueRange, index) {
          /** @type {string} リアクションキー */
          const reactionKey = validKeys[index];
          /** @type {string} セル値 */
          let cellValue = '';
          
          if (valueRange && valueRange.values && valueRange.values[0] && valueRange.values[0][0]) {
            cellValue = valueRange.values[0][0];
          }
          
          if (cellValue) {
            /** @type {Array} リアクション配列 */
            const reactions = parseReactionString(cellValue);
            reactionData[reactionKey].count = reactions.length;
            reactionData[reactionKey].reacted = reactions.indexOf(userEmail) !== -1;
          }
        });
      }
    }

    return reactionData;
  } catch (e) {
    errorLog('getRowReactions エラー: ' + e.message);
    return {
      UNDERSTAND: { count: 0, reacted: false },
      LIKE: { count: 0, reacted: false },
      CURIOUS: { count: 0, reacted: false }
    };
  }
}

// =================================================================
// 追加のコアファンクション（Code.gsから移行）
// =================================================================

/**
 * 軽量な件数チェック（新着通知用）
 * 実際のデータではなく件数のみを返す
 */
/**
 * 回答ボードのデータを強制的に再読み込み
 */

// =================================================================
// ユーティリティ関数
// =================================================================

/**
 * 現在のユーザーのステータス情報を取得（AppSetupPage.htmlから呼び出される）
 * @returns {object} ユーザーのステータス情報
 */

/**
 * isActive状態を更新（AppSetupPage.htmlから呼び出される）
 * @param {boolean} isActive - 新しいisActive状態
 * @returns {object} 更新結果
 */
function updateIsActiveStatus(requestUserId, isActive) {
  verifyUserAccess(requestUserId);
  try {
    /** @type {string} アクティブユーザーメール */
    const activeUserEmail = getCurrentUserEmail();
    if (!activeUserEmail) {
      return {
        status: 'error',
        message: 'ユーザーが認証されていません'
      };
    }

    // 現在のユーザー情報を取得
    /** @type {Object} ユーザー情報 */
    const userInfo = findUserByEmail(activeUserEmail);
    if (!userInfo) {
      return {
        status: 'error',
        message: 'ユーザーがデータベースに登録されていません'
      };
    }

    // 編集者権限があるか確認（自分自身の状態変更も含む）
    if (!isTrue(userInfo.isActive)) {
      return {
        status: 'error',
        message: 'この操作を実行する権限がありません'
      };
    }

    // isActive状態を更新
    /** @type {string} 新しいアクティブ値 */
    const newIsActiveValue = isActive ? 'true' : 'false';
    /** @type {Object} 更新結果 */
    const updateResult = updateUser(userInfo.userId, {
      isActive: newIsActiveValue,
      lastAccessedAt: new Date().toISOString()
    });

    if (updateResult.success) {
      /** @type {string} ステータスメッセージ */
      const statusMessage = isActive
        ? 'アプリが正常に有効化されました'
        : 'アプリが正常に無効化されました';

      return {
        status: 'success',
        message: statusMessage,
        newStatus: newIsActiveValue
      };
    } else {
      return {
        status: 'error',
        message: 'データベースの更新に失敗しました'
      };
    }
  } catch (e) {
    errorLog('updateIsActiveStatus エラー: ' + e.message);
    return {
      status: 'error',
      message: 'isActive状態の更新に失敗しました: ' + e.message
    };
  }
}

/**
 * セットアップページへのアクセス権限を確認
 * @returns {boolean} アクセス権限があればtrue
 */
function hasSetupPageAccess() {
  try {
    /** @type {string} アクティブユーザーメール */
    const activeUserEmail = getCurrentUserEmail();
    if (!activeUserEmail) {
      return false;
    }

    // データベースに登録され、かつisActiveがtrueのユーザーのみアクセス可能
    /** @type {Object} ユーザー情報 */
    const userInfo = findUserByEmail(activeUserEmail);
    return userInfo && isTrue(userInfo.isActive);
  } catch (e) {
    errorLog('hasSetupPageAccess エラー: ' + e.message);
    return false;
  }
}

/**
 * メールアドレスからドメインを抽出
 * @param {string} email - メールアドレス
 * @returns {string} ドメイン部分
 */
function getEmailDomain(email) {
  return email.split('@')[1] || '';
}

/**
 * Drive APIサービスを取得
 * @returns {object} Drive APIサービス
 */
function getDriveService() {
  /** @type {string} アクセストークン */
  const accessToken = getServiceAccountTokenCached();
  return {
    accessToken: accessToken,
    baseUrl: 'https://www.googleapis.com/drive/v3',
    files: {
      get: function(params) {
        /** @type {string} URL */
        const url = this.baseUrl + '/files/' + params.fileId + '?fields=' + encodeURIComponent(params.fields);
        /** @type {Object} レスポンス */
        const response = UrlFetchApp.fetch(url, {
          headers: { 'Authorization': 'Bearer ' + this.accessToken }
        });
        return JSON.parse(response.getContentText());
      }
    }
  };
}

/**
 * 現在のアクセスユーザーがデバッグを有効にすべきかを判定
 * @returns {boolean} デバッグモードを有効にすべき場合はtrue
 */
function shouldEnableDebugMode() {
  try {
    // PropertiesServiceでDEBUG_MODEが有効に設定されているかをチェック
    const debugMode = PropertiesService.getScriptProperties().getProperty('DEBUG_MODE') === 'true';
    
    // デバッグモードが有効な場合のみtrueを返す
    return debugMode;
  } catch (error) {
    // エラー時は安全側に倒してfalseを返す
    warnLog('shouldEnableDebugMode error:', error.message);
    return false;
  }
}

/**
 * デプロイユーザーかどうか判定
 * 1. GASスクリプトの編集権限を確認
 * 2. 教育機関ドメインの管理者権限を確認
 * @returns {boolean} 管理者権限を持つ場合 true
 */
function isSystemAdmin() {
  try {
    /** @type {Object} プロパティ */
    const props = PropertiesService.getScriptProperties();
    /** @type {string} 管理者メール */
    const adminEmail = props.getProperty(SCRIPT_PROPS_KEYS.ADMIN_EMAIL);
    /** @type {string} 現在のユーザーメールアドレス */
  const currentUserEmail = getCurrentUserEmail();
    return adminEmail && currentUserEmail && adminEmail === currentUserEmail;
  } catch (e) {
    errorLog('isSystemAdmin エラー: ' + e.message);
    return false;
  }
}

function isDeployUser() {
  try {
    /** @type {Object} プロパティ */
    const props = PropertiesService.getScriptProperties();
    /** @type {string} 管理者メール */
    const adminEmail = props.getProperty(SCRIPT_PROPS_KEYS.ADMIN_EMAIL);
    /** @type {string} 現在のユーザーメールアドレス */
  const currentUserEmail = getCurrentUserEmail();
    
    // デバッグログ
    if (shouldEnableDebugMode()) {
      infoLog('isDeployUser check - adminEmail: ' + adminEmail);
      infoLog('isDeployUser check - currentUserEmail: ' + currentUserEmail);
      infoLog('isDeployUser check - result: ' + (adminEmail === currentUserEmail));
    }
    
    return adminEmail && currentUserEmail && adminEmail === currentUserEmail;
  } catch (e) {
    errorLog('isDeployUser エラー: ' + e.message);
    return false;
  }
}

// =================================================================
// マルチテナント対応関数（同時アクセス対応）
// =================================================================

/**
 * マルチテナント対応: 設定保存・公開
 * @param {string} requestUserId - リクエストされたユーザーID（第一引数）
 * @param {Object} settingsData - 設定データ
 * @returns {Object} 実行結果
 */

/**
 * マルチテナント対応: アクティブフォーム情報取得
 * @param {string} requestUserId - リクエストされたユーザーID
 * @returns {Object} フォーム情報
 */
/**
 * 管理者によるユーザー削除（AppSetupPage.html用ラッパー）
 * @param {string} targetUserId - 削除対象ユーザーID
 * @param {string} reason - 削除理由
 */
async function deleteUserAccountByAdminForUI(targetUserId, reason) {
  try {
    const result = await deleteUserAccountByAdmin(targetUserId, reason);
    return {
      status: 'success',
      message: result.message,
      deletedUser: result.deletedUser
    };
  } catch (error) {
    errorLog('deleteUserAccountByAdmin wrapper error:', error.message);
    return {
      status: 'error',
      message: error.message
    };
  }
}

/**
 * 削除ログ取得（AppSetupPage.html用ラッパー）
 */
function getDeletionLogsForUI() {
  try {
    const logs = getDeletionLogs();
    return {
      status: 'success',
      logs: logs
    };
  } catch (error) {
    errorLog('getDeletionLogs wrapper error:', error.message);
    return {
      status: 'error',
      message: error.message
    };
  }
}

/**
 * 診断ログ一覧を取得（UI用ラッパー）
 * @param {number} limit - 取得する件数の上限
 */
function getDiagnosticLogsForUI(limit) {
  try {
    const logs = getDiagnosticLogs(limit || 50);
    return {
      status: 'success',
      logs: logs
    };
  } catch (error) {
    errorLog('getDiagnosticLogs wrapper error:', error.message);
    return {
      status: 'error',
      message: error.message
    };
  }
}

function getAllUsersForAdminForUI(requestUserId) {
  try {
    const result = getAllUsersForAdmin();
    return {
      status: 'success',
      users: result
    };
  } catch (error) {
    errorLog('getAllUsersForAdminForUI wrapper error:', error.message);
    return {
      status: 'error',
      message: error.message
    };
  }
}

/**
 * カスタムセットアップの実行
 * @param {string} requestUserId リクエストユーザーID
 * @param {Object} config セットアップ設定
 * @returns {Object} 実行結果
 */
function executeCustomSetup(requestUserId, config) {
  try {
    // アクセス権限確認
    if (!verifyUserAccess(requestUserId)) {
      throw new Error('アクセス権限がありません');
    }

    // 設定検証
    if (!config || typeof config !== 'object') {
      throw new Error('設定情報が不正です');
    }

    // 必要な設定項目の確認
    const requiredFields = ['spreadsheetId', 'sheetName'];
    for (const field of requiredFields) {
      if (!config[field]) {
        throw new Error('必要な設定項目が不足しています: ' + field);
      }
    }

    /** @type {Object} ユーザー情報 */
    const userInfo = findUserById(requestUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    // ステップ1: スプレッドシートアクセス確認
    try {
      const spreadsheet = SpreadsheetApp.openById(config.spreadsheetId);
      if (!spreadsheet) {
        throw new Error('指定されたスプレッドシートにアクセスできません');
      }
    } catch (accessError) {
      throw new Error('スプレッドシートアクセスエラー: ' + accessError.message);
    }

    // ステップ2: ヘッダー自動マッピング
    /** @type {Object} ヘッダーマッピング結果 */
    const headerMapping = autoMapHeaders(config.headers || [], config.sheetName);
    
    if (!headerMapping || !headerMapping.success) {
      throw new Error('ヘッダーの自動マッピングに失敗しました');
    }

    // ステップ3: 設定の構築
    /** @type {Object} シート設定 */
    const sheetConfig = {
      timestampHeader: headerMapping.mapping.timestamp || 'タイムスタンプ',
      classHeader: headerMapping.mapping.class || 'クラス',
      nameHeader: headerMapping.mapping.name || '名前',
      emailHeader: headerMapping.mapping.email || 'メールアドレス',
      opinionHeader: headerMapping.mapping.opinion || '回答',
      reasonHeader: headerMapping.mapping.reason || '理由',
      guessedConfig: headerMapping.mapping,
      lastModified: new Date().toISOString()
    };

    // ステップ4: 設定保存
    const saveResult = saveSheetConfig(
      requestUserId,
      config.spreadsheetId,
      config.sheetName,
      sheetConfig,
      { updatePublished: true }
    );

    if (!saveResult || !saveResult.success) {
      throw new Error('設定の保存に失敗しました');
    }

    // ステップ5: キャッシュクリア
    if (typeof clearExecutionUserInfoCache === 'function') {
      clearExecutionUserInfoCache();
    }
    
    if (typeof cacheManager !== 'undefined' && cacheManager.clearByPattern) {
      cacheManager.clearByPattern('publishedData_' + requestUserId + '_');
    }

    // ステップ6: 成功レスポンス
    logInfo('カスタムセットアップが完了しました', {
      userId: requestUserId,
      spreadsheetId: config.spreadsheetId,
      sheetName: config.sheetName
    });

    return {
      success: true,
      message: 'カスタムセットアップが完了しました',
      data: {
        spreadsheetId: config.spreadsheetId,
        sheetName: config.sheetName,
        headerMapping: headerMapping.mapping,
        setupTimestamp: new Date().toISOString()
      }
    };

  } catch (error) {
    logError(error, 'executeCustomSetup', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.SYSTEM);
    
    return {
      success: false,
      error: error.message,
      message: 'カスタムセットアップに失敗しました'
    };
  }
}

/**
 * 管理者によるユーザーのisActive制御（AppSetupPage.html用ラッパー）
 * @param {string} targetUserId - 対象ユーザーID
 * @param {boolean} isActive - 設定する状態
 * @returns {Object} 実行結果
 */
function updateUserActiveStatusForUI(targetUserId, isActive) {
  try {
    // 権限チェック: システム管理者のみ許可
    if (!isSystemAdmin()) {
      return {
        status: 'error',
        message: 'この操作を実行する権限がありません'
      };
    }

    if (!targetUserId) {
      return { status: 'error', message: 'ユーザーIDが指定されていません' };
    }

    /** @type {string} 新しいアクティブ値 */
    const newIsActiveValue = isActive ? 'true' : 'false';
    /** @type {Object} 更新結果 */
    const updateResult = updateUser(targetUserId, {
      isActive: newIsActiveValue,
      lastAccessedAt: new Date().toISOString()
    });

    if (updateResult && updateResult.status === 'success') {
      return {
        status: 'success',
        message: isActive ? 'ユーザーを有効化しました' : 'ユーザーを無効化しました',
        newStatus: newIsActiveValue
      };
    } else {
      return { status: 'error', message: 'データベースの更新に失敗しました' };
    }
  } catch (error) {
    errorLog('updateUserActiveStatusForUI エラー:', error);
    return {
      status: 'error',
      message: error.message || String(error)
    };
  }
}

/**
 * 統合カスタムセットアップ処理（QuickStart同様の完全自動化）
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @param {object} config - カスタムフォーム設定
 * @returns {object} セットアップ結果
 */
function customSetup(requestUserId, config) {
  verifyUserAccess(requestUserId);
  try {

    // ステップ1: ユーザー入力の取得と検証
    const formTitle = (config && config.formTitle ? String(config.formTitle).trim() : '');
    if (!formTitle) {
      throw new Error('formTitleは必須です');
    }
    const sheetName = (config && config.sheetName ? String(config.sheetName).trim() : '');

    const userInfo = findUserById(requestUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    const userEmail = userInfo.adminEmail;
    let configJson = {};
    try {
      configJson = JSON.parse(userInfo.configJson || '{}');
    } catch (parseError) {
      warnLog('ConfigJson parse error in executeCustomSetup:', parseError.message);
      configJson = {};
    }

    // ステップ2: Googleフォームとスプレッドシートの作成
    const formAndSsInfo = createCustomFormAndSheet(userEmail, requestUserId, config);

    // ステップ3: AI列判定の実行
    const aiDetectionResult = performAutoAIDetection(requestUserId, formAndSsInfo.spreadsheetId, formAndSsInfo.sheetName);

    // ステップ4: ユーザーデータベースレコードの統合更新
    const sheetKey = 'sheet_' + formAndSsInfo.sheetName;
    configJson.formCreated = true;
    configJson.formUrl = formAndSsInfo.formUrl;
    configJson.editFormUrl = formAndSsInfo.editFormUrl;
    configJson.publishedSpreadsheetId = formAndSsInfo.spreadsheetId;
    configJson.publishedSheetName = formAndSsInfo.sheetName;
    configJson.setupStatus = 'completed';
    configJson.lastModified = new Date().toISOString();
    configJson[sheetKey] = configJson[sheetKey] || {};
    if (aiDetectionResult && aiDetectionResult.guessedConfig) {
      const guessed = aiDetectionResult.guessedConfig;
      configJson[sheetKey].opinionHeader = guessed.opinionHeader || '';
      configJson[sheetKey].nameHeader = guessed.nameHeader || '';
      configJson[sheetKey].classHeader = guessed.classHeader || '';
      configJson[sheetKey].reasonHeader = guessed.reasonHeader || '';
    }

    const updateResult = updateUser(requestUserId, {
      spreadsheetId: formAndSsInfo.spreadsheetId,
      spreadsheetUrl: formAndSsInfo.spreadsheetUrl,
      configJson: JSON.stringify(configJson),
      lastAccessedAt: new Date().toISOString()
    });
    if (updateResult.status !== 'success') {
      throw new Error('データベース保存に失敗: ' + updateResult.message);
    }

    // ステップ5: キャッシュのクリア
    clearExecutionUserInfoCache();
    invalidateUserCache(requestUserId, userEmail, formAndSsInfo.spreadsheetId, true);

    // ステップ6: 結果の返却
    const appUrls = generateUserUrls(requestUserId);
    return {
      status: 'success',
      message: 'カスタムセットアップが完了しました',
      webAppUrl: appUrls.webAppUrl,
      adminUrl: appUrls.adminUrl,
      viewUrl: appUrls.viewUrl,
      setupUrl: appUrls.setupUrl,
      formUrl: formAndSsInfo.formUrl,
      editFormUrl: formAndSsInfo.editFormUrl,
      spreadsheetUrl: formAndSsInfo.spreadsheetUrl,
      aiDetectedHeaders: aiDetectionResult.guessedConfig || {},
      sheetName: formAndSsInfo.sheetName,
      spreadsheetId: formAndSsInfo.spreadsheetId
    };

  } catch (e) {
    errorLog('❌ customSetup エラー: ' + e.message);

    // エラー時はセットアップ状態をリセット
    try {
      const userInfo = findUserById(requestUserId);
      if (userInfo) {
        const currentConfig = JSON.parse(userInfo.configJson || '{}');
        currentConfig.setupStatus = 'error';
        currentConfig.lastError = e.message;
        currentConfig.errorAt = new Date().toISOString();

        updateUser(requestUserId, { configJson: JSON.stringify(currentConfig) });
        invalidateUserCache(requestUserId, userInfo.adminEmail || userInfo.email, null, false);
        clearExecutionUserInfoCache();
      }
    } catch (updateError) {
      errorLog('エラー状態の更新に失敗: ' + updateError.message);
    }

    return {
      status: 'error',
      message: 'カスタムセットアップに失敗しました: ' + e.message,
      webAppUrl: '',
      adminUrl: '',
      viewUrl: '',
      setupUrl: ''
    };
  }
}
/**
 * カスタムフォームとスプレッドシート作成の統合処理
 * @param {string} userEmail - ユーザーのメールアドレス
 * @param {string} requestUserId - ユーザーID
 * @param {object} config - カスタム設定
 * @returns {object} 作成されたファイル情報
 */
function createCustomFormAndSheet(userEmail, requestUserId, config) {
  const convertedConfig = {
    mainQuestion: {
      title: config.mainQuestion || 'あなたの考えや気づいたことを教えてください',
      type: config.responseType || config.questionType || 'text',
      choices: config.questionChoices || config.choices || [],
      includeOthers: config.includeOthers || false
    },
    enableClass: config.enableClass || false,
    classQuestion: {
      choices: config.classChoices || ['クラス1', 'クラス2', 'クラス3', 'クラス4']
    }
  };

  const overrides = { customConfig: convertedConfig };
  if (config.formTitle) {
    overrides.formTitle = config.formTitle;
  } else {
    overrides.titlePrefix = 'カスタムフォーム';
  }

  const formAndSsInfo = createUnifiedForm('custom', userEmail, requestUserId, overrides);

  if (config.sheetName && config.sheetName !== formAndSsInfo.sheetName) {
    try {
      const ss = SpreadsheetApp.openById(formAndSsInfo.spreadsheetId);
      const sheet = ss.getSheetByName(formAndSsInfo.sheetName);
      if (sheet) {
        sheet.setName(config.sheetName);
        formAndSsInfo.sheetName = config.sheetName;
      }
    } catch (e) {
      warnLog('シート名の変更に失敗しました: ' + e.message);
    }
  }

  return formAndSsInfo;
}

/**
 * 高精度AI列判定の自動実行
 * @param {string} requestUserId - ユーザーID
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {object} AI判定結果
 */
function performAutoAIDetection(requestUserId, spreadsheetId, sheetName) {
  try {

    const sheetDetails = getSheetDetails(requestUserId, spreadsheetId, sheetName);

    if (!sheetDetails || !sheetDetails.guessedConfig) {
      throw new Error('AI列判定の実行に失敗しました');
    }

    infoLog('✅ AI列判定自動実行完了', {
      opinionHeader: sheetDetails.guessedConfig.opinionHeader,
      nameHeader: sheetDetails.guessedConfig.nameHeader,
      classHeader: sheetDetails.guessedConfig.classHeader
    });

    return {
      success: true,
      aiDetected: true,
      guessedConfig: sheetDetails.guessedConfig,
      allHeaders: sheetDetails.allHeaders,
      message: 'AI列判定が正常に完了しました'
    };

  } catch (error) {
    errorLog('❌ AI列判定自動実行エラー', { error: error.message });
    return {
      success: false,
      aiDetected: false,
      message: 'AI列判定に失敗しました: ' + error.message,
      error: error.message
    };
  }
}

/**
 * カスタムフォーム作成（AdminPanel.html用ラッパー）
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @param {object} config - フォーム設定
 */
function createCustomFormUI(requestUserId, config) {
  try {
    verifyUserAccess(requestUserId);
    const activeUserEmail = getCurrentUserEmail();

    // AdminPanelのconfig構造を内部形式に変換（createCustomForm の処理を統合）
    const convertedConfig = {
      mainQuestion: {
        title: config.mainQuestion || 'あなたの考えや気づいたことを教えてください',
        type: config.responseType || config.questionType || 'text', // responseTypeを優先して使用
        choices: config.questionChoices || config.choices || [], // questionChoicesを優先して使用
        includeOthers: config.includeOthers || false
      },
      enableClass: config.enableClass || false,
      classQuestion: {
        choices: config.classChoices || ['クラス1', 'クラス2', 'クラス3', 'クラス4']
      }
    };
    const overrides = { customConfig: convertedConfig };
    if (config.formTitle) {
      overrides.formTitle = config.formTitle;
    } else {
      overrides.titlePrefix = 'カスタムフォーム';
    }

    const result = createUnifiedForm('custom', activeUserEmail, requestUserId, overrides);

    // 新規追加: カスタムフォーム作成後のシートアクティベーション
    if (result.sheetName) {
      try {
        const activeSheetResult = setActiveSheet(requestUserId, result.sheetName);
        infoLog('✅ カスタムフォーム作成後のシートアクティベーション完了:', activeSheetResult);
      } catch (sheetError) {
        warnLog('⚠️ シートアクティベーション失敗（処理継続）:', sheetError.message);
        // シートアクティベーション失敗時のログ詳細化
        errorLog('シートアクティベーション失敗詳細:', {
          requestUserId: requestUserId,
          sheetName: result.sheetName,
          error: sheetError.message
        });
      }
    } else {
      warnLog('⚠️ createUnifiedForm結果にsheetNameが含まれていません:', result);
    }

    // ユーザー専用フォルダを作成・取得してファイルを移動
    let folder = null;
    const moveResults = { form: false, spreadsheet: false };
    try {
      folder = createUserFolder(activeUserEmail);

      if (folder) {
        // フォームファイルを移動
        try {
          const formFile = DriveApp.getFileById(result.formId);
          if (formFile) {
            // 既にフォルダに存在するかチェック
            const formParents = formFile.getParents();
            let isFormAlreadyInFolder = false;

            while (formParents.hasNext()) {
              if (formParents.next().getId() === folder.getId()) {
                isFormAlreadyInFolder = true;
                break;
              }
            }

            if (!isFormAlreadyInFolder) {
              // 推奨メソッドmoveTo()を使用してファイル移動
              formFile.moveTo(folder);
              moveResults.form = true;
              infoLog('✅ カスタムフォームファイル移動完了: %s (ID: %s)', folder.getName(), formFile.getId());
            }
          }
        } catch (formMoveError) {
          warnLog('カスタムフォームファイル移動エラー:', formMoveError.message);
        }

        // スプレッドシートファイルを移動
        try {
          const ssFile = DriveApp.getFileById(result.spreadsheetId);
          if (ssFile) {
            // 既にフォルダに存在するかチェック
            const ssParents = ssFile.getParents();
            let isSsAlreadyInFolder = false;

            while (ssParents.hasNext()) {
              if (ssParents.next().getId() === folder.getId()) {
                isSsAlreadyInFolder = true;
                break;
              }
            }

            if (!isSsAlreadyInFolder) {
              // 推奨メソッドmoveTo()を使用してファイル移動
              ssFile.moveTo(folder);
              moveResults.spreadsheet = true;
              infoLog('✅ カスタムスプレッドシートファイル移動完了: %s (ID: %s)', folder.getName(), ssFile.getId());
            }
          }
        } catch (ssMoveError) {
          warnLog('カスタムスプレッドシートファイル移動エラー:', ssMoveError.message);
        }

      }
    } catch (folderError) {
      warnLog('カスタムセットアップフォルダ処理エラー:', folderError.message);
    }

    // 既存ユーザーの情報を更新（スプレッドシート情報を追加）
    const existingUser = findUserById(requestUserId);
    if (existingUser) {

      const updatedConfigJson = JSON.parse(existingUser.configJson || '{}');
      updatedConfigJson.formUrl = result.formUrl;
      updatedConfigJson.editFormUrl = result.editFormUrl || result.formUrl;
      updatedConfigJson.formCreated = true;
      updatedConfigJson.lastFormCreatedAt = new Date().toISOString();
      updatedConfigJson.setupStatus = 'completed';
      updatedConfigJson.appPublished = true;
      // publishedSheetName と publishedSpreadsheetId の安全な設定
      if (result.spreadsheetId && typeof result.spreadsheetId === 'string') {
        updatedConfigJson.publishedSpreadsheetId = result.spreadsheetId;
        infoLog('✅ publishedSpreadsheetId設定完了:', result.spreadsheetId);
      } else {
        errorLog('❌ 無効なspreadsheetId:', result.spreadsheetId);
        throw new Error('フォーム作成は成功しましたが、スプレッドシートIDが無効です');
      }

      if (result.sheetName && typeof result.sheetName === 'string' && result.sheetName.trim() !== '' && result.sheetName !== 'true') {
        updatedConfigJson.publishedSheetName = result.sheetName;
        infoLog('✅ publishedSheetName設定完了:', result.sheetName);
      } else {
        errorLog('❌ 無効なsheetName:', result.sheetName);
        throw new Error('フォーム作成は成功しましたが、シート名が無効です: ' + result.sheetName);
      }
      updatedConfigJson.folderId = folder ? folder.getId() : '';
      updatedConfigJson.folderUrl = folder ? folder.getUrl() : '';

      // カスタムフォーム設定情報を保存
      // カスタムフォーム設定情報をシート固有のキーの下に保存
      const sheetKey = 'sheet_' + result.sheetName;
      updatedConfigJson[sheetKey] = {
        ...(updatedConfigJson[sheetKey] || {}), // 既存のシート設定を保持
        formTitle: config.formTitle,
        mainQuestion: config.mainQuestion,
        questionType: config.questionType,
        choices: config.choices,
        includeOthers: config.includeOthers,
        enableClass: config.enableClass,
        classChoices: config.classChoices,
        formUrl: result.formUrl, // シート固有のフォームURL保存
        editFormUrl: result.editFormUrl || result.formUrl, // 編集用URL保存
        lastModified: new Date().toISOString()
      };

      // 以前の実行で誤ってトップレベルに追加された可能性のあるプロパティを削除
      delete updatedConfigJson.formTitle;
      delete updatedConfigJson.mainQuestion;
      delete updatedConfigJson.questionType;
      delete updatedConfigJson.choices;
      delete updatedConfigJson.includeOthers;
      delete updatedConfigJson.enableClass;
      delete updatedConfigJson.classChoices;

      // 新しく作成されたスプレッドシート情報をメインのユーザー情報として更新
      const updateData = {
        spreadsheetId: result.spreadsheetId,
        spreadsheetUrl: result.spreadsheetUrl,
        configJson: JSON.stringify(updatedConfigJson),
        lastAccessedAt: new Date().toISOString()
      };
      updateUser(requestUserId, updateData);

      // カスタムフォーム作成後の包括的キャッシュ同期（Quick Startと同様）
      synchronizeCacheAfterCriticalUpdate(requestUserId, activeUserEmail, existingUser.spreadsheetId, result.spreadsheetId);
    } else {
      warnLog('createCustomFormUI - user not found:', requestUserId);
    }

    return {
      status: 'success',
      message: 'カスタムフォームが正常に作成されました！',
      formUrl: result.formUrl,
      spreadsheetUrl: result.spreadsheetUrl,
      formTitle: result.formTitle,
      sheetName: result.sheetName,
      folderId: folder ? folder.getId() : '',
      folderUrl: folder ? folder.getUrl() : '',
      filesMovedToFolder: moveResults
    };
  } catch (error) {
    errorLog('createCustomFormUI error:', error.message);
    return {
      status: 'error',
      message: error.message
    };
  }
}

/**
 * クイックスタート用フォーム作成（UI用ラッパー）
 * @param {string} requestUserId - リクエスト元のユーザーID
 */
function createQuickStartFormUI(requestUserId) {
  try {
    verifyUserAccess(requestUserId);
    const activeUserEmail = getCurrentUserEmail();

    // createQuickStartForm の処理を統合（直接 createUnifiedForm を呼び出し）
    const result = createUnifiedForm('quickstart', activeUserEmail, requestUserId);

    // QuickStart作成後のシートアクティベーション
    if (result.sheetName) {
      try {
        const activeSheetResult = setActiveSheet(requestUserId, result.sheetName);
        infoLog('✅ QuickStart作成後のシートアクティベーション完了:', activeSheetResult);
      } catch (sheetError) {
        warnLog('⚠️ QuickStartシートアクティベーション失敗（処理継続）:', sheetError.message);
      }
    } else {
      warnLog('⚠️ QuickStart createUnifiedForm結果にsheetNameが含まれていません:', result);
    }

    // 既存ユーザーの情報を更新
    const existingUser = findUserById(requestUserId);
    if (existingUser) {
      const updatedConfigJson = JSON.parse(existingUser.configJson || '{}');
      updatedConfigJson.formUrl = result.viewFormUrl || result.formUrl;
      updatedConfigJson.editFormUrl = result.editFormUrl;
      updatedConfigJson.formCreated = true;
      updatedConfigJson.setupStatus = 'completed';
      updatedConfigJson.appPublished = true;

      // QuickStart用のpublishedSheetNameとpublishedSpreadsheetIdの安全な設定
      if (result.spreadsheetId && typeof result.spreadsheetId === 'string') {
        updatedConfigJson.publishedSpreadsheetId = result.spreadsheetId;
        infoLog('✅ QuickStart publishedSpreadsheetId設定完了:', result.spreadsheetId);
      } else {
        errorLog('❌ QuickStart 無効なspreadsheetId:', result.spreadsheetId);
      }

      if (result.sheetName && typeof result.sheetName === 'string' && result.sheetName.trim() !== '' && result.sheetName !== 'true') {
        updatedConfigJson.publishedSheetName = result.sheetName;
        infoLog('✅ QuickStart publishedSheetName設定完了:', result.sheetName);
      } else {
        errorLog('❌ QuickStart 無効なsheetName:', result.sheetName);
      }

      const updateData = {
        spreadsheetId: result.spreadsheetId,
        spreadsheetUrl: result.spreadsheetUrl,
        configJson: JSON.stringify(updatedConfigJson),
        lastAccessedAt: new Date().toISOString()
      };

      updateUser(requestUserId, updateData);
    }

    return {
      status: 'success',
      message: 'クイックスタートフォームが正常に作成されました！',
      formUrl: result.formUrl,
      spreadsheetUrl: result.spreadsheetUrl,
      formTitle: result.formTitle
    };
  } catch (error) {
    errorLog('createQuickStartFormUI error:', error.message);
    return {
      status: 'error',
      message: error.message
    };
  }
}

async function deleteCurrentUserAccount(requestUserId) {
  try {
    if (!requestUserId) {
      throw new Error('認証エラー: ユーザーIDが指定されていません');
    }
    verifyUserAccess(requestUserId);
    const result = await deleteUserAccount(requestUserId);

    return {
      status: 'success',
      message: 'アカウントが正常に削除されました',
      result: result
    };
  } catch (error) {
    errorLog('deleteCurrentUserAccount error:', error.message);
    return {
      status: 'error',
      message: error.message
    };
  }
}

/**
 * シートを有効化（AdminPanel.html用のシンプル版）
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @param {string} sheetName - シート名
 */
function activateSheetSimple(requestUserId, sheetName) {
  try {
    verifyUserAccess(requestUserId);
    const userInfo = findUserById(requestUserId);
    if (!userInfo || !userInfo.spreadsheetId) {
      throw new Error('ユーザー情報またはスプレッドシートIDが見つかりません');
    }

    return activateSheet(requestUserId, userInfo.spreadsheetId, sheetName);
  } catch (error) {
    errorLog('activateSheetSimple error:', error.message);
    return {
      status: 'error',
      message: error.message
    };
  }
}

/**
 * ログインフロー修正の検証テスト
 * @returns {object} テスト結果
 */

/**
 * セキュリティ強化とキャッシュ最適化のパフォーマンステスト
 * @returns {object} パフォーマンステスト結果
 */
/**
 * 高信頼性ログイン状態確認（段階的検索戦略）
 * @returns {Object} Login status result
 */
function getLoginStatus() {
  const startTime = Date.now();
  let activeUserEmail = null;
  
  try {
    activeUserEmail = getCurrentUserEmail();
    if (!activeUserEmail) {
      return { status: 'error', message: 'ログインユーザーの情報を取得できませんでした。' };
    }
    
    // 段階的検索戦略（書き込み直後の読み取り一貫性を考慮）
    let userInfo = null;
    let searchSuccess = false;
    let searchAttempts = [];

    // Stage 1: キャッシュ確認（短期間のみ）
    const cacheKey = 'login_status_' + activeUserEmail;
    try {
      const cached = CacheService.getScriptCache().get(cacheKey);
      if (cached) {
        const parsedCache = JSON.parse(cached);
        // キャッシュが古い場合（書き込み直後等）は無視
        const cacheAge = Date.now() - (parsedCache._timestamp || 0);
        if (cacheAge < 10000) { // 10秒以内のキャッシュのみ使用
          return parsedCache;
        } else {
        }
      }
    } catch (cacheError) {
    }

    // Stage 2: 通常検索（キャッシュ付き）
    try {
      userInfo = findUserByEmail(activeUserEmail);
      searchAttempts.push({ method: 'findUserByEmail', success: !!userInfo });
      if (userInfo) {
        searchSuccess = true;
      }
    } catch (stage2Error) {
      searchAttempts.push({ method: 'findUserByEmail', error: stage2Error.message });
      warnLog('getLoginStatus: Stage 2失敗:', stage2Error.message);
    }

    // Stage 3: 強制フレッシュ検索（データベース直接）
    if (!searchSuccess) {
      try {
        userInfo = fetchUserFromDatabase('adminEmail', activeUserEmail, { 
          forceFresh: true, 
          retryOnce: true 
        });
        searchAttempts.push({ method: 'fetchUserFromDatabase_forceFresh', success: !!userInfo });
        if (userInfo) {
          searchSuccess = true;
        }
      } catch (stage3Error) {
        searchAttempts.push({ method: 'fetchUserFromDatabase_forceFresh', error: stage3Error.message });
        warnLog('getLoginStatus: Stage 3失敗:', stage3Error.message);
      }
    }

    // Stage 4: 最終確認検索（複数方法で確認）
    if (!searchSuccess) {
      const finalMethods = [
        () => fetchUserFromDatabase('adminEmail', activeUserEmail, { forceFresh: true }),
        () => {
          // 少し待ってから再度検索（書き込み直後の場合）
          Utilities.sleep(1000);
          return fetchUserFromDatabase('adminEmail', activeUserEmail, { forceFresh: true });
        }
      ];

      for (let i = 0; i < finalMethods.length && !searchSuccess; i++) {
        try {
          userInfo = finalMethods[i]();
          const methodName = 'final_method_' + (i + 1);
          searchAttempts.push({ method: methodName, success: !!userInfo });
          if (userInfo) {
            searchSuccess = true;
            break;
          }
        } catch (finalError) {
          searchAttempts.push({ method: 'final_method_' + (i + 1), error: finalError.message });
        }
      }
    }

    const searchElapsed = Date.now() - startTime;
    
    // 検索結果の処理
    let result;
    
    if (userInfo && searchSuccess) {
      // ユーザーが見つかった場合の状態判定
      const isActive = (userInfo.isActive === true || 
                       userInfo.isActive === 'true' || 
                       String(userInfo.isActive).toLowerCase() === 'true');
      
      const urls = generateUserUrls(userInfo.userId);
      
      if (isActive) {
        result = {
          status: 'existing_user',
          userId: userInfo.userId,
          adminUrl: urls.adminUrl,
          viewUrl: urls.viewUrl,
          message: 'ログインが完了しました',
          _timestamp: Date.now(),
          _searchElapsed: searchElapsed
        };
        infoLog('✅ getLoginStatus: アクティブユーザー確認完了', {
          email: activeUserEmail,
          userId: userInfo.userId,
          elapsed: searchElapsed + 'ms',
          attempts: searchAttempts.length
        });
      } else {
        result = {
          status: 'setup_required',
          userId: userInfo.userId,
          adminUrl: urls.adminUrl,
          viewUrl: urls.viewUrl,
          message: 'セットアップを完了してください',
          _timestamp: Date.now(),
          _searchElapsed: searchElapsed
        };
        warnLog('⚠️ getLoginStatus: 非アクティブユーザー検出', {
          email: activeUserEmail,
          userId: userInfo.userId
        });
      }
    } else {
      // ユーザーが見つからない場合
      result = { 
        status: 'unregistered', 
        userEmail: activeUserEmail,
        _timestamp: Date.now(),
        _searchElapsed: searchElapsed,
        searchAttempts: searchAttempts
      };
      warnLog('⚠️ getLoginStatus: ユーザーが見つかりません', {
        email: activeUserEmail,
        attemptCount: searchAttempts.length,
        elapsed: searchElapsed + 'ms',
        attempts: searchAttempts
      });
    }
    
    // 結果をキャッシュ（成功時は長め、失敗時は短め）
    try {
      const cacheTtl = searchSuccess ? 60 : 15; // 成功時60秒、失敗時15秒
      CacheService.getScriptCache().put(cacheKey, JSON.stringify(result), cacheTtl);
    } catch (cacheError) {
      // キャッシュ保存エラーは無視
    }
    
    return result;

  } catch (error) {
    errorLog('❌ getLoginStatus 重大エラー:', {
      error: error.message,
      email: activeUserEmail || 'unknown',
      elapsed: (Date.now() - startTime) + 'ms'
    });
    return { 
      status: 'error', 
      message: error.message,
      _timestamp: Date.now(),
      _searchElapsed: Date.now() - startTime
    };
  }
}

/**
 * 統合ユーザー登録確認（高信頼性フロー）
 * @returns {Object} Registration result
 */
function confirmUserRegistration() {
  const startTime = Date.now();
  
  try {
    const activeUserEmail = getCurrentUserEmail();
    if (!activeUserEmail) {
      return { status: 'error', message: 'ユーザー情報を取得できませんでした。' };
    }
    
    infoLog('confirmUserRegistration: 統合登録確認開始', { email: activeUserEmail });
    
    // 関連キャッシュのクリア（登録確認時は新鮮な状態で開始）
    try {
      const cache = CacheService.getScriptCache();
      cache.remove('login_status_' + activeUserEmail);
      cache.remove('email_' + activeUserEmail);
    } catch (cacheError) {
      // キャッシュクリアの失敗は無視
    }

    // 段階的な既存ユーザー確認戦略
    let existingUser = null;
    let userFound = false;
    const searchStages = [
      { name: 'immediate', method: () => findUserByEmail(activeUserEmail) },
      { name: 'forceFresh', method: () => fetchUserFromDatabase('adminEmail', activeUserEmail, { forceFresh: true }) },
      { name: 'retryFresh', method: () => {
        Utilities.sleep(500); // 短い待機後再検索
        return fetchUserFromDatabase('adminEmail', activeUserEmail, { forceFresh: true, retryOnce: true });
      }}
    ];

    for (const stage of searchStages) {
      if (userFound) break;
      
      try {
        existingUser = stage.method();
        if (existingUser && existingUser.userId && existingUser.adminEmail) {
          userFound = true;
          infoLog('✅ confirmUserRegistration: 既存ユーザー発見', {
            stage: stage.name,
            userId: existingUser.userId,
            email: activeUserEmail
          });
          break;
        }
      } catch (stageError) {
        warnLog('confirmUserRegistration: 検索ステージエラー', {
          stage: stage.name,
          error: stageError.message
        });
      }
    }
    
    if (existingUser && userFound) {
      // 既存ユーザーの処理
      const isActive = (existingUser.isActive === true || 
                       String(existingUser.isActive).toLowerCase() === 'true');
      const urls = generateUserUrls(existingUser.userId);
      
      // ログイン状態を適切にキャッシュ
      try {
        const loginStatus = {
          status: isActive ? 'existing_user' : 'setup_required',
          userId: existingUser.userId,
          _timestamp: Date.now()
        };
        CacheService.getScriptCache().put('login_status_' + activeUserEmail, JSON.stringify(loginStatus), 300);
      } catch (cacheError) {
        // キャッシュ設定失敗は無視
      }
      
      const elapsedTime = Date.now() - startTime;
      infoLog('✅ confirmUserRegistration: 既存ユーザー確認完了', {
        email: activeUserEmail,
        userId: existingUser.userId,
        isActive: isActive,
        elapsed: elapsedTime + 'ms'
      });
      
      return {
        status: isActive ? 'existing_user' : 'setup_required',
        userId: existingUser.userId,
        adminUrl: urls.adminUrl,
        viewUrl: urls.viewUrl,
        message: isActive ? '既に登録済みのため、管理パネルへ移動します' : 'セットアップを完了してください',
        _processTime: elapsedTime
      };
    }
    
    // 新規ユーザー登録の実行
    
    try {
      const registrationResult = registerNewUser(activeUserEmail);
      
      // 登録成功の場合の追加検証
      if (registrationResult && registrationResult.userId) {
        // 登録直後の検証（登録が確実に完了したことを確認）
        try {
          const verificationUser = fetchUserFromDatabase('userId', registrationResult.userId, { 
            forceFresh: true 
          });
          
          if (!verificationUser) {
            warnLog('confirmUserRegistration: 登録後検証で見つからない:', registrationResult.userId);
          } else {
          }
        } catch (verifyError) {
          warnLog('confirmUserRegistration: 登録後検証でエラー:', verifyError.message);
        }
        
        // ログイン状態の適切なキャッシュ設定
        try {
          const loginStatus = {
            status: 'existing_user',
            userId: registrationResult.userId,
            _timestamp: Date.now()
          };
          CacheService.getScriptCache().put('login_status_' + activeUserEmail, JSON.stringify(loginStatus), 300);
        } catch (cacheError) {
          // キャッシュ設定失敗は無視
        }
      }
      
      const totalElapsed = Date.now() - startTime;
      infoLog('✅ confirmUserRegistration: 新規ユーザー登録完了', {
        email: activeUserEmail,
        userId: registrationResult?.userId,
        elapsed: totalElapsed + 'ms'
      });
      
      // 結果に処理時間を追加
      if (registrationResult && typeof registrationResult === 'object') {
        registrationResult._processTime = totalElapsed;
      }
      
      return registrationResult;
      
    } catch (registrationError) {
      const totalElapsed = Date.now() - startTime;
      errorLog('❌ confirmUserRegistration: 登録処理エラー', {
        email: activeUserEmail,
        error: registrationError.message,
        elapsed: totalElapsed + 'ms'
      });
      
      return { 
        status: 'error', 
        message: registrationError.message,
        _processTime: totalElapsed
      };
    }
    
  } catch (error) {
    const totalElapsed = Date.now() - startTime;
    errorLog('❌ confirmUserRegistration: 重大エラー', {
      email: activeUserEmail || 'unknown',
      error: error.message,
      elapsed: totalElapsed + 'ms'
    });
    
    return { 
      status: 'error', 
      message: error.message,
      _processTime: totalElapsed
    };
  }
}

// =================================================================
// 統合API（フェーズ2最適化）
// =================================================================

/**
 * 統合初期データ取得API - OPTIMIZED
 * 従来の5つのAPI呼び出し（getCurrentUserStatus, getUserId, getAppConfig, getSheetDetails）を統合
 * @param {string} requestUserId - リクエスト元のユーザーID（省略可能）
 * @param {string} targetSheetName - 詳細を取得するシート名（省略可能）
 * @param {boolean} lightweightMode - 軽量モード（シート詳細取得をスキップ）
 * @returns {Object} 統合された初期データ
 */
function getInitialData(requestUserId, targetSheetName, lightweightMode) {
  try {
    /** @type {number} 開始時間 */
    const startTime = new Date().getTime();

    // === ステップ1: ユーザー認証とユーザー情報取得（キャッシュ活用） ===
    /** @type {string} アクティブユーザーメール */
    const activeUserEmail = getCurrentUserEmail();
    /** @type {string} 現在のユーザーID */
    const currentUserId = requestUserId;

    // UserID の解決
    if (!currentUserId) {
      currentUserId = getUserId();
    }

    // Phase3 Optimization: Use execution-level cache to avoid duplicate database queries
    clearExecutionUserInfoCache(); // Clear any stale cache

    // 軽量モード時またはキャッシュバイパス時の追加キャッシュクリア
    if (lightweightMode || targetSheetName === 'BYPASS_CACHE') {
      try {
        // 統一キャッシュクリアを実行
        if (typeof performUnifiedCacheClear === 'function') {
          performUnifiedCacheClear(currentUserId, activeUserEmail, null, 'execution');
        }
        // データベースキャッシュも強制クリア
        if (typeof clearDatabaseCache === 'function') {
          clearDatabaseCache();
        }
      } catch (cacheError) {
        // エラーは無視
      }
    }

    // ユーザー認証
    verifyUserAccess(currentUserId);
    
    // 軽量モードまたは強制更新時は、確実に最新データを取得
    /** @type {Object|undefined} ユーザー情報 */
    let userInfo;
    if (lightweightMode || targetSheetName === 'BYPASS_CACHE') {
      userInfo = findUserByIdFresh(currentUserId);
      if (!userInfo) {
        throw new Error('ユーザー情報が見つかりません（強制更新）');
      }
    } else {
      userInfo = getOrFetchUserInfo(currentUserId, 'userId', {
        useExecutionCache: true,
        ttl: 300
      });
      if (!userInfo) {
        throw new Error('ユーザー情報が見つかりません');
      }
    }

    // === ステップ1.5: データ整合性の自動チェックと修正 ===
    try {
      /** @type {Object} 整合性結果 */
      const consistencyResult = fixUserDataConsistency(currentUserId);
      if (consistencyResult.updated) {
        // 修正後は最新データを再取得
        clearExecutionUserInfoCache();
        userInfo = getOrFetchUserInfo(currentUserId, 'userId', {
          useExecutionCache: true,
          ttl: 300
        });
      }
    } catch (consistencyError) {
      // エラーが発生しても初期化処理は続行
    }

    // === ステップ2: 設定データの取得と自動修復 ===
    /** @type {Object} 設定JSONオブジェクト */
    let configJson = {};
    try {
      configJson = JSON.parse(userInfo.configJson || '{}');
    } catch (parseError) {
      warnLog('ConfigJson parse error, using empty config:', parseError.message);
      configJson = {};
    }

    // --- 統一された自動修復システム ---
    const healingResult = performAutoHealing(userInfo, configJson, currentUserId);
    if (healingResult.updated) {
      configJson = healingResult.configJson;
      userInfo.configJson = JSON.stringify(configJson);
    }

    // === ステップ3: シート一覧とアプリURL生成 ===
    /** @type {Array} シートリスト */
    const sheets = getSheetsList(currentUserId);
    /** @type {Object} アプリケーションURL群 */
    const appUrls = generateUserUrls(currentUserId);

    // === ステップ4: 回答数とリアクション数の取得 ===
    /** @type {number} 回答数 */
    let answerCount = 0;
    /** @type {number} 総リアクション数 */
    let totalReactions = 0;
    try {
      if (configJson.publishedSpreadsheetId && configJson.publishedSheetName) {
        /** @type {Object} レスポンスデータ */
        const responseData = getResponsesData(currentUserId, configJson.publishedSheetName);
        if (responseData.status === 'success') {
          answerCount = responseData.data.length;
          totalReactions = answerCount * 2; // 暫定値
        }
      }
    } catch (err) {
      warnLog('Answer count retrieval failed:', err.message);
    }

    // === ステップ5: セットアップステップの決定 ===
    /** @type {number} セットアップステップ */
    let setupStep = 1;
    try {
      setupStep = getSetupStep(userInfo, configJson);
    } catch (stepError) {
      warnLog('setupStep決定でエラー、デフォルト値(1)を使用:', stepError.message);
      setupStep = 1;
    }

    // 公開シート設定とヘッダー情報を取得
    /** @type {string} 公開シート名 */
    const publishedSheetName = configJson.publishedSheetName || '';
    /** @type {string} シート設定キー */
    const sheetConfigKey = publishedSheetName ? 'sheet_' + publishedSheetName : '';
    /** @type {Object|undefined} アクティブシート設定 */
    const activeSheetConfig = sheetConfigKey && configJson[sheetConfigKey]
      ? configJson[sheetConfigKey]
      : {};

    /** @type {string} 意見ヘッダー */
    const opinionHeader = activeSheetConfig.opinionHeader || '';
    /** @type {string} 名前ヘッダー */
    const nameHeader = activeSheetConfig.nameHeader || '';
    /** @type {string} クラスヘッダー */
    const classHeader = activeSheetConfig.classHeader || '';

    // === ベース応答の構築 ===
    /** @type {Object} レスポンスオブジェクト */
    const response = {
      // ユーザー情報
      userInfo: {
        userId: userInfo.userId,
        adminEmail: userInfo.adminEmail,
        isActive: userInfo.isActive,
        lastAccessedAt: userInfo.lastAccessedAt,
        spreadsheetId: userInfo.spreadsheetId,
        spreadsheetUrl: userInfo.spreadsheetUrl,
        configJson: userInfo.configJson
      },
      // アプリ設定
      appUrls: appUrls,
      setupStep: setupStep,
      activeSheetName: configJson.publishedSheetName || null,
      webAppUrl: appUrls.webApp,
      isPublished: !!configJson.appPublished,
      answerCount: answerCount,
      totalReactions: totalReactions,
      config: {
        publishedSheetName: publishedSheetName,
        opinionHeader: opinionHeader,
        nameHeader: nameHeader,
        classHeader: classHeader,
        showNames: configJson.showNames || false,
        showCounts: configJson.showCounts !== undefined ? configJson.showCounts : false,
        displayMode: configJson.displayMode || 'anonymous',
        setupStatus: configJson.setupStatus || 'initial',
        isPublished: !!configJson.appPublished,
      },
      // シート情報
      allSheets: sheets,
      sheetNames: sheets,
      // カスタムフォーム情報
      customFormInfo: configJson.formUrl ? {
        title: configJson.formTitle || 'カスタムフォーム',
        mainQuestion: configJson.mainQuestion || opinionHeader || configJson.publishedSheetName || '質問',
        formUrl: configJson.formUrl
      } : null,
      // メタ情報
      _meta: {
        apiVersion: 'integrated_v1',
        executionTime: null,
        includedApis: ['getCurrentUserStatus', 'getUserId', 'getAppConfig']
      }
    };

    // === ステップ6: シート詳細の取得（オプション）- 最適化版 ===
    // 軽量モード時はシート詳細の取得をスキップ
    /** @type {boolean} シート詳細を含むか */
    const includeSheetDetails = !lightweightMode && (targetSheetName || configJson.publishedSheetName);

    // publishedSheetNameが空の場合のフォールバック処理
    if (!includeSheetDetails && userInfo.spreadsheetId && configJson) {
      warnLog('⚠️ シート名が指定されていません。デフォルトシート名を検索中...');
      try {
        // 一般的なシート名パターンを試す
        const commonSheetNames = ['フォームの回答 1', 'フォーム回答 1', 'Form Responses 1', 'Sheet1', 'シート1'];
        const tempService = getSheetsServiceCached();
        const spreadsheetInfo = getSpreadsheetsData(tempService, userInfo.spreadsheetId);

        if (spreadsheetInfo && spreadsheetInfo.sheets && spreadsheetInfo.sheets.length > 0) {
          // 既知のシート名から最初に見つかったものを使用
          for (const commonName of commonSheetNames) {
            if (spreadsheetInfo.sheets.some(sheet => sheet.properties.title === commonName)) {
              includeSheetDetails = commonName;
              infoLog('✅ フォールバックシート名を使用:', commonName);
              break;
            }
          }

          // それでも見つからない場合は最初のシートを使用
          if (!includeSheetDetails) {
            includeSheetDetails = spreadsheetInfo.sheets[0].properties.title;
            infoLog('✅ 最初のシートを使用:', includeSheetDetails);
          }
        }
      } catch (fallbackError) {
        warnLog('⚠️ フォールバックシート名検索に失敗:', fallbackError.message);
      }
    }

    if (includeSheetDetails && userInfo.spreadsheetId) {
      try {
        // 最適化: getSheetsServiceの重複呼び出しを避けるため、一度だけ作成して再利用
        /** @type {Object} 共有シーツサービス */
        const sharedSheetsService = getSheetsServiceCached();

        // ExecutionContext を最適化版で作成（sheetsService と userInfo を渡して重複作成を回避）
        const context = createExecutionContext(currentUserId, {
          reuseService: sharedSheetsService,
          reuseUserInfo: userInfo
        });
        /** @type {Object} シート詳細 */
        const sheetDetails = getSheetDetailsFromContext(context, userInfo.spreadsheetId, includeSheetDetails);
        response.sheetDetails = sheetDetails;
        response._meta.includedApis.push('getSheetDetails');
        infoLog('✅ シート詳細を統合応答に追加 (最適化版):', includeSheetDetails);
        // getInitialData内で生成したcontextの変更をコミット
        commitAllChanges(context);
      } catch (sheetErr) {
        warnLog('Sheet details retrieval failed:', sheetErr.message);
        response.sheetDetailsError = sheetErr.message;
      }
    }

    // === 実行時間の記録 ===
    /** @type {number} 終了時間 */
    const endTime = new Date().getTime();
    response._meta.executionTime = endTime - startTime;

    return response;

  } catch (error) {
    errorLog('❌ getInitialData error:', error);
    
    // エラーレスポンスにも最低限のuserInfo構造を含める
    let fallbackUserInfo = null;
    try {
      const activeUserEmail = getCurrentUserEmail();
      if (activeUserEmail) {
        fallbackUserInfo = {
          userId: requestUserId || 'unknown',
          adminEmail: activeUserEmail,
          isActive: false,
          lastAccessedAt: null,
          spreadsheetId: '',
          spreadsheetUrl: '',
          configJson: '{}'
        };
      }
    } catch (userError) {
      errorLog('❌ Fallback userInfo creation failed:', userError);
    }
    
    return {
      status: 'error',
      message: error.message,
      data: {
        setupStep: 1,
        appUrls: null,
        activeSheetName: null
      },
      userInfo: fallbackUserInfo,
      timestamp: new Date().toISOString(),
      _meta: {
        apiVersion: 'integrated_v1',
        error: error.message,
        includedApis: ['error_fallback'],
        executionTime: 0
      }
    };
  }
}

/**
 * 手動データ整合性修正（管理パネル用）
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @returns {object} 修正結果
 */
function fixDataConsistencyManual(requestUserId) {
  try {
    verifyUserAccess(requestUserId);

    /** @type {Object} 結果 */
    const result = fixUserDataConsistency(requestUserId);

    if (result.status === 'success') {
      if (result.updated) {
        return {
          status: 'success',
          message: 'データ整合性の問題を修正しました。ページを再読み込みしてください。',
          details: result.message
        };
      } else {
        return {
          status: 'success',
          message: 'データ整合性に問題は見つかりませんでした。',
          details: result.message
        };
      }
    } else {
      return {
        status: 'error',
        message: 'データ整合性修正中にエラーが発生しました: ' + result.message
      };
    }

  } catch (error) {
    errorLog('❌ 手動データ整合性修正エラー:', error);
    return {
      status: 'error',
      message: '修正処理中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * getSheetDetailsの内部実装（統合API用）
 * @param {string} requestUserId - ユーザーID
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} シート詳細
 */

/**
 * アプリケーション状態取得（UI用）
 * @returns {Object} アプリケーション状態情報
 */
function getApplicationStatusForUI() {
  try {
    const accessCheck = checkApplicationAccess();
    const isEnabled = getApplicationEnabled();
    const adminEmail = getCurrentUserEmail();

    return {
      status: 'success',
      isEnabled: isEnabled,
      isSystemAdmin: accessCheck.isSystemAdmin,
      adminEmail: adminEmail,
      lastUpdated: new Date().toISOString(),
      message: accessCheck.accessReason
    };
  } catch (error) {
    errorLog('getApplicationStatusForUI エラー:', error);
    return {
      status: 'error',
      message: error.message
    };
  }
}

/**
 * アプリケーション状態設定（UI用）
 * @param {boolean} enabled - 有効化するかどうか
 * @returns {Object} 設定結果
 */
function setApplicationStatusForUI(enabled) {
  try {
    const result = setApplicationEnabled(enabled);
    return {
      status: 'success',
      enabled: result.enabled,
      message: result.message,
      timestamp: result.timestamp,
      adminEmail: result.adminEmail
    };
  } catch (error) {
    errorLog('setApplicationStatusForUI エラー:', error);
    return {
      status: 'error',
      message: error.message
    };
  }
}

/**
 * エラーバウンダリからのテスト用関数
 * ErrorBoundaryで呼び出される関数の実装
 * @returns {Object} テスト結果
 */
function testForceLogoutRedirect() {
  try {
    return {
      status: 'success',
      message: 'テスト関数が正常に動作しています',
      timestamp: new Date().toISOString(),
      function: 'testForceLogoutRedirect'
    };
  } catch (error) {
    errorLog('testForceLogoutRedirect エラー:', error);
    return {
      status: 'error',
      message: error.message,
      timestamp: new Date().toISOString()
    };
  }
}

/**
 * フロー進捗を取得する関数（BackendProgressSync用）
 * @param {string} flowId - フローID（例: 'customSetup', 'quickStart'）
 * @returns {Object} 進捗情報
 */
function getFlowProgress(flowId) {
  try {
    
    // 現在はシンプルなスタブ実装
    // 実際のバックエンド処理状況を取得するロジックを将来実装
    const mockProgress = {
      flowId: flowId,
      progress: 50, // 0-100の進捗率
      currentStep: flowId === 'customSetup' ? 3 : 2,
      totalSteps: flowId === 'customSetup' ? 7 : 4,
      stepName: flowId === 'customSetup' ? 'AI列判定実行中' : 'フォーム作成中',
      stepDetail: '処理中...',
      estimatedTimeRemaining: 120,
      isComplete: false,
      timestamp: new Date().toISOString()
    };
    
    return {
      status: 'success',
      data: mockProgress
    };
  } catch (error) {
    logError(error, 'getFlowProgress', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.SYSTEM, { flowId });
    return {
      status: 'error',
      message: '進捗取得に失敗しました: ' + error.message,
      flowId: flowId
    };
  }
}

/**
 * スプレッドシートの作成日を取得
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {string|null} 作成日のISO文字列、取得失敗時はnull
 */
function getSpreadsheetCreatedDate(spreadsheetId) {
  try {
    if (!spreadsheetId) {
      warnLog('⚠️ getSpreadsheetCreatedDate: spreadsheetIdが指定されていません');
      return null;
    }
    
    const file = DriveApp.getFileById(spreadsheetId);
    const createdDate = file.getDateCreated();
    
    return createdDate.toISOString();
    
  } catch (error) {
    warnLog('⚠️ スプレッドシート作成日取得失敗:', error.message);
    return null;
  }
}

/**
 * ユーザーのスプレッドシート作成日を取得（管理パネル用API）
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @returns {object} API応答
 */
function getSpreadsheetCreatedDateAPI(requestUserId) {
  try {
    verifyUserAccess(requestUserId);
    
    const userInfo = getUserInfo(requestUserId);
    if (!userInfo || !userInfo.spreadsheetId) {
      return {
        status: 'error',
        message: 'スプレッドシートが見つかりません'
      };
    }
    
    const createdDate = getSpreadsheetCreatedDate(userInfo.spreadsheetId);
    
    return {
      status: 'success',
      createdDate: createdDate,
      spreadsheetId: userInfo.spreadsheetId
    };
    
  } catch (error) {
    errorLog('❌ getSpreadsheetCreatedDateAPI エラー:', error);
    return {
      status: 'error',
      message: '作成日取得中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * API function to update user data (e.g., switch active spreadsheet)
 * @param {string} requestUserId - Requesting user ID
 * @param {Object} updateData - Data to update (e.g., {spreadsheetId: "..."})
 * @returns {Object} API response with status
 */
function updateUserAPI(requestUserId, updateData) {
  try {
    // Verify user access
    verifyUserAccess(requestUserId);
    
    // Validate input data
    if (!updateData || typeof updateData !== 'object') {
      return {
        status: 'error',
        message: '更新データが無効です'
      };
    }
    
    // Allow only safe fields to be updated
    const allowedFields = ['spreadsheetId'];
    const filteredUpdateData = {};
    
    for (const field of allowedFields) {
      if (updateData.hasOwnProperty(field)) {
        filteredUpdateData[field] = updateData[field];
      }
    }
    
    if (Object.keys(filteredUpdateData).length === 0) {
      return {
        status: 'error',
        message: '更新可能なフィールドが含まれていません'
      };
    }
    
    // Perform the update
    updateUser(requestUserId, filteredUpdateData);
    
    // Clear relevant caches after spreadsheet switch
    if (filteredUpdateData.spreadsheetId) {
      // Clear user-specific cache
      clearExecutionUserInfoCache();
      // Clear any sheet-specific cache
      try {
        invalidateUserCache(requestUserId, null, null, 'all', null);
      } catch (cacheError) {
        warnLog('Cache invalidation warning:', cacheError.message);
      }
      
      // CRITICAL: Verify that the update was actually persisted
      const maxVerificationAttempts = 3;
      let verificationSuccess = false;
      
      for (let attempt = 1; attempt <= maxVerificationAttempts; attempt++) {
        try {
          // Small delay to allow for database consistency
          Utilities.sleep(200 * attempt);
          
          // Force fresh read from database
          clearExecutionUserInfoCache();
          const verifiedUserInfo = findUserByIdFresh(requestUserId);
          
          if (verifiedUserInfo && verifiedUserInfo.spreadsheetId === filteredUpdateData.spreadsheetId) {
            verificationSuccess = true;
            break;
          } else {
            warnLog('⚠️ Update verification failed on attempt ' + attempt + ':', {
              expected: filteredUpdateData.spreadsheetId,
              actual: verifiedUserInfo ? verifiedUserInfo.spreadsheetId : 'null',
              userInfo: !!verifiedUserInfo
            });
          }
        } catch (verificationError) {
          warnLog('⚠️ Update verification error on attempt ' + attempt + ':', verificationError.message);
        }
      }
      
      if (!verificationSuccess) {
        errorLog('❌ CRITICAL: Update verification failed after ' + maxVerificationAttempts + ' attempts');
        return {
          status: 'error',
          message: 'データベース更新の検証に失敗しました。管理者にお問い合わせください。'
        };
      }
    }
    
    return {
      status: 'success',
      message: 'ユーザー情報が正常に更新されました',
      updatedFields: Object.keys(filteredUpdateData)
    };
    
  } catch (error) {
    errorLog('❌ updateUserAPI エラー:', error);
    return {
      status: 'error',
      message: 'ユーザー情報更新中にエラーが発生しました: ' + error.message
    };
  }
}
