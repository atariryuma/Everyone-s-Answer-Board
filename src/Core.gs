/**
 * @fileoverview StudyQuest - Core Functions (最適化版)
 * 主要な業務ロジックとAPI エンドポイント
 */

// デバッグログ関数の定義（テスト環境対応）
if (typeof debugLog === 'undefined') {
  function debugLog(message, ...args) {
    console.log('[DEBUG]', message, ...args);
  }
}

// Import standardized error handling functions
if (typeof logError === 'undefined') {
  throw new Error('errorHandler.gs must be loaded before Core.gs');
}

if (typeof warnLog === 'undefined') {
  function warnLog(message, ...args) {
    console.warn('[WARN]', message, ...args);
  }
}

if (typeof infoLog === 'undefined') {
  function infoLog(message, ...args) {
    console.log('[INFO]', message, ...args);
  }
}

/**
 * セットアップステップを判定する（サーバー側統一実装）
 * フロントエンドとバックエンドで共通のステップ判定ロジックを提供
 * @param {Object} userInfo - ユーザー情報
 * @param {Object} configJson - 設定JSON（オブジェクト形式）
 * @returns {number} setupStep (1-3)
 */
function getSetupStep(userInfo, configJson) {
  // Step 1: データソース未設定
  if (!userInfo || !userInfo.spreadsheetId || userInfo.spreadsheetId.trim() === '') {
    return 1;
  }
  
  // configが存在しない場合は必ずStep 2
  if (!configJson || typeof configJson !== 'object') {
    return 2;
  }
  
  // 公開済み状態の最優先判定（データ不整合に関係なく公開済みならStep 3）
  const isCurrentlyPublished = (
    configJson.appPublished === true ||
    (configJson.setupStatus === 'completed' && 
     configJson.formCreated === true && 
     configJson.formUrl && configJson.formUrl.trim())
  );
  
  if (isCurrentlyPublished) {
    return 3;
  }
  
  // デフォルト: セットアップ継続中
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
    debugLog('clearActiveSheet開始: userId=' + requestUserId);

    const userInfo = getConfigUserInfo(requestUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません。');
    }

    const configJson = JSON.parse(userInfo.configJson || '{}');

    debugLog('🔍 公開停止前の設定:', {
      publishedSheetName: configJson.publishedSheetName,
      publishedSpreadsheetId: configJson.publishedSpreadsheetId,
      appPublished: configJson.appPublished
    });

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
    if (debugMode) debugLog('🔧 setupStep統一判定: Step 1 - データソース未設定');
    return 1;
  }

  // configJsonが存在しない場合は必ずStep 2
  if (!configJson || typeof configJson !== 'object') {
    if (debugMode) debugLog('🔧 setupStep統一判定: Step 2 - configJson未設定');
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
    if (debugMode) {
      debugLog('🔧 setupStep統一判定: Step 2 - セットアップ未完了', {
        setupStatus,
        formCreated,
        hasFormUrl
      });
    }
    return 2;
  }

  // Step 3: セットアップ完了（すべての条件をクリア）
  if (setupStatus === 'completed' && formCreated && hasFormUrl) {
    if (debugMode) debugLog('🔧 setupStep統一判定: Step 3 - セットアップ完了');
    return 3;
  }

  // フォールバック: 不明な状態はStep 2として扱う
  if (debugMode) {
    debugLog('🔧 setupStep統一判定: フォールバック - Step 2', {
      setupStatus,
      formCreated,
      hasFormUrl
    });
  }
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
      errors.push(`必須フィールド '${field}' が未定義です`);
    } else if (typeof config[field] !== expectedType) {
      errors.push(`フィールド '${field}' の型が不正です。期待値: ${expectedType}, 実際の値: ${typeof config[field]}`);
    }
  }

  // 文字列フィールドの特別な検証
  if (config.publishedSheetName === 'true') {
    errors.push("publishedSheetNameが不正な値 'true' になっています");
  }

  // setupStatusの値チェック
  const validSetupStatuses = ['pending', 'completed', 'error', 'reconfiguring'];
  if (config.setupStatus && !validSetupStatuses.includes(config.setupStatus)) {
    errors.push(`setupStatusの値が不正です: ${config.setupStatus}`);
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
      changes.push(`setupStatus: ${configJson.setupStatus} → completed (form作成済み)`);
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
        logError(`Auto-healing後の状態が無効: ${validation.errors}`, 'autoHealConfig', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.DATABASE);
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
        debugLog('📋 Auto-healing実行:', changes.join(', '));
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
    debugLog('🔧 ExecutionLevel SheetsService: 初回作成');
    _executionSheetsServiceCache = getSheetsService();
  } else {
    debugLog('♻️ ExecutionLevel SheetsService: キャッシュから取得');
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
    debugLog('🔍 Starting header integrity validation for userId:', userId);

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

    const config = JSON.parse(userInfo.configJson || '{}');
    const sheetConfigKey = 'sheet_' + (config.publishedSheetName || sheetName);
    const sheetConfig = config[sheetConfigKey] || {};

    const opinionHeader = sheetConfig.opinionHeader || config.publishedSheetName || 'お題';

    debugLog('getOpinionHeaderSafely:', {
      userId: userId,
      sheetName: sheetName,
      opinionHeader: opinionHeader
    });

    return opinionHeader;
  } catch (e) {
    logError(e, 'getOpinionHeaderSafely', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.SYSTEM);
    return 'お題';
  }
}

/**
 * 新規ユーザーを登録する（データベース登録のみ）
 * フォーム作成はクイックスタートで実行される
 */
function registerNewUser(adminEmail) {
  const activeUser = Session.getActiveUser();
  if (adminEmail !== activeUser.getEmail()) {
    throw new Error('認証エラー: 操作を実行しているユーザーとメールアドレスが一致しません。');
  }

  // ドメイン制限チェック
  const domainInfo = getDeployUserDomainInfo();
  if (domainInfo.deployDomain && domainInfo.deployDomain !== '' && !domainInfo.isDomainMatch) {
    throw new Error(`ドメインアクセスが制限されています。許可されたドメイン: ${domainInfo.deployDomain}, 現在のドメイン: ${domainInfo.currentDomain}`);
  }

  // 既存ユーザーチェック（1ユーザー1行の原則）
  const existingUser = findUserByEmail(adminEmail);
  let userId, appUrls;

  if (existingUser) {
    // 既存ユーザーの場合は最小限の更新のみ（設定は保護）
    userId = existingUser.userId;
    const existingConfig = JSON.parse(existingUser.configJson || '{}');

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

  var initialConfig = {
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

  var userData = {
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
    // 生成されたユーザー情報のキャッシュをクリア
    invalidateUserCache(userId, adminEmail, null, false);
  } catch (e) {
    logDatabaseError(e, 'userRegistration', { userId: userInfo.userId, email: userInfo.email });
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
    var reactingUserEmail = Session.getActiveUser().getEmail();
    var ownerUserId = requestUserId; // requestUserId を使用

    // ボードオーナーの情報をDBから取得（キャッシュ利用）
    var boardOwnerInfo = findUserById(ownerUserId);
    if (!boardOwnerInfo) {
      throw new Error('無効なボードです。');
    }

    var result = processReaction(
      boardOwnerInfo.spreadsheetId,
      sheetName,
      rowIndex,
      reactionKey,
      reactingUserEmail
    );

    // Page.html期待形式に変換
    if (result && result.status === 'success') {
      // 更新後のリアクション情報を取得
      var updatedReactions = getRowReactions(boardOwnerInfo.spreadsheetId, sheetName, rowIndex, reactingUserEmail);

      return {
        status: "ok",
        reactions: updatedReactions
      };
    } else {
      throw new Error(result.message || 'リアクションの処理に失敗しました');
    }
  } catch (e) {
    logError(e, 'addReaction', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.SYSTEM, { userId, rowIndex, reaction });
    return {
      status: "error",
      message: e.message
    };
  } finally {
    // 実行終了時にユーザー情報キャッシュをクリア
    clearExecutionUserInfoCache();
  }
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
      throw new Error(`バッチサイズが制限を超えています (最大${MAX_BATCH_SIZE}件)`);
    }

    debugLog('🔄 バッチリアクション処理開始:', batchOperations.length + '件');

    var reactingUserEmail = Session.getActiveUser().getEmail();
    var ownerUserId = requestUserId;

    // ボードオーナーの情報をDBから取得（キャッシュ利用）
    var boardOwnerInfo = findUserById(ownerUserId);
    if (!boardOwnerInfo) {
      throw new Error('無効なボードです。');
    }

    // バッチ処理結果を格納
    var batchResults = [];
    var processedRows = new Set(); // 重複行の追跡

    // 既存のsheetNameを取得（最初の操作から）
    var sheetName = getCurrentSheetName(boardOwnerInfo.spreadsheetId);

    // バッチ操作を順次処理（既存のprocessReaction関数を再利用）
    for (var i = 0; i < batchOperations.length; i++) {
      var operation = batchOperations[i];

      try {
        // 入力検証
        if (!operation.rowIndex || !operation.reaction) {
          warnLog('無効な操作をスキップ:', operation);
          continue;
        }

        // 既存のprocessReaction関数を使用（100%互換性保証）
        var result = processReaction(
          boardOwnerInfo.spreadsheetId,
          sheetName,
          operation.rowIndex,
          operation.reaction,
          reactingUserEmail
        );

        if (result && result.status === 'success') {
          // 更新後のリアクション情報を取得
          var updatedReactions = getRowReactions(
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
    var finalResults = [];
    processedRows.forEach(function(rowIndex) {
      try {
        var latestReactions = getRowReactions(
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
    var spreadsheet = openSpreadsheetOptimized(spreadsheetId);
    var sheets = spreadsheet.getSheets();

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

  // Session.getActiveUser().getEmail()を直接使用（verifyUserAccessではemail比較が必要）
  const activeUserEmail = Session.getActiveUser().getEmail();
  debugLog(`verifyUserAccess start: userId=${requestUserId}, email=${activeUserEmail}`);
  if (!activeUserEmail) {
    throw new Error('認証エラー: アクティブユーザーの情報を取得できませんでした');
  }

  // requestUserIdがnull/undefinedの場合は新規ユーザーとして扱い、基本的なセッション検証のみ実行
  if (requestUserId == null) {
    debugLog(`✅ 新規ユーザーセッション検証成功: ${activeUserEmail}`);
    return;
  }

  const requestedUserInfo = findUserById(requestUserId);

  if (!requestedUserInfo) {
    throw new Error(`認証エラー: 指定されたユーザーID (${requestUserId}) が見つかりません。`);
  }

  // 管理者かどうかを確認
  if (activeUserEmail !== requestedUserInfo.adminEmail) {
    const config = JSON.parse(requestedUserInfo.configJson || '{}');
    if (config.appPublished === true) {
      debugLog(`✅ 公開ボード閲覧許可: ${activeUserEmail} -> ${requestUserId}`);
      return;
    }
    throw new Error(`権限エラー: ${activeUserEmail} はユーザーID ${requestUserId} のデータにアクセスする権限がありません。`);
  }
  debugLog(`✅ ユーザーアクセス検証成功: ${activeUserEmail} は ${requestUserId} のデータにアクセスできます。`);
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
    // キャッシュキー生成（パフォーマンス向上）
    var requestKey = `publishedData_${requestUserId}_${classFilter}_${sortOrder}_${adminMode}`;

    // キャッシュバイパス時は直接実行
    if (bypassCache === true) {
      debugLog('🔄 キャッシュバイパス：最新データを直接取得');
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
      var currentUserId = requestUserId; // requestUserId を使用
      debugLog('getPublishedSheetData: userId=%s, classFilter=%s, sortOrder=%s, adminMode=%s', currentUserId, classFilter, sortOrder, adminMode);

      var userInfo = getOrFetchUserInfo(currentUserId, 'userId', {
        useExecutionCache: true,
        ttl: 300
      });
      if (!userInfo) {
        throw new Error('ユーザー情報が見つかりません');
      }
      debugLog('getPublishedSheetData: userInfo=%s', JSON.stringify(userInfo));

      var configJson = JSON.parse(userInfo.configJson || '{}');
      debugLog('getPublishedSheetData: configJson=%s', JSON.stringify(configJson));

    // セットアップ状況を確認
    var setupStatus = configJson.setupStatus || 'pending';

    // 公開対象のスプレッドシートIDとシート名を取得
    var publishedSpreadsheetId = configJson.publishedSpreadsheetId;
    var publishedSheetName = configJson.publishedSheetName;

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
    var sheetKey = 'sheet_' + publishedSheetName;
    var sheetConfig = configJson[sheetKey] || {};
    debugLog('getPublishedSheetData: sheetConfig=%s', JSON.stringify(sheetConfig));

    // Check if current user is the board owner
    var isOwner = (configJson.ownerId === currentUserId);
    debugLog('getPublishedSheetData: isOwner=%s, ownerId=%s, currentUserId=%s', isOwner, configJson.ownerId, currentUserId);

    // データ取得
    var sheetData = getSheetData(currentUserId, publishedSheetName, classFilter, sortOrder, adminMode);
    debugLog('getPublishedSheetData: sheetData status=%s, totalCount=%s', sheetData.status, sheetData.totalCount);

    // 診断: スプレッドシートとシートの存在確認
    try {
      debugLog('🔍 診断: スプレッドシート詳細情報');
      debugLog('  publishedSpreadsheetId:', publishedSpreadsheetId);
      debugLog('  publishedSheetName:', publishedSheetName);
      debugLog('  データ取得結果:', {
        status: sheetData.status,
        totalCount: sheetData.totalCount,
        hasData: !!(sheetData.data && sheetData.data.length > 0),
        hasHeaders: !!(sheetData.headers && sheetData.headers.length > 0)
      });

      if (sheetData.totalCount === 0) {
        debugLog('⚠️ 診断: データが0件です。原因を調査します...');
        var spreadsheet = openSpreadsheetOptimized(publishedSpreadsheetId);
        var sheet = spreadsheet.getSheetByName(publishedSheetName);
        if (sheet) {
          var lastRow = sheet.getLastRow();
          var lastColumn = sheet.getLastColumn();
          debugLog('  スプレッドシート実状態:', {
            lastRow: lastRow,
            lastColumn: lastColumn,
            hasData: lastRow > 1, // ヘッダー行を除く
            範囲: `A1:${String.fromCharCode(64 + lastColumn)}${lastRow}`
          });

          if (lastRow <= 1) {
            debugLog('⚠️ 診断結果: スプレッドシートにデータ行がありません（ヘッダーのみ）');
          }
        } else {
          debugLog('❌ 診断結果: 指定されたシート名が見つかりません');
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
    var mainHeaderName;
    if (setupStatus === 'pending') {
      mainHeaderName = 'セットアップ中...';
    } else {
      mainHeaderName = sheetConfig.opinionHeader || COLUMN_HEADERS.OPINION;
    }

    // その他のヘッダーフィールドも安全に取得
    var reasonHeaderName, classHeaderName, nameHeaderName;
    if (setupStatus === 'pending') {
      reasonHeaderName = 'セットアップ中...';
      classHeaderName = 'セットアップ中...';
      nameHeaderName = 'セットアップ中...';
    } else {
      reasonHeaderName = sheetConfig.reasonHeader || COLUMN_HEADERS.REASON;
      classHeaderName = sheetConfig.classHeader !== undefined ? sheetConfig.classHeader : COLUMN_HEADERS.CLASS;
      nameHeaderName = sheetConfig.nameHeader !== undefined ? sheetConfig.nameHeader : COLUMN_HEADERS.NAME;
    }
    debugLog('getPublishedSheetData: Configured Headers - mainHeaderName=%s, reasonHeaderName=%s, classHeaderName=%s, nameHeaderName=%s', mainHeaderName, reasonHeaderName, classHeaderName, nameHeaderName);

    // ヘッダーインデックスマップを取得（キャッシュされた実際のマッピング）
    var headerIndices = getHeaderIndices(publishedSpreadsheetId, publishedSheetName);
    debugLog('getPublishedSheetData: Available headerIndices=%s', JSON.stringify(headerIndices));

    // 動的列名のマッピング: 設定された名前と実際のヘッダーを照合
    var mappedIndices = mapConfigToActualHeaders({
      opinionHeader: mainHeaderName,
      reasonHeader: reasonHeaderName,
      classHeader: classHeaderName,
      nameHeader: nameHeaderName
    }, headerIndices);
    debugLog('getPublishedSheetData: Mapped indices=%s', JSON.stringify(mappedIndices));

    var formattedData = formatSheetDataForFrontend(sheetData.data, mappedIndices, headerIndices, adminMode, isOwner, sheetData.displayMode);

    debugLog('getPublishedSheetData: formattedData length=%s', formattedData.length);
    debugLog('getPublishedSheetData: formattedData content=%s', JSON.stringify(formattedData));

    // ボードのタイトルを実際のスプレッドシートのヘッダーから取得
    let headerTitle = publishedSheetName || '今日のお題';
    if (mappedIndices.opinionHeader !== undefined) {
      for (var actualHeader in headerIndices) {
        if (headerIndices[actualHeader] === mappedIndices.opinionHeader) {
          headerTitle = actualHeader;
          debugLog('getPublishedSheetData: Using actual header as title: "%s"', headerTitle);
          break;
        }
      }
    }

    var finalDisplayMode = (adminMode === true) ? DISPLAY_MODES.NAMED : (configJson.displayMode || DISPLAY_MODES.ANONYMOUS);

    var result = {
      header: headerTitle,
      sheetName: publishedSheetName,
      showCounts: (adminMode === true) ? true : (configJson.showCounts === true),
      displayMode: finalDisplayMode,
      data: formattedData,
      rows: formattedData // 後方互換性のため
    };

    debugLog('🔍 最終結果:', {
      adminMode: adminMode,
      originalDisplayMode: sheetData.displayMode,
      finalDisplayMode: finalDisplayMode,
      dataCount: formattedData.length,
      showCounts: result.showCounts
    });
    debugLog('getPublishedSheetData: Returning result=%s', JSON.stringify(result));
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
    debugLog('🔄 増分データ取得開始: sinceRowCount=%s', sinceRowCount);

    var currentUserId = requestUserId; // requestUserId を使用

    var userInfo = getOrFetchUserInfo(currentUserId, 'userId', {
      useExecutionCache: true,
      ttl: 300
    });
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    var configJson = JSON.parse(userInfo.configJson || '{}');
    var setupStatus = configJson.setupStatus || 'pending';
    var publishedSpreadsheetId = configJson.publishedSpreadsheetId;
    var publishedSheetName = configJson.publishedSheetName;

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
    var ss = openSpreadsheetOptimized(publishedSpreadsheetId);
      debugLog('DEBUG: Spreadsheet object obtained: %s', ss ? ss.getName() : 'null');

      var sheet = ss.getSheetByName(publishedSheetName);
      debugLog('DEBUG: Sheet object obtained: %s', sheet ? sheet.getName() : 'null');

    if (!sheet) {
      throw new Error('指定されたシートが見つかりません: ' + publishedSheetName);
    }

    var lastRow = sheet.getLastRow(); // スプレッドシートの最終行
    var headerRow = 1; // ヘッダー行は1行目と仮定

    // 実際に読み込むべき開始行を計算 (sinceRowCountはデータ行数なので、+1してヘッダーを考慮)
    // sinceRowCountが0の場合、ヘッダーの次の行から読み込む
    var startRowToRead = sinceRowCount + headerRow + 1;

    // 新しいデータがない場合
    if (lastRow < startRowToRead) {
      debugLog('🔍 増分データ分析: 新しいデータなし。lastRow=%s, startRowToRead=%s', lastRow, startRowToRead);
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
    var numRowsToRead = lastRow - startRowToRead + 1;

    // 必要なデータのみをスプレッドシートから直接取得
    // getRange(row, column, numRows, numColumns)
    // ここでは全列を取得すると仮定 (A列から最終列まで)
    var lastColumn = sheet.getLastColumn();
    var rawNewData = sheet.getRange(startRowToRead, 1, numRowsToRead, lastColumn).getValues();

    debugLog('📥 スプレッドシートから直接取得した新しいデータ:', rawNewData.length, '件');

    // ヘッダーインデックスマップを取得（キャッシュされた実際のマッピング）
    var headerIndices = getHeaderIndices(publishedSpreadsheetId, publishedSheetName);

    // 動的列名のマッピング: 設定された名前と実際のヘッダーを照合
    var sheetConfig = configJson['sheet_' + publishedSheetName] || {};
    var mainHeaderName = sheetConfig.opinionHeader || COLUMN_HEADERS.OPINION;
    var reasonHeaderName = sheetConfig.reasonHeader || COLUMN_HEADERS.REASON;
    var classHeaderName = sheetConfig.classHeader !== undefined ? sheetConfig.classHeader : COLUMN_HEADERS.CLASS;
    var nameHeaderName = sheetConfig.nameHeader !== undefined ? sheetConfig.nameHeader : COLUMN_HEADERS.NAME;
    var mappedIndices = mapConfigToActualHeaders({
      opinionHeader: mainHeaderName,
      reasonHeader: reasonHeaderName,
      classHeader: classHeaderName,
      nameHeader: nameHeaderName
    }, headerIndices);

    // ユーザー情報と管理者モードの取得
    var isOwner = (configJson.ownerId === currentUserId);
    var displayMode = configJson.displayMode || DISPLAY_MODES.ANONYMOUS;

    // 新しいデータを既存の処理パイプラインと同様に加工
    var headers = sheet.getRange(headerRow, 1, 1, lastColumn).getValues()[0];
    var rosterMap = buildRosterMap([]); // roster is not used
    var processedData = rawNewData.map(function(row, idx) {
      return processRowData(row, headers, headerIndices, rosterMap, displayMode, startRowToRead + idx, isOwner);
    });

    // 取得した生データをPage.htmlが期待する形式にフォーマット
    var formattedNewData = formatSheetDataForFrontend(processedData, mappedIndices, headerIndices, adminMode, isOwner, displayMode);

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
    var currentUserId = requestUserId; // requestUserId を使用

    if (!currentUserId) {
      warnLog('getAvailableSheets: No current user ID set');
      return [];
    }

    var sheets = getSheetsList(currentUserId);

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
    var userInfo = findUserById(userId);
    if (!userInfo || !userInfo.spreadsheetId) {
      warnLog('getSheetsList: User info or spreadsheetId not found for userId:', userId);
      return [];
    }

    var spreadsheet = openSpreadsheetOptimized(userInfo.spreadsheetId);
    var sheets = spreadsheet.getSheets();

    var sheetList = sheets.map(function(sheet) {
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
    var currentUserId = requestUserId; // requestUserId を使用

    var userInfo = findUserById(currentUserId);
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
  var currentUserEmail = Session.getActiveUser().getEmail();

  return rawData.map(function(row, index) {
    var classIndex = mappedIndices.classHeader;
    var opinionIndex = mappedIndices.opinionHeader;
    var reasonIndex = mappedIndices.reasonHeader;
    var nameIndex = mappedIndices.nameHeader;

    debugLog('formatSheetDataForFrontend: Row %s - classIndex=%s, opinionIndex=%s, reasonIndex=%s, nameIndex=%s', index, classIndex, opinionIndex, reasonIndex, nameIndex);
    debugLog('formatSheetDataForFrontend: Row data length=%s, first few values=%s', row.originalData ? row.originalData.length : 'undefined', row.originalData ? JSON.stringify(row.originalData.slice(0, 5)) : 'undefined');

    var nameValue = '';
    var shouldShowName = (adminMode === true || displayMode === DISPLAY_MODES.NAMED || isOwner);
    var hasNameIndex = nameIndex !== undefined;
    var hasOriginalData = row.originalData && row.originalData.length > 0;

    if (shouldShowName && hasNameIndex && hasOriginalData) {
      nameValue = row.originalData[nameIndex] || '';
    }

    if (!nameValue && shouldShowName && hasOriginalData) {
      var emailIndex = headerIndices[COLUMN_HEADERS.EMAIL];
      if (emailIndex !== undefined && row.originalData[emailIndex]) {
        nameValue = row.originalData[emailIndex].split('@')[0];
      }
    }

    debugLog('🔍 サーバー側名前データ詳細:', {
      rowIndex: row.rowNumber || (index + 2),
      shouldShowName: shouldShowName,
      adminMode: adminMode,
      displayMode: displayMode,
      isOwner: isOwner,
      nameIndex: nameIndex,
      hasNameIndex: hasNameIndex,
      hasOriginalData: hasOriginalData,
      originalDataLength: row.originalData ? row.originalData.length : 'undefined',
      nameValue: nameValue,
      rawNameData: hasOriginalData && nameIndex !== undefined ? row.originalData[nameIndex] : 'N/A'
    });

    // リアクション状態を判定するヘルパー関数
    function checkReactionState(reactionKey) {
      var columnName = COLUMN_HEADERS[reactionKey];
      var columnIndex = headerIndices[columnName];
      var count = 0;
      var reacted = false;

      if (columnIndex !== undefined && row.originalData && row.originalData[columnIndex]) {
        var reactionString = row.originalData[columnIndex].toString();
        if (reactionString) {
          var reactions = parseReactionString(reactionString);
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
    var reasonValue = '';
    if (reasonIndex !== undefined && row.originalData && row.originalData[reasonIndex] !== undefined) {
      reasonValue = row.originalData[reasonIndex] || '';
    }

    // 意見と理由の取得（マッピングが利用できない場合はprocessedRowから取得）
    var opinionValue = '';
    var finalReasonValue = reasonValue;

    if (opinionIndex !== undefined && row.originalData && row.originalData[opinionIndex]) {
      opinionValue = row.originalData[opinionIndex];
    } else if (row.opinion) {
      // フォールバック: processedRowから取得
      opinionValue = row.opinion;
    }

    if (!finalReasonValue && row.reason) {
      // フォールバック: processedRowから取得
      finalReasonValue = row.reason;
    }

    debugLog('formatSheetDataForFrontend: Content extraction for row %s - opinion="%s", reason="%s"',
             index, opinionValue.substring(0, 50), finalReasonValue.substring(0, 50));

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
    var currentUserId = requestUserId;
    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    var configJson = JSON.parse(userInfo.configJson || '{}');

    // --- 統一された自動修復システム ---
    const healingResult = performAutoHealing(userInfo, configJson, currentUserId);
    if (healingResult.updated) {
      configJson = healingResult.configJson;
    }

    var sheets = getSheetsList(currentUserId);
    var appUrls = generateUserUrls(currentUserId);

    // 回答数を取得
    var answerCount = 0;
    var totalReactions = 0;
    try {
      if (configJson.publishedSpreadsheetId && configJson.publishedSheetName) {
        var responseData = getResponsesData(currentUserId, configJson.publishedSheetName);
        if (responseData.status === 'success') {
          answerCount = responseData.data.length;
          // リアクション数の概算計算（詳細実装は後回し）
          totalReactions = answerCount * 2; // 暫定値
        }
      }
    } catch (countError) {
      warnLog('回答数の取得に失敗: ' + countError.message);
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
      formUrl: configJson.formUrl || '',
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

    var currentUserId = userId;

    // 最適化モード: 事前取得済みuserInfoを使用、なければ取得
    var userInfo = options.userInfo || findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    var configJson = JSON.parse(userInfo.configJson || '{}');

    // シート設定を更新
    var sheetKey = 'sheet_' + sheetName;
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

    var currentUserId = userId;

    // 最適化モード: 事前取得済みuserInfoを使用、なければ取得
    var userInfo = options.userInfo || findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    var configJson = JSON.parse(userInfo.configJson || '{}');

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
// セットアップ関数
// =================================================================

/**
 * アプリケーションの初期セットアップ（管理者が手動で実行） (マルチテナント対応版)
 */
function setupApplication(credsJson, dbId) {
  try {
    JSON.parse(credsJson);
    if (typeof dbId !== 'string' || dbId.length !== 44) {
      throw new Error('無効なスプレッドシートIDです。IDは44文字の文字列である必要があります。');
    }

    var props = PropertiesService.getScriptProperties();
    props.setProperty(SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS, credsJson);
    props.setProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID, dbId);

    var adminEmail = Session.getEffectiveUser().getEmail();
    if (adminEmail) {
      props.setProperty(SCRIPT_PROPS_KEYS.ADMIN_EMAIL, adminEmail);
    }

    // データベースシートの初期化
    initializeDatabaseSheet(dbId);

    infoLog('✅ セットアップが正常に完了しました。');
    return { status: 'success', message: 'セットアップが正常に完了しました。' };
  } catch (e) {
    logError(e, 'customSetup', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.SYSTEM, { userId });
    throw new Error('セットアップに失敗しました: ' + e.message);
  }
}

/**
 * セットアップ状態をテストする (マルチテナント対応版)
 */
function testSetup() {
  try {
    var props = PropertiesService.getScriptProperties();
    var dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
    var creds = props.getProperty(SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS);

    if (!dbId) {
      return { status: 'error', message: 'データベーススプレッドシートIDが設定されていません。' };
    }

    if (!creds) {
      return { status: 'error', message: 'サービスアカウント認証情報が設定されていません。' };
    }

    // データベースへの接続テスト
    try {
      var userInfo = findUserByEmail(Session.getActiveUser().getEmail());
      return {
        status: 'success',
        message: 'セットアップは正常に完了しています。システムは使用準備が整いました。',
        details: {
          databaseConnected: true,
          userCount: userInfo ? 'ユーザー登録済み' : '未登録',
          serviceAccountConfigured: true
        }
      };
    } catch (dbError) {
      return {
        status: 'warning',
        message: '設定は保存されていますが、データベースアクセスに問題があります。',
        details: { error: dbError.message }
      };
    }

  } catch (e) {
    logError(e, 'testCustomSetup', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.SYSTEM, { userId });
    return { status: 'error', message: 'セットアップテストに失敗しました: ' + e.message };
  }
}

// =================================================================
// ヘルパー関数
// =================================================================

// include 関数は main.gs で定義されています
function getResponsesData(userId, sheetName) {
  var userInfo = getOrFetchUserInfo(userId, 'userId', {
    useExecutionCache: true,
    ttl: 300
  });
  if (!userInfo) {
    return { status: 'error', message: 'ユーザー情報が見つかりません' };
  }

  try {
    var service = getSheetsServiceCached();
    var spreadsheetId = userInfo.spreadsheetId;
    var range = "'" + (sheetName || 'フォームの回答 1') + "'!A:Z";

    var response = batchGetSheetsData(service, spreadsheetId, [range]);
    var values = response.valueRanges[0].values || [];

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
 * 現在のユーザーのステータスを取得し、UIに表示するための情報を返します。(マルチテナント対応版)
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @returns {object} ユーザーのステータス情報
 */
function getCurrentUserStatus(requestUserId) {
  try {
    const activeUserEmail = Session.getActiveUser().getEmail();

    // requestUserIdが無効な場合は、メールアドレスでユーザーを検索
    let userInfo;
    if (requestUserId && requestUserId.trim() !== '') {
      verifyUserAccess(requestUserId);
      userInfo = findUserById(requestUserId);
    } else {
      userInfo = findUserByEmail(activeUserEmail);
    }

    if (!userInfo) {
      return { status: 'error', message: 'ユーザー情報が見つかりません。' };
    }

    return {
      status: 'success',
      userInfo: {
        userId: userInfo.userId,
        adminEmail: userInfo.adminEmail,
        isActive: userInfo.isActive,
        lastAccessedAt: userInfo.lastAccessedAt
      }
    };
  } catch (e) {
    logError(e, 'getCurrentUserStatus', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.SYSTEM, { userId });
    return { status: 'error', message: 'ステータス取得に失敗しました: ' + e.message };
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
    var currentUserId = requestUserId; // requestUserId を使用

    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    var configJson = JSON.parse(userInfo.configJson || '{}');

    // フォーム回答数を取得
    var answerCount = 0;
    try {
      if (configJson.publishedSpreadsheetId && configJson.publishedSheet) {
        var responseData = getResponsesData(currentUserId, configJson.publishedSheet);
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
    // Session.getActiveUser().getEmail() が requestUserId の adminEmail と一致するかどうかを verifyUserAccess で既にチェック済み
    // ここでは単に userInfo.adminEmail と Session.getActiveUser().getEmail() が一致するかを返す
    return Session.getActiveUser().getEmail() === userInfo.adminEmail;
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
  const key = `rowCount_${spreadsheetId}_${sheetName}_${classFilter}`;
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
    const configJson = JSON.parse(userInfo.configJson || '{}');

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
    var currentUserId = requestUserId; // requestUserId を使用

    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    var configJson = JSON.parse(userInfo.configJson || '{}');

    if (configJson.editFormUrl) {
      try {
        var formId = extractFormIdFromUrl(configJson.editFormUrl);
        var form = FormApp.openById(formId);

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
    var currentUserId = requestUserId; // requestUserId を使用

    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    var configJson = JSON.parse(userInfo.configJson || '{}');

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
    var currentUserId = requestUserId; // requestUserId を使用

    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    // 管理者権限チェック - 統一されたcheckAdmin関数を使用
    if (!checkAdmin(requestUserId)) {
      throw new Error('ハイライト機能は管理者のみ使用できます');
    }

    var result = processHighlightToggle(
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
  var userEmail = setupContext.userEmail;
  var requestUserId = setupContext.requestUserId;

  // ステップ1: ユーザー専用フォルダを作成
  debugLog('📁 ステップ1: フォルダ作成中...');
  var folder = createUserFolder(userEmail);

  // ステップ2: Googleフォームとスプレッドシートを作成
  debugLog('📝 ステップ2: フォーム作成中...');
  var formAndSsInfo = createUnifiedForm('study', userEmail, requestUserId);

  // 作成したファイルをフォルダに移動（改善版：冗長処理除去と安全な移動処理）
  if (folder) {
    var moveResults = { form: false, spreadsheet: false };
    var moveErrors = [];

    try {
      var formFile = DriveApp.getFileById(formAndSsInfo.formId);
      var ssFile = DriveApp.getFileById(formAndSsInfo.spreadsheetId);

      // フォームファイルの移動処理
      try {
        // 既にフォルダに存在するかチェック（重複移動を防止）
        var formParents = formFile.getParents();
        var isFormAlreadyInFolder = false;

        while (formParents.hasNext()) {
          if (formParents.next().getId() === folder.getId()) {
            isFormAlreadyInFolder = true;
            break;
          }
        }

        if (!isFormAlreadyInFolder) {
          debugLog('📝 フォームファイルを移動中: %s → %s', formFile.getId(), folder.getName());
          folder.addFile(formFile);
          // ルートフォルダから削除（適切なタイミングで実行）
          DriveApp.getRootFolder().removeFile(formFile);
          moveResults.form = true;
          infoLog('✅ フォームファイル移動完了');
        } else {
          debugLog('ℹ️ フォームファイルは既にフォルダに存在します');
          moveResults.form = true;
        }
      } catch (formMoveError) {
        moveErrors.push('フォームファイル移動エラー: ' + formMoveError.message);
        errorLog('❌ フォームファイルの移動に失敗:', formMoveError.message);
      }

      // スプレッドシートファイルの移動処理
      try {
        // 既にフォルダに存在するかチェック（重複移動を防止）
        var ssParents = ssFile.getParents();
        var isSsAlreadyInFolder = false;

        while (ssParents.hasNext()) {
          if (ssParents.next().getId() === folder.getId()) {
            isSsAlreadyInFolder = true;
            break;
          }
        }

        if (!isSsAlreadyInFolder) {
          debugLog('📊 スプレッドシートファイルを移動中: %s → %s', ssFile.getId(), folder.getName());
          folder.addFile(ssFile);
          // ルートフォルダから削除（適切なタイミングで実行）
          DriveApp.getRootFolder().removeFile(ssFile);
          moveResults.spreadsheet = true;
          infoLog('✅ スプレッドシートファイル移動完了');
        } else {
          debugLog('ℹ️ スプレッドシートファイルは既にフォルダに存在します');
          moveResults.spreadsheet = true;
        }
      } catch (ssMoveError) {
        moveErrors.push('スプレッドシートファイル移動エラー: ' + ssMoveError.message);
        errorLog('❌ スプレッドシートファイルの移動に失敗:', ssMoveError.message);
      }

      // 移動結果のログ出力
      if (moveResults.form && moveResults.spreadsheet) {
        infoLog('✅ 全ファイルのフォルダ移動が完了: ' + folder.getName());
      } else {
        warnLog('⚠️ 一部のファイル移動に失敗しましたが、処理を継続します');
        debugLog('移動結果: フォーム=%s, スプレッドシート=%s', moveResults.form, moveResults.spreadsheet);
        if (moveErrors.length > 0) {
          debugLog('移動エラー詳細: %s', moveErrors.join('; '));
        }
      }

    } catch (generalError) {
      errorLog('❌ ファイル移動処理で予期しないエラー:', generalError.message);
      // ファイル移動失敗は致命的ではないため、処理は継続
      debugLog('ファイルはマイドライブに残りますが、システムは正常に動作します');
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
  var requestUserId = setupContext.requestUserId;
  var configJson = setupContext.configJson;
  var userEmail = setupContext.userEmail;
  var formAndSsInfo = createdFiles.formAndSsInfo;
  var folder = createdFiles.folder;

  debugLog('💾 ステップ3: データベース更新中...');
  debugLog('📊 新規作成されたファイル情報:');
  debugLog('  📝 フォームID:', formAndSsInfo.formId);
  debugLog('  📊 スプレッドシートID:', formAndSsInfo.spreadsheetId);
  debugLog('  📄 シート名:', formAndSsInfo.sheetName);

  // クイックスタート用の適切な初期設定を作成
  var sheetConfigKey = 'sheet_' + formAndSsInfo.sheetName;
  var quickStartSheetConfig = {
    opinionHeader: '今日の学習について、あなたの考えや感想を聞かせてください',
    reasonHeader: 'そう考える理由や体験があれば教えてください（任意）',
    nameHeader: '名前',
    classHeader: 'クラス',
    formUrl: formAndSsInfo.viewFormUrl || formAndSsInfo.formUrl, // シート固有のフォームURL保存
    editFormUrl: formAndSsInfo.editFormUrl, // 編集用URL保存
    lastModified: new Date().toISOString()
  };

  debugLog('📝 クイックスタート用質問文設定:');
  debugLog('  💭 メイン質問:', quickStartSheetConfig.opinionHeader);
  debugLog('  💡 理由質問:', quickStartSheetConfig.reasonHeader);

  // 型安全性確保: publishedSheetNameの明示的文字列変換
  var safeSheetName = formAndSsInfo.sheetName;
  if (typeof safeSheetName !== 'string') {
    errorLog('❌ quickStartSetup: formAndSsInfo.sheetNameが文字列ではありません:', typeof safeSheetName, safeSheetName);
    safeSheetName = String(safeSheetName); // 強制的に文字列化
  }
  if (safeSheetName === 'true' || safeSheetName === 'false') {
    errorLog('❌ quickStartSetup: 無効なシート名が検出されました:', safeSheetName);
    throw new Error('無効なシート名: ' + safeSheetName);
  }

  // 6時間自動停止機能の設定
  var publishedAt = new Date().toISOString();
  var autoStopMinutes = 360; // 6時間 = 360分
  var scheduledEndAt = new Date(Date.now() + (autoStopMinutes * 60 * 1000)).toISOString();

  var updatedConfig = {
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
  debugLog('💾 ユーザーデータベースを新しいセットアップで更新中...');
  var updateData = {
    spreadsheetId: formAndSsInfo.spreadsheetId,
    spreadsheetUrl: formAndSsInfo.spreadsheetUrl,
    configJson: JSON.stringify(updatedConfig),
    lastAccessedAt: new Date().toISOString()
  };

  debugLog('📋 データベース更新内容:');
  debugLog('  📊 新スプレッドシートID:', updateData.spreadsheetId);
  debugLog('  🔗 新スプレッドシートURL:', updateData.spreadsheetUrl);

  updateUser(requestUserId, updateData);

  infoLog('✅ ユーザーデータベース更新完了!');

  // 重要: 新しいセットアップ完了後に包括的キャッシュ同期を実行（二重保証）
  debugLog('🗑️ 新しいセットアップ確実反映のための包括的キャッシュ同期中...');

  // updateUserで既にキャッシュ同期されているが、クイックスタートの場合は追加で確実性を高める
  synchronizeCacheAfterCriticalUpdate(requestUserId, userEmail, null, formAndSsInfo.spreadsheetId);

  // 最終検証: 更新が正確に反映されているかデータベースから直接確認
  debugLog('🔍 データベース更新の最終検証中...');
  debugLog('🎯 期待するスプレッドシートID:', formAndSsInfo.spreadsheetId);

  // 少し待ってからデータベースを確認（更新の反映を待つ）
  Utilities.sleep(500);

  var verificationUserInfo = findUserByIdFresh(requestUserId);
  debugLog('📊 検証結果:');
  debugLog('  取得したスプレッドシートID:', verificationUserInfo ? verificationUserInfo.spreadsheetId : 'null');
  debugLog('  期待値との一致:', verificationUserInfo && verificationUserInfo.spreadsheetId === formAndSsInfo.spreadsheetId);

  if (verificationUserInfo && verificationUserInfo.spreadsheetId === formAndSsInfo.spreadsheetId) {
    infoLog('✅ データベース更新検証成功: 新しいスプレッドシートID確認');
  } else {
    warnLog('⚠️ データベース更新検証で不一致が検出されましたが、処理を続行します');
    debugLog('📝 注意: この不一致は一時的なキャッシュの問題の可能性があります');

    // 検証失敗でもQuick Startを続行する（フォーム作成は成功している）
    debugLog('🚀 フォーム作成は成功しているため、Quick Startを続行します');
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
    debugLog('🌐 QuickStart自動公開開始', { requestUserId, sheetName });
    
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
    
    debugLog('🔍 自動公開前の事前確認完了', {
      userId: requestUserId,
      hasSpreadsheet: !!userInfo.spreadsheetId,
      targetSheet: trimmedSheetName
    });

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
  var requestUserId = setupContext.requestUserId;
  var formAndSsInfo = createdFiles.formAndSsInfo;

  // 最終検証：新規作成されたファイルの確認
  debugLog('🔍 最終検証 - 作成されたファイル:');
  debugLog('  📝 フォームID:', formAndSsInfo.formId);
  debugLog('  📊 スプレッドシートID:', formAndSsInfo.spreadsheetId);
  debugLog('  📄 シート名:', formAndSsInfo.sheetName);
  debugLog('  🔗 フォームURL:', formAndSsInfo.formUrl);
  debugLog('  🔗 スプレッドシートURL:', formAndSsInfo.spreadsheetUrl);

  // 公開結果の検証
  var isPublished = publishResult && publishResult.success && publishResult.published;
  var publishMessage = isPublished 
    ? '回答ボードが自動的に公開されました！' 
    : '回答ボードが作成されました。管理パネルから手動で公開してください。';
  
  debugLog('🔍 公開結果検証:', {
    hasPublishResult: !!publishResult,
    isPublished: isPublished,
    publishMessage: publishMessage
  });

  infoLog('✅ クイックスタートセットアップ完了: ' + requestUserId);

  var appUrls = generateUserUrls(requestUserId);
  
  // 拡張されたレスポンス情報
  var response = {
    status: 'success',
    message: 'クイックスタートが完了しました！' + publishMessage,
    webAppUrl: appUrls.webAppUrl,
    adminUrl: appUrls.adminUrl,
    viewUrl: appUrls.viewUrl,
    setupUrl: appUrls.setupUrl,
    formUrl: updatedConfig.formUrl,
    editFormUrl: updatedConfig.editFormUrl,
    spreadsheetUrl: formAndSsInfo.spreadsheetUrl,
    folderUrl: updatedConfig.folderUrl,
    // 進捗システム用の詳細情報
    setupComplete: true,
    autoPublished: isPublished,
    publishResult: publishResult,
    sheetName: formAndSsInfo.sheetName,
    formId: formAndSsInfo.formId,
    spreadsheetId: formAndSsInfo.spreadsheetId,
    // フロントエンド完了通知用のタイムスタンプ
    completedAt: new Date().toISOString(),
    // 成功ステップの詳細
    completedSteps: [
      'ユーザー専用フォルダの作成',
      'Googleフォームとスプレッドシートの作成',
      'データベース更新とキャッシュ管理',
      isPublished ? '回答ボードの自動公開' : '回答ボード作成（手動公開待ち）',
      'キャッシュクリアと最終化',
      'セットアップ完了'
    ]
  };
  
  debugLog('📤 QuickStart最終レスポンス生成完了:', response);
  return response;
}

/**
 * クイックスタートセットアップのコンテキストを初期化
 * @param {string} requestUserId - ユーザーID
 * @returns {object} セットアップコンテキスト
 */
function initializeQuickStartContext(requestUserId) {
  debugLog('🚀 クイックスタートセットアップ開始: ' + requestUserId);

  // ユーザー情報の取得
  var userInfo = findUserById(requestUserId);
  if (!userInfo) {
    throw new Error('ユーザー情報が見つかりません');
  }

  var configJson = JSON.parse(userInfo.configJson || '{}');
  var userEmail = userInfo.adminEmail;

  // クイックスタート繰り返し実行を許可
  // 既存のセットアップがある場合は完全にリセットして新しいセットアップで上書きする
  if (configJson.formCreated || userInfo.spreadsheetId) {
    debugLog('⚠️ 既存のセットアップが検出されました。新しいセットアップで完全に上書きします。');
    debugLog('🔄 既存スプレッドシートID:', userInfo.spreadsheetId);

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
    debugLog('🗑️ 既存スプレッドシートIDをクリアして新規作成を強制します');

    // データベースレベルでもスプレッドシート情報をクリア（確実性を高める）
    updateUser(requestUserId, {
      spreadsheetId: '',
      spreadsheetUrl: '',
      configJson: JSON.stringify(configJson)
    });

    // userInfo オブジェクトもクリア
    userInfo.spreadsheetId = null;
    userInfo.spreadsheetUrl = null;

    // 更新後のキャッシュを強制同期
    synchronizeCacheAfterCriticalUpdate(requestUserId, userEmail, userInfo.spreadsheetId, null);
  } else {
    debugLog('✨ 初回セットアップを開始します');
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
    debugLog('🚀 QuickStart: セットアップ開始', { requestUserId });
    
    // ステップ0: セットアップコンテキストを初期化
    debugLog('📋 QuickStart: ステップ0 - コンテキスト初期化中...');
    var setupContext = initializeQuickStartContext(requestUserId);
    var configJson = setupContext.configJson;
    var userEmail = setupContext.userEmail;
    var userInfo = setupContext.userInfo;
    debugLog('✅ QuickStart: ステップ0完了 - コンテキスト初期化成功');

    // ステップ1-2: ファイル作成とフォルダ管理を実行
    debugLog('📁 QuickStart: ステップ1-2 - ファイル・フォルダ作成中...');
    var createdFiles = createQuickStartFiles(setupContext);
    var formAndSsInfo = createdFiles.formAndSsInfo;
    var folder = createdFiles.folder;
    debugLog('✅ QuickStart: ステップ1-2完了 - ファイル作成成功', {
      formId: formAndSsInfo.formId,
      spreadsheetId: formAndSsInfo.spreadsheetId,
      sheetName: formAndSsInfo.sheetName
    });

    // ステップ3: データベース更新とキャッシュ管理を実行
    debugLog('💾 QuickStart: ステップ3 - データベース更新中...');
    var updatedConfig = updateQuickStartDatabase(setupContext, createdFiles);
    debugLog('✅ QuickStart: ステップ3完了 - データベース更新成功');

    // ステップ4: 自動公開処理（重要な新機能）
    debugLog('🌐 QuickStart: ステップ4 - 自動公開処理中...');
    var publishResult = performAutoPublish(requestUserId, formAndSsInfo.sheetName);
    debugLog('✅ QuickStart: ステップ4完了 - 自動公開成功', publishResult);

    // ステップ5: キャッシュクリアと最終化
    debugLog('🔄 QuickStart: ステップ5 - キャッシュクリア中...');
    clearExecutionUserInfoCache();
    invalidateUserCache(requestUserId, userEmail, createdFiles.formAndSsInfo.spreadsheetId, true);
    clearExecutionUserInfoCache();
    debugLog('✅ QuickStart: ステップ5完了 - キャッシュクリア成功');

    // ステップ6: 最終レスポンス生成
    debugLog('📤 QuickStart: ステップ6 - レスポンス生成中...');
    var finalResponse = generateQuickStartResponse(setupContext, createdFiles, updatedConfig, publishResult);
    debugLog('🎉 QuickStart: 全工程完了！', finalResponse);
    
    return finalResponse;

  } catch (e) {
    errorLog('❌ quickStartSetup エラー: ' + e.message);

    // エラー時はセットアップ状態をリセット
    try {
      var currentConfig = JSON.parse(userInfo.configJson || '{}');
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
    var rootFolderName = "StudyQuest - ユーザーデータ";
    var userFolderName = "StudyQuest - " + userEmail + " - ファイル";

    // ルートフォルダを検索または作成
    var rootFolder;
    var folders = DriveApp.getFoldersByName(rootFolderName);
    if (folders.hasNext()) {
      rootFolder = folders.next();
    } else {
      rootFolder = DriveApp.createFolder(rootFolderName);
    }

    // ユーザー専用フォルダを作成
    var userFolders = rootFolder.getFoldersByName(userFolderName);
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
    var service = getSheetsServiceCached();
    var headerIndices = getHeaderIndices(spreadsheetId, sheetName);
    var highlightColumnIndex = headerIndices[COLUMN_HEADERS.HIGHLIGHT];

    if (highlightColumnIndex === undefined) {
      throw new Error('ハイライト列が見つかりません');
    }

    // 現在の値を取得
    var range = "'" + sheetName + "'!" + String.fromCharCode(65 + highlightColumnIndex) + rowIndex;
    var response = batchGetSheetsData(service, spreadsheetId, [range]);
    var isHighlighted = false;
    if (response && response.valueRanges && response.valueRanges[0] &&
        response.valueRanges[0].values && response.valueRanges[0].values[0] &&
        response.valueRanges[0].values[0][0]) {
      isHighlighted = response.valueRanges[0].values[0][0] === 'true';
    }

    // 値を切り替え
    var newValue = isHighlighted ? 'false' : 'true';
    updateSheetsData(service, spreadsheetId, range, [[newValue]]);

    // ハイライト更新後のキャッシュ無効化
    try {
      if (typeof cacheManager !== 'undefined' && typeof cacheManager.invalidateSheetData === 'function') {
        cacheManager.invalidateSheetData(spreadsheetId, sheetName);
        debugLog('ハイライト更新後のキャッシュ無効化完了: ' + spreadsheetId);
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
  debugLog('getHeaderIndices received in core.gs: spreadsheetId=%s, sheetName=%s', spreadsheetId, sheetName);

  var cacheKey = 'hdr_' + spreadsheetId + '_' + sheetName;
  var indices = getHeadersCached(spreadsheetId, sheetName);

  // 理由列が取得できていない場合はキャッシュを無効化して再取得
  if (!indices || indices[COLUMN_HEADERS.REASON] === undefined) {
    debugLog('getHeaderIndices: Reason header missing, refreshing cache for key=%s', cacheKey);
    cacheManager.remove(cacheKey);
    indices = getHeadersCached(spreadsheetId, sheetName);
  }

  return indices;
}

function getSheetColumns(userId, sheetId) {
  verifyUserAccess(userId);
  try {
    var userInfo = findUserById(userId);
    if (!userInfo || !userInfo.spreadsheetId) {
      throw new Error('ユーザー情報またはスプレッドシートIDが見つかりません');
    }

    var spreadsheet = openSpreadsheetOptimized(userInfo.spreadsheetId);
    var sheet = spreadsheet.getSheetById(sheetId);

    if (!sheet) {
      throw new Error('指定されたシートが見つかりません: ' + sheetId);
    }

    var lastColumn = sheet.getLastColumn();
    if (lastColumn === 0) {
      return []; // 列がない場合
    }

    var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];

    // クライアントサイドが期待する形式に変換
    var columns = headers.map(function(headerName) {
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
    var formIdMatch = url.match(/\/forms\/d\/([a-zA-Z0-9-_]+)/);
    if (formIdMatch && formIdMatch[1]) {
      return formIdMatch[1];
    }

    // Alternative pattern for e/ URLs
    var eFormIdMatch = url.match(/\/forms\/d\/e\/([a-zA-Z0-9-_]+)/);
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

      var service = getSheetsServiceCached();
      var headerIndices = getHeaderIndices(spreadsheetId, sheetName);

      // すべてのリアクション列を取得してユーザーの重複リアクションをチェック
      var allReactionRanges = [];
      var allReactionColumns = {};
      var targetReactionColumnIndex = null;

      // 全リアクション列の情報を準備
      REACTION_KEYS.forEach(function(key) {
        var columnName = COLUMN_HEADERS[key];
        var columnIndex = headerIndices[columnName];
        if (columnIndex !== undefined) {
          var range = "'" + sheetName + "'!" + String.fromCharCode(65 + columnIndex) + rowIndex;
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
      var response = batchGetSheetsData(service, spreadsheetId, allReactionRanges);
      var updateData = [];
      var userAction = null;
      var targetCount = 0;

      // 各リアクション列を処理
      var rangeIndex = 0;
      REACTION_KEYS.forEach(function(key) {
        if (!allReactionColumns[key]) return;

        var currentReactionString = '';
        if (response && response.valueRanges && response.valueRanges[rangeIndex] &&
            response.valueRanges[rangeIndex].values && response.valueRanges[rangeIndex].values[0] &&
            response.valueRanges[rangeIndex].values[0][0]) {
          currentReactionString = response.valueRanges[rangeIndex].values[0][0];
        }

        var currentReactions = parseReactionString(currentReactionString);
        var userIndex = currentReactions.indexOf(reactingUserEmail);

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
            debugLog('他のリアクションから削除: ' + reactingUserEmail + ' from ' + key);
          }
        }

        // 更新データを準備
        var updatedReactionString = currentReactions.join(', ');
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

      debugLog('リアクション切り替え完了: ' + reactingUserEmail + ' → ' + reactionKey + ' (' + userAction + ')', {
        updatedRanges: updateData.length,
        targetCount: targetCount,
        allColumns: Object.keys(allReactionColumns)
      });

      // リアクション更新後のキャッシュ無効化
      try {
        if (typeof cacheManager !== 'undefined' && typeof cacheManager.invalidateSheetData === 'function') {
          cacheManager.invalidateSheetData(spreadsheetId, sheetName);
          debugLog('リアクション更新後のキャッシュ無効化完了: ' + spreadsheetId);
        }
      } catch (cacheError) {
        warnLog('リアクション後のキャッシュ無効化エラー:', cacheError.message);
      }

      return {
        status: 'success',
        message: 'リアクションを更新しました。',
        action: userAction,
        count: targetCount
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
// NOTE: unpublishBoard関数の重複を回避するため、config.gsの実装を使用
// function unpublishBoard(requestUserId) {
//   verifyUserAccess(requestUserId);
//   try {
//     var currentUserId = requestUserId;
//     var userInfo = findUserById(currentUserId);
//     if (!userInfo) {
//       throw new Error('ユーザー情報が見つかりません');
//     }

//     var configJson = JSON.parse(userInfo.configJson || '{}');

//     configJson.publishedSpreadsheetId = '';
//     configJson.publishedSheetName = '';
//     configJson.appPublished = false; // 公開状態をfalseにする
//     configJson.setupStatus = 'completed'; // 公開停止後もセットアップは完了状態とする

//     updateUser(currentUserId, { configJson: JSON.stringify(configJson) });
//     invalidateUserCache(currentUserId, userInfo.adminEmail, userInfo.spreadsheetId, true);

//     infoLog('✅ 回答ボードの公開を停止しました: %s', currentUserId);
//     return { status: 'success', message: '回答ボードの公開を停止しました。' };
//   } catch (e) {
//     errorLog('公開停止エラー: ' + e.message);
//     return { status: 'error', message: '回答ボードの公開停止に失敗しました: ' + e.message };
//   }
// }

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
    var userEmail = options.userEmail;
    var userId = options.userId;
    var formDescription = options.formDescription || 'みんなの回答ボードへの投稿フォームです。';

    // タイムスタンプ生成（日時を含む）
    var now = new Date();
    var dateTimeString = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年MM月dd日 HH:mm');

    // フォームタイトル生成
    var formTitle = options.formTitle || ('みんなの回答ボード ' + dateTimeString);

    // フォーム作成
    var form = FormApp.create(formTitle);
    form.setDescription(formDescription);

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
    var spreadsheetResult = createLinkedSpreadsheet(userEmail, form, dateTimeString);

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
    var config = getQuestionConfig(questionType, customConfig);

    form.setCollectEmail(false);

    if (questionType === 'simple') {
      var classItem = form.addListItem();
      classItem.setTitle(config.classQuestion.title);
      classItem.setChoiceValues(config.classQuestion.choices);
      classItem.setRequired(true);

      var nameItem = form.addTextItem();
      nameItem.setTitle(config.nameQuestion.title);
      nameItem.setRequired(true);

      var mainItem = form.addTextItem();
      mainItem.setTitle(config.mainQuestion.title);
      mainItem.setRequired(true);

      var reasonItem = form.addParagraphTextItem();
      reasonItem.setTitle(config.reasonQuestion.title);
      reasonItem.setHelpText(config.reasonQuestion.helpText);
      var validation = FormApp.createParagraphTextValidation()
        .requireTextLengthLessThanOrEqualTo(140)
        .build();
      reasonItem.setValidation(validation);
      reasonItem.setRequired(false);
    } else if (questionType === 'custom' && customConfig) {
      debugLog('addUnifiedQuestions - custom mode with config:', JSON.stringify(customConfig));

      // クラス選択肢（有効な場合のみ）
      if (customConfig.enableClass && customConfig.classQuestion && customConfig.classQuestion.choices && customConfig.classQuestion.choices.length > 0) {
        var classItem = form.addListItem();
        classItem.setTitle('クラス');
        classItem.setChoiceValues(customConfig.classQuestion.choices);
        classItem.setRequired(true);
      }

      // 名前欄（常にオン）
      var nameItem = form.addTextItem();
      nameItem.setTitle('名前');
      nameItem.setRequired(false);

      // メイン質問
      var mainQuestionTitle = customConfig.mainQuestion ? customConfig.mainQuestion.title : '今回のテーマについて、あなたの考えや意見を聞かせてください';
      var mainItem;
      var questionType = customConfig.mainQuestion ? customConfig.mainQuestion.type : 'text';

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
      var reasonItem = form.addParagraphTextItem();
      reasonItem.setTitle('そう考える理由や体験があれば教えてください。');
      reasonItem.setRequired(false);
    } else {
      var classItem = form.addTextItem();
      classItem.setTitle(config.classQuestion.title);
      classItem.setHelpText(config.classQuestion.helpText);
      classItem.setRequired(true);

      var mainItem = form.addParagraphTextItem();
      mainItem.setTitle(config.mainQuestion.title);
      mainItem.setHelpText(config.mainQuestion.helpText);
      mainItem.setRequired(true);

      var reasonItem = form.addParagraphTextItem();
      reasonItem.setTitle(config.reasonQuestion.title);
      reasonItem.setHelpText(config.reasonQuestion.helpText);
      reasonItem.setRequired(false);

      var nameItem = form.addTextItem();
      nameItem.setTitle(config.nameQuestion.title);
      nameItem.setHelpText(config.nameQuestion.helpText);
      nameItem.setRequired(false);
    }

    debugLog('フォームに統一質問を追加しました: ' + questionType);

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
  var config = {
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
      title: '今日の学習について、あなたの考えや感想を聞かせてください',
      helpText: '',
      choices: ['気づいたことがある。', '疑問に思うことがある。', 'もっと知りたいことがある。'],
      type: 'paragraph'
    },
    reasonQuestion: {
      title: 'そう考える理由や体験があれば教えてください。',
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
    cfg.formTitle = `フォーム作成 - ${timestamp}`;

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
    var currentUserId = userId;
    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    var configJson = JSON.parse(userInfo.configJson || '{}');
    configJson.savedClassChoices = classChoices;
    configJson.lastClassChoicesUpdate = new Date().toISOString();

    updateUser(currentUserId, {
      configJson: JSON.stringify(configJson)
    });

    debugLog('クラス選択肢が保存されました:', classChoices);
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
    var currentUserId = userId;
    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      return { status: 'error', message: 'ユーザー情報が見つかりません' };
    }

    var configJson = JSON.parse(userInfo.configJson || '{}');
    var savedClassChoices = configJson.savedClassChoices || ['クラス1', 'クラス2', 'クラス3', 'クラス4'];

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
      mainQuestion: '今日の学習について、あなたの考えや感想を聞かせてください',
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
      throw new Error(`未知のプリセットタイプ: ${presetType}`);
    }

    const now = new Date();
    const dateTimeString = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年MM月dd日 HH:mm');

    // タイトル生成（上書き可能）
    const titlePrefix = overrides.titlePrefix || preset.titlePrefix;
    const formTitle = overrides.formTitle || `${titlePrefix} ${dateTimeString}`;

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
    errorLog(`createUnifiedForm Error (${presetType}):`, error.message);
    throw new Error(`フォーム作成に失敗しました (${presetType}): ` + error.message);
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
    var spreadsheetName = userEmail + ' - 回答データ - ' + dateTimeString;

    // 新しいスプレッドシートを作成
    var spreadsheetObj = SpreadsheetApp.create(spreadsheetName);
    var spreadsheetId = spreadsheetObj.getId();

    // スプレッドシートの共有設定を同一ドメイン閲覧可能に設定
    try {
      var file = DriveApp.getFileById(spreadsheetId);
      if (!file) {
        throw new Error('スプレッドシートが見つかりません: ' + spreadsheetId);
      }

      // 同一ドメインで閲覧可能に設定（教育機関対応）
      file.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.VIEW);
      debugLog('スプレッドシートを同一ドメイン閲覧可能に設定しました: ' + spreadsheetId);

      // 作成者（現在のユーザー）は所有者として保持
      debugLog('作成者は所有者として権限を保持: ' + userEmail);

    } catch (sharingError) {
      warnLog('共有設定の変更に失敗しましたが、処理を続行します: ' + sharingError.message);
    }

    // フォームの回答先としてスプレッドシートを設定
    form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheetId);

    // スプレッドシート名を再設定（フォーム設定後に名前が変わる場合があるため）
    spreadsheetObj.rename(spreadsheetName);

    // シート名を取得（通常は「フォームの回答 1」）
    var sheets = spreadsheetObj.getSheets();
    var sheetName = String(sheets[0].getName());
    // シート名が不正な値でないことを確認
    if (!sheetName || sheetName === 'true') {
      sheetName = 'Sheet1'; // または適切なデフォルト値
      warnLog('不正なシート名が検出されました。デフォルトのシート名を使用します: ' + sheetName);
    }

    // サービスアカウントとスプレッドシートを共有（失敗しても処理継続）
    try {
      shareSpreadsheetWithServiceAccount(spreadsheetId);
      debugLog('サービスアカウントとの共有完了: ' + spreadsheetId);
    } catch (shareError) {
      warnLog('サービスアカウント共有エラー（処理継続）:', shareError.message);
      // 権限エラーの場合でも、スプレッドシート作成自体は成功とみなす
      debugLog('スプレッドシート作成は完了しました。サービスアカウント共有は後で設定してください。');
    }

    return {
      spreadsheetId: spreadsheetId,
      spreadsheetUrl: spreadsheetObj.getUrl(),
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
    var serviceAccountEmail = getServiceAccountEmail();

    if (!serviceAccountEmail || serviceAccountEmail === 'サービスアカウント未設定' || serviceAccountEmail === 'サービスアカウント設定エラー') {
      throw new Error('サービスアカウントのメールアドレスが取得できません: ' + serviceAccountEmail);
    }

    debugLog('サービスアカウント共有開始:', serviceAccountEmail, 'スプレッドシート:', spreadsheetId);

    // DriveAppを使用してスプレッドシートをサービスアカウントと共有
    try {
      var file = DriveApp.getFileById(spreadsheetId);
      if (!file) {
        throw new Error('スプレッドシートが見つかりません: ' + spreadsheetId);
      }
      file.addEditor(serviceAccountEmail);
    } catch (driveError) {
      errorLog('DriveApp error:', driveError.message);
      throw new Error('Drive API操作に失敗しました: ' + driveError.message);
    }

    debugLog('サービスアカウント共有成功:', serviceAccountEmail);

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
    debugLog('全スプレッドシートのサービスアカウント共有開始');

    var allUsers = getAllUsers();
    var results = [];
    var successCount = 0;
    var errorCount = 0;

    for (var i = 0; i < allUsers.length; i++) {
      var user = allUsers[i];
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
          debugLog('共有成功:', user.adminEmail, user.spreadsheetId);
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

    debugLog('全スプレッドシート共有完了:', successCount + '件成功', errorCount + '件失敗');

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
    var props = PropertiesService.getScriptProperties();
    var serviceAccountCreds = JSON.parse(props.getProperty(SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS));
    var serviceAccountEmail = serviceAccountCreds.client_email;

    var spreadsheet = openSpreadsheetOptimized(spreadsheetId);

    // サービスアカウントを編集者として追加
    if (serviceAccountEmail) {
      spreadsheet.addEditor(serviceAccountEmail);
      debugLog('サービスアカウント (' + serviceAccountEmail + ') をスプレッドシートの編集者として追加しました。');

      // セッション管理：サービスアカウントアクセス権限の記録
      try {
        const sessionData = {
          serviceAccountEmail: serviceAccountEmail,
          spreadsheetId: spreadsheetId,
          accessGranted: new Date().toISOString(),
          accessType: 'service_account_editor',
          securityLevel: 'domain_view'
        };
        debugLog('サービスアカウントアクセス権限を記録しました:', JSON.stringify(sessionData));
      } catch (sessionLogError) {
        warnLog('セッション記録でエラー:', sessionLogError.message);
      }
    }

    // 同一ドメインユーザーは共有設定により閲覧可能
    debugLog('同一ドメインユーザーは共有設定により閲覧可能です');

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
    debugLog('スプレッドシートアクセス権限の修復を開始: ' + userEmail + ' -> ' + spreadsheetId);

    // DriveApp経由で共有設定を変更
    var file;
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
      debugLog('スプレッドシートをドメイン全体で編集可能に設定しました');
    } catch (domainSharingError) {
      warnLog('ドメイン共有設定に失敗: ' + domainSharingError.message);

      // ドメイン共有に失敗した場合は個別にユーザーを追加
      try {
        file.addEditor(userEmail);
        debugLog('ユーザーを個別に編集者として追加しました: ' + userEmail);
      } catch (individualError) {
        errorLog('個別ユーザー追加も失敗: ' + individualError.message);
      }
    }

    // SpreadsheetApp経由でも編集者として追加
    try {
      var spreadsheet = openSpreadsheetOptimized(spreadsheetId);
      spreadsheet.addEditor(userEmail);
      debugLog('SpreadsheetApp経由でユーザーを編集者として追加: ' + userEmail);
    } catch (spreadsheetAddError) {
      warnLog('SpreadsheetApp経由の追加で警告: ' + spreadsheetAddError.message);
    }

    // サービスアカウントも確認
    var props = PropertiesService.getScriptProperties();
    var serviceAccountCreds = JSON.parse(props.getProperty(SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS));
    var serviceAccountEmail = serviceAccountCreds.client_email;

    if (serviceAccountEmail) {
      try {
        file.addEditor(serviceAccountEmail);
        debugLog('サービスアカウントも編集者として追加: ' + serviceAccountEmail);
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
    var spreadsheet = openSpreadsheetOptimized(spreadsheetId);
    var sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.getSheets()[0];

    var additionalHeaders = [
      COLUMN_HEADERS.UNDERSTAND,
      COLUMN_HEADERS.LIKE,
      COLUMN_HEADERS.CURIOUS,
      COLUMN_HEADERS.HIGHLIGHT
    ];

    // 効率的にヘッダー情報を取得
    var currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var startCol = currentHeaders.length + 1;

    // バッチでヘッダーを追加
    sheet.getRange(1, startCol, 1, additionalHeaders.length).setValues([additionalHeaders]);

    // スタイリングを一括適用
    var allHeadersRange = sheet.getRange(1, 1, 1, currentHeaders.length + additionalHeaders.length);
    allHeadersRange.setFontWeight('bold').setBackground('#E3F2FD');

    // 自動リサイズ（エラーが出ても続行）
    try {
      sheet.autoResizeColumns(1, allHeadersRange.getNumColumns());
    } catch (resizeError) {
      warnLog('Auto-resize failed:', resizeError.message);
    }

    debugLog('リアクション列を追加しました: ' + sheetName);
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
  var cacheKey = `sheetData_${userId}_${sheetName}_${classFilter}_${sortMode}`;

  // 管理モードの場合はキャッシュをバイパス（最新データを取得）
  if (adminMode === true) {
    debugLog('🔄 管理モード：シートデータキャッシュをバイパス');
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
      var userInfo = getOrFetchUserInfo(userId, 'userId', {
    useExecutionCache: true,
    ttl: 300
  });
      if (!userInfo) {
        throw new Error('ユーザー情報が見つかりません');
      }

      var spreadsheetId = userInfo.spreadsheetId;
      var service = getSheetsServiceCached();

      // フォーム回答データのみを取得（名簿機能は使用しない）
      var ranges = [sheetName + '!A:Z'];

      var responses = batchGetSheetsData(service, spreadsheetId, ranges);
      debugLog('DEBUG: batchGetSheetsData responses: %s', JSON.stringify(responses));
      var sheetData = responses.valueRanges[0].values || [];
      debugLog('DEBUG: sheetData length: %s', sheetData.length);

    // 名簿機能は使用せず、空の配列を設定
    var rosterData = [];


    if (sheetData.length === 0) {
      return {
        status: 'success',
        data: [],
        headers: [],
        totalCount: 0
      };
    }

    var headers = sheetData[0];
    var dataRows = sheetData.slice(1);

    // ヘッダーインデックスを取得（キャッシュ利用）
    var headerIndices = getHeaderIndices(spreadsheetId, sheetName);

    // 名簿マップを作成（キャッシュ利用）
    var rosterMap = buildRosterMap(rosterData);

    // 表示モードとシート固有設定を取得
    var configJson = JSON.parse(userInfo.configJson || '{}');
    var displayMode = configJson.displayMode || DISPLAY_MODES.ANONYMOUS;

    // シート固有の設定を取得（最新のAI判定結果を反映）
    var sheetKey = 'sheet_' + sheetName;
    var sheetConfig = configJson[sheetKey] || {};

    // AI判定結果またはguessedConfigがある場合、それを優先使用
    var effectiveHeaderConfig = sheetConfig.guessedConfig || sheetConfig || {};
    debugLog('🔍 executeGetSheetData: シート設定取得完了 sheetKey=%s, hasGuessedConfig=%s', sheetKey, !!effectiveHeaderConfig.opinionHeader);

    // フォールバック設定: 設定が空の場合はデフォルト構造を提供
    if (!effectiveHeaderConfig.opinionHeader && headers.length > 1) {
      effectiveHeaderConfig.opinionHeader = headers[1]; // 通常、タイムスタンプの次が最初の質問
      debugLog('🔄 executeGetSheetData: フォールバック設定 - opinionHeader: %s', effectiveHeaderConfig.opinionHeader);
    }
    if (!effectiveHeaderConfig.reasonHeader && headers.length > 2) {
      effectiveHeaderConfig.reasonHeader = headers[2]; // 2番目の質問を理由として設定
      debugLog('🔄 executeGetSheetData: フォールバック設定 - reasonHeader: %s', effectiveHeaderConfig.reasonHeader);
    }

    // 統合デバッグ: 設定とデータの詳細ログ
    debugLog('🔍 executeGetSheetData: 統合デバッグ情報', {
      headers: headers,
      headersLength: headers.length,
      configJson: configJson,
      sheetKey: sheetKey,
      sheetConfig: sheetConfig,
      effectiveHeaderConfig: effectiveHeaderConfig,
      dataRowsLength: dataRows.length,
      firstRowSample: dataRows.length > 0 ? dataRows[0] : 'なし'
    });

    // Check if current user is the board owner
    var isOwner = (configJson.ownerId === userId);
    debugLog('getSheetData: isOwner=%s, ownerId=%s, userId=%s', isOwner, configJson.ownerId, userId);

    // データを処理
    var processedData = dataRows.map(function(row, index) {
      return processRowData(row, headers, headerIndices, rosterMap, displayMode, index + 2, isOwner);
    });

    // フィルタリング
    var filteredData = processedData;
    if (classFilter && classFilter !== 'すべて') {
      var classIndex = headerIndices[COLUMN_HEADERS.CLASS];
      if (classIndex !== undefined) {
        filteredData = processedData.filter(function(row) {
          return row.originalData[classIndex] === classFilter;
        });
      }
    }

    // ソート適用
    var sortedData = applySortMode(filteredData, sortMode || 'newest');

    // カスタムフォーム設定がある場合のヘッダー情報を優先
    var effectiveHeaders = headers;
    var mainQuestionHeader = headers[0]; // デフォルトは最初の列をメイン質問とする

    // AI判定結果またはシート設定からメイン質問を特定
    if (effectiveHeaderConfig.opinionHeader) {
      var opinionIndex = headers.indexOf(effectiveHeaderConfig.opinionHeader);
      if (opinionIndex !== -1) {
        mainQuestionHeader = effectiveHeaderConfig.opinionHeader;
        debugLog('🎯 executeGetSheetData: AI判定結果からメインヘッダー設定 - %s', mainQuestionHeader);
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
      var userInfo = findUserById(userId);
      var configJson = JSON.parse(userInfo.configJson || '{}');
      var sheetKey = 'sheet_' + sheetName;
      var sheetConfig = configJson[sheetKey] || {};
      var effectiveHeaderConfig = sheetConfig.guessedConfig || sheetConfig || {};

      // 設定からメインヘッダーを復元
      var fallbackHeader = effectiveHeaderConfig.opinionHeader || sheetName;

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
    debugLog('getSheetsList: Start for userId:', userId);
    var userInfo = findUserById(userId);
    if (!userInfo) {
      warnLog('getSheetsList: User not found:', userId);
      return [];
    }

    debugLog('getSheetsList: UserInfo found:', {
      userId: userInfo.userId,
      adminEmail: userInfo.adminEmail,
      spreadsheetId: userInfo.spreadsheetId,
      spreadsheetUrl: userInfo.spreadsheetUrl
    });

    if (!userInfo.spreadsheetId) {
      warnLog('getSheetsList: No spreadsheet ID for user:', userId);
      return [];
    }

    debugLog('getSheetsList: User\'s spreadsheetId:', userInfo.spreadsheetId);

    var service = getSheetsServiceCached();
    if (!service) {
      errorLog('❌ getSheetsList: Sheets service not initialized');
      return [];
    }

    infoLog('✅ getSheetsList: Service validated successfully');

    debugLog('getSheetsList: SheetsService obtained, attempting to fetch spreadsheet data...');

    var spreadsheet;
    try {
      spreadsheet = getSpreadsheetsData(service, userInfo.spreadsheetId);
    } catch (accessError) {
      warnLog('getSheetsList: アクセスエラーを検出。サービスアカウント権限を修復中...', accessError.message);

      // サービスアカウントの権限修復を試行
      try {
        addServiceAccountToSpreadsheet(userInfo.spreadsheetId);
        debugLog('getSheetsList: サービスアカウント権限を追加しました。再試行中...');

        // 少し待ってから再試行
        Utilities.sleep(1000);
        spreadsheet = getSpreadsheetsData(service, userInfo.spreadsheetId);

      } catch (repairError) {
        errorLog('getSheetsList: 権限修復に失敗:', repairError.message);

        // 最終手段：ユーザー権限での修復も試行
        try {
          var currentUserEmail = Session.getActiveUser().getEmail();
          if (currentUserEmail === userInfo.adminEmail) {
            repairUserSpreadsheetAccess(currentUserEmail, userInfo.spreadsheetId);
            debugLog('getSheetsList: ユーザー権限での修復を実行しました。');
          }
        } catch (finalRepairError) {
          errorLog('getSheetsList: 最終修復も失敗:', finalRepairError.message);
        }

        return [];
      }
    }

    debugLog('getSheetsList: Raw spreadsheet data:', spreadsheet);
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

    var sheets = spreadsheet.sheets.map(function(sheet) {
      if (!sheet.properties) {
        warnLog('getSheetsList: Sheet missing properties:', sheet);
        return null;
      }
      return {
        name: sheet.properties.title,
        id: sheet.properties.sheetId
      };
    }).filter(function(sheet) { return sheet !== null; });

    debugLog('getSheetsList: Successfully returning', sheets.length, 'sheets:', sheets);
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
  var processedRow = {
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
    var columnName = COLUMN_HEADERS[reactionKey];
    var columnIndex = headerIndices[columnName];

    if (columnIndex !== undefined && row[columnIndex]) {
      var reactions = parseReactionString(row[columnIndex]);
      var count = reactions.length;

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
  var highlightIndex = headerIndices[COLUMN_HEADERS.HIGHLIGHT];
  if (highlightIndex !== undefined && row[highlightIndex]) {
    processedRow.isHighlighted = row[highlightIndex].toString().toLowerCase() === 'true';
  }

  // スコア計算
  processedRow.score = calculateRowScore(processedRow);

  // 名前の表示処理（フォーム入力の名前を使用）
  var nameIndex = headerIndices[COLUMN_HEADERS.NAME];
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
  var emailIndex = headerIndices[COLUMN_HEADERS.EMAIL];
  if (emailIndex !== undefined && row[emailIndex]) {
    processedRow.email = row[emailIndex];
  }

  // タイムスタンプ（フォーム回答時刻）
  var timestampIndex = headerIndices[COLUMN_HEADERS.TIMESTAMP];
  if (timestampIndex !== undefined && row[timestampIndex]) {
    processedRow.timestamp = row[timestampIndex];
  }

  return processedRow;
}

/**
 * 行のスコアを計算
 */
function calculateRowScore(rowData) {
  var baseScore = 1.0;

  // いいね！による加算
  var likeBonus = rowData.likeCount * SCORING_CONFIG.LIKE_MULTIPLIER_FACTOR;

  // その他のリアクションも軽微な加算
  var reactionBonus = (rowData.understandCount + rowData.curiousCount) * 0.01;

  // ハイライトによる大幅加算
  var highlightBonus = rowData.isHighlighted ? 0.5 : 0;

  // ランダム要素（同じスコアの項目をランダムに並べるため）
  var randomFactor = Math.random() * SCORING_CONFIG.RANDOM_SCORE_FACTOR;

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
    var j = Math.floor(Math.random() * (i + 1));
    var temp = array[i];
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
  var mappedIndices = {};
  var availableHeaders = Object.keys(actualHeaderIndices);
  debugLog('mapConfigToActualHeaders: Available headers in spreadsheet: %s', JSON.stringify(availableHeaders));

  // 各設定ヘッダーでマッピングを試行
  for (var configKey in configHeaders) {
    var configHeaderName = configHeaders[configKey];
    var mappedIndex = undefined;

    debugLog('mapConfigToActualHeaders: Trying to map %s = "%s"', configKey, configHeaderName);

    if (configHeaderName && actualHeaderIndices.hasOwnProperty(configHeaderName)) {
      // 完全一致
      mappedIndex = actualHeaderIndices[configHeaderName];
      debugLog('mapConfigToActualHeaders: Exact match found for %s: index %s', configKey, mappedIndex);
    } else if (configHeaderName) {
      // 部分一致で検索（大文字小文字を区別しない）
      var normalizedConfigName = configHeaderName.toLowerCase().trim();

      for (var actualHeader in actualHeaderIndices) {
        var normalizedActualHeader = actualHeader.toLowerCase().trim();
        if (normalizedActualHeader === normalizedConfigName) {
          mappedIndex = actualHeaderIndices[actualHeader];
          debugLog('mapConfigToActualHeaders: Case-insensitive match found for %s: "%s" -> index %s', configKey, actualHeader, mappedIndex);
          break;
        }
      }

      // 部分一致検索
      if (mappedIndex === undefined) {
        for (var actualHeader in actualHeaderIndices) {
          var normalizedActualHeader = actualHeader.toLowerCase().trim();
          if (normalizedActualHeader.includes(normalizedConfigName) || normalizedConfigName.includes(normalizedActualHeader)) {
            mappedIndex = actualHeaderIndices[actualHeader];
            debugLog('mapConfigToActualHeaders: Partial match found for %s: "%s" -> index %s', configKey, actualHeader, mappedIndex);
            break;
          }
        }
      }
    }

    // opinionHeader（メイン質問）の特別処理：見つからない場合は最も長い質問様ヘッダーを使用
    if (mappedIndex === undefined && configKey === 'opinionHeader') {
      var standardHeaders = ['タイムスタンプ', 'メールアドレス', 'クラス', '名前', '理由', 'なるほど！', 'いいね！', 'もっと知りたい！', 'ハイライト'];
      var questionHeaders = [];

      for (var header in actualHeaderIndices) {
        var isStandardHeader = false;
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
        var longestHeader = questionHeaders.reduce(function(prev, current) {
          return (prev.header.length > current.header.length) ? prev : current;
        });
        mappedIndex = longestHeader.index;
        debugLog('mapConfigToActualHeaders: Auto-detected main question header for %s: "%s" -> index %s', configKey, longestHeader.header, mappedIndex);
      }
    }

    // reasonHeader（理由列）の特別処理：見つからない場合は理由らしいヘッダーを自動検出
    if (mappedIndex === undefined && configKey === 'reasonHeader') {
      var reasonKeywords = ['理由', 'reason', 'なぜ', 'why', '根拠', 'わけ'];

      debugLog('mapConfigToActualHeaders: Searching for reason header with keywords: %s', JSON.stringify(reasonKeywords));

      for (var header in actualHeaderIndices) {
        var normalizedHeader = header.toLowerCase().trim();
        for (var k = 0; k < reasonKeywords.length; k++) {
          if (normalizedHeader.includes(reasonKeywords[k]) || reasonKeywords[k].includes(normalizedHeader)) {
            mappedIndex = actualHeaderIndices[header];
            debugLog('mapConfigToActualHeaders: Auto-detected reason header for %s: "%s" -> index %s (keyword: %s)', configKey, header, mappedIndex, reasonKeywords[k]);
            break;
          }
        }
        if (mappedIndex !== undefined) break;
      }

      // より広範囲の検索（部分一致）
      if (mappedIndex === undefined) {
        for (var header in actualHeaderIndices) {
          var normalizedHeader = header.toLowerCase().trim();
          if (normalizedHeader.indexOf('理由') !== -1 || normalizedHeader.indexOf('reason') !== -1) {
            mappedIndex = actualHeaderIndices[header];
            debugLog('mapConfigToActualHeaders: Found reason header by partial match for %s: "%s" -> index %s', configKey, header, mappedIndex);
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
        debugLog('mapConfigToActualHeaders: Using fallback position for %s: index %s', configKey, mappedIndex);
      } else if (configKey === 'reasonHeader') {
        // 理由列が見つからない場合、2番目の列を使用
        mappedIndex = 2;
        debugLog('mapConfigToActualHeaders: Using fallback position for %s: index %s', configKey, mappedIndex);
      }
    }

    mappedIndices[configKey] = mappedIndex;

    if (mappedIndex === undefined) {
      debugLog('mapConfigToActualHeaders: WARNING - No match found for %s = "%s"', configKey, configHeaderName);
    }
  }

  debugLog('mapConfigToActualHeaders: Final mapping result: %s', JSON.stringify(mappedIndices));
  return mappedIndices;
}

/**
 * 特定の行のリアクション情報を取得
 */
function getRowReactions(spreadsheetId, sheetName, rowIndex, userEmail) {
  try {
    var service = getSheetsServiceCached();
    var headerIndices = getHeaderIndices(spreadsheetId, sheetName);

    var reactionData = {
      UNDERSTAND: { count: 0, reacted: false },
      LIKE: { count: 0, reacted: false },
      CURIOUS: { count: 0, reacted: false }
    };

    // 各リアクション列からデータを取得
    ['UNDERSTAND', 'LIKE', 'CURIOUS'].forEach(function(reactionKey) {
      var columnName = COLUMN_HEADERS[reactionKey];
      var columnIndex = headerIndices[columnName];

      if (columnIndex !== undefined) {
        var range = sheetName + '!' + String.fromCharCode(65 + columnIndex) + rowIndex;
        try {
          var response = batchGetSheetsData(service, spreadsheetId, [range]);
          var cellValue = '';
          if (response && response.valueRanges && response.valueRanges[0] &&
              response.valueRanges[0].values && response.valueRanges[0].values[0] &&
              response.valueRanges[0].values[0][0]) {
            cellValue = response.valueRanges[0].values[0][0];
          }

          if (cellValue) {
            var reactions = parseReactionString(cellValue);
            reactionData[reactionKey].count = reactions.length;
            reactionData[reactionKey].reacted = reactions.indexOf(userEmail) !== -1;
          }
        } catch (cellError) {
          warnLog('リアクション取得エラー(' + reactionKey + '): ' + cellError.message);
        }
      }
    });

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
    var activeUserEmail = Session.getActiveUser().getEmail();
    if (!activeUserEmail) {
      return {
        status: 'error',
        message: 'ユーザーが認証されていません'
      };
    }

    // 現在のユーザー情報を取得
    var userInfo = findUserByEmail(activeUserEmail);
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
    var newIsActiveValue = isActive ? 'true' : 'false';
    var updateResult = updateUser(userInfo.userId, {
      isActive: newIsActiveValue,
      lastAccessedAt: new Date().toISOString()
    });

    if (updateResult.success) {
      var statusMessage = isActive
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
    var activeUserEmail = Session.getActiveUser().getEmail();
    if (!activeUserEmail) {
      return false;
    }

    // データベースに登録され、かつisActiveがtrueのユーザーのみアクセス可能
    var userInfo = findUserByEmail(activeUserEmail);
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
  var accessToken = getServiceAccountTokenCached();
  return {
    accessToken: accessToken,
    baseUrl: 'https://www.googleapis.com/drive/v3',
    files: {
      get: function(params) {
        var url = this.baseUrl + '/files/' + params.fileId + '?fields=' + encodeURIComponent(params.fields);
        var response = UrlFetchApp.fetch(url, {
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
  return isSystemAdmin();
}

/**
 * デプロイユーザーかどうか判定
 * 1. GASスクリプトの編集権限を確認
 * 2. 教育機関ドメインの管理者権限を確認
 * @returns {boolean} 管理者権限を持つ場合 true
 */
function isSystemAdmin() {
  try {
    var props = PropertiesService.getScriptProperties();
    var adminEmail = props.getProperty(SCRIPT_PROPS_KEYS.ADMIN_EMAIL);
    var currentUserEmail = Session.getActiveUser().getEmail();
    return adminEmail && currentUserEmail && adminEmail === currentUserEmail;
  } catch (e) {
    errorLog('isSystemAdmin エラー: ' + e.message);
    return false;
  }
}

function isDeployUser() {
  return isSystemAdmin();
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
function deleteUserAccountByAdminForUI(targetUserId, reason) {
  try {
    const result = deleteUserAccountByAdmin(targetUserId, reason);
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
 * 統合カスタムセットアップ処理（QuickStart同様の完全自動化）
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @param {object} config - カスタムフォーム設定
 * @returns {object} セットアップ結果
 */
function customSetup(requestUserId, config) {
  verifyUserAccess(requestUserId);
  try {
    debugLog('🎨 CustomSetup: 統合セットアップ開始', { requestUserId, config });
    
    // ステップ1: カスタムフォーム設定の解析と検証
    debugLog('📋 CustomSetup: ステップ1 - 設定解析中...');
    var setupContext = initializeCustomSetupContext(requestUserId, config);
    debugLog('✅ CustomSetup: ステップ1完了 - 設定解析成功');

    // ステップ2: Googleフォームとスプレッドシートの作成
    debugLog('📝 CustomSetup: ステップ2 - フォーム・スプレッドシート作成中...');
    var createdFiles = createCustomFormAndSheet(setupContext);
    var formAndSsInfo = createdFiles.formAndSsInfo;
    var folder = createdFiles.folder;
    debugLog('✅ CustomSetup: ステップ2完了 - ファイル作成成功', {
      formId: formAndSsInfo.formId,
      spreadsheetId: formAndSsInfo.spreadsheetId,
      sheetName: formAndSsInfo.sheetName
    });

    // ステップ3: 高精度AI列判定の自動実行
    debugLog('🤖 CustomSetup: ステップ3 - AI列判定実行中...');
    var aiDetectionResult = performAutoAIDetection(requestUserId, formAndSsInfo.spreadsheetId, formAndSsInfo.sheetName);
    debugLog('✅ CustomSetup: ステップ3完了 - AI列判定成功', aiDetectionResult);

    // ステップ4: 設定の自動保存と適用
    debugLog('💾 CustomSetup: ステップ4 - 設定保存中...');
    var saveResult = applyAutoConfiguration(requestUserId, formAndSsInfo.spreadsheetId, formAndSsInfo.sheetName, aiDetectionResult);
    debugLog('✅ CustomSetup: ステップ4完了 - 設定保存成功');

    // ステップ5: 回答ボードの自動公開
    debugLog('🌐 CustomSetup: ステップ5 - 自動公開処理中...');
    var publishResult = performAutoPublish(requestUserId, formAndSsInfo.sheetName);
    debugLog('✅ CustomSetup: ステップ5完了 - 自動公開成功', publishResult);

    // ステップ6: キャッシュクリアと最終化
    debugLog('🔄 CustomSetup: ステップ6 - 最終化処理中...');
    clearExecutionUserInfoCache();
    invalidateUserCache(requestUserId, setupContext.userEmail, formAndSsInfo.spreadsheetId, true);
    clearExecutionUserInfoCache();
    debugLog('✅ CustomSetup: ステップ6完了 - 最終化成功');

    // ステップ7: 最終レスポンス生成と完了通知
    debugLog('📤 CustomSetup: ステップ7 - レスポンス生成中...');
    var finalResponse = generateCustomSetupResponse(setupContext, createdFiles, saveResult, aiDetectionResult, publishResult);
    debugLog('🎉 CustomSetup: 全工程完了！', finalResponse);
    
    return finalResponse;

  } catch (e) {
    errorLog('❌ customSetup エラー: ' + e.message);

    // エラー時はセットアップ状態をリセット
    try {
      var userInfo = findUserById(requestUserId);
      if (userInfo) {
        var currentConfig = JSON.parse(userInfo.configJson || '{}');
        currentConfig.setupStatus = 'error';
        currentConfig.lastError = e.message;
        currentConfig.errorAt = new Date().toISOString();

        updateUser(requestUserId, {
          configJson: JSON.stringify(currentConfig)
        });
        invalidateUserCache(requestUserId, userInfo.adminEmail, null, false);
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
 * カスタムセットアップのコンテキストを初期化
 * @param {string} requestUserId - ユーザーID
 * @param {object} config - カスタムフォーム設定
 * @returns {object} セットアップコンテキスト
 */
function initializeCustomSetupContext(requestUserId, config) {
  debugLog('🚀 カスタムセットアップのコンテキスト初期化開始: ' + requestUserId);

  // ユーザー情報の取得
  var userInfo = findUserById(requestUserId);
  if (!userInfo) {
    throw new Error('ユーザー情報が見つかりません');
  }

  var userEmail = userInfo.adminEmail;
  var configJson = JSON.parse(userInfo.configJson || '{}');

  // カスタム設定の詳細ログ
  debugLog('📋 カスタム設定詳細:', {
    mainQuestion: config.mainQuestion,
    responseType: config.responseType,
    enableClass: config.enableClass,
    classChoices: config.classChoices
  });

  return {
    requestUserId: requestUserId,
    userInfo: userInfo,
    userEmail: userEmail,
    configJson: configJson,
    customConfig: config
  };
}

/**
 * カスタムフォームとスプレッドシート作成の統合処理
 * @param {object} setupContext - セットアップコンテキスト
 * @returns {object} 作成されたファイル情報
 */
function createCustomFormAndSheet(setupContext) {
  var userEmail = setupContext.userEmail;
  var requestUserId = setupContext.requestUserId;
  var config = setupContext.customConfig;

  // ステップ1: ユーザー専用フォルダを作成
  debugLog('📁 カスタムセットアップ: フォルダ作成中...');
  var folder = createUserFolder(userEmail);

  // ステップ2: AdminPanelのconfig構造を内部形式に変換
  const convertedConfig = {
    mainQuestion: {
      title: config.mainQuestion || '今日の学習について、あなたの考えや感想を聞かせてください',
      type: config.responseType || config.questionType || 'text',
      choices: config.questionChoices || config.choices || [],
      includeOthers: config.includeOthers || false
    },
    enableClass: config.enableClass || false,
    classQuestion: {
      choices: config.classChoices || ['クラス1', 'クラス2', 'クラス3', 'クラス4']
    }
  };

  const overrides = {
    titlePrefix: config.formTitle || 'カスタムフォーム',
    customConfig: convertedConfig
  };

  // ステップ3: 統合フォーム作成
  const formAndSsInfo = createUnifiedForm('custom', userEmail, requestUserId, overrides);

  // ステップ4: 作成したファイルをフォルダに移動
  if (folder) {
    moveFilesToFolder(folder, formAndSsInfo.formId, formAndSsInfo.spreadsheetId);
  }

  return {
    formAndSsInfo: formAndSsInfo,
    folder: folder
  };
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
    debugLog('🤖 AI列判定自動実行開始', { requestUserId, spreadsheetId, sheetName });
    
    // getSheetDetailsを使用してAI判定を実行
    const sheetDetails = getSheetDetails(requestUserId, spreadsheetId, sheetName);
    
    if (!sheetDetails || !sheetDetails.guessedConfig) {
      throw new Error('AI列判定の実行に失敗しました');
    }
    
    infoLog('✅ AI列判定自動実行完了', {
      opinionColumn: sheetDetails.guessedConfig.opinionColumn,
      nameColumn: sheetDetails.guessedConfig.nameColumn,
      classColumn: sheetDetails.guessedConfig.classColumn
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
 * 自動設定適用処理
 * @param {string} requestUserId - ユーザーID
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @param {object} aiDetectionResult - AI判定結果
 * @returns {object} 保存結果
 */
function applyAutoConfiguration(requestUserId, spreadsheetId, sheetName, aiDetectionResult) {
  try {
    debugLog('💾 自動設定適用開始', { requestUserId, sheetName });
    
    if (!aiDetectionResult.success || !aiDetectionResult.guessedConfig) {
      throw new Error('AI判定結果が無効です');
    }
    
    // AI判定結果を設定として保存
    const saveResult = saveSheetConfig(requestUserId, spreadsheetId, sheetName, aiDetectionResult.guessedConfig);
    
    if (!saveResult || !saveResult.success) {
      throw new Error('設定保存に失敗しました: ' + (saveResult?.message || 'unknown error'));
    }
    
    infoLog('✅ 自動設定適用完了', saveResult);
    
    return {
      success: true,
      configured: true,
      savedConfig: aiDetectionResult.guessedConfig,
      message: '設定が正常に保存されました',
      details: saveResult
    };
    
  } catch (error) {
    errorLog('❌ 自動設定適用エラー', { error: error.message });
    return {
      success: false,
      configured: false,
      message: '設定適用に失敗しました: ' + error.message,
      error: error.message
    };
  }
}

/**
 * カスタムセットアップの最終レスポンス生成
 * @param {object} setupContext - セットアップコンテキスト
 * @param {object} createdFiles - 作成されたファイル情報
 * @param {object} saveResult - 保存結果
 * @param {object} aiDetectionResult - AI判定結果
 * @param {object} publishResult - 公開結果
 * @returns {object} 最終レスポンス
 */
function generateCustomSetupResponse(setupContext, createdFiles, saveResult, aiDetectionResult, publishResult) {
  var requestUserId = setupContext.requestUserId;
  var formAndSsInfo = createdFiles.formAndSsInfo;

  // 最終検証：新規作成されたファイルの確認
  debugLog('🔍 カスタムセットアップ最終検証 - 作成されたファイル:');
  debugLog('  📝 フォームID:', formAndSsInfo.formId);
  debugLog('  📊 スプレッドシートID:', formAndSsInfo.spreadsheetId);
  debugLog('  📄 シート名:', formAndSsInfo.sheetName);

  // AI判定と公開結果の検証
  var isAIDetected = aiDetectionResult && aiDetectionResult.success && aiDetectionResult.aiDetected;
  var isConfigured = saveResult && saveResult.success && saveResult.configured;
  var isPublished = publishResult && publishResult.success && publishResult.published;
  
  var setupMessage = 'カスタムセットアップが完了しました！';
  if (isAIDetected && isConfigured && isPublished) {
    setupMessage += 'AI列判定、設定保存、自動公開がすべて正常に完了しました。';
  } else if (isAIDetected && isConfigured) {
    setupMessage += 'AI列判定と設定保存が完了しました。管理パネルから手動で公開してください。';
  } else {
    setupMessage += 'フォームが作成されました。管理パネルで設定を確認してください。';
  }

  infoLog('✅ カスタムセットアップ完了: ' + requestUserId);

  var appUrls = generateUserUrls(requestUserId);
  
  // 拡張されたレスポンス情報
  var response = {
    status: 'success',
    message: setupMessage,
    webAppUrl: appUrls.webAppUrl,
    adminUrl: appUrls.adminUrl,
    viewUrl: appUrls.viewUrl,
    setupUrl: appUrls.setupUrl,
    formUrl: formAndSsInfo.formUrl,
    editFormUrl: formAndSsInfo.editFormUrl,
    spreadsheetUrl: formAndSsInfo.spreadsheetUrl,
    folderUrl: createdFiles.folder ? createdFiles.folder.getUrl() : '',
    // カスタムセットアップ固有の詳細情報
    setupComplete: true,
    aiDetected: isAIDetected,
    autoConfigured: isConfigured,
    autoPublished: isPublished,
    aiDetectionResult: aiDetectionResult,
    saveResult: saveResult,
    publishResult: publishResult,
    sheetName: formAndSsInfo.sheetName,
    formId: formAndSsInfo.formId,
    spreadsheetId: formAndSsInfo.spreadsheetId,
    // フロントエンド完了通知用のタイムスタンプ
    completedAt: new Date().toISOString(),
    // 成功ステップの詳細
    completedSteps: [
      'カスタムフォーム設定の解析',
      'Googleフォームとスプレッドシートの作成',
      isAIDetected ? '高精度AI列判定の実行' : 'AI列判定（失敗）',
      isConfigured ? '設定の自動保存と適用' : '設定保存（失敗）',
      isPublished ? '回答ボードの自動公開' : '回答ボード作成（手動公開待ち）',
      'キャッシュクリアと最終化',
      'カスタムセットアップ完了'
    ]
  };
  
  debugLog('📤 カスタムセットアップ最終レスポンス生成完了:', response);
  return response;
}

/**
 * ファイルをフォルダに移動する処理
 * @param {object} folder - 移動先フォルダ
 * @param {string} formId - フォームID
 * @param {string} spreadsheetId - スプレッドシートID
 */
function moveFilesToFolder(folder, formId, spreadsheetId) {
  try {
    var formFile = DriveApp.getFileById(formId);
    var ssFile = DriveApp.getFileById(spreadsheetId);

    // 既にフォルダに存在するかチェック
    var isFormInFolder = false;
    var isSSInFolder = false;

    var formParents = formFile.getParents();
    while (formParents.hasNext()) {
      if (formParents.next().getId() === folder.getId()) {
        isFormInFolder = true;
        break;
      }
    }

    var ssParents = ssFile.getParents();
    while (ssParents.hasNext()) {
      if (ssParents.next().getId() === folder.getId()) {
        isSSInFolder = true;
        break;
      }
    }

    // 必要に応じてファイルを移動
    if (!isFormInFolder) {
      folder.addFile(formFile);
      DriveApp.getRootFolder().removeFile(formFile);
      debugLog('📁 フォームファイルをフォルダに移動完了');
    }

    if (!isSSInFolder) {
      folder.addFile(ssFile);
      DriveApp.getRootFolder().removeFile(ssFile);
      debugLog('📁 スプレッドシートファイルをフォルダに移動完了');
    }

  } catch (error) {
    warnLog('⚠️ ファイル移動処理で警告:', error.message);
    // ファイル移動の失敗は致命的エラーではないため、処理を継続
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
    const activeUserEmail = Session.getActiveUser().getEmail();

    // AdminPanelのconfig構造を内部形式に変換（createCustomForm の処理を統合）
    const convertedConfig = {
      mainQuestion: {
        title: config.mainQuestion || '今日の学習について、あなたの考えや感想を聞かせてください',
        type: config.responseType || config.questionType || 'text', // responseTypeを優先して使用
        choices: config.questionChoices || config.choices || [], // questionChoicesを優先して使用
        includeOthers: config.includeOthers || false
      },
      enableClass: config.enableClass || false,
      classQuestion: {
        choices: config.classChoices || ['クラス1', 'クラス2', 'クラス3', 'クラス4']
      }
    };

    debugLog('createCustomFormUI - converted config:', JSON.stringify(convertedConfig));

    const overrides = {
      titlePrefix: config.formTitle || 'カスタムフォーム',
      customConfig: convertedConfig
    };

    const result = createUnifiedForm('custom', activeUserEmail, requestUserId, overrides);

    // 新規追加: カスタムフォーム作成後のシートアクティベーション
    if (result.sheetName) {
      try {
        debugLog('🎯 カスタムフォーム作成後のシートアクティベーション開始:', result.sheetName);
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
      debugLog('📁 ユーザーフォルダ作成とファイル移動開始...');
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
              folder.addFile(formFile);
              DriveApp.getRootFolder().removeFile(formFile);
              moveResults.form = true;
              infoLog('✅ カスタムフォームファイル移動完了');
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
              folder.addFile(ssFile);
              DriveApp.getRootFolder().removeFile(ssFile);
              moveResults.spreadsheet = true;
              infoLog('✅ カスタムスプレッドシートファイル移動完了');
            }
          }
        } catch (ssMoveError) {
          warnLog('カスタムスプレッドシートファイル移動エラー:', ssMoveError.message);
        }

        debugLog('📁 カスタムセットアップファイル移動結果:', moveResults);
      }
    } catch (folderError) {
      warnLog('カスタムセットアップフォルダ処理エラー:', folderError.message);
    }

    // 既存ユーザーの情報を更新（スプレッドシート情報を追加）
    const existingUser = findUserById(requestUserId);
    if (existingUser) {
      debugLog('createCustomFormUI - updating user data for:', requestUserId);

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

      debugLog('createCustomFormUI - update data:', JSON.stringify(updateData));

      updateUser(requestUserId, updateData);

      // カスタムフォーム作成後の包括的キャッシュ同期（Quick Startと同様）
      debugLog('🗑️ カスタムフォーム作成後の包括的キャッシュ同期中...');
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
    const activeUserEmail = Session.getActiveUser().getEmail();

    // createQuickStartForm の処理を統合（直接 createUnifiedForm を呼び出し）
    const result = createUnifiedForm('quickstart', activeUserEmail, requestUserId);

    // QuickStart作成後のシートアクティベーション
    if (result.sheetName) {
      try {
        debugLog('🎯 QuickStart作成後のシートアクティベーション開始:', result.sheetName);
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

function deleteCurrentUserAccount(requestUserId) {
  try {
    if (!requestUserId) {
      throw new Error('認証エラー: ユーザーIDが指定されていません');
    }
    verifyUserAccess(requestUserId);
    const result = deleteUserAccount(requestUserId);

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
 * 現在ログイン中のユーザーのメールアドレスを取得する簡易関数
 * @returns {string} ユーザーのメールアドレス
 */
function getCurrentUserEmail() {
  try {
    return Session.getActiveUser().getEmail() || '';
  } catch (error) {
    errorLog('getCurrentUserEmail error:', error);
    return '';
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
 * 統合ログインフロー処理 - セキュリティ強化とパフォーマンス最適化
 * 従来の複数API呼び出し（getCurrentUserStatus → getExistingBoard → registerNewUser）を1回に集約
 * @returns {object} ログイン結果とリダイレクトURL
 */
function processLoginFlow() {
  try {
    // アクティブユーザーのメールアドレスを取得
    var activeUserEmail = Session.getActiveUser().getEmail();
    if (!activeUserEmail) {
      return {
        status: 'error',
        message: 'ログインユーザーの情報を取得できませんでした。'
      };
    }

    debugLog('processLoginFlow: ログインフロー開始 -', activeUserEmail);

    // ログインフロー専用の短期キャッシュをチェック（30秒キャッシュ）
    var cacheKey = 'login_flow_' + activeUserEmail;
    var cached = null;
    try {
      var cachedString = CacheService.getScriptCache().get(cacheKey);
      if (cachedString) {
        cached = JSON.parse(cachedString);
        debugLog('processLoginFlow: キャッシュから高速返却');
        return cached;
      }
    } catch (e) {
      warnLog('processLoginFlow: キャッシュ読み込みエラー -', e.message);
      cached = null;
    }

    // ユーザー情報をメールアドレスで検索（キャッシュ活用）
    var userInfo = cacheManager.get(
      'email_' + activeUserEmail,
      function() { return findUserByEmail(activeUserEmail); },
      { ttl: 300, enableMemoization: true }
    );

    var result;

    if (userInfo) {
      // ユーザーが存在する場合（アクティブ状態に関係なく既存ユーザーとして処理）
      debugLog('processLoginFlow: 既存ユーザー検出 -', userInfo.userId, 'isActive:', userInfo.isActive, 'type:', typeof userInfo.isActive);

      // アクティブ状態の安全なチェック（undefinedの場合はtrueと仮定）
      var isUserActive = (userInfo.isActive === undefined ||
                         userInfo.isActive === true ||
                         userInfo.isActive === 'true' ||
                         String(userInfo.isActive).toLowerCase() === 'true');

      if (isUserActive) {
        // アクティブユーザー - 最終アクセス時刻のみ更新（設定は保護）
        try {
          updateUser(userInfo.userId, {
            lastAccessedAt: new Date().toISOString(),
            isActive: 'true'
          });
          debugLog('processLoginFlow: 既存ユーザーの最終アクセス時刻を更新しました（設定保護）');
        } catch (updateError) {
          warnLog('processLoginFlow: 最終アクセス時刻更新エラー:', updateError.message);
          // 更新エラーは無視して続行
        }

        var appUrls = generateUserUrls(userInfo.userId);
        result = {
          status: 'existing_user',
          userId: userInfo.userId,
          adminUrl: appUrls.adminUrl,
          viewUrl: appUrls.viewUrl,
          message: 'おかえりなさい！'
        };
        debugLog('processLoginFlow: 既存アクティブユーザー（設定保護）-', userInfo.userId);
      } else {
        // 明示的に非アクティブなユーザー
        var appUrls = generateUserUrls(userInfo.userId);
        result = {
          status: 'inactive_user',
          userId: userInfo.userId,
          adminUrl: appUrls.adminUrl,
          viewUrl: appUrls.viewUrl,
          message: 'アカウントが無効化されています。管理者にお問い合わせください。'
        };
        debugLog('processLoginFlow: 既存非アクティブユーザー -', userInfo.userId);
      }
    } else {
      // 新規ユーザー - 自動登録処理
      debugLog('processLoginFlow: 新規ユーザー登録開始 -', activeUserEmail);

      // registerNewUser を呼び出して新規登録
      var registrationResult = registerNewUser(activeUserEmail);

      if (registrationResult && registrationResult.adminUrl) {
        result = {
          status: 'new_user',
          userId: registrationResult.userId,
          adminUrl: registrationResult.adminUrl,
          viewUrl: registrationResult.viewUrl,
          message: '新規ユーザー登録が完了しました'
        };
        debugLog('processLoginFlow: 新規ユーザー登録完了 -', registrationResult.userId);
      } else {
        throw new Error('新規ユーザー登録に失敗しました: ' + (registrationResult ? registrationResult.message : '不明なエラー'));
      }
    }

    // 結果を短期キャッシュに保存（30秒）
    try {
      CacheService.getScriptCache().put(cacheKey, JSON.stringify(result), 30);
    } catch (e) {
      warnLog('processLoginFlow: キャッシュ保存エラー -', e.message);
      // キャッシュ保存に失敗しても処理は続行
    }

    debugLog('processLoginFlow: 処理完了 -', result.status, result.userId);
    return result;

  } catch (error) {
    errorLog('processLoginFlow: エラー発生 -', error.message);
    errorLog('processLoginFlow: エラースタック -', error.stack);

    // エラータイプに応じた詳細なメッセージを提供
    var errorMessage = 'ログイン処理中にエラーが発生しました';
    if (error.message.includes('not a function')) {
      errorMessage = 'システム関数の呼び出しエラーが発生しました。しばらく待ってから再試行してください。';
    } else if (error.message.includes('permission') || error.message.includes('権限')) {
      errorMessage = 'アクセス権限に問題があります。管理者にお問い合わせください。';
    } else if (error.message.includes('network') || error.message.includes('timeout')) {
      errorMessage = 'ネットワークエラーが発生しました。接続を確認して再試行してください。';
    } else {
      errorMessage = 'ログイン処理中に予期しないエラーが発生しました: ' + error.message;
    }

    return {
      status: 'error',
      message: errorMessage,
      errorType: error.name,
      timestamp: new Date().toISOString()
    };
  }
}

/**
 * Returns the current login status without auto-registering new users.
 * @returns {Object} Login status result
 */
function getLoginStatus() {
  try {
    var activeUserEmail = Session.getActiveUser().getEmail();
    if (!activeUserEmail) {
      return { status: 'error', message: 'ログインユーザーの情報を取得できませんでした。' };
    }

    var cacheKey = 'login_status_' + activeUserEmail;
    try {
      var cached = CacheService.getScriptCache().get(cacheKey);
      if (cached) return JSON.parse(cached);
    } catch (e) {
      warnLog('getLoginStatus: キャッシュ読み込みエラー -', e.message);
    }

    var userInfo = cacheManager.get(
      'email_' + activeUserEmail,
      function() { return findUserByEmail(activeUserEmail); },
      { ttl: 300, enableMemoization: true }
    );

    var result;
    if (userInfo && (userInfo.isActive === true || String(userInfo.isActive).toLowerCase() === 'true')) {
      var urls = generateUserUrls(userInfo.userId);
      result = {
        status: 'existing_user',
        userId: userInfo.userId,
        adminUrl: urls.adminUrl,
        viewUrl: urls.viewUrl,
        message: 'ログインが完了しました'
      };
    } else if (userInfo) {
      var urls = generateUserUrls(userInfo.userId);
      result = {
        status: 'setup_required',
        userId: userInfo.userId,
        adminUrl: urls.adminUrl,
        viewUrl: urls.viewUrl,
        message: 'セットアップを完了してください'
      };
    } else {
      result = { status: 'unregistered', userEmail: activeUserEmail };
    }

    try {
      CacheService.getScriptCache().put(cacheKey, JSON.stringify(result), 30);
    } catch (e) {
      warnLog('getLoginStatus: キャッシュ保存エラー -', e.message);
    }

    return result;
  } catch (error) {
    errorLog('getLoginStatus error:', error);
    return { status: 'error', message: error.message };
  }
}

/**
 * Confirms registration for the active user on demand.
 * @returns {Object} Registration result
 */
function confirmUserRegistration() {
  try {
    var activeUserEmail = Session.getActiveUser().getEmail();
    if (!activeUserEmail) {
      return { status: 'error', message: 'ユーザー情報を取得できませんでした。' };
    }
    return registerNewUser(activeUserEmail);
  } catch (error) {
    errorLog('confirmUserRegistration error:', error);
    return { status: 'error', message: error.message };
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
 * @returns {Object} 統合された初期データ
 */
function getInitialData(requestUserId, targetSheetName) {
  debugLog('🚀 getInitialData: 統合初期化開始', { requestUserId, targetSheetName });

  try {
    var startTime = new Date().getTime();

    // === ステップ1: ユーザー認証とユーザー情報取得（キャッシュ活用） ===
    var activeUserEmail = Session.getActiveUser().getEmail();
    var currentUserId = requestUserId;

    // UserID の解決
    if (!currentUserId) {
      currentUserId = getUserId();
    }

    // Phase3 Optimization: Use execution-level cache to avoid duplicate database queries
    clearExecutionUserInfoCache(); // Clear any stale cache

    // ユーザー認証
    verifyUserAccess(currentUserId);
    var userInfo = getOrFetchUserInfo(currentUserId, 'userId', {
      useExecutionCache: true,
      ttl: 300
    }); // Use cached version
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    // === ステップ1.5: データ整合性の自動チェックと修正 ===
    try {
      debugLog('🔍 データ整合性の自動チェック開始...');
      var consistencyResult = fixUserDataConsistency(currentUserId);
      if (consistencyResult.updated) {
        infoLog('✅ データ整合性が自動修正されました');
        // 修正後は最新データを再取得
        clearExecutionUserInfoCache();
        userInfo = getOrFetchUserInfo(currentUserId, 'userId', {
          useExecutionCache: true,
          ttl: 300
        });
      }
    } catch (consistencyError) {
      warnLog('⚠️ データ整合性チェック中にエラー:', consistencyError.message);
      // エラーが発生しても初期化処理は続行
    }

    // === ステップ2: 設定データの取得と自動修復 ===
    var configJson = JSON.parse(userInfo.configJson || '{}');

    // --- 統一された自動修復システム ---
    const healingResult = performAutoHealing(userInfo, configJson, currentUserId);
    if (healingResult.updated) {
      configJson = healingResult.configJson;
      userInfo.configJson = JSON.stringify(configJson);
    }

    // === ステップ3: シート一覧とアプリURL生成 ===
    var sheets = getSheetsList(currentUserId);
    var appUrls = generateUserUrls(currentUserId);

    // === ステップ4: 回答数とリアクション数の取得 ===
    var answerCount = 0;
    var totalReactions = 0;
    try {
      if (configJson.publishedSpreadsheetId && configJson.publishedSheetName) {
        var responseData = getResponsesData(currentUserId, configJson.publishedSheetName);
        if (responseData.status === 'success') {
          answerCount = responseData.data.length;
          totalReactions = answerCount * 2; // 暫定値
        }
      }
    } catch (err) {
      warnLog('Answer count retrieval failed:', err.message);
    }

    // === ステップ5: セットアップステップの決定 ===
    var setupStep = 1;
    try {
      setupStep = getSetupStep(userInfo, configJson);
      debugLog('setupStep決定完了', { userId: currentUserId, setupStep: setupStep });
    } catch (stepError) {
      warnLog('setupStep決定でエラー、デフォルト値(1)を使用:', stepError.message);
      setupStep = 1;
    }

    // 公開シート設定とヘッダー情報を取得
    var publishedSheetName = configJson.publishedSheetName || '';
    var sheetConfigKey = publishedSheetName ? 'sheet_' + publishedSheetName : '';
    var activeSheetConfig = sheetConfigKey && configJson[sheetConfigKey]
      ? configJson[sheetConfigKey]
      : {};

    var opinionHeader = activeSheetConfig.opinionHeader || '';
    var nameHeader = activeSheetConfig.nameHeader || '';
    var classHeader = activeSheetConfig.classHeader || '';

    // === ベース応答の構築 ===
    var response = {
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
    var includeSheetDetails = targetSheetName || configJson.publishedSheetName;

    // デバッグ: シート詳細取得パラメータの確認
    debugLog('🔍 getInitialData: シート詳細取得パラメータ確認:', {
      targetSheetName: targetSheetName,
      publishedSheetName: configJson.publishedSheetName,
      includeSheetDetails: includeSheetDetails,
      hasSpreadsheetId: !!userInfo.spreadsheetId,
      willIncludeSheetDetails: !!(includeSheetDetails && userInfo.spreadsheetId)
    });

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
        var sharedSheetsService = getSheetsServiceCached();

        // ExecutionContext を最適化版で作成（sheetsService と userInfo を渡して重複作成を回避）
        const context = createExecutionContext(currentUserId, {
          reuseService: sharedSheetsService,
          reuseUserInfo: userInfo
        });
        var sheetDetails = getSheetDetailsFromContext(context, userInfo.spreadsheetId, includeSheetDetails);
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
    var endTime = new Date().getTime();
    response._meta.executionTime = endTime - startTime;

    debugLog('🎯 getInitialData: 統合初期化完了', {
      executionTime: response._meta.executionTime + 'ms',
      userId: currentUserId,
      setupStep: setupStep,
      hasSheetDetails: !!response.sheetDetails
    });

    return response;

  } catch (error) {
    errorLog('❌ getInitialData error:', error);
    return {
      status: 'error',
      message: error.message,
      _meta: {
        apiVersion: 'integrated_v1',
        error: error.message
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
    debugLog('🔧 手動データ整合性修正実行:', requestUserId);

    var result = fixUserDataConsistency(requestUserId);

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
    const adminEmail = Session.getActiveUser().getEmail();

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
    debugLog('🧪 testForceLogoutRedirect called');
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
