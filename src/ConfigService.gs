/**
 * @fileoverview ConfigService - 統一設定管理サービス (遅延初期化対応)
 *
 * 🎯 責任範囲:
 * - configJSON の CRUD操作
 * - 設定検証・サニタイズ
 * - 動的設定生成（URL等）
 * - 設定マイグレーション
 *
 * 🔄 GAS Best Practices準拠:
 * - 遅延初期化パターン (各公開関数先頭でinit)
 * - ファイル読み込み順序非依存設計
 * - グローバル副作用排除
 */

/* global ServiceFactory, URL, DatabaseOperations, validateUrl, createErrorResponse, validateSpreadsheetId */

// ===========================================
// 🔧 Zero-Dependency ConfigService (ServiceFactory版)
// ===========================================

/**
 * ConfigService - ゼロ依存アーキテクチャ
 * ServiceFactoryパターンによる依存関係除去
 * PROPS_KEYS, DB依存を完全排除
 */

/**
 * ServiceFactory統合初期化
 * 依存関係チェックなしの即座初期化
 * @returns {boolean} 初期化成功可否
 */
function initConfigServiceZero() {
  return ServiceFactory.getUtils().initService('ConfigService');
}



/**
 * デフォルト設定取得
 * @param {string} userId - ユーザーID
 * @returns {Object} デフォルト設定
 */
function getDefaultConfig(userId) {
  // 🚀 Zero-dependency: ServiceFactory不要（静的デフォルト値のため）
  return {
    userId,
    setupStatus: 'pending',
    appPublished: false,
    displaySettings: {
      showNames: false,
      showReactions: false
    },
    userPermissions: {
      isOwner: false,
      isSystemAdmin: false,
      accessLevel: 'viewer',
      canEdit: false,
      canView: true,
      canReact: true
    },
    setupStep: 1,
    completionScore: 0
  };
}

/**
 * 設定JSONパース・修復
 * @param {string} configJson - 設定JSON文字列
 * @param {string} userId - ユーザーID
 * @returns {Object} パース済み設定
 */
function parseAndRepairConfig(configJson, userId) {
  try {
    const config = JSON.parse(configJson || '{}');

    // 基本修復
    const repairedConfig = repairNestedConfig(config, userId);

    // 必須フィールド確保
    return ensureRequiredFields(repairedConfig, userId);
  } catch (parseError) {
    console.warn('parseAndRepairConfig: JSON解析失敗 - デフォルト設定を使用', {
      userId,
      error: parseError.message
    });
    return getDefaultConfig(userId);
  }
}

/**
 * ネストされた設定構造の修復
 * @param {Object} config - 設定オブジェクト
 * @param {string} userId - ユーザーID
 * @returns {Object} 修復済み設定
 */
function repairNestedConfig(config, userId) {
  const repaired = { ...config };

  // displaySettings修復
  if (!repaired.displaySettings || typeof repaired.displaySettings !== 'object') {
    repaired.displaySettings = {
      showNames: false,
      showReactions: false
    };
  }

  // columnMapping修復
  if (!repaired.columnMapping || typeof repaired.columnMapping !== 'object') {
    repaired.columnMapping = { mapping: {} };
  }

  // userPermissions修復
  if (!repaired.userPermissions) {
    repaired.userPermissions = generateUserPermissions(userId);
  }

  return repaired;
}

/**
 * 必須フィールドの確保
 * @param {Object} config - 設定オブジェクト
 * @param {string} userId - ユーザーID
 * @returns {Object} 必須フィールド確保済み設定
 */
function ensureRequiredFields(config, userId) {
  return {
    userId,
    setupStatus: config.setupStatus || 'pending',
    appPublished: Boolean(config.appPublished),
    spreadsheetId: config.spreadsheetId || '',
    sheetName: config.sheetName || '',
    formUrl: config.formUrl || '',
    displaySettings: config.displaySettings || {
      showNames: false,
      showReactions: false
    },
    columnMapping: config.columnMapping || { mapping: {} },
    userPermissions: config.userPermissions || generateUserPermissions(userId),
    setupStep: config.setupStep || determineSetupStep(null, JSON.stringify(config)),
    completionScore: calculateCompletionScore(config),
    lastModified: new Date().toISOString()
  };
}

/**
 * 動的URL付きの設定拡張
 * @param {Object} baseConfig - 基本設定
 * @param {string} userId - ユーザーID
 * @returns {Object} URL拡張済み設定
 */
function enhanceConfigWithDynamicUrls(baseConfig, userId) {
  const enhanced = { ...baseConfig };

  try {
    const webAppUrl = ServiceFactory.getUtils().getWebAppUrl();

    enhanced.dynamicUrls = {
      webAppUrl,
      adminPanelUrl: `${webAppUrl}?mode=admin&userId=${userId}`,
      viewBoardUrl: `${webAppUrl}?mode=view&userId=${userId}`,
      setupUrl: `${webAppUrl}?mode=setup&userId=${userId}`
    };

    // システムメタデータ追加
    enhanced.systemMeta = {
      lastGenerated: new Date().toISOString(),
      version: '3.0.0-flat',
      architecture: 'gas-flat-functions'
    };

  } catch (error) {
    console.warn('enhanceConfigWithDynamicUrls: URL生成エラー', error.message);
    enhanced.dynamicUrls = {};
    enhanced.systemMeta = {};
  }

  return enhanced;
}

/**
 * ユーザー権限生成
 * @param {string} userId - ユーザーID
 * @returns {Object} 権限オブジェクト
 */
function generateUserPermissions(_userId) {
  try {
    const session = ServiceFactory.getSession();
    const currentEmail = session.email;
    if (!currentEmail) {
      return {
        isOwner: false,
        isSystemAdmin: false,
        accessLevel: 'viewer',
        canEdit: false,
        canView: true,
        canReact: true
      };
    }

    const isSystemAdmin = checkIfSystemAdmin(currentEmail);

    return {
      isOwner: true, // 現在のユーザーは自分の設定のオーナー
      isSystemAdmin,
      accessLevel: isSystemAdmin ? 'admin' : 'owner',
      canEdit: true,
      canView: true,
      canReact: true,
      canDelete: isSystemAdmin,
      canManageUsers: isSystemAdmin
    };

  } catch (error) {
    console.error('generateUserPermissions: エラー', error.message);
    return {
      isOwner: false,
      isSystemAdmin: false,
      accessLevel: 'viewer',
      canEdit: false,
      canView: true,
      canReact: true
    };
  }
}

// ===========================================
// 💾 設定保存・更新
// ===========================================

/**
 * 設定保存（統合版）
 * @param {string} userId - ユーザーID
 * @param {Object} config - 保存する設定
 * @returns {Object} 保存結果
 */
function saveUserConfig(userId, config) {
  try {
    // 検証・サニタイズ
    const validatedConfig = validateAndSanitizeConfig(config, userId);

    if (!validatedConfig.success) {
      return {
        success: false,
        message: validatedConfig.message,
        errors: validatedConfig.errors
      };
    }

    const sanitizedConfig = validatedConfig.data;
    sanitizedConfig.lastModified = new Date().toISOString();

    // データベース保存
    const updateResult = DatabaseOperations.updateUserConfig(userId, JSON.stringify(sanitizedConfig));

    if (!updateResult) {
      throw new Error('データベース更新に失敗しました');
    }

    // キャッシュクリア
    clearConfigCache(userId);

    console.info('saveUserConfig: 設定保存完了', {
      userId,
      setupStatus: sanitizedConfig.setupStatus,
      completionScore: sanitizedConfig.completionScore
    });

    return {
      success: true,
      message: '設定を保存しました',
      data: sanitizedConfig
    };

  } catch (error) {
    console.error('saveUserConfig: エラー', {
      userId,
      error: error.message
    });
    return {
      success: false,
      message: '設定の保存に失敗しました',
      error: error.message
    };
  }
}


// ===========================================
// ✅ 設定検証・サニタイズ
// ===========================================

/**
 * 設定検証・サニタイズ（統合版）
 * @param {Object} config - 設定オブジェクト
 * @param {string} userId - ユーザーID
 * @returns {Object} 検証結果
 */
function validateAndSanitizeConfig(config, userId) {
  const errors = [];

  try {
    // 基本構造検証
    if (!config || typeof config !== 'object') {
      errors.push('設定は有効なオブジェクトである必要があります');
      return createErrorResponse('無効な設定形式', { errors });
    }

    const sanitized = { ...config };

    // userId検証
    if (!validateConfigUserId(userId)) {
      errors.push('無効なユーザーID形式');
    }
    sanitized.userId = userId;

    // spreadsheetId検証
    if (sanitized.spreadsheetId && !validateSpreadsheetId(sanitized.spreadsheetId)) {
      errors.push('無効なスプレッドシートID形式');
      sanitized.spreadsheetId = '';
    }

    // formUrl検証
    if (sanitized.formUrl && !validateUrl(sanitized.formUrl)?.isValid) {
      errors.push('無効なフォームURL形式');
      sanitized.formUrl = '';
    }

    // displaySettings サニタイズ
    if (sanitized.displaySettings) {
      sanitized.displaySettings = sanitizeDisplaySettings(sanitized.displaySettings);
    }

    // columnMapping サニタイズ
    if (sanitized.columnMapping) {
      sanitized.columnMapping = sanitizeColumnMapping(sanitized.columnMapping);
    }

    // エラーがある場合は失敗を返す
    if (errors.length > 0) {
      return {
        success: false,
        message: '設定検証エラーがあります',
        errors,
        data: sanitized
      };
    }

    return {
      success: true,
      message: '設定検証完了',
      data: sanitized,
      errors: []
    };

  } catch (error) {
    console.error('validateAndSanitizeConfig: エラー', error.message);
    return {
      success: false,
      message: '設定検証中にエラーが発生しました',
      errors: [error.message]
    };
  }
}

// ===========================================
// 🔧 ヘルパー関数
// ===========================================

/**
 * 表示設定サニタイズ
 * @param {Object} displaySettings - 表示設定
 * @returns {Object} サニタイズ済み表示設定
 */
function sanitizeDisplaySettings(displaySettings) {
  return {
    showNames: Boolean(displaySettings.showNames),
    showReactions: Boolean(displaySettings.showReactions),
    theme: String(displaySettings.theme || 'default').substring(0, 50),
    pageSize: Math.min(Math.max(Number(displaySettings.pageSize) || 20, 1), 100)
  };
}

/**
 * 列マッピングサニタイズ
 * @param {Object} columnMapping - 列マッピング
 * @returns {Object} サニタイズ済み列マッピング
 */
function sanitizeColumnMapping(columnMapping) {
  const sanitized = { mapping: {} };

  if (columnMapping.mapping && typeof columnMapping.mapping === 'object') {
    const validFields = ['answer', 'reason', 'class', 'name', 'timestamp', 'email'];

    validFields.forEach(field => {
      const index = columnMapping.mapping[field];
      if (typeof index === 'number' && index >= 0 && index < 100) {
        sanitized.mapping[field] = index;
      }
    });
  }

  return sanitized;
}

/**
 * ユーザーID検証
 * @param {string} userId - ユーザーID
 * @returns {boolean} 有効性
 */
function validateConfigUserId(userId) {
  return typeof userId === 'string' && userId.length > 0 && userId.length <= 100;
}

/**
 * スプレッドシートID検証
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {boolean} 有効性
 */

/**
 * フォームURL検証
 * @param {string} formUrl - フォームURL
 * @returns {boolean} 有効性
 */

// ===========================================
// 📊 システム状態・診断
// ===========================================

/**
 * セットアップステップ判定
 * @param {Object} userInfo - ユーザー情報
 * @param {string} configJson - 設定JSON
 * @returns {number} セットアップステップ (1-3)
 */
function determineSetupStep(userInfo, configJson) {
  try {
    const config = JSON.parse(configJson || '{}');

    if (!config.spreadsheetId) {
      return 1; // データソース未設定
    }

    if (!config.formUrl || config.setupStatus !== 'completed') {
      return 2; // 設定未完了
    }

    if (config.appPublished) {
      return 3; // 完了・公開済み
    }

    return 2; // 設定完了だが未公開
  } catch (error) {
    console.error('determineSetupStep: エラー', error.message);
    return 1;
  }
}

/**
 * システムセットアップ状態確認
 * @returns {boolean} セットアップ済みかどうか
 */
function isSystemSetup() {
  try {
    const session = ServiceFactory.getSession();
    const currentEmail = session.email;
    if (!currentEmail) return false;

    // 🔧 ServiceFactory経由で直接データベースから取得
    const db = ServiceFactory.getDB();
    if (!db) return false;

    const user = db.findUserByEmail(currentEmail);
    return !!(user && user.configJson);
  } catch (error) {
    console.error('isSystemSetup: エラー', error.message);
    return false;
  }
}

/**
 * 設定完了スコア計算
 * @param {Object} config - 設定オブジェクト
 * @returns {number} 完了スコア (0-100)
 */
function calculateCompletionScore(config) {
  if (!config) return 0;

  let score = 0;
  const maxScore = 100;

  // 基本設定 (50点)
  if (config.spreadsheetId) score += 25;
  if (config.sheetName) score += 25;

  // フォーム設定 (30点)
  if (config.formUrl) score += 30;

  // 表示設定 (10点)
  if (config.displaySettings) score += 10;

  // 列マッピング (10点)
  if (config.columnMapping && config.columnMapping.mapping) score += 10;

  return Math.min(score, maxScore);
}

/**
 * 設定キャッシュクリア
 * @param {string} userId - ユーザーID
 */
function clearConfigCache(userId) {
  try {
    const cacheKey = `config_${userId}`;
    ServiceFactory.getCache().remove(cacheKey);
    console.info('clearConfigCache: キャッシュクリア完了', { userId });
  } catch (error) {
    console.warn('clearConfigCache: キャッシュクリアエラー', error.message);
  }
}

/**
 * 全設定キャッシュクリア
 */
function clearAllConfigCache() {
  try {
    // 個別キャッシュクリアは困難なため、プリフィックスパターンでクリア
    console.info('clearAllConfigCache: 全設定キャッシュクリア実行');
    // Note: GASにはワイルドカードクリア機能がないため、必要に応じて個別にクリア
  } catch (error) {
    console.warn('clearAllConfigCache: エラー', error.message);
  }
}

/**
 * システム診断
 * @returns {Object} 診断結果
 */

/**
 * コアシステムプロパティ確認 - 3つの必須項目をすべてチェック
 * @returns {boolean} 3つすべて存在すれば true
 */
function hasCoreSystemProps() {
  try {
    const props = ServiceFactory.getProperties();

    // 3つの必須項目をすべてチェック（依存関係なしで直接指定）
    const adminEmail = props.getProperty('ADMIN_EMAIL');
    const dbId = props.getProperty('DATABASE_SPREADSHEET_ID');
    const creds = props.getProperty('SERVICE_ACCOUNT_CREDS');

    if (!adminEmail || !dbId || !creds) {
      console.warn('hasCoreSystemProps: 必須項目不足', {
        hasAdmin: !!adminEmail,
        hasDb: !!dbId,
        hasCreds: !!creds
      });
      return false;
    }

    // SERVICE_ACCOUNT_CREDSのJSON検証
    try {
      const parsed = JSON.parse(creds);
      if (!parsed || typeof parsed !== 'object' || !parsed.client_email) {
        console.warn('hasCoreSystemProps: SERVICE_ACCOUNT_CREDS JSON不正');
        return false;
      }
    } catch (jsonError) {
      console.warn('hasCoreSystemProps: SERVICE_ACCOUNT_CREDS JSON解析失敗', jsonError.message);
      return false;
    }

    return true;
  } catch (error) {
    console.error('hasCoreSystemProps: エラー', error.message);
    return false;
  }
}

// ===========================================
// 📱 アプリケーション管理
// ===========================================





// ===========================================
// 🔧 ヘルパー関数（依存関数）
// ===========================================


/**
 * システム管理者確認
 * @param {string} email - メールアドレス
 * @returns {boolean} システム管理者かどうか
 */
function checkIfSystemAdmin(email) {
  try {
    const adminEmail = ServiceFactory.getProperties().getProperty('ADMIN_EMAIL');
    return adminEmail && adminEmail === email;
  } catch (error) {
    console.error('checkIfSystemAdmin: エラー', error.message);
    return false;
  }
}
