/**
 * @fileoverview ConfigService - 統一設定管理サービス (遅延初期化対応)
 *
 * 責任範囲:
 * - configJSON の CRUD操作
 * - 設定検証・サニタイズ
 * - 動的設定生成（URL等）
 * - 設定マイグレーション
 *
 * GAS Best Practices準拠:
 * - 遅延初期化パターン (各公開関数先頭でinit)
 * - ファイル読み込み順序非依存設計
 * - グローバル副作用排除
 */

/* global getCurrentEmail, findUserById, updateUser, validateEmail, CACHE_DURATION, TIMEOUT_MS, SYSTEM_LIMITS, validateConfig, URL, validateUrl, createErrorResponse, validateSpreadsheetId, findUserByEmail, findUserBySpreadsheetId, openSpreadsheet, UserService, isAdministrator, SLEEP_MS, getSheetInfo */




/**
 * ConfigService - ゼロ依存アーキテクチャ
 * GAS-Nativeパターンによる直接APIアクセス
 * PROPS_KEYS, DB依存を完全排除
 */


/**
 * FormApp権限の安全チェック - GAS 2025ベストプラクティス準拠（実行時テスト強化版）
 * V8ランタイム対応の軽量権限テスト（実際のAPI呼び出しで検証）
 * @returns {Object} 権限チェック結果
 */
function validateFormAppAccess() {
  try {
    if (typeof FormApp === 'undefined') {
      return {
        hasAccess: false,
        reason: 'FORMAPP_NOT_AVAILABLE',
        message: 'FormAppサービスが利用できません'
      };
    }

    try {
      const hasUi = typeof FormApp.getUi === 'function';

      if (!hasUi) {
        return {
          hasAccess: false,
          reason: 'FORMAPP_UI_NOT_AVAILABLE',
          message: 'FormApp基本機能が利用できません'
        };
      }

      if (typeof FormApp.openByUrl !== 'function') {
        return {
          hasAccess: false,
          reason: 'OPENBYURL_NOT_AVAILABLE',
          message: 'FormApp.openByUrl機能が利用できません'
        };
      }

      return {
        hasAccess: true,
        reason: 'ACCESS_GRANTED',
        message: 'FormAppへのアクセス権限が確認されました（実行時テスト済み）'
      };

    } catch (runtimeError) {
      return {
        hasAccess: false,
        reason: 'RUNTIME_PERMISSION_ERROR',
        message: `FormApp実行時権限エラー: ${runtimeError.message || '詳細不明'}`,
        error: runtimeError.message
      };
    }

  } catch (error) {
    return {
      hasAccess: false,
      reason: 'PERMISSION_ERROR',
      message: error && error.message ? `FormApp権限エラー: ${error.message}` : 'FormApp権限エラー: 詳細不明',
      error: error.message
    };
  }
}


/**
 * デフォルト設定取得
 * @param {string} userId - ユーザーID
 * @returns {Object} デフォルト設定
 */
function getDefaultConfig(userId) {
  return {
    userId,
    setupStatus: 'pending',
    isPublished: false,
    displaySettings: {
      showNames: false,
      showReactions: false,
      theme: 'default',
      pageSize: 20
    },
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

    const repairedConfig = repairNestedConfig(config, userId);

    return ensureRequiredFields(repairedConfig, userId);
  } catch (parseError) {
    console.warn('parseAndRepairConfig: JSON解析失敗 - デフォルト設定を使用', {
      operation: 'parseAndRepairConfig',
      userId: userId && typeof userId === 'string' ? `${userId.substring(0, 8)}***` : 'N/A',
      configLength: configJson?.length || 0,
      error: parseError.message,
      stack: parseError.stack
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

  if (!repaired.displaySettings || typeof repaired.displaySettings !== 'object') {
    repaired.displaySettings = {
      showNames: false,
      showReactions: false,
      theme: 'default',
      pageSize: 20
    };
  }

  if (!repaired.columnMapping || typeof repaired.columnMapping !== 'object') {
    repaired.columnMapping = {};
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
    isPublished: Boolean(config.isPublished),
    spreadsheetId: config.spreadsheetId || '',
    sheetName: config.sheetName || '',
    formUrl: config.formUrl || '',
    displaySettings: config.displaySettings || { showNames: false, showReactions: false, theme: 'default', pageSize: 20 },
    columnMapping: config.columnMapping,
    completionScore: calculateCompletionScore(config)
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
    const webAppUrl = getWebAppUrl(); // eslint-disable-line no-undef

    enhanced.dynamicUrls = {
      webAppUrl: webAppUrl || '',
      adminPanelUrl: webAppUrl ? `${webAppUrl}?mode=admin&userId=${userId}` : '',
      viewBoardUrl: webAppUrl ? `${webAppUrl}?mode=view&userId=${userId}` : '',
      setupUrl: webAppUrl ? `${webAppUrl}?mode=setup&userId=${userId}` : '',
      manualUrl: webAppUrl ? `${webAppUrl}?mode=manual` : ''
    };

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
 * 公開専用設定検証（厳格版）
 * ✅ CRITICAL FIX: 公開時の必須フィールドを厳格に検証
 * @param {Object} config - 設定オブジェクト
 * @param {string} userId - ユーザーID
 * @returns {Object} 検証結果
 */
function validatePublishConfig(config, userId) {
  try {
    const baseValidation = validateAndSanitizeConfig(config, userId);
    if (!baseValidation.success) {
      return baseValidation;
    }

    const errors = [];

    if (!config.spreadsheetId || typeof config.spreadsheetId !== 'string' || !config.spreadsheetId.trim()) {
      errors.push('公開にはスプレッドシートIDが必要です');
    }

    if (!config.sheetName || typeof config.sheetName !== 'string' || !config.sheetName.trim()) {
      errors.push('公開にはシート名が必要です');
    }

    if (!config.columnMapping || typeof config.columnMapping !== 'object') {
      errors.push('公開には列マッピングが必要です');
    } else if (Object.keys(config.columnMapping).length === 0) {
      errors.push('公開には列マッピングが必要です（空のマッピングは無効）');
    } else {
      const answerColumn = config.columnMapping.answer;
      if (answerColumn === undefined || answerColumn === null || (typeof answerColumn === 'number' && answerColumn < 0)) {
        errors.push('公開には回答列（answer）の設定が必要です');
      }
    }

    if (errors.length > 0) {
      return {
        success: false,
        message: '公開に必要な設定が不足しています',
        errors,
        data: baseValidation.data
      };
    }

    return baseValidation;

  } catch (error) {
    console.error('validatePublishConfig: エラー', error.message);
    return {
      success: false,
      message: '公開設定検証中にエラーが発生しました',
      errors: [error.message]
    };
  }
}

/**
 * 設定検証・サニタイズ（統合版）
 * @param {Object} config - 設定オブジェクト
 * @param {string} userId - ユーザーID
 * @returns {Object} 検証結果
 */
function validateAndSanitizeConfig(config, userId) {
  try {
    if (config.columnMapping) {
    }

    const validationResult = validateConfig(config);
    if (!validationResult.isValid) {
      return {
        success: false,
        message: '設定検証エラーがあります',
        errors: validationResult.errors,
        data: validationResult.sanitized || config
      };
    }

    const errors = [];
    if (!validateConfigUserId(userId)) {
      errors.push('無効なユーザーID形式');
    }

    const sanitized = { ...validationResult.sanitized };
    sanitized.userId = userId;

    if (sanitized.displaySettings) {
      sanitized.displaySettings = sanitizeDisplaySettings(sanitized.displaySettings);
    }
    if (sanitized.columnMapping) {
      sanitized.columnMapping = sanitizeMapping(sanitized.columnMapping);
    }

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




/**
 * 表示設定サニタイズ
 * @param {Object} displaySettings - 表示設定
 * @returns {Object} サニタイズ済み表示設定
 */
function sanitizeDisplaySettings(displaySettings) {
  return {
    showNames: Boolean(displaySettings.showNames),
    showReactions: Boolean(displaySettings.showReactions),
    theme: String(displaySettings.theme || 'default').substring(0, SYSTEM_LIMITS.PREVIEW_LENGTH),
    pageSize: Math.min(Math.max(Number(displaySettings.pageSize) || SYSTEM_LIMITS.DEFAULT_PAGE_SIZE, 1), SYSTEM_LIMITS.MAX_PAGE_SIZE)
  };
}

/**
 * 列マッピングサニタイズ
 * @param {Object} columnMapping - 列マッピング
 * @returns {Object} サニタイズ済み列マッピング
 */
function sanitizeMapping(columnMapping) {
  const sanitized = {};
  const validFields = ['answer', 'reason', 'class', 'name', 'timestamp', 'email'];

  validFields.forEach(field => {
    const index = columnMapping[field];
    if (typeof index === 'number' && index >= 0 && index < 100) {
      sanitized[field] = index;
    }
  });

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





/**
 * システムセットアップ状態確認
 * @returns {boolean} セットアップ済みかどうか
 */
function isSystemSetup() {
  try {
    const userConfigResult = getBatchedUserConfig(); // eslint-disable-line no-undef
    if (!userConfigResult.success || !userConfigResult.user) {
      return false;
    }

    return !!(userConfigResult.user && userConfigResult.user.configJson);
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

  if (config.spreadsheetId) score += 25;
  if (config.sheetName) score += 25;

  if (config.formUrl) score += 30;

  if (config.displaySettings) score += 10;

  if (config.columnMapping && Object.keys(config.columnMapping).length > 0) score += 10;

  return Math.min(score, maxScore);
}

/**
 * 設定キャッシュクリア
 * @param {string} userId - ユーザーID
 */
function clearConfigCache(userId) {
  try {
    const cache = CacheService.getScriptCache();

    const keysToRemove = [
      `config_${userId}`,           // 設定キャッシュ
      `user_${userId}`,             // ユーザーキャッシュ
      `board_data_${userId}`,       // ボードデータキャッシュ
      `admin_panel_${userId}`,      // 管理パネルデータ
      `question_text_${userId}`     // 質問テキストキャッシュ
    ];

    if (keysToRemove.length > 0) {
      cache.removeAll(keysToRemove);
    }
  } catch (error) {
    console.warn('clearConfigCache: キャッシュクリアエラー', error.message);
  }
}

/**
 * 全設定キャッシュクリア（特定ユーザー群用）
 */
function clearAllConfigCache(userIds = []) {
  try {
    if (!Array.isArray(userIds) || userIds.length === 0) {
      console.warn('clearAllConfigCache: ユーザーID配列が空または無効');
      return;
    }

    const allKeysToRemove = [];
    userIds.forEach(userId => {
      if (typeof userId === 'string' && userId.trim()) {
        allKeysToRemove.push(
          `config_${userId}`,
          `user_${userId}`,
          `board_data_${userId}`,
          `admin_panel_${userId}`,
          `question_text_${userId}`
        );
      }
    });

    if (allKeysToRemove.length > 0) {
      CacheService.getScriptCache().removeAll(allKeysToRemove);
    }
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
    const props = PropertiesService.getScriptProperties();

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











/**
 * 動的questionText取得（configJson最適化対応 + パフォーマンス最適化済み）
 * headers配列とcolumnMappingから実際の問題文を動的取得
 * preloadedHeadersが提供された場合は重複スプレッドシートアクセスを回避
 * @param {Object} config - ユーザー設定オブジェクト
 * @param {Object} context - アクセスコンテキスト（target user info for cross-user access）
 * @param {Array} preloadedHeaders - 事前取得されたヘッダー配列（パフォーマンス最適化用）
 * @returns {string} 問題文テキスト
 */
function getQuestionText(config, context = {}, preloadedHeaders = null) {
  try {

    const answerIndex = config?.columnMapping?.answer;

    if (typeof answerIndex === 'number' && config?.headers?.[answerIndex]) {
      const questionText = config.headers[answerIndex];
      if (questionText && typeof questionText === 'string' && questionText.trim()) {
        return questionText.trim();
      }
    }

    if (typeof answerIndex === 'number' && preloadedHeaders && preloadedHeaders[answerIndex]) {
      const questionText = preloadedHeaders[answerIndex];
      if (questionText && typeof questionText === 'string' && questionText.trim()) {
        return questionText.trim();
      }
    }

    if (typeof answerIndex === 'number' && config?.spreadsheetId && config?.sheetName) {
      try {
        // ✅ CRITICAL: 同一ドメイン共有設定で対応（サービスアカウント不使用）
        const currentEmail = getCurrentEmail();
        try {
          const dataAccess = openSpreadsheet(config.spreadsheetId, { useServiceAccount: false });
          const { spreadsheet } = dataAccess;
          const sheet = spreadsheet.getSheetByName(config.sheetName);
          if (sheet) {
            const { lastCol, headers } = getSheetInfo(sheet);
            if (lastCol > 0 && headers && headers[answerIndex]) {
              const questionText = headers[answerIndex];
              if (questionText && typeof questionText === 'string' && questionText.trim()) {
                return questionText.trim();
              }
            }
          }
        } catch (dataAccessError) {
          console.warn('⚠️ getQuestionText: openSpreadsheet access failed:', dataAccessError.message);
        }
      } catch (dynamicError) {
        console.warn('⚠️ getQuestionText: Dynamic headers fetch failed:', dynamicError.message);
      }
    }

    if (config?.formTitle && typeof config.formTitle === 'string' && config.formTitle.trim()) {
      return config.formTitle.trim();
    }

    return 'Everyone\'s Answer Board';
  } catch (error) {
    console.error('❌ getQuestionText ERROR:', error.message);
    return 'Everyone\'s Answer Board';
  }
}


/**
 * システム管理者確認
 * @param {string} email - メールアドレス
 * @returns {boolean} システム管理者かどうか
 */




/**
 * 統一設定読み込みAPI - V8最適化、変数チェック強化
 * main.gs内の直接JSON.parse()操作を置換する統一関数
 * @param {string} userId - ユーザーID
 * @param {Object} preloadedUser - 事前取得済みユーザーオブジェクト（パフォーマンス最適化用）
 * @returns {Object} {success: boolean, config: Object, message?: string, userId?: string}
 */
function getUserConfig(userId, preloadedUser = null) {
  if (!userId || typeof userId !== 'string' || !userId.trim()) {
    return {
      success: false,
      config: getDefaultConfig(null),
      message: 'Invalid userId provided'
    };
  }

  try {
    const user = preloadedUser || findUserById(userId, {
      requestingUser: getCurrentEmail()
    });
    if (!user) {
      return {
        success: false,
        config: getDefaultConfig(userId),
        message: 'User not found',
        userId
      };
    }

    if (!user.configJson || typeof user.configJson !== 'string') {
      return {
        success: true,
        config: getDefaultConfig(userId),
        message: 'No config found, using defaults',
        userId
      };
    }

    const config = parseAndRepairConfig(user.configJson, userId);

    return {
      success: true,
      config,
      userId,
      message: 'Config loaded successfully'
    };

  } catch (error) {
    console.error('getUserConfig error:', {
      userId: userId ? `${userId.substring(0, 8)}***` : 'N/A',
      error: error.message
    });

    return {
      success: false,
      config: getDefaultConfig(userId),
      message: error && error.message ? `Config load error: ${error.message}` : 'Config load error: 詳細不明',
      userId
    };
  }
}

/**
 * 統一設定保存API - 検証+サニタイズ必須
 * main.gs内の直接JSON.stringify()操作を置換する統一関数
 * @param {string} userId - ユーザーID
 * @param {Object} config - 保存する設定オブジェクト
 * @param {Object} options - 保存オプション
 * @returns {Object} {success: boolean, message: string, data?: Object}
 */
function saveUserConfig(userId, config, options = {}) {
  if (!userId || typeof userId !== 'string' || !userId.trim()) {
    return {
      success: false,
      message: 'Invalid userId provided'
    };
  }

  if (!config || typeof config !== 'object') {
    return {
      success: false,
      message: 'Invalid config object provided'
    };
  }

  try {
    const user = findUserById(userId, { requestingUser: getCurrentEmail() });
    if (!user) {
      return {
        success: false,
        message: 'User not found'
      };
    }

    if (config.etag) {
      if (user.configJson) {
        try {
          const currentConfig = JSON.parse(user.configJson);
          const currentETag = currentConfig.etag || user.lastModified;

          if (currentETag && config.etag !== currentETag) {
            console.warn('saveUserConfig: ETag mismatch detected', {
              userId: userId && typeof userId === 'string' ? `${userId.substring(0, 8)}***` : 'N/A',
              requestETag: config.etag,
              currentETag
            });

            return {
              success: false,
              error: 'etag_mismatch',
              message: 'Configuration has been modified by another user',
              currentConfig
            };
          }
        } catch (parseError) {
          console.warn('saveUserConfig: Current config parse error for ETag validation:', parseError.message);
        }
      }
    }

    const validation = validateAndSanitizeConfig(config, userId);
    if (!validation.success) {
      return {
        success: false,
        message: validation.message,
        errors: validation.errors
      };
    }

    const cleanedConfig = cleanConfigFields(validation.data, options);


    if (!cleanedConfig.lastAccessedAt) {
      cleanedConfig.lastAccessedAt = new Date().toISOString();
    }

    const currentTime = new Date().toISOString();
    cleanedConfig.etag = `${currentTime}_${Math.random().toString(36).substring(2, 15)}`;

    const updateResult = updateUser(userId, {
      configJson: JSON.stringify(cleanedConfig)
    });

    if (!updateResult || !updateResult.success) {
      return {
        success: false,
        message: updateResult?.message || 'Database update failed'
      };
    }

    clearConfigCache(userId);

    if (user.spreadsheetId) {
      const saValidationKey = `sa_validation_${user.spreadsheetId}`;
      try {
        CacheService.getScriptCache().remove(saValidationKey);
      } catch (cacheRemoveError) {
        console.warn('saveUserConfig: Failed to clear SA validation cache:', cacheRemoveError.message);
      }
    }

    return {
      success: true,
      message: 'Config saved successfully',
      data: cleanedConfig,
      config: cleanedConfig, // フロントエンド互換性のため
      etag: cleanedConfig.etag, // 楽観的ロック用
      userId
    };

  } catch (error) {
    console.error('saveUserConfig error:', {
      userId: userId ? `${userId.substring(0, 8)}***` : 'N/A',
      error: error.message
    });

    return {
      success: false,
      message: error && error.message ? `Config save error: ${error.message}` : 'Config save error: 詳細不明'
    };
  }
}


/**
 * 共通フィールドクリーンアップ - main.gs内の個別delete操作を統一
 * @param {Object} config - 設定オブジェクト
 * @param {Object} options - クリーンアップオプション
 * @returns {Object} クリーンアップ済み設定
 */
function cleanConfigFields(config, options = {}) {
  if (!config || typeof config !== 'object') {
    return config;
  }

  const cleanedConfig = { ...config };

  const fieldsToRemove = [
    'setupComplete',  // レガシーフィールド
    'isDraft',        // レガシーフィールド
    'questionText'    // 動的生成されるフィールド
  ];


  fieldsToRemove.forEach(field => {
    if (field in cleanedConfig) {
      delete cleanedConfig[field];
    }
  });

  if (options.updateAccessTime !== false) {
    cleanedConfig.lastAccessedAt = new Date().toISOString();
  }

  return cleanedConfig;
}
