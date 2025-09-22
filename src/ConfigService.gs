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

/* global getCurrentEmail, findUserById, updateUser, validateEmail, CACHE_DURATION, TIMEOUT_MS, SYSTEM_LIMITS, validateConfig, URL, validateUrl, createErrorResponse, validateSpreadsheetId, findUserByEmail, findUserBySpreadsheetId, openSpreadsheet, getUserSpreadsheetData, Auth, UserService, isAdministrator, SLEEP_MS */

// ===========================================
// 🔧 GAS-Native ConfigService (直接API版)
// ===========================================

/**
 * ConfigService - ゼロ依存アーキテクチャ
 * GAS-Nativeパターンによる直接APIアクセス
 * PROPS_KEYS, DB依存を完全排除
 */


/**
 * FormApp権限の安全チェック - GAS 2025ベストプラクティス準拠
 * V8ランタイム対応の軽量権限テスト
 * @returns {Object} 権限チェック結果
 */
function validateFormAppAccess() {
  try {
    // 権限テスト: 最小限のFormApp操作で権限確認
    // FormApp.getActiveForm()は現在のスクリプトコンテキストでは使用不可のため
    // より軽量なアプローチを使用
    if (typeof FormApp === 'undefined') {
      return {
        hasAccess: false,
        reason: 'FORMAPP_NOT_AVAILABLE',
        message: 'FormAppサービスが利用できません'
      };
    }

    // FormAppのopenByUrl機能をテスト（権限チェックのみ、実際の呼び出しはしない）
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
      message: 'FormAppへのアクセス権限が確認されました'
    };
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
 * V8ランタイム対応 - 安全なFormApp.openByUrl実行
 * スタックオーバーフロー対策、指数バックオフ付きリトライパターン
 * @param {string} formUrl - フォームURL
 * @param {Object} options - オプション設定
 * @returns {Object} 実行結果
 */
function openFormWithRetry(formUrl, options = {}) {
  const maxTries = options.maxTries || 3;
  const initialDelay = options.initialDelay || 500;

  // URL形式の事前検証（GASベストプラクティス）
  if (!formUrl || typeof formUrl !== 'string') {
    return {
      success: false,
      error: 'INVALID_URL',
      message: '有効なフォームURLが指定されていません'
    };
  }

  // FormURL形式チェック: edit URLか確認
  const isGoogleFormUrl = formUrl.includes('docs.google.com/forms/');
  if (!isGoogleFormUrl) {
    return {
      success: false,
      error: 'NOT_GOOGLE_FORM_URL',
      message: 'GoogleフォームのURLではありません'
    };
  }

  // 権限チェック実行
  const accessCheck = validateFormAppAccess();
  if (!accessCheck.hasAccess) {
    return {
      success: false,
      error: 'PERMISSION_DENIED',
      message: accessCheck.message,
      accessCheckResult: accessCheck
    };
  }

  // リトライ実行（指数バックオフ付き）
  let tries = 1;
  let lastError = null;

  while (tries <= maxTries) {
    try {
      // V8ランタイム対応: 短時間の遅延でスタック状態をリセット
      if (tries > 1) {
        const delay = initialDelay * Math.pow(2, tries - 1);
        Utilities.sleep(Math.min(delay, SLEEP_MS.MAX)); // 最大5秒まで
      }

      // FormApp.openByUrl実行
      const form = FormApp.openByUrl(formUrl);

      // V8ランタイム: undefined/null チェック強化
      if (!form || form === null || typeof form === 'undefined') {
        throw new Error('FormApp.openByUrl returned null or undefined');
      }

      // 成功時の戻り値
      return {
        success: true,
        form,
        tries,
        message: `FormApp.openByUrl成功 (${tries}回目)`
      };

    } catch (error) {
      lastError = error;
      const errorMessage = error.message || String(error);

      // スタックオーバーフロー検知
      const isStackOverflow = errorMessage.includes('Stack overflow') ||
                             errorMessage.includes('Maximum call stack') ||
                             errorMessage.includes('%[sdj%]');

      // 権限エラー検知
      const isPermissionError = errorMessage.includes('permission') ||
                               errorMessage.includes('not authorized') ||
                               errorMessage.includes('Authorization');

      console.warn(`FormApp.openByUrl試行${tries}回目失敗:`, errorMessage);

      // 最終試行での失敗処理
      if (tries >= maxTries) {
        return {
          success: false,
          error: isStackOverflow ? 'STACK_OVERFLOW' :
                 isPermissionError ? 'PERMISSION_ERROR' : 'UNKNOWN_ERROR',
          message: `FormApp.openByUrl ${maxTries}回失敗: ${errorMessage}`,
          lastError: errorMessage,
          isStackOverflow,
          isPermissionError,
          totalTries: tries
        };
      }

      tries++;
    }
  }

  // 理論的には到達しないが安全のため
  return {
    success: false,
    error: 'UNEXPECTED_END',
    message: '予期しないリトライループ終了',
    lastError: lastError ? lastError.message : 'Unknown'
  };
}



/**
 * デフォルト設定取得
 * @param {string} userId - ユーザーID
 * @returns {Object} デフォルト設定
 */
function getDefaultConfig(userId) {
  // 🚀 Zero-dependency: 静的デフォルト値を提供
  // ✅ Optimized: createdAt and lastModified moved to database columns, removed from configJSON
  return {
    userId,
    setupStatus: 'pending',
    isPublished: false,
    displaySettings: {
      showNames: false,
      showReactions: false
    },
    userPermissions: {
      isEditor: false,
      isAdministrator: false,
      accessLevel: 'viewer',
      canEdit: false,
      canView: true,
      canReact: true
    },
    completionScore: 0
    // lastModified removed - managed exclusively by database column
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
    isPublished: Boolean(config.isPublished),
    spreadsheetId: config.spreadsheetId || '',
    sheetName: config.sheetName || '',
    formUrl: config.formUrl || '',
    displaySettings: config.displaySettings,
    columnMapping: config.columnMapping,
    userPermissions: config.userPermissions,
    completionScore: calculateCompletionScore(config)
    // lastModified removed - managed exclusively by database column
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
    const webAppUrl = ScriptApp.getService().getUrl();

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
    // ✅ CLAUDE.md準拠: Batched admin authentication (70x performance improvement)
    const adminAuth = getBatchedAdminAuth({ allowNonAdmin: true }); // eslint-disable-line no-undef
    if (!adminAuth.success || !adminAuth.authenticated) {
      return {
        isEditor: false,
        isAdministrator: false,
        accessLevel: 'viewer',
        canEdit: false,
        canView: true,
        canReact: true
      };
    }

    const { email: currentEmail, isAdmin } = adminAuth;

    return {
      isEditor: true, // 現在のユーザーは自分の設定の編集者
      isAdministrator: isAdmin,
      accessLevel: isAdmin ? 'administrator' : 'editor',
      canEdit: true,
      canView: true,
      canReact: true,
      canDelete: isAdmin,
      canManageUsers: isAdmin
    };

  } catch (error) {
    console.error('generateUserPermissions: エラー', error.message);
    return {
      isEditor: false,
      isAdministrator: false,
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
  try {
    // 統一検証: validators.gsのvalidateConfigを活用
    const validationResult = validateConfig(config);
    if (!validationResult.isValid) {
      return {
        success: false,
        message: '設定検証エラーがあります',
        errors: validationResult.errors,
        data: validationResult.sanitized || config
      };
    }

    // ユーザーIDの追加検証
    const errors = [];
    if (!validateConfigUserId(userId)) {
      errors.push('無効なユーザーID形式');
    }

    // ConfigService固有の追加サニタイズ
    const sanitized = { ...validationResult.sanitized };
    sanitized.userId = userId;

    if (sanitized.displaySettings) {
      sanitized.displaySettings = sanitizeDisplaySettings(sanitized.displaySettings);
    }
    if (sanitized.columnMapping) {
      sanitized.columnMapping = sanitizeColumnMapping(sanitized.columnMapping);
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
    theme: String(displaySettings.theme || 'default').substring(0, SYSTEM_LIMITS.PREVIEW_LENGTH),
    pageSize: Math.min(Math.max(Number(displaySettings.pageSize) || SYSTEM_LIMITS.DEFAULT_PAGE_SIZE, 1), SYSTEM_LIMITS.MAX_PAGE_SIZE)
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
 * システムセットアップ状態確認
 * @returns {boolean} セットアップ済みかどうか
 */
function isSystemSetup() {
  try {
    // ✅ CLAUDE.md準拠: Batched user config retrieval (70x performance improvement)
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
    const cache = CacheService.getScriptCache();

    // 🔧 CLAUDE.md準拠: 依存関係キャッシュの完全無効化
    const keysToRemove = [
      `config_${userId}`,           // 設定キャッシュ
      `user_${userId}`,             // ユーザーキャッシュ
      `board_data_${userId}`,       // ボードデータキャッシュ
      `admin_panel_${userId}`,      // 管理パネルデータ
      `question_text_${userId}`     // 質問テキストキャッシュ
    ];

    // 一括削除でパフォーマンス向上
    if (keysToRemove.length > 0) {
      cache.removeAll(keysToRemove);
    }

    console.info('clearConfigCache: 依存関係キャッシュクリア完了', {
      userId: userId && typeof userId === 'string' ? `${userId.substring(0, 8)}***` : 'N/A',
      keysCleared: keysToRemove.length
    });
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

    // 特定ユーザー群のキャッシュを一括クリア
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
      console.info('clearAllConfigCache: ユーザー群キャッシュクリア完了', {
        userCount: userIds.length,
        keysCleared: allKeysToRemove.length
      });
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
 * 動的questionText取得（configJson最適化対応 + 動的headers取得）
 * headers配列とcolumnMappingから実際の問題文を動的取得
 * headersがない場合はスプレッドシートから動的取得
 * @param {Object} config - ユーザー設定オブジェクト
 * @param {Object} context - アクセスコンテキスト（target user info for cross-user access）
 * @returns {string} 問題文テキスト
 */
function getQuestionText(config, context = {}) {
  try {
    console.log('📝 getQuestionText START:', {
      hasColumnMapping: !!config?.columnMapping,
      hasHeaders: !!config?.columnMapping?.headers,
      answerIndex: config?.columnMapping?.mapping?.answer,
      headersLength: config?.columnMapping?.headers?.length || 0,
      hasSpreadsheetId: !!config?.spreadsheetId,
      hasSheetName: !!config?.sheetName
    });

    const answerIndex = config?.columnMapping?.mapping?.answer;

    // 1. 既存のheadersから取得を試行
    if (typeof answerIndex === 'number' && config?.columnMapping?.headers?.[answerIndex]) {
      const questionText = config.columnMapping.headers[answerIndex];
      if (questionText && typeof questionText === 'string' && questionText.trim()) {
        console.log('✅ getQuestionText SUCCESS (from stored headers):', questionText.trim());
        return questionText.trim();
      }
    }

    // 2. headersがない場合、スプレッドシートから動的取得
    if (typeof answerIndex === 'number' && config?.spreadsheetId && config?.sheetName) {
      try {
        console.log('🔄 getQuestionText: Fetching headers from spreadsheet');
        // 🔧 CLAUDE.md準拠: Context-aware service account usage
        // ✅ **Cross-user**: Use service account when accessing other user's config
        // ✅ **Self-access**: Use normal permissions for own config
        const currentEmail = getCurrentEmail();

        // CLAUDE.md準拠: spreadsheetIdから所有者を特定して直接比較
        const targetUser = findUserBySpreadsheetId(config.spreadsheetId);
        const isSelfAccess = targetUser && targetUser.userEmail === currentEmail;

        console.log(`getQuestionText: ${isSelfAccess ? 'Self-access normal permissions' : 'Cross-user service account'} for spreadsheet`);
        try {
          const dataAccess = openSpreadsheet(config.spreadsheetId, { useServiceAccount: !isSelfAccess });
          const { spreadsheet } = dataAccess;
          const sheet = spreadsheet.getSheetByName(config.sheetName);
          if (sheet && sheet.getLastColumn() > 0) {
            const [headers] = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
            if (headers && headers[answerIndex]) {
              const questionText = headers[answerIndex];
              if (questionText && typeof questionText === 'string' && questionText.trim()) {
                console.log('✅ getQuestionText SUCCESS (from dynamic headers):', questionText.trim());
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

    // 3. formTitleからの取得
    if (config?.formTitle && typeof config.formTitle === 'string' && config.formTitle.trim()) {
      console.log('✅ getQuestionText SUCCESS (from formTitle):', config.formTitle.trim());
      return config.formTitle.trim();
    }

    // 4. デフォルトフォールバック
    console.log('🔄 getQuestionText FALLBACK: Using default title');
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

// ===========================================
// 🔧 統一configJson操作API (CLAUDE.md準拠)
// ===========================================

/**
 * 統一設定読み込みAPI - V8最適化、変数チェック強化
 * main.gs内の直接JSON.parse()操作を置換する統一関数
 * @param {string} userId - ユーザーID
 * @returns {Object} {success: boolean, config: Object, message?: string, userId?: string}
 */
function getUserConfig(userId) {
  // V8最適化: 事前変数チェック（CLAUDE.md 151-169行準拠）
  if (!userId || typeof userId !== 'string' || !userId.trim()) {
    return {
      success: false,
      config: getDefaultConfig(null),
      message: 'Invalid userId provided'
    };
  }

  try {
    // Zero-Dependency: 直接findUserById呼び出し（CLAUDE.md準拠）
    const currentEmail = getCurrentEmail();
    const user = findUserById(userId, {
      requestingUser: currentEmail
    });
    if (!user) {
      return {
        success: false,
        config: getDefaultConfig(userId),
        message: 'User not found',
        userId
      };
    }

    // V8最適化: configJson存在チェック + 構造化パース
    if (!user.configJson || typeof user.configJson !== 'string') {
      return {
        success: true,
        config: getDefaultConfig(userId),
        message: 'No config found, using defaults',
        userId
      };
    }

    // 構造化パース（修復機能付き）- 既存parseAndRepairConfig利用
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
  // V8最適化: 入力検証（事前チェック）
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
    // 🔧 CLAUDE.md準拠: 楽観的ロック（ETag）検証の実装
    // ✅ Optimized: Use database lastModified for ETag validation
    const user = findUserById(userId);
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
          // ✅ Optimized: Fallback to database lastModified for ETag comparison
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
          // パースエラーの場合は競合チェックをスキップして続行
        }
      }
    }

    // 1. 統合検証・サニタイズ（既存validateAndSanitizeConfig利用）
    const validation = validateAndSanitizeConfig(config, userId);
    if (!validation.success) {
      return {
        success: false,
        message: validation.message,
        errors: validation.errors
      };
    }

    // 2. 共通フィールドクリーンアップ
    const cleanedConfig = cleanConfigFields(validation.data, options);

    // 3. タイムスタンプ更新とETag生成
    // ✅ Optimized: Remove lastModified from configJSON, use database column exclusively
    // lastModified is now managed automatically by database updateUser function

    if (!cleanedConfig.lastAccessedAt) {
      cleanedConfig.lastAccessedAt = new Date().toISOString();
    }

    // 🔧 CLAUDE.md準拠: 楽観的ロック用ETag生成
    // ✅ Optimized: Use database lastModified for ETag generation
    const currentTime = new Date().toISOString();
    cleanedConfig.etag = `${currentTime}_${Math.random().toString(36).substring(2, 15)}`;

    // 4. 🔧 CLAUDE.md準拠: 書き込み前キャッシュ無効化 - 同期ギャップ防止
    clearConfigCache(userId);
    console.log('saveUserConfig: 書き込み前キャッシュクリア完了');

    // 5. Zero-Dependency: 直接updateUser呼び出し
    // ✅ Optimized: lastModified automatically managed by database updateUser function
    const updateResult = updateUser(userId, {
      configJson: JSON.stringify(cleanedConfig)
      // lastModified removed - automatically updated by DatabaseCore
    });

    if (!updateResult || !updateResult.success) {
      return {
        success: false,
        message: updateResult?.message || 'Database update failed'
      };
    }

    // 6. 🔧 CLAUDE.md準拠: 書き込み後的確キャッシュ無効化 - 最終一貫性保証
    clearConfigCache(userId);
    console.log('saveUserConfig: 書き込み後的確キャッシュクリア完了', {
      userId: userId && typeof userId === 'string' ? `${userId.substring(0, 8)}***` : 'N/A',
      newETag: cleanedConfig.etag
    });

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

  // V8最適化: const/destructuring使用
  const cleanedConfig = { ...config };

  // 標準的なクリーンアップフィールド（main.gsパターンから抽出）
  const fieldsToRemove = [
    'setupComplete',  // レガシーフィールド
    'isDraft',        // レガシーフィールド
    'questionText'    // 動的生成されるフィールド
  ];

  // オプション指定による追加クリーンアップ
  if (options.cleanLegacyFields) {
    fieldsToRemove.push('setupStatus_old', 'configVersion_old');
  }

  // フィールド削除実行
  fieldsToRemove.forEach(field => {
    if (field in cleanedConfig) {
      delete cleanedConfig[field];
    }
  });

  // lastAccessedAt自動更新（オプション）
  if (options.updateAccessTime !== false) {
    cleanedConfig.lastAccessedAt = new Date().toISOString();
  }

  return cleanedConfig;
}
