/**
 * システム管理統合モジュール（シンプル版）
 * 必要最小限の機能に絞り込み - 2025年9月版
 * 
 * 機能：
 * - システムヘルスチェック（セキュリティ・DB接続）
 * - configJSON重複ネスト修正
 * - ユーザー設定デフォルトリセット
 * - 基本診断機能
 */

/** @OnlyCurrentDoc */

// =============================================================================
// システムヘルスチェック機能
// =============================================================================

/**
 * 🔐 セキュリティ設定確認（サービスアカウント・トークン）
 * @returns {Object} セキュリティチェック結果
 */
function testSecurity() {
  try {
    const props = PropertiesService.getScriptProperties();

    const hasServiceAccount = !!props.getProperty(PROPS_KEYS.SERVICE_ACCOUNT_CREDS);
    const hasDatabaseId = !!props.getProperty(PROPS_KEYS.DATABASE_SPREADSHEET_ID);
    const hasAdminEmail = !!props.getProperty(PROPS_KEYS.ADMIN_EMAIL);

    const isComplete = hasServiceAccount && hasDatabaseId && hasAdminEmail;

    // サービスアカウントトークン生成テスト
    let tokenTest = false;
    if (hasServiceAccount) {
      try {
        const token = getServiceAccountTokenCached();
        tokenTest = !!token;
      } catch (error) {
        console.warn('トークン生成テスト失敗:', error.message);
      }
    }

    return {
      success: true,
      isComplete,
      hasServiceAccount,
      hasDatabaseId,
      hasAdminEmail,
      tokenTest,
      details: {
        serviceAccount: hasServiceAccount ? '設定済み' : '未設定',
        databaseId: hasDatabaseId ? '設定済み' : '未設定', 
        adminEmail: hasAdminEmail ? '設定済み' : '未設定',
        tokenGeneration: tokenTest ? '正常' : '失敗'
      }
    };
  } catch (error) {
    return {
      success: false,
      error: error.message,
      isComplete: false
    };
  }
}

/**
 * 📊 データベース接続確認
 * @returns {Object} DB接続チェック結果  
 */
function testDatabaseConnection() {
  try {
    const dbId = getSecureDatabaseId();
    const service = getSheetsServiceCached();
    const sheetName = DB_CONFIG.SHEET_NAME;

    // シートの基本情報取得テスト
    const sheetData = batchGetSheetsData(service, dbId, [`'${sheetName}'!A1:E1`]);
    const headers = sheetData.valueRanges[0].values?.[0] || [];

    const hasCorrectHeaders = 
      headers[0] === 'userId' &&
      headers[1] === 'userEmail' &&
      headers[2] === 'isActive' &&
      headers[3] === 'configJson' &&
      headers[4] === 'lastModified';

    return {
      success: true,
      databaseId: dbId,
      sheetName: sheetName,
      hasCorrectHeaders,
      headerCount: headers.length,
      headers: headers
    };
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 🔍 システム統合チェック
 * @returns {Object} 全体的なシステム状態
 */
function checkSetupStatus() {
  const security = testSecurity();
  const database = testDatabaseConnection(); 
  const userStats = getUserCount();

  return {
    isComplete: security.isComplete && database.success,
    security,
    database,
    userStats,
    timestamp: new Date().toISOString()
  };
}

// =============================================================================
// configJSON修正機能
// =============================================================================

/**
 * 🚨 configJSON重複ネスト修正
 * JSONの重複・ネスト問題を解決
 * @returns {Object} 修正結果
 */
function fixConfigJsonNestingImpl() {
  console.log('🔧 SystemManager.fixConfigJsonNesting: 重複ネスト修正開始');
  
  try {
    const users = DB.getAllUsers();
    const results = {
      total: users.length,
      fixed: 0,
      errors: [],
      details: []
    };

    for (const user of users) {
      try {
        const config = JSON.parse(user.configJson || '{}');
        
        // configJsonフィールドが存在する場合は除去
        if ('configJson' in config) {
          console.log(`🚨 重複ネスト発見: ${user.userId}`);
          delete config.configJson;
          
          // ConfigManager経由で修正保存
          const success = ConfigManager.saveConfig(user.userId, config);
          
          if (success) {
            results.fixed++;
            results.details.push({
              userId: user.userId,
              userEmail: user.userEmail,
              status: 'fixed'
            });
            console.log(`✅ 修正完了: ${user.userEmail}`);
          } else {
            results.errors.push({
              userId: user.userId,
              userEmail: user.userEmail,
              error: '保存失敗'
            });
          }
        }
      } catch (error) {
        results.errors.push({
          userId: user.userId,
          userEmail: user.userEmail,
          error: error.message
        });
      }
    }

    console.log('✅ SystemManager.fixConfigJsonNesting: 修正完了', {
      total: results.total,
      fixed: results.fixed,
      errors: results.errors.length
    });

    return {
      success: results.errors.length === 0,
      ...results,
      timestamp: new Date().toISOString()
    };

  } catch (error) {
    console.error('❌ SystemManager.fixConfigJsonNesting: エラー:', error.message);
    return {
      success: false,
      error: error.message,
      timestamp: new Date().toISOString()
    };
  }
}

/**
 * 🆕 ユーザー設定をデフォルトにリセット
 * configJSONと重複フィールドを完全初期化
 * @param {string} userId - 対象ユーザーID
 * @returns {Object} リセット結果
 */
function resetUserConfigToDefaultImpl(userId) {
  console.log('🔄 resetUserConfigToDefault: デフォルトリセット開始', userId);
  
  try {
    // デフォルト設定定義
    const defaultConfig = {
      createdAt: new Date().toISOString(),
      lastAccessedAt: new Date().toISOString(),
      setupStatus: 'pending',
      appPublished: false,
      displaySettings: {
        showNames: false,
        showReactions: false
      },
      configVersion: '2.0',
      claudeMdCompliant: true
    };

    // ユーザー存在確認
    const user = DB.findUserById(userId);
    if (!user) {
      throw new Error('ユーザーが見つかりません');
    }

    // 5フィールド構造でのクリーンな更新データ
    const cleanUserData = {
      userId: userId,
      userEmail: user.userEmail,
      isActive: true,
      configJson: JSON.stringify(defaultConfig),
      lastModified: new Date().toISOString()
    };

    // データベース更新（重複フィールド完全削除）
    const updateResult = DB.updateUser(userId, cleanUserData);
    
    if (updateResult) {
      console.log('✅ ユーザー設定デフォルトリセット完了:', userId);
      return {
        success: true,
        userId: userId,
        userEmail: user.userEmail,
        resetTo: 'default',
        defaultConfig: defaultConfig,
        timestamp: new Date().toISOString()
      };
    } else {
      throw new Error('データベース更新に失敗');
    }

  } catch (error) {
    console.error('❌ resetUserConfigToDefault エラー:', error.message);
    return {
      success: false,
      userId: userId,
      error: error.message,
      timestamp: new Date().toISOString()
    };
  }
}

// =============================================================================
// 基本診断機能
// =============================================================================

/**
 * 👥 ユーザー数取得
 * @returns {Object} ユーザー統計
 */
function getUserCount() {
  try {
    const users = DB.getAllUsers();
    const activeUsers = users.filter(user => user.isActive);
    
    return {
      total: users.length,
      active: activeUsers.length,
      inactive: users.length - activeUsers.length
    };
  } catch (error) {
    return {
      total: 0,
      active: 0,
      inactive: 0,
      error: error.message
    };
  }
}

/**
 * 🔍 設定整合性確認
 * @returns {Object} 設定チェック結果
 */
function testConfigIntegrity() {
  try {
    const users = DB.getAllUsers();
    let validConfigs = 0;
    let invalidConfigs = 0;
    const errors = [];

    for (const user of users) {
      try {
        const config = JSON.parse(user.configJson || '{}');
        
        // 基本フィールド確認
        if (config.setupStatus && config.displaySettings) {
          validConfigs++;
        } else {
          invalidConfigs++;
          errors.push({
            userId: user.userId,
            userEmail: user.userEmail,
            issue: '必須フィールド不足'
          });
        }
      } catch (parseError) {
        invalidConfigs++;
        errors.push({
          userId: user.userId,
          userEmail: user.userEmail,
          issue: 'JSON解析エラー'
        });
      }
    }

    return {
      success: true,
      total: users.length,
      valid: validConfigs,
      invalid: invalidConfigs,
      errors: errors
    };
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

// =============================================================================
// SystemManager メインオブジェクト
// =============================================================================

const SystemManager = {
  testSecurity,
  testDatabaseConnection, 
  checkSetupStatus,
  fixConfigJsonNesting: fixConfigJsonNestingImpl,
  resetUserConfigToDefault: resetUserConfigToDefaultImpl,
  getUserCount,
  testConfigIntegrity
};

// =============================================================================
// グローバルアクセス関数（Web UI用）
// =============================================================================

/**
 * Web UI用のconfigJSON修正関数
 */
function fixConfigJsonNesting() {
  return SystemManager.fixConfigJsonNesting();
}

/**
 * Web UI用のデフォルトリセット関数
 */
function resetUserConfigToDefault(userId) {
  return SystemManager.resetUserConfigToDefault(userId);
}

/**
 * システム診断関数（管理用）
 */
function runSystemDiagnostics() {
  console.log('🔍 システム診断実行開始');
  
  const diagnostics = {
    security: SystemManager.testSecurity(),
    database: SystemManager.testDatabaseConnection(),
    userStats: SystemManager.getUserCount(),
    configIntegrity: SystemManager.testConfigIntegrity(),
    timestamp: new Date().toISOString()
  };
  
  console.log('🔍 システム診断完了:', diagnostics);
  return diagnostics;
}