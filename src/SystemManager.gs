/**
 * システム管理・テスト・最適化統合モジュール
 * 分散していたテスト・最適化・診断機能を一元管理
 * 2025年2月版 - Claude.md準拠
 */

/** @OnlyCurrentDoc */

// =============================================================================
// テスト実行関数群
// =============================================================================

/**
 * メインのデータベース最適化テスト関数
 * GASエディタから直接実行可能
 * @returns {Object} 実行結果
 */
function testDatabaseMigration() {
  try {
    console.info('🔧 データベース最適化テスト開始（2025年2月版）');

    // 新しい最適化関数を実行
    const result = SystemManager.optimizeDatabaseConfigJson();

    console.info('🔧 最適化結果:', result);

    return {
      status: 'success',
      message: '✅ データベース最適化完了',
      details: {
        processedRows: result.processedRows,
        optimizedRows: result.optimizedRows,
        errorRows: result.errorRows,
        timestamp: result.timestamp,
        changes: [
          'appNameフィールドを削除',
          'opinionHeaderフィールドを追加',
          '表示設定デフォルト値をfalse（心理的安全性重視）に変更',
          '不要フィールドをクリーンアップ',
        ],
      },
    };
  } catch (error) {
    console.error('🔧 データベース最適化テストエラー:', error.message);
    return {
      status: 'error',
      message: '❌ データベース最適化エラー: ' + error.message,
    };
  }
}

/**
 * データベース構造最適化テスト（14項目→9項目）
 * GASエディタから直接実行可能
 * @returns {Object} 最適化結果
 */
function testSchemaOptimization() {
  try {
    console.info('🔄 データベース構造最適化テスト開始');
    
    // 構造最適化を実行
    const result = SystemManager.migrateToSimpleSchema();
    
    console.info('🔄 最適化結果:', result);
    
    return {
      status: result.status,
      message: result.message,
      details: result.details,
      summary: {
        削除項目: ['spreadsheetUrl', 'formUrl', 'columnMappingJson', 'publishedAt', 'appUrl'],
        統合先: 'configJson',
        新項目数: 9,
        旧項目数: 14
      },
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    console.error('🔄 データベース構造最適化テストエラー:', error.message);
    return {
      status: 'error',
      message: '❌ データベース構造最適化テストエラー: ' + error.message,
      error: error.stack
    };
  }
}

/**
 * システム全体の診断テスト
 * @returns {Object} 診断結果
 */
function testSystemDiagnosis() {
  try {
    console.info('🏥 システム診断開始');

    const results = {
      database: SystemManager.testDatabaseConnection(),
      userCount: SystemManager.getUserCount(),
      configIntegrity: SystemManager.testConfigIntegrity(),
      securityCheck: SystemManager.testSecurity(),
      timestamp: new Date().toISOString(),
    };

    return {
      status: 'success',
      message: '✅ システム診断完了',
      results: results,
    };
  } catch (error) {
    console.error('🏥 システム診断エラー:', error.message);
    return {
      status: 'error',
      message: '❌ システム診断エラー: ' + error.message,
    };
  }
}

/**
 * セットアップテスト（既存のtestSetupを統合）
 * @returns {Object} セットアップ状態
 */
function testSetup() {
  try {
    console.info('⚙️ セットアップテスト開始');

    const setupStatus = SystemManager.checkSetupStatus();

    return {
      status: 'success',
      message: setupStatus.isComplete ? '✅ セットアップ完了' : '⚠️ セットアップ未完了',
      details: setupStatus,
    };
  } catch (error) {
    console.error('⚙️ セットアップテストエラー:', error.message);
    return {
      status: 'error',
      message: '❌ セットアップテストエラー: ' + error.message,
    };
  }
}

// =============================================================================
// SystemManager 名前空間 - 統合システム管理
// =============================================================================

const SystemManager = {
  // ---------------------------------------------------------------------------
  // データベース最適化機能
  // ---------------------------------------------------------------------------

  /**
   * データベースのconfigJsonを最適化
   * @returns {Object} 最適化結果
   */
  optimizeDatabaseConfigJson() {
    try {
      console.info('🔧 データベース最適化開始: configJson構造の更新');

      const dbId = getSecureDatabaseId();
      const service = getSheetsServiceCached();
      const sheetName = DB_CONFIG.SHEET_NAME;

      // 全データを取得
      const currentData = batchGetSheetsData(service, dbId, [`'${sheetName}'!A:N`]);
      const values = currentData.valueRanges[0].values || [];

      if (values.length <= 1) {
        console.warn('🔧 最適化対象データが存在しません');
        return {
          success: false,
          reason: 'データが存在しません',
          processedRows: 0,
          optimizedRows: 0,
          errorRows: 0,
        };
      }

      // ヘッダー行を確認
      const headers = values[0];
      const configJsonIndex = headers.indexOf('configJson');
      const columnMappingJsonIndex = headers.indexOf('columnMappingJson');

      if (configJsonIndex === -1) {
        throw new Error('configJson列が見つかりません');
      }

      // データ行の最適化処理
      const dataRows = values.slice(1);
      const optimizedRows = [];
      let optimizedCount = 0;
      let errorCount = 0;

      for (let i = 0; i < dataRows.length; i++) {
        const row = [...dataRows[i]]; // 行をコピー
        const configJsonStr = row[configJsonIndex];

        if (!configJsonStr) {
          optimizedRows.push(row);
          continue;
        }

        try {
          let config = JSON.parse(configJsonStr);
          let isModified = false;

          // 1. appNameフィールドを削除
          if ('appName' in config) {
            delete config.appName;
            isModified = true;
            console.log(`行${i + 2}: appNameフィールドを削除`);
          }

          // 2. opinionHeaderを追加
          if (!config.opinionHeader) {
            let opinionHeader = 'お題';
            if (config.columnMapping && config.columnMapping.answer) {
              opinionHeader = config.columnMapping.answer;
            }
            config.opinionHeader = opinionHeader;
            isModified = true;
            console.log(`行${i + 2}: opinionHeaderを追加: "${opinionHeader}"`);
          }

          // 3. displaySettingsのデフォルト値を更新
          if (!config.displaySettings) {
            config.displaySettings = { showNames: false, showReactions: false };
            isModified = true;
          } else if (
            config.displaySettings.showNames !== false ||
            config.displaySettings.showReactions !== false
          ) {
            config.displaySettings.showNames = false;
            config.displaySettings.showReactions = false;
            isModified = true;
          }

          // 4. 不要なフィールドのクリーンアップ
          const unnecessaryFields = ['appName', 'applicationName', 'publishedSpreadsheetId'];
          for (const field of unnecessaryFields) {
            if (field in config) {
              delete config[field];
              isModified = true;
            }
          }

          if (isModified) {
            row[configJsonIndex] = JSON.stringify(config);
            optimizedCount++;
          }
        } catch (error) {
          console.error(`行${i + 2}: configJson処理エラー:`, error.message);
          errorCount++;
        }

        optimizedRows.push(row);
      }

      // 変更があった場合のみデータベースを更新
      if (optimizedCount > 0) {
        const dataRange = `'${sheetName}'!A2:N${optimizedRows.length + 1}`;
        updateSheetsData(service, dbId, dataRange, optimizedRows);

        // キャッシュをクリア
        if (typeof userIndexCache !== 'undefined') {
          userIndexCache.byUserId.clear();
          userIndexCache.byEmail.clear();
          userIndexCache.lastUpdate = 0;
        }
      }

      console.info('🔧 データベース最適化完了:', {
        処理行数: dataRows.length,
        最適化行数: optimizedCount,
        エラー行数: errorCount,
      });

      return {
        success: true,
        processedRows: dataRows.length,
        optimizedRows: optimizedCount,
        errorRows: errorCount,
        timestamp: new Date().toISOString(),
      };
    } catch (error) {
      console.error('🔧 データベース最適化エラー:', error.message);
      throw new Error(`データベース最適化に失敗しました: ${error.message}`);
    }
  },

  /**
   * データベースの現在の状態を分析
   * @returns {Object} 分析結果
   */
  analyzeDatabaseState() {
    try {
      console.info('📊 データベース分析開始');

      const dbId = getSecureDatabaseId();
      const service = getSheetsServiceCached();
      const sheetName = DB_CONFIG.SHEET_NAME;

      const currentData = batchGetSheetsData(service, dbId, [`'${sheetName}'!A:N`]);
      const values = currentData.valueRanges[0].values || [];

      if (values.length <= 1) {
        return {
          totalRows: 0,
          hasAppName: 0,
          needsOpinionHeader: 0,
          needsDisplayUpdate: 0,
        };
      }

      const headers = values[0];
      const configJsonIndex = headers.indexOf('configJson');
      const dataRows = values.slice(1);

      let hasAppName = 0;
      let needsOpinionHeader = 0;
      let needsDisplayUpdate = 0;

      for (const row of dataRows) {
        const configJsonStr = row[configJsonIndex];
        if (!configJsonStr) continue;

        try {
          const config = JSON.parse(configJsonStr);

          if ('appName' in config) hasAppName++;
          if (!config.opinionHeader) needsOpinionHeader++;
          if (
            !config.displaySettings ||
            config.displaySettings.showNames !== false ||
            config.displaySettings.showReactions !== false
          ) {
            needsDisplayUpdate++;
          }
        } catch (e) {
          // パースエラーは無視
        }
      }

      return {
        totalRows: dataRows.length,
        hasAppName: hasAppName,
        needsOpinionHeader: needsOpinionHeader,
        needsDisplayUpdate: needsDisplayUpdate,
      };
    } catch (error) {
      console.error('📊 データベース分析エラー:', error.message);
      throw error;
    }
  },

  // ---------------------------------------------------------------------------
  // 診断・テスト機能
  // ---------------------------------------------------------------------------

  /**
   * データベース接続テスト
   * @returns {Object} 接続結果
   */
  testDatabaseConnection() {
    try {
      const dbId = getSecureDatabaseId();
      const service = getSheetsServiceCached();
      const sheetName = DB_CONFIG.SHEET_NAME;

      // 簡単な読み取りテスト
      const testData = batchGetSheetsData(service, dbId, [`'${sheetName}'!A1:A2`]);

      return {
        success: true,
        databaseId: dbId ? dbId.substring(0, 10) + '...' : 'なし',
        sheetName: sheetName,
        hasData: testData.valueRanges[0].values && testData.valueRanges[0].values.length > 0,
      };
    } catch (error) {
      return {
        success: false,
        error: error.message,
      };
    }
  },

  /**
   * ユーザー数を取得
   * @returns {Object} ユーザー統計
   */
  getUserCount() {
    try {
      const dbId = getSecureDatabaseId();
      const service = getSheetsServiceCached();
      const sheetName = DB_CONFIG.SHEET_NAME;

      const data = batchGetSheetsData(service, dbId, [`'${sheetName}'!A:A`]);
      const rows = data.valueRanges[0].values || [];

      return {
        total: Math.max(0, rows.length - 1), // ヘッダー行を除く
        timestamp: new Date().toISOString(),
      };
    } catch (error) {
      return {
        total: 0,
        error: error.message,
      };
    }
  },

  /**
   * configJson整合性テスト
   * @returns {Object} 整合性結果
   */
  testConfigIntegrity() {
    try {
      const analysisResult = this.analyzeDatabaseState();

      const totalRows = analysisResult.totalRows;
      const issues =
        analysisResult.hasAppName +
        analysisResult.needsOpinionHeader +
        analysisResult.needsDisplayUpdate;

      return {
        totalRows: totalRows,
        issuesFound: issues,
        isHealthy: issues === 0,
        details: analysisResult,
      };
    } catch (error) {
      return {
        isHealthy: false,
        error: error.message,
      };
    }
  },

  /**
   * セキュリティ設定テスト
   * @returns {Object} セキュリティ状態
   */
  testSecurity() {
    try {
      const props = PropertiesService.getScriptProperties();

      const hasServiceAccount = !!props.getProperty(PROPS_KEYS.SERVICE_ACCOUNT_CREDS);
      const hasDatabaseId = !!props.getProperty(PROPS_KEYS.DATABASE_SPREADSHEET_ID);
      const hasAdminEmail = !!props.getProperty(PROPS_KEYS.ADMIN_EMAIL);

      return {
        hasServiceAccount: hasServiceAccount,
        hasDatabaseId: hasDatabaseId,
        hasAdminEmail: hasAdminEmail,
        isComplete: hasServiceAccount && hasDatabaseId && hasAdminEmail,
      };
    } catch (error) {
      return {
        isComplete: false,
        error: error.message,
      };
    }
  },

  /**
   * セットアップ状態チェック
   * @returns {Object} セットアップ状態
   */
  checkSetupStatus() {
    try {
      const security = this.testSecurity();
      const database = this.testDatabaseConnection();
      const userStats = this.getUserCount();

      return {
        isComplete: security.isComplete && database.success,
        security: security,
        database: database,
        userCount: userStats.total,
        timestamp: new Date().toISOString(),
      };
    } catch (error) {
      return {
        isComplete: false,
        error: error.message,
      };
    }
  },

  // ---------------------------------------------------------------------------
  // 最適化・メンテナンス機能
  // ---------------------------------------------------------------------------

  /**
   * 単一ユーザーのconfigJsonを最適化
   * @param {string} userId - ユーザーID
   * @returns {Object} 最適化結果
   */
  optimizeUserConfigJson(userId) {
    try {
      console.info(`👤 ユーザー ${userId} のconfigJson最適化開始`);

      const userInfo = DB.findUserById(userId);
      if (!userInfo || !userInfo.configJson) {
        return { success: false, reason: 'ユーザーまたはconfigJsonが存在しません' };
      }

      let config = JSON.parse(userInfo.configJson);
      let isModified = false;

      // 最適化処理
      if ('appName' in config) {
        delete config.appName;
        isModified = true;
      }

      if (!config.opinionHeader) {
        let opinionHeader = 'お題';
        if (config.columnMapping && config.columnMapping.answer) {
          opinionHeader = config.columnMapping.answer;
        }
        config.opinionHeader = opinionHeader;
        isModified = true;
      }

      if (!config.displaySettings) {
        config.displaySettings = { showNames: false, showReactions: false };
        isModified = true;
      } else if (
        config.displaySettings.showNames !== false ||
        config.displaySettings.showReactions !== false
      ) {
        config.displaySettings.showNames = false;
        config.displaySettings.showReactions = false;
        isModified = true;
      }

      if (isModified) {
        DB.updateUser(userId, { configJson: JSON.stringify(config) });
        console.info(`👤 ユーザー ${userId} のconfigJson最適化完了`);
        return { success: true, userId: userId, updated: true };
      }

      return { success: true, userId: userId, updated: false, reason: '変更不要' };
    } catch (error) {
      console.error(`👤 ユーザー ${userId} の最適化エラー:`, error.message);
      return { success: false, userId: userId, error: error.message };
    }
  },

  /**
   * データベース構造を最適化（重複項目削除）
   * 14項目 → 9項目に簡素化
   * @returns {Object} 最適化結果
   */
  migrateToSimpleSchema() {
    try {
      console.info('🔄 データベース構造最適化開始（14項目→9項目）');
      
      const allUsers = DB.getAllUsers();
      const results = {
        totalUsers: allUsers.length,
        migratedUsers: 0,
        errors: [],
        removedColumns: ['spreadsheetUrl', 'formUrl', 'columnMappingJson', 'publishedAt', 'appUrl'],
        timestamp: new Date().toISOString()
      };
      
      // 各ユーザーのconfigJsonを統合・移行
      allUsers.forEach((user) => {
        try {
          let config = {};
          
          // 既存のconfigJsonをベースに
          if (user.configJson) {
            config = JSON.parse(user.configJson);
          }
          
          // 重複項目をconfigJsonに統合
          if (user.formUrl && !config.formUrl) {
            config.formUrl = user.formUrl;
          }
          
          if (user.columnMappingJson && !config.columnMapping) {
            try {
              config.columnMapping = JSON.parse(user.columnMappingJson);
            } catch (e) {
              console.warn(`columnMappingJson解析エラー (userId: ${user.userId}):`, e.message);
            }
          }
          
          // opinionHeaderの追加（未設定の場合）
          if (!config.opinionHeader && config.columnMapping?.answer) {
            config.opinionHeader = config.columnMapping.answer;
          }
          
          // プライバシー設定のデフォルト化
          if (!config.displaySettings) {
            config.displaySettings = {
              showNames: false,
              showReactions: false
            };
          }
          
          // appNameの削除（不要）
          delete config.appName;
          
          // 更新対象データ（新スキーマ項目のみ）
          const updateData = {
            configJson: JSON.stringify(config),
            lastModified: new Date().toISOString()
          };
          
          DB.updateUser(user.userId, updateData);
          results.migratedUsers++;
          
        } catch (error) {
          results.errors.push({
            userId: user.userId,
            userEmail: user.userEmail,
            error: error.message
          });
        }
      });
      
      console.info('🔄 データベース構造最適化完了:', {
        総ユーザー数: results.totalUsers,
        移行成功: results.migratedUsers,
        エラー数: results.errors.length,
        削除項目: results.removedColumns
      });
      
      return {
        status: 'success',
        message: `✅ データベース構造最適化完了（${results.migratedUsers}/${results.totalUsers}ユーザー）`,
        details: results
      };
      
    } catch (error) {
      console.error('🔄 データベース構造最適化エラー:', error.message);
      return {
        status: 'error', 
        message: '❌ データベース構造最適化エラー: ' + error.message,
        error: error.stack
      };
    }
  },

  /**
   * システム全体のクリーンアップ
   * @returns {Object} クリーンアップ結果
   */
  performSystemCleanup() {
    try {
      console.info('🧹 システムクリーンアップ開始');

      const results = {
        cacheCleared: false,
        configOptimized: false,
        errors: [],
      };

      // キャッシュクリア
      try {
        if (typeof userIndexCache !== 'undefined') {
          userIndexCache.byUserId.clear();
          userIndexCache.byEmail.clear();
          userIndexCache.lastUpdate = 0;
        }
        results.cacheCleared = true;
      } catch (error) {
        results.errors.push('キャッシュクリアエラー: ' + error.message);
      }

      // config最適化
      try {
        const optimizeResult = this.optimizeDatabaseConfigJson();
        results.configOptimized = optimizeResult.success;
        results.optimizedRows = optimizeResult.optimizedRows;
      } catch (error) {
        results.errors.push('config最適化エラー: ' + error.message);
      }

      console.info('🧹 システムクリーンアップ完了:', results);

      return {
        success: results.errors.length === 0,
        results: results,
        timestamp: new Date().toISOString(),
      };
    } catch (error) {
      console.error('🧹 システムクリーンアップエラー:', error.message);
      return {
        success: false,
        error: error.message,
      };
    }
  },
};

// =============================================================================
// レガシー関数の統合（既存システムとの互換性維持）
// =============================================================================

/**
 * レガシー最適化関数の統合版
 * ConfigOptimizer.gsの機能を統合
 */
function optimizeConfigJson(currentConfig, userInfo) {
  return SystemManager.optimizeUserConfigJson(userInfo.userId);
}

/**
 * レガシーユーザー最適化関数
 */
function optimizeUserDatabase(userId) {
  return SystemManager.optimizeUserConfigJson(userId);
}

/**
 * レガシー診断関数
 */
function diagnoseDatabase(targetUserId) {
  if (targetUserId) {
    return SystemManager.optimizeUserConfigJson(targetUserId);
  } else {
    return SystemManager.checkSetupStatus();
  }
}

// =============================================================================
// 外部実行用のエントリーポイント
// =============================================================================

/**
 * データベース最適化の対話的実行
 * スプレッドシートから実行可能
 */
function runDatabaseOptimizationInteractive() {
  try {
    const ui = SpreadsheetApp.getUi();

    // 分析結果を表示
    const analysisResult = SystemManager.analyzeDatabaseState();

    const response = ui.alert(
      'データベース最適化',
      `以下の変更を実行します:\n\n` +
        `- 処理対象行数: ${analysisResult.totalRows}\n` +
        `- appName削除対象: ${analysisResult.hasAppName}行\n` +
        `- opinionHeader追加対象: ${analysisResult.needsOpinionHeader}行\n` +
        `- 表示設定更新対象: ${analysisResult.needsDisplayUpdate}行\n\n` +
        `実行しますか？`,
      ui.ButtonSet.YES_NO
    );

    if (response === ui.Button.YES) {
      const result = SystemManager.optimizeDatabaseConfigJson();

      ui.alert(
        '最適化完了',
        `データベースの最適化が完了しました。\n\n` +
          `- 処理行数: ${result.processedRows}\n` +
          `- 最適化行数: ${result.optimizedRows}\n` +
          `- エラー行数: ${result.errorRows}`,
        ui.ButtonSet.OK
      );

      return result;
    } else {
      return { status: 'cancelled', message: 'ユーザーによってキャンセルされました' };
    }
  } catch (error) {
    console.error('対話的最適化エラー:', error.message);
    return { status: 'error', message: error.message };
  }
}

/**
 * システム診断レポート生成
 * スプレッドシートから実行可能
 */
function generateSystemReport() {
  try {
    const diagnosis = testSystemDiagnosis();
    const setupTest = testSetup();

    console.log('=== システム診断レポート ===');
    console.log('診断結果:', diagnosis);
    console.log('セットアップ状態:', setupTest);
    console.log('===============================');

    return {
      diagnosis: diagnosis,
      setup: setupTest,
      timestamp: new Date().toISOString(),
    };
  } catch (error) {
    console.error('システムレポート生成エラー:', error.message);
    return { error: error.message };
  }
}
