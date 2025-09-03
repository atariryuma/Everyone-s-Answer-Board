/**
 * @fileoverview データベース移行機能 - レガシー構造→5フィールド超効率化移行
 * configJSON中心型アーキテクチャへの自動移行
 */

/**
 * 🚀 データベースレガシー構造→5フィールド超効率化移行
 * configJSON中心設計でパフォーマンス大幅向上を実現
 */
function migrateToConfigJsonCentric() {
  console.log('🚀 データベース超効率化移行開始：レガシー構造→5フィールド構造');
  
  try {
    // 1. 現在のデータベーススキーマ確認
    const currentSchema = getCurrentDatabaseSchema();
    console.log('📋 現在のスキーマ:', currentSchema);
    
    if (currentSchema.columnCount === 5) {
      console.log('✅ すでに5項目最適化済みです');
      return { success: true, message: 'すでに最適化済み' };
    }
    
    // 2. 全ユーザーデータ取得
    const allUsers = getAllUsersForMigration();
    console.log(`📊 移行対象ユーザー数: ${allUsers.length}`);
    
    // 3. データベースバックアップ
    const backupResult = createDatabaseBackup();
    console.log('💾 バックアップ作成:', backupResult.success);
    
    // 4. 5項目構造に変換
    const convertedUsers = allUsers.map(user => convertUserToConfigJsonCentric(user));
    console.log(`⚡ データ変換完了: ${convertedUsers.length}件`);
    
    // 5. 新しいスキーマでデータベースを再構築
    const migrationResult = rebuildDatabaseWith5Columns(convertedUsers);
    
    if (migrationResult.success) {
      console.log('🎉 データベース超効率化移行完了！');
      return {
        success: true,
        message: '5項目超効率化移行完了',
        performance: {
          previousColumns: currentSchema.columnCount,
          newColumns: 5,
          reduction: `${Math.round((1 - 5/currentSchema.columnCount) * 100)}%削減`,
          migratedUsers: convertedUsers.length
        }
      };
    } else {
      throw new Error(`移行失敗: ${migrationResult.error}`);
    }
    
  } catch (error) {
    console.error('❌ 移行エラー:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * ユーザーデータを5項目configJSON中心型に変換
 * @param {Object} user - 既存ユーザーデータ
 * @returns {Object} - 5項目最適化ユーザーデータ
 */
function convertUserToConfigJsonCentric(user) {
  // configJSON統合型データ構造作成
  const optimizedConfigJson = {
    // 🔄 DB列から移行するデータ
    createdAt: user.createdAt || new Date().toISOString(),
    lastAccessedAt: user.lastAccessedAt || new Date().toISOString(),
    spreadsheetId: user.spreadsheetId || '',
    sheetName: user.sheetName || '',
    
    // 📊 動的生成URL（キャッシュ）
    spreadsheetUrl: user.spreadsheetId ? 
      `https://docs.google.com/spreadsheets/d/${user.spreadsheetId}` : '',
    appUrl: `${ScriptApp.getService().getUrl()}?mode=view&userId=${user.userId}`,
    
    // ⚙️ 既存configJsonの内容をマージ
    ...(() => {
      try {
        return JSON.parse(user.configJson || '{}');
      } catch (e) {
        console.warn(`configJson parse error for ${user.userId}:`, e);
        return {};
      }
    })(),
    
    // 📅 メタデータ更新
    migratedAt: new Date().toISOString(),
    schemaVersion: '5-column-optimized'
  };
  
  // 5項目最適化構造
  return {
    userId: user.userId,
    userEmail: user.userEmail,
    isActive: user.isActive !== false, // デフォルトtrue
    configJson: JSON.stringify(optimizedConfigJson),
    lastModified: new Date().toISOString()
  };
}

/**
 * 現在のデータベーススキーマを取得
 */
function getCurrentDatabaseSchema() {
  const service = getServiceAccountTokenCached();
  const dbId = PropertiesService.getScriptProperties().getProperty('DATABASE_SPREADSHEET_ID');
  const sheetName = DB_CONFIG.SHEET_NAME;
  
  try {
    const data = batchGetSheetsData(service, dbId, [`'${sheetName}'!1:1`]);
    const headers = data.valueRanges[0].values[0] || [];
    
    return {
      columnCount: headers.length,
      headers: headers,
      is5ColumnOptimized: headers.length === 5
    };
  } catch (error) {
    console.error('スキーマ取得エラー:', error);
    return { columnCount: 0, headers: [], is5ColumnOptimized: false };
  }
}

/**
 * 移行用の全ユーザーデータ取得（従来形式）
 */
function getAllUsersForMigration() {
  const service = getServiceAccountTokenCached();
  const dbId = PropertiesService.getScriptProperties().getProperty('DATABASE_SPREADSHEET_ID');
  const sheetName = DB_CONFIG.SHEET_NAME;
  
  try {
    const data = batchGetSheetsData(service, dbId, [`'${sheetName}'!A:Z`]);
    const values = data.valueRanges[0].values || [];
    
    if (values.length < 2) return [];
    
    const headers = values[0];
    const dataRows = values.slice(1);
    
    return dataRows.map(row => {
      const user = {};
      headers.forEach((header, index) => {
        user[header] = row[index] || '';
      });
      return user;
    });
  } catch (error) {
    console.error('ユーザーデータ取得エラー:', error);
    return [];
  }
}

/**
 * データベースバックアップ作成
 */
function createDatabaseBackup() {
  try {
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const backupSheetName = `Users_Backup_${timestamp}`;
    
    const service = getServiceAccountTokenCached();
    const dbId = PropertiesService.getScriptProperties().getProperty('DATABASE_SPREADSHEET_ID');
    
    // 現在のUsersシートをバックアップシートとしてコピー
    const spreadsheet = SpreadsheetApp.openById(dbId);
    const usersSheet = spreadsheet.getSheetByName('Users');
    const backupSheet = usersSheet.copyTo(spreadsheet);
    backupSheet.setName(backupSheetName);
    
    console.log(`💾 バックアップ作成完了: ${backupSheetName}`);
    
    return {
      success: true,
      backupSheetName: backupSheetName
    };
  } catch (error) {
    console.error('バックアップ作成エラー:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 5項目構造でデータベースを再構築
 */
function rebuildDatabaseWith5Columns(convertedUsers) {
  try {
    const service = getServiceAccountTokenCached();
    const dbId = PropertiesService.getScriptProperties().getProperty('DATABASE_SPREADSHEET_ID');
    const sheetName = DB_CONFIG.SHEET_NAME;
    
    // 新しいヘッダー行
    const newHeaders = DB_CONFIG.HEADERS;
    
    // データ行を5項目形式に変換
    const newDataRows = convertedUsers.map(user => [
      user.userId,
      user.userEmail,
      user.isActive,
      user.configJson,
      user.lastModified
    ]);
    
    // ヘッダー + データ行
    const newSheetData = [newHeaders, ...newDataRows];
    
    // スプレッドシートを更新
    const updateResult = updateSheetData(service, dbId, sheetName, newSheetData);
    
    if (updateResult) {
      console.log(`🎉 5項目データベース再構築完了: ${newDataRows.length}行`);
      return {
        success: true,
        newStructure: {
          columns: newHeaders.length,
          rows: newDataRows.length
        }
      };
    } else {
      throw new Error('スプレッドシート更新に失敗しました');
    }
    
  } catch (error) {
    console.error('データベース再構築エラー:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 🧪 移行テスト（安全確認）
 */
function testConfigJsonCentricMigration() {
  console.log('🧪 configJSON中心型移行テスト開始');
  
  // テストデータ作成
  const testUser = {
    userId: 'test-user-123',
    userEmail: 'test@example.com',
    createdAt: '2025-01-01T00:00:00Z',
    lastAccessedAt: '2025-01-02T12:00:00Z',
    isActive: 'TRUE',
    spreadsheetId: 'test-sheet-456',
    sheetName: 'テストシート',
    configJson: JSON.stringify({
      setupStatus: 'completed',
      appPublished: true,
      displaySettings: { showNames: true }
    }),
    lastModified: '2025-01-02T08:00:00Z'
  };
  
  // 変換テスト
  const converted = convertUserToConfigJsonCentric(testUser);
  
  // 検証
  const configData = JSON.parse(converted.configJson);
  
  const testResults = {
    hasUserId: !!converted.userId,
    hasUserEmail: !!converted.userEmail,
    hasIsActive: converted.isActive !== undefined,
    hasConfigJson: !!converted.configJson,
    hasLastModified: !!converted.lastModified,
    
    // configJson内容検証
    configHasSpreadsheetId: !!configData.spreadsheetId,
    configHasSheetName: !!configData.sheetName,
    configHasCreatedAt: !!configData.createdAt,
    configHasMigrationInfo: !!configData.migratedAt,
    configHasSchemaVersion: configData.schemaVersion === '5-column-optimized'
  };
  
  const allPassed = Object.values(testResults).every(result => result === true);
  
  console.log('🧪 移行テスト結果:', testResults);
  console.log(allPassed ? '✅ テスト合格' : '❌ テスト失敗');
  
  return {
    success: allPassed,
    testResults: testResults,
    convertedSample: converted
  };
}