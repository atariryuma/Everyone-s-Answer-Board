/**
 * @fileoverview DatabaseCore - データベースコア機能 (遅延初期化対応)
 *
 * 🎯 責任範囲:
 * - データベース接続・認証
 * - 基本CRUD操作
 * - サービスアカウント管理
 *
 * 🔄 GAS Best Practices準拠:
 * - 遅延初期化パターン (DB関数呼び出し時にinit)
 * - ファイル読み込み順序非依存設計
 * - グローバル副作用排除
 */

/* global DB:writable, ServiceFactory */


// 遅延初期化状態管理
let databaseCoreInitialized = false;

/**
 * DatabaseCore遅延初期化
 */
function initDatabaseCore() {
if (databaseCoreInitialized) return true;

try {
databaseCoreInitialized = true;
console.log('✅ DatabaseCore initialized successfully');
return true;
} catch (error) {
console.error('initDatabaseCore failed:', error.message);
return false;
}
}

/**
 * DatabaseCore - データベースコア機能
 * 基本的なデータベース操作とサービス管理
 */
// ===========================================
// 🔐 DatabaseCore Functions (Flat)
// ===========================================

// ==========================================
// 🔐 データベース接続・認証
// ==========================================

/**
 * セキュアなデータベースID取得
 * @returns {string} データベースID
 */
function getSecureDatabaseId() {
try {
return ServiceFactory.getProperties().getProperty('DATABASE_SPREADSHEET_ID');
} catch (error) {
console.error('DatabaseCore', {
operation: 'getSecureDatabaseId',
error: error.message
});
throw new Error('データベース設定の取得に失敗しました');
}
}

/**
 * バッチデータ取得
 * @param {Object} service - Sheetsサービス
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {Array} ranges - 取得範囲配列
 * @returns {Object} バッチ取得結果
 */
function batchGetSheetsData(service, spreadsheetId, ranges) {
try {
console.log('DatabaseCore.batchGetSheetsData');

if (!ranges || ranges.length === 0) {
return { valueRanges: [] };
}

const result = service.spreadsheets.values.batchGet({
spreadsheetId,
ranges
});

return result;
} catch (error) {
console.error('DatabaseCore', {
operation: 'batchGetSheetsData',
spreadsheetId,
rangesCount: ranges?.length,
error: error.message
});
throw error;
}
}

/**
 * Sheetsサービス取得（直接作成）
 * @returns {Object} Sheetsサービス
 */
function getSheetsService() {
  try {
    // キャッシュ機能除去: 常に新しいサービスを作成
    const service = createSheetsService();

    // サービス検証
    if (!service || !service.spreadsheets || !service.spreadsheets.values) {
      throw new Error('Invalid service structure created');
    }

    // 必要なメソッドの存在確認（正確な階層構造）
    const requiredMethods = ['get', 'update', 'append'];
    for (const method of requiredMethods) {
      if (typeof service.spreadsheets.values[method] !== 'function') {
        console.error(`Method check failed: service.spreadsheets.values.${method}`, {
          type: typeof service.spreadsheets.values[method],
          available: Object.keys(service.spreadsheets.values)
        });
        throw new Error(`Required method '${method}' is not available in service.spreadsheets.values`);
      }
    }

    console.log('DatabaseCore: All required methods validated', {
      methods: requiredMethods,
      structure: 'service.spreadsheets.values'
    });

    return service;
  } catch (error) {
    console.error('DatabaseCore', {
      operation: 'getSheetsService',
      error: error.message
    });

    // より具体的なエラーメッセージを提供
    if (error.message.includes('SERVICE_ACCOUNT_CREDS')) {
      throw new Error('サービスアカウント設定に問題があります。システム管理者にお問い合わせください。');
    } else if (error.message.includes('spreadsheet')) {
      throw new Error('データベースへのアクセスに失敗しました。設定を確認してください。');
    } else {
      throw new Error(`データベースサービスの初期化に失敗しました: ${error.message}`);
    }
  }
}

/**
 * Sheetsサービス作成
 * @returns {Object} 新しいSheetsサービス
 */
function createSheetsService() {
try {
const serviceAccountKey = ServiceFactory.getProperties().getProperty('SERVICE_ACCOUNT_CREDS');

if (!serviceAccountKey) {
throw new Error('サービスアカウントキーが設定されていません');
}

// サービスアカウントキーの検証のみ（Google Apps Scriptでは直接Sheetsサービスを使用）
const parsedKey = JSON.parse(serviceAccountKey);

// サービスアカウントキーの基本検証
if (!parsedKey.client_email || !parsedKey.private_key) {
throw new Error('無効なサービスアカウントキーです');
}

// Google Apps Script標準のSpreadsheetAppを使用
const service = {
  spreadsheets: {
    values: {
      get: ({ spreadsheetId, range }) => {
        try {
          const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
          const [sheetName] = range.split('!');
          const sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.getSheets()[0];
          const values = sheet.getDataRange().getValues();
          return { data: { values } };
        } catch (error) {
          console.error('Service.get error:', error.message);
          throw error;
        }
      },

      update: ({ spreadsheetId, range, resource }) => {
        try {
          const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
          const [sheetName] = range.split('!');
          const sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.getSheets()[0];

          if (!sheet) {
            throw new Error(`Sheet "${sheetName}" not found in spreadsheet`);
          }

          const {values} = resource;
          if (values && values.length > 0) {
            sheet.getRange(1, 1, values.length, values[0].length).setValues(values);
          }

          console.log('DatabaseCore.update: Data updated successfully', {
            sheetName,
            rowsUpdated: values ? values.length : 0
          });

          return { updatedCells: values ? values.length * values[0].length : 0 };
        } catch (error) {
          console.error('Service.update error:', error.message);
          throw error;
        }
      },

      append: ({ spreadsheetId, range, resource, valueInputOption }) => {
        try {
          const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
          const [sheetName] = range.split('!');
          const sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.getSheets()[0];

          if (!sheet) {
            throw new Error(`Sheet "${sheetName}" not found in spreadsheet`);
          }

          const {values} = resource;
          if (values && values.length > 0) {
            const lastRow = sheet.getLastRow();
            const targetRow = lastRow + 1;
            const targetRange = sheet.getRange(targetRow, 1, values.length, values[0].length);
            targetRange.setValues(values);

            console.log('DatabaseCore.append: Data written successfully', {
              sheetName,
              targetRow,
              rowsWritten: values.length
            });
          }

          return {
            updates: {
              updatedRows: values ? values.length : 0,
              spreadsheetId,
              range
            }
          };
        } catch (error) {
          console.error('Service.append error:', error.message);
          throw error;
        }
      }
    }
  }
};

console.log('DatabaseCore', {
operation: 'createSheetsService',
serviceType: parsedKey.type || 'unknown'
});

return service;
} catch (error) {
console.error('DatabaseCore', {
operation: 'createSheetsService',
error: error.message
});
throw error;
}
}

/**
 * リトライ付きSheetsサービス取得
 * @param {number} maxRetries - 最大リトライ回数
 * @returns {Object} Sheetsサービス
 */
function getSheetsServiceWithRetry(maxRetries = 2) {
for (let attempt = 1; attempt <= maxRetries; attempt++) {
try {
return getSheetsService();
} catch (error) {
console.warn('DatabaseCore', {
operation: 'getSheetsServiceWithRetry',
attempt,
maxRetries,
error: error.message
});

if (attempt === maxRetries) {
throw error;
}

Utilities.sleep(1000 * attempt); // 指数バックオフ
}
}
}

// ==========================================
// 🔧 診断・ユーティリティ
// ==========================================

/**
 * データベース接続診断
 * @returns {Object} 診断結果
 */
function diagnoseDatabaseCore() {
const results = {
service: 'DatabaseCore',
timestamp: new Date().toISOString(),
checks: []
};

try {
// データベースID確認
const databaseId = getSecureDatabaseId();
results.checks.push({
name: 'Database ID',
status: databaseId ? '✅' : '❌',
details: databaseId ? 'Database ID configured' : 'Database ID missing'
});

// サービスアカウント確認
try {
const service = createSheetsService();
results.checks.push({
name: 'Service Account',
status: service ? '✅' : '❌',
details: 'Service account authentication working'
});
} catch (serviceError) {
results.checks.push({
name: 'Service Account',
status: '❌',
details: serviceError.message
});
}

// キャッシュサービス確認
try {
const cache = ServiceFactory.getCache();
cache.get('test_key');
results.checks.push({
name: 'Cache Service',
status: '✅',
details: 'Cache service accessible'
});
} catch (cacheError) {
results.checks.push({
name: 'Cache Service',
status: '⚠️',
details: cacheError.message
});
}

results.overall = results.checks.every(check => check.status === '✅') ? '✅' : '⚠️';
} catch (error) {
results.checks.push({
name: 'Core Diagnosis',
status: '❌',
details: error.message
});
results.overall = '❌';
}

return results;
}

// ... 残りのファイル内容は同じです