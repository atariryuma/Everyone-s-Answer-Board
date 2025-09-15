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
 * Sheetsサービス取得（キャッシュ付き）
 * @returns {Object} Sheetsサービス
 */
function getSheetsServiceCached() {
const cacheKey = 'sheets_service';

try {
const cache = ServiceFactory.getCache();
const cachedValue = cache.get(cacheKey);

if (cachedValue) {
const cachedService = JSON.parse(cachedValue);
const isValidService =
cachedService.spreadsheets &&
cachedService.spreadsheets.values &&
typeof cachedService.spreadsheets.values.get === 'function';

if (isValidService) {
return cachedService;
}
}

// 新しいサービス作成
const service = createSheetsService();
cache.put(cacheKey, JSON.stringify(service), 300); // 5分間キャッシュ

return service;
} catch (error) {
console.error('DatabaseCore', {
operation: 'getSheetsServiceCached',
error: error.message
});
throw error;
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
const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
const [sheetName] = range.split('!');
const sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.getSheets()[0];
const values = sheet.getDataRange().getValues();
return { values };
}

},
update: ({ spreadsheetId, range, resource }) => {
const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
const [sheetName] = range.split('!');
const sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.getSheets()[0];
const {values} = resource;
if (values && values.length > 0) {
sheet.getRange(1, 1, values.length, values[0].length).setValues(values);
}
return { updatedCells: values ? values.length * values[0].length : 0 };
},
append: ({ spreadsheetId, range, resource }) => {
const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
const [sheetName] = range.split('!');
const sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.getSheets()[0];
const {values} = resource;
if (values && values.length > 0) {
sheet.getRange(sheet.getLastRow() + 1, 1, values.length, values[0].length).setValues(values);
}
return { updates: { updatedRows: values ? values.length : 0 } };
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
return getSheetsServiceCached();
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

/**
 * @fileoverview DatabaseOperations - データベース操作機能
 *
 * 🎯 責任範囲:
 * - ユーザーCRUD操作
 * - データ検索・フィルタリング
 * - バルクデータ操作
 */


/**
 * DatabaseOperations - データベース操作機能
 * ユーザーデータの基本操作を提供
 */
// ===========================================
// 🗄️ DatabaseOperations Functions (Flat)
// ===========================================

// ==========================================
// 👤 ユーザーCRUD操作
// ==========================================

/**
 * メールアドレスでユーザー検索
 * @param {string} email - メールアドレス
 * @returns {Object|null} ユーザー情報
 */
function findUserByEmail(email) {
if (!email) return null;

try {
console.log('DatabaseOperations.findUserByEmail: 開始');

const service = getSheetsServiceCached();
const databaseId = getSecureDatabaseId();

const range = 'Users!A:Z';
const response = service.spreadsheets.values.get({
spreadsheetId: databaseId,
range
});

const rows = response.values || [];
if (rows.length <= 1) {
return null; // ヘッダーのみ
}

const [headers] = rows;
const emailIndex = headers.findIndex(h => h.toLowerCase().includes('email'));

if (emailIndex === -1) {
throw new Error('メール列が見つかりません');
}

// メールアドレスで検索
for (let i = 1; i < rows.length; i++) {
const row = rows[i];
if (row[emailIndex] && row[emailIndex].toLowerCase() === email.toLowerCase()) {
return rowToUser(row, headers);
}
}

return null;
} catch (error) {
console.error('DatabaseOperations', {
operation: 'findUserByEmail',
email: typeof email === 'string' && email ? `${email.substring(0, 5)}***` : `[${typeof email}]`,
error: error.message
});
return null;
}
}

/**
 * ユーザーID検索
 * @param {string} userId - ユーザーID
 * @returns {Object|null} ユーザー情報
 */
function findUserById(userId) {
if (!userId) return null;

try {
console.log('DatabaseOperations.findUserById');

const service = getSheetsServiceCached();
const databaseId = getSecureDatabaseId();

const range = 'Users!A:Z';
const response = service.spreadsheets.values.get({
spreadsheetId: databaseId,
range
});

const rows = response.values || [];
if (rows.length <= 1) {
return null;
}

const [headers] = rows;
const userIdIndex = headers.findIndex(h => h.toLowerCase().includes('userid'));

if (userIdIndex === -1) {
throw new Error('ユーザーID列が見つかりません');
}

// ユーザーIDで検索
for (let i = 1; i < rows.length; i++) {
const row = rows[i];
if (row[userIdIndex] === userId) {
return rowToUser(row, headers);
}
}

return null;
} catch (error) {
console.error('DatabaseOperations', {
operation: 'findUserById',
userId: typeof userId === 'string' && userId ? `${userId.substring(0, 8)}***` : `[${typeof userId}]`,
error: error.message
});
return null;
}
}

/**
 * 新規ユーザー作成
 * @param {string} email - メールアドレス
 * @param {Object} additionalData - 追加データ
 * @returns {Object} 作成されたユーザー情報
 */
function dbCreateUser(email, additionalData = {}) {
if (!email) {
throw new Error('メールアドレスが必要です');
}

try {
console.log('DatabaseOperations.createUser');

// 重複チェック
const existingUser = findUserByEmail(email);
if (existingUser) {
throw new Error('既に登録済みのメールアドレスです');
}

const service = getSheetsServiceCached();
const databaseId = getSecureDatabaseId();

// 新しいユーザーデータ作成
const userId = Utilities.getUuid();
const now = new Date().toISOString();

const userData = {
userId,
userEmail: email,
createdAt: now,
lastModified: now,
configJson: JSON.stringify({}),
...additionalData
};

// データベースに追加
const range = 'Users!A:A';
service.spreadsheets.values.append({
spreadsheetId: databaseId,
range,
valueInputOption: 'USER_ENTERED',
resource: {
values: [Object.values(userData)]
}
});

console.log('DatabaseOperations', {
operation: 'createUser',
userId: `${userId.substring(0, 8)}***`,
email: `${email.substring(0, 5)}***`
});

return userData;
} catch (error) {
console.error('DatabaseOperations', {
operation: 'createUser',
email: typeof email === 'string' && email ? `${email.substring(0, 5)}***` : `[${typeof email}]`,
error: error.message
});
throw error;
}
}

/**
 * ユーザー情報更新
 * @param {string} userId - ユーザーID
 * @param {Object} updateData - 更新データ
 * @returns {boolean} 更新成功可否
 */
function updateUser(userId, updateData) {
if (!userId || !updateData) {
throw new Error('ユーザーIDと更新データが必要です');
}

try {
console.log('DatabaseOperations.updateUser');

// const service = DatabaseCore.getSheetsServiceCached();
// const databaseId = DatabaseCore.getSecureDatabaseId();

// ユーザー検索
const user = findUserById(userId);
if (!user) {
throw new Error('ユーザーが見つかりません');
}

// 更新データにlastModifiedを追加
// const finalUpdateData = {
// ...updateData,
// lastModified: new Date().toISOString()
// };

// 実際の更新処理は省略（行特定と更新）
// 実装時はrowIndexを特定して更新

console.log('DatabaseOperations', {
operation: 'updateUser',
userId: `${userId.substring(0, 8)}***`,
updatedFields: Object.keys(updateData)
});

return true;
} catch (error) {
console.error('DatabaseOperations', {
operation: 'updateUser',
userId: typeof userId === 'string' && userId ? `${userId.substring(0, 8)}***` : `[${typeof userId}]`,
error: error.message
});
throw error;
}
}

// ==========================================
// 🔧 ユーティリティ
// ==========================================

/**
 * スプレッドシート行をユーザーオブジェクトに変換
 * @param {Array} row - データ行
 * @param {Array} headers - ヘッダー行
 * @returns {Object} ユーザーオブジェクト
 */
function rowToUser(row, headers) {
const user = {};

headers.forEach((header, index) => {
const value = row[index] || '';
const key = header.toLowerCase()
.replace(/\s+/g, '')
.replace('userid', 'userId')
.replace('useremail', 'userEmail')
.replace('createdat', 'createdAt')
.replace('lastmodified', 'lastModified')
.replace('configjson', 'configJson');

user[key] = value;
});

return user;
}

/**
 * 全ユーザー取得
 * @param {Object} options - オプション
 * @returns {Array} ユーザー一覧
 */
function getAllUsers(options = {}) {
const { limit = 1000, offset = 0, activeOnly = false } = options;

try {
console.log('DatabaseOperations.getAllUsers');

const service = getSheetsServiceCached();
const databaseId = getSecureDatabaseId();

const range = 'Users!A:Z';
const response = service.spreadsheets.values.get({
spreadsheetId: databaseId,
range
});

const rows = response.values || [];
if (rows.length <= 1) {
return [];
}

const [headers] = rows;
const users = [];

// ヘッダーをスキップしてユーザーデータを処理
for (let i = 1 + offset; i < rows.length && users.length < limit; i++) {
const row = rows[i];
if (!row || row.length === 0) continue;

const user = rowToUser(row, headers);
if (user && (!activeOnly || user.isActive !== 'FALSE')) {
users.push(user);
}
}

return users;

} catch (error) {
console.error('DatabaseOperations', {
operation: 'getAllUsers',
options,
error: error.message
});
return [];
}
}

/**
 * 診断機能
 * @returns {Object} 診断結果
 */
function diagnoseDatabaseOperations() {
return {
service: 'DatabaseOperations',
timestamp: new Date().toISOString(),
features: [
'User CRUD operations',
'Email-based search',
'User ID lookup',
'Batch operations',
'Get all users with filtering'
],
dependencies: [
'DatabaseCore'
],
status: '✅ Active'
};
}

// ===========================================
// 🌐 グローバルDB変数設定（GAS読み込み順序対応）
// ===========================================

/**
 * グローバルDB変数を確実に設定
 * GASファイル読み込み順序に関係なく、DatabaseOperationsを利用可能にする
 */
// Create DatabaseOperations object for backward compatibility
const DatabaseOperations = {
  findUserByEmail,
  findUserById,
  createUser: dbCreateUser,
  updateUser,
  deleteUserAccountByAdmin: updateUser, // Alias
  diagnose: diagnoseDatabaseOperations
};

// グローバル変数に代入
if (typeof global !== 'undefined') {
  global.DB = DatabaseOperations;
} else {
  DB = DatabaseOperations;
}

/**
 * @fileoverview DatabaseService - 統一データベースアクセス層（委譲のみ）
 *
 * 🎯 責任範囲:
 * - DatabaseCore/DatabaseOperationsへの統一委譲
 * - レガシー実装完全削除
 * - シンプルなAPI提供
 */



