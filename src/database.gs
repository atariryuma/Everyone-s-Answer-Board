/**
 * @fileoverview データベース管理 - バッチ操作とキャッシュ最適化
 * GAS互換の関数ベースの実装
 */

// データベース管理のための定数
var USER_CACHE_TTL = 300; // 5分
var DB_BATCH_SIZE = 100;

/**
 * 削除ログ用シート設定
 */
const DELETE_LOG_SHEET_CONFIG = {
  SHEET_NAME: 'DeleteLogs',
  HEADERS: ['timestamp', 'executorEmail', 'targetUserId', 'targetEmail', 'reason', 'deleteType']
};

/**
 * 削除ログを記録
 * @param {string} executorEmail - 削除実行者のメール
 * @param {string} targetUserId - 削除対象のユーザーID
 * @param {string} targetEmail - 削除対象のメール
 * @param {string} reason - 削除理由
 * @param {string} deleteType - 削除タイプ ("self" | "admin")
 */
function logAccountDeletion(executorEmail, targetUserId, targetEmail, reason, deleteType) {
  try {
    const props = PropertiesService.getScriptProperties();
    const dbId = props.getProperty('DATABASE_SPREADSHEET_ID');
    
    if (!dbId) {
      console.warn('削除ログの記録をスキップします: データベースIDが設定されていません');
      return;
    }
    
    const service = getSheetsService();
    const logSheetName = DELETE_LOG_SHEET_CONFIG.SHEET_NAME;
    
    // ログシートの存在確認・作成
    try {
      const spreadsheetInfo = getSpreadsheetsData(service, dbId);
      const logSheetExists = spreadsheetInfo.sheets.some(sheet => 
        sheet.properties.title === logSheetName
      );
      
      if (!logSheetExists) {
        // ログシートを作成
        const addSheetRequest = {
          addSheet: {
            properties: {
              title: logSheetName
            }
          }
        };
        
        batchUpdateSpreadsheet(service, dbId, {
          requests: [addSheetRequest]
        });
        
        // ヘッダー行を追加
        appendSheetsData(service, dbId, `'${logSheetName}'!A1`, [DELETE_LOG_SHEET_CONFIG.HEADERS]);
      }
    } catch (sheetError) {
      console.warn('ログシートの確認に失敗しましたが、処理を続行します:', sheetError.message);
    }
    
    // ログエントリを追加
    const logEntry = [
      new Date().toISOString(),
      executorEmail,
      targetUserId,
      targetEmail,
      reason || '',
      deleteType
    ];
    
    appendSheetsData(service, dbId, `'${logSheetName}'!A:F`, [logEntry]);
    console.log('✅ 削除ログを記録しました:', logEntry);
    
  } catch (error) {
    console.error('削除ログの記録に失敗しました:', error.message);
    // ログ記録の失敗は削除処理自体を止めない
  }
}

/**
 * 全ユーザー一覧を取得（管理者用）
 */
function getAllUsersForAdmin() {
  try {
    // 管理者権限チェック
    if (!isDeployUser()) {
      throw new Error('この機能にアクセスする権限がありません。');
    }
    
    const props = PropertiesService.getScriptProperties();
    const dbId = props.getProperty('DATABASE_SPREADSHEET_ID');
    
    if (!dbId) {
      throw new Error('データベースIDが設定されていません');
    }
    
    const service = getSheetsService();
    const sheetName = DB_SHEET_CONFIG.SHEET_NAME;
    
    const data = batchGetSheetsData(service, dbId, [`'${sheetName}'!A:I`]);
    const values = data.valueRanges[0].values || [];
    
    if (values.length <= 1) {
      return []; // ヘッダーのみの場合
    }
    
    const headers = values[0];
    const users = [];
    
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const user = {};
      
      headers.forEach((header, index) => {
        user[header] = row[index] || '';
      });
      
      // 設定情報をパース
      try {
        user.configJson = JSON.parse(user.configJson || '{}');
      } catch (e) {
        user.configJson = {};
      }
      
      users.push(user);
    }
    
    console.log(`✅ 管理者用ユーザー一覧を取得: ${users.length}件`);
    return users;
    
  } catch (error) {
    console.error('getAllUsersForAdmin error:', error.message);
    throw new Error('ユーザー一覧の取得に失敗しました: ' + error.message);
  }
}

/**
 * 管理者による他ユーザーアカウント削除
 * @param {string} targetUserId - 削除対象のユーザーID
 * @param {string} reason - 削除理由
 */
function deleteUserAccountByAdmin(targetUserId, reason) {
  try {
    // 管理者権限チェック
    if (!isDeployUser()) {
      throw new Error('この機能にアクセスする権限がありません。');
    }
    
    // 削除理由の検証
    if (!reason || reason.trim().length < 20) {
      throw new Error('削除理由は20文字以上で入力してください。');
    }
    
    const executorEmail = Session.getActiveUser().getEmail();
    
    // 削除対象ユーザー情報を取得
    const targetUserInfo = findUserById(targetUserId);
    if (!targetUserInfo) {
      throw new Error('削除対象のユーザーが見つかりません。');
    }
    
    // 自分自身の削除を防ぐ
    if (targetUserInfo.adminEmail === executorEmail) {
      throw new Error('自分自身のアカウントは管理者削除機能では削除できません。個人用削除機能をご利用ください。');
    }
    
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    
    try {
      // データベースからユーザー行を削除
      const props = PropertiesService.getScriptProperties();
      const dbId = props.getProperty('DATABASE_SPREADSHEET_ID');
      
      if (!dbId) {
        throw new Error('データベースIDが設定されていません');
      }
      
      const service = getSheetsService();
      const sheetName = DB_SHEET_CONFIG.SHEET_NAME;
      
      // データベーススプレッドシートの情報を取得
      const spreadsheetInfo = getSpreadsheetsData(service, dbId);
      const targetSheetId = spreadsheetInfo.sheets.find(sheet => 
        sheet.properties.title === sheetName
      )?.properties.sheetId;
      
      if (targetSheetId === null || targetSheetId === undefined) {
        throw new Error(`データベースシート「${sheetName}」が見つかりません`);
      }
      
      // データを取得してユーザー行を特定
      const data = batchGetSheetsData(service, dbId, [`'${sheetName}'!A:I`]);
      const values = data.valueRanges[0].values || [];
      
      let rowToDelete = -1;
      for (let i = values.length - 1; i >= 1; i--) {
        if (values[i][0] === targetUserId) {
          rowToDelete = i + 1; // スプレッドシートは1ベース
          break;
        }
      }
      
      if (rowToDelete !== -1) {
        const deleteRequest = {
          deleteDimension: {
            range: {
              sheetId: targetSheetId,
              dimension: 'ROWS',
              startIndex: rowToDelete - 1, // APIは0ベース
              endIndex: rowToDelete
            }
          }
        };
        
        batchUpdateSpreadsheet(service, dbId, {
          requests: [deleteRequest]
        });
        
        console.log(`✅ 管理者削除完了: row ${rowToDelete}, sheetId ${targetSheetId}`);
      } else {
        throw new Error('削除対象のユーザー行が見つかりませんでした');
      }
      
      // 削除ログを記録
      logAccountDeletion(
        executorEmail,
        targetUserId,
        targetUserInfo.adminEmail,
        reason.trim(),
        'admin'
      );
      
      // 関連キャッシュを削除
      invalidateUserCache(targetUserId, targetUserInfo.adminEmail, targetUserInfo.spreadsheetId, false);
      
    } finally {
      lock.releaseLock();
    }
    
    const successMessage = `管理者によりアカウント「${targetUserInfo.adminEmail}」が削除されました。\n削除理由: ${reason.trim()}`;
    console.log(successMessage);
    return {
      success: true,
      message: successMessage,
      deletedUser: {
        userId: targetUserId,
        email: targetUserInfo.adminEmail
      }
    };
    
  } catch (error) {
    console.error('deleteUserAccountByAdmin error:', error.message);
    throw new Error('管理者によるアカウント削除に失敗しました: ' + error.message);
  }
}

/**
 * 削除権限チェック
 * @param {string} targetUserId - 対象ユーザーID
 */
function canDeleteUser(targetUserId) {
  try {
    const currentUserEmail = Session.getActiveUser().getEmail();
    const targetUser = findUserById(targetUserId);
    
    if (!targetUser) {
      return false;
    }
    
    // 本人削除 OR 管理者削除
    return (currentUserEmail === targetUser.adminEmail) || isDeployUser();
  } catch (error) {
    console.error('canDeleteUser error:', error.message);
    return false;
  }
}

/**
 * 削除ログ一覧を取得（管理者用）
 */
function getDeletionLogs() {
  try {
    // 管理者権限チェック
    if (!isDeployUser()) {
      throw new Error('この機能にアクセスする権限がありません。');
    }
    
    const props = PropertiesService.getScriptProperties();
    const dbId = props.getProperty('DATABASE_SPREADSHEET_ID');
    
    if (!dbId) {
      throw new Error('データベースIDが設定されていません');
    }
    
    const service = getSheetsService();
    const logSheetName = DELETE_LOG_SHEET_CONFIG.SHEET_NAME;
    
    try {
      // ログシートの存在確認
      const spreadsheetInfo = getSpreadsheetsData(service, dbId);
      const logSheetExists = spreadsheetInfo.sheets.some(sheet => 
        sheet.properties.title === logSheetName
      );
      
      if (!logSheetExists) {
        console.log('削除ログシートが存在しません');
        return []; // ログシートがない場合は空の配列を返す
      }
      
      // ログデータを取得
      const data = batchGetSheetsData(service, dbId, [`'${logSheetName}'!A:F`]);
      const values = data.valueRanges[0].values || [];
      
      if (values.length <= 1) {
        return []; // ヘッダーのみまたは空の場合
      }
      
      const headers = values[0];
      const logs = [];
      
      // データをオブジェクト形式に変換（最新順に並び替え）
      for (let i = values.length - 1; i >= 1; i--) {
        const row = values[i];
        const log = {};
        
        headers.forEach((header, index) => {
          log[header] = row[index] || '';
        });
        
        logs.push(log);
      }
      
      console.log(`✅ 削除ログを取得: ${logs.length}件`);
      return logs;
      
    } catch (sheetError) {
      console.warn('ログシートの読み取りに失敗:', sheetError.message);
      return []; // シートアクセスエラーの場合は空を返す
    }
    
  } catch (error) {
    console.error('getDeletionLogs error:', error.message);
    throw new Error('削除ログの取得に失敗しました: ' + error.message);
  }
}

/**
 * 最適化されたSheetsサービスを取得
 * データベース操作専用：サービスアカウントのみ使用
 * @returns {object} Sheets APIサービス
 */
function getSheetsService() {
  try {
    var accessToken = getServiceAccountTokenCached();
    if (!accessToken) {
      throw new Error('サービスアカウントトークンの取得に失敗しました。データベースにアクセスできません。');
    }
    return createSheetsService(accessToken);
  } catch (error) {
    console.error('getSheetsService error:', error.message);
    throw new Error('データベースアクセス用サービスアカウントの認証に失敗しました: ' + error.message);
  }
}


/**
 * ユーザー情報を効率的に検索（統合キャッシュ優先）
 * @param {string} userId - ユーザーID
 * @param {boolean} bypassCache - キャッシュを無視してDB直接検索
 * @returns {object|null} ユーザー情報
 */
function findUserById(userId, bypassCache = false) {
  console.log('🔍 findUserById: 検索開始', { userId, bypassCache });
  
  // 循環依存防止: 再帰深度の追跡
  if (!findUserById._recursionDepth) {
    findUserById._recursionDepth = 0;
  }
  
  if (findUserById._recursionDepth > 2) {
    console.warn('🚨 findUserById: 循環依存を検出、直接DB検索にフォールバック');
    findUserById._recursionDepth = 0;
    return fetchUserFromDatabase('userId', userId);
  }
  
  findUserById._recursionDepth++;
  
  try {
    // Core.gsの統合キャッシュシステムが利用可能な場合はそれを使用
    if (typeof getCachedUserInfoUnified === 'function') {
      console.log('🔍 統合キャッシュシステム使用');
      const result = getCachedUserInfoUnified(userId, bypassCache);
      findUserById._recursionDepth--;
      return result;
    }
  } catch (error) {
    findUserById._recursionDepth--;
    throw error;
  }
  
  // フォールバック: 従来のキャッシュシステム
  console.log('🔍 従来キャッシュシステム使用');
  var cacheKey = 'user_' + userId;
  const result = cacheManager.get(
    cacheKey,
    function() { 
      console.log('🔍 findUserById: キャッシュミス、データベース検索', { userId });
      const dbResult = fetchUserFromDatabase('userId', userId);
      console.log('🔍 findUserById: データベース検索結果', { 
        userId, 
        found: !!dbResult, 
        adminEmail: dbResult?.adminEmail 
      });
      return dbResult;
    },
    { ttl: 300, enableMemoization: true }
  );
  console.log('🔍 findUserById: 最終結果', { 
    userId, 
    found: !!result, 
    adminEmail: result?.adminEmail 
  });
  findUserById._recursionDepth--;
  return result;
}

/**
 * ロック競合を避けるための軽量ユーザー検索
 * 登録処理中にhandleDirectExecAccessで使用
 * サービスアカウント専用でデータベースにアクセス
 * @param {string} email メールアドレス
 * @returns {object|null} ユーザー情報またはnull
 */
function findUserByEmailNonBlocking(email) {
  try {
    console.log('🔍 findUserByEmailNonBlocking: 検索開始', { email });
    if (!email) {
      console.log('🔍 findUserByEmailNonBlocking: メールアドレスが空');
      return null;
    }
    
    // 軽量検索でもサービスアカウント経由でアクセス
    const result = fetchUserFromDatabase('adminEmail', email);
    console.log('🔍 findUserByEmailNonBlocking: 検索結果', { 
      email, 
      found: !!result, 
      userId: result?.userId 
    });
    return result;
  } catch (error) {
    console.error('findUserByEmailNonBlocking error:', error);
    return null;
  }
}

/**
 * メールアドレスでユーザー検索
 * @param {string} email - メールアドレス
 * @returns {object|null} ユーザー情報
 */
function findUserByEmail(email) {
  var cacheKey = 'email_' + email;
  return cacheManager.get(
    cacheKey,
    function() { return fetchUserFromDatabase('adminEmail', email); },
    { ttl: 300, enableMemoization: true }
  );
}

/**
 * データベースからユーザーを取得（インデックス最適化版）
 * サービスアカウント専用でデータベースにアクセス
 * @param {string} field - 検索フィールド
 * @param {string} value - 検索値
 * @returns {object|null} ユーザー情報
 */
function fetchUserFromDatabase(field, value) {
  try {
    var props = PropertiesService.getScriptProperties();
    var dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
    var sheetName = DB_SHEET_CONFIG.SHEET_NAME;
    
    // インデックスキャッシュを利用
    var indexCacheKey = 'db_index_' + field;
    var userData = cacheManager.get(indexCacheKey, function() {
      return buildDatabaseIndex(field);
    }, { ttl: 600, enableMemoization: true }); // 10分間キャッシュ
    
    // インデックスから直接検索
    var targetUser = userData[value];
    if (targetUser) {
      console.log('fetchUserFromDatabase (Indexed) - Found user:', {
        field: field,
        value: value,
        userId: targetUser.userId,
        adminEmail: targetUser.adminEmail
      });
      return targetUser;
    }
    
    return null;
  } catch (error) {
    console.error('ユーザー検索エラー (' + field + ':' + value + '):', error);
    // フォールバック: 従来の線形検索
    console.warn('インデックス検索失敗、線形検索にフォールバック');
    try {
      return fetchUserFromDatabaseLinear(field, value);
    } catch (fallbackError) {
      const processedError = processError(fallbackError, {
        function: 'fetchUserFromDatabase',
        field: field,
        value: value,
        operation: 'user_search'
      });
      throw new Error(processedError.userMessage);
    }
  }
}

/**
 * データベースインデックスを構築
 * @param {string} field - インデックス対象フィールド
 * @returns {object} フィールド値をキーとするユーザーオブジェクト
 */
function buildDatabaseIndex(field) {
  var props = PropertiesService.getScriptProperties();
  var dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
  var sheetName = DB_SHEET_CONFIG.SHEET_NAME;
  
  var service = getSheetsService();
  var data = batchGetSheetsData(service, dbId, ["'" + sheetName + "'!A:I"]);
  var values = data.valueRanges[0].values || [];
  
  var index = {};
  if (values.length === 0) return index;
  
  var headers = values[0];
  var fieldIndex = headers.indexOf(field);
  
  if (fieldIndex === -1) return index;
  
  // インデックス構築（O(n)の一回限り）
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var key = row[fieldIndex];
    if (key) {
      var user = {};
      headers.forEach(function(header, index) {
        user[header] = row[index] || '';
      });
      index[key] = user;
    }
  }
  
  console.log('データベースインデックス構築完了:', { field: field, entries: Object.keys(index).length });
  return index;
}

/**
 * 従来の線形検索（フォールバック用）
 * @param {string} field - 検索フィールド
 * @param {string} value - 検索値
 * @returns {object|null} ユーザー情報
 */
function fetchUserFromDatabaseLinear(field, value) {
  try {
    var props = PropertiesService.getScriptProperties();
    var dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
    var sheetName = DB_SHEET_CONFIG.SHEET_NAME;
    
    var service = getSheetsService();
    var data = batchGetSheetsData(service, dbId, ["'" + sheetName + "'!A:I"]);
    var values = data.valueRanges[0].values || [];
    
    if (values.length === 0) return null;
    
    var headers = values[0];
    var fieldIndex = headers.indexOf(field);
    
    if (fieldIndex === -1) return null;
    
    for (var i = 1; i < values.length; i++) {
      if (values[i][fieldIndex] === value) {
        var user = {};
        headers.forEach(function(header, index) {
          user[header] = values[i][index] || '';
        });
        
        console.log('fetchUserFromDatabaseLinear - Found user:', {
          field: field,
          value: value,
          userId: user.userId,
          adminEmail: user.adminEmail
        });
        
        return user;
      }
    }
    
    return null;
  } catch (error) {
    console.error('線形検索エラー (' + field + ':' + value + '):', error);
    const processedError = handleSheetsApiError(error, {
      function: 'fetchUserFromDatabaseLinear',
      field: field,
      value: value,
      operation: 'linear_search'
    });
    throw new Error(processedError.userMessage);
  }
}

/**
 * データの整合性検証用のシグネチャを生成
 * @param {array} values - スプレッドシートの値配列
 * @param {array} headers - ヘッダー行
 * @returns {string} データシグネチャ
 */
function generateDataSignature(values, headers) {
  try {
    // データ構造の基本的な特徴をハッシュ化
    var signature = {
      headerCount: headers ? headers.length : 0,
      rowCount: values ? values.length : 0,
      firstRowLength: (values && values.length > 1 && values[1]) ? values[1].length : 0,
      headerHash: headers ? headers.join('|').substring(0, 50) : '',
      timestamp: Date.now()
    };
    return JSON.stringify(signature);
  } catch (e) {
    console.warn('generateDataSignature: シグネチャ生成エラー', { error: e.message });
    return JSON.stringify({ error: true, timestamp: Date.now() });
  }
}

/**
 * 行番号インデックスを構築
 * @param {array} values - シートのデータ行
 * @param {array} headers - ヘッダー行
 * @param {string} field - インデックス対象フィールド
 * @returns {object} フィールド値をキーとする行番号マップ
 */
function buildRowIndexForField(values, headers, field) {
  var index = {};
  var fieldIndex = headers.indexOf(field);
  
  // データ整合性チェック
  if (!values || !Array.isArray(values) || values.length === 0) {
    console.warn('buildRowIndexForField: 無効なvalues配列', { values: values });
    return { index: {}, metadata: { valid: false, reason: 'invalid_values' } };
  }
  
  if (!headers || !Array.isArray(headers)) {
    console.warn('buildRowIndexForField: 無効なheaders配列', { headers: headers });
    return { index: {}, metadata: { valid: false, reason: 'invalid_headers' } };
  }
  
  console.log('buildRowIndexForField: 開始', { 
    field: field, 
    fieldIndex: fieldIndex,
    valuesLength: values.length,
    headers: headers 
  });
  
  if (fieldIndex === -1) {
    console.warn('buildRowIndexForField: フィールドが見つかりません', { field: field, headers: headers });
    return { index: {}, metadata: { valid: false, reason: 'field_not_found' } };
  }
  
  // データハッシュを生成してキャッシュ有効性を検証
  var dataSignature = generateDataSignature(values, headers);
  var validEntries = 0;
  var skippedEntries = 0;
  
  // 行番号インデックス構築
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (!row || !Array.isArray(row)) {
      console.warn('buildRowIndexForField: 不正な行データ', { rowIndex: i, row: row });
      skippedEntries++;
      continue;
    }
    
    // 行の整合性をチェック
    if (i + 1 > values.length) {
      console.error('buildRowIndexForField: 行番号が範囲外', { i: i, valuesLength: values.length });
      skippedEntries++;
      continue;
    }
    
    var key = row[fieldIndex];
    if (key) {
      // 重複キーの検証
      if (index[key]) {
        console.warn('buildRowIndexForField: 重複キー検出', { key: key, existingIndex: index[key], newIndex: i + 1 });
      }
      index[key] = { 
        rowIndex: i + 1, // 1-based index
        arrayIndex: i,   // 0-based index for validation
        lastValidated: new Date().toISOString()
      };
      validEntries++;
      console.log('buildRowIndexForField: エントリー追加', { key: key, rowIndex: i + 1 });
    }
  }
  
  console.log('行インデックス構築完了:', { 
    field: field, 
    entries: Object.keys(index).length,
    validEntries: validEntries,
    skippedEntries: skippedEntries,
    indexKeys: Object.keys(index).slice(0, 5) // 最初の5つのキーを表示
  });
  
  return {
    index: index,
    metadata: {
      valid: true,
      dataSignature: dataSignature,
      fieldIndex: fieldIndex,
      totalRows: values.length,
      validEntries: validEntries,
      skippedEntries: skippedEntries,
      createdAt: new Date().toISOString()
    }
  };
}

/**
 * キャッシュの整合性を検証し、必要に応じて修復する
 * @param {string} userId - 検証対象のユーザーID
 * @returns {object} 検証結果
 */
function validateAndRepairCache(userId) {
  try {
    console.log('🔍 [CACHE_REPAIR] キャッシュ整合性検証開始:', userId);
    
    var repairActions = [];
    var issues = [];
    
    // 1. データベースから最新データを取得
    var props = PropertiesService.getScriptProperties();
    var dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
    var service = getSheetsService();
    var sheetName = DB_SHEET_CONFIG.SHEET_NAME;
    
    var data = batchGetSheetsData(service, dbId, ["'" + sheetName + "'!A:I"]);
    var values = data.valueRanges[0].values || [];
    var headers = values[0];
    
    // 2. キャッシュされたインデックスを取得
    var cachedIndex = cacheManager.get('db_index_userId');
    if (cachedIndex) {
      console.log('🔍 [CACHE_REPAIR] キャッシュインデックス検証中');
      
      // 新形式か旧形式かを判定
      var indexData = {};
      if (cachedIndex.metadata && cachedIndex.index) {
        indexData = cachedIndex.index;
        console.log('✅ [CACHE_REPAIR] 新形式インデックス検出');
      } else if (typeof cachedIndex === 'object') {
        indexData = cachedIndex;
        console.log('⚠️ [CACHE_REPAIR] 旧形式インデックス検出（更新推奨）');
        issues.push('旧形式インデックス使用中');
      }
      
      // 3. 特定ユーザーのインデックス整合性をチェック
      var cachedUserIndex = indexData[userId];
      if (cachedUserIndex) {
        var expectedRowIndex = -1;
        var userIdIndex = headers.indexOf('userId');
        
        // 実際のデータでユーザーを検索
        for (var i = 1; i < values.length; i++) {
          if (values[i] && values[i][userIdIndex] === userId) {
            expectedRowIndex = i + 1;
            break;
          }
        }
        
        // インデックスと実際のデータを比較
        var cachedRowIndex = typeof cachedUserIndex === 'object' ? cachedUserIndex.rowIndex : cachedUserIndex;
        
        if (expectedRowIndex !== cachedRowIndex) {
          console.warn('❌ [CACHE_REPAIR] インデックス不整合検出:', {
            userId: userId,
            cachedRowIndex: cachedRowIndex,
            expectedRowIndex: expectedRowIndex
          });
          issues.push(`インデックス不整合: キャッシュ=${cachedRowIndex}, 実際=${expectedRowIndex}`);
          repairActions.push('インデックス再構築');
        } else {
          console.log('✅ [CACHE_REPAIR] インデックス整合性OK');
        }
      } else {
        console.warn('⚠️ [CACHE_REPAIR] ユーザーインデックス未存在:', userId);
        issues.push('ユーザーインデックス未存在');
        repairActions.push('インデックス再構築');
      }
    } else {
      console.warn('⚠️ [CACHE_REPAIR] インデックスキャッシュ未存在');
      issues.push('インデックスキャッシュ未存在');
      repairActions.push('インデックス初期化');
    }
    
    // 4. 必要に応じて修復を実行
    if (repairActions.length > 0) {
      console.log('🔧 [CACHE_REPAIR] 修復実行:', repairActions);
      
      // インデックスを再構築
      cacheManager.clear('db_index_userId');
      var newIndex = buildRowIndexForField(values, headers, 'userId');
      cacheManager.set('db_index_userId', newIndex, { ttl: 600 }); // 10分間キャッシュ
      
      console.log('✅ [CACHE_REPAIR] インデックス修復完了');
      repairActions.push('インデックス修復完了');
    }
    
    return {
      status: 'success',
      userId: userId,
      issues: issues,
      repairActions: repairActions,
      isHealthy: issues.length === 0
    };
    
  } catch (error) {
    console.error('❌ [CACHE_REPAIR] 修復エラー:', error.message);
    return {
      status: 'error',
      userId: userId,
      error: error.message,
      issues: ['修復処理エラー'],
      repairActions: [],
      isHealthy: false
    };
  }
}

/**
 * システム全体のキャッシュ健全性チェックと自動修復
 * @returns {object} システム健全性レポート
 */
function performSystemCacheHealthCheck() {
  try {
    console.log('🔍 [SYSTEM_HEALTH] システムキャッシュ健全性チェック開始');
    
    var healthReport = {
      timestamp: new Date().toISOString(),
      overallHealth: 'good',
      issues: [],
      repairActions: [],
      cacheStats: {}
    };
    
    // キャッシュ統計を取得
    try {
      healthReport.cacheStats = {
        totalCaches: cacheManager.size || 0,
        hasDbIndex: !!cacheManager.get('db_index_userId'),
        memoryUsage: 'unknown' // GASでは取得困難
      };
    } catch (e) {
      console.warn('⚠️ [SYSTEM_HEALTH] キャッシュ統計取得失敗:', e.message);
    }
    
    // 期限切れキャッシュのクリーンアップ
    try {
      cacheManager.clearExpired();
      healthReport.repairActions.push('期限切れキャッシュクリーンアップ');
    } catch (e) {
      console.warn('⚠️ [SYSTEM_HEALTH] キャッシュクリーンアップ失敗:', e.message);
      healthReport.issues.push('キャッシュクリーンアップ失敗');
    }
    
    // 全体的な健全性を判定
    if (healthReport.issues.length > 0) {
      healthReport.overallHealth = 'degraded';
    }
    if (healthReport.issues.length > 3) {
      healthReport.overallHealth = 'poor';
    }
    
    console.log('✅ [SYSTEM_HEALTH] 健全性チェック完了:', healthReport);
    return healthReport;
    
  } catch (error) {
    console.error('❌ [SYSTEM_HEALTH] 健全性チェックエラー:', error.message);
    return {
      timestamp: new Date().toISOString(),
      overallHealth: 'error',
      issues: ['健全性チェック実行エラー'],
      repairActions: [],
      error: error.message
    };
  }
}

/**
 * ユーザー情報をキャッシュ優先で取得し、見つからなければDBから直接取得
 * @param {string} userId - ユーザーID
 * @returns {object|null} ユーザー情報
 */
function getUserWithFallback(userId) {
  // 入力検証
  if (!userId || typeof userId !== 'string') {
    console.warn('getUserWithFallback: Invalid userId:', userId);
    return null;
  }

  // 可能な限りキャッシュを利用
  var user = findUserById(userId);
  if (!user) {
    handleMissingUser(userId);
  }
  return user;
}

/**
 * ユーザー情報を一括更新
 * @param {string} userId - ユーザーID
 * @param {object} updateData - 更新データ
 * @returns {object} 更新結果
 */
function updateUser(userId, updateData) {
  try {
    console.log('updateUser: 開始', { userId: userId, updateData: updateData });
    
    var props = PropertiesService.getScriptProperties();
    var dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
    var service = getSheetsService();
    var sheetName = DB_SHEET_CONFIG.SHEET_NAME;
    
    // 現在のデータを取得
    var data = batchGetSheetsData(service, dbId, ["'" + sheetName + "'!A:I"]);
    var values = data.valueRanges[0].values || [];
    
    console.log('updateUser: データ取得完了', { 
      valuesLength: values.length,
      hasHeaders: values.length > 0 && Array.isArray(values[0]),
      headers: values.length > 0 ? values[0] : 'なし',
      sampleRows: values.length > 1 ? values.slice(1, 3) : 'なし'
    });
    
    if (values.length === 0) {
      throw new Error('データベースが空です');
    }
    
    var headers = values[0];
    var userIdIndex = headers.indexOf('userId');
    var rowIndex = -1;
    
    console.log('updateUser: ヘッダー情報', { 
      headers: headers,
      userIdIndex: userIdIndex,
      userIdExists: userIdIndex !== -1
    });
    
    // インデックスを使用してユーザーの行を効率的に特定
    // 強化されたインデックスシステムを使用
    var indexResult = cacheManager.get('db_index_userId', function() {
      return buildRowIndexForField(values, headers, 'userId');
    }, { ttl: 60 }); // 1分間に短縮
    
    // デバッグ用：キャッシュ無効化オプション
    if (updateData.__debug_clearCache) {
      cacheManager.clear('db_index_userId');
      indexResult = buildRowIndexForField(values, headers, 'userId');
      console.log('updateUser: キャッシュクリア後のインデックス再構築');
    }
    
    // インデックス結果の検証
    var userIndexData = {};
    var indexValid = false;
    
    if (indexResult && indexResult.metadata && indexResult.metadata.valid) {
      userIndexData = indexResult.index;
      indexValid = true;
      console.log('updateUser: 有効なインデックス使用', { 
        validEntries: indexResult.metadata.validEntries,
        skippedEntries: indexResult.metadata.skippedEntries,
        dataSignature: indexResult.metadata.dataSignature 
      });
    } else if (indexResult && typeof indexResult === 'object' && !indexResult.metadata) {
      // 旧形式のインデックス（後方互換性）
      userIndexData = indexResult;
      indexValid = true;
      console.warn('updateUser: 旧形式インデックスを使用（更新推奨）');
    } else {
      console.error('updateUser: インデックスが無効、線形検索に切り替え', { indexResult: indexResult });
      indexValid = false;
      
      // 🔧 自動修復: インデックスが無効な場合はキャッシュ修復を試行
      try {
        console.log('🔧 [AUTO_REPAIR] インデックス無効につき自動修復開始');
        var repairResult = validateAndRepairCache(userId);
        if (repairResult.status === 'success' && repairResult.repairActions.length > 0) {
          console.log('✅ [AUTO_REPAIR] キャッシュ修復成功、再取得試行');
          // 修復後に再度インデックスを取得
          indexResult = cacheManager.get('db_index_userId');
          if (indexResult && indexResult.metadata && indexResult.metadata.valid) {
            userIndexData = indexResult.index;
            indexValid = true;
            console.log('✅ [AUTO_REPAIR] 修復後のインデックス利用可能');
          }
        }
      } catch (repairError) {
        console.warn('⚠️ [AUTO_REPAIR] 自動修復失敗:', repairError.message);
      }
    }
    
    console.log('updateUser: ユーザーインデックス情報', { 
      hasIndexData: !!userIndexData,
      indexValid: indexValid,
      userId: userId,
      cachedRowIndex: userIndexData ? userIndexData[userId] : 'なし'
    });
    
    var cachedRowIndex = userIndexData[userId];
    var rowIndex = -1;
    if (cachedRowIndex) {
      // 新形式（構造化インデックス）か旧形式かを判定
      if (typeof cachedRowIndex === 'object' && cachedRowIndex.rowIndex) {
        // 新形式：構造化インデックス
        rowIndex = cachedRowIndex.rowIndex;
        var arrayIndex = cachedRowIndex.arrayIndex;
        console.log('updateUser: 構造化インデックスから取得', { 
          rowIndex: rowIndex, 
          arrayIndex: arrayIndex,
          lastValidated: cachedRowIndex.lastValidated 
        });
        
        // 二重検証：rowIndexとarrayIndexの整合性をチェック
        if (rowIndex >= 1 && rowIndex <= values.length && arrayIndex >= 0 && arrayIndex < values.length) {
          var cachedRow = values[arrayIndex];
          if (cachedRow && cachedRow[userIdIndex] === userId) {
            console.log('updateUser: 構造化インデックスが有効');
          } else {
            console.warn('updateUser: 構造化インデックスが無効、線形検索に切り替え', {
              rowIndex: rowIndex,
              arrayIndex: arrayIndex,
              cachedRowExists: !!cachedRow,
              cachedRowUserId: cachedRow ? cachedRow[userIdIndex] : 'N/A',
              expectedUserId: userId
            });
            rowIndex = -1;
          }
        } else {
          console.warn('updateUser: 構造化インデックスの範囲が無効', { 
            rowIndex: rowIndex, 
            arrayIndex: arrayIndex, 
            valuesLength: values.length 
          });
          rowIndex = -1;
        }
      } else {
        // 旧形式：単純な行番号
        rowIndex = cachedRowIndex;
        console.log('updateUser: 旧形式インデックスから取得', { rowIndex: rowIndex });
        
        if (rowIndex >= 1 && rowIndex <= values.length) {
          var cachedRow = values[rowIndex - 1];
          if (cachedRow && cachedRow[userIdIndex] === userId) {
            console.log('updateUser: 旧形式インデックスが有効');
          } else {
            console.warn('updateUser: 旧形式インデックスが無効、線形検索に切り替え', {
              rowIndex: rowIndex,
              cachedRowExists: !!cachedRow,
              cachedRowUserId: cachedRow ? cachedRow[userIdIndex] : 'N/A',
              expectedUserId: userId
            });
            rowIndex = -1;
          }
        } else {
          console.warn('updateUser: 旧形式インデックスが範囲外', { rowIndex: rowIndex, valuesLength: values.length });
          rowIndex = -1;
        }
      }
    }
    
    if (rowIndex === -1) {
      // フォールバック: 線形検索
      console.log('updateUser: 線形検索開始');
      for (var i = 1; i < values.length; i++) {
        var currentRow = values[i];
        if (currentRow && currentRow[userIdIndex] === userId) {
          rowIndex = i + 1; // 1-based index
          console.log('updateUser: ユーザー発見', { rowIndex: rowIndex, arrayIndex: i });
          break;
        }
      }
    }

    // キャッシュが古く行番号が範囲外の場合は再計算
    if (rowIndex < 1 || rowIndex > values.length) {
      for (var i = 1; i < values.length; i++) {
        if (values[i][userIdIndex] === userId) {
          rowIndex = i + 1; // 1-based index
          break;
        }
      }
      if (rowIndex < 1 || rowIndex > values.length) {
        throw new Error('更新対象のユーザーが見つかりません');
      }
    }
    
    if (rowIndex === -1) {
      console.error('updateUser: 最終的にユーザーが見つかりませんでした', {
        userId: userId,
        searchedRows: values.length - 1,
        allUserIds: values.slice(1).map(function(row, index) {
          return {
            rowIndex: index + 1,
            userId: row && row[userIdIndex] ? row[userIdIndex] : 'N/A',
            rowData: row
          };
        })
      });
      
      // 自動修復を試行：ユーザーが存在しない場合は作成
      console.log('updateUser: ユーザーが存在しないため、自動作成を試行');
      try {
        // 基本的なユーザーデータを作成
        var basicUserData = {
          userId: userId,
          adminEmail: updateData.adminEmail || 'unknown@example.com',
          createdAt: new Date().toISOString(),
          lastAccessedAt: new Date().toISOString(),
          setupStatus: 'user_created'
        };
        
        var newUser = createUserAtomic(basicUserData);
        if (newUser) {
          console.log('updateUser: ユーザー自動作成成功', { userId: userId });
          
          // 作成後に元の更新データを適用
          setTimeout(function() {
            try {
              updateUser(userId, updateData);
            } catch (retryError) {
              console.error('updateUser: 再試行失敗', { error: retryError.message });
            }
          }, 100);
          
          return { success: true, message: 'ユーザーが自動作成されました', created: true };
        }
      } catch (createError) {
        console.error('updateUser: ユーザー自動作成失敗', { error: createError.message });
      }
      
      throw new Error('更新対象のユーザーが見つかりません (自動作成も失敗)');
    }
    
    // 配列の範囲チェック
    if (rowIndex < 1 || rowIndex > values.length) {
      console.error('updateUser: rowIndexが範囲外です', {
        rowIndex: rowIndex,
        valuesLength: values.length,
        userId: userId
      });
      throw new Error('更新対象のユーザーデータが見つかりません (行インデックスが範囲外)');
    }
    
    // バックアップ用に現在のデータを保存
    var targetRow = values[rowIndex - 1]; // 1-based to 0-based
    
    console.log('updateUser: 最終行データ検証', {
      rowIndex: rowIndex,
      arrayIndex: rowIndex - 1,
      valuesLength: values.length,
      targetRowType: typeof targetRow,
      targetRowIsArray: Array.isArray(targetRow),
      targetRowLength: targetRow ? targetRow.length : 'N/A',
      targetRowUserId: targetRow && targetRow[userIdIndex] ? targetRow[userIdIndex] : 'N/A',
      expectedUserId: userId
    });
    
    if (!targetRow || !Array.isArray(targetRow)) {
      console.error('updateUser: 対象行がundefinedまたは配列ではありません', {
        rowIndex: rowIndex,
        valuesLength: values.length,
        targetRow: targetRow,
        userId: userId,
        allRowsDebug: values.map(function(row, index) {
          return {
            index: index,
            isArray: Array.isArray(row),
            length: row ? row.length : 'N/A',
            userId: row && row[userIdIndex] ? row[userIdIndex] : 'N/A'
          };
        })
      });
      throw new Error('更新対象のユーザーデータが見つかりません (行データが不正)');
    }
    
    // 最終確認：行のユーザーIDが一致するか
    if (targetRow[userIdIndex] !== userId) {
      console.error('updateUser: 行のユーザーIDが一致しません', {
        rowIndex: rowIndex,
        actualUserId: targetRow[userIdIndex],
        expectedUserId: userId,
        targetRow: targetRow
      });
      throw new Error('更新対象のユーザーデータが見つかりません (ユーザーIDが一致しません)');
    }
    
    // バッチ更新リクエストを作成
    var requests = Object.keys(updateData).map(function(key) {
      var colIndex = headers.indexOf(key);
      if (colIndex === -1) {
        console.warn('未知のフィールドをスキップ:', key);
        return null;
      }
      
      return {
        range: "'" + sheetName + "'!" + String.fromCharCode(65 + colIndex) + rowIndex,
        values: [[updateData[key]]]
      };
    }).filter(function(item) { return item !== null; });
    
    if (requests.length === 0) {
      console.warn('更新するフィールドがありません');
      return { success: true, message: '更新フィールドなし' };
    }
    
    console.log('💾 [DEBUG] Updating user data:', {
      userId: userId,
      updateFields: Object.keys(updateData),
      requestCount: requests.length
    });
    
    try {
      batchUpdateSheetsData(service, dbId, requests);
      console.log('✅ [DEBUG] Database update completed successfully');
      
      // 更新成功を検証 - データを再取得して確認
      var verificationData = batchGetSheetsData(service, dbId, ["'" + sheetName + "'!" + rowIndex + ":" + rowIndex]);
      if (verificationData.valueRanges[0].values && verificationData.valueRanges[0].values.length > 0) {
        console.log('✅ [DEBUG] Database update verification successful');
      } else {
        throw new Error('更新後のデータ検証に失敗しました');
      }
    } catch (updateError) {
      console.error('⚠️ [ERROR] Database update failed:', updateError);
      // ロールバックは実装しない（Google Sheets APIの制限）
      // 代わりにエラーを伝播して呼び出し元で処理
      throw new Error('データベース更新に失敗しました: ' + updateError.message);
    }
    
    // スプレッドシートIDが更新された場合、サービスアカウントと共有
    if (updateData.spreadsheetId) {
      try {
        shareSpreadsheetWithServiceAccount(updateData.spreadsheetId);
        console.log('ユーザー更新時のサービスアカウント共有完了:', updateData.spreadsheetId);
      } catch (shareError) {
        console.error('ユーザー更新時のサービスアカウント共有エラー:', shareError.message);
        console.error('スプレッドシートの更新は完了しましたが、サービスアカウントとの共有に失敗しました。手動で共有してください。');
      }
    }
    
    // 統合キャッシュシステムに対応した無効化処理
    
    // 1. 実行レベルキャッシュを即座にクリア
    if (typeof clearExecutionUserInfoCache === 'function') {
      clearExecutionUserInfoCache(userId);
      console.log('✅ 実行レベルキャッシュクリア完了: updateUser後');
    }
    
    // 2. 従来のキャッシュシステムも無効化
    var userInfo = findUserById(userId, true); // キャッシュバイパスして最新情報を取得
    var email = updateData.adminEmail || (userInfo ? userInfo.adminEmail : null);
    var spreadsheetId = updateData.spreadsheetId || (userInfo ? userInfo.spreadsheetId : null);
    
    // 3. 段階的キャッシュ無効化（部分 → 全体）
    invalidateUserCache(userId, email, spreadsheetId, false);
    
    // 4. データベースインデックスキャッシュも更新が必要な場合はクリア
    if (updateData.adminEmail || updateData.userId) {
      // ユーザーIDやメールアドレスが変更された場合はインデックスキャッシュもクリア
      cacheManager.clearByPattern('db_index_');
      console.log('✅ データベースインデックスキャッシュもクリア');
    }
    
    console.log('✅ updateUser: キャッシュ無効化完了', { userId, email, spreadsheetId });
    
    return { success: true };
  } catch (error) {
    console.error('ユーザー更新エラー:', {
      userId: userId,
      error: error.message,
      stack: error.stack,
      updateData: updateData
    });
    
    // 特定のエラーケースの処理
    if (error.message && error.message.includes('ユーザーが自動作成されました')) {
      return { success: true, message: error.message, created: true };
    }
    
    throw error;
  }
}

/**
 * 🎯 プロダクション環境向け: ユーザー検索・作成（リトライ機能付き）- 統合版
 * メール特化ロックシステムを使用
 * @param {string} adminEmail - ユーザーメールアドレス
 * @param {object} additionalData - 追加データ
 * @param {number} maxRetries - 最大リトライ回数（デフォルト: 3）
 * @param {number} initialBackoff - 初期待機時間（ミリ秒、デフォルト: 1000）
 * @returns {object} { userId, isNewUser, userInfo }
 */
function findOrCreateUserProductionWithRetry(adminEmail, additionalData = {}, maxRetries = 3, initialBackoff = 1000) {
  let attempts = 0;
  let lastError = null;

  while (attempts < maxRetries) {
    try {
      // メール特化ロックシステムを直接使用
      const result = findOrCreateUserWithEmailLock(adminEmail, additionalData);
      
      // 成功した場合は結果を返す
      return result;
      
    } catch (error) {
      lastError = error;
      
      // EmailLock関連エラーの場合のみリトライ
      if (error.message && (
        error.message.includes('LOCK_TIMEOUT') || 
        error.message.includes('SCRIPT_LOCK_TIMEOUT') ||
        error.message.includes('EMAIL_ALREADY_PROCESSING') ||
        error.message.includes('Lock wait timeout')
      )) {
        attempts++;
        
        if (attempts < maxRetries) {
          const waitTime = initialBackoff * Math.pow(2, attempts - 1) + Math.random() * 500;
          console.warn(`findOrCreateUserProductionWithRetry: EmailLock失敗, リトライ ${attempts}/${maxRetries} (${waitTime}ms待機)...`, { adminEmail });
          Utilities.sleep(waitTime);
        }
      } else {
        // EMAIL_ALREADY_PROCESSING以外のエラーは即時スロー
        throw error;
      }
    }
  }

  // リトライ上限に達した場合
  console.error(`findOrCreateUserProductionWithRetry 最終エラー: ${lastError.message}`, {
    adminEmail: adminEmail,
    attempts: attempts,
    lastErrorMessage: lastError.toString()
  });
  
  // 最後のタイムアウトエラーをスロー
  throw new Error(`ユーザー情報の確保に失敗しました: LOCK_TIMEOUT`);
}

/**
 * 🎯 原子的なユーザー検索・作成操作（Upsert）
 * 重複を防ぐためのメイン関数
 * @param {string} adminEmail - ユーザーメールアドレス
 * @param {object} additionalData - 追加データ（オプション）
 * @returns {object} { userId, isNewUser, userInfo }
 */
/**
 * @deprecated Use findOrCreateUserWithEmailLock instead
 * 競合するロック戦略を防ぐため無効化
 */
function findOrCreateUser(adminEmail, additionalData = {}) {
  console.warn('findOrCreateUser is deprecated. Use findOrCreateUserWithEmailLock instead.');
  // 直接EmailLockシステムにリダイレクト
  return findOrCreateUserWithEmailLock(adminEmail, additionalData);
}

/**
 * 🎯 適応的ロックでのユーザー作成試行
 * @param {string} adminEmail - メールアドレス
 * @param {object} additionalData - 追加データ
 * @returns {object|null} 成功時は結果、失敗時はnull
 */
function attemptWithAdaptiveLock(adminEmail, additionalData) {
  const lock = LockService.getScriptLock();
  const maxWaitTime = 10000; // 10秒に短縮
  
  try {
    if (!lock.waitLock(maxWaitTime)) {
      debugLog('attemptWithAdaptiveLock: タイムアウト', { adminEmail, waitTime: maxWaitTime });
      return null; // フォールバックに委謗
    }

    debugLog('attemptWithAdaptiveLock: ロック取得成功', { adminEmail });

    // 1. 既存ユーザー確認
    let existingUser = findUserByEmail(adminEmail);
    
    if (existingUser) {
      debugLog('findOrCreateUser: 既存ユーザー発見', { userId: existingUser.userId, adminEmail });
      
      // 必要に応じて既存ユーザー情報を更新
      if (Object.keys(additionalData).length > 0) {
        const updateData = {
          lastAccessedAt: new Date().toISOString(),
          isActive: 'true',
          ...additionalData
        };
        updateUser(existingUser.userId, updateData);
        debugLog('findOrCreateUser: 既存ユーザー更新完了', { userId: existingUser.userId });
      }
      
      return {
        userId: existingUser.userId,
        isNewUser: false,
        userInfo: existingUser
      };
    }

    // 2. 新規ユーザー作成
    debugLog('findOrCreateUser: 新規ユーザー作成開始', { adminEmail });
    
    const userId = generateConsistentUserId(adminEmail);
    const userData = {
      userId: userId,
      adminEmail: adminEmail,
      createdAt: new Date().toISOString(),
      lastAccessedAt: new Date().toISOString(),
      isActive: 'true',
      configJson: '{}',
      spreadsheetId: '',
      spreadsheetUrl: '',
      ...additionalData
    };

    // 原子的作成（重複チェックなし - ロック内なので安全）
    createUserAtomic(userData);
    
    debugLog('findOrCreateUser: 新規ユーザー作成完了', { userId, adminEmail });
    
    return {
      userId: userId,
      isNewUser: true,
      userInfo: userData
    };

  } catch (error) {
    console.error('findOrCreateUser エラー:', error);
    throw error;
  } finally {
    // 必ずロックを解除
    try {
      lock.releaseLock();
      debugLog('findOrCreateUser: ロック解除完了', { adminEmail });
    } catch (e) {
      console.warn('ロック解除エラー:', e.message);
    }
  }
}

/**
 * ⚡ 軽量ロックでのユーザー作成試行
 * 適応的ロック失敗時のフォールバック
 * @param {string} adminEmail - メールアドレス
 * @param {object} additionalData - 追加データ
 * @returns {object} { userId, isNewUser, userInfo }
 */
function attemptWithLightweightLock(adminEmail, additionalData) {
  const lock = LockService.getScriptLock();
  const timeout = 3000; // 3秒

  try {
    if (!lock.waitLock(timeout)) {
      debugLog('attemptWithLightweightLock: タイムアウト', { adminEmail, timeout });
      const existingUser = findUserByEmail(adminEmail);
      if (existingUser) {
        debugLog('attemptWithLightweightLock: 既存ユーザー返却', { userId: existingUser.userId });
        return {
          userId: existingUser.userId,
          isNewUser: false,
          userInfo: existingUser
        };
      }
      throw new Error('ユーザー情報の確保に失敗しました: attemptWithLightweightLock');
    }

    debugLog('attemptWithLightweightLock: ロック取得成功', { adminEmail });

    const existingUser = findUserByEmail(adminEmail);

    if (existingUser) {
      if (Object.keys(additionalData).length > 0) {
        const updateData = {
          lastAccessedAt: new Date().toISOString(),
          isActive: 'true',
          ...additionalData
        };
        updateUser(existingUser.userId, updateData);
      }

      return {
        userId: existingUser.userId,
        isNewUser: false,
        userInfo: existingUser
      };
    }

    const userId = generateConsistentUserId(adminEmail);
    const userData = {
      userId: userId,
      adminEmail: adminEmail,
      createdAt: new Date().toISOString(),
      lastAccessedAt: new Date().toISOString(),
      isActive: 'true',
      configJson: '{}',
      spreadsheetId: '',
      spreadsheetUrl: '',
      ...additionalData
    };

    createUserAtomic(userData);

    debugLog('attemptWithLightweightLock: 新規ユーザー作成完了', { userId });

    return {
      userId: userId,
      isNewUser: true,
      userInfo: userData
    };
  } catch (error) {
    console.error('attemptWithLightweightLock エラー:', error);
    throw error;
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {
      console.warn('軽量ロック解除エラー:', e.message);
    }
  }
}

/**
 * 🔧 一貫したユーザーID生成
 * メールアドレスから決定論的にUUIDを生成
 * @param {string} adminEmail - メールアドレス
 * @returns {string} 一意のユーザーID
 */
function generateConsistentUserId(adminEmail) {
  // メールアドレスからハッシュ値を生成
  const hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, adminEmail, Utilities.Charset.UTF_8);
  const hexString = hash.map(byte => (byte + 256).toString(16).slice(-2)).join('');
  
  // UUID v4 フォーマットに整形
  const uuid = [
    hexString.slice(0, 8),
    hexString.slice(8, 12),
    '4' + hexString.slice(13, 16), // version 4
    ((parseInt(hexString.slice(16, 17), 16) & 0x3) | 0x8).toString(16) + hexString.slice(17, 20), // variant
    hexString.slice(20, 32)
  ].join('-');
  
  debugLog('generateConsistentUserId', { adminEmail, userId: uuid });
  return uuid;
}

/**
 * ⚡ 原子的ユーザー作成（重複チェックなし）
 * findOrCreateUser内でのみ使用 - ロック保護下での高速作成
 * サービスアカウント専用でデータベースにアクセス
 * @param {object} userData - 作成するユーザーデータ
 * @returns {object} 作成されたユーザーデータ
 */
function createUserAtomic(userData) {
  const props = PropertiesService.getScriptProperties();
  const dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
  const sheetName = DB_SHEET_CONFIG.SHEET_NAME;

  const newRow = DB_SHEET_CONFIG.HEADERS.map(function(header) { 
    return userData[header] || ''; 
  });
  
  // サービスアカウント専用でアクセス
  const service = getSheetsService();
  appendSheetsData(service, dbId, "'" + sheetName + "'!A1", [newRow]);
  console.log('User created via Service Account');
  
  // キャッシュ無効化
  invalidateUserCache(userData.userId, userData.adminEmail, userData.spreadsheetId, false);
  
  return userData;
}

/**
 * 🔄 レガシー互換: 新規ユーザーをデータベースに作成
 * @deprecated findOrCreateUser を使用してください
 * @param {object} userData - 作成するユーザーデータ
 * @returns {object} 作成されたユーザーデータ
 */
function createUser(userData) {
  console.warn('createUser は非推奨です。findOrCreateUser を使用してください。');
  
  // 既存の動作を維持（後方互換性のため）
  var existingUser = findUserByEmail(userData.adminEmail);
  if (existingUser) {
    throw new Error('このメールアドレスは既に登録されています。');
  }

  return createUserAtomic(userData);
}

/**
 * データベースシートを初期化
 * @param {string} spreadsheetId - データベースのスプレッドシートID
 */
function initializeDatabaseSheet(spreadsheetId) {
  var service = getSheetsService();
  var sheetName = DB_SHEET_CONFIG.SHEET_NAME;

  try {
    // シートが存在するか確認
    var spreadsheet = getSpreadsheetsData(service, spreadsheetId);
    var sheetExists = spreadsheet.sheets.some(function(s) { 
      return s.properties.title === sheetName; 
    });

    if (!sheetExists) {
      // シートが存在しない場合は作成
      batchUpdateSpreadsheet(service, spreadsheetId, {
        requests: [{ addSheet: { properties: { title: sheetName } } }]
      });
    }
    
    // ヘッダーを書き込み
    var headerRange = "'" + sheetName + "'!A1:" + String.fromCharCode(65 + DB_SHEET_CONFIG.HEADERS.length - 1) + '1';
    updateSheetsData(service, spreadsheetId, headerRange, [DB_SHEET_CONFIG.HEADERS]);

    debugLog('データベースシート「' + sheetName + '」の初期化が完了しました。');
  } catch (e) {
    console.error('データベースシートの初期化に失敗: ' + e.message);
    throw new Error('データベースシートの初期化に失敗しました。サービスアカウントに編集者権限があるか確認してください。詳細: ' + e.message);
  }
}

/**
 * ユーザーが見つからない場合のキャッシュ処理
 * @param {string} userId - キャッシュ削除対象のユーザーID
 */
function handleMissingUser(userId) {
  try {
    // 最適化: 特定のユーザーキャッシュのみ削除（全体のキャッシュクリアは避ける）
    if (userId) {
      // 特定ユーザーのキャッシュを削除
      cacheManager.remove('user_' + userId);
      
      // 関連するメールキャッシュも削除（可能な場合）
      try {
        var userProps = PropertiesService.getUserProperties();
        var currentUserId = userProps.getProperty('CURRENT_USER_ID');
        if (currentUserId === userId) {
          // 現在のユーザーIDが無効な場合はクリア
          userProps.deleteProperty('CURRENT_USER_ID');
          debugLog('[Cache] Cleared invalid CURRENT_USER_ID: ' + userId);
        }
      } catch (propsError) {
        console.warn('Failed to clear user properties:', propsError.message);
      }
    }
    
    // 全データベースキャッシュのクリアは最後の手段として実行
    // 頻繁な実行を避けるため、確実にデータ不整合がある場合のみ実行
    var shouldClearAll = false;
    
    // 判定条件: 複数のユーザーで問題が発生している場合のみ全クリア
    if (shouldClearAll) {
      clearDatabaseCache();
    }
    
    debugLog('[Cache] Handled missing user: ' + userId);
  } catch (error) {
    console.error('handleMissingUser error:', error.message);
    // エラーが発生した場合のみ全キャッシュクリア
    clearDatabaseCache();
  }
}


// =================================================================
// 最適化されたSheetsサービス関数群
// =================================================================

/**
 * 最適化されたSheetsサービスを作成
 * @param {string} accessToken - アクセストークン
 * @returns {object} Sheetsサービスオブジェクト
 */
function createSheetsService(accessToken) {
  return {
    accessToken: accessToken,
    baseUrl: 'https://sheets.googleapis.com/v4/spreadsheets',
    spreadsheets: {
      values: {
        get: function(options) {
          var url = 'https://sheets.googleapis.com/v4/spreadsheets/' + 
                   options.spreadsheetId + '/values/' + encodeURIComponent(options.range);
          
          var response = UrlFetchApp.fetch(url, {
            headers: { 'Authorization': 'Bearer ' + accessToken },
            muteHttpExceptions: true,
            followRedirects: true,
            validateHttpsCertificates: true
          });
          
          if (response.getResponseCode() !== 200) {
            throw new Error('Sheets API error: ' + response.getResponseCode() + ' - ' + response.getContentText());
          }
          
          return JSON.parse(response.getContentText());
        }
      },
      get: function(options) {
        var url = 'https://sheets.googleapis.com/v4/spreadsheets/' + options.spreadsheetId;
        if (options.fields) {
          url += '?fields=' + encodeURIComponent(options.fields);
        }
        
        var response = UrlFetchApp.fetch(url, {
          headers: { 'Authorization': 'Bearer ' + accessToken },
          muteHttpExceptions: true,
          followRedirects: true,
          validateHttpsCertificates: true
        });
        
        if (response.getResponseCode() !== 200) {
          throw new Error('Sheets API error: ' + response.getResponseCode() + ' - ' + response.getContentText());
        }
        
        return JSON.parse(response.getContentText());
      }
    }
  };
}

/**
 * 強化されたバッチ取得（最適化版）
 * @param {object} service - Sheetsサービス
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string[]} ranges - 取得範囲の配列
 * @param {object} options - オプション { priority, maxRetries, batchSize }
 * @returns {object} レスポンス
 */
function batchGetSheetsData(service, spreadsheetId, ranges, options = {}) {
  const { 
    priority = 'normal',
    maxRetries = 2,
    batchSize = 10 // 1回のAPIコールで処理する範囲数の上限
  } = options;

  // 入力検証
  if (!ranges || ranges.length === 0) {
    return { valueRanges: [] };
  }

  // キャッシュキーの最適化（範囲が多い場合はハッシュ化）
  const rangesKey = ranges.length > 5 ? 
    Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(ranges))
      .map(byte => (byte + 256).toString(16).slice(-2)).join('').substr(0, 16) :
    JSON.stringify(ranges);
    
  const cacheKey = `batchGet_${spreadsheetId}_${rangesKey}`;
  
  return cacheManager.get(cacheKey, () => {
    try {
      // 大量の範囲を分割して処理
      if (ranges.length > batchSize) {
        return processLargeBatch(service, spreadsheetId, ranges, batchSize, maxRetries);
      }
      
      // 通常のバッチ処理
      return executeBatchGet(service, spreadsheetId, ranges, maxRetries);
      
    } catch (error) {
      const processedError = handleSheetsApiError(error, {
        function: 'batchGetSheetsData',
        spreadsheetId: spreadsheetId,
        rangeCount: ranges.length,
        operation: 'batch_get'
      });
      throw new Error(processedError.userMessage);
    }
  }, { 
    ttl: priority === 'high' ? 60 : 180, // 高優先度は短いキャッシュ
    priority: priority 
  });
}

/**
 * 大量バッチの分割処理
 * @private
 */
function processLargeBatch(service, spreadsheetId, ranges, batchSize, maxRetries) {
  const results = { valueRanges: [] };
  const chunks = [];
  
  // 範囲を分割
  for (let i = 0; i < ranges.length; i += batchSize) {
    chunks.push(ranges.slice(i, i + batchSize));
  }
  
  console.log(`[BatchGet] 大量データを${chunks.length}個のチャンクに分割`);
  
  // 各チャンクを順次処理（並列処理は API制限回避のため避ける）
  for (let i = 0; i < chunks.length; i++) {
    try {
      const chunkResult = executeBatchGet(service, spreadsheetId, chunks[i], maxRetries);
      results.valueRanges.push(...chunkResult.valueRanges);
      
      // レート制限対策（チャンク間で少し待機）
      if (i < chunks.length - 1) {
        Utilities.sleep(100); // 100ms待機
      }
      
    } catch (error) {
      console.error(`[BatchGet] チャンク ${i + 1} の処理に失敗:`, error.message);
      throw error;
    }
  }
  
  return results;
}

/**
 * バッチ取得の実行
 * @private
 */
function executeBatchGet(service, spreadsheetId, ranges, maxRetries) {
  return retryWithBackoff(() => {
    const url = service.baseUrl + '/' + spreadsheetId + '/values:batchGet?' + 
      ranges.map(range => 'ranges=' + encodeURIComponent(range)).join('&') +
      '&valueRenderOption=UNFORMATTED_VALUE&dateTimeRenderOption=FORMATTED_STRING';
    
    const response = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + service.accessToken },
      muteHttpExceptions: true,
      followRedirects: true,
      validateHttpsCertificates: true
    });
    
    if (response.getResponseCode() !== 200) {
      const errorMsg = `Sheets API error: ${response.getResponseCode()} - ${response.getContentText()}`;
      console.error('[BatchGet]', errorMsg);
      throw new Error(errorMsg);
    }
    
    return JSON.parse(response.getContentText());
    
  }, { 
    maxRetries: maxRetries,
    baseDelay: 1000,
    maxDelay: 5000
  });
}

/**
 * バッチ更新
 * @param {object} service - Sheetsサービス
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {object[]} requests - 更新リクエストの配列
 * @returns {object} レスポンス
 */
function batchUpdateSheetsData(service, spreadsheetId, requests) {
  try {
    var url = service.baseUrl + '/' + spreadsheetId + '/values:batchUpdate';
    
    var response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + service.accessToken },
      payload: JSON.stringify({
        data: requests,
        valueInputOption: 'RAW'
      }),
      muteHttpExceptions: true,
      followRedirects: true,
      validateHttpsCertificates: true
    });
    
    if (response.getResponseCode() !== 200) {
      throw new Error('Sheets API error: ' + response.getResponseCode() + ' - ' + response.getContentText());
    }
    
    return JSON.parse(response.getContentText());
  } catch (error) {
    console.error('batchUpdateSheetsData error:', error.message);
    throw new Error('データ更新に失敗しました: ' + error.message);
  }
}

/**
 * データ追加
 * @param {object} service - Sheetsサービス
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} range - 範囲
 * @param {array} values - 値の配列
 * @returns {object} レスポンス
 */
function appendSheetsData(service, spreadsheetId, range, values) {
  var url = service.baseUrl + '/' + spreadsheetId + '/values/' + encodeURIComponent(range) + 
    ':append?valueInputOption=RAW&insertDataOption=INSERT_ROWS';
  
  var response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + service.accessToken },
    payload: JSON.stringify({ values: values })
  });
  
  return JSON.parse(response.getContentText());
}

/**
 * スプレッドシート情報取得
 * @param {object} service - Sheetsサービス
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {object} スプレッドシート情報
 */
function getSpreadsheetsData(service, spreadsheetId) {
  try {
    // シート情報を含む基本的なメタデータを取得するために fields パラメータを追加
    var url = service.baseUrl + '/' + spreadsheetId + '?fields=sheets.properties';
    var response = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + service.accessToken },
      muteHttpExceptions: true,
      followRedirects: true,
      validateHttpsCertificates: true
    });
    
    if (response.getResponseCode() !== 200) {
      console.error('Sheets API response code:', response.getResponseCode());
      console.error('Sheets API response body:', response.getContentText());
      
      if (response.getResponseCode() === 403) {
        var errorResponse = JSON.parse(response.getContentText());
        if (errorResponse.error && errorResponse.error.message === 'The caller does not have permission') {
          var serviceAccountEmail = getServiceAccountEmail();
          throw new Error('スプレッドシートへのアクセス権限がありません。サービスアカウント（' + serviceAccountEmail + '）をスプレッドシートの編集者として共有してください。');
        }
      }
      
      throw new Error('Sheets API error: ' + response.getResponseCode() + ' - ' + response.getContentText());
    }
    
    var result = JSON.parse(response.getContentText());
    console.log('getSpreadsheetsData success: Found', result.sheets ? result.sheets.length : 0, 'sheets');
    return result;
  } catch (error) {
    console.error('getSpreadsheetsData error:', error.message);
    console.error('getSpreadsheetsData error stack:', error.stack);
    throw new Error('スプレッドシート情報取得に失敗しました: ' + error.message);
  }
}

/**
 * すべてのユーザー情報を取得
 * @returns {Array} ユーザー情報配列
 */
function getAllUsers() {
  try {
    var props = PropertiesService.getScriptProperties();
    var dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
    var service = getSheetsService();
    var sheetName = DB_SHEET_CONFIG.SHEET_NAME;
    
    var data = batchGetSheetsData(service, dbId, ["'" + sheetName + "'!A:I"]);
    var values = data.valueRanges[0].values || [];
    
    if (values.length <= 1) {
      return []; // ヘッダーのみの場合は空配列を返す
    }
    
    var headers = values[0];
    var users = [];
    
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      var user = {};
      
      for (var j = 0; j < headers.length; j++) {
        user[headers[j]] = row[j] || '';
      }
      
      users.push(user);
    }
    
    return users;
    
  } catch (error) {
    console.error('getAllUsers エラー:', error.message);
    throw new Error('全ユーザー情報の取得に失敗しました: ' + error.message);
  }
}

/**
 * データ更新
 * @param {object} service - Sheetsサービス
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} range - 範囲
 * @param {array} values - 値の配列
 * @returns {object} レスポンス
 */
function updateSheetsData(service, spreadsheetId, range, values) {
  var url = service.baseUrl + '/' + spreadsheetId + '/values/' + encodeURIComponent(range) + 
    '?valueInputOption=RAW';
  
  var response = UrlFetchApp.fetch(url, {
    method: 'put',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + service.accessToken },
    payload: JSON.stringify({ values: values })
  });
  
  return JSON.parse(response.getContentText());
}

/**
 * スプレッドシート構造更新
 * @param {object} service - Sheetsサービス
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {object} requestBody - リクエストボディ
 * @returns {object} レスポンス
 */
function batchUpdateSpreadsheet(service, spreadsheetId, requestBody) {
  var url = service.baseUrl + '/' + spreadsheetId + ':batchUpdate';
  var response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + service.accessToken },
    payload: JSON.stringify(requestBody)
  });
  return JSON.parse(response.getContentText());
}

/**
 * @deprecated この関数は非推奨です。サービスアカウント経由のアクセスを使用してください
 * データベースシートを取得
 * @returns {object} データベースシート
 */
function getDbSheet() {
  throw new Error('getDbSheet関数は非推奨です。データベースアクセスはサービスアカウント経由でのみ行ってください。');
}

/**
 * 現在のユーザーのアカウント情報をデータベースから完全に削除する
 * @returns {string} 成功メッセージ
 */
function deleteUserAccount(userId) {
  try {
    if (!userId) {
      throw new Error('ユーザーが特定できません。');
    }
    
    // ユーザー情報を取得して、関連情報を得る
    const userInfo = findUserById(userId);
    if (!userInfo) {
      throw new Error('データベースにユーザー情報が見つかりません。');
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(15000); // 15秒待機

    try {
      // データベース（シート）からユーザー行を削除（サービスアカウント経由）
      var props = PropertiesService.getScriptProperties();
      var dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
      if (!dbId) {
        throw new Error('データベースIDが設定されていません');
      }
      
      var service = getSheetsService();
      if (!service) {
        throw new Error('Sheets APIサービスの初期化に失敗しました');
      }
      
      var sheetName = DB_SHEET_CONFIG.SHEET_NAME;
      
      // データベーススプレッドシートの情報を取得してsheetIdを確認
      var spreadsheetInfo = getSpreadsheetsData(service, dbId);
      
      var targetSheetId = null;
      for (var i = 0; i < spreadsheetInfo.sheets.length; i++) {
        if (spreadsheetInfo.sheets[i].properties.title === sheetName) {
          targetSheetId = spreadsheetInfo.sheets[i].properties.sheetId;
          break;
        }
      }
      
      if (targetSheetId === null) {
        throw new Error('データベースシート「' + sheetName + '」が見つかりません');
      }
      
      console.log('Found database sheet with sheetId:', targetSheetId);
      
      // データを取得
      var data = batchGetSheetsData(service, dbId, ["'" + sheetName + "'!A:I"]);
      var values = data.valueRanges[0].values || [];
      
      // ユーザーIDに基づいて行を探す（A列がIDと仮定）
      var rowToDelete = -1;
      for (let i = values.length - 1; i >= 1; i--) {
        if (values[i][0] === userId) {
          rowToDelete = i + 1; // スプレッドシートは1ベース
          break;
        }
      }
      
      if (rowToDelete !== -1) {
        console.log('Deleting row:', rowToDelete, 'from sheetId:', targetSheetId);
        
        // 行を削除（正しいsheetIdを使用）
        var deleteRequest = {
          deleteDimension: {
            range: {
              sheetId: targetSheetId,
              dimension: 'ROWS',
              startIndex: rowToDelete - 1, // APIは0ベース
              endIndex: rowToDelete
            }
          }
        };
        
        batchUpdateSpreadsheet(service, dbId, {
          requests: [deleteRequest]
        });
        
        console.log('Row deletion completed successfully');
      } else {
        console.warn('User row not found for deletion, userId:', userId);
      }
      
      // 削除ログを記録
      logAccountDeletion(
        userInfo.adminEmail,
        userId,
        userInfo.adminEmail,
        '自己削除',
        'self'
      );
      
      // 関連するすべてのキャッシュを削除
      invalidateUserCache(userId, userInfo.adminEmail, userInfo.spreadsheetId, false);
      
      // UserPropertiesからも関連情報を削除
      const userProps = PropertiesService.getUserProperties();
      userProps.deleteProperty('CURRENT_USER_ID');

    } finally {
      lock.releaseLock();
    }
    
    const successMessage = 'アカウント「' + userInfo.adminEmail + '」が正常に削除されました。';
    console.log(successMessage);
    return successMessage;

  } catch (error) {
    console.error('アカウント削除エラー:', error.message);
    console.error('アカウント削除エラー詳細:', error.stack);
    
    // より詳細なエラー情報を提供
    var errorMessage = 'アカウントの削除に失敗しました: ' + error.message;
    
    if (error.message.includes('No grid with id')) {
      errorMessage += '\n詳細: データベースシートのIDが正しく取得できませんでした。データベース設定を確認してください。';
    } else if (error.message.includes('permissions')) {
      errorMessage += '\n詳細: サービスアカウントの権限が不足している可能性があります。';
    } else if (error.message.includes('not found')) {
      errorMessage += '\n詳細: 削除対象のユーザーまたはシートが見つかりませんでした。';
    }
    
    throw new Error(errorMessage);
  }
}

/**
 * データベースのセキュリティ状態を検証
 * @returns {object} 検証結果
 */
function validateDatabaseSecurity() {
  try {
    const props = PropertiesService.getScriptProperties();
    const dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
    const serviceAccountEmail = getServiceAccountEmail();
    
    if (!dbId) {
      return {
        isSecure: false,
        message: 'データベーススプレッドシートIDが設定されていません'
      };
    }
    
    if (serviceAccountEmail === 'サービスアカウント未設定') {
      return {
        isSecure: false,
        message: 'サービスアカウントが設定されていません'
      };
    }
    
    // サービスアカウントでのアクセステスト
    try {
      const service = getSheetsService();
      const data = batchGetSheetsData(service, dbId, ["'" + DB_SHEET_CONFIG.SHEET_NAME + "'!A1:A1"]);
      
      return {
        isSecure: true,
        message: 'データベースはサービスアカウント専用で保護されています',
        serviceAccount: serviceAccountEmail,
        databaseId: dbId
      };
    } catch (accessError) {
      return {
        isSecure: false,
        message: 'サービスアカウントでデータベースにアクセスできません: ' + accessError.message,
        serviceAccount: serviceAccountEmail,
        databaseId: dbId
      };
    }
  } catch (error) {
    return {
      isSecure: false,
      message: 'セキュリティ検証中にエラーが発生しました: ' + error.message
    };
  }
}

/**
 * データベースセキュリティの状態を管理者向けに表示
 * @returns {string} セキュリティ状態レポート
 */
function getDatabaseSecurityReport() {
  const validation = validateDatabaseSecurity();
  
  let report = '=== データベースセキュリティ状態 ===\n';
  report += '状態: ' + (validation.isSecure ? '✅ 安全' : '❌ 要注意') + '\n';
  report += 'メッセージ: ' + validation.message + '\n';
  
  if (validation.serviceAccount) {
    report += 'サービスアカウント: ' + validation.serviceAccount + '\n';
  }
  
  if (validation.databaseId) {
    report += 'データベースID: ' + validation.databaseId + '\n';
  }
  
  report += '\n=== セキュリティ設定 ===\n';
  report += '• データベース操作: サービスアカウント専用\n';
  report += '• 個人スプレッドシート: ユーザー権限\n';
  report += '• Google Drive操作: ユーザー権限\n';
  
  return report;
}
