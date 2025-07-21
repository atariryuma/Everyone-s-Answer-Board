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
    const dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
    
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
    const dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
    
    if (!dbId) {
      throw new Error('データベースIDが設定されていません');
    }
    
    const service = getSheetsService();
    const sheetName = DB_SHEET_CONFIG.SHEET_NAME;
    
    const data = batchGetSheetsData(service, dbId, [`'${sheetName}'!A:H`]);
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
      const dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
      
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
      const data = batchGetSheetsData(service, dbId, [`'${sheetName}'!A:H`]);
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
    const dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
    
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
 * @returns {object} Sheets APIサービス
 */
function getSheetsService() {
  try {
    console.log('🔧 getSheetsService: サービス取得開始');
    
    var accessToken;
    try {
      accessToken = getServiceAccountTokenCached();
    } catch (tokenError) {
      console.error('❌ Failed to get service account token:', tokenError.message);
      throw new Error('サービスアカウントトークンの取得に失敗しました: ' + tokenError.message);
    }

    if (!accessToken) {
      console.error('❌ Access token is null or undefined after generation.');
      throw new Error('サービスアカウントトークンが取得できませんでした。');
    }
    
    console.log('✅ Access token obtained successfully');
    
    var service = createSheetsService(accessToken);
    if (!service || !service.baseUrl) {
      console.error('❌ Failed to create sheets service or service object is invalid');
      throw new Error('Sheets APIサービスの初期化に失敗しました。');
    }
    
    console.log('✅ Sheets service created successfully');
    console.log('DEBUG: getSheetsService returning:', JSON.stringify(service, null, 2));
    return service;
    
  } catch (error) {
    console.error('❌ getSheetsService error:', error.message);
    console.error('❌ Error stack:', error.stack);
    throw error; // エラーを再スロー
  }
}

/**
 * ユーザー情報を効率的に検索（キャッシュ優先）
 * @param {string} userId - ユーザーID
 * @returns {object|null} ユーザー情報
 */
function findUserById(userId) {
  var cacheKey = 'user_' + userId;
  return cacheManager.get(
    cacheKey,
    function() { return fetchUserFromDatabase('userId', userId); },
    { ttl: 300, enableMemoization: true }
  );
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
 * データベースからユーザーを取得
 * @param {string} field - 検索フィールド
 * @param {string} value - 検索値
 * @returns {object|null} ユーザー情報
 */
function fetchUserFromDatabase(field, value) {
  try {
    var props = PropertiesService.getScriptProperties();
    var dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
    
    if (!dbId) {
      console.error('fetchUserFromDatabase: データベースIDが設定されていません');
      return null;
    }
    
    var service = getSheetsService();
    var sheetName = DB_SHEET_CONFIG.SHEET_NAME;
    
    console.log('fetchUserFromDatabase - 検索開始:', {
      field: field,
      value: value,
      dbId: dbId,
      sheetName: sheetName
    });
    
    var data = batchGetSheetsData(service, dbId, ["'" + sheetName + "'!A:H"]);
    var values = data.valueRanges[0].values || [];
    
    console.log('fetchUserFromDatabase - スプレッドシートデータ取得完了:', {
      totalRows: values.length,
      hasHeaders: values.length > 0,
      headers: values.length > 0 ? values[0] : 'なし'
    });
    
    if (values.length === 0) {
      console.warn('fetchUserFromDatabase: データが見つかりません');
      return null;
    }
    
    var headers = values[0];
    var fieldIndex = headers.indexOf(field);
    
    if (fieldIndex === -1) {
      console.error('fetchUserFromDatabase: 指定されたフィールドが見つかりません:', {
        field: field,
        availableHeaders: headers
      });
      return null;
    }
    
    console.log('fetchUserFromDatabase - フィールド検索開始:', {
      fieldIndex: fieldIndex,
      searchValue: value,
      searchValueType: typeof value
    });
    
    for (var i = 1; i < values.length; i++) {
      var currentRow = values[i];
      var currentValue = currentRow[fieldIndex];
      
      // 値の比較を厳密に行う（文字列の trim と型変換）
      var normalizedCurrentValue = currentValue ? String(currentValue).trim() : '';
      var normalizedSearchValue = value ? String(value).trim() : '';
      
      console.log('fetchUserFromDatabase - 行比較:', {
        rowIndex: i,
        currentValue: currentValue,
        normalizedCurrentValue: normalizedCurrentValue,
        searchValue: value,
        normalizedSearchValue: normalizedSearchValue,
        isMatch: normalizedCurrentValue === normalizedSearchValue
      });
      
      if (normalizedCurrentValue === normalizedSearchValue) {
        var user = {};
        headers.forEach(function(header, index) {
          var rawValue = currentRow[index];
          // 空文字の場合は undefined ではなく空文字を保持
          user[header] = rawValue !== undefined && rawValue !== null ? rawValue : '';
        });
        
        // isActive フィールドの型変換を確実に行う
        if (user.hasOwnProperty('isActive')) {
          if (user.isActive === true || user.isActive === 'true' || user.isActive === 'TRUE') {
            user.isActive = true;
          } else if (user.isActive === false || user.isActive === 'false' || user.isActive === 'FALSE') {
            user.isActive = false;
          } else {
            // デフォルトでアクティブとする
            user.isActive = true;
          }
        }
        
        // デバッグログでマッピングを確認
        console.log('fetchUserFromDatabase - ユーザー発見:', {
          field: field,
          value: value,
          userId: user.userId,
          adminEmail: user.adminEmail,
          createdAt: user.createdAt,
          isActive: user.isActive,
          isActiveType: typeof user.isActive,
          configJson: user.configJson ? '設定あり(' + String(user.configJson).length + '文字)' : '設定なし',
          allFields: Object.keys(user),
          rawRowData: currentRow
        });
        
        return user;
      }
    }
    
    console.warn('fetchUserFromDatabase - ユーザーが見つかりません:', {
      field: field,
      value: value,
      totalSearchedRows: values.length - 1
    });
    
    return null;
  } catch (error) {
    console.error('fetchUserFromDatabase - エラー発生 (' + field + ':' + value + '):', error.message, error.stack);
    return null;
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
  var transactionLog = [];
  var rollbackData = null;
  
  try {
    console.log('updateUser: トランザクション開始 - userId:', userId);
    transactionLog.push('Transaction started for userId: ' + userId);
    
    // Pre-flight checks with database health validation
    var dbHealthCheck = testAndRepairDatabaseConnection();
    if (!dbHealthCheck.isHealthy) {
      throw new Error('データベース接続エラー: ' + dbHealthCheck.issues.join(', '));
    }
    
    if (dbHealthCheck.repaired) {
      console.log('updateUser: データベース接続を修復しました');
      transactionLog.push('Database connection repaired');
    }
    
    var props = PropertiesService.getScriptProperties();
    var dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
    var service = getSheetsService();
    var sheetName = DB_SHEET_CONFIG.SHEET_NAME;
    
    // Validate input data
    if (!userId || typeof userId !== 'string') {
      throw new Error('無効なuserID: ' + userId);
    }
    
    if (!updateData || typeof updateData !== 'object') {
      throw new Error('無効な更新データ: ' + typeof updateData);
    }
    
    transactionLog.push('Input validation passed');
    
    // Get current data with retry logic
    var data, values;
    var maxRetries = 3;
    var retryDelay = 1000;
    
    for (var retry = 0; retry < maxRetries; retry++) {
      try {
        data = batchGetSheetsData(service, dbId, ["'" + sheetName + "'!A:H"]);
        values = data.valueRanges[0].values || [];
        break;
      } catch (dataError) {
        console.warn('updateUser: データ取得試行 ' + (retry + 1) + ' 失敗:', dataError.message);
        if (retry === maxRetries - 1) {
          throw new Error('データベースからの情報取得に失敗しました: ' + dataError.message);
        }
        Utilities.sleep(retryDelay);
        retryDelay *= 2;
      }
    }
    
    if (values.length === 0) {
      throw new Error('データベースが空です');
    }
    
    transactionLog.push('Current data retrieved successfully');
    
    var headers = values[0];
    var userIdIndex = headers.indexOf('userId');
    var rowIndex = -1;
    var currentRowData = null;
    
    // Find user row and store current data for rollback
    for (var i = 1; i < values.length; i++) {
      if (values[i][userIdIndex] === userId) {
        rowIndex = i + 1; // 1-based index
        currentRowData = values[i].slice(); // Create copy for rollback
        break;
      }
    }
    
    if (rowIndex === -1) {
      throw new Error('更新対象のユーザーが見つかりません: ' + userId);
    }
    
    // Store rollback data
    rollbackData = {
      rowIndex: rowIndex,
      headers: headers,
      originalData: currentRowData
    };
    
    transactionLog.push('User row found at index: ' + rowIndex);
    
    // Create batch update requests with validation
    var requests = [];
    var validUpdateFields = [];
    
    Object.keys(updateData).forEach(function(key) {
      var colIndex = headers.indexOf(key);
      if (colIndex === -1) {
        console.warn('updateUser: 不明な列をスキップ:', key);
        return;
      }
      
      // Validate data types and constraints
      var value = updateData[key];
      if (key === 'configJson' && value) {
        try {
          JSON.parse(value); // Validate JSON format
        } catch (jsonError) {
          throw new Error('無効なJSON形式のconfigJson: ' + jsonError.message);
        }
      }
      
      requests.push({
        range: "'" + sheetName + "'!" + String.fromCharCode(65 + colIndex) + rowIndex,
        values: [[value]]
      });
      
      validUpdateFields.push(key);
    });
    
    if (requests.length === 0) {
      console.warn('updateUser: 更新可能なフィールドがありません');
      return { success: true, message: '更新するフィールドがありませんでした', fieldsUpdated: [] };
    }
    
    transactionLog.push('Prepared ' + requests.length + ' update requests for fields: ' + validUpdateFields.join(', '));
    
    // Execute batch update with retry and verification
    for (var updateRetry = 0; updateRetry < maxRetries; updateRetry++) {
      try {
        batchUpdateSheetsData(service, dbId, requests);
        console.log('updateUser: バッチ更新完了 (試行 ' + (updateRetry + 1) + ')');
        transactionLog.push('Batch update completed on attempt ' + (updateRetry + 1));
        break;
      } catch (updateError) {
        console.warn('updateUser: バッチ更新試行 ' + (updateRetry + 1) + ' 失敗:', updateError.message);
        
        if (updateRetry === maxRetries - 1) {
          throw new Error('データベース更新に失敗しました: ' + updateError.message);
        }
        
        Utilities.sleep(retryDelay);
      }
    }
    
    // Handle spreadsheet sharing if needed
    if (updateData.spreadsheetId) {
      try {
        shareSpreadsheetWithServiceAccount(updateData.spreadsheetId);
        console.log('updateUser: サービスアカウント共有完了:', updateData.spreadsheetId);
        transactionLog.push('Service account sharing completed for: ' + updateData.spreadsheetId);
      } catch (shareError) {
        console.error('updateUser: サービスアカウント共有エラー:', shareError.message);
        transactionLog.push('Service account sharing failed: ' + shareError.message);
        // Note: This is not a critical failure for the user update transaction
        console.warn('スプレッドシートの更新は完了しましたが、サービスアカウントとの共有に失敗しました。手動で共有してください。');
      }
    }
    
    // Verify update success by reading back the data
    try {
      Utilities.sleep(200); // Brief pause for consistency
      var verificationData = batchGetSheetsData(service, dbId, ["'" + sheetName + "'!" + String.fromCharCode(65 + userIdIndex) + rowIndex]);
      var verifiedUserId = verificationData.valueRanges[0].values[0][0];
      
      if (verifiedUserId !== userId) {
        throw new Error('データ整合性エラー: 更新確認に失敗しました');
      }
      
      transactionLog.push('Update verification successful');
    } catch (verificationError) {
      console.warn('updateUser: 更新確認で警告:', verificationError.message);
      // Continue as update likely succeeded
    }
    
    // Cache invalidation (optimized based on changed fields)
    var userInfo = findUserById(userId);
    var email = updateData.adminEmail || (userInfo ? userInfo.adminEmail : null);
    var spreadsheetId = updateData.spreadsheetId || (userInfo ? userInfo.spreadsheetId : null);
    
    invalidateUserCache(userId, email, spreadsheetId, false);
    transactionLog.push('Cache invalidation completed');
    
    console.log('updateUser: トランザクション完了 - フィールド:', validUpdateFields.join(', '));
    
    return { 
      success: true, 
      message: 'ユーザー情報を正常に更新しました',
      fieldsUpdated: validUpdateFields,
      transactionLog: transactionLog
    };
    
  } catch (error) {
    console.error('updateUser: トランザクションエラー:', error.message);
    transactionLog.push('Transaction failed: ' + error.message);
    
    // Attempt rollback if we have the necessary data
    if (rollbackData && rollbackData.originalData) {
      try {
        console.log('updateUser: ロールバック試行中...');
        
        var rollbackRequests = [];
        for (var i = 0; i < rollbackData.headers.length && i < rollbackData.originalData.length; i++) {
          if (rollbackData.originalData[i] !== undefined) {
            rollbackRequests.push({
              range: "'" + sheetName + "'!" + String.fromCharCode(65 + i) + rollbackData.rowIndex,
              values: [[rollbackData.originalData[i]]]
            });
          }
        }
        
        if (rollbackRequests.length > 0) {
          batchUpdateSheetsData(service, dbId, rollbackRequests);
          console.log('updateUser: ロールバック完了');
          transactionLog.push('Rollback completed successfully');
        }
        
      } catch (rollbackError) {
        console.error('updateUser: ロールバック失敗:', rollbackError.message);
        transactionLog.push('Rollback failed: ' + rollbackError.message);
        // Add rollback failure to the original error message
        error.message += ' (ロールバックも失敗: ' + rollbackError.message + ')';
      }
    }
    
    // Enhanced error reporting
    var enhancedError = new Error('ユーザー更新に失敗しました: ' + error.message);
    enhancedError.transactionLog = transactionLog;
    enhancedError.originalError = error;
    enhancedError.rollbackData = rollbackData;
    
    throw enhancedError;
  }
}

/**
 * 新規ユーザーをデータベースに作成
 * @param {object} userData - 作成するユーザーデータ
 * @returns {object} 作成されたユーザーデータ
 */
function createUser(userData) {
  // 同時登録による重複を防ぐためロックを取得
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    // メールアドレスの重複チェック
    var existingUser = findUserByEmail(userData.adminEmail);
    if (existingUser) {
      throw new Error('このメールアドレスは既に登録されています。');
    }

    var props = PropertiesService.getScriptProperties();
    var dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
    var service = getSheetsService();
    var sheetName = DB_SHEET_CONFIG.SHEET_NAME;

    var newRow = DB_SHEET_CONFIG.HEADERS.map(function(header) {
      return userData[header] || '';
    });
  
    appendSheetsData(service, dbId, "'" + sheetName + "'!A1", [newRow]);

    // 最適化: 新規ユーザー作成時は対象キャッシュのみ無効化
    invalidateUserCache(userData.userId, userData.adminEmail, userData.spreadsheetId, false);

    return userData;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Polls the database until a user record becomes available.
 * @param {string} userId - The ID of the user to fetch.
 * @param {number} maxWaitMs - Maximum wait time in milliseconds.
 * @param {number} intervalMs - Poll interval in milliseconds.
 * @returns {boolean} true if found within the wait window.
 */
function waitForUserRecord(userId, maxWaitMs, intervalMs) {
  var start = Date.now();
  while (Date.now() - start < maxWaitMs) {
    if (fetchUserFromDatabase('userId', userId)) return true;
    Utilities.sleep(intervalMs);
  }
  return false;
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
 * バッチ取得
 * @param {object} service - Sheetsサービス
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string[]} ranges - 取得範囲の配列
 * @returns {object} レスポンス
 */
function batchGetSheetsData(service, spreadsheetId, ranges) {
  console.log('DEBUG: batchGetSheetsData received service:', JSON.stringify(service, null, 2));
  // API呼び出しをキャッシュ化（短期間）
  var cacheKey = `batchGet_${spreadsheetId}_${JSON.stringify(ranges)}`;
  
  return cacheManager.get(cacheKey, () => {
    const currentService = service; // Capture service in local variable
    try {
      console.log('DEBUG: Accessing service.baseUrl. Service object:', JSON.stringify(currentService, null, 2));
      var url = currentService.baseUrl + '/' + spreadsheetId + '/values:batchGet?' + 
        ranges.map(function(range) { return 'ranges=' + encodeURIComponent(range); }).join('&');
      
      var response = UrlFetchApp.fetch(url, {
        headers: { 'Authorization': 'Bearer ' + service.accessToken },
        muteHttpExceptions: true,
        followRedirects: true,
        validateHttpsCertificates: true
      });
      
      if (response.getResponseCode() !== 200) {
        throw new Error('Sheets API error: ' + response.getResponseCode() + ' - ' + response.getContentText());
      }
      
      return JSON.parse(response.getContentText());
    } catch (error) {
      console.error('batchGetSheetsData error:', error.message);
      throw new Error('データ取得に失敗しました: ' + error.message);
    }
  }, { ttl: 120 }); // 2分間キャッシュ（API制限対策）
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
    if (!service || !service.baseUrl) {
      throw new Error('Sheets APIサービスオブジェクトが無効です。baseUrlが見つかりません。');
    }
    // シート情報を含む基本的なメタデータを取得するために fields パラメータを追加
    var baseUrl = service.baseUrl; // Store baseUrl in a local variable
    var url = baseUrl + '/' + spreadsheetId + '?fields=sheets.properties';
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
 * データベース接続の健全性をテストし、必要に応じて修復を試みる
 * @returns {object} {isHealthy: boolean, issues: array, repaired: boolean, dbInfo: object}
 */
function testAndRepairDatabaseConnection() {
  var result = {
    isHealthy: false,
    issues: [],
    repaired: false,
    dbInfo: null
  };
  
  try {
    console.log('testAndRepairDatabaseConnection: データベース健全性チェック開始');
    
    // Step 1: Check if database spreadsheet ID is configured
    var props = PropertiesService.getScriptProperties();
    var dbSpreadsheetId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
    
    if (!dbSpreadsheetId) {
      result.issues.push('データベースのスプレッドシートIDが設定されていません');
      return result;
    }
    
    result.dbInfo = {
      spreadsheetId: dbSpreadsheetId
    };
    
    // Step 2: Validate and repair service account
    var serviceAccountValidation = validateAndRepairServiceAccount();
    if (!serviceAccountValidation.isValid) {
      result.issues.push('サービスアカウント設定エラー: ' + serviceAccountValidation.issues.join(', '));
      return result;
    }
    
    if (serviceAccountValidation.repaired) {
      result.repaired = true;
      console.log('testAndRepairDatabaseConnection: サービスアカウントを修復しました');
    }
    
    // Step 3: Test basic spreadsheet access
    var maxRetries = 3;
    var retryDelay = 1000;
    var accessTest = false;
    
    for (var retry = 0; retry < maxRetries; retry++) {
      try {
        var file = DriveApp.getFileById(dbSpreadsheetId);
        if (!file) {
          result.issues.push('データベーススプレッドシートが見つかりません: ' + dbSpreadsheetId);
          return result;
        }
        
        // Try to access basic file properties
        result.dbInfo.name = file.getName();
        result.dbInfo.url = file.getUrl();
        accessTest = true;
        
        console.log('testAndRepairDatabaseConnection: スプレッドシートアクセス成功 (試行 ' + (retry + 1) + ')');
        break;
        
      } catch (accessError) {
        console.warn('testAndRepairDatabaseConnection: アクセス試行 ' + (retry + 1) + ' 失敗:', accessError.message);
        
        if (retry === maxRetries - 1) {
          result.issues.push('データベーススプレッドシートへのアクセスに失敗: ' + accessError.message);
          return result;
        }
        
        Utilities.sleep(retryDelay);
        retryDelay *= 2;
      }
    }
    
    if (!accessTest) {
      result.issues.push('データベーススプレッドシートアクセステストに失敗しました');
      return result;
    }
    
    // Step 4: Test Sheets API connection
    try {
      var service = getSheetsService();
      if (!service) {
        result.issues.push('Sheets APIサービスの初期化に失敗しました');
        return result;
      }
      
      // Simple API test - try to get spreadsheet metadata
      var request = {
        'url': 'https://sheets.googleapis.com/v4/spreadsheets/' + dbSpreadsheetId + '?fields=sheets.properties.title',
        'method': 'GET',
        'headers': {
          'Authorization': 'Bearer ' + service.getAccessToken()
        }
      };
      
      var response = UrlFetchApp.fetch(request.url, {
        method: request.method,
        headers: request.headers
      });
      
      if (response.getResponseCode() !== 200) {
        if (response.getResponseCode() === 403) {
          // Try to repair permissions
          console.log('testAndRepairDatabaseConnection: 権限エラー検出、修復を試行中...');
          try {
            shareSpreadsheetWithServiceAccount(dbSpreadsheetId);
            result.repaired = true;
            console.log('testAndRepairDatabaseConnection: データベース権限を修復しました');
            
            // Retry the API call
            var retryResponse = UrlFetchApp.fetch(request.url, {
              method: request.method,
              headers: request.headers
            });
            
            if (retryResponse.getResponseCode() !== 200) {
              result.issues.push('権限修復後もAPI アクセスに失敗: ' + retryResponse.getResponseCode());
              return result;
            }
          } catch (repairError) {
            result.issues.push('データベース権限修復に失敗: ' + repairError.message);
            return result;
          }
        } else {
          result.issues.push('Sheets API エラー: ' + response.getResponseCode() + ' - ' + response.getContentText());
          return result;
        }
      }
      
      var apiResult = JSON.parse(response.getContentText());
      result.dbInfo.sheets = apiResult.sheets || [];
      
    } catch (apiError) {
      result.issues.push('Sheets API テストに失敗: ' + apiError.message);
      return result;
    }
    
    // Step 5: Test basic read operation
    try {
      var testData = getAllUsers();
      result.dbInfo.userCount = testData ? testData.length : 0;
      console.log('testAndRepairDatabaseConnection: データベース読み込みテスト成功 - ユーザー数:', result.dbInfo.userCount);
    } catch (readError) {
      console.warn('testAndRepairDatabaseConnection: データ読み込みテストで警告:', readError.message);
      // This is not a critical failure, continue
    }
    
    result.isHealthy = true;
    console.log('testAndRepairDatabaseConnection: データベース健全性チェック完了');
    
    return result;
    
  } catch (error) {
    result.issues.push('予期しないエラー: ' + error.message);
    console.error('testAndRepairDatabaseConnection: 致命的エラー:', error.message);
    return result;
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
    
    var data = batchGetSheetsData(service, dbId, ["'" + sheetName + "'!A:H"]);
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
 * データベースシートを取得
 * @returns {object} データベースシート
 */
function getDbSheet() {
  try {
    var props = PropertiesService.getScriptProperties();
    var dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
    if (!dbId) {
      throw new Error('データベースIDが設定されていません');
    }
    
    var ss = SpreadsheetApp.openById(dbId);
    var sheet = ss.getSheetByName(DB_SHEET_CONFIG.SHEET_NAME);
    if (!sheet) {
      throw new Error('データベースシートが見つかりません: ' + DB_SHEET_CONFIG.SHEET_NAME);
    }
    
    return sheet;
  } catch (e) {
    console.error('getDbSheet error:', e.message);
    throw new Error('データベースシートの取得に失敗しました: ' + e.message);
  }
}

/**
 * 現在のユーザーのアカウント情報をデータベースから完全に削除する。
 * Google Drive 上の関連ファイルやフォルダは保持したままにする。
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
      var data = batchGetSheetsData(service, dbId, ["'" + sheetName + "'!A:H"]);
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

      // Google Drive のデータは保持するため何も操作しない
      
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
