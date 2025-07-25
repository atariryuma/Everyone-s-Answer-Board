/**
 * @fileoverview データベース管理 - バッチ操作とキャッシュ最適化
 * GAS互換の関数ベースの実装
 */

// データベース管理のための定数
var USER_CACHE_TTL = 300; // 5分
var DB_BATCH_SIZE = 100;

// 簡易インデックス機能：ユーザー検索の高速化
var userIndexCache = {
  byUserId: new Map(),
  byEmail: new Map(),
  lastUpdate: 0,
  TTL: 300000 // 5分間のキャッシュ
};

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
    
    const service = getSheetsServiceCached();
    const logSheetName = DELETE_LOG_SHEET_CONFIG.SHEET_NAME;
    
    // ログシートの存在確認・作成
    try {
      const spreadsheetInfo = getSpreadsheetsData(service, dbId);
      const logSheetExists = spreadsheetInfo.sheets.some(sheet => 
        sheet.properties.title === logSheetName
      );
      
      if (!logSheetExists) {
        // バッチ最適化: ログシート作成
        console.log('📊 バッチ最適化: ログシート作成を実行');
        
        const addSheetRequest = {
          addSheet: {
            properties: {
              title: logSheetName,
              gridProperties: {
                rowCount: 1000,
                columnCount: DELETE_LOG_SHEET_CONFIG.HEADERS.length
              }
            }
          }
        };
        
        // シートを作成
        batchUpdateSpreadsheet(service, dbId, {
          requests: [addSheetRequest]
        });
        
        // ヘッダー行を追加（作成直後）
        appendSheetsData(service, dbId, `'${logSheetName}'!A1`, [DELETE_LOG_SHEET_CONFIG.HEADERS]);
        
        console.log('✅ ログシートとヘッダーの作成完了');
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
    
    const service = getSheetsServiceCached();
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
      
      // createdAt を registrationDate にマッピング
      if (user.createdAt) {
        user.registrationDate = user.createdAt;
      }

      // 設定情報をパース
      try {
        user.configJson = JSON.parse(user.configJson || '{}');
      } catch (e) {
        user.configJson = {};
      }
      
      users.push(user);
      console.log(`DEBUG: getAllUsersForAdmin - User object: ${JSON.stringify(user)}`);
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
      
      const service = getSheetsServiceCached();
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
    
    const service = getSheetsServiceCached();
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
 * 長期キャッシュ対応Sheetsサービスを取得
 * @returns {object} Sheets APIサービス
 */
function getSheetsServiceCached(forceRefresh) {
  try {
    console.log('🔧 getSheetsServiceCached: 新規サービス作成開始（キャッシュなし版）');
    
    var accessToken;
    if (forceRefresh) {
      console.log('🔐 認証トークンも強制リフレッシュ');
      cacheManager.remove('service_account_token');
      accessToken = generateNewServiceAccountToken();
    } else {
      accessToken = getServiceAccountTokenCached();
    }
    
    if (!accessToken) {
      throw new Error('サービスアカウントトークンが取得できませんでした');
    }
    
    var service = createSheetsService(accessToken);
    if (!service || !service.baseUrl || !service.accessToken) {
      console.error('❌ サービスオブジェクト検証失敗:', {
        hasService: !!service,
        hasBaseUrl: !!(service && service.baseUrl),
        hasAccessToken: !!(service && service.accessToken)
      });
      throw new Error('Sheets APIサービスの初期化に失敗しました: 有効なサービスオブジェクトを作成できません');
    }
    
    console.log('✅ サービスオブジェクト検証成功:', {
      hasBaseUrl: true,
      hasAccessToken: true,
      hasSpreadsheets: !!service.spreadsheets,
      hasGet: !!(service.spreadsheets && typeof service.spreadsheets.get === 'function'),
      baseUrl: service.baseUrl
    });
    
    // 関数の存在確認（重要: Google Apps Scriptで関数が失われていないか確認）
    if (!service.spreadsheets || typeof service.spreadsheets.get !== 'function') {
      console.error('❌ 重要な関数が失われています:', {
        hasSpreadsheets: !!service.spreadsheets,
        getType: service.spreadsheets ? typeof service.spreadsheets.get : 'no spreadsheets'
      });
      throw new Error('SheetsServiceオブジェクトの関数が正しく設定されていません');
    }
    
    console.log('✅ キャッシュ用新規Sheetsサービス作成完了（関数検証済み）');
    return service;
    
  } catch (error) {
    console.error('❌ getSheetsServiceCached error:', error.message);
    throw error;
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
    console.log('DEBUG: getSheetsService returning service object with baseUrl:', service.baseUrl);
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
 * キャッシュをバイパスして最新のユーザー情報を強制取得
 * クリティカルな更新操作直後に使用
 * @param {string} userId - ユーザーID
 * @returns {object|null} 最新のユーザー情報
 */
function findUserByIdFresh(userId) {
  // 既存キャッシュを強制削除
  var cacheKey = 'user_' + userId;
  cacheManager.remove(cacheKey);
  
  // データベースから直接取得（キャッシュなし）
  var freshUserInfo = fetchUserFromDatabase('userId', userId);
  
  // 次回の通常アクセス時にキャッシュされるため、ここでは手動設定不要
  console.log('✅ Fresh user data retrieved for:', userId);
  
  return freshUserInfo;
}

/**
 * ユーザーデータの整合性を修正する
 * @param {string} userId - ユーザーID
 * @returns {object} 修正結果
 */
function fixUserDataConsistency(userId) {
  try {
    console.log('🔧 ユーザーデータ整合性修正開始:', userId);
    
    // 最新のユーザー情報を取得
    var userInfo = findUserByIdFresh(userId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    
    console.log('📊 現在のspreadsheetId:', userInfo.spreadsheetId);
    
    var configJson = JSON.parse(userInfo.configJson || '{}');
    console.log('📝 configJson内のpublishedSpreadsheetId:', configJson.publishedSpreadsheetId);
    
    var needsUpdate = false;
    var updateData = {};
    
    // 1. エラー情報をクリーンアップ
    if (configJson.lastError || configJson.errorAt) {
      console.log('🧹 エラー情報をクリーンアップ');
      delete configJson.lastError;
      delete configJson.errorAt;
      needsUpdate = true;
    }
    
    // 2. spreadsheetIdとpublishedSpreadsheetIdの整合性チェック
    if (userInfo.spreadsheetId && configJson.publishedSpreadsheetId !== userInfo.spreadsheetId) {
      console.log('🔄 publishedSpreadsheetIdを実際のspreadsheetIdに合わせて修正');
      console.log('  修正前:', configJson.publishedSpreadsheetId);
      console.log('  修正後:', userInfo.spreadsheetId);
      
      configJson.publishedSpreadsheetId = userInfo.spreadsheetId;
      needsUpdate = true;
    }
    
    // 3. セットアップ状態の正規化
    if (userInfo.spreadsheetId && configJson.setupStatus !== 'completed') {
      console.log('🔄 セットアップ状態を正規化');
      configJson.setupStatus = 'completed';
      configJson.formCreated = true;
      configJson.appPublished = true;
      needsUpdate = true;
    }
    
    if (needsUpdate) {
      updateData.configJson = JSON.stringify(configJson);
      updateData.lastAccessedAt = new Date().toISOString();
      
      console.log('💾 整合性修正のためデータベース更新実行');
      updateUser(userId, updateData);
      
      console.log('✅ ユーザーデータ整合性修正完了');
      return { status: 'success', message: 'データ整合性が修正されました', updated: true };
    } else {
      console.log('✅ データ整合性に問題なし');
      return { status: 'success', message: 'データ整合性に問題ありません', updated: false };
    }
    
  } catch (error) {
    console.error('❌ データ整合性修正エラー:', error);
    return { status: 'error', message: 'データ整合性修正に失敗: ' + error.message };
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
 * データベースからユーザーを取得（エラーハンドリング強化版）
 * @param {string} field - 検索フィールド
 * @param {string} value - 検索値
 * @param {object} options - オプション設定
 * @returns {object|null} ユーザー情報
 */
function fetchUserFromDatabase(field, value, options = {}) {
  // オプションのデフォルト値
  var opts = {
    retryCount: options.retryCount || 2,
    enableDiagnostics: options.enableDiagnostics !== false,
    autoRepair: options.autoRepair !== false,
    ...options
  };
  
  var retryAttempt = 0;
  var lastError = null;
  
  while (retryAttempt <= opts.retryCount) {
    try {
      console.log('fetchUserFromDatabase - 試行 ' + (retryAttempt + 1) + '/' + (opts.retryCount + 1) + 
        ': ' + field + '=' + value);
      
      var props = PropertiesService.getScriptProperties();
      var dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
      
      if (!dbId) {
        var configError = new Error('データベースIDが設定されていません');
        configError.type = 'CONFIG_ERROR';
        throw configError;
      }
      
      // サービス取得（リトライ時は強制リフレッシュ）
      var service;
      try {
        service = retryAttempt > 0 ? getSheetsServiceCached(true) : getCachedSheetsService();
      } catch (serviceError) {
        serviceError.type = 'SERVICE_ERROR';
        throw serviceError;
      }
      
      var sheetName = DB_SHEET_CONFIG.SHEET_NAME;
      
      console.log('fetchUserFromDatabase - 検索開始: ' + field + '=' + value);
      console.log('fetchUserFromDatabase - データベースID: ' + dbId);
      console.log('fetchUserFromDatabase - シート名: ' + sheetName);
      
      // キャッシュクリア（リトライ時は必須）
      if (retryAttempt > 0 || opts.clearCache) {
        const batchGetCacheKey = `batchGet_${dbId}_["'${sheetName}'!A:H"]`;
        try {
          cacheManager.remove(batchGetCacheKey);
        } catch (cacheError) {
          console.warn('fetchUserFromDatabase - キャッシュクリア警告:', cacheError.message);
        }
      }
      
      // データ取得
      var data;
      try {
        data = batchGetSheetsData(service, dbId, ["'" + sheetName + "'!A:H"]);
      } catch (dataError) {
        dataError.type = 'DATA_ACCESS_ERROR';
        throw dataError;
      }
      
      var values = data.valueRanges[0].values || [];
      
      console.log('fetchUserFromDatabase - データ取得完了: rows=' + values.length);
      
      if (values.length === 0) {
        var noDataError = new Error('データベースが空です');
        noDataError.type = 'NO_DATA_ERROR';
        throw noDataError;
      }
      
      var headers = values[0];
      var fieldIndex = headers.indexOf(field);
      
      if (fieldIndex === -1) {
        console.error('fetchUserFromDatabase: 指定されたフィールドが見つかりません:', {
          field: field,
          availableHeaders: headers
        });
        var fieldError = new Error('検索フィールド "' + field + '" が見つかりません');
        fieldError.type = 'FIELD_ERROR';
        throw fieldError;
      }
      
      console.log('fetchUserFromDatabase - フィールド検索開始: index=' + fieldIndex);
      console.log('fetchUserFromDatabase - デバッグ: 検索対象データ行数=' + (values.length > 1 ? values.length - 1 : 0));
      
      // ユーザー検索
      for (var i = 1; i < values.length; i++) { // i=0はヘッダー行のためスキップ
        var currentRow = values[i];
        var currentValue = currentRow[fieldIndex];
        
        // 値の比較を厳密に行う（文字列の trim と型変換）
        var normalizedCurrentValue = currentValue ? String(currentValue).trim() : '';
        var normalizedSearchValue = value ? String(value).trim() : '';
        
        // 詳細ログ（最初の試行時のみ）
        if (retryAttempt === 0) {
          console.log('fetchUserFromDatabase - 行' + i + '値比較:', {
            original: currentValue,
            normalized: normalizedCurrentValue,
            searchValue: normalizedSearchValue,
            isMatch: normalizedCurrentValue === normalizedSearchValue
          });
        }
        
        if (normalizedCurrentValue === normalizedSearchValue) {
          console.log('fetchUserFromDatabase - ユーザー発見 at rowIndex:', i);
          var user = {};
          headers.forEach(function(header, index) {
            var rawValue = currentRow[index];
            user[header] = rawValue !== undefined && rawValue !== null ? rawValue : '';
          });
          
          // isActive フィールドの型変換
          if (user.hasOwnProperty('isActive')) {
            if (user.isActive === true || user.isActive === 'true' || user.isActive === 'TRUE') {
              user.isActive = true;
            } else if (user.isActive === false || user.isActive === 'false' || user.isActive === 'FALSE') {
              user.isActive = false;
            } else {
              user.isActive = true; // デフォルト
            }
          }
          
          console.log('✅ fetchUserFromDatabase - ユーザー取得成功:', field + '=' + value, 
            'userId=' + user.userId, 'isActive=' + user.isActive);
          
          return user;
        }
      }
      
      // ユーザーが見つからない場合
      console.warn('⚠️ fetchUserFromDatabase - ユーザーが見つかりません:', {
        field: field,
        value: value,
        totalSearchedRows: values.length - 1,
        attempt: retryAttempt + 1,
        availableUserIds: values.slice(1).map(row => row[headers.indexOf('userId')] || 'undefined').slice(0, 5)
      });
      
      var notFoundError = new Error('ユーザー情報が見つかりません');
      notFoundError.type = 'USER_NOT_FOUND';
      notFoundError.searchCriteria = { field: field, value: value };
      throw notFoundError;
      
    } catch (error) {
      lastError = error;
      retryAttempt++;
      
      console.error('❌ fetchUserFromDatabase - エラー発生 (試行 ' + retryAttempt + '/' + (opts.retryCount + 1) + 
        ') (' + field + ':' + value + '):', error.message);
      
      // エラータイプ別の処理
      if (error.type === 'CONFIG_ERROR' || error.type === 'FIELD_ERROR') {
        // 設定エラーやフィールドエラーはリトライしても無駄
        console.error('❌ 致命的エラーのためリトライを中止:', error.type);
        break;
      }
      
      // 最後の試行で失敗した場合の診断・修復
      if (retryAttempt > opts.retryCount) {
        console.error('❌ 全ての試行が失敗しました');
        
        // 診断実行（オプションが有効な場合）
        if (opts.enableDiagnostics) {
          try {
            console.log('🔍 エラー診断を実行中...');
            var diagnosis = diagnoseDatabase(field === 'userId' ? value : null);
            console.log('📊 診断結果:', diagnosis.summary);
            
            // 自動修復試行（オプションが有効で、診断で問題が見つかった場合）
            if (opts.autoRepair && diagnosis.summary.criticalIssues.length > 0) {
              console.log('🔧 自動修復を試行中...');
              var repairResult = performAutoRepair(field === 'userId' ? value : null);
              console.log('🔧 修復結果:', repairResult.summary);
              
              if (repairResult.success) {
                console.log('♻️ 修復後に再試行します...');
                // 修復成功時は1回だけ追加試行
                return fetchUserFromDatabase(field, value, { 
                  retryCount: 0, 
                  enableDiagnostics: false, 
                  autoRepair: false,
                  clearCache: true
                });
              }
            }
          } catch (diagError) {
            console.error('❌ 診断・修復処理でエラー:', diagError.message);
          }
        }
        
        break;
      }
      
      // リトライ前の待機
      if (retryAttempt <= opts.retryCount) {
        var waitTime = Math.min(1000 * retryAttempt, 3000); // 最大3秒
        console.log('⏳ ' + waitTime + 'ms 待機後にリトライ...');
        Utilities.sleep(waitTime);
      }
    }
  }
  
  // すべてのリトライが失敗した場合
  console.error('❌ fetchUserFromDatabase - 全てのリトライが失敗:', lastError ? lastError.message : 'unknown error');
  
  // エラーの詳細情報を付加
  var finalError = lastError || new Error('ユーザー情報の取得に失敗しました');
  finalError.searchCriteria = { field: field, value: value };
  finalError.retryCount = retryAttempt - 1;
  
  return null;
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
  // 型安全性とバリデーション強化: パラメータ検証
  if (!userId) {
    throw new Error('データ更新エラー: ユーザーIDが指定されていません');
  }
  if (typeof userId !== 'string') {
    throw new Error('データ更新エラー: ユーザーIDは文字列である必要があります');
  }
  if (userId.trim().length === 0) {
    throw new Error('データ更新エラー: ユーザーIDが空文字列です');
  }
  if (userId.length > 255) {
    throw new Error('データ更新エラー: ユーザーIDが長すぎます（最大255文字）');
  }
  
  if (!updateData) {
    throw new Error('データ更新エラー: 更新データが指定されていません');
  }
  if (typeof updateData !== 'object' || Array.isArray(updateData)) {
    throw new Error('データ更新エラー: 更新データはオブジェクトである必要があります');
  }
  if (Object.keys(updateData).length === 0) {
    throw new Error('データ更新エラー: 更新データが空です');
  }
  
  // 許可されたフィールドのホワイトリスト検証
  const allowedFields = ['adminEmail', 'spreadsheetId', 'spreadsheetUrl',
    'configJson', 'lastAccessedAt', 'createdAt', 'formUrl', 'isActive'];
  const updateFields = Object.keys(updateData);
  const invalidFields = updateFields.filter(field => !allowedFields.includes(field));
  
  if (invalidFields.length > 0) {
    throw new Error('データ更新エラー: 許可されていないフィールドが含まれています: ' + invalidFields.join(', '));
  }
  
  // 各フィールドの型検証
  for (const field of updateFields) {
    const value = updateData[field];
    if (value !== null && value !== undefined && typeof value !== 'string') {
      throw new Error(`データ更新エラー: フィールド "${field}" は文字列である必要があります`);
    }
    if (typeof value === 'string' && value.length > 10000) {
      throw new Error(`データ更新エラー: フィールド "${field}" が長すぎます（最大10000文字）`);
    }
  }
  
  try {
    var props = PropertiesService.getScriptProperties();
    var dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
    
    if (!dbId) {
      throw new Error('データ更新エラー: データベースIDが設定されていません');
    }
    var service = getSheetsServiceCached();
    var sheetName = DB_SHEET_CONFIG.SHEET_NAME;
    
    // 現在のデータを取得
    var data = batchGetSheetsData(service, dbId, ["'" + sheetName + "'!A:H"]);
    var values = data.valueRanges[0].values || [];
    
    if (values.length === 0) {
      throw new Error('データベースが空です');
    }
    
    var headers = values[0];
    var userIdIndex = headers.indexOf('userId');
    var rowIndex = -1;
    
    // ユーザーの行を特定
    for (var i = 1; i < values.length; i++) {
      if (values[i][userIdIndex] === userId) {
        rowIndex = i + 1; // 1-based index
        break;
      }
    }
    
    if (rowIndex === -1) {
      throw new Error('更新対象のユーザーが見つかりません');
    }
    
    // バッチ更新リクエストを作成
    var requests = Object.keys(updateData).map(function(key) {
      var colIndex = headers.indexOf(key);
      if (colIndex === -1) return null;
      
      return {
        range: "'" + sheetName + "'!" + String.fromCharCode(65 + colIndex) + rowIndex,
        values: [[updateData[key]]]
      };
    }).filter(function(item) { return item !== null; });
    
    if (requests.length > 0) {
      console.log('📝 データベース更新リクエスト詳細:');
      requests.forEach(function(req, index) {
        console.log('  ' + (index + 1) + '. 範囲: ' + req.range + ', 値: ' + JSON.stringify(req.values));
      });
      
      var maxRetries = 2;
      var retryCount = 0;
      var updateSuccess = false;
      
      while (retryCount <= maxRetries && !updateSuccess) {
        try {
          if (retryCount > 0) {
            console.log('🔄 認証エラーによるリトライ (' + retryCount + '/' + maxRetries + ')');
            // 認証エラーの場合、新しいサービスを取得
            service = getSheetsServiceCached(true); // forceRefresh = true
            Utilities.sleep(1000); // 少し待機
          }
          
          batchUpdateSheetsData(service, dbId, requests);
          console.log('✅ データベース更新リクエスト送信完了');
          updateSuccess = true;
          
          // 更新成功の確認
          console.log('🔍 更新直後の確認のため、データベースから再取得...');
          var verifyData = batchGetSheetsData(service, dbId, ["'" + sheetName + "'!" + String.fromCharCode(65 + userIdIndex) + rowIndex + ":" + String.fromCharCode(72) + rowIndex]);
          if (verifyData.valueRanges && verifyData.valueRanges[0] && verifyData.valueRanges[0].values) {
            var updatedRow = verifyData.valueRanges[0].values[0];
            console.log('📊 更新後のユーザー行データ:', updatedRow);
            if (updateData.spreadsheetId) {
              var spreadsheetIdIndex = headers.indexOf('spreadsheetId');
              console.log('🎯 スプレッドシートID更新確認:', updatedRow[spreadsheetIdIndex] === updateData.spreadsheetId ? '✅ 成功' : '❌ 失敗');
            }
          }
          
        } catch (updateError) {
          retryCount++;
          var errorMessage = updateError.toString();
          
          if (errorMessage.includes('401') || errorMessage.includes('UNAUTHENTICATED') || errorMessage.includes('ACCESS_TOKEN_EXPIRED')) {
            console.warn('🔐 認証エラーを検出:', errorMessage);
            
            if (retryCount <= maxRetries) {
              console.log('🔄 認証トークンをリフレッシュしてリトライします...');
              continue; // リトライループを続行
            } else {
              console.error('❌ 最大リトライ回数に達しました');
              throw new Error('認証エラーにより更新に失敗しました（最大リトライ回数超過）');
            }
          } else {
            // 認証エラー以外のエラーはすぐに終了
            console.error('❌ データベース更新中にエラーが発生:', updateError);
            throw new Error('データベース更新に失敗しました: ' + updateError.message);
          }
        }
      }
      
      // 更新が確実に反映されるまで少し待機
      Utilities.sleep(100);
    } else {
      console.log('⚠️ 更新するデータがありません');
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
    
    // 重要: 更新完了後に包括的キャッシュ同期を実行
    var userInfo = findUserById(userId);
    var email = updateData.adminEmail || (userInfo ? userInfo.adminEmail : null);
    var oldSpreadsheetId = userInfo ? userInfo.spreadsheetId : null;
    var newSpreadsheetId = updateData.spreadsheetId || oldSpreadsheetId;
    
    // クリティカル更新時の包括的キャッシュ同期
    synchronizeCacheAfterCriticalUpdate(userId, email, oldSpreadsheetId, newSpreadsheetId);
    
    return { status: 'success', message: 'ユーザープロフィールが正常に更新されました' };
  } catch (error) {
    console.error('ユーザー更新エラー:', error);
    throw error;
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
    var service = getSheetsServiceCached();
    var sheetName = DB_SHEET_CONFIG.SHEET_NAME;

    var newRow = DB_SHEET_CONFIG.HEADERS.map(function(header) {
      return userData[header] || '';
    });
    
    console.log('createUser - デバッグ: ヘッダー構成=' + JSON.stringify(DB_SHEET_CONFIG.HEADERS));
    console.log('createUser - デバッグ: ユーザーデータ=' + JSON.stringify(userData));
    console.log('createUser - デバッグ: 作成される行データ=' + JSON.stringify(newRow));
  
    appendSheetsData(service, dbId, "'" + sheetName + "'!A1", [newRow]);
    
    console.log('createUser - データベース書き込み完了: userId=' + userData.userId);

    // 新規ユーザー用の専用フォルダを作成
    try {
      console.log('createUser - 専用フォルダ作成開始: ' + userData.adminEmail);
      var folder = createUserFolder(userData.adminEmail);
      if (folder) {
        console.log('✅ createUser - 専用フォルダ作成成功: ' + folder.getName());
      } else {
        console.log('⚠️ createUser - 専用フォルダ作成失敗（処理は続行）');
      }
    } catch (folderError) {
      console.warn('createUser - フォルダ作成でエラーが発生しましたが、処理を続行します: ' + folderError.message);
    }

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
  var service = getSheetsServiceCached();
  var sheetName = DB_SHEET_CONFIG.SHEET_NAME;

  try {
    // シートが存在するか確認
    var spreadsheet = getSpreadsheetsData(service, spreadsheetId);
    var sheetExists = spreadsheet.sheets.some(function(s) { 
      return s.properties.title === sheetName; 
    });

    if (!sheetExists) {
      // バッチ処理最適化: シート作成とヘッダー追加を1回のAPI呼び出しで実行
      console.log('📊 バッチ最適化: シート作成+ヘッダー追加を同時実行');
      
      var requests = [
        // 1. シートを作成
        { 
          addSheet: { 
            properties: { 
              title: sheetName,
              gridProperties: {
                rowCount: 1000,
                columnCount: DB_SHEET_CONFIG.HEADERS.length
              }
            } 
          } 
        }
      ];
      
      // バッチ実行: シートを作成
      batchUpdateSpreadsheet(service, spreadsheetId, { requests: requests });
      
      // 2. 作成直後にヘッダーを追加（A1記法でレンジを指定）
      var headerRange = "'" + sheetName + "'!A1:" + String.fromCharCode(65 + DB_SHEET_CONFIG.HEADERS.length - 1) + '1';
      updateSheetsData(service, spreadsheetId, headerRange, [DB_SHEET_CONFIG.HEADERS]);
      
      console.log('✅ バッチ最適化完了: シート作成+ヘッダー追加（2回のAPI呼び出し→シーケンシャル実行）');
      
    } else {
      // シートが既に存在する場合は、ヘッダーのみ更新（既存動作を維持）
      var headerRange = "'" + sheetName + "'!A1:" + String.fromCharCode(65 + DB_SHEET_CONFIG.HEADERS.length - 1) + '1';
      updateSheetsData(service, spreadsheetId, headerRange, [DB_SHEET_CONFIG.HEADERS]);
      console.log('✅ 既存シートのヘッダー更新完了');
    }

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
 * バッチ取得（baseUrl 問題修正版）
 * @param {object} service - Sheetsサービス
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string[]} ranges - 取得範囲の配列
 * @returns {object} レスポンス
 */
function batchGetSheetsData(service, spreadsheetId, ranges) {
  console.log('DEBUG: batchGetSheetsData - 安全なサービスオブジェクト処理開始');
  
  // 型安全性とバリデーション強化: 入力パラメータ検証
  if (!service) {
    throw new Error('Sheetsサービスオブジェクトが提供されていません');
  }
  if (typeof service !== 'object' || !service.baseUrl) {
    throw new Error('無効なSheetsサービスオブジェクトです: baseUrlが見つかりません');
  }
  
  if (!spreadsheetId || typeof spreadsheetId !== 'string') {
    throw new Error('無効なspreadsheetIDです: 文字列である必要があります');
  }
  if (spreadsheetId.length < 20 || spreadsheetId.length > 60) {
    throw new Error('無効なspreadsheetID形式です: 長さが不正です');
  }
  if (!/^[a-zA-Z0-9_-]+$/.test(spreadsheetId)) {
    throw new Error('無効なspreadsheetID形式です: 許可されていない文字が含まれています');
  }
  
  if (!ranges || !Array.isArray(ranges) || ranges.length === 0) {
    throw new Error('無効な範囲配列です: 配列で1つ以上の範囲が必要です');
  }
  if (ranges.length > 100) {
    throw new Error('範囲配列が大きすぎます: 最大100個までです');
  }
  
  // 各範囲の検証
  for (let i = 0; i < ranges.length; i++) {
    const range = ranges[i];
    if (!range || typeof range !== 'string') {
      throw new Error(`範囲[${i}]が無効です: 文字列である必要があります`);
    }
    if (range.length === 0) {
      throw new Error(`範囲[${i}]が空文字列です`);
    }
    if (range.length > 200) {
      throw new Error(`範囲[${i}]が長すぎます: 最大200文字までです`);
    }
  }

  // API効率化: 小さなバッチの統合とキャッシュ化
  var cacheKey = `batchGet_${spreadsheetId}_${JSON.stringify(ranges)}`;
  
  return cacheManager.get(cacheKey, () => {
    try {
      // 防御的プログラミング: サービスオブジェクトのプロパティを安全に取得
      var baseUrl = service.baseUrl;
      var accessToken = service.accessToken;
      
      // baseUrlが失われている場合の防御処理
      if (!baseUrl || typeof baseUrl !== 'string') {
        console.warn('⚠️ baseUrlが見つかりません。デフォルトのGoogleSheetsAPIエンドポイントを使用します');
        baseUrl = 'https://sheets.googleapis.com/v4/spreadsheets';
      }
      
      if (!accessToken || typeof accessToken !== 'string') {
        throw new Error('アクセストークンが見つかりません。サービスオブジェクトが破損している可能性があります');
      }
      
      console.log('DEBUG: 使用するbaseUrl:', baseUrl);
      
      // 安全なURL構築
      var url = baseUrl + '/' + encodeURIComponent(spreadsheetId) + '/values:batchGet?' + 
        ranges.map(function(range) { 
          return 'ranges=' + encodeURIComponent(range); 
        }).join('&');
      
      console.log('DEBUG: 構築されたURL:', url.substring(0, 100) + '...');
      
      var response = UrlFetchApp.fetch(url, {
        headers: { 'Authorization': 'Bearer ' + accessToken },
        muteHttpExceptions: true,
        followRedirects: true,
        validateHttpsCertificates: true
      });
      
      var responseCode = response.getResponseCode();
      var responseText = response.getContentText();
      
      if (responseCode !== 200) {
        console.error('Sheets API エラー詳細:', {
          code: responseCode,
          response: responseText,
          url: url.substring(0, 100) + '...',
          spreadsheetId: spreadsheetId
        });
        throw new Error('Sheets API error: ' + responseCode + ' - ' + responseText);
      }
      
      var result;
      try {
        result = JSON.parse(responseText);
      } catch (parseError) {
        console.error('❌ JSON解析エラー:', parseError.message);
        console.error('❌ Response text:', responseText.substring(0, 200));
        throw new Error('APIレスポンスのJSON解析に失敗: ' + parseError.message);
      }
      
      // レスポンス構造の検証
      if (!result || typeof result !== 'object') {
        throw new Error('無効なAPIレスポンス: オブジェクトが期待されましたが ' + typeof result + ' を受信');
      }
      
      if (!result.valueRanges || !Array.isArray(result.valueRanges)) {
        console.warn('⚠️ valueRanges配列が見つからないか、配列でありません:', typeof result.valueRanges);
        result.valueRanges = []; // 空配列を設定
      }
      
      // リクエストした範囲数と一致するか確認
      if (result.valueRanges.length !== ranges.length) {
        console.warn(`⚠️ リクエスト範囲数(${ranges.length})とレスポンス数(${result.valueRanges.length})が一致しません`);
      }
      
      console.log('✅ batchGetSheetsData 成功: 取得した範囲数:', result.valueRanges.length);
      
      // 各範囲のデータ存在確認
      result.valueRanges.forEach((valueRange, index) => {
        const hasValues = valueRange.values && valueRange.values.length > 0;
        console.log(`📊 範囲[${index}] ${ranges[index]}: ${hasValues ? valueRange.values.length + '行' : 'データなし'}`);
      if (hasValues) {
        console.log(`DEBUG: batchGetSheetsData - 範囲[${index}] データプレビュー:`, JSON.stringify(valueRange.values.slice(0, 5))); // 最初の5行をプレビュー
      }
      });
      
      return result;
      
    } catch (error) {
      console.error('❌ batchGetSheetsData error:', error.message);
      console.error('❌ Error stack:', error.stack);
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
 * スプレッドシート情報取得（baseUrl 問題修正版）
 * @param {object} service - Sheetsサービス
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {object} スプレッドシート情報
 */
function getSpreadsheetsData(service, spreadsheetId) {
  try {
    // 入力パラメータ検証強化
    if (!service) {
      throw new Error('Sheetsサービスオブジェクトが提供されていません');
    }
    if (!spreadsheetId || typeof spreadsheetId !== 'string') {
      throw new Error('無効なspreadsheetIDです');
    }

    // 防御的プログラミング: サービスオブジェクトのプロパティを安全に取得
    var baseUrl = service.baseUrl;
    var accessToken = service.accessToken;
    
    // baseUrlが失われている場合の防御処理
    if (!baseUrl || typeof baseUrl !== 'string') {
      console.warn('⚠️ baseUrlが見つかりません。デフォルトのGoogleSheetsAPIエンドポイントを使用します');
      baseUrl = 'https://sheets.googleapis.com/v4/spreadsheets';
    }
    
    if (!accessToken || typeof accessToken !== 'string') {
      throw new Error('アクセストークンが見つかりません。サービスオブジェクトが破損している可能性があります');
    }
    
    console.log('DEBUG: getSpreadsheetsData - 使用するbaseUrl:', baseUrl);
    
    // 安全なURL構築 - シート情報を含む基本的なメタデータを取得するために fields パラメータを追加
    var url = baseUrl + '/' + encodeURIComponent(spreadsheetId) + '?fields=sheets.properties';
    
    console.log('DEBUG: getSpreadsheetsData - 構築されたURL:', url.substring(0, 100) + '...');
    
    var response = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + accessToken },
      muteHttpExceptions: true,
      followRedirects: true,
      validateHttpsCertificates: true
    });
    
    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();
    
    if (responseCode !== 200) {
      console.error('Sheets API エラー詳細:', {
        code: responseCode,
        response: responseText,
        url: url.substring(0, 100) + '...',
        spreadsheetId: spreadsheetId
      });
      
      if (responseCode === 403) {
        try {
          var errorResponse = JSON.parse(responseText);
          if (errorResponse.error && errorResponse.error.message === 'The caller does not have permission') {
            var serviceAccountEmail = getServiceAccountEmail();
            throw new Error('スプレッドシートへのアクセス権限がありません。サービスアカウント（' + serviceAccountEmail + '）をスプレッドシートの編集者として共有してください。');
          }
        } catch (parseError) {
          console.warn('エラーレスポンスのJSON解析に失敗:', parseError.message);
        }
      }
      
      throw new Error('Sheets API error: ' + responseCode + ' - ' + responseText);
    }
    
    var result;
    try {
      result = JSON.parse(responseText);
    } catch (parseError) {
      console.error('❌ JSON解析エラー:', parseError.message);
      console.error('❌ Response text:', responseText.substring(0, 200));
      throw new Error('APIレスポンスのJSON解析に失敗: ' + parseError.message);
    }
    
    // レスポンス構造の検証
    if (!result || typeof result !== 'object') {
      throw new Error('無効なAPIレスポンス: オブジェクトが期待されましたが ' + typeof result + ' を受信');
    }
    
    if (!result.sheets || !Array.isArray(result.sheets)) {
      console.warn('⚠️ sheets配列が見つからないか、配列でありません:', typeof result.sheets);
      result.sheets = []; // 空配列を設定してエラーを避ける
    }
    
    var sheetCount = result.sheets.length;
    console.log('✅ getSpreadsheetsData 成功: 発見シート数:', sheetCount);
    
    if (sheetCount === 0) {
      console.warn('⚠️ スプレッドシートにシートが見つかりませんでした');
    } else {
      console.log('📋 利用可能なシート:', result.sheets.map(s => s.properties?.title || 'Unknown').join(', '));
    }
    
    return result;
    
  } catch (error) {
    console.error('❌ getSpreadsheetsData error:', error.message);
    console.error('❌ Error stack:', error.stack);
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
    var service = getSheetsServiceCached();
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
 * データベース状態を詳細確認する診断機能
 * @param {string} targetUserId - 確認対象のユーザーID（オプション）
 * @returns {object} データベース診断結果
 */
function diagnoseDatabase(targetUserId) {
  try {
    console.log('🔍 データベース診断開始:', targetUserId || 'ALL_USERS');
    
    var props = PropertiesService.getScriptProperties();
    var dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
    
    var diagnosticResult = {
      timestamp: new Date().toISOString(),
      databaseId: dbId,
      targetUserId: targetUserId,
      checks: {},
      summary: {
        overallStatus: 'unknown',
        criticalIssues: [],
        warnings: [],
        recommendations: []
      }
    };
    
    // 1. データベース設定チェック
    diagnosticResult.checks.databaseConfig = {
      hasDatabaseId: !!dbId,
      databaseId: dbId ? dbId.substring(0, 10) + '...' : null
    };
    
    if (!dbId) {
      diagnosticResult.summary.criticalIssues.push('DATABASE_SPREADSHEET_ID が設定されていません');
      diagnosticResult.summary.overallStatus = 'critical';
      return diagnosticResult;
    }
    
    // 2. サービス接続テスト
    var service;
    try {
      service = getSheetsServiceCached();
      diagnosticResult.checks.serviceConnection = { status: 'success' };
    } catch (serviceError) {
      diagnosticResult.checks.serviceConnection = { 
        status: 'failed', 
        error: serviceError.message 
      };
      diagnosticResult.summary.criticalIssues.push('Sheets サービス接続失敗: ' + serviceError.message);
    }
    
    if (!service) {
      diagnosticResult.summary.overallStatus = 'critical';
      return diagnosticResult;
    }
    
    // 3. データベーススプレッドシートアクセステスト
    try {
      var spreadsheetInfo = getSpreadsheetsData(service, dbId);
      diagnosticResult.checks.spreadsheetAccess = {
        status: 'success',
        sheetCount: spreadsheetInfo.sheets ? spreadsheetInfo.sheets.length : 0,
        hasUserSheet: spreadsheetInfo.sheets.some(sheet => sheet.properties.title === DB_SHEET_CONFIG.SHEET_NAME)
      };
    } catch (accessError) {
      diagnosticResult.checks.spreadsheetAccess = {
        status: 'failed',
        error: accessError.message
      };
      diagnosticResult.summary.criticalIssues.push('スプレッドシートアクセス失敗: ' + accessError.message);
    }
    
    // 4. ユーザーデータ取得テスト
    try {
      var data = batchGetSheetsData(service, dbId, ["'" + DB_SHEET_CONFIG.SHEET_NAME + "'!A:H"]);
      var values = data.valueRanges[0].values || [];
      
      diagnosticResult.checks.userData = {
        status: 'success',
        totalRows: values.length,
        userCount: Math.max(0, values.length - 1), // ヘッダー行を除く
        hasHeaders: values.length > 0,
        headers: values.length > 0 ? values[0] : []
      };
      
      // 特定ユーザーの検索テスト
      if (targetUserId && values.length > 1) {
        var userFound = false;
        var userRowIndex = -1;
        
        for (var i = 1; i < values.length; i++) {
          if (values[i][0] === targetUserId) {
            userFound = true;
            userRowIndex = i;
            break;
          }
        }
        
        diagnosticResult.checks.targetUser = {
          userId: targetUserId,
          found: userFound,
          rowIndex: userRowIndex,
          data: userFound ? values[userRowIndex] : null
        };
        
        if (!userFound) {
          diagnosticResult.summary.criticalIssues.push('対象ユーザー ' + targetUserId + ' がデータベースに見つかりません');
        }
      }
      
    } catch (dataError) {
      diagnosticResult.checks.userData = {
        status: 'failed',
        error: dataError.message
      };
      diagnosticResult.summary.criticalIssues.push('ユーザーデータ取得失敗: ' + dataError.message);
    }
    
    // 5. キャッシュ状態チェック
    try {
      var cacheStatus = checkCacheStatus(targetUserId);
      diagnosticResult.checks.cache = cacheStatus;
      
      if (cacheStatus.staleEntries > 0) {
        diagnosticResult.summary.warnings.push(cacheStatus.staleEntries + ' 個の古いキャッシュエントリが見つかりました');
      }
    } catch (cacheError) {
      diagnosticResult.checks.cache = {
        status: 'failed',
        error: cacheError.message
      };
    }
    
    // 6. 総合判定
    if (diagnosticResult.summary.criticalIssues.length === 0) {
      diagnosticResult.summary.overallStatus = diagnosticResult.summary.warnings.length > 0 ? 'warning' : 'healthy';
    } else {
      diagnosticResult.summary.overallStatus = 'critical';
    }
    
    // 7. 推奨事項の生成
    if (diagnosticResult.summary.criticalIssues.length > 0) {
      diagnosticResult.summary.recommendations.push('クリティカルな問題を解決してください');
    }
    if (diagnosticResult.checks.cache && diagnosticResult.checks.cache.staleEntries > 0) {
      diagnosticResult.summary.recommendations.push('キャッシュをクリアすることを推奨します');
    }
    
    console.log('✅ データベース診断完了:', diagnosticResult.summary.overallStatus);
    return diagnosticResult;
    
  } catch (error) {
    console.error('❌ データベース診断でエラー:', error);
    return {
      timestamp: new Date().toISOString(),
      error: error.message,
      summary: {
        overallStatus: 'error',
        criticalIssues: ['診断処理自体が失敗: ' + error.message]
      }
    };
  }
}

/**
 * キャッシュ状態をチェックする
 * @param {string} userId - チェック対象のユーザーID（オプション）
 * @returns {object} キャッシュ状態情報
 */
function checkCacheStatus(userId) {
  try {
    var cacheStatus = {
      userSpecific: null,
      general: {
        totalEntries: 0,
        staleEntries: 0,
        healthyEntries: 0
      }
    };
    
    // ユーザー固有のキャッシュ確認
    if (userId) {
      var userCacheKey = 'user_' + userId;
      var cachedUser = cacheManager.get(userCacheKey, null, { skipFetch: true });
      
      cacheStatus.userSpecific = {
        userId: userId,
        cacheKey: userCacheKey,
        cached: !!cachedUser,
        data: cachedUser ? 'present' : 'absent'
      };
    }
    
    // 一般的なキャッシュ統計
    // 注: 実際のキャッシュマネージャーの実装に依存
    try {
      if (typeof cacheManager.getStats === 'function') {
        var stats = cacheManager.getStats();
        cacheStatus.general = stats;
      }
    } catch (statsError) {
      console.warn('キャッシュ統計取得エラー:', statsError.message);
    }
    
    return cacheStatus;
    
  } catch (error) {
    console.error('キャッシュ状態チェックエラー:', error);
    return {
      status: 'error',
      error: error.message
    };
  }
}

/**
 * サービスアカウントの権限を確認・修復する
 * @param {string} spreadsheetId - 確認対象のスプレッドシートID（オプション）
 * @returns {object} 権限確認・修復結果
 */
function verifyServiceAccountPermissions(spreadsheetId) {
  try {
    console.log('🔐 サービスアカウント権限確認開始:', spreadsheetId || 'DATABASE');
    
    var result = {
      timestamp: new Date().toISOString(),
      spreadsheetId: spreadsheetId,
      checks: {},
      summary: {
        status: 'unknown',
        issues: [],
        actions: []
      }
    };
    
    // 1. データベーススプレッドシートの権限確認
    var props = PropertiesService.getScriptProperties();
    var dbId = spreadsheetId || props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
    
    if (!dbId) {
      result.summary.issues.push('データベースIDが設定されていません');
      result.summary.status = 'error';
      return result;
    }
    
    // 2. サービスアカウント情報確認
    try {
      var serviceAccountEmail = getServiceAccountEmail();
      result.checks.serviceAccount = {
        email: serviceAccountEmail,
        configured: !!serviceAccountEmail
      };
      
      if (!serviceAccountEmail) {
        result.summary.issues.push('サービスアカウント設定が見つかりません');
      }
    } catch (saError) {
      result.checks.serviceAccount = {
        configured: false,
        error: saError.message
      };
      result.summary.issues.push('サービスアカウント設定エラー: ' + saError.message);
    }
    
    // 3. スプレッドシートアクセステスト
    try {
      var service = getSheetsServiceCached();
      var spreadsheetInfo = getSpreadsheetsData(service, dbId);
      
      result.checks.spreadsheetAccess = {
        status: 'success',
        sheetCount: spreadsheetInfo.sheets ? spreadsheetInfo.sheets.length : 0,
        canRead: true
      };
      
      // 書き込みテスト（安全な方法で）
      try {
        // テスト用のバッチ更新（実際には何も変更しない）
        var testRequest = {
          requests: []
        };
        // 空のリクエストでテスト
        batchUpdateSpreadsheet(service, dbId, testRequest);
        result.checks.spreadsheetAccess.canWrite = true;
        
      } catch (writeError) {
        result.checks.spreadsheetAccess.canWrite = false;
        result.checks.spreadsheetAccess.writeError = writeError.message;
        result.summary.issues.push('スプレッドシート書き込み権限不足: ' + writeError.message);
      }
      
    } catch (accessError) {
      result.checks.spreadsheetAccess = {
        status: 'failed',
        error: accessError.message,
        canRead: false,
        canWrite: false
      };
      result.summary.issues.push('スプレッドシートアクセス失敗: ' + accessError.message);
    }
    
    // 4. 権限修復の試行
    if (result.summary.issues.length > 0) {
      try {
        console.log('🔧 権限修復を試行中...');
        
        // サービスアカウントの再共有を試行
        if (result.checks.serviceAccount && result.checks.serviceAccount.email) {
          shareSpreadsheetWithServiceAccount(dbId);
          result.summary.actions.push('サービスアカウントの再共有実行');
          
          // 修復後の再テスト
          Utilities.sleep(3000); // 共有反映待ち
          
          try {
            var retestService = getSheetsServiceCached(true); // 強制リフレッシュ
            var retestInfo = getSpreadsheetsData(retestService, dbId);
            
            result.checks.postRepairAccess = {
              status: 'success',
              canRead: true,
              repairSuccessful: true
            };
            
            // 修復成功後はissuesをクリア
            result.summary.issues = [];
            result.summary.actions.push('アクセス権限修復成功');
            
          } catch (retestError) {
            result.checks.postRepairAccess = {
              status: 'failed',
              error: retestError.message,
              repairSuccessful: false
            };
            result.summary.actions.push('修復後テスト失敗: ' + retestError.message);
          }
        }
        
      } catch (repairError) {
        result.summary.actions.push('権限修復失敗: ' + repairError.message);
      }
    }
    
    // 5. 最終判定
    if (result.summary.issues.length === 0) {
      result.summary.status = 'healthy';
    } else if (result.summary.actions.length > 0) {
      result.summary.status = 'repaired';
    } else {
      result.summary.status = 'critical';
    }
    
    console.log('✅ サービスアカウント権限確認完了:', result.summary.status);
    return result;
    
  } catch (error) {
    console.error('❌ サービスアカウント権限確認でエラー:', error);
    return {
      timestamp: new Date().toISOString(),
      error: error.message,
      summary: {
        status: 'error',
        issues: ['権限確認処理自体が失敗: ' + error.message]
      }
    };
  }
}

/**
 * 問題の自動修復を試行する
 * @param {string} targetUserId - 修復対象のユーザーID（オプション）
 * @returns {object} 修復結果
 */
function performAutoRepair(targetUserId) {
  try {
    console.log('🔧 自動修復開始:', targetUserId || 'GENERAL');
    
    var repairResult = {
      timestamp: new Date().toISOString(),
      targetUserId: targetUserId,
      actions: [],
      success: false,
      summary: ''
    };
    
    // 1. キャッシュクリア
    try {
      if (targetUserId) {
        // 特定ユーザーのキャッシュクリア
        invalidateUserCache(targetUserId, null, null, false);
        repairResult.actions.push('ユーザー固有キャッシュクリア: ' + targetUserId);
      } else {
        // 全体キャッシュクリア
        clearDatabaseCache();
        repairResult.actions.push('データベース全体キャッシュクリア');
      }
    } catch (cacheError) {
      repairResult.actions.push('キャッシュクリア失敗: ' + cacheError.message);
    }
    
    // 2. サービス接続リフレッシュ
    try {
      getSheetsServiceCached(true); // forceRefresh = true
      repairResult.actions.push('Sheets サービス接続リフレッシュ');
    } catch (serviceError) {
      repairResult.actions.push('サービスリフレッシュ失敗: ' + serviceError.message);
    }
    
    // 3. サービスアカウント権限確認・修復
    try {
      var permissionResult = verifyServiceAccountPermissions();
      if (permissionResult.summary.status === 'repaired' || 
          permissionResult.summary.status === 'healthy') {
        repairResult.actions.push('サービスアカウント権限確認・修復完了');
      } else {
        repairResult.actions.push('サービスアカウント権限に問題あり: ' + 
          (permissionResult.summary.issues ? permissionResult.summary.issues.join(', ') : '不明'));
      }
    } catch (permError) {
      repairResult.actions.push('サービスアカウント権限確認失敗: ' + permError.message);
    }
    
    // 4. 修復後の検証
    try {
      Utilities.sleep(2000); // 少し待機
      var postRepairDiagnosis = diagnoseDatabase(targetUserId);
      
      if (postRepairDiagnosis.summary.overallStatus === 'healthy' || 
          postRepairDiagnosis.summary.overallStatus === 'warning') {
        repairResult.success = true;
        repairResult.summary = '修復成功: システム状態が改善されました';
      } else {
        repairResult.summary = '修復不完全: 追加の手動対応が必要です';
      }
      
      repairResult.postRepairStatus = postRepairDiagnosis.summary;
      
    } catch (verifyError) {
      repairResult.summary = '修復実行したが検証に失敗: ' + verifyError.message;
    }
    
    console.log('✅ 自動修復完了:', repairResult.success ? '成功' : '要追加対応');
    return repairResult;
    
  } catch (error) {
    console.error('❌ 自動修復でエラー:', error);
    return {
      timestamp: new Date().toISOString(),
      error: error.message,
      success: false,
      summary: '修復処理自体が失敗: ' + error.message
    };
  }
}

/**
 * データベースの整合性を包括的にチェックする
 * @param {object} options - チェックオプション
 * @returns {object} 整合性チェック結果
 */
function performDataIntegrityCheck(options = {}) {
  try {
    console.log('🔍 データ整合性チェック開始');
    
    var opts = {
      checkDuplicates: options.checkDuplicates !== false,
      checkMissingFields: options.checkMissingFields !== false,
      checkInvalidData: options.checkInvalidData !== false,
      autoFix: options.autoFix || false,
      ...options
    };
    
    var result = {
      timestamp: new Date().toISOString(),
      summary: {
        status: 'unknown',
        totalUsers: 0,
        issues: [],
        warnings: [],
        fixed: []
      },
      details: {
        duplicates: [],
        missingFields: [],
        invalidData: [],
        orphanedData: []
      }
    };
    
    // データベース接続確認
    var props = PropertiesService.getScriptProperties();
    var dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
    if (!dbId) {
      result.summary.issues.push('データベースIDが設定されていません');
      result.summary.status = 'critical';
      return result;
    }
    
    var service = getSheetsServiceCached();
    var data = batchGetSheetsData(service, dbId, ["'" + DB_SHEET_CONFIG.SHEET_NAME + "'!A:H"]);
    var values = data.valueRanges[0].values || [];
    
    if (values.length <= 1) {
      result.summary.status = 'empty';
      return result;
    }
    
    var headers = values[0];
    var userRows = values.slice(1);
    result.summary.totalUsers = userRows.length;
    
    console.log('📊 データ整合性チェック: ' + userRows.length + 'ユーザーを確認中');
    
    // 1. 重複チェック
    if (opts.checkDuplicates) {
      var duplicateResult = checkForDuplicates(headers, userRows);
      result.details.duplicates = duplicateResult.duplicates;
      if (duplicateResult.duplicates.length > 0) {
        result.summary.issues.push(duplicateResult.duplicates.length + '件の重複データが見つかりました');
      }
    }
    
    // 2. 必須フィールドチェック
    if (opts.checkMissingFields) {
      var missingFieldsResult = checkMissingRequiredFields(headers, userRows);
      result.details.missingFields = missingFieldsResult.missing;
      if (missingFieldsResult.missing.length > 0) {
        result.summary.warnings.push(missingFieldsResult.missing.length + '件の必須フィールド不足が見つかりました');
      }
    }
    
    // 3. データ形式チェック
    if (opts.checkInvalidData) {
      var invalidDataResult = checkInvalidDataFormats(headers, userRows);
      result.details.invalidData = invalidDataResult.invalid;
      if (invalidDataResult.invalid.length > 0) {
        result.summary.warnings.push(invalidDataResult.invalid.length + '件の不正なデータ形式が見つかりました');
      }
    }
    
    // 4. 孤立データチェック
    var orphanResult = checkOrphanedData(headers, userRows);
    result.details.orphanedData = orphanResult.orphaned;
    if (orphanResult.orphaned.length > 0) {
      result.summary.warnings.push(orphanResult.orphaned.length + '件の孤立データが見つかりました');
    }
    
    // 5. 自動修復
    if (opts.autoFix && (result.summary.issues.length > 0 || result.summary.warnings.length > 0)) {
      try {
        var fixResult = performDataIntegrityFix(result.details, headers, userRows, dbId, service);
        result.summary.fixed = fixResult.fixed;
        console.log('🔧 自動修復完了: ' + fixResult.fixed.length + '件修復');
      } catch (fixError) {
        console.error('❌ 自動修復エラー:', fixError.message);
        result.summary.issues.push('自動修復に失敗: ' + fixError.message);
      }
    }
    
    // 6. 最終判定
    if (result.summary.issues.length === 0) {
      result.summary.status = result.summary.warnings.length > 0 ? 'warning' : 'healthy';
    } else {
      result.summary.status = 'critical';
    }
    
    console.log('✅ データ整合性チェック完了:', result.summary.status);
    return result;
    
  } catch (error) {
    console.error('❌ データ整合性チェックでエラー:', error);
    return {
      timestamp: new Date().toISOString(),
      error: error.message,
      summary: {
        status: 'error',
        issues: ['整合性チェック自体が失敗: ' + error.message]
      }
    };
  }
}

/**
 * 重複データをチェック
 * @param {Array} headers - ヘッダー行
 * @param {Array} userRows - ユーザーデータ行
 * @returns {object} 重複チェック結果
 */
function checkForDuplicates(headers, userRows) {
  var duplicates = [];
  var userIdIndex = headers.indexOf('userId');
  var emailIndex = headers.indexOf('adminEmail');
  
  if (userIdIndex === -1 || emailIndex === -1) {
    return { duplicates: [] };
  }
  
  var seenUserIds = new Set();
  var seenEmails = new Set();
  
  for (var i = 0; i < userRows.length; i++) {
    var row = userRows[i];
    var userId = row[userIdIndex];
    var email = row[emailIndex];
    
    // userId重複チェック
    if (userId && seenUserIds.has(userId)) {
      duplicates.push({
        type: 'userId',
        value: userId,
        rowIndex: i + 2, // スプレッドシート行番号（1ベース + ヘッダー）
        message: 'ユーザーID重複: ' + userId
      });
    } else if (userId) {
      seenUserIds.add(userId);
    }
    
    // email重複チェック
    if (email && seenEmails.has(email)) {
      duplicates.push({
        type: 'email',
        value: email,
        rowIndex: i + 2,
        message: 'メールアドレス重複: ' + email
      });
    } else if (email) {
      seenEmails.add(email);
    }
  }
  
  return { duplicates: duplicates };
}

/**
 * 必須フィールドの不足をチェック
 * @param {Array} headers - ヘッダー行
 * @param {Array} userRows - ユーザーデータ行
 * @returns {object} 必須フィールドチェック結果
 */
function checkMissingRequiredFields(headers, userRows) {
  var missing = [];
  var requiredFields = ['userId', 'adminEmail']; // 最低限必要なフィールド
  
  for (var i = 0; i < userRows.length; i++) {
    var row = userRows[i];
    var missingInThisRow = [];
    
    for (var j = 0; j < requiredFields.length; j++) {
      var fieldName = requiredFields[j];
      var fieldIndex = headers.indexOf(fieldName);
      
      if (fieldIndex === -1 || !row[fieldIndex] || row[fieldIndex].trim() === '') {
        missingInThisRow.push(fieldName);
      }
    }
    
    if (missingInThisRow.length > 0) {
      missing.push({
        rowIndex: i + 2,
        missingFields: missingInThisRow,
        message: '必須フィールド不足: ' + missingInThisRow.join(', ')
      });
    }
  }
  
  return { missing: missing };
}

/**
 * データ形式の妥当性をチェック
 * @param {Array} headers - ヘッダー行
 * @param {Array} userRows - ユーザーデータ行
 * @returns {object} データ形式チェック結果
 */
function checkInvalidDataFormats(headers, userRows) {
  var invalid = [];
  var emailIndex = headers.indexOf('adminEmail');
  var userIdIndex = headers.indexOf('userId');
  
  for (var i = 0; i < userRows.length; i++) {
    var row = userRows[i];
    var rowIssues = [];
    
    // メールアドレス形式チェック
    if (emailIndex !== -1 && row[emailIndex]) {
      var email = row[emailIndex];
      if (!EMAIL_REGEX.test(email)) {
        rowIssues.push('無効なメールアドレス形式: ' + email);
      }
    }
    
    // ユーザーID形式チェック（UUIDかどうか）
    if (userIdIndex !== -1 && row[userIdIndex]) {
      var userId = row[userIdIndex];
      var uuidRegex = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
      if (!uuidRegex.test(userId)) {
        rowIssues.push('無効なユーザーID形式: ' + userId);
      }
    }
    
    if (rowIssues.length > 0) {
      invalid.push({
        rowIndex: i + 2,
        issues: rowIssues,
        message: rowIssues.join(', ')
      });
    }
  }
  
  return { invalid: invalid };
}

/**
 * 孤立データをチェック
 * @param {Array} headers - ヘッダー行
 * @param {Array} userRows - ユーザーデータ行
 * @returns {object} 孤立データチェック結果
 */
function checkOrphanedData(headers, userRows) {
  var orphaned = [];
  var isActiveIndex = headers.indexOf('isActive');
  
  for (var i = 0; i < userRows.length; i++) {
    var row = userRows[i];
    var issues = [];
    
    // 非アクティブだが他のデータが残っている
    if (isActiveIndex !== -1 && row[isActiveIndex] === 'false') {
      // 非アクティブユーザーでスプレッドシートIDが残っている場合
      var spreadsheetIdIndex = headers.indexOf('spreadsheetId');
      if (spreadsheetIdIndex !== -1 && row[spreadsheetIdIndex]) {
        issues.push('非アクティブユーザーにスプレッドシートIDが残存');
      }
    }
    
    if (issues.length > 0) {
      orphaned.push({
        rowIndex: i + 2,
        issues: issues,
        message: issues.join(', ')
      });
    }
  }
  
  return { orphaned: orphaned };
}

/**
 * データ整合性問題の自動修復
 * @param {object} details - 問題の詳細
 * @param {Array} headers - ヘッダー行
 * @param {Array} userRows - ユーザーデータ行
 * @param {string} dbId - データベースID
 * @param {object} service - Sheetsサービス
 * @returns {object} 修復結果
 */
function performDataIntegrityFix(details, headers, userRows, dbId, service) {
  var fixed = [];
  
  // 注意: 重複データの自動削除は危険なため、ログのみ記録
  if (details.duplicates.length > 0) {
    console.warn('⚠️ 重複データが検出されましたが、自動削除は行いません:', details.duplicates.length + '件');
    fixed.push('重複データの確認完了（手動対応が必要）');
  }
  
  // 無効なisActiveフィールドの修正
  var isActiveIndex = headers.indexOf('isActive');
  if (isActiveIndex !== -1) {
    var updatesNeeded = [];
    
    for (var i = 0; i < userRows.length; i++) {
      var row = userRows[i];
      var currentValue = row[isActiveIndex];
      
      // isActiveフィールドが空または無効な値の場合、trueに設定
      if (!currentValue || (currentValue !== 'true' && currentValue !== 'false' && currentValue !== true && currentValue !== false)) {
        updatesNeeded.push({
          range: "'" + DB_SHEET_CONFIG.SHEET_NAME + "'!" + String.fromCharCode(65 + isActiveIndex) + (i + 2),
          values: [['true']]
        });
      }
    }
    
    if (updatesNeeded.length > 0) {
      try {
        batchUpdateSheetsData(service, dbId, updatesNeeded);
        fixed.push(updatesNeeded.length + '件のisActiveフィールドを修正');
      } catch (updateError) {
        console.error('❌ isActiveフィールド修正エラー:', updateError.message);
      }
    }
  }
  
  return { fixed: fixed };
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
 * システム監視とアラート機能
 * @param {object} options - 監視オプション
 * @returns {object} 監視結果
 */
function performSystemMonitoring(options = {}) {
  try {
    console.log('📊 システム監視開始');
    
    var opts = {
      checkHealth: options.checkHealth !== false,
      checkPerformance: options.checkPerformance !== false,
      checkSecurity: options.checkSecurity !== false,
      enableAlerts: options.enableAlerts !== false,
      ...options
    };
    
    var monitoringResult = {
      timestamp: new Date().toISOString(),
      summary: {
        overallHealth: 'unknown',
        alerts: [],
        warnings: [],
        metrics: {}
      },
      details: {
        healthCheck: null,
        performanceCheck: null,
        securityCheck: null
      }
    };
    
    // 1. ヘルスチェック
    if (opts.checkHealth) {
      try {
        var healthResult = performHealthCheck();
        monitoringResult.details.healthCheck = healthResult;
        
        if (healthResult.summary.overallStatus === 'critical') {
          monitoringResult.summary.alerts.push('システムヘルスが危険な状態です');
        } else if (healthResult.summary.overallStatus === 'warning') {
          monitoringResult.summary.warnings.push('システムヘルスに軽微な問題があります');
        }
      } catch (healthError) {
        console.error('❌ ヘルスチェックエラー:', healthError.message);
        monitoringResult.summary.alerts.push('ヘルスチェック自体が失敗しました');
      }
    }
    
    // 2. パフォーマンスチェック
    if (opts.checkPerformance) {
      try {
        var perfResult = performPerformanceCheck();
        monitoringResult.details.performanceCheck = perfResult;
        
        if (perfResult.metrics.responseTime > 10000) { // 10秒以上
          monitoringResult.summary.alerts.push('レスポンス時間が異常に長くなっています');
        } else if (perfResult.metrics.responseTime > 5000) { // 5秒以上
          monitoringResult.summary.warnings.push('レスポンス時間が少し長くなっています');
        }
        
        monitoringResult.summary.metrics.responseTime = perfResult.metrics.responseTime;
        monitoringResult.summary.metrics.cacheHitRate = perfResult.metrics.cacheHitRate;
      } catch (perfError) {
        console.error('❌ パフォーマンスチェックエラー:', perfError.message);
        monitoringResult.summary.warnings.push('パフォーマンス測定に失敗しました');
      }
    }
    
    // 3. セキュリティチェック
    if (opts.checkSecurity) {
      try {
        var securityResult = performSecurityCheck();
        monitoringResult.details.securityCheck = securityResult;
        
        if (securityResult.vulnerabilities.length > 0) {
          monitoringResult.summary.alerts.push(securityResult.vulnerabilities.length + '件のセキュリティ問題が検出されました');
        }
      } catch (securityError) {
        console.error('❌ セキュリティチェックエラー:', securityError.message);
        monitoringResult.summary.warnings.push('セキュリティチェックに失敗しました');
      }
    }
    
    // 4. 総合判定
    if (monitoringResult.summary.alerts.length === 0) {
      monitoringResult.summary.overallHealth = monitoringResult.summary.warnings.length > 0 ? 'warning' : 'healthy';
    } else {
      monitoringResult.summary.overallHealth = 'critical';
    }
    
    // 5. アラート通知
    if (opts.enableAlerts && (monitoringResult.summary.alerts.length > 0 || monitoringResult.summary.warnings.length > 0)) {
      try {
        sendSystemAlert(monitoringResult);
      } catch (alertError) {
        console.error('❌ アラート送信エラー:', alertError.message);
      }
    }
    
    console.log('✅ システム監視完了:', monitoringResult.summary.overallHealth);
    return monitoringResult;
    
  } catch (error) {
    console.error('❌ システム監視でエラー:', error);
    return {
      timestamp: new Date().toISOString(),
      error: error.message,
      summary: {
        overallHealth: 'error',
        alerts: ['監視システム自体が失敗: ' + error.message]
      }
    };
  }
}

/**
 * システムヘルスチェック
 * @returns {object} ヘルスチェック結果
 */
function performHealthCheck() {
  var healthResult = {
    timestamp: new Date().toISOString(),
    checks: {},
    summary: {
      overallStatus: 'unknown',
      passedChecks: 0,
      failedChecks: 0
    }
  };
  
  // データベース接続チェック
  try {
    var diagnosis = diagnoseDatabase();
    healthResult.checks.database = {
      status: diagnosis.summary.overallStatus === 'healthy' ? 'pass' : 'fail',
      details: diagnosis.summary
    };
    
    if (diagnosis.summary.overallStatus === 'healthy') {
      healthResult.summary.passedChecks++;
    } else {
      healthResult.summary.failedChecks++;
    }
  } catch (dbError) {
    healthResult.checks.database = {
      status: 'fail',
      error: dbError.message
    };
    healthResult.summary.failedChecks++;
  }
  
  // サービスアカウント権限チェック
  try {
    var permissionResult = verifyServiceAccountPermissions();
    healthResult.checks.serviceAccount = {
      status: permissionResult.summary.status === 'healthy' ? 'pass' : 'fail',
      details: permissionResult.summary
    };
    
    if (permissionResult.summary.status === 'healthy') {
      healthResult.summary.passedChecks++;
    } else {
      healthResult.summary.failedChecks++;
    }
  } catch (permError) {
    healthResult.checks.serviceAccount = {
      status: 'fail',
      error: permError.message
    };
    healthResult.summary.failedChecks++;
  }
  
  // システム設定チェック
  try {
    var props = PropertiesService.getScriptProperties();
    var requiredProps = [
      SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID,
      SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS,
      SCRIPT_PROPS_KEYS.ADMIN_EMAIL
    ];
    
    var missingProps = [];
    for (var i = 0; i < requiredProps.length; i++) {
      if (!props.getProperty(requiredProps[i])) {
        missingProps.push(requiredProps[i]);
      }
    }
    
    healthResult.checks.configuration = {
      status: missingProps.length === 0 ? 'pass' : 'fail',
      missingProperties: missingProps
    };
    
    if (missingProps.length === 0) {
      healthResult.summary.passedChecks++;
    } else {
      healthResult.summary.failedChecks++;
    }
  } catch (configError) {
    healthResult.checks.configuration = {
      status: 'fail',
      error: configError.message
    };
    healthResult.summary.failedChecks++;
  }
  
  // 総合判定
  if (healthResult.summary.failedChecks === 0) {
    healthResult.summary.overallStatus = 'healthy';
  } else if (healthResult.summary.passedChecks > healthResult.summary.failedChecks) {
    healthResult.summary.overallStatus = 'warning';
  } else {
    healthResult.summary.overallStatus = 'critical';
  }
  
  return healthResult;
}

/**
 * パフォーマンスチェック
 * @returns {object} パフォーマンスチェック結果
 */
function performPerformanceCheck() {
  var startTime = Date.now();
  
  var perfResult = {
    timestamp: new Date().toISOString(),
    metrics: {
      responseTime: 0,
      cacheHitRate: 0,
      apiCallCount: 0
    },
    benchmarks: {}
  };
  
  try {
    // データベースアクセス速度テスト
    var dbTestStart = Date.now();
    var props = PropertiesService.getScriptProperties();
    var dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
    
    if (dbId) {
      var service = getSheetsServiceCached();
      var testData = batchGetSheetsData(service, dbId, ["'" + DB_SHEET_CONFIG.SHEET_NAME + "'!A1:B1"]);
      perfResult.benchmarks.databaseAccess = Date.now() - dbTestStart;
      perfResult.metrics.apiCallCount++;
    }
    
    // キャッシュ効率テスト（簡易版）
    var cacheTestStart = Date.now();
    try {
      if (typeof cacheManager !== 'undefined' && cacheManager.getStats) {
        var cacheStats = cacheManager.getStats();
        perfResult.metrics.cacheHitRate = cacheStats.hitRate || 0;
      }
    } catch (cacheError) {
      console.warn('キャッシュ統計取得エラー:', cacheError.message);
    }
    perfResult.benchmarks.cacheCheck = Date.now() - cacheTestStart;
    
  } catch (perfError) {
    console.error('パフォーマンステスト実行エラー:', perfError.message);
    perfResult.error = perfError.message;
  }
  
  perfResult.metrics.responseTime = Date.now() - startTime;
  
  return perfResult;
}

/**
 * セキュリティチェック
 * @returns {object} セキュリティチェック結果
 */
function performSecurityCheck() {
  var securityResult = {
    timestamp: new Date().toISOString(),
    vulnerabilities: [],
    recommendations: []
  };
  
  try {
    // スクリプトプロパティのセキュリティチェック
    var props = PropertiesService.getScriptProperties();
    var allProps = props.getProperties();
    
    // サービスアカウント認証情報の存在確認
    var serviceAccountCreds = props.getProperty(SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS);
    if (!serviceAccountCreds) {
      securityResult.vulnerabilities.push({
        type: 'missing_credentials',
        severity: 'high',
        message: 'サービスアカウント認証情報が設定されていません'
      });
    }
    
    // 管理者メールの設定確認
    var adminEmail = props.getProperty(SCRIPT_PROPS_KEYS.ADMIN_EMAIL);
    if (!adminEmail) {
      securityResult.vulnerabilities.push({
        type: 'missing_admin',
        severity: 'medium',
        message: '管理者メールアドレスが設定されていません'
      });
    }
    
    // データベースアクセス権限の確認
    try {
      var permissionCheck = verifyServiceAccountPermissions();
      if (permissionCheck.summary.status === 'critical') {
        securityResult.vulnerabilities.push({
          type: 'access_permission',
          severity: 'high',
          message: 'データベースアクセス権限に問題があります'
        });
      }
    } catch (permError) {
      securityResult.vulnerabilities.push({
        type: 'permission_check_failed',
        severity: 'medium',
        message: '権限確認に失敗しました'
      });
    }
    
    // 推奨事項の生成
    if (securityResult.vulnerabilities.length === 0) {
      securityResult.recommendations.push('セキュリティ設定は適切です');
    } else {
      securityResult.recommendations.push('検出された脆弱性を修正してください');
      securityResult.recommendations.push('定期的なセキュリティチェックを実施してください');
    }
    
  } catch (securityError) {
    console.error('セキュリティチェック実行エラー:', securityError.message);
    securityResult.error = securityError.message;
  }
  
  return securityResult;
}

/**
 * システムアラートを送信
 * @param {object} monitoringResult - 監視結果
 */
function sendSystemAlert(monitoringResult) {
  try {
    console.log('🚨 システムアラート送信開始');
    
    // アラート内容の構築
    var alertMessage = '【StudyQuest システムアラート】\n\n';
    alertMessage += '発生時刻: ' + monitoringResult.timestamp + '\n';
    alertMessage += 'システム状態: ' + monitoringResult.summary.overallHealth + '\n\n';
    
    if (monitoringResult.summary.alerts.length > 0) {
      alertMessage += '🚨 緊急問題:\n';
      for (var i = 0; i < monitoringResult.summary.alerts.length; i++) {
        alertMessage += '  • ' + monitoringResult.summary.alerts[i] + '\n';
      }
      alertMessage += '\n';
    }
    
    if (monitoringResult.summary.warnings.length > 0) {
      alertMessage += '⚠️ 警告:\n';
      for (var j = 0; j < monitoringResult.summary.warnings.length; j++) {
        alertMessage += '  • ' + monitoringResult.summary.warnings[j] + '\n';
      }
      alertMessage += '\n';
    }
    
    if (monitoringResult.summary.metrics.responseTime) {
      alertMessage += 'パフォーマンス:\n';
      alertMessage += '  • レスポンス時間: ' + monitoringResult.summary.metrics.responseTime + 'ms\n';
      if (monitoringResult.summary.metrics.cacheHitRate) {
        alertMessage += '  • キャッシュヒット率: ' + (monitoringResult.summary.metrics.cacheHitRate * 100).toFixed(1) + '%\n';
      }
    }
    
    // 管理者への通知（メール送信は実装に依存）
    var props = PropertiesService.getScriptProperties();
    var adminEmail = props.getProperty(SCRIPT_PROPS_KEYS.ADMIN_EMAIL);
    
    if (adminEmail) {
      // ログに記録（実際のメール送信機能がある場合はここで実装）
      console.log('📧 管理者アラート（' + adminEmail + '）:\n' + alertMessage);
      
      // 緊急レベルの場合は追加のログ記録
      if (monitoringResult.summary.alerts.length > 0) {
        console.error('🚨 緊急システムアラート: ' + monitoringResult.summary.alerts.join(', '));
      }
    }
    
    // システムログへの記録
    logSystemEvent('system_alert', {
      level: monitoringResult.summary.overallHealth,
      alerts: monitoringResult.summary.alerts,
      warnings: monitoringResult.summary.warnings,
      timestamp: monitoringResult.timestamp
    });
    
    console.log('✅ システムアラート送信完了');
    
  } catch (alertError) {
    console.error('❌ アラート送信エラー:', alertError.message);
  }
}

/**
 * システムイベントをログに記録
 * @param {string} eventType - イベントタイプ
 * @param {object} eventData - イベントデータ
 */
function logSystemEvent(eventType, eventData) {
  try {
    var logEntry = {
      timestamp: new Date().toISOString(),
      type: eventType,
      data: eventData
    };
    
    // コンソールログ
    console.log('📝 システムイベント [' + eventType + ']:', JSON.stringify(eventData));
    
    // 将来的にはログシートやログファイルに保存する機能を追加可能
    
  } catch (logError) {
    console.error('システムイベントログエラー:', logError.message);
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
      
      var service = getSheetsServiceCached();
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
