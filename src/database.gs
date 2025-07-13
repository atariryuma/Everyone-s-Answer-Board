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
    var accessToken = getServiceAccountTokenCached();
    if (!accessToken) {
      console.error('Failed to get service account token');
      return null;
    }
    return createSheetsService(accessToken);
  } catch (error) {
    console.error('getSheetsService error:', error.message);
    return null;
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
    var service = getSheetsService();
    var sheetName = DB_SHEET_CONFIG.SHEET_NAME;
    
    var data = batchGetSheetsData(service, dbId, ["'" + sheetName + "'!A:H"]);
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
        
        // デバッグログでマッピングを確認
        console.log('fetchUserFromDatabase - Found user:', {
          field: field,
          value: value,
          userId: user.userId,
          adminEmail: user.adminEmail,
          createdAt: user.createdAt
        });
        
        return user;
      }
    }
    
    return null;
  } catch (error) {
    console.error('ユーザー検索エラー (' + field + ':' + value + '):', error);
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
  try {
    var props = PropertiesService.getScriptProperties();
    var dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
    var service = getSheetsService();
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
    
    // バックアップ用に現在のデータを保存
    var originalValues = values[rowIndex - 1].slice(); // 1-based to 0-based
    
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
    
    // 最適化: 変更された内容に基づいてキャッシュを選択的に無効化
    var userInfo = findUserById(userId);
    var email = updateData.adminEmail || (userInfo ? userInfo.adminEmail : null);
    var spreadsheetId = updateData.spreadsheetId || (userInfo ? userInfo.spreadsheetId : null);
    
    // キャッシュを無効化（必要最小限）
    invalidateUserCache(userId, email, spreadsheetId, false);
    
    return { success: true };
  } catch (error) {
    console.error('ユーザー更新エラー:', error);
    throw error;
  }
}

/**
 * 🎯 原子的なユーザー検索・作成操作（Upsert）
 * 重複を防ぐためのメイン関数
 * @param {string} adminEmail - ユーザーメールアドレス
 * @param {object} additionalData - 追加データ（オプション）
 * @returns {object} { userId, isNewUser, userInfo }
 */
function findOrCreateUser(adminEmail, additionalData = {}) {
  if (!adminEmail || !adminEmail.includes('@')) {
    throw new Error('有効なメールアドレスが必要です');
  }

  // 🔒 適応的ロックメカニズム
  const result = attemptWithAdaptiveLock(adminEmail, additionalData);
  if (result) {
    return result;
  }
  
  // フォールバック: 軽量ロックで再試行
  return attemptWithLightweightLock(adminEmail, additionalData);
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
 * @param {object} userData - 作成するユーザーデータ
 * @returns {object} 作成されたユーザーデータ
 */
function createUserAtomic(userData) {
  const props = PropertiesService.getScriptProperties();
  const dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
  const service = getSheetsService();
  const sheetName = DB_SHEET_CONFIG.SHEET_NAME;

  const newRow = DB_SHEET_CONFIG.HEADERS.map(function(header) { 
    return userData[header] || ''; 
  });
  
  appendSheetsData(service, dbId, "'" + sheetName + "'!A1", [newRow]);
  
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
 * バッチ取得
 * @param {object} service - Sheetsサービス
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string[]} ranges - 取得範囲の配列
 * @returns {object} レスポンス
 */
function batchGetSheetsData(service, spreadsheetId, ranges) {
  // API呼び出しをキャッシュ化（短期間）
  var cacheKey = `batchGet_${spreadsheetId}_${JSON.stringify(ranges)}`;
  
  return cacheManager.get(cacheKey, () => {
    try {
      var url = service.baseUrl + '/' + spreadsheetId + '/values:batchGet?' + 
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
