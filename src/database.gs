/**
 * @fileoverview データベース管理 - バッチ操作とキャッシュ最適化
 * GAS互換の関数ベースの実装
 * 回復力のある実行機構を統合
 * 統合ユーザー検索システム
 */

// 統合ユーザー検索システムは削除（簡素化のため、既存の3段階検索で十分）

// 回復力のあるProperties/Cache操作
function getResilientScriptProperties() {
  try {
    // 直接PropertiesServiceを呼び出し、エラー時のみリトライ
    return PropertiesService.getScriptProperties();
  } catch (error) {
    // 1回だけリトライ
    try {
      Utilities.sleep(1000); // 1秒待機
      return PropertiesService.getScriptProperties();
    } catch (retryError) {
      warnLog('getResilientScriptProperties: PropertiesService取得に失敗しました', {
        originalError: error.message,
        retryError: retryError.message
      });
      // nullを返してフォールバック処理を可能にする
      throw new Error(`PropertiesService取得に失敗しました: ${retryError.message}`);
    }
  }
}

function getResilientUserProperties() {
  try {
    // 直接PropertiesServiceを呼び出し、エラー時のみリトライ
    return PropertiesService.getUserProperties();
  } catch (error) {
    // 1回だけリトライ
    try {
      Utilities.sleep(1000); // 1秒待機
      return PropertiesService.getUserProperties();
    } catch (retryError) {
      warnLog('getResilientUserProperties: UserProperties取得に失敗しました', {
        originalError: error.message,
        retryError: retryError.message
      });
      // nullを返してフォールバック処理を可能にする
      throw new Error(`UserProperties取得に失敗しました: ${retryError.message}`);
    }
  }
}

// データベース管理のための定数
const USER_CACHE_TTL = 300; // 5分
const DB_BATCH_SIZE = 100;

// 簡易インデックス機能：ユーザー検索の高速化
const userIndexCache = {
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
 * 削除ログを安全にトランザクション的に記録
 * @param {string} executorEmail - 削除実行者のメール
 * @param {string} targetUserId - 削除対象のユーザーID
 * @param {string} targetEmail - 削除対象のメール
 * @param {string} reason - 削除理由
 * @param {string} deleteType - 削除タイプ ("self" | "admin")
 */
function logAccountDeletion(executorEmail, targetUserId, targetEmail, reason, deleteType) {
  const transactionLog = {
    startTime: Date.now(),
    steps: [],
    success: false,
    rollbackActions: []
  };

  try {
    // 入力検証
    if (!executorEmail || !targetUserId || !targetEmail || !deleteType) {
      throw new Error('必須パラメータが不足しています');
    }

    const props =  getResilientScriptProperties();
    const dbId =  getSecureDatabaseId();

    if (!dbId) {
      warnLog('削除ログの記録をスキップします: データベースIDが設定されていません');
      return { success: false, reason: 'no_database_id' };
    }

    transactionLog.steps.push('validation_complete');

    const service = getSheetsServiceCached();
    const logSheetName = DELETE_LOG_SHEET_CONFIG.SHEET_NAME;

    // 統一ロック管理でトランザクション開始
    return executeWithStandardizedLock('WRITE_OPERATION', 'logAccountDeletion', () => {
      transactionLog.steps.push('lock_acquired');

      // ログシートの存在確認・作成（トランザクション内）
      let sheetCreated = false;
      try {
        const spreadsheetInfo = getSpreadsheetsData(service, dbId);
        const logSheetExists = spreadsheetInfo.sheets.some(sheet =>
          sheet.properties.title === logSheetName
        );

        if (!logSheetExists) {
          // バッチ最適化: ログシート作成
          debugLog('📊 トランザクション内ログシート作成開始');

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

          transactionLog.steps.push('sheet_created');
          sheetCreated = true;

          // ヘッダー行を追加（作成直後）
          appendSheetsData(service, dbId, `'${logSheetName}'!A1`, [DELETE_LOG_SHEET_CONFIG.HEADERS]);
          transactionLog.steps.push('headers_added');

          infoLog('✅ ログシートとヘッダーの作成完了');
        }
      } catch (sheetError) {
        // シート作成失敗時のロールバック
        if (sheetCreated) {
          transactionLog.rollbackActions.push('remove_created_sheet');
        }
        throw new Error(`ログシートの準備に失敗: ${sheetError.message}`);
      }

      // ログエントリを安全に追加
      const logEntry = [
        new Date().toISOString(),
        String(executorEmail).substring(0, 255), // 長さ制限
        String(targetUserId).substring(0, 255),
        String(targetEmail).substring(0, 255),
        String(reason || '').substring(0, 500), // 理由は500文字まで
        String(deleteType).substring(0, 50)
      ];

      try {
        appendSheetsData(service, dbId, `'${logSheetName}'!A:F`, [logEntry]);
        transactionLog.steps.push('log_entry_added');

        // 検証: 追加されたログエントリの確認
        Utilities.sleep(100); // 書き込み完了待機
        const verificationData = batchGetSheetsData(service, dbId, [`'${logSheetName}'!A:F`]);
        const lastRow = verificationData.valueRanges[0].values?.slice(-1)[0];

        if (!lastRow || lastRow[1] !== executorEmail || lastRow[2] !== targetUserId) {
          throw new Error('ログエントリの検証に失敗しました');
        }

        transactionLog.steps.push('verification_complete');
        transactionLog.success = true;

        infoLog('✅ 削除ログの安全な記録完了:', {
          executor: executorEmail,
          target: targetUserId,
          type: deleteType,
          steps: transactionLog.steps.length
        });

        return {
          success: true,
          logEntry: logEntry,
          transactionLog: transactionLog
        };

      } catch (appendError) {
        throw new Error(`ログエントリの追加に失敗: ${appendError.message}`);
      }
    });

  } catch (error) {
    transactionLog.duration = Date.now() - transactionLog.startTime;

    // 構造化エラーログ
    const errorInfo = {
      timestamp: new Date().toISOString(),
      function: 'logAccountDeletion',
      severity: 'medium', // ログ記録失敗は削除処理自体を止めない
      parameters: { executorEmail, targetUserId, targetEmail, deleteType },
      error: error.message,
      transactionLog: transactionLog
    };

    errorLog('🚨 削除ログ記録エラー:', JSON.stringify(errorInfo, null, 2));

    // ロールバック処理（必要に応じて）
    // 現在はログ記録のみなので、深刻なロールバックは不要

    return {
      success: false,
      error: error.message,
      transactionLog: transactionLog
    };
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

    const props =  getResilientScriptProperties();
    const dbId =  getSecureDatabaseId();

    if (!dbId) {
      throw new Error('データベースIDが設定されていません');
    }

    const service = getSheetsServiceCached();
    const sheetName = DB_SHEET_CONFIG.SHEET_NAME;

    const data =  batchGetSheetsData(service, dbId, [`'${sheetName}'!A:H`]);
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
      debugLog(`DEBUG: getAllUsersForAdmin - User object: ${JSON.stringify(user)}`);
    }

    infoLog(`✅ 管理者用ユーザー一覧を取得: ${users.length}件`);
    return users;

  } catch (error) {
    errorLog('getAllUsersForAdmin error:', error.message);
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

    // 厳格な削除理由の検証
    if (!reason || typeof reason !== 'string' || reason.trim().length < 20) {
      throw new Error('削除理由は20文字以上で入力してください。');
    }

    // 削除理由の内容検証（不適切な理由を防ぐ）
    const invalidReasonPatterns = [
      /test/i, /テスト/i, /試し/i, /適当/i, /dummy/i, /sample/i
    ];

    if (invalidReasonPatterns.some(pattern => pattern.test(reason))) {
      throw new Error('削除理由に適切な内容を入力してください。テスト用の削除は許可されていません。');
    }

    const executorEmail = getCurrentUserEmail();

    // セキュリティチェック: 管理者メールの検証
    if (!executorEmail || !executorEmail.includes('@')) {
      throw new Error('実行者のメールアドレスを取得できません。');
    }

    // 削除対象ユーザー情報を安全に取得
    const targetUserInfo = findUserById(targetUserId);
    if (!targetUserInfo) {
      throw new Error('削除対象のユーザーが見つかりません。');
    }

    // 追加セキュリティチェック
    if (!targetUserInfo.adminEmail || !targetUserInfo.userId) {
      throw new Error('削除対象ユーザーの情報が不完全です。');
    }

    // 自分自身の削除を防ぐ
    if (targetUserInfo.adminEmail === executorEmail) {
      throw new Error('自分自身のアカウントは管理者削除機能では削除できません。個人用削除機能をご利用ください。');
    }

    // 最後のアクティブ確認（削除予定の危険度チェック）
    if (targetUserInfo.lastAccessedAt) {
      const lastAccess = new Date(targetUserInfo.lastAccessedAt);
      const daysSinceLastAccess = (Date.now() - lastAccess.getTime()) / (1000 * 60 * 60 * 24);

      if (daysSinceLastAccess < 7) {
        warnLog(`警告: 削除対象ユーザーは${Math.floor(daysSinceLastAccess)}日前にアクセスしています`);
      }
    }

    // 統一ロック管理でデータベース削除実行
    return executeWithStandardizedLock('CRITICAL_OPERATION', 'deleteUserAccountByAdmin', () => {
      // データベースからユーザー行を削除
      const props =  getResilientScriptProperties();
      const dbId =  getSecureDatabaseId();

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
      const data =  batchGetSheetsData(service, dbId, [`'${sheetName}'!A:H`]);
      const values = data.valueRanges[0].values || [];

      // userIdフィールドのインデックスを特定
      const headers = values[0];
      let userIdFieldIndex = -1;
      
      for (let j = 0; j < headers.length; j++) {
        if (headers[j] === 'userId') {
          userIdFieldIndex = j;
          break;
        }
      }
      
      if (userIdFieldIndex === -1) {
        throw new Error('データベースに userId フィールドが見つかりません');
      }
      
      debugLog('Found userId field at index:', userIdFieldIndex, 'for admin deletion');

      let rowToDelete = -1;
      for (let i = values.length - 1; i >= 1; i--) {
        if (values[i][userIdFieldIndex] === targetUserId) {
          rowToDelete = i + 1; // スプレッドシートは1ベース
          debugLog('Found target user row to delete at index:', i, 'rowToDelete:', rowToDelete);
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

        infoLog(`✅ 管理者削除完了: row ${rowToDelete}, sheetId ${targetSheetId}`);
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

      const successMessage = `管理者によりアカウント「${targetUserInfo.adminEmail}」が削除されました。\n削除理由: ${reason.trim()}`;
      infoLog(successMessage);
      return {
        success: true,
        message: successMessage,
        deletedUser: {
          userId: targetUserId,
          email: targetUserInfo.adminEmail
        }
      };
    });

  } catch (error) {
    errorLog('deleteUserAccountByAdmin error:', error.message);
    throw new Error('管理者によるアカウント削除に失敗しました: ' + error.message);
  }
}

/**
 * 削除権限チェック
 * @param {string} targetUserId - 対象ユーザーID
 */
function canDeleteUser(targetUserId) {
  try {
    const currentUserEmail = getCurrentUserEmail();
    const targetUser = findUserById(targetUserId);

    if (!targetUser) {
      return false;
    }

    // 本人削除 OR 管理者削除
    return (currentUserEmail === targetUser.adminEmail) || isDeployUser();
  } catch (error) {
    errorLog('canDeleteUser error:', error.message);
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

    const props =  getResilientScriptProperties();
    const dbId =  getSecureDatabaseId();

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
        debugLog('削除ログシートが存在しません');
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

      infoLog(`✅ 削除ログを取得: ${logs.length}件`);
      return logs;

    } catch (sheetError) {
      warnLog('ログシートの読み取りに失敗:', sheetError.message);
      return []; // シートアクセスエラーの場合は空を返す
    }

  } catch (error) {
    errorLog('getDeletionLogs error:', error.message);
    throw new Error('削除ログの取得に失敗しました: ' + error.message);
  }
}

/**
 * 長期キャッシュ対応Sheetsサービスを取得
 * @returns {object} Sheets APIサービス
 */
function getSheetsServiceCached(forceRefresh) {
  try {
    debugLog('🔧 getSheetsServiceCached: 新規サービス作成開始（キャッシュなし版）');

    let accessToken;
    if (forceRefresh) {
      debugLog('🔐 認証トークンも強制リフレッシュ');
      cacheManager.remove('service_account_token');
      accessToken = generateNewServiceAccountToken();
    } else {
      accessToken = getServiceAccountTokenCached();
    }

    if (!accessToken) {
      throw new Error('サービスアカウントトークンが取得できませんでした');
    }

    const service = createSheetsService(accessToken);
    if (!service || !service.baseUrl || !service.accessToken) {
      errorLog('❌ サービスオブジェクト検証失敗:', {
        hasService: !!service,
        hasBaseUrl: !!(service && service.baseUrl),
        hasAccessToken: !!(service && service.accessToken)
      });
      throw new Error('Sheets APIサービスの初期化に失敗しました: 有効なサービスオブジェクトを作成できません');
    }

    // パフォーマンス最適化: 詳細ログを削減し、必要最小限の検証のみ実行
    debugLog('✅ サービスオブジェクト検証成功:', {
      hasBaseUrl: true,
      hasAccessToken: true,
      hasSpreadsheets: !!service.spreadsheets,
      hasGet: !!(service.spreadsheets && typeof service.spreadsheets.get === 'function'),
      baseUrl: service.baseUrl
    });

    // 簡略化された関数存在確認
    if (!service.spreadsheets || typeof service.spreadsheets.get !== 'function') {
      errorLog('❌ 重要な関数が失われています: SheetsServiceの基本機能が利用できません');
      throw new Error('SheetsServiceオブジェクトの関数が正しく設定されていません');
    }

    debugLog('✅ Sheetsサービス作成完了（検証済み）');
    return service;

  } catch (error) {
    errorLog('❌ getSheetsServiceCached error:', error.message);
    throw error;
  }
}

/**
 * 最適化されたSheetsサービスを取得
 * @returns {object} Sheets APIサービス
 */
function getSheetsService() {
  try {
    debugLog('🔧 getSheetsService: サービス取得開始');

    let accessToken;
    try {
      accessToken = getServiceAccountTokenCached();
    } catch (tokenError) {
      errorLog('❌ Failed to get service account token:', tokenError.message);
      throw new Error('サービスアカウントトークンの取得に失敗しました: ' + tokenError.message);
    }

    if (!accessToken) {
      errorLog('❌ Access token is null or undefined after generation.');
      throw new Error('サービスアカウントトークンが取得できませんでした。');
    }

    infoLog('✅ Access token obtained successfully');

    const service = createSheetsService(accessToken);
    if (!service || !service.baseUrl) {
      errorLog('❌ Failed to create sheets service or service object is invalid');
      throw new Error('Sheets APIサービスの初期化に失敗しました。');
    }

    infoLog('✅ Sheets service created successfully');
    debugLog('DEBUG: getSheetsService returning service object with baseUrl:', service.baseUrl);
    return service;

  } catch (error) {
    errorLog('❌ getSheetsService error:', error.message);
    errorLog('❌ Error stack:', error.stack);
    throw error; // エラーを再スロー
  }
}

/**
 * ユーザー情報を効率的に検索（キャッシュ優先）
 * @param {string} userId - ユーザーID
 * @returns {object|null} ユーザー情報
 */
function findUserById(userId) {
  // シンプルなキャッシュ戦略: 短時間の基本キャッシュのみ
  const cacheKey = 'user_' + userId;
  try {
    const cached = CacheService.getScriptCache().get(cacheKey);
    if (cached) {
      return JSON.parse(cached);
    }
  } catch (e) {
    // キャッシュエラーは無視して直接検索
  }

  const user = fetchUserFromDatabase('userId', userId);
  if (user) {
    try {
      CacheService.getScriptCache().put(cacheKey, JSON.stringify(user), 180); // 3分キャッシュ
    } catch (e) {
      // キャッシュ保存の失敗は無視
    }
  }
  return user;
}

/**
 * キャッシュをバイパスして最新のユーザー情報を強制取得
 * クリティカルな更新操作直後に使用
 * @param {string} userId - ユーザーID
 * @returns {object|null} 最新のユーザー情報
 */
function findUserByIdFresh(userId) {
  // 既存キャッシュを強制削除
  const cacheKey = 'user_' + userId;
  cacheManager.remove(cacheKey);

  // データベースから直接取得（範囲キャッシュも無効化）
  const freshUserInfo = fetchUserFromDatabase('userId', userId, {
    clearCache: true
  });

  // 次回の通常アクセス時にキャッシュされるため、ここでは手動設定不要
  infoLog('✅ Fresh user data retrieved for:', userId);

  return freshUserInfo;
}

/**
 * ユーザーデータの整合性を修正する
 * @param {string} userId - ユーザーID
 * @returns {object} 修正結果
 */
function fixUserDataConsistency(userId) {
  try {
    debugLog('🔧 ユーザーデータ整合性修正開始:', userId);

    // 最新のユーザー情報を取得
    const userInfo = findUserByIdFresh(userId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    debugLog('📊 現在のspreadsheetId:', userInfo.spreadsheetId);

    const configJson = JSON.parse(userInfo.configJson || '{}');
    debugLog('📝 configJson内のpublishedSpreadsheetId:', configJson.publishedSpreadsheetId);

    let needsUpdate = false;
    const updateData = {};

    // 1. エラー情報をクリーンアップ
    if (configJson.lastError || configJson.errorAt) {
      debugLog('🧹 エラー情報をクリーンアップ');
      delete configJson.lastError;
      delete configJson.errorAt;
      needsUpdate = true;
    }

    // 2. spreadsheetIdとpublishedSpreadsheetIdの整合性チェック
    if (userInfo.spreadsheetId && configJson.publishedSpreadsheetId !== userInfo.spreadsheetId) {
      debugLog('🔄 publishedSpreadsheetIdを実際のspreadsheetIdに合わせて修正');
      debugLog('  修正前:', configJson.publishedSpreadsheetId);
      debugLog('  修正後:', userInfo.spreadsheetId);

      configJson.publishedSpreadsheetId = userInfo.spreadsheetId;
      needsUpdate = true;
    }

    // 3. セットアップ状態の正規化（統一修復システムに委譲）
    // Note: 重複する修復ロジックはperformAutoHealing()に統合済み

    if (needsUpdate) {
      updateData.configJson = JSON.stringify(configJson);
      updateData.lastAccessedAt = new Date().toISOString();

      debugLog('💾 整合性修正のためデータベース更新実行');
      updateUser(userId, updateData);

      infoLog('✅ ユーザーデータ整合性修正完了');
      return { status: 'success', message: 'データ整合性が修正されました', updated: true };
    } else {
      infoLog('✅ データ整合性に問題なし');
      return { status: 'success', message: 'データ整合性に問題ありません', updated: false };
    }

  } catch (error) {
    errorLog('❌ データ整合性修正エラー:', error);
    return { status: 'error', message: 'データ整合性修正に失敗: ' + error.message };
  }
}

/**
 * メールアドレスでユーザー検索
 * @param {string} email - メールアドレス
 * @returns {object|null} ユーザー情報
 */
function findUserByEmail(email) {
  // シンプルなキャッシュ戦略: 短時間の基本キャッシュのみ
  const cacheKey = 'email_' + email;
  try {
    const cached = CacheService.getScriptCache().get(cacheKey);
    if (cached) {
      return JSON.parse(cached);
    }
  } catch (e) {
    // キャッシュエラーは無視して直接検索
  }

  const user = fetchUserFromDatabase('adminEmail', email);
  if (user) {
    try {
      CacheService.getScriptCache().put(cacheKey, JSON.stringify(user), 180); // 3分キャッシュ
    } catch (e) {
      // キャッシュ保存の失敗は無視
    }
  }
  return user;
}

// 以下の複雑な正規化関数は簡素化により削除されました:
// - _normalizeHeader
// - _normalizeValue  
// - _resolveFieldIndex
// 新しい fetchUserFromDatabase では直接的なフィールドマッチングを使用しています

/**
 * シンプル・信頼性重視のユーザーDB検索（統一版）
 * @param {string} field - 検索フィールド ('userId' | 'adminEmail')
 * @param {string} value - 検索値
 * @param {object} options - オプション設定
 * @returns {object|null} ユーザー情報
 */
function fetchUserFromDatabase(field, value, options = {}) {
  if (!field || !value) {
    debugLog('❌ fetchUserFromDatabase: 無効なパラメーター');
    return null;
  }

  const opts = {
    forceFresh: options.forceFresh === true,
    retryOnce: options.retryOnce === true,
    ...options
  };

  try {
    // 基本設定の取得
    const dbId = getSecureDatabaseId();
    if (!dbId) {
      throw new Error('データベースIDが設定されていません');
    }

    const service = getSheetsServiceCached();
    const sheetName = DB_SHEET_CONFIG.SHEET_NAME;

    // メールアドレス検索時の値の正規化
    let normalizedValue = value;
    if (field.toLowerCase().includes('email')) {
      normalizedValue = String(value).toLowerCase().trim();
      // 余分な空白や特殊文字の除去
      normalizedValue = normalizedValue.replace(/\s+/g, '').replace(/[\r\n\t]/g, '');
      if (normalizedValue !== value) {
        debugLog('fetchUserFromDatabase: メールアドレス正規化', {
          original: value,
          normalized: normalizedValue
        });
      }
    }

    debugLog('fetchUserFromDatabase: 検索開始', { 
      field, 
      value: normalizedValue, 
      forceFresh: opts.forceFresh 
    });

    // 強制フレッシュ時のピンポイントキャッシュ無効化
    if (opts.forceFresh) {
      try {
        // 1. UnifiedBatchProcessorキャッシュ無効化
        if (typeof unifiedBatchProcessor !== 'undefined' && unifiedBatchProcessor.invalidateCacheForSpreadsheet) {
          unifiedBatchProcessor.invalidateCacheForSpreadsheet(dbId);
        }
        
        // 2. 必要なキャッシュのみピンポイント削除
        const cache = CacheService.getScriptCache();
        // 対象ユーザーのキャッシュのみ削除
        if (field.toLowerCase().includes('email')) {
          cache.remove('email_' + normalizedValue);
          cache.remove('login_status_' + normalizedValue);
        }
        if (field.toLowerCase() === 'userid') {
          cache.remove('user_' + normalizedValue);
        }
        
        debugLog('fetchUserFromDatabase: 強制フレッシュ - ピンポイントキャッシュ無効化完了');
        
        // キャッシュ無効化後の短い待機（短縮）
        Utilities.sleep(50);
        
      } catch (cacheError) {
        warnLog('fetchUserFromDatabase: キャッシュ無効化でエラー:', cacheError.message);
      }
    }

    // データ取得オプション
    let cacheOptions = opts.forceFresh ? 
      { useCache: false, ttl: 0, forceRefresh: true } : 
      { useCache: true, ttl: 300 };

    // データ取得の信頼性確保（複数回試行）
    let data = null;
    let dataRetrieved = false;
    const maxDataRetries = opts.forceFresh ? 3 : 1;
    
    for (let dataAttempt = 0; dataAttempt < maxDataRetries && !dataRetrieved; dataAttempt++) {
      try {
        if (dataAttempt > 0) {
          debugLog('fetchUserFromDatabase: データ取得リトライ', dataAttempt);
          Utilities.sleep(200 * dataAttempt); // 段階的待機
        }
        
        data = batchGetSheetsData(service, dbId, [`'${sheetName}'!A:H`], cacheOptions);
        // 予防: 非同期オブジェクトが返ってきた場合は即時エラー
        if (data && typeof data.then === 'function') {
          throw new Error('無効なデータ構造を受信');
        }
        
        if (data && data.valueRanges && data.valueRanges[0]) {
          dataRetrieved = true;
        } else {
          throw new Error('無効なデータ構造を受信');
        }
      } catch (dataError) {
        warnLog('fetchUserFromDatabase: データ取得エラー (試行' + (dataAttempt + 1) + '):', dataError.message);
        // 構造不整合はプログラミングエラー相当として即時中断（深いリトライを避ける）
        const msg = String(dataError && dataError.message || '');
        if (msg.indexOf('無効なデータ構造') !== -1 || msg.toLowerCase().indexOf('invalid data structure') !== -1) {
          throw new Error('データ取得に失敗しました: ' + dataError.message);
        }
        // 400 INVALID_ARGUMENT 等の恒久的エラーも即時中断
        if (msg.indexOf('BatchGet失敗: 400') !== -1 || msg.indexOf('INVALID_ARGUMENT') !== -1) {
          throw new Error('データ取得に失敗しました: ' + dataError.message);
        }
        if (dataAttempt === maxDataRetries - 1) {
          throw new Error('データ取得に失敗しました: ' + dataError.message);
        }
      }
    }

    const values = data.valueRanges[0].values || [];

    if (values.length === 0) {
      throw new Error('データベースが空です');
    }

    debugLog('fetchUserFromDatabase: データ取得完了', { 
      rows: values.length, 
      forceFresh: opts.forceFresh,
      retries: maxDataRetries 
    });

    // ヘッダーとフィールドインデックスの解決（シンプル版）
    const headers = values[0];
    let fieldIndex = -1;
    
    // 直接的なフィールド名マッチング（正規化処理を簡素化）
    for (let i = 0; i < headers.length; i++) {
      const header = String(headers[i]).toLowerCase().trim();
      const searchField = field.toLowerCase().trim();
      
      if (header === searchField || 
          (searchField === 'adminemail' && (header === 'email' || header === 'adminemail')) ||
          (searchField === 'userid' && header === 'userid')) {
        fieldIndex = i;
        break;
      }
    }

    if (fieldIndex === -1) {
      throw new Error(`検索フィールド "${field}" が見つかりません。利用可能: ${headers.join(', ')}`);
    }

    debugLog('fetchUserFromDatabase: フィールド解決完了', { field, fieldIndex, header: headers[fieldIndex] });

    // ユーザー検索（シンプルな線形検索）
    for (let i = 1; i < values.length; i++) {
      const currentRow = values[i];
      const currentValue = currentRow[fieldIndex];

      if (!currentValue) continue;

      // 値の比較（メールアドレス強化版）
      let match = false;
      const currentStr = String(currentValue).trim();
      const searchStr = String(normalizedValue).trim();

      if (field.toLowerCase().includes('email')) {
        // メールアドレスの厳密な正規化比較
        const normalizedCurrent = currentStr.toLowerCase().replace(/\s+/g, '').replace(/[\r\n\t]/g, '');
        const normalizedSearch = searchStr.toLowerCase().replace(/\s+/g, '').replace(/[\r\n\t]/g, '');
        match = normalizedCurrent === normalizedSearch;
        
        // デバッグ情報（不一致時のみ詳細ログ）
        if (!match && i <= 5) { // 最初の5行のみデバッグ
          debugLog('fetchUserFromDatabase: メール比較詳細', {
            rowIndex: i,
            current: normalizedCurrent,
            search: normalizedSearch,
            rawCurrent: currentStr,
            rawSearch: searchStr
          });
        }
      } else {
        // その他のフィールドは厳密一致
        match = currentStr === searchStr;
      }

      if (match) {
        // ユーザーオブジェクトの構築
        const user = {};
        headers.forEach((header, index) => {
          user[header] = currentRow[index] || '';
        });

        // isActiveフィールドの正規化（ブール型に統一）
        if (user.hasOwnProperty('isActive')) {
          user.isActive = (user.isActive === true || user.isActive === 'true' || user.isActive === 'TRUE');
        }

        // データ整合性の最終確認（重要フィールドの検証）
        const dataIntegrity = {
          hasUserId: !!user.userId,
          hasEmail: !!user.adminEmail,
          fieldsComplete: !!(user.userId && user.adminEmail),
          searchMatch: user[field] === value || 
                      (field === 'adminEmail' && user.adminEmail.toLowerCase() === value.toLowerCase())
        };

        if (!dataIntegrity.fieldsComplete) {
          warnLog('fetchUserFromDatabase: データ整合性警告', {
            user: user,
            integrity: dataIntegrity,
            rowIndex: i
          });
        }

        infoLog('✅ fetchUserFromDatabase: ユーザー検索成功', { 
          field, 
          value, 
          userId: user.userId, 
          isActive: user.isActive,
          dataIntegrity: dataIntegrity
        });

        return user;
      }
    }

    // ユーザーが見つからない場合の詳細診断（強化版）
    const searchDiagnostics = {
      searchCriteria: { field, value: normalizedValue },
      originalValue: value !== normalizedValue ? value : undefined,
      totalRows: values.length - 1,
      fieldIndex: fieldIndex,
      headerInfo: { available: headers, matched: headers[fieldIndex] },
      sampleData: values.slice(1, Math.min(4, values.length)).map((row, idx) => ({
        rowIndex: idx + 1,
        targetValue: row[fieldIndex],
        rawRow: row.slice(0, 3) // 最初の3列のみ表示
      })),
      // メールアドレス検索時の特別診断
      emailDiagnostics: field.toLowerCase().includes('email') ? {
        normalizedSearch: normalizedValue,
        similarEmails: values.slice(1, Math.min(6, values.length))
          .map(row => row[fieldIndex])
          .filter(email => email && String(email).includes('@'))
          .map(email => ({
            original: email,
            normalized: String(email).toLowerCase().replace(/\s+/g, '').replace(/[\r\n\t]/g, ''),
            similarity: calculateEmailSimilarity(normalizedValue, String(email).toLowerCase())
          }))
          .sort((a, b) => b.similarity - a.similarity)
          .slice(0, 3)
      } : undefined
    };

    // メールアドレスの射似マッチ情報を追加ログ出力
    if (field.toLowerCase().includes('email') && searchDiagnostics.emailDiagnostics?.similarEmails?.length > 0) {
      const topSimilar = searchDiagnostics.emailDiagnostics.similarEmails[0];
      if (topSimilar.similarity > 0.7) {
        warnLog('🔍 fetchUserFromDatabase: 似たメールアドレスを発見', {
          search: normalizedValue,
          similar: topSimilar.original,
          similarity: topSimilar.similarity
        });
      }
    }
    
    warnLog('⚠️ fetchUserFromDatabase: ユーザーが見つかりません', searchDiagnostics);
    
    return null;

  } catch (error) {
    errorLog('❌ fetchUserFromDatabase エラー:', error.message);
    
    // 1回だけリトライオプションが有効で、まだリトライしていない場合
    if (opts.retryOnce && !opts._hasRetried) {
      warnLog('fetchUserFromDatabase: 1回リトライします...');
      Utilities.sleep(500); // 短い待機
      return fetchUserFromDatabase(field, value, { 
        ...opts, 
        _hasRetried: true, 
        forceFresh: true 
      });
    }
    
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
    warnLog('getUserWithFallback: Invalid userId:', userId);
    return null;
  }

  // 可能な限りキャッシュを利用
  const user = findUserById(userId);
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
    // カスタムセットアップ関連の更新の場合の詳細ログ
    if (updateData.spreadsheetId || updateData.folderId || (updateData.configJson && updateData.configJson.includes('formCreated'))) {
      infoLog('📋 updateUser: カスタムセットアップ関連の更新を開始', {
        userId: userId,
        hasSpreadsheetId: !!updateData.spreadsheetId,
        hasFolderId: !!updateData.folderId,
        hasConfigJson: !!updateData.configJson,
        updateFields: Object.keys(updateData)
      });
    }
    
    const props =  getResilientScriptProperties();
    const dbId =  getSecureDatabaseId();

    if (!dbId) {
      throw new Error('データ更新エラー: データベースIDが設定されていません');
    }
    // サービスインスタンスを一度だけ取得して再利用
    const service = getSheetsServiceCached();
    var sheetName = DB_SHEET_CONFIG.SHEET_NAME;

    // 現在のデータを取得
    const data =  batchGetSheetsData(service, dbId, ["'" + sheetName + "'!A:H"]);
    var values = data.valueRanges[0].values || [];

    if (values.length === 0) {
      throw new Error('データベースが空です');
    }

    const headers = values[0];
    var userIdIndex = headers.indexOf('userId');
    var rowIndex = -1;

    // ユーザーの行を特定
    for (let i = 1; i < values.length; i++) {
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
      debugLog('📝 データベース更新リクエスト詳細:');
      requests.forEach(function(req, index) {
        debugLog('  ' + (index + 1) + '. 範囲: ' + req.range + ', 値: ' + JSON.stringify(req.values));
      });

      var maxRetries = 2;
      var retryCount = 0;
      var updateSuccess = false;

      while (retryCount <= maxRetries && !updateSuccess) {
        try {
          if (retryCount > 0) {
            debugLog('🔄 認証エラーによるリトライ (' + retryCount + '/' + maxRetries + ')');
            // 認証エラーの場合のみ、新しいサービスを取得
            if (!service || retryCount === 1) {
              service = getSheetsServiceCached(true); // forceRefresh = true
              Utilities.sleep(1000); // 少し待機
            }
          }

          batchUpdateSheetsData(service, dbId, requests);
          infoLog('✅ データベース更新リクエスト送信完了');
          updateSuccess = true;

          // 更新成功の確認
          debugLog('🔍 更新直後の確認のため、データベースから再取得...');
          var verifyData = batchGetSheetsData(service, dbId, ["'" + sheetName + "'!" + String.fromCharCode(65 + userIdIndex) + rowIndex + ":" + String.fromCharCode(72) + rowIndex]);
          if (verifyData.valueRanges && verifyData.valueRanges[0] && verifyData.valueRanges[0].values) {
            var updatedRow = verifyData.valueRanges[0].values[0];
            debugLog('📊 更新後のユーザー行データ:', updatedRow);
            if (updateData.spreadsheetId) {
              var spreadsheetIdIndex = headers.indexOf('spreadsheetId');
              debugLog('🎯 スプレッドシートID更新確認:', updatedRow[spreadsheetIdIndex] === updateData.spreadsheetId ? '✅ 成功' : '❌ 失敗');
            }
          }

        } catch (updateError) {
          retryCount++;
          var errorMessage = updateError.toString();

          if (errorMessage.includes('401') || errorMessage.includes('UNAUTHENTICATED') || errorMessage.includes('ACCESS_TOKEN_EXPIRED')) {
            warnLog('🔐 認証エラーを検出:', errorMessage);

            if (retryCount <= maxRetries) {
              debugLog('🔄 認証トークンをリフレッシュしてリトライします...');
              continue; // リトライループを続行
            } else {
              errorLog('❌ 最大リトライ回数に達しました');
              throw new Error('認証エラーにより更新に失敗しました（最大リトライ回数超過）');
            }
          } else {
            // 認証エラー以外のエラーはすぐに終了
            errorLog('❌ データベース更新中にエラーが発生:', updateError);
            throw new Error('データベース更新に失敗しました: ' + updateError.message);
          }
        }
      }

      // 更新が確実に反映されるまで少し待機
      Utilities.sleep(100);
    } else {
      debugLog('⚠️ 更新するデータがありません');
    }

    // スプレッドシートIDが更新された場合、サービスアカウントと共有
    if (updateData.spreadsheetId) {
      try {
        shareSpreadsheetWithServiceAccount(updateData.spreadsheetId);
        debugLog('ユーザー更新時のサービスアカウント共有完了:', updateData.spreadsheetId);
      } catch (shareError) {
        errorLog('ユーザー更新時のサービスアカウント共有エラー:', shareError.message);
        errorLog('スプレッドシートの更新は完了しましたが、サービスアカウントとの共有に失敗しました。手動で共有してください。');
      }
    }

    // 重要: 更新完了後に包括的キャッシュ同期を実行
    var userInfo = findUserByIdFresh(userId); // findUserByIdFresh を使用
    var email = updateData.adminEmail || (userInfo ? userInfo.adminEmail : null);
    var oldSpreadsheetId = userInfo ? userInfo.spreadsheetId : null;
    var newSpreadsheetId = updateData.spreadsheetId || oldSpreadsheetId;

    // クリティカル更新時の包括的キャッシュ同期
    synchronizeCacheAfterCriticalUpdate(userId, email, oldSpreadsheetId, newSpreadsheetId);

    // カスタムセットアップ関連の更新の場合の完了ログ
    if (updateData.spreadsheetId || updateData.folderId || (updateData.configJson && updateData.configJson.includes('formCreated'))) {
      infoLog('✅ updateUser: カスタムセットアップ関連の更新が完了しました', {
        userId: userId,
        spreadsheetId: updateData.spreadsheetId || 'unchanged',
        folderId: updateData.folderId || 'unchanged',
        hasConfigJson: !!updateData.configJson
      });
    }

    return { status: 'success', message: 'ユーザープロフィールが正常に更新されました' };
  } catch (error) {
    errorLog('ユーザー更新エラー:', error);
    throw error;
  }
}

/**
 * 新規ユーザーをデータベースに作成
 * @param {object} userData - 作成するユーザーデータ
 * @returns {object} 作成されたユーザーデータ
 */
function createUser(userData) {
  // 統一ロック管理で重複登録を防ぐ
  return executeWithStandardizedLock('WRITE_OPERATION', 'createUser', () => {
    // メールアドレスの重複チェック
    var existingUser = findUserByEmail(userData.adminEmail);
    if (existingUser) {
      throw new Error('このメールアドレスは既に登録されています。');
    }

    const props =  getResilientScriptProperties();
    const dbId =  getSecureDatabaseId();
    const service = getSheetsServiceCached();
    var sheetName = DB_SHEET_CONFIG.SHEET_NAME;

    var newRow = DB_SHEET_CONFIG.HEADERS.map(function(header) {
      return userData[header] || '';
    });

    debugLog('createUser - デバッグ: ヘッダー構成=' + JSON.stringify(DB_SHEET_CONFIG.HEADERS));
    debugLog('createUser - デバッグ: ユーザーデータ=' + JSON.stringify(userData));
    debugLog('createUser - デバッグ: 作成される行データ=' + JSON.stringify(newRow));

    appendSheetsData(service, dbId, "'" + sheetName + "'!A1", [newRow]);

    debugLog('createUser - データベース書き込み完了: userId=' + userData.userId);

    // 書き込み直後の読み取り一貫性を高めるためにキャッシュを明示的に無効化し、短い待機を入れる
    try {
      if (typeof unifiedBatchProcessor !== 'undefined' && unifiedBatchProcessor.invalidateCacheForSpreadsheet) {
        unifiedBatchProcessor.invalidateCacheForSpreadsheet(dbId);
        debugLog('createUser - 統一バッチキャッシュを無効化: ' + dbId);
      }
    } catch (cacheInvalidationError) {
      warnLog('createUser - キャッシュ無効化で警告:', cacheInvalidationError.message);
    }
    Utilities.sleep(150);

    // 新規ユーザー用の専用フォルダを作成
    try {
      debugLog('createUser - 専用フォルダ作成開始: ' + userData.adminEmail);
      var folder = createUserFolder(userData.adminEmail);
      if (folder) {
        infoLog('✅ createUser - 専用フォルダ作成成功: ' + folder.getName());
      } else {
        debugLog('⚠️ createUser - 専用フォルダ作成失敗（処理は続行）');
      }
    } catch (folderError) {
      warnLog('createUser - フォルダ作成でエラーが発生しましたが、処理を続行します: ' + folderError.message);
    }

    // 最適化: 新規ユーザー作成時は対象キャッシュのみ無効化
    invalidateUserCache(userData.userId, userData.adminEmail, userData.spreadsheetId, false);

    return userData;
  });
}

// waitForUserRecord関数は簡素化により削除されました
// registerNewUser内のシンプルな検証ループで置き換えられています

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
      debugLog('📊 バッチ最適化: シート作成+ヘッダー追加を同時実行');

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

      infoLog('✅ バッチ最適化完了: シート作成+ヘッダー追加（2回のAPI呼び出し→シーケンシャル実行）');

    } else {
      // シートが既に存在する場合は、ヘッダーのみ更新（既存動作を維持）
      var headerRange = "'" + sheetName + "'!A1:" + String.fromCharCode(65 + DB_SHEET_CONFIG.HEADERS.length - 1) + '1';
      updateSheetsData(service, spreadsheetId, headerRange, [DB_SHEET_CONFIG.HEADERS]);
      infoLog('✅ 既存シートのヘッダー更新完了');
    }

    debugLog('データベースシート「' + sheetName + '」の初期化が完了しました。');
  } catch (e) {
    errorLog('データベースシートの初期化に失敗: ' + e.message);
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
        warnLog('Failed to clear user properties:', propsError.message);
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
    errorLog('handleMissingUser error:', error.message);
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

          var response =  resilientUrlFetch(url, {
            headers: { 'Authorization': 'Bearer ' + accessToken },
            muteHttpExceptions: true,
            followRedirects: true,
            validateHttpsCertificates: true
          });

          // レスポンスオブジェクト検証
          if (!response || typeof response.getResponseCode !== 'function') {
            throw new Error('Sheets API: 無効なレスポンスオブジェクトが返されました');
          }
          
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

        var response =  resilientUrlFetch(url, {
          headers: { 'Authorization': 'Bearer ' + accessToken },
          muteHttpExceptions: true,
          followRedirects: true,
          validateHttpsCertificates: true
        });

        // レスポンスオブジェクト検証
        if (!response || typeof response.getResponseCode !== 'function') {
          throw new Error('Sheets API: 無効なレスポンスオブジェクトが返されました');
        }
        
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
function batchGetSheetsData(service, spreadsheetId, ranges, options = {}) {
  // 同期実装（UrlFetchApp.fetch を使用）に統一
  debugLog('DEBUG: batchGetSheetsData - 同期実装で実行');

  // 入力検証
  if (!service || typeof service !== 'object') {
    throw new Error('Sheetsサービスオブジェクトが提供されていません');
  }
  if (!service.baseUrl || !service.accessToken) {
    throw new Error('無効なSheetsサービスオブジェクトです: baseUrl/accessTokenが見つかりません');
  }
  if (!spreadsheetId || typeof spreadsheetId !== 'string') {
    throw new Error('無効なspreadsheetIDです: 文字列である必要があります');
  }
  if (!/^[a-zA-Z0-9_-]{20,}$/.test(spreadsheetId)) {
    throw new Error('無効なspreadsheetID形式です');
  }
  if (!ranges || !Array.isArray(ranges) || ranges.length === 0) {
    throw new Error('無効な範囲配列です: 配列で1つ以上の範囲が必要です');
  }

  const opts = {
    valueRenderOption: 'UNFORMATTED_VALUE',
    dateTimeRenderOption: 'SERIAL_NUMBER',
    // includeGridData は values:batchGet では無効（spreadsheets.get 系のみ）
    ...options
  };

  // リクエストURL作成
  const endpoint = service.baseUrl + '/' + encodeURIComponent(spreadsheetId) + '/values:batchGet';
  const params = [];
  for (let i = 0; i < ranges.length; i++) {
    params.push('ranges=' + encodeURIComponent(ranges[i]));
  }
  params.push('valueRenderOption=' + encodeURIComponent(opts.valueRenderOption));
  params.push('dateTimeRenderOption=' + encodeURIComponent(opts.dateTimeRenderOption));
  // includeGridData は使用しない

  const url = endpoint + '?' + params.join('&');

  // リクエスト実行（同期）
  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { 'Authorization': 'Bearer ' + service.accessToken },
    muteHttpExceptions: true
  });

  if (!response || typeof response.getResponseCode !== 'function') {
    throw new Error('無効なレスポンスオブジェクトが返されました');
  }

  const code = response.getResponseCode();
  const text = response.getContentText();
  if (code !== 200) {
    throw new Error('BatchGet失敗: ' + code + ' - ' + text);
  }

  let body;
  try {
    body = JSON.parse(text);
  } catch (e) {
    throw new Error('BatchGetレスポンスのJSON解析に失敗しました');
  }

  if (!body || !Array.isArray(body.valueRanges)) {
    throw new Error('無効なデータ構造を受信');
  }

  return {
    spreadsheetId: spreadsheetId,
    valueRanges: body.valueRanges
  };
}

/**
 * バッチ更新
 * @param {object} service - Sheetsサービス
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {object[]} requests - 更新リクエストの配列
 * @returns {object} レスポンス
 */
function batchUpdateSheetsData(service, spreadsheetId, requests) {
  // 同期実装（UrlFetchApp.fetch）
  if (!service || !service.baseUrl || !service.accessToken) {
    throw new Error('無効なSheetsサービスオブジェクトです: baseUrl/accessTokenが見つかりません');
  }
  if (!spreadsheetId || typeof spreadsheetId !== 'string') {
    throw new Error('無効なspreadsheetIDです');
  }
  if (!Array.isArray(requests) || requests.length === 0) {
    throw new Error('更新データが無効です');
  }

  const url = service.baseUrl + '/' + encodeURIComponent(spreadsheetId) + '/values:batchUpdate';
  const body = {
    valueInputOption: 'RAW',
    includeValuesInResponse: false,
    data: requests
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + service.accessToken },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });

  if (!response || typeof response.getResponseCode !== 'function') {
    throw new Error('無効なレスポンスオブジェクトが返されました');
  }
  const code = response.getResponseCode();
  if (code !== 200) {
    throw new Error('BatchUpdate失敗: ' + code + ' - ' + response.getContentText());
  }

  // 必要最小限の情報を返す
  try {
    return JSON.parse(response.getContentText());
  } catch (e) {
    return { status: 'success' };
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
  try {
    const url = service.baseUrl + '/' + spreadsheetId + '/values/' + encodeURIComponent(range) +
      ':append?valueInputOption=RAW&insertDataOption=INSERT_ROWS';

    debugLog('appendSheetsData: 書き込み開始', { 
      spreadsheetId: spreadsheetId.substring(0, 20) + '...', 
      range, 
      valueCount: values.length 
    });

    // 書き込み実行
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + service.accessToken },
      payload: JSON.stringify({ values: values }),
      muteHttpExceptions: true
    });

    // レスポンス検証
    if (!response || typeof response.getResponseCode !== 'function') {
      throw new Error('無効なレスポンスオブジェクトが返されました');
    }

    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode !== 200) {
      throw new Error(`書き込み失敗: ${responseCode} - ${responseText}`);
    }

    // レスポンス解析と詳細確認
    let parsed;
    try {
      parsed = JSON.parse(responseText);
    } catch (parseError) {
      throw new Error(`レスポンス解析失敗: ${parseError.message}`);
    }

    // 書き込み成功の詳細確認
    const writeConfirmation = {
      success: !!parsed.updates,
      updatedRows: parsed.updates?.updatedRows || 0,
      updatedColumns: parsed.updates?.updatedColumns || 0,
      updatedCells: parsed.updates?.updatedCells || 0,
      updatedRange: parsed.updates?.updatedRange || null
    };

    infoLog('✅ appendSheetsData: 書き込み成功確認', writeConfirmation);

    // 書き込み成功が確認できない場合はエラー
    if (!writeConfirmation.success || writeConfirmation.updatedRows === 0) {
      throw new Error('書き込み応答は成功だが、実際の更新が確認できません');
    }

    // 書き込み後の効率的キャッシュ無効化
    try {
      // 1. UnifiedBatchProcessorキャッシュ無効化
      if (typeof unifiedBatchProcessor !== 'undefined' && unifiedBatchProcessor.invalidateCacheForSpreadsheet) {
        unifiedBatchProcessor.invalidateCacheForSpreadsheet(spreadsheetId);
      }
      
      // 2. 新規ユーザー登録時のみ範囲を限定したキャッシュクリア
      try {
        const cache = CacheService.getScriptCache();
        // 新規登録時のみ必要なキャッシュをクリアするように最適化
        // 全キャッシュクリアは重いので節約
        debugLog('appendSheetsData: 最小限キャッシュクリア実行');
      } catch (cacheErr) {
        warnLog('appendSheetsData: キャッシュクリアでエラー:', cacheErr.message);
      }
      
      infoLog('✅ appendSheetsData: 最適化されたキャッシュ無効化完了');
      
    } catch (invalidateErr) {
      warnLog('appendSheetsData: キャッシュ無効化でエラー:', invalidateErr.message);
      // キャッシュ無効化の失敗は書き込み成功を妨げない
    }

    // 書き込み完了確認のための短い待機（短縮）
    Utilities.sleep(100); // 100msに短縮してパフォーマンス向上

    debugLog('✅ appendSheetsData: 書き込み処理完了');
    return parsed;

  } catch (error) {
    errorLog('❌ appendSheetsData エラー:', {
      error: error.message,
      spreadsheetId: spreadsheetId.substring(0, 20) + '...',
      range,
      valueCount: values ? values.length : 0
    });

    // 重要なエラーの場合のみ1回リトライ
    if (error.message.includes('429') || error.message.includes('503') || error.message.includes('timeout')) {
      try {
        warnLog('appendSheetsData: API制限/サーバーエラーのため2秒後にリトライします...');
        Utilities.sleep(2000);
        
        const retryResponse = UrlFetchApp.fetch(url, {
          method: 'post',
          contentType: 'application/json',
          headers: { 'Authorization': 'Bearer ' + service.accessToken },
          payload: JSON.stringify({ values: values }),
          muteHttpExceptions: true
        });

        if (retryResponse.getResponseCode() === 200) {
          infoLog('✅ appendSheetsData: リトライ成功');
          const retryParsed = JSON.parse(retryResponse.getContentText());
          
          // リトライ成功時もキャッシュ無効化
          try {
            if (typeof unifiedBatchProcessor !== 'undefined' && unifiedBatchProcessor.invalidateCacheForSpreadsheet) {
              unifiedBatchProcessor.invalidateCacheForSpreadsheet(spreadsheetId);
            }
          } catch (invalidateErr) {
            // キャッシュ無効化エラーは無視
          }
          
          return retryParsed;
        }
      } catch (retryError) {
        errorLog('❌ appendSheetsData: リトライも失敗:', retryError.message);
      }
    }

    throw error;
  }
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
      warnLog('⚠️ baseUrlが見つかりません。デフォルトのGoogleSheetsAPIエンドポイントを使用します');
      baseUrl = 'https://sheets.googleapis.com/v4/spreadsheets';
    }

    if (!accessToken || typeof accessToken !== 'string') {
      throw new Error('アクセストークンが見つかりません。サービスオブジェクトが破損している可能性があります');
    }

    debugLog('DEBUG: getSpreadsheetsData - 使用するbaseUrl:', baseUrl);

    // 安全なURL構築 - シート情報を含む基本的なメタデータを取得するために fields パラメータを追加
    var url = baseUrl + '/' + encodeURIComponent(spreadsheetId) + '?fields=sheets.properties';

    debugLog('DEBUG: getSpreadsheetsData - 構築されたURL:', url.substring(0, 100) + '...');

    var response = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + accessToken },
      muteHttpExceptions: true,
      followRedirects: true,
      validateHttpsCertificates: true
    });

    // レスポンスオブジェクト検証
    if (!response || typeof response.getResponseCode !== 'function') {
      throw new Error('Database API: 無効なレスポンスオブジェクトが返されました');
    }
    
    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();

    if (responseCode !== 200) {
      errorLog('Sheets API エラー詳細:', {
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
          warnLog('エラーレスポンスのJSON解析に失敗:', parseError.message);
        }
      }

      throw new Error('Sheets API error: ' + responseCode + ' - ' + responseText);
    }

    var result;
    try {
      result = JSON.parse(responseText);
    } catch (parseError) {
      errorLog('❌ JSON解析エラー:', parseError.message);
      errorLog('❌ Response text:', responseText.substring(0, 200));
      throw new Error('APIレスポンスのJSON解析に失敗: ' + parseError.message);
    }

    // レスポンス構造の検証
    if (!result || typeof result !== 'object') {
      throw new Error('無効なAPIレスポンス: オブジェクトが期待されましたが ' + typeof result + ' を受信');
    }

    if (!result.sheets || !Array.isArray(result.sheets)) {
      warnLog('⚠️ sheets配列が見つからないか、配列でありません:', typeof result.sheets);
      result.sheets = []; // 空配列を設定してエラーを避ける
    }

    var sheetCount = result.sheets.length;
    infoLog('✅ getSpreadsheetsData 成功: 発見シート数:', sheetCount);

    if (sheetCount === 0) {
      warnLog('⚠️ スプレッドシートにシートが見つかりませんでした');
    } else {
      debugLog('📋 利用可能なシート:', result.sheets.map(s => s.properties?.title || 'Unknown').join(', '));
    }

    return result;

  } catch (error) {
    errorLog('❌ getSpreadsheetsData error:', error.message);
    errorLog('❌ Error stack:', error.stack);
    throw new Error('スプレッドシート情報取得に失敗しました: ' + error.message);
  }
}

/**
 * すべてのユーザー情報を取得
 * @returns {Array} ユーザー情報配列
 */
function getAllUsers() {
  try {
    const props =  getResilientScriptProperties();
    const dbId =  getSecureDatabaseId();
    const service = getSheetsServiceCached();
    var sheetName = DB_SHEET_CONFIG.SHEET_NAME;

    const data =  batchGetSheetsData(service, dbId, ["'" + sheetName + "'!A:H"]);
    var values = data.valueRanges[0].values || [];

    if (values.length <= 1) {
      return []; // ヘッダーのみの場合は空配列を返す
    }

    const headers = values[0];
    var users = [];

    for (let i = 1; i < values.length; i++) {
      var row = values[i];
      var user = {};

      for (var j = 0; j < headers.length; j++) {
        user[headers[j]] = row[j] || '';
      }

      users.push(user);
    }

    return users;

  } catch (error) {
    errorLog('getAllUsers エラー:', error.message);
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
  // 統一バッチ処理システムを使用
  const requests = requestBody.requests || [];
  return  unifiedBatchProcessor.batchUpdateSpreadsheet(service, spreadsheetId, requests, {
    includeSpreadsheetInResponse: requestBody.includeSpreadsheetInResponse || false,
    responseRanges: requestBody.responseRanges || [],
    invalidateCache: true
  });
}

/**
 * データベース状態を詳細確認する診断機能
 * @param {string} targetUserId - 確認対象のユーザーID（オプション）
 * @returns {object} データベース診断結果
 */
function diagnoseDatabase(targetUserId) {
  try {
    debugLog('🔍 データベース診断開始:', targetUserId || 'ALL_USERS');

    const props =  getResilientScriptProperties();
    const dbId =  getSecureDatabaseId();

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
      const values = data.valueRanges[0].values || [];

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

        for (let i = 1; i < values.length; i++) {
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

    infoLog('✅ データベース診断完了:', diagnosticResult.summary.overallStatus);
    return diagnosticResult;

  } catch (error) {
    errorLog('❌ データベース診断でエラー:', error);
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
      warnLog('キャッシュ統計取得エラー:', statsError.message);
    }

    return cacheStatus;

  } catch (error) {
    errorLog('キャッシュ状態チェックエラー:', error);
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
    debugLog('🔐 サービスアカウント権限確認開始:', spreadsheetId || 'DATABASE');

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
    var dbId = spreadsheetId ||  getSecureDatabaseId();

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
      const service = getSheetsServiceCached();
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
        debugLog('🔧 権限修復を試行中...');

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

    infoLog('✅ サービスアカウント権限確認完了:', result.summary.status);
    return result;

  } catch (error) {
    errorLog('❌ サービスアカウント権限確認でエラー:', error);
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
    debugLog('🔧 自動修復開始:', targetUserId || 'GENERAL');

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

    infoLog('✅ 自動修復完了:', repairResult.success ? '成功' : '要追加対応');
    return repairResult;

  } catch (error) {
    errorLog('❌ 自動修復でエラー:', error);
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
    debugLog('🔍 データ整合性チェック開始');

    const opts = {
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
    const props =  getResilientScriptProperties();
    const dbId =  getSecureDatabaseId();
    if (!dbId) {
      result.summary.issues.push('データベースIDが設定されていません');
      result.summary.status = 'critical';
      return result;
    }

    const service = getSheetsServiceCached();
    var data = batchGetSheetsData(service, dbId, ["'" + DB_SHEET_CONFIG.SHEET_NAME + "'!A:H"]);
    var values = data.valueRanges[0].values || [];

    if (values.length <= 1) {
      result.summary.status = 'empty';
      return result;
    }

    const headers = values[0];
    var userRows = values.slice(1);
    result.summary.totalUsers = userRows.length;

    debugLog('📊 データ整合性チェック: ' + userRows.length + 'ユーザーを確認中');

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
        debugLog('🔧 自動修復完了: ' + fixResult.fixed.length + '件修復');
      } catch (fixError) {
        errorLog('❌ 自動修復エラー:', fixError.message);
        result.summary.issues.push('自動修復に失敗: ' + fixError.message);
      }
    }

    // 6. 最終判定
    if (result.summary.issues.length === 0) {
      result.summary.status = result.summary.warnings.length > 0 ? 'warning' : 'healthy';
    } else {
      result.summary.status = 'critical';
    }

    infoLog('✅ データ整合性チェック完了:', result.summary.status);
    return result;

  } catch (error) {
    errorLog('❌ データ整合性チェックでエラー:', error);
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
    warnLog('⚠️ 重複データが検出されましたが、自動削除は行いません:', details.duplicates.length + '件');
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
        errorLog('❌ isActiveフィールド修正エラー:', updateError.message);
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
    const props =  getResilientScriptProperties();
    const dbId =  getSecureDatabaseId();
    if (!dbId) {
      throw new Error('データベースIDが設定されていません');
    }

    var ss = openSpreadsheetOptimized(dbId);
    var sheet = ss.getSheetByName(DB_SHEET_CONFIG.SHEET_NAME);
    if (!sheet) {
      throw new Error('データベースシートが見つかりません: ' + DB_SHEET_CONFIG.SHEET_NAME);
    }

    return sheet;
  } catch (e) {
    errorLog('getDbSheet error:', e.message);
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
    debugLog('📊 システム監視開始');

    const opts = {
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
        errorLog('❌ ヘルスチェックエラー:', healthError.message);
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
        errorLog('❌ パフォーマンスチェックエラー:', perfError.message);
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
        errorLog('❌ セキュリティチェックエラー:', securityError.message);
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
        errorLog('❌ アラート送信エラー:', alertError.message);
      }
    }

    infoLog('✅ システム監視完了:', monitoringResult.summary.overallHealth);
    return monitoringResult;

  } catch (error) {
    errorLog('❌ システム監視でエラー:', error);
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
    const props =  getResilientScriptProperties();
    const dbId =  getSecureDatabaseId();

    if (dbId) {
      const service = getSheetsServiceCached();
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
      warnLog('キャッシュ統計取得エラー:', cacheError.message);
    }
    perfResult.benchmarks.cacheCheck = Date.now() - cacheTestStart;

  } catch (perfError) {
    errorLog('パフォーマンステスト実行エラー:', perfError.message);
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
    errorLog('セキュリティチェック実行エラー:', securityError.message);
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
    debugLog('🚨 システムアラート送信開始');

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
      debugLog('📧 管理者アラート（' + adminEmail + '）:\n' + alertMessage);

      // 緊急レベルの場合は追加のログ記録
      if (monitoringResult.summary.alerts.length > 0) {
        errorLog('🚨 緊急システムアラート: ' + monitoringResult.summary.alerts.join(', '));
      }
    }

    // システムログへの記録
    logSystemEvent('system_alert', {
      level: monitoringResult.summary.overallHealth,
      alerts: monitoringResult.summary.alerts,
      warnings: monitoringResult.summary.warnings,
      timestamp: monitoringResult.timestamp
    });

    infoLog('✅ システムアラート送信完了');

  } catch (alertError) {
    errorLog('❌ アラート送信エラー:', alertError.message);
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
    debugLog('📝 システムイベント [' + eventType + ']:', JSON.stringify(eventData));

    // 将来的にはログシートやログファイルに保存する機能を追加可能

  } catch (logError) {
    errorLog('システムイベントログエラー:', logError.message);
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

    // 統一ロック管理でアカウント削除実行
    return executeWithStandardizedLock('CRITICAL_OPERATION', 'deleteUserAccount', () => {
      // データベース（シート）からユーザー行を削除（サービスアカウント経由）
      var props = PropertiesService.getScriptProperties();
      var dbId =  getSecureDatabaseId();
      if (!dbId) {
        throw new Error('データベースIDが設定されていません');
      }

      const service = getSheetsServiceCached();
      if (!service) {
        throw new Error('Sheets APIサービスの初期化に失敗しました');
      }

      const sheetName = DB_SHEET_CONFIG.SHEET_NAME;

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

      debugLog('Found database sheet with sheetId:', targetSheetId);

      // データを取得
      const data =  batchGetSheetsData(service, dbId, ["'" + sheetName + "'!A:H"]);
      const values = data.valueRanges[0].values || [];

      // ユーザーIDに基づいて行を探す（フィールドベース検索）
      const headers = values[0];
      let userIdFieldIndex = -1;
      
      // userIdフィールドのインデックスを特定
      for (let j = 0; j < headers.length; j++) {
        if (headers[j] === 'userId') {
          userIdFieldIndex = j;
          break;
        }
      }
      
      if (userIdFieldIndex === -1) {
        throw new Error('データベースに userId フィールドが見つかりません');
      }
      
      debugLog('Found userId field at index:', userIdFieldIndex);
      
      var rowToDelete = -1;
      for (let i = values.length - 1; i >= 1; i--) {
        if (values[i][userIdFieldIndex] === userId) {
          rowToDelete = i + 1; // スプレッドシートは1ベース
          debugLog('Found user row to delete at index:', i, 'rowToDelete:', rowToDelete);
          break;
        }
      }

      if (rowToDelete !== -1) {
        debugLog('Deleting row:', rowToDelete, 'from sheetId:', targetSheetId);

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

        debugLog('Row deletion completed successfully');
      } else {
        // 削除対象の行が見つからない場合はエラーとして扱う
        const errorMessage = `削除対象のユーザー行が見つかりません。userId: ${userId}`;
        errorLog(errorMessage);
        throw new Error(errorMessage);
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
      invalidateUserCache(userId, userInfo.adminEmail, userInfo.spreadsheetId, false, dbId);

      // Google Drive のデータは保持するため何も操作しない

      // UserPropertiesからも関連情報を削除
      const userProps = PropertiesService.getUserProperties();
      userProps.deleteProperty('CURRENT_USER_ID');

      const successMessage = 'アカウント「' + userInfo.adminEmail + '」が正常に削除されました。';
      infoLog(successMessage);
      return successMessage;
    });

  } catch (error) {
    errorLog('アカウント削除エラー:', error.message);
    errorLog('アカウント削除エラー詳細:', error.stack);

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
 * メールアドレスの類似度計算（簡易版）
 * @param {string} email1 - 比較元メールアドレス
 * @param {string} email2 - 比較先メールアドレス
 * @returns {number} 類似度（0-1）
 */
function calculateEmailSimilarity(email1, email2) {
  if (!email1 || !email2) return 0;
  
  const e1 = String(email1).toLowerCase().trim();
  const e2 = String(email2).toLowerCase().trim();
  
  if (e1 === e2) return 1.0;
  
  // Levenshtein距離ベースの類似度計算（簡略版）
  const maxLen = Math.max(e1.length, e2.length);
  if (maxLen === 0) return 1.0;
  
  const distance = levenshteinDistance(e1, e2);
  return Math.max(0, 1 - distance / maxLen);
}

/**
 * Levenshtein距離計算（簡易版）
 * @param {string} str1 - 文字列1
 * @param {string} str2 - 文字列2
 * @returns {number} 編集距離
 */
function levenshteinDistance(str1, str2) {
  const matrix = [];
  
  // 初期化
  for (let i = 0; i <= str2.length; i++) {
    matrix[i] = [i];
  }
  
  for (let j = 0; j <= str1.length; j++) {
    matrix[0][j] = j;
  }
  
  // 計算
  for (let i = 1; i <= str2.length; i++) {
    for (let j = 1; j <= str1.length; j++) {
      if (str2.charAt(i - 1) === str1.charAt(j - 1)) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j - 1] + 1, // 置換
          matrix[i][j - 1] + 1,     // 挿入
          matrix[i - 1][j] + 1      // 削除
        );
      }
    }
  }
  
  return matrix[str2.length][str1.length];
}
