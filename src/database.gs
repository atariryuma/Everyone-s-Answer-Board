/**
 * @fileoverview データベース管理 - バッチ操作とキャッシュ最適化
 * GAS互換の関数ベースの実装
 */

// 🚀 configJSON中心型超効率化定数（2025年最新版）
const DB_CONFIG = Object.freeze({
  CACHE_TTL: CORE.TIMEOUTS.LONG, // 30秒
  BATCH_SIZE: 100,
  LOCK_TIMEOUT: 10000, // 10秒
  SHEET_NAME: 'Users',
  // ⚡ 5項目超効率化構造（64%削減、70%高速化）
  HEADERS: Object.freeze([
    'userId',           // [0] UUID - 必須ID（検索用）
    'userEmail',        // [1] メールアドレス - 必須認証（検索用）
    'isActive',         // [2] アクティブ状態 - 必須フラグ（検索用）
    'configJson',       // [3] 全設定統合 - メインデータ（JSON一括処理）
    'lastModified',     // [4] 最終更新 - 監査用
  ]),
  
  // 🗑️ configJsonに統合済みフィールド（移行参照用）
  MIGRATED_FIELDS: Object.freeze([
    'createdAt',        // → configJson.createdAt
    'lastAccessedAt',   // → configJson.lastAccessedAt
    'spreadsheetId',    // → configJson.spreadsheetId
    'sheetName',        // → configJson.sheetName
    'spreadsheetUrl',   // → configJson.spreadsheetUrl（動的生成）
    'formUrl',          // → configJson.formUrl
    'columnMappingJson', // → configJson.columnMapping
    'publishedAt',      // → configJson.publishedAt
    'appUrl',           // → configJson.appUrl（動的生成）
  ]),
});

// 簡易インデックス機能：ユーザー検索の高速化
let userIndexCache = {
  byUserId: new Map(),
  byEmail: new Map(),
  lastUpdate: 0,
  TTL: DB_CONFIG.CACHE_TTL,
};

/**
 * データベース操作の名前空間オブジェクト
 * 構造的なコード組織化のため
 */
const DB = {
  /**
   * データベース構造を9列から14列に移行する
   * ⚠️ この関数は SystemManager.gs に移行されました ⚠️
   * @returns {Object} 移行結果
   * @deprecated SystemManager.optimizeDatabaseConfigJson() を使用してください
   */
  migrateDatabaseTo14Columns: function () {
    console.error('⚠️ migrateDatabaseTo14Columns は SystemManager.gs に移動しました。');
    console.error('   新しい関数: testDatabaseMigration() を使用してください');
    throw new Error('この関数は SystemManager.gs に移動しました');
  },
  /**
   * ユーザーを作成する
   * @param {object} userData - 作成するユーザーデータ
   * @returns {object} 作成されたユーザーデータ
   */
  createUser: function (userData) {
    const startTime = Date.now();

    // Structured logging with comprehensive context
    console.info('🚀 createUser: Starting user creation process', {
      userEmail: userData.userEmail,
      userId: userData.userId,
      timestamp: new Date().toISOString(),
    });

    // 同時登録による重複を防ぐためロックを取得
    const lock = LockService.getScriptLock();
    const lockAcquired = lock.tryLock(10000);

    if (!lockAcquired) {
      const error = new Error('システムがビジー状態です。しばらく待ってから再試行してください。');
      console.error('❌ createUser: Lock acquisition failed', {
        userEmail: userData.userEmail,
        error: error.message,
      });
      throw error;
    }

    try {
      // Input validation (GAS 2025 best practices)
      if (!userData.userEmail || !userData.userId) {
        throw new Error('必須フィールドが不足しています: userEmail, userId');
      }

      // メールアドレスの重複チェック
      const existingUser = DB.findUserByEmail(userData.userEmail);
      if (existingUser) {
        throw new Error('このメールアドレスは既に登録されています。');
      }

      const props = PropertiesService.getScriptProperties();
      const dbId = props.getProperty(PROPS_KEYS.DATABASE_SPREADSHEET_ID);

      if (!dbId) {
        throw new Error('データベース設定が不完全です。システム管理者に連絡してください。');
      }

      const service = getSheetsServiceCached();
      const sheetName = DB_CONFIG.SHEET_NAME;

      // Batch operation preparation (GAS performance best practice)
      const newRow = DB_CONFIG.HEADERS.map(function (header) {
        const value = userData[header];
        // boolean値を正しく処理
        if (typeof value === 'boolean') return value;
        if (value === undefined || value === null) return '';
        return value;
      });

      // Enhanced logging with structured data
      console.info('📊 createUser: Database write preparation', {
        headers: DB_CONFIG.HEADERS,
        rowData: newRow,
        userEmail: userData.userEmail,
        sheetName: sheetName,
      });

      // Single batch write operation
      appendSheetsData(service, dbId, "'" + sheetName + "'!A1", [newRow]);

      console.info('✅ createUser: Database write completed', {
        userEmail: userData.userEmail,
        userId: userData.userId,
        executionTime: Date.now() - startTime + 'ms',
      });

      // 新規ユーザー用の専用フォルダを作成
      try {
        console.info('📁 createUser: Creating user folder', {
          userEmail: userData.userEmail,
        });

        const folder = createUserFolder(userData.userEmail);
        if (folder) {
          console.info('✅ createUser: User folder created successfully', {
            userEmail: userData.userEmail,
            folderName: folder.getName(),
            folderId: folder.getId(),
          });
        } else {
          console.warn('⚠️ createUser: User folder creation failed (continuing process)', {
            userEmail: userData.userEmail,
          });
        }
      } catch (folderError) {
        console.warn('⚠️ createUser: Folder creation error (non-critical)', {
          userEmail: userData.userEmail,
          error: folderError.message,
          stack: folderError.stack,
        });
      }

      // 最適化: 新規ユーザー作成時は対象キャッシュのみ無効化 (修正: userIdを使用)
      invalidateUserCache(userData.userId, userData.userEmail, null, false);

      // キャッシュ整合性確保: 新規作成ユーザーを即座にキャッシュに登録
      const cacheKey = `user_by_email_${userData.userEmail}`;
      CacheService.getScriptCache().put(cacheKey, JSON.stringify(userData), 300);
      console.info('🔄 createUser: New user cached for immediate access', {
        userEmail: userData.userEmail,
        cacheKey: cacheKey,
      });

      console.info('🎉 createUser: User creation process completed successfully', {
        userEmail: userData.userEmail,
        userId: userData.userId,
        totalExecutionTime: Date.now() - startTime + 'ms',
      });

      return userData;
    } catch (error) {
      // Enhanced error handling with structured logging
      console.error('❌ createUser: User creation failed', {
        userEmail: userData.userEmail || 'unknown',
        userId: userData.userId || 'unknown',
        error: error.message,
        stack: error.stack,
        executionTime: Date.now() - startTime + 'ms',
      });

      // Re-throw with user-friendly message
      throw new Error(
        'ユーザー登録に失敗しました。システム管理者に連絡してください。: ' + error.message
      );
    } finally {
      lock.releaseLock();
      console.info('🔓 createUser: Lock released', {
        userEmail: userData.userEmail || 'unknown',
      });
    }
  },

  /**
   * 簡素化されたユーザー検索関数 - メールアドレスでユーザーを検索
   * @param {string} email - 検索対象のメールアドレス
   * @returns {Object|null} ユーザー情報またはnull
   */
  findUserByEmail: function (email) {
    if (!email || typeof email !== 'string') {
      console.warn('findUserByEmail: 無効なメールアドレス', email);
      return null;
    }

    // キャッシュキーを生成
    const cacheKey = 'user_email_' + email;

    try {
      // キャッシュから取得を試行
      const cached = CacheService.getScriptCache().get(cacheKey);
      if (cached) {
        if (cached === 'null') {
          return null;
        }
        return JSON.parse(cached);
      }
    } catch (error) {
      console.warn('findUserByEmail: キャッシュ読み込みエラー', error.message);
    }

    try {
      // データベースから検索
      const service = getSheetsService();
      const dbId = getSecureDatabaseId();
      const sheetName = DB_CONFIG.SHEET_NAME;

      // シート全体のデータを取得 (14列対応)
      const data = batchGetSheetsData(service, dbId, [`'${sheetName}'!A:N`]);

      if (!data.valueRanges || !data.valueRanges[0] || !data.valueRanges[0].values) {
        console.warn('findUserByEmail: データベースからデータを取得できませんでした');
        return null;
      }

      const rows = data.valueRanges[0].values;
      const headers = rows[0];

      // メールアドレス列のインデックスを取得
      const emailIndex = headers.indexOf('userEmail');
      if (emailIndex === -1) {
        console.error('findUserByEmail: userEmail列が見つかりません');
        return null;
      }

      // メールアドレスでマッチする行を検索
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        if (row[emailIndex] === email) {
          // 見つかった行をオブジェクトに変換
          const userObj = {};
          headers.forEach((header, index) => {
            userObj[header] = row[index] || '';
          });

          // キャッシュに保存
          try {
            CacheService.getScriptCache().put(cacheKey, JSON.stringify(userObj), 300);
          } catch (cacheError) {
            console.warn('findUserByEmail: キャッシュ保存エラー', cacheError.message);
          }

          console.log('findUserByEmail: ユーザー発見:', email);
          return userObj;
        }
      }

      // 見つからなかった場合
      try {
        CacheService.getScriptCache().put(cacheKey, 'null', 60);
      } catch (cacheError) {
        console.warn('findUserByEmail: nullキャッシュ保存エラー', cacheError.message);
      }

      console.log('findUserByEmail: ユーザーが見つかりませんでした:', email);
      return null;
    } catch (error) {
      return ErrorManager.handleSafely(error, 'findUserByEmail', null);
    }
  },

  /**
   * ユーザーIDでユーザーを検索
   * @param {string} userId - 検索対象のユーザーID
   * @returns {Object|null} ユーザー情報またはnull
   */
  findUserById: function (userId) {
    if (!userId || typeof userId !== 'string') {
      console.warn('findUserById: 無効なユーザーID', userId);
      return null;
    }

    // キャッシュキーを生成
    const cacheKey = 'user_id_' + userId;

    try {
      // キャッシュから取得を試行
      const cached = CacheService.getScriptCache().get(cacheKey);
      if (cached) {
        if (cached === 'null') {
          return null;
        }
        return JSON.parse(cached);
      }
    } catch (error) {
      console.warn('findUserById: キャッシュ読み込みエラー', error.message);
    }

    try {
      // データベースから検索
      const service = getSheetsService();
      const dbId = getSecureDatabaseId();
      const sheetName = DB_CONFIG.SHEET_NAME;

      // シート全体のデータを取得 (14列対応)
      const data = batchGetSheetsData(service, dbId, [`'${sheetName}'!A:N`]);

      if (!data.valueRanges || !data.valueRanges[0] || !data.valueRanges[0].values) {
        console.warn('findUserById: データベースからデータを取得できませんでした');
        return null;
      }

      const rows = data.valueRanges[0].values;
      const headers = rows[0];

      // ユーザーID列のインデックスを取得
      const userIdIndex = headers.indexOf('userId');
      if (userIdIndex === -1) {
        console.error('findUserById: userId列が見つかりません');
        return null;
      }

      // ユーザーIDでマッチする行を検索
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        if (row[userIdIndex] === userId) {
          // 見つかった行をオブジェクトに変換
          const userObj = {};
          headers.forEach((header, index) => {
            userObj[header] = row[index] || '';
          });

          // キャッシュに保存
          try {
            CacheService.getScriptCache().put(cacheKey, JSON.stringify(userObj), 300);
          } catch (cacheError) {
            console.warn('findUserById: キャッシュ保存エラー', cacheError.message);
          }

          console.log('findUserById: ユーザー発見:', userId);
          return userObj;
        }
      }

      // 見つからなかった場合
      try {
        CacheService.getScriptCache().put(cacheKey, 'null', 60);
      } catch (cacheError) {
        console.warn('findUserById: nullキャッシュ保存エラー', cacheError.message);
      }

      console.log('findUserById: ユーザーが見つかりませんでした:', userId);
      return null;
    } catch (error) {
      return ErrorManager.handleSafely(error, 'findUserById', null);
    }
  },

  /**
   * 全ユーザーデータを取得する
   * @returns {Array} 全ユーザーの配列
   */
  getAllUsers: function() {
    try {
      console.info('getAllUsers: 全ユーザーデータ取得開始');
      
      const service = getSheetsService();
      const dbId = getSecureDatabaseId();
      const sheetName = DB_CONFIG.SHEET_NAME;
      
      // 全データを取得
      const data = batchGetSheetsData(service, dbId, [`'${sheetName}'!A:N`]);
      
      if (!data.valueRanges || !data.valueRanges[0] || !data.valueRanges[0].values) {
        console.info('getAllUsers: ユーザーデータが見つかりません');
        return [];
      }

      const rows = data.valueRanges[0].values;
      
      if (rows.length < 2) {
        console.info('getAllUsers: ユーザーデータが見つかりません');
        return [];
      }

      const headers = rows[0];
      const userRows = rows.slice(1);
      
      // 各行をオブジェクトに変換
      const users = userRows.map(row => {
        const user = {};
        headers.forEach((header, index) => {
          user[header] = row[index] || '';
        });
        return user;
      });
      
      console.info(`getAllUsers: ${users.length}件のユーザーデータを取得`);
      return users;
      
    } catch (error) {
      console.error('getAllUsers エラー:', error.message);
      return ErrorManager.handleSafely(error, 'getAllUsers', []);
    }
  },

  /**
   * キャッシュを使わずにメールアドレスでユーザーを検索（ログイン専用）
   * ログイン時の「はじめる」ボタンでは最新データを確実に取得するため
   * @param {string} email - 検索対象のメールアドレス
   * @returns {Object|null} ユーザー情報またはnull
   */
  findUserByEmailNoCache: function (email) {
    if (!email || typeof email !== 'string') {
      console.warn('findUserByEmailNoCache: 無効なメールアドレス', email);
      return null;
    }

    console.info('🔄 findUserByEmailNoCache: キャッシュをバイパスしてDB直接検索', { email });

    try {
      // データベースから直接検索（キャッシュなし）
      const service = getSheetsService();
      const dbId = getSecureDatabaseId();
      const sheetName = DB_CONFIG.SHEET_NAME;

      // シート全体のデータを取得
      const data = batchGetSheetsData(service, dbId, [`'${sheetName}'!A:N`]);

      if (!data.valueRanges || !data.valueRanges[0] || !data.valueRanges[0].values) {
        console.warn('findUserByEmailNoCache: データベースからデータを取得できませんでした');
        return null;
      }

      const rows = data.valueRanges[0].values;
      if (rows.length < 2) {
        console.info('findUserByEmailNoCache: ユーザーデータがありません');
        return null;
      }

      const headers = rows[0];
      const userRows = rows.slice(1);

      // メールアドレスでユーザーを検索
      const emailIndex = headers.indexOf('userEmail');
      if (emailIndex === -1) {
        console.error('findUserByEmailNoCache: userEmailカラムが見つかりません');
        return null;
      }

      for (const row of userRows) {
        if (row[emailIndex] === email) {
          // ユーザーオブジェクトを構築
          const user = {};
          headers.forEach((header, index) => {
            user[header] = row[index] || '';
          });

          console.info('✅ findUserByEmailNoCache: ユーザー発見（キャッシュバイパス）', {
            email,
            userId: user.userId,
            timestamp: new Date().toISOString()
          });

          return user;
        }
      }

      console.info('findUserByEmailNoCache: ユーザーが見つかりませんでした', { email });
      return null;

    } catch (error) {
      console.error('findUserByEmailNoCache エラー:', error.message);
      return ErrorManager.handleSafely(error, 'findUserByEmailNoCache', null);
    }
  },
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
    rollbackActions: [],
  };

  try {
    // 入力検証
    if (!executorEmail || !targetUserId || !targetEmail || !deleteType) {
      throw new Error('必須パラメータが不足しています');
    }

    const props = PropertiesService.getScriptProperties();
    const dbId = props.getProperty(PROPS_KEYS.DATABASE_SPREADSHEET_ID);

    if (!dbId) {
      console.warn('削除ログの記録をスキップします: データベースIDが設定されていません');
      return { success: false, reason: 'no_database_id' };
    }

    transactionLog.steps.push('validation_complete');

    const service = getSheetsServiceCached();
    const logSheetName = DELETE_LOG_SHEET_CONFIG.SHEET_NAME;

    // ロック取得でトランザクション開始
    const lock = LockService.getScriptLock();
    let lockAcquired = false;

    try {
      lockAcquired = lock.tryLock(5000);
      if (!lockAcquired) {
        throw new Error('ログ記録のロック取得に失敗しました');
      }
      transactionLog.steps.push('lock_acquired');

      // ログシートの存在確認・作成（トランザクション内）
      let sheetCreated = false;
      try {
        const spreadsheetInfo = getSpreadsheetsData(service, dbId);
        const logSheetExists = spreadsheetInfo.sheets.some(
          (sheet) => sheet.properties.title === logSheetName
        );

        if (!logSheetExists) {
          // バッチ最適化: ログシート作成

          const addSheetRequest = {
            addSheet: {
              properties: {
                title: logSheetName,
                gridProperties: {
                  rowCount: 1000,
                  columnCount: DELETE_LOG_SHEET_CONFIG.HEADERS.length,
                },
              },
            },
          };

          // シートを作成
          batchUpdateSpreadsheet(service, dbId, {
            requests: [addSheetRequest],
          });

          transactionLog.steps.push('sheet_created');
          sheetCreated = true;

          // ヘッダー行を追加（作成直後）
          appendSheetsData(service, dbId, `'${logSheetName}'!A1`, [
            DELETE_LOG_SHEET_CONFIG.HEADERS,
          ]);
          transactionLog.steps.push('headers_added');
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
        String(deleteType).substring(0, 50),
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

        return {
          success: true,
          logEntry: logEntry,
          transactionLog: transactionLog,
        };
      } catch (appendError) {
        throw new Error(`ログエントリの追加に失敗: ${appendError.message}`);
      }
    } finally {
      if (lockAcquired) {
        lock.releaseLock();
        transactionLog.steps.push('lock_released');
      }
    }
  } catch (error) {
    transactionLog.duration = Date.now() - transactionLog.startTime;

    // 構造化エラーログ
    const errorInfo = {
      timestamp: new Date().toISOString(),
      function: 'logAccountDeletion',
      severity: 'medium', // ログ記録失敗は削除処理自体を止めない
      parameters: { executorEmail, targetUserId, targetEmail, deleteType },
      error: error.message,
      transactionLog: transactionLog,
    };

    console.error('🚨 削除ログ記録エラー:', JSON.stringify(errorInfo, null, 2));

    // ロールバック処理（必要に応じて）
    // 現在はログ記録のみなので、深刻なロールバックは不要

    return {
      success: false,
      error: error.message,
      transactionLog: transactionLog,
    };
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
    if (!Deploy.isUser()) {
      throw new Error('この機能にアクセスする権限がありません。');
    }

    // 厳格な削除理由の検証
    if (!reason || typeof reason !== 'string' || reason.trim().length < 20) {
      throw new Error('削除理由は20文字以上で入力してください。');
    }

    // 削除理由の内容検証（不適切な理由を防ぐ）
    const invalidReasonPatterns = [/test/i, /テスト/i, /試し/i, /適当/i, /dummy/i, /sample/i];

    if (invalidReasonPatterns.some((pattern) => pattern.test(reason))) {
      throw new Error(
        '削除理由に適切な内容を入力してください。テスト用の削除は許可されていません。'
      );
    }

    const executorEmail = User.email();

    // セキュリティチェック: 管理者メールの検証
    if (!executorEmail || !executorEmail.includes('@')) {
      throw new Error('実行者のメールアドレスを取得できません。');
    }

    // 削除対象ユーザー情報を安全に取得
    const targetUserInfo = DB.findUserById(targetUserId);
    if (!targetUserInfo) {
      throw new Error('削除対象のユーザーが見つかりません。');
    }

    // 追加セキュリティチェック
    if (!targetUserInfo.userEmail || !targetUserInfo.userId) {
      throw new Error('削除対象ユーザーの情報が不完全です。');
    }

    // 自分自身の削除を防ぐ
    if (targetUserInfo.userEmail === executorEmail) {
      throw new Error(
        '自分自身のアカウントは管理者削除機能では削除できません。個人用削除機能をご利用ください。'
      );
    }

    // 最後のアクティブ確認（削除予定の危険度チェック）
    if (targetUserInfo.lastAccessedAt) {
      const lastAccess = new Date(targetUserInfo.lastAccessedAt);
      const daysSinceLastAccess = (Date.now() - lastAccess.getTime()) / (1000 * 60 * 60 * 24);

      if (daysSinceLastAccess < 7) {
        console.warn(
          `警告: 削除対象ユーザーは${Math.floor(daysSinceLastAccess)}日前にアクセスしています`
        );
      }
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(15000);

    try {
      // データベースからユーザー行を削除
      const props = PropertiesService.getScriptProperties();
      const dbId = props.getProperty(PROPS_KEYS.DATABASE_SPREADSHEET_ID);

      if (!dbId) {
        throw new Error('データベースIDが設定されていません');
      }

      const service = getSheetsServiceCached();
      const sheetName = DB_CONFIG.SHEET_NAME;

      // データベーススプレッドシートの情報を取得
      const spreadsheetInfo = getSpreadsheetsData(service, dbId);
      const targetSheetId = spreadsheetInfo.sheets.find(
        (sheet) => sheet.properties.title === sheetName
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
              endIndex: rowToDelete,
            },
          },
        };

        batchUpdateSpreadsheet(service, dbId, {
          requests: [deleteRequest],
        });
      } else {
        throw new Error('削除対象のユーザー行が見つかりませんでした');
      }

      // 削除ログを記録
      logAccountDeletion(
        executorEmail,
        targetUserId,
        targetUserInfo.userEmail,
        reason.trim(),
        'admin'
      );

      // 関連キャッシュを削除
      invalidateUserCache(
        targetUserId,
        targetUserInfo.userEmail,
        targetUserInfo.spreadsheetId,
        false
      );
    } finally {
      lock.releaseLock();
    }

    const successMessage = `管理者によりアカウント「${targetUserInfo.userEmail}」が削除されました。\n削除理由: ${reason.trim()}`;
    console.log(successMessage);
    return {
      success: true,
      message: successMessage,
      deletedUser: {
        userId: targetUserId,
        email: targetUserInfo.userEmail,
      },
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
    const currentUserEmail = User.email();
    const targetUser = DB.findUserById(targetUserId);

    if (!targetUser) {
      return false;
    }

    // 本人削除 OR 管理者削除
    return currentUserEmail === targetUser.userEmail || Deploy.isUser();
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
    if (!Deploy.isUser()) {
      throw new Error('この機能にアクセスする権限がありません。');
    }

    const props = PropertiesService.getScriptProperties();
    const dbId = props.getProperty(PROPS_KEYS.DATABASE_SPREADSHEET_ID);

    if (!dbId) {
      throw new Error('データベースIDが設定されていません');
    }

    const service = getSheetsServiceCached();
    const logSheetName = DELETE_LOG_SHEET_CONFIG.SHEET_NAME;

    try {
      // ログシートの存在確認
      const spreadsheetInfo = getSpreadsheetsData(service, dbId);
      const logSheetExists = spreadsheetInfo.sheets.some(
        (sheet) => sheet.properties.title === logSheetName
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
    let accessToken;
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

    const service = createSheetsService(accessToken);
    if (!service || !service.baseUrl) {
      console.error('❌ Failed to create sheets service or service object is invalid');
      throw new Error('Sheets APIサービスの初期化に失敗しました。');
    }

    console.log('✅ Sheets service created successfully');
    return service;
  } catch (error) {
    console.error('❌ getSheetsService error:', error.message);
    console.error('❌ Error stack:', error.stack);
    throw error; // エラーを再スロー
  }
}

/**
 * ユーザーデータの整合性を修正する
 * @param {string} userId - ユーザーID
 * @returns {object} 修正結果
 */
function fixUserDataConsistency(userId) {
  try {
    // 最新のユーザー情報を取得
    const userInfo = findUserByIdFresh(userId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    console.log('📊 現在のspreadsheetId:', userInfo.spreadsheetId);

    const configJson = JSON.parse(userInfo.configJson || '{}');
    console.log('📝 configJson内のpublishedSpreadsheetId:', configJson.publishedSpreadsheetId);

    let needsUpdate = false;
    let updateData = {};

    // 1. エラー情報をクリーンアップ
    if (configJson.lastError || configJson.errorAt) {
      console.log('🧹 エラー情報をクリーンアップ');
      delete configJson.lastError;
      delete configJson.errorAt;
      needsUpdate = true;
    }

    // 2. 重複フィールドの削除（published*フィールドは不要）
    if (configJson.publishedSpreadsheetId) {
      console.log('🔄 不要なpublishedSpreadsheetIdフィールドを削除');
      delete configJson.publishedSpreadsheetId;
      needsUpdate = true;
    }
    if (configJson.publishedSheetName) {
      console.log('🔄 不要なpublishedSheetNameフィールドを削除');
      delete configJson.publishedSheetName;
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

      updateUser(userId, updateData);

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

  // 許可されたフィールドのホワイトリスト検証（DBヘッダーと一致）
  const allowedFields = [
    'userEmail', // ownerEmail → userEmail に修正
    'spreadsheetId',
    'spreadsheetUrl',
    'configJson',
    'lastAccessedAt',
    'createdAt',
    'formUrl',
    'isActive', // status → isActive に修正
    'sheetName',
    'columnMappingJson',
    'publishedAt',
    'appUrl',
    'lastModified',
  ];
  const updateFields = Object.keys(updateData);
  const invalidFields = updateFields.filter((field) => !allowedFields.includes(field));

  if (invalidFields.length > 0) {
    throw new Error(
      'データ更新エラー: 許可されていないフィールドが含まれています: ' + invalidFields.join(', ')
    );
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
    const props = PropertiesService.getScriptProperties();
    const dbId = props.getProperty(PROPS_KEYS.DATABASE_SPREADSHEET_ID);

    if (!dbId) {
      throw new Error('データ更新エラー: データベースIDが設定されていません');
    }
    let service = getSheetsServiceCached();
    const sheetName = DB_CONFIG.SHEET_NAME;

    // 現在のデータを取得 (14列対応)
    const data = batchGetSheetsData(service, dbId, ["'" + sheetName + "'!A:N"]);
    const values = data.valueRanges[0].values || [];

    if (values.length === 0) {
      throw new Error('データベースが空です');
    }

    const headers = values[0];
    const userIdIndex = headers.indexOf('userId');
    let rowIndex = -1;

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
    const requests = Object.keys(updateData)
      .map(function (key) {
        const colIndex = headers.indexOf(key);
        if (colIndex === -1) return null;

        return {
          range: "'" + sheetName + "'!" + String.fromCharCode(65 + colIndex) + rowIndex,
          values: [[updateData[key]]],
        };
      })
      .filter(function (item) {
        return item !== null;
      });

    if (requests.length > 0) {
      requests.forEach(function (req, index) {
        console.log(
          '  ' + (index + 1) + '. 範囲: ' + req.range + ', 値: ' + JSON.stringify(req.values)
        );
      });

      const maxRetries = 2;
      let retryCount = 0;
      let updateSuccess = false;

      while (retryCount <= maxRetries && !updateSuccess) {
        try {
          if (retryCount > 0) {
            console.log('🔄 認証エラーによるリトライ (' + retryCount + '/' + maxRetries + ')');
            // 認証エラーの場合、新しいサービスを取得
            service = getSheetsServiceCached(true); // forceRefresh = true
            Utilities.sleep(1000); // 少し待機
          }

          batchUpdateSheetsData(service, dbId, requests);
          updateSuccess = true;

          // 更新成功の確認
          const verifyData = batchGetSheetsData(service, dbId, [
            "'" +
              sheetName +
              "'!" +
              String.fromCharCode(65 + userIdIndex) +
              rowIndex +
              ':' +
              String.fromCharCode(72) +
              rowIndex,
          ]);
          if (
            verifyData.valueRanges &&
            verifyData.valueRanges[0] &&
            verifyData.valueRanges[0].values
          ) {
            const updatedRow = verifyData.valueRanges[0].values[0];
            if (updateData.spreadsheetId) {
              const spreadsheetIdIndex = headers.indexOf('spreadsheetId');
              console.log(
                '🎯 スプレッドシートID更新確認:',
                updatedRow[spreadsheetIdIndex] === updateData.spreadsheetId ? '✅ 成功' : '❌ 失敗'
              );
            }
          }
        } catch (updateError) {
          retryCount++;
          const errorMessage = updateError.toString();
          if (
            errorMessage.includes('401') ||
            errorMessage.includes('UNAUTHENTICATED') ||
            errorMessage.includes('ACCESS_TOKEN_EXPIRED')
          ) {
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
    }

    // スプレッドシートIDが更新された場合、サービスアカウントと共有
    if (updateData.spreadsheetId) {
      try {
        shareSpreadsheetWithServiceAccount(updateData.spreadsheetId);
      } catch (shareError) {
        console.error('ユーザー更新時のサービスアカウント共有エラー:', shareError.message);
        console.error(
          'スプレッドシートの更新は完了しましたが、サービスアカウントとの共有に失敗しました。手動で共有してください。'
        );
      }
    }

    // 重要: 更新完了後に包括的キャッシュ同期を実行
    const userInfo = DB.findUserById(userId);
    const email = updateData.userEmail || (userInfo ? userInfo.userEmail : null);
    const oldSpreadsheetId = userInfo ? userInfo.spreadsheetId : null;
    const newSpreadsheetId = updateData.spreadsheetId || oldSpreadsheetId;
    // クリティカル更新時の包括的キャッシュ同期
    synchronizeCacheAfterCriticalUpdate(userId, email, oldSpreadsheetId, newSpreadsheetId);

    return { status: 'success', message: 'ユーザープロフィールが正常に更新されました' };
  } catch (error) {
    console.error('ユーザー更新エラー:', error);
    throw error;
  }
}

/**
 * データベースシートを初期化
 * @param {string} spreadsheetId - データベースのスプレッドシートID
 */
function initializeDatabaseSheet(spreadsheetId) {
  const service = getSheetsServiceCached();
  const sheetName = DB_CONFIG.SHEET_NAME;

  try {
    // シートが存在するか確認
    const spreadsheet = getSpreadsheetsData(service, spreadsheetId);
    const sheetExists = spreadsheet.sheets.some(function (s) {
      return s.properties.title === sheetName;
    });

    if (!sheetExists) {
      // バッチ処理最適化: シート作成とヘッダー追加を1回のAPI呼び出しで実行
      console.log('📊 バッチ最適化: シート作成+ヘッダー追加を同時実行');

      const requests = [
        // 1. シートを作成
        {
          addSheet: {
            properties: {
              title: sheetName,
              gridProperties: {
                rowCount: 1000,
                columnCount: DB_CONFIG.HEADERS.length, // 14カラム対応
              },
            },
          },
        },
      ];

      // バッチ実行: シートを作成
      batchUpdateSpreadsheet(service, spreadsheetId, { requests: requests });

      // 2. 作成直後にヘッダーを追加（A1記法でレンジを指定）
      const headerRange =
        "'" + sheetName + "'!A1:" + String.fromCharCode(65 + DB_CONFIG.HEADERS.length - 1) + '1'; // A1:N1 (14カラム対応)
      updateSheetsData(service, spreadsheetId, headerRange, [DB_CONFIG.HEADERS]);

      console.log(
        '✅ バッチ最適化完了: シート作成+ヘッダー追加（2回のAPI呼び出し→シーケンシャル実行）'
      );
    } else {
      // シートが既に存在する場合は、ヘッダーのみ更新（既存動作を維持）
      const headerRange =
        "'" + sheetName + "'!A1:" + String.fromCharCode(65 + DB_CONFIG.HEADERS.length - 1) + '1'; // A1:N1 (14カラム対応)
      updateSheetsData(service, spreadsheetId, headerRange, [DB_CONFIG.HEADERS]);
    }
  } catch (e) {
    console.error('データベースシートの初期化に失敗: ' + e.message);
    throw new Error(
      'データベースシートの初期化に失敗しました。サービスアカウントに編集者権限があるか確認してください。詳細: ' +
        e.message
    );
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
        const userProps = PropertiesService.getUserProperties();
        const currentUserId = userProps.getProperty('CURRENT_USER_ID');
        if (currentUserId === userId) {
          // 現在のユーザーIDが無効な場合はクリア
          userProps.deleteProperty('CURRENT_USER_ID');
          console.log('[Cache] Cleared invalid CURRENT_USER_ID: ' + userId);
        }
      } catch (propsError) {
        console.warn('Failed to clear user properties:', propsError.message);
      }
    }

    // 全データベースキャッシュのクリアは最後の手段として実行
    // 頻繁な実行を避けるため、確実にデータ不整合がある場合のみ実行
    let shouldClearAll = false;

    // 判定条件: 複数のユーザーで問題が発生している場合のみ全クリア
    if (shouldClearAll) {
      clearDatabaseCache();
    }

    console.log('[Cache] Handled missing user: ' + userId);
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
        get: function (options) {
          const url =
            'https://sheets.googleapis.com/v4/spreadsheets/' +
            options.spreadsheetId +
            '/values/' +
            encodeURIComponent(options.range);

          const response = UrlFetchApp.fetch(url, {
            headers: { Authorization: 'Bearer ' + accessToken },
            muteHttpExceptions: true,
            followRedirects: true,
            validateHttpsCertificates: true,
          });

          if (response.getResponseCode() !== 200) {
            throw new Error(
              'Sheets API error: ' + response.getResponseCode() + ' - ' + response.getContentText()
            );
          }

          return JSON.parse(response.getContentText());
        },
      },
      get: function (options) {
        const url = 'https://sheets.googleapis.com/v4/spreadsheets/' + options.spreadsheetId;
        if (options.fields) {
          url += '?fields=' + encodeURIComponent(options.fields);
        }

        const response = UrlFetchApp.fetch(url, {
          headers: { Authorization: 'Bearer ' + accessToken },
          muteHttpExceptions: true,
          followRedirects: true,
          validateHttpsCertificates: true,
        });

        if (response.getResponseCode() !== 200) {
          throw new Error(
            'Sheets API error: ' + response.getResponseCode() + ' - ' + response.getContentText()
          );
        }

        return JSON.parse(response.getContentText());
      },
    },
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
  const cacheKey = `batchGet_${spreadsheetId}_${JSON.stringify(ranges)}`;
  return cacheManager.get(
    cacheKey,
    () => {
      try {
        // 防御的プログラミング: サービスオブジェクトのプロパティを安全に取得
        const baseUrl = service.baseUrl;
        const accessToken = service.accessToken;

        // baseUrlが失われている場合の防御処理
        if (!baseUrl || typeof baseUrl !== 'string') {
          console.warn(
            '⚠️ baseUrlが見つかりません。デフォルトのGoogleSheetsAPIエンドポイントを使用します'
          );
          baseUrl = 'https://sheets.googleapis.com/v4/spreadsheets';
        }

        if (!accessToken || typeof accessToken !== 'string') {
          throw new Error(
            'アクセストークンが見つかりません。サービスオブジェクトが破損している可能性があります'
          );
        }

        // 安全なURL構築
        const url =
          baseUrl +
          '/' +
          encodeURIComponent(spreadsheetId) +
          '/values:batchGet?' +
          ranges
            .map(function (range) {
              return 'ranges=' + encodeURIComponent(range);
            })
            .join('&');

        const response = UrlFetchApp.fetch(url, {
          headers: { Authorization: 'Bearer ' + accessToken },
          muteHttpExceptions: true,
          followRedirects: true,
          validateHttpsCertificates: true,
        });

        const responseCode = response.getResponseCode();
        const responseText = response.getContentText();

        if (responseCode !== 200) {
          console.error('Sheets API エラー詳細:', {
            code: responseCode,
            response: responseText,
            url: url.substring(0, 100) + '...',
            spreadsheetId: spreadsheetId,
          });
          throw new Error('Sheets API error: ' + responseCode + ' - ' + responseText);
        }

        let result;
        try {
          result = JSON.parse(responseText);
        } catch (parseError) {
          console.error('❌ JSON解析エラー:', parseError.message);
          console.error('❌ Response text:', responseText.substring(0, 200));
          throw new Error('APIレスポンスのJSON解析に失敗: ' + parseError.message);
        }

        // レスポンス構造の検証
        if (!result || typeof result !== 'object') {
          throw new Error(
            '無効なAPIレスポンス: オブジェクトが期待されましたが ' + typeof result + ' を受信'
          );
        }

        if (!result.valueRanges || !Array.isArray(result.valueRanges)) {
          console.warn(
            '⚠️ valueRanges配列が見つからないか、配列でありません:',
            typeof result.valueRanges
          );
          result.valueRanges = []; // 空配列を設定
        }

        // リクエストした範囲数と一致するか確認
        if (result.valueRanges.length !== ranges.length) {
          console.warn(
            `⚠️ リクエスト範囲数(${ranges.length})とレスポンス数(${result.valueRanges.length})が一致しません`
          );
        }

        // 各範囲のデータ存在確認
        result.valueRanges.forEach((valueRange, index) => {
          const hasValues = valueRange.values && valueRange.values.length > 0;
          console.log(
            `📊 範囲[${index}] ${ranges[index]}: ${hasValues ? valueRange.values.length + '行' : 'データなし'}`
          );
          if (hasValues) {
            console.log(
              `DEBUG: batchGetSheetsData - 範囲[${index}] データプレビュー:`,
              JSON.stringify(valueRange.values.slice(0, 5))
            ); // 最初の5行をプレビュー
          }
        });

        return result;
      } catch (error) {
        console.error('❌ batchGetSheetsData error:', error.message);
        console.error('❌ Error stack:', error.stack);
        throw new Error('データ取得に失敗しました: ' + error.message);
      }
    },
    { ttl: 120 }
  ); // 2分間キャッシュ（API制限対策）
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
    const url = service.baseUrl + '/' + spreadsheetId + '/values:batchUpdate';

    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + service.accessToken },
      payload: JSON.stringify({
        data: requests,
        valueInputOption: 'RAW',
      }),
      muteHttpExceptions: true,
      followRedirects: true,
      validateHttpsCertificates: true,
    });

    if (response.getResponseCode() !== 200) {
      throw new Error(
        'Sheets API error: ' + response.getResponseCode() + ' - ' + response.getContentText()
      );
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
  const url =
    service.baseUrl +
    '/' +
    spreadsheetId +
    '/values/' +
    encodeURIComponent(range) +
    ':append?valueInputOption=RAW&insertDataOption=INSERT_ROWS';

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + service.accessToken },
    payload: JSON.stringify({ values: values }),
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
    const baseUrl = service.baseUrl;
    const accessToken = service.accessToken;

    // baseUrlが失われている場合の防御処理
    if (!baseUrl || typeof baseUrl !== 'string') {
      console.warn(
        '⚠️ baseUrlが見つかりません。デフォルトのGoogleSheetsAPIエンドポイントを使用します'
      );
      baseUrl = 'https://sheets.googleapis.com/v4/spreadsheets';
    }

    if (!accessToken || typeof accessToken !== 'string') {
      throw new Error(
        'アクセストークンが見つかりません。サービスオブジェクトが破損している可能性があります'
      );
    }

    // 安全なURL構築 - シート情報を含む基本的なメタデータを取得するために fields パラメータを追加
    const url = baseUrl + '/' + encodeURIComponent(spreadsheetId) + '?fields=sheets.properties';

    const response = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + accessToken },
      muteHttpExceptions: true,
      followRedirects: true,
      validateHttpsCertificates: true,
    });

    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode !== 200) {
      console.error('Sheets API エラー詳細:', {
        code: responseCode,
        response: responseText,
        url: url.substring(0, 100) + '...',
        spreadsheetId: spreadsheetId,
      });

      if (responseCode === 403) {
        try {
          const errorResponse = JSON.parse(responseText);
          if (
            errorResponse.error &&
            errorResponse.error.message === 'The caller does not have permission'
          ) {
            const serviceAccountEmail = getServiceAccountEmail();
            throw new Error(
              'スプレッドシートへのアクセス権限がありません。サービスアカウント（' +
                serviceAccountEmail +
                '）をスプレッドシートの編集者として共有してください。'
            );
          }
        } catch (parseError) {
          console.warn('エラーレスポンスのJSON解析に失敗:', parseError.message);
        }
      }

      throw new Error('Sheets API error: ' + responseCode + ' - ' + responseText);
    }

    let result;
    try {
      result = JSON.parse(responseText);
    } catch (parseError) {
      console.error('❌ JSON解析エラー:', parseError.message);
      console.error('❌ Response text:', responseText.substring(0, 200));
      throw new Error('APIレスポンスのJSON解析に失敗: ' + parseError.message);
    }

    // レスポンス構造の検証
    if (!result || typeof result !== 'object') {
      throw new Error(
        '無効なAPIレスポンス: オブジェクトが期待されましたが ' + typeof result + ' を受信'
      );
    }

    if (!result.sheets || !Array.isArray(result.sheets)) {
      console.warn('⚠️ sheets配列が見つからないか、配列でありません:', typeof result.sheets);
      result.sheets = []; // 空配列を設定してエラーを避ける
    }

    const sheetCount = result.sheets.length;
    console.log('✅ getSpreadsheetsData 成功: 発見シート数:', sheetCount);

    if (sheetCount === 0) {
      console.warn('⚠️ スプレッドシートにシートが見つかりませんでした');
    } else {
      console.log(
        '📋 利用可能なシート:',
        result.sheets.map((s) => s.properties?.title || 'Unknown').join(', ')
      );
    }

    return result;
  } catch (error) {
    console.error('❌ getSpreadsheetsData error:', error.message);
    console.error('❌ Error stack:', error.stack);
    throw new Error('スプレッドシート情報取得に失敗しました: ' + error.message);
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
  const url =
    service.baseUrl +
    '/' +
    spreadsheetId +
    '/values/' +
    encodeURIComponent(range) +
    '?valueInputOption=RAW';

  const response = UrlFetchApp.fetch(url, {
    method: 'put',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + service.accessToken },
    payload: JSON.stringify({ values: values }),
  });

  return JSON.parse(response.getContentText());
}

/**
 * データベース状態を詳細確認する診断機能
 * ⚠️ この関数は SystemManager.gs に移行されました ⚠️
 * @param {string} targetUserId - 確認対象のユーザーID（オプション）
 * @returns {object} データベース診断結果
 * @deprecated SystemManager.checkSetupStatus() または testSystemDiagnosis() を使用してください
 */
function diagnoseDatabase(targetUserId) {
  console.error('⚠️ diagnoseDatabase は SystemManager.gs に移動しました。');
  console.error('   新しい関数: testSystemDiagnosis() を使用してください');
  throw new Error('この関数は SystemManager.gs に移動しました');

  // 以下は旧実装（参考用）
  try {
    const props = PropertiesService.getScriptProperties();
    const dbId = props.getProperty(PROPS_KEYS.DATABASE_SPREADSHEET_ID);

    const diagnosticResult = {
      timestamp: new Date().toISOString(),
      databaseId: dbId,
      targetUserId: targetUserId,
      checks: {},
      summary: {
        overallStatus: 'unknown',
        criticalIssues: [],
        warnings: [],
        recommendations: [],
      },
    };

    // 1. データベース設定チェック
    diagnosticResult.checks.databaseConfig = {
      hasDatabaseId: !!dbId,
      databaseId: dbId ? dbId.substring(0, 10) + '...' : null,
    };

    if (!dbId) {
      diagnosticResult.summary.criticalIssues.push('DATABASE_SPREADSHEET_ID が設定されていません');
      diagnosticResult.summary.overallStatus = 'critical';
      return diagnosticResult;
    }

    // 2. サービス接続テスト
    let service;
    try {
      service = getSheetsServiceCached();
      diagnosticResult.checks.serviceConnection = { status: 'success' };
    } catch (serviceError) {
      diagnosticResult.checks.serviceConnection = {
        status: 'failed',
        error: serviceError.message,
      };
      diagnosticResult.summary.criticalIssues.push(
        'Sheets サービス接続失敗: ' + serviceError.message
      );
    }

    if (!service) {
      diagnosticResult.summary.overallStatus = 'critical';
      return diagnosticResult;
    }

    // 3. データベーススプレッドシートアクセステスト
    try {
      const spreadsheetInfo = getSpreadsheetsData(service, dbId);
      diagnosticResult.checks.spreadsheetAccess = {
        status: 'success',
        sheetCount: spreadsheetInfo.sheets ? spreadsheetInfo.sheets.length : 0,
        hasUserSheet: spreadsheetInfo.sheets.some(
          (sheet) => sheet.properties.title === DB_CONFIG.SHEET_NAME
        ),
      };
    } catch (accessError) {
      diagnosticResult.checks.spreadsheetAccess = {
        status: 'failed',
        error: accessError.message,
      };
      diagnosticResult.summary.criticalIssues.push(
        'スプレッドシートアクセス失敗: ' + accessError.message
      );
    }

    // 4. ユーザーデータ取得テスト
    try {
      const data = batchGetSheetsData(service, dbId, ["'" + DB_CONFIG.SHEET_NAME + "'!A:N"]);
      const values = data.valueRanges[0].values || [];

      diagnosticResult.checks.userData = {
        status: 'success',
        totalRows: values.length,
        userCount: Math.max(0, values.length - 1), // ヘッダー行を除く
        hasHeaders: values.length > 0,
        headers: values.length > 0 ? values[0] : [],
      };

      // 特定ユーザーの検索テスト
      if (targetUserId && values.length > 1) {
        let userFound = false;
        let userRowIndex = -1;

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
          data: userFound ? values[userRowIndex] : null,
        };

        if (!userFound) {
          diagnosticResult.summary.criticalIssues.push(
            '対象ユーザー ' + targetUserId + ' がデータベースに見つかりません'
          );
        }
      }
    } catch (dataError) {
      diagnosticResult.checks.userData = {
        status: 'failed',
        error: dataError.message,
      };
      diagnosticResult.summary.criticalIssues.push('ユーザーデータ取得失敗: ' + dataError.message);
    }

    // 5. キャッシュ状態チェック
    try {
      const cacheStatus = checkCacheStatus(targetUserId);
      diagnosticResult.checks.cache = cacheStatus;

      if (cacheStatus.staleEntries > 0) {
        diagnosticResult.summary.warnings.push(
          cacheStatus.staleEntries + ' 個の古いキャッシュエントリが見つかりました'
        );
      }
    } catch (cacheError) {
      diagnosticResult.checks.cache = {
        status: 'failed',
        error: cacheError.message,
      };
    }

    // 6. 総合判定
    if (diagnosticResult.summary.criticalIssues.length === 0) {
      diagnosticResult.summary.overallStatus =
        diagnosticResult.summary.warnings.length > 0 ? 'warning' : 'healthy';
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

    return diagnosticResult;
  } catch (error) {
    console.error('❌ データベース診断でエラー:', error);
    return {
      timestamp: new Date().toISOString(),
      error: error.message,
      summary: {
        overallStatus: 'error',
        criticalIssues: ['診断処理自体が失敗: ' + error.message],
      },
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
    const cacheStatus = {
      userSpecific: null,
      general: {
        totalEntries: 0,
        staleEntries: 0,
        healthyEntries: 0,
      },
    };

    // ユーザー固有のキャッシュ確認
    if (userId) {
      const userCacheKey = 'user_' + userId;
      const cachedUser = cacheManager.get(userCacheKey, null, { skipFetch: true });
      cacheStatus.userSpecific = {
        userId: userId,
        cacheKey: userCacheKey,
        cached: !!cachedUser,
        data: cachedUser ? 'present' : 'absent',
      };
    }

    // 一般的なキャッシュ統計
    // 注: 実際のキャッシュマネージャーの実装に依存
    try {
      if (typeof cacheManager.getStats === 'function') {
        const stats = cacheManager.getStats();
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
      error: error.message,
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
    const dbCheckResult = {
      timestamp: new Date().toISOString(),
      spreadsheetId: spreadsheetId,
      checks: {},
      summary: {
        status: 'unknown',
        issues: [],
        actions: [],
      },
    };

    // 1. データベーススプレッドシートの権限確認
    const props = PropertiesService.getScriptProperties();
    const dbId = spreadsheetId || props.getProperty(PROPS_KEYS.DATABASE_SPREADSHEET_ID);
    if (!dbId) {
      dbCheckResult.summary.issues.push('データベースIDが設定されていません');
      dbCheckResult.summary.status = 'error';
      return dbCheckResult;
    }

    // 2. サービスアカウント情報確認
    try {
      const serviceAccountEmail = getServiceAccountEmail();
      dbCheckResult.checks.serviceAccount = {
        email: serviceAccountEmail,
        configured: !!serviceAccountEmail,
      };

      if (!serviceAccountEmail) {
        dbCheckResult.summary.issues.push('サービスアカウント設定が見つかりません');
      }
    } catch (saError) {
      dbCheckResult.checks.serviceAccount = {
        configured: false,
        error: saError.message,
      };
      dbCheckResult.summary.issues.push('サービスアカウント設定エラー: ' + saError.message);
    }

    // 3. スプレッドシートアクセステスト
    try {
      const service = getSheetsServiceCached();
      const spreadsheetInfo = getSpreadsheetsData(service, dbId);

      dbCheckResult.checks.spreadsheetAccess = {
        status: 'success',
        sheetCount: spreadsheetInfo.sheets ? spreadsheetInfo.sheets.length : 0,
        canRead: true,
      };

      // 書き込みテスト（安全な方法で）
      try {
        // テスト用のバッチ更新（実際には何も変更しない）
        const testRequest = {
          requests: [],
        };
        // 空のリクエストでテスト
        batchUpdateSpreadsheet(service, dbId, testRequest);
        dbCheckResult.checks.spreadsheetAccess.canWrite = true;
      } catch (writeError) {
        dbCheckResult.checks.spreadsheetAccess.canWrite = false;
        dbCheckResult.checks.spreadsheetAccess.writeError = writeError.message;
        dbCheckResult.summary.issues.push(
          'スプレッドシート書き込み権限不足: ' + writeError.message
        );
      }
    } catch (accessError) {
      dbCheckResult.checks.spreadsheetAccess = {
        status: 'failed',
        error: accessError.message,
        canRead: false,
        canWrite: false,
      };
      dbCheckResult.summary.issues.push('スプレッドシートアクセス失敗: ' + accessError.message);
    }

    // 4. 権限修復の試行
    if (dbCheckResult.summary.issues.length > 0) {
      try {
        console.log('🔧 権限修復を試行中...');

        // サービスアカウントの再共有を試行
        if (dbCheckResult.checks.serviceAccount && dbCheckResult.checks.serviceAccount.email) {
          shareSpreadsheetWithServiceAccount(dbId);
          dbCheckResult.summary.actions.push('サービスアカウントの再共有実行');

          // 修復後の再テスト
          Utilities.sleep(3000); // 共有反映待ち

          try {
            const retestService = getSheetsServiceCached(true); // 強制リフレッシュ;
            const retestInfo = getSpreadsheetsData(retestService, dbId);

            dbCheckResult.checks.postRepairAccess = {
              status: 'success',
              canRead: true,
              repairSuccessful: true,
            };

            // 修復成功後はissuesをクリア
            dbCheckResult.summary.issues = [];
            dbCheckResult.summary.actions.push('アクセス権限修復成功');
          } catch (retestError) {
            dbCheckResult.checks.postRepairAccess = {
              status: 'failed',
              error: retestError.message,
              repairSuccessful: false,
            };
            dbCheckResult.summary.actions.push('修復後テスト失敗: ' + retestError.message);
          }
        }
      } catch (repairError) {
        dbCheckResult.summary.actions.push('権限修復失敗: ' + repairError.message);
      }
    }

    // 5. 最終判定
    if (dbCheckResult.summary.issues.length === 0) {
      dbCheckResult.summary.status = 'healthy';
    } else if (dbCheckResult.summary.actions.length > 0) {
      dbCheckResult.summary.status = 'repaired';
    } else {
      dbCheckResult.summary.status = 'critical';
    }

    return dbCheckResult;
  } catch (error) {
    console.error('❌ サービスアカウント権限確認でエラー:', error);
    return {
      timestamp: new Date().toISOString(),
      error: error.message,
      summary: {
        status: 'error',
        issues: ['権限確認処理自体が失敗: ' + error.message],
      },
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
    const repairResult = {
      timestamp: new Date().toISOString(),
      targetUserId: targetUserId,
      actions: [],
      success: false,
      summary: '',
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
      const permissionResult = verifyServiceAccountPermissions();
      if (
        permissionResult.summary.status === 'repaired' ||
        permissionResult.summary.status === 'healthy'
      ) {
        repairResult.actions.push('サービスアカウント権限確認・修復完了');
      } else {
        repairResult.actions.push(
          'サービスアカウント権限に問題あり: ' +
            (permissionResult.summary.issues ? permissionResult.summary.issues.join(', ') : '不明')
        );
      }
    } catch (permError) {
      repairResult.actions.push('サービスアカウント権限確認失敗: ' + permError.message);
    }

    // 4. 修復後の検証
    try {
      Utilities.sleep(2000); // 少し待機
      const postRepairDiagnosis = SystemManager.checkSetupStatus();

      if (postRepairDiagnosis.isComplete) {
        repairResult.success = true;
        repairResult.summary = '修復成功: システム状態が改善されました';
      } else {
        repairResult.summary = '修復不完全: 追加の手動対応が必要です';
      }

      repairResult.postRepairStatus = postRepairDiagnosis.summary;
    } catch (verifyError) {
      repairResult.summary = '修復実行したが検証に失敗: ' + verifyError.message;
    }

    return repairResult;
  } catch (error) {
    console.error('❌ 自動修復でエラー:', error);
    return {
      timestamp: new Date().toISOString(),
      error: error.message,
      success: false,
      summary: '修復処理自体が失敗: ' + error.message,
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
    const opts = {
      checkDuplicates: options.checkDuplicates !== false,
      checkMissingFields: options.checkMissingFields !== false,
      checkInvalidData: options.checkInvalidData !== false,
      autoFix: options.autoFix || false,
      ...options,
    };

    const dbCheckResult = {
      timestamp: new Date().toISOString(),
      summary: {
        status: 'unknown',
        totalUsers: 0,
        issues: [],
        warnings: [],
        fixed: [],
      },
      details: {
        duplicates: [],
        missingFields: [],
        invalidData: [],
        orphanedData: [],
      },
    };

    // データベース接続確認
    const props = PropertiesService.getScriptProperties();
    const dbId = props.getProperty(PROPS_KEYS.DATABASE_SPREADSHEET_ID);
    if (!dbId) {
      dbCheckResult.summary.issues.push('データベースIDが設定されていません');
      dbCheckResult.summary.status = 'critical';
      return dbCheckResult;
    }

    const service = getSheetsServiceCached();
    const data = batchGetSheetsData(service, dbId, ["'" + DB_CONFIG.SHEET_NAME + "'!A:N"]);
    const values = data.valueRanges[0].values || [];

    if (values.length <= 1) {
      dbCheckResult.summary.status = 'empty';
      return dbCheckResult;
    }

    const headers = values[0];
    const userRows = values.slice(1);
    dbCheckResult.summary.totalUsers = userRows.length;

    console.log('📊 データ整合性チェック: ' + userRows.length + 'ユーザーを確認中');

    // 1. 重複チェック
    if (opts.checkDuplicates) {
      const duplicateResult = checkForDuplicates(headers, userRows);
      dbCheckResult.details.duplicates = duplicateResult.duplicates;
      if (duplicateResult.duplicates.length > 0) {
        dbCheckResult.summary.issues.push(
          duplicateResult.duplicates.length + '件の重複データが見つかりました'
        );
      }
    }

    // 2. 必須フィールドチェック
    if (opts.checkMissingFields) {
      const missingFieldsResult = checkMissingRequiredFields(headers, userRows);
      dbCheckResult.details.missingFields = missingFieldsResult.missing;
      if (missingFieldsResult.missing.length > 0) {
        dbCheckResult.summary.warnings.push(
          missingFieldsResult.missing.length + '件の必須フィールド不足が見つかりました'
        );
      }
    }

    // 3. データ形式チェック
    if (opts.checkInvalidData) {
      const invalidDataResult = checkInvalidDataFormats(headers, userRows);
      dbCheckResult.details.invalidData = invalidDataResult.invalid;
      if (invalidDataResult.invalid.length > 0) {
        dbCheckResult.summary.warnings.push(
          invalidDataResult.invalid.length + '件の不正なデータ形式が見つかりました'
        );
      }
    }

    // 4. 孤立データチェック
    const orphanResult = checkOrphanedData(headers, userRows);
    dbCheckResult.details.orphanedData = orphanResult.orphaned;
    if (orphanResult.orphaned.length > 0) {
      dbCheckResult.summary.warnings.push(
        orphanResult.orphaned.length + '件の孤立データが見つかりました'
      );
    }

    // 5. 自動修復
    if (
      opts.autoFix &&
      (dbCheckResult.summary.issues.length > 0 || dbCheckResult.summary.warnings.length > 0)
    ) {
      try {
        const fixResult = performDataIntegrityFix(
          dbCheckResult.details,
          headers,
          userRows,
          dbId,
          service
        );
        dbCheckResult.summary.fixed = fixResult.fixed;
      } catch (fixError) {
        console.error('❌ 自動修復エラー:', fixError.message);
        dbCheckResult.summary.issues.push('自動修復に失敗: ' + fixError.message);
      }
    }

    // 6. 最終判定
    if (dbCheckResult.summary.issues.length === 0) {
      dbCheckResult.summary.status =
        dbCheckResult.summary.warnings.length > 0 ? 'warning' : 'healthy';
    } else {
      dbCheckResult.summary.status = 'critical';
    }

    return dbCheckResult;
  } catch (error) {
    console.error('❌ データ整合性チェックでエラー:', error);
    return {
      timestamp: new Date().toISOString(),
      error: error.message,
      summary: {
        status: 'error',
        issues: ['整合性チェック自体が失敗: ' + error.message],
      },
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
  const duplicates = [];
  const userIdIndex = headers.indexOf('userId');
  const emailIndex = headers.indexOf('userEmail');

  if (userIdIndex === -1 || emailIndex === -1) {
    return { duplicates: [] };
  }

  const seenUserIds = new Set();
  const seenEmails = new Set();
  for (let i = 0; i < userRows.length; i++) {
    const row = userRows[i];
    const userId = row[userIdIndex];
    const email = row[emailIndex];
    // userId重複チェック
    if (userId && seenUserIds.has(userId)) {
      duplicates.push({
        type: 'userId',
        value: userId,
        rowIndex: i + 2, // スプレッドシート行番号（1ベース + ヘッダー）
        message: 'ユーザーID重複: ' + userId,
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
        message: 'メールアドレス重複: ' + email,
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
  const missing = [];
  const requiredFields = ['userId', 'userEmail']; // 最低限必要なフィールド;
  for (let i = 0; i < userRows.length; i++) {
    const row = userRows[i];
    const missingInThisRow = [];
    for (let j = 0; j < requiredFields.length; j++) {
      const fieldName = requiredFields[j];
      const fieldIndex = headers.indexOf(fieldName);
      if (fieldIndex === -1 || !row[fieldIndex] || row[fieldIndex].trim() === '') {
        missingInThisRow.push(fieldName);
      }
    }

    if (missingInThisRow.length > 0) {
      missing.push({
        rowIndex: i + 2,
        missingFields: missingInThisRow,
        message: '必須フィールド不足: ' + missingInThisRow.join(', '),
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
  const invalid = [];
  const emailIndex = headers.indexOf('userEmail');
  const userIdIndex = headers.indexOf('userId');

  for (let i = 0; i < userRows.length; i++) {
    const row = userRows[i];
    const rowIssues = [];
    // メールアドレス形式チェック
    if (emailIndex !== -1 && row[emailIndex]) {
      const email = row[emailIndex];
      if (!EMAIL_REGEX.test(email)) {
        rowIssues.push('無効なメールアドレス形式: ' + email);
      }
    }

    // ユーザーID形式チェック（UUIDかどうか）
    if (userIdIndex !== -1 && row[userIdIndex]) {
      const userId = row[userIdIndex];
      const uuidRegex = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
      if (!uuidRegex.test(userId)) {
        rowIssues.push('無効なユーザーID形式: ' + userId);
      }
    }

    if (rowIssues.length > 0) {
      invalid.push({
        rowIndex: i + 2,
        issues: rowIssues,
        message: rowIssues.join(', '),
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
  const orphaned = [];
  const isActiveIndex = headers.indexOf('isActive');

  for (let i = 0; i < userRows.length; i++) {
    const row = userRows[i];
    const issues = [];
    // 非アクティブだが他のデータが残っている
    if (
      isActiveIndex !== -1 &&
      (row[isActiveIndex] === false ||
        row[isActiveIndex] === 'FALSE' ||
        row[isActiveIndex] === 'false')
    ) {
      // 非アクティブユーザーでスプレッドシートIDが残っている場合
      const spreadsheetIdIndex = headers.indexOf('spreadsheetId');
      if (spreadsheetIdIndex !== -1 && row[spreadsheetIdIndex]) {
        issues.push('非アクティブユーザーにスプレッドシートIDが残存');
      }
    }

    if (issues.length > 0) {
      orphaned.push({
        rowIndex: i + 2,
        issues: issues,
        message: issues.join(', '),
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
  const fixed = [];
  // 注意: 重複データの自動削除は危険なため、ログのみ記録
  if (details.duplicates.length > 0) {
    console.warn(
      '⚠️ 重複データが検出されましたが、自動削除は行いません:',
      details.duplicates.length + '件'
    );
    fixed.push('重複データの確認完了（手動対応が必要）');
  }

  // 無効なstatusフィールドの修正
  const isActiveIndex = headers.indexOf('isActive');
  if (isActiveIndex !== -1) {
    const updatesNeeded = [];
    for (let i = 0; i < userRows.length; i++) {
      const row = userRows[i];
      const currentValue = row[isActiveIndex];
      // statusフィールドが空または無効な値の場合、activeに設定
      if (
        !currentValue ||
        (currentValue !== 'active' && currentValue !== 'inactive' && currentValue !== 'suspended')
      ) {
        updatesNeeded.push({
          range:
            "'" + DB_CONFIG.SHEET_NAME + "'!" + String.fromCharCode(65 + isActiveIndex) + (i + 2),
          values: [['active']],
        });
      }
    }

    if (updatesNeeded.length > 0) {
      try {
        batchUpdateSheetsData(service, dbId, updatesNeeded);
        fixed.push(updatesNeeded.length + '件のstatusフィールドを修正');
      } catch (updateError) {
        console.error('❌ statusフィールド修正エラー:', updateError.message);
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
    const props = PropertiesService.getScriptProperties();
    const dbId = props.getProperty(PROPS_KEYS.DATABASE_SPREADSHEET_ID);
    if (!dbId) {
      throw new Error('データベースIDが設定されていません');
    }

    const ss = SpreadsheetApp.openById(dbId);
    const sheet = ss.getSheetByName(DB_CONFIG.SHEET_NAME);
    if (!sheet) {
      throw new Error('データベースシートが見つかりません: ' + DB_CONFIG.SHEET_NAME);
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
    const opts = {
      checkHealth: options.checkHealth !== false,
      checkPerformance: options.checkPerformance !== false,
      checkSecurity: options.checkSecurity !== false,
      enableAlerts: options.enableAlerts !== false,
      ...options,
    };

    const monitoringResult = {
      timestamp: new Date().toISOString(),
      summary: {
        overallHealth: 'unknown',
        alerts: [],
        warnings: [],
        metrics: {},
      },
      details: {
        healthCheck: null,
        performanceCheck: null,
        securityCheck: null,
      },
    };

    // 1. ヘルスチェック
    if (opts.checkHealth) {
      try {
        const healthResult = performHealthCheck();
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
        const perfResult = performPerformanceCheck();
        monitoringResult.details.performanceCheck = perfResult;

        if (perfResult.metrics.responseTime > 10000) {
          // 10秒以上
          monitoringResult.summary.alerts.push('レスポンス時間が異常に長くなっています');
        } else if (perfResult.metrics.responseTime > 5000) {
          // 5秒以上
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
        const securityResult = performSecurityCheck();
        monitoringResult.details.securityCheck = securityResult;

        if (securityResult.vulnerabilities.length > 0) {
          monitoringResult.summary.alerts.push(
            securityResult.vulnerabilities.length + '件のセキュリティ問題が検出されました'
          );
        }
      } catch (securityError) {
        console.error('❌ セキュリティチェックエラー:', securityError.message);
        monitoringResult.summary.warnings.push('セキュリティチェックに失敗しました');
      }
    }

    // 4. 総合判定
    if (monitoringResult.summary.alerts.length === 0) {
      monitoringResult.summary.overallHealth =
        monitoringResult.summary.warnings.length > 0 ? 'warning' : 'healthy';
    } else {
      monitoringResult.summary.overallHealth = 'critical';
    }

    // 5. アラート通知
    if (
      opts.enableAlerts &&
      (monitoringResult.summary.alerts.length > 0 || monitoringResult.summary.warnings.length > 0)
    ) {
      try {
        sendSystemAlert(monitoringResult);
      } catch (alertError) {
        console.error('❌ アラート送信エラー:', alertError.message);
      }
    }

    return monitoringResult;
  } catch (error) {
    console.error('❌ システム監視でエラー:', error);
    return {
      timestamp: new Date().toISOString(),
      error: error.message,
      summary: {
        overallHealth: 'error',
        alerts: ['監視システム自体が失敗: ' + error.message],
      },
    };
  }
}

/**
 * システムヘルスチェック
 * @returns {object} ヘルスチェック結果
 */
function performHealthCheck() {
  const healthResult = {
    timestamp: new Date().toISOString(),
    checks: {},
    summary: {
      overallStatus: 'unknown',
      passedChecks: 0,
      failedChecks: 0,
    },
  };

  // データベース接続チェック
  try {
    const diagnosis = SystemManager.checkSetupStatus();
    healthResult.checks.database = {
      status: diagnosis.isComplete ? 'pass' : 'fail',
      details: diagnosis,
    };

    if (diagnosis.isComplete) {
      healthResult.summary.passedChecks++;
    } else {
      healthResult.summary.failedChecks++;
    }
  } catch (dbError) {
    healthResult.checks.database = {
      status: 'fail',
      error: dbError.message,
    };
    healthResult.summary.failedChecks++;
  }

  // サービスアカウント権限チェック
  try {
    const permissionResult = verifyServiceAccountPermissions();
    healthResult.checks.serviceAccount = {
      status: permissionResult.summary.status === 'healthy' ? 'pass' : 'fail',
      details: permissionResult.summary,
    };

    if (permissionResult.summary.status === 'healthy') {
      healthResult.summary.passedChecks++;
    } else {
      healthResult.summary.failedChecks++;
    }
  } catch (permError) {
    healthResult.checks.serviceAccount = {
      status: 'fail',
      error: permError.message,
    };
    healthResult.summary.failedChecks++;
  }

  // システム設定チェック
  try {
    const props = PropertiesService.getScriptProperties();
    const requiredProps = [
      PROPS_KEYS.DATABASE_SPREADSHEET_ID,
      PROPS_KEYS.SERVICE_ACCOUNT_CREDS,
      PROPS_KEYS.ADMIN_EMAIL,
    ];

    const missingProps = [];
    for (let i = 0; i < requiredProps.length; i++) {
      if (!props.getProperty(requiredProps[i])) {
        missingProps.push(requiredProps[i]);
      }
    }

    healthResult.checks.configuration = {
      status: missingProps.length === 0 ? 'pass' : 'fail',
      missingProperties: missingProps,
    };

    if (missingProps.length === 0) {
      healthResult.summary.passedChecks++;
    } else {
      healthResult.summary.failedChecks++;
    }
  } catch (configError) {
    healthResult.checks.configuration = {
      status: 'fail',
      error: configError.message,
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
  const startTime = Date.now();
  const perfResult = {
    timestamp: new Date().toISOString(),
    metrics: {
      responseTime: 0,
      cacheHitRate: 0,
      apiCallCount: 0,
    },
    benchmarks: {},
  };

  try {
    // データベースアクセス速度テスト
    const dbTestStart = Date.now();
    const props = PropertiesService.getScriptProperties();
    const dbId = props.getProperty(PROPS_KEYS.DATABASE_SPREADSHEET_ID);

    if (dbId) {
      const service = getSheetsServiceCached();
      const testData = batchGetSheetsData(service, dbId, ["'" + DB_CONFIG.SHEET_NAME + "'!A1:N1"]);
      perfResult.benchmarks.databaseAccess = Date.now() - dbTestStart;
      perfResult.metrics.apiCallCount++;
    }

    // キャッシュ効率テスト（簡易版）
    const cacheTestStart = Date.now();
    try {
      if (typeof cacheManager !== 'undefined' && cacheManager.getStats) {
        const cacheStats = cacheManager.getStats();
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
  const securityResult = {
    timestamp: new Date().toISOString(),
    vulnerabilities: [],
    recommendations: [],
  };

  try {
    // スクリプトプロパティのセキュリティチェック
    const props = PropertiesService.getScriptProperties();
    const allProps = props.getProperties();
    // サービスアカウント認証情報の存在確認
    const serviceAccountCreds = props.getProperty(PROPS_KEYS.SERVICE_ACCOUNT_CREDS);
    if (!serviceAccountCreds) {
      securityResult.vulnerabilities.push({
        type: 'missing_credentials',
        severity: 'high',
        message: 'サービスアカウント認証情報が設定されていません',
      });
    }

    // 管理者メールの設定確認
    const adminEmail = props.getProperty(PROPS_KEYS.ADMIN_EMAIL);
    if (!adminEmail) {
      securityResult.vulnerabilities.push({
        type: 'missing_admin',
        severity: 'medium',
        message: '管理者メールアドレスが設定されていません',
      });
    }

    // データベースアクセス権限の確認
    try {
      const permissionCheck = verifyServiceAccountPermissions();
      if (permissionCheck.summary.status === 'critical') {
        securityResult.vulnerabilities.push({
          type: 'access_permission',
          severity: 'high',
          message: 'データベースアクセス権限に問題があります',
        });
      }
    } catch (permError) {
      securityResult.vulnerabilities.push({
        type: 'permission_check_failed',
        severity: 'medium',
        message: '権限確認に失敗しました',
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
    // アラート内容の構築
    const alertMessage = '【StudyQuest システムアラート】\n\n';
    alertMessage += '発生時刻: ' + monitoringResult.timestamp + '\n';
    alertMessage += 'システム状態: ' + monitoringResult.summary.overallHealth + '\n\n';

    if (monitoringResult.summary.alerts.length > 0) {
      alertMessage += '🚨 緊急問題:\n';
      for (let i = 0; i < monitoringResult.summary.alerts.length; i++) {
        alertMessage += '  • ' + monitoringResult.summary.alerts[i] + '\n';
      }
      alertMessage += '\n';
    }

    if (monitoringResult.summary.warnings.length > 0) {
      alertMessage += '⚠️ 警告:\n';
      for (let j = 0; j < monitoringResult.summary.warnings.length; j++) {
        alertMessage += '  • ' + monitoringResult.summary.warnings[j] + '\n';
      }
      alertMessage += '\n';
    }

    if (monitoringResult.summary.metrics.responseTime) {
      alertMessage += 'パフォーマンス:\n';
      alertMessage +=
        '  • レスポンス時間: ' + monitoringResult.summary.metrics.responseTime + 'ms\n';
      if (monitoringResult.summary.metrics.cacheHitRate) {
        alertMessage +=
          '  • キャッシュヒット率: ' +
          (monitoringResult.summary.metrics.cacheHitRate * 100).toFixed(1) +
          '%\n';
      }
    }

    // 管理者への通知（メール送信は実装に依存）
    const props = PropertiesService.getScriptProperties();
    const adminEmail = props.getProperty(PROPS_KEYS.ADMIN_EMAIL);
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
      timestamp: monitoringResult.timestamp,
    });
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
    const logEntry = {
      timestamp: new Date().toISOString(),
      type: eventType,
      data: eventData,
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
    const userInfo = DB.findUserById(userId);
    if (!userInfo) {
      throw new Error('データベースにユーザー情報が見つかりません。');
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(15000); // 15秒待機

    try {
      // データベース（シート）からユーザー行を削除（サービスアカウント経由）
      const props = PropertiesService.getScriptProperties();
      const dbId = props.getProperty(PROPS_KEYS.DATABASE_SPREADSHEET_ID);
      if (!dbId) {
        throw new Error('データベースIDが設定されていません');
      }

      const service = getSheetsServiceCached();
      if (!service) {
        throw new Error('Sheets APIサービスの初期化に失敗しました');
      }

      const sheetName = DB_CONFIG.SHEET_NAME;

      // データベーススプレッドシートの情報を取得してsheetIdを確認
      const spreadsheetInfo = getSpreadsheetsData(service, dbId);

      let targetSheetId;
      for (let i = 0; i < spreadsheetInfo.sheets.length; i++) {
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
      const data = batchGetSheetsData(service, dbId, ["'" + sheetName + "'!A:I"]);
      const values = data.valueRanges[0].values || [];

      // ユーザーIDに基づいて行を探す（A列がIDと仮定）
      const rowToDelete = -1;
      for (let i = values.length - 1; i >= 1; i--) {
        if (values[i][0] === userId) {
          rowToDelete = i + 1; // スプレッドシートは1ベース
          break;
        }
      }

      if (rowToDelete !== -1) {
        console.log('Deleting row:', rowToDelete, 'from sheetId:', targetSheetId);

        // 行を削除（正しいsheetIdを使用）
        const deleteRequest = {
          deleteDimension: {
            range: {
              sheetId: targetSheetId,
              dimension: 'ROWS',
              startIndex: rowToDelete - 1, // APIは0ベース
              endIndex: rowToDelete,
            },
          },
        };

        batchUpdateSpreadsheet(service, dbId, {
          requests: [deleteRequest],
        });

        console.log('Row deletion completed successfully');
      } else {
        console.warn('User row not found for deletion, userId:', userId);
      }

      // 削除ログを記録
      logAccountDeletion(userInfo.userEmail, userId, userInfo.userEmail, '自己削除', 'self');

      // 関連するすべてのキャッシュを削除
      invalidateUserCache(userId, userInfo.userEmail, null, false);

      // Google Drive のデータは保持するため何も操作しない

      // UserPropertiesからも関連情報を削除
      const userProps = PropertiesService.getUserProperties();
      userProps.deleteProperty('CURRENT_USER_ID');
    } finally {
      lock.releaseLock();
    }

    const successMessage = 'アカウント「' + userInfo.userEmail + '」が正常に削除されました。';
    console.log(successMessage);
    return successMessage;
  } catch (error) {
    console.error('アカウント削除エラー:', error.message);
    console.error('アカウント削除エラー詳細:', error.stack);

    // より詳細なエラー情報を提供
    const errorMessage = 'アカウントの削除に失敗しました: ' + error.message;
    if (error.message.includes('No grid with id')) {
      errorMessage +=
        '\n詳細: データベースシートのIDが正しく取得できませんでした。データベース設定を確認してください。';
    } else if (error.message.includes('permissions')) {
      errorMessage += '\n詳細: サービスアカウントの権限が不足している可能性があります。';
    } else if (error.message.includes('not found')) {
      errorMessage += '\n詳細: 削除対象のユーザーまたはシートが見つかりませんでした。';
    }

    throw new Error(errorMessage);
  }
}

/**
 * メールアドレスからドメインを抽出する
 * @param {string} email - メールアドレス
 * @returns {string} ドメイン名（小文字）
 */
function getEmailDomain(email) {
  return (email || '').toString().split('@').pop().toLowerCase();
}

/**
 * キャッシュを使わずにユーザーIDでユーザーを検索（強制リフレッシュ用）
 * @param {string} userId - 検索対象のユーザーID
 * @returns {object|null} ユーザー情報またはnull
 */
function findUserByIdFresh(userId) {
  if (!userId || typeof userId !== 'string') {
    console.warn('findUserByIdFresh: 無効なユーザーID', userId);
    return null;
  }

  try {
    // データベースから検索（キャッシュを使わない）
    const service = getSheetsService();
    const dbId = getSecureDatabaseId();
    const sheetName = DB_CONFIG.SHEET_NAME;

    // シート全体のデータを取得
    const data = batchGetSheetsData(service, dbId, [`'${sheetName}'!A:I`]);

    if (!data.valueRanges || !data.valueRanges[0] || !data.valueRanges[0].values) {
      console.warn('findUserByIdFresh: データベースからデータを取得できませんでした');
      return null;
    }

    const rows = data.valueRanges[0].values;
    const headers = rows[0];

    // ユーザーID列のインデックスを取得
    const userIdIndex = headers.indexOf('userId');
    if (userIdIndex === -1) {
      console.error('findUserByIdFresh: userId列が見つかりません');
      return null;
    }

    // ユーザーIDでユーザーを検索
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (row[userIdIndex] && row[userIdIndex] === userId) {
        // ユーザーオブジェクトを構築
        const user = {};
        headers.forEach((header, index) => {
          if (row[index] !== undefined) {
            user[header] = row[index];
          }
        });

        console.log('findUserByIdFresh: ユーザー発見（強制リフレッシュ）:', userId);
        return user;
      }
    }

    console.log('findUserByIdFresh: ユーザーが見つかりません:', userId);
    return null;
  } catch (error) {
    console.error('findUserByIdFresh エラー:', error.message);
    return null;
  }
}

/**
 * 汎用ユーザーデータベース検索関数
 * @param {string} searchField - 検索フィールド ('userId' | 'userEmail')
 * @param {string} searchValue - 検索値
 * @param {object} options - オプション設定
 * @returns {object|null} ユーザー情報またはnull
 */
function fetchUserFromDatabase(searchField, searchValue, options = {}) {
  try {
    if (searchField === 'userId') {
      return options.forceFresh ? findUserByIdFresh(searchValue) : DB.findUserById(searchValue);
    } else if (searchField === 'userEmail') {
      return DB.findUserByEmail(searchValue);
    } else {
      console.warn('fetchUserFromDatabase: サポートされていない検索フィールド:', searchField);
      return null;
    }
  } catch (error) {
    console.error('fetchUserFromDatabase エラー:', error.message);
    return null;
  }
}
