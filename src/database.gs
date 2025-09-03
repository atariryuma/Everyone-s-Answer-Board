/**
 * @fileoverview データベース管理 - CLAUDE.md完全準拠版
 * configJSON中心型5フィールド構造の完全実装
 */

/**
 * CLAUDE.md準拠：configJSON中心型データベーススキーマ（5フィールド構造）
 * 絶対遵守：この構造以外は使用禁止
 */
const DB_CONFIG = Object.freeze({
  CACHE_TTL: CORE.TIMEOUTS.LONG, // 30秒
  BATCH_SIZE: 100,
  LOCK_TIMEOUT: 10000, // 10秒
  SHEET_NAME: 'Users',
  
  /**
   * CLAUDE.md絶対遵守：5フィールド構造
   * 全データはconfigJsonに統合し、単一JSON操作で全データ取得・更新
   */
  HEADERS: Object.freeze([
    'userId',           // [0] UUID - 必須ID（検索用）
    'userEmail',        // [1] メールアドレス - 必須認証（検索用）
    'isActive',         // [2] アクティブ状態 - 必須フラグ（検索用）
    'configJson',       // [3] 全設定統合 - メインデータ（JSON一括処理）
    'lastModified',     // [4] 最終更新 - 監査用
  ]),
  
  // CLAUDE.md準拠：A:E範囲（5列のみ）
  RANGE: 'A:E',
  HEADER_RANGE: 'A1:E1'
});

/**
 * CLAUDE.md準拠：統一データソース原則
 * データベース操作の名前空間オブジェクト
 */
const DB = {
  
  /**
   * ユーザー作成（CLAUDE.md準拠版）
   * @param {Object} userData - ユーザーデータ
   * @returns {Object} 作成結果
   */
  createUser: function(userData) {
    try {
      console.info('📊 createUser: configJSON中心型ユーザー作成開始', {
        userId: userData.userId,
        userEmail: userData.userEmail,
        timestamp: new Date().toISOString()
      });

      // Input validation (GAS 2025 best practices)
      if (!userData.userEmail || !userData.userId) {
        throw new Error('必須フィールドが不足しています: userEmail, userId');
      }

      // 重複チェック
      const existingUser = this.findUserByEmail(userData.userEmail);
      if (existingUser) {
        throw new Error('このメールアドレスは既に登録されています。');
      }

      // 同時登録による重複を防ぐためロックを取得
      const lock = LockService.getScriptLock();
      const lockAcquired = lock.tryLock(DB_CONFIG.LOCK_TIMEOUT);

      if (!lockAcquired) {
        const error = new Error('システムがビジー状態です。しばらく待ってから再試行してください。');
        console.error('❌ createUser: Lock acquisition failed', {
          userEmail: userData.userEmail,
          error: error.message,
        });
        throw error;
      }

      try {
        const props = PropertiesService.getScriptProperties();
        const dbId = props.getProperty(PROPS_KEYS.DATABASE_SPREADSHEET_ID);

        if (!dbId) {
          throw new Error('データベース設定が不完全です。システム管理者に連絡してください。');
        }

          const sheetName = DB_CONFIG.SHEET_NAME;

        // SpreadsheetAppを使用
        const spreadsheet = SpreadsheetApp.openById(dbId);
        const sheet = spreadsheet.getSheetByName(sheetName);
        if (!sheet) {
          throw new Error(`シート '${sheetName}' が見つかりません`);
        }

        // CLAUDE.md準拠：configJSON構築（全データを統合）
        const configJson = this.buildConfigJson(userData);

        // CLAUDE.md準拠：5フィールドのみでデータ構築
        const newRow = [
          userData.userId,
          userData.userEmail,
          userData.isActive || 'TRUE',
          JSON.stringify(configJson),
          new Date().toISOString()
        ];

        console.info('📊 createUser: CLAUDE.md準拠5フィールド構築完了', {
          userId: userData.userId,
          configJsonSize: JSON.stringify(configJson).length,
          headers: DB_CONFIG.HEADERS,
          timestamp: new Date().toISOString()
        });

        // 新しい行を追加
        sheet.appendRow(newRow);

        console.info('✅ createUser: configJSON中心型ユーザー作成完了', {
          userId: userData.userId,
          userEmail: userData.userEmail,
          configJsonFields: Object.keys(configJson).length,
          row: nextRow
        });

        return {
          success: true,
          userId: userData.userId,
          userEmail: userData.userEmail,
          configJson: configJson,
          timestamp: new Date().toISOString()
        };

      } finally {
        lock.releaseLock();
      }

    } catch (error) {
      console.error('❌ createUser: configJSON中心型作成エラー:', {
        userId: userData.userId,
        error: error.message,
        stack: error.stack
      });
      throw error;
    }
  },

  /**
   * CLAUDE.md準拠：configJSON構築
   * 全データをconfigJsonに統合（統一データソース原則）
   */
  buildConfigJson: function(userData) {
    const now = new Date().toISOString();
    
    return {
      // 監査情報（旧DB列から移行）
      createdAt: userData.createdAt || now,
      lastAccessedAt: userData.lastAccessedAt || now,
      
      // データソース情報（旧DB列から移行）
      spreadsheetId: userData.spreadsheetId || null,
      spreadsheetUrl: userData.spreadsheetUrl || null,
      sheetName: userData.sheetName || null,
      
      // フォーム・マッピング情報
      formUrl: userData.formUrl || null,
      columnMapping: userData.columnMapping || {},
      
      // アプリ設定
      setupStatus: userData.setupStatus || 'pending',
      appPublished: userData.appPublished || false,
      displaySettings: userData.displaySettings || {
        showNames: false, // CLAUDE.md準拠：心理的安全性重視
        showReactions: false
      },
      publishedAt: userData.publishedAt || null,
      appUrl: userData.appUrl || null,
      
      // メタデータ
      configJsonVersion: '1.0',
      claudeMdCompliant: true,
      createdWith: 'configJSON中心型システム'
    };
  },

  /**
   * メールでユーザー検索（CLAUDE.md準拠版）
   * @param {string} email - メールアドレス
   * @returns {Object|null} ユーザー情報またはnull
   */
  findUserByEmail: function(email) {
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
      console.log('🔍 findUserByEmail: configJSON中心型検索開始', { email });

      const dbId = getSecureDatabaseId();
      const sheetName = DB_CONFIG.SHEET_NAME;

      // SpreadsheetAppを使用してデータ取得
      const spreadsheet = SpreadsheetApp.openById(dbId);
      const sheet = spreadsheet.getSheetByName(sheetName);
      if (!sheet) {
        console.warn('findUserByEmail: シートが見つかりません');
        return null;
      }

      const rows = sheet.getDataRange().getValues();
      if (rows.length < 2) {
        console.info('findUserByEmail: ユーザーデータが存在しません');
        return null;
      }

      const headers = rows[0];
      
      // CLAUDE.md準拠：userEmailは2列目（インデックス1）
      const emailIndex = 1;

      // メールアドレスでマッチする行を検索
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        if (row[emailIndex] === email) {
          // CLAUDE.md準拠：5フィールド構造でユーザーオブジェクト構築
          const userObj = this.parseUserRow(headers, row);

          // キャッシュに保存
          try {
            CacheService.getScriptCache().put(cacheKey, JSON.stringify(userObj), DB_CONFIG.CACHE_TTL);
          } catch (cacheError) {
            console.warn('findUserByEmail: キャッシュ保存エラー', cacheError.message);
          }

          console.log('✅ findUserByEmail: configJSON中心型ユーザー発見', {
            email,
            userId: userObj.userId,
            configFields: Object.keys(userObj.parsedConfig).length
          });

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
      console.error('❌ findUserByEmail: configJSON中心型検索エラー:', {
        email,
        error: error.message
      });
      return null;
    }
  },

  /**
   * ユーザーIDでユーザー検索（CLAUDE.md準拠版）
   * @param {string} userId - ユーザーID
   * @returns {Object|null} ユーザー情報またはnull
   */
  findUserById: function(userId) {
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
      console.log('🔍 findUserById: configJSON中心型検索開始', { userId });

      const dbId = getSecureDatabaseId();
      const sheetName = DB_CONFIG.SHEET_NAME;

      // SpreadsheetAppを使用してデータ取得
      const spreadsheet = SpreadsheetApp.openById(dbId);
      const sheet = spreadsheet.getSheetByName(sheetName);
      if (!sheet) {
        console.warn('findUserById: シートが見つかりません');
        return null;
      }

      const rows = sheet.getDataRange().getValues();
      if (rows.length < 2) {
        console.info('findUserById: ユーザーデータが存在しません');
        return null;
      }

      const headers = rows[0];
      
      // CLAUDE.md準拠：userIdは1列目（インデックス0）
      const userIdIndex = 0;

      // ユーザーIDでマッチする行を検索
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        if (row[userIdIndex] === userId) {
          // CLAUDE.md準拠：5フィールド構造でユーザーオブジェクト構築
          const userObj = this.parseUserRow(headers, row);

          // キャッシュに保存
          try {
            CacheService.getScriptCache().put(cacheKey, JSON.stringify(userObj), DB_CONFIG.CACHE_TTL);
          } catch (cacheError) {
            console.warn('findUserById: キャッシュ保存エラー', cacheError.message);
          }

          console.log('✅ findUserById: configJSON中心型ユーザー発見', {
            userId,
            userEmail: userObj.userEmail,
            configFields: Object.keys(userObj.parsedConfig).length
          });

          return userObj;
        }
      }

      // 見つからなかった場合
      try {
        CacheService.getScriptCache().put(cacheKey, 'null', 60);
      } catch (cacheError) {
        console.warn('findUserById: nullキャッシュ保存エラー', cacheError.message);
      }

      console.info('findUserById: ユーザーが見つかりませんでした', { userId });
      return null;

    } catch (error) {
      console.error('❌ findUserById: configJSON中心型検索エラー:', {
        userId,
        error: error.message
      });
      return null;
    }
  },

  /**
   * CLAUDE.md準拠：行データをユーザーオブジェクトに変換
   * 統一データソース原則：configJsonから全データを展開
   */
  parseUserRow: function(headers, row) {
    // CLAUDE.md準拠：5フィールド構造
    const userObj = {
      userId: row[0] || '',
      userEmail: row[1] || '',
      isActive: row[2] || 'TRUE',
      configJson: row[3] || '{}',
      lastModified: row[4] || ''
    };

    // configJsonをパース（統一データソース原則）
    try {
      userObj.parsedConfig = JSON.parse(userObj.configJson);
    } catch (e) {
      console.warn('configJson解析エラー:', {
        userId: userObj.userId,
        error: e.message
      });
      userObj.parsedConfig = {};
    }

    // CLAUDE.md準拠：統一データソースからデータ展開
    const config = userObj.parsedConfig;
    userObj.spreadsheetId = config.spreadsheetId || null;
    userObj.sheetName = config.sheetName || null;
    userObj.columnMapping = config.columnMapping || {};
    userObj.displaySettings = config.displaySettings || { showNames: false, showReactions: false };
    userObj.appPublished = config.appPublished || false;
    userObj.formUrl = config.formUrl || null;
    userObj.createdAt = config.createdAt || null;
    userObj.lastAccessedAt = config.lastAccessedAt || null;
    userObj.publishedAt = config.publishedAt || null;
    userObj.appUrl = config.appUrl || null;

    return userObj;
  },

  /**
   * ユーザー更新（CLAUDE.md準拠版）
   * @param {string} userId - ユーザーID
   * @param {Object} updateData - 更新データ
   * @returns {Object} 更新結果
   */
  updateUser: function(userId, updateData) {
    try {
      console.info('📝 updateUser: configJSON中心型更新開始', {
        userId,
        updateFields: Object.keys(updateData),
        timestamp: new Date().toISOString()
      });

      // 現在のユーザーデータを取得
      const currentUser = this.findUserById(userId);
      if (!currentUser) {
        throw new Error('更新対象のユーザーが見つかりません');
      }

      // CLAUDE.md準拠：現在のconfigJsonと更新データをマージ
      const currentConfig = currentUser.parsedConfig || {};
      const updatedConfig = { ...currentConfig };

      // 更新データをconfigJsonに統合（統一データソース原則）
      Object.keys(updateData).forEach(key => {
        if (key === 'userEmail' || key === 'isActive') {
          // 基本フィールドはそのまま
          return;
        }
        // その他はすべてconfigJsonに統合
        updatedConfig[key] = updateData[key];
      });

      // lastModifiedを更新
      updatedConfig.lastModified = new Date().toISOString();

      // CLAUDE.md準拠：5フィールドで更新データを構築
      const dbUpdateData = {
        userEmail: updateData.userEmail || currentUser.userEmail,
        isActive: updateData.isActive || currentUser.isActive,
        configJson: JSON.stringify(updatedConfig),
        lastModified: updatedConfig.lastModified
      };

      // データベース更新実行
      this.updateUserInDatabase(userId, dbUpdateData);

      // キャッシュクリア
      this.clearUserCache(userId, currentUser.userEmail);

      console.info('✅ updateUser: configJSON中心型更新完了', {
        userId,
        updatedFields: Object.keys(updateData),
        configSize: dbUpdateData.configJson.length
      });

      return {
        success: true,
        userId,
        updatedConfig: updatedConfig,
        timestamp: updatedConfig.lastModified
      };

    } catch (error) {
      console.error('❌ updateUser: configJSON中心型更新エラー:', {
        userId,
        error: error.message,
        stack: error.stack
      });
      throw error;
    }
  },

  /**
   * CLAUDE.md準拠：データベース物理更新（5フィールドのみ）
   */
  updateUserInDatabase: function(userId, dbUpdateData) {
    const dbId = getSecureDatabaseId();
    const service = getSheetsServiceCached();
    const sheetName = DB_CONFIG.SHEET_NAME;

    // CLAUDE.md準拠：5列のみ取得
    const data = batchGetSheetsData(service, dbId, [`'${sheetName}'!${DB_CONFIG.RANGE}`]);
    const values = data.valueRanges[0].values || [];

    if (values.length === 0) {
      throw new Error('データベースが空です');
    }

    // ユーザーの行を特定
    let rowIndex = -1;
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === userId) { // userIdは1列目（インデックス0）
        rowIndex = i + 1; // 1-based index
        break;
      }
    }

    if (rowIndex === -1) {
      throw new Error('更新対象のユーザーが見つかりません');
    }

    // CLAUDE.md準拠：5列更新
    const updateRow = [
      userId, // 変更しない
      dbUpdateData.userEmail,
      dbUpdateData.isActive,
      dbUpdateData.configJson,
      dbUpdateData.lastModified
    ];

    const range = `'${sheetName}'!A${rowIndex}:E${rowIndex}`;
    updateSheetsData(service, dbId, range, [updateRow]);

    console.log('💾 CLAUDE.md準拠：5フィールド物理更新完了', {
      userId,
      row: rowIndex,
      range: range,
      configJsonSize: dbUpdateData.configJson.length
    });
  },

  /**
   * ユーザーキャッシュクリア
   */
  clearUserCache: function(userId, userEmail) {
    try {
      const cache = CacheService.getScriptCache();
      cache.remove('user_id_' + userId);
      if (userEmail) {
        cache.remove('user_email_' + userEmail);
      }
    } catch (error) {
      console.warn('キャッシュクリアエラー:', error.message);
    }
  },

  /**
   * 全ユーザー取得（CLAUDE.md準拠版）
   */
  getAllUsers: function() {
    try {
      console.info('📋 getAllUsers: configJSON中心型全ユーザー取得開始');

      const service = getSheetsService();
      const dbId = getSecureDatabaseId();
      const sheetName = DB_CONFIG.SHEET_NAME;

      // CLAUDE.md準拠：5列のみ取得
      const data = batchGetSheetsData(service, dbId, [`'${sheetName}'!${DB_CONFIG.RANGE}`]);

      if (!data.valueRanges || !data.valueRanges[0] || !data.valueRanges[0].values) {
        console.info('getAllUsers: ユーザーデータが見つかりません');
        return [];
      }

      const rows = data.valueRanges[0].values;
      if (rows.length < 2) {
        console.info('getAllUsers: データ行が存在しません');
        return [];
      }

      const headers = rows[0];
      const userRows = rows.slice(1);

      // 各行をユーザーオブジェクトに変換
      const users = userRows.map(row => {
        return this.parseUserRow(headers, row);
      });

      console.info(`✅ getAllUsers: configJSON中心型で${users.length}件のユーザーデータを取得`);
      return users;

    } catch (error) {
      console.error('❌ getAllUsers: configJSON中心型取得エラー:', error.message);
      return [];
    }
  },

  /**
   * キャッシュなしでメールアドレスでユーザーを検索（ログイン専用）
   */
  findUserByEmailNoCache: function(email) {
    if (!email || typeof email !== 'string') {
      console.warn('findUserByEmailNoCache: 無効なメールアドレス', email);
      return null;
    }

    console.info('🔄 findUserByEmailNoCache: キャッシュをバイパスしてDB直接検索', { email });

    try {
      const service = getSheetsService();
      const dbId = getSecureDatabaseId();
      const sheetName = DB_CONFIG.SHEET_NAME;

      // CLAUDE.md準拠：5列のみ取得
      const data = batchGetSheetsData(service, dbId, [`'${sheetName}'!${DB_CONFIG.RANGE}`]);

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

      // メールアドレス列のインデックスを取得（CLAUDE.md準拠：2列目）
      const emailIndex = 1;

      for (const row of userRows) {
        if (row[emailIndex] === email) {
          const user = this.parseUserRow(headers, row);

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
      return null;
    }
  }
};

/**
 * CLAUDE.md準拠：グローバル関数（後方互換性）
 */
function updateUser(userId, updateData) {
  return DB.updateUser(userId, updateData);
}