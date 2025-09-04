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
  createUser(userData) {
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
          row: newRow
        });

        return {
          success: true,
          userId: userData.userId,
          userEmail: userData.userEmail,
          configJson,
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
  /**
   * 🚀 最小限configJSON構築（root cause fix）
   * 新規ユーザー作成時は最小限のみ、設定時に追加
   */
  buildConfigJson(userData) {
    const now = new Date().toISOString();
    
    // ✅ userDataがconfigJson文字列を既に持っている場合はそれを使用
    if (userData.configJson && typeof userData.configJson === 'string') {
      try {
        return JSON.parse(userData.configJson);
      } catch (error) {
        console.warn('buildConfigJson: configJson解析エラー、最小構成で再構築', error.message);
      }
    }
    
    // 🎯 最小限configJSON構築（JSON bloat完全回避）
    return {
      setupStatus: userData.setupStatus || 'pending',
      appPublished: userData.appPublished || false,
      displaySettings: userData.displaySettings || {
        showNames: false,    // CLAUDE.md準拠：心理的安全性重視
        showReactions: false
      },
      createdAt: userData.createdAt || now,
      lastModified: userData.lastModified || now
    };
  },

  /**
   * メールでユーザー検索（CLAUDE.md準拠版）
   * @param {string} email - メールアドレス
   * @returns {Object|null} ユーザー情報またはnull
   */
  findUserByEmail(email) {
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
  findUserById(userId) {
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
  parseUserRow(headers, row) {
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

    // ✅ CLAUDE.md完全準拠：5フィールド構造を厳格実装
    // configJSON中心型 - 外側フィールド展開を完全削除
    // parsedConfigは参照専用、userObjへの展開は行わない（データ重複排除）

    return userObj;
  },

  /**
   * ユーザー更新（CLAUDE.md準拠版）
   * @param {string} userId - ユーザーID
   * @param {Object} updateData - 更新データ
   * @returns {Object} 更新結果
   */
  updateUser(userId, updateData) {
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

      // ✅ CLAUDE.md準拠：5フィールド構造に厳格準拠
      Object.keys(updateData).forEach(key => {
        if (key === 'userEmail' || key === 'isActive' || key === 'lastModified') {
          // 5フィールド構造の基本フィールドはそのまま
          return;
        }
        // その他はすべてconfigJsonに統合（統一データソース原則）
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
        updatedConfig,
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
   * SpreadsheetApp統一版
   */
  updateUserInDatabase(userId, dbUpdateData) {
    const dbId = getSecureDatabaseId();
    const sheetName = DB_CONFIG.SHEET_NAME;

    // SpreadsheetAppを使用してデータ取得
    const spreadsheet = SpreadsheetApp.openById(dbId);
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error('データベースシートが見つかりません');
    }

    const values = sheet.getDataRange().getValues();
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

    // SpreadsheetAppを使用してデータ更新
    sheet.getRange(rowIndex, 1, 1, 5).setValues([updateRow]);

    console.log('💾 CLAUDE.md準拠：5フィールド物理更新完了（SpreadsheetApp統一版）', {
      userId,
      row: rowIndex,
      configJsonSize: dbUpdateData.configJson.length
    });
  },

  /**
   * ユーザーキャッシュクリア
   */
  clearUserCache(userId, userEmail) {
    try {
      const cache = CacheService.getScriptCache();
      cache.remove(`user_id_${  userId}`);
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
  getAllUsers() {
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
  findUserByEmailNoCache(email) {
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
  },

  /**
   * 🗑️ ユーザー削除（管理者専用）
   * @param {string} targetUserId 削除対象ユーザーID
   * @param {string} reason 削除理由
   * @returns {Object} 削除結果
   */
  deleteUserAccountByAdmin(targetUserId, reason) {
    try {
      // 1. 基本検証
      if (!targetUserId || !reason || reason.length < 10) {
        throw new Error('削除対象ユーザーIDと削除理由（10文字以上）が必要です');
      }

      // 2. 管理者権限確認
      const currentUserEmail = User.email();
      const props = PropertiesService.getScriptProperties();
      const adminEmail = props.getProperty('ADMIN_EMAIL');
      
      if (currentUserEmail !== adminEmail) {
        throw new Error('管理者権限が必要です');
      }

      // 3. 削除対象ユーザー情報取得（キャッシュバイパス）
      const targetUser = this.findUserByIdNoCache(targetUserId);
      if (!targetUser) {
        throw new Error('削除対象ユーザーが見つかりません');
      }

      // 4. 自己削除防止
      if (targetUser.userEmail === currentUserEmail) {
        throw new Error('自分のアカウントは削除できません');
      }

      console.log('🗑️ ユーザー削除開始', { 
        targetUserId, 
        targetEmail: targetUser.userEmail,
        reason,
        executor: currentUserEmail 
      });

      // 5. データベースから削除
      const service = getSheetsService();
      const dbId = getSecureDatabaseId();
      const sheetName = DB_CONFIG.SHEET_NAME;

      // 全ユーザーデータを取得
      const data = batchGetSheetsData(service, dbId, [`'${sheetName}'!${DB_CONFIG.RANGE}`]);
      const rows = data.valueRanges[0].values;
      
      if (!rows || rows.length < 2) {
        throw new Error('データベースにユーザーデータがありません');
      }

      // ターゲットユーザーの行を特定
      let targetRowIndex = -1;
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === targetUserId) { // userId列（0番目）で判定
          targetRowIndex = i + 1; // Sheets APIは1ベース
          break;
        }
      }

      if (targetRowIndex === -1) {
        throw new Error('削除対象ユーザーがデータベースに見つかりません');
      }

      // 6. スプレッドシートから行削除
      const spreadsheet = SpreadsheetApp.openById(dbId);
      const sheet = spreadsheet.getSheetByName(sheetName);
      sheet.deleteRow(targetRowIndex);

      // 7. 🔥 重要：キャッシュ完全クリア
      this.invalidateUserCache(targetUserId, targetUser.userEmail);

      // 8. 削除ログ記録
      this.logAccountDeletion(targetUserId, targetUser.userEmail, reason, currentUserEmail);

      console.log('✅ ユーザー削除完了', { 
        targetUserId, 
        targetEmail: targetUser.userEmail,
        rowIndex: targetRowIndex 
      });

      return {
        success: true,
        message: `ユーザー ${targetUser.userEmail} を正常に削除しました`,
        deletedUser: {
          userId: targetUserId,
          email: targetUser.userEmail
        }
      };

    } catch (error) {
      console.error('❌ ユーザー削除エラー:', error.message);
      throw error;
    }
  },

  /**
   * 🚨 キャッシュ完全無効化
   * @param {string} userId ユーザーID  
   * @param {string} userEmail ユーザーメール
   */
  invalidateUserCache(userId, userEmail) {
    try {
      const cache = CacheService.getScriptCache();
      
      // ユーザー関連の全キャッシュキーをクリア
      const cacheKeys = [
        `user_${userId}`,
        `userinfo_${userId}`,
        'user_email_' + userEmail,
        `config_${userId}`,
        'all_users', // 全ユーザーリストキャッシュ
        'user_count'  // ユーザー数キャッシュ
      ];

      cacheKeys.forEach(key => {
        cache.remove(key);
      });

      console.log('🔥 キャッシュ完全クリア完了', { userId, userEmail, clearedKeys: cacheKeys.length });

    } catch (error) {
      console.warn('キャッシュクリアエラー:', error.message);
    }
  },

  /**
   * 📝 削除ログ記録
   */
  logAccountDeletion(targetUserId, targetEmail, reason, executorEmail) {
    try {
      const dbId = getSecureDatabaseId();
      const logSheetName = 'DeletionLogs';
      
      const spreadsheet = SpreadsheetApp.openById(dbId);
      let logSheet = spreadsheet.getSheetByName(logSheetName);
      
      // ログシートが存在しない場合は作成
      if (!logSheet) {
        logSheet = spreadsheet.insertSheet(logSheetName);
        logSheet.getRange(1, 1, 1, 6).setValues([
          ['timestamp', 'executorEmail', 'targetUserId', 'targetEmail', 'reason', 'deleteType']
        ]);
      }

      // ログエントリを追加
      logSheet.appendRow([
        new Date().toISOString(),
        executorEmail,
        targetUserId,
        targetEmail,
        reason,
        'admin_deletion'
      ]);

      console.log('📝 削除ログ記録完了', { targetUserId, targetEmail });

    } catch (error) {
      console.warn('削除ログ記録エラー:', error.message);
    }
  }
};

/**
 * CLAUDE.md準拠：グローバル関数（後方互換性）
 */
function updateUser(userId, updateData) {
  return DB.updateUser(userId, updateData);
}

function deleteUserAccountByAdmin(targetUserId, reason) {
  return DB.deleteUserAccountByAdmin(targetUserId, reason);
}

/**
 * 📊 データベースシートの初期化
 * @param {string} spreadsheetId - データベーススプレッドシートID
 */
function initializeDatabaseSheet(spreadsheetId) {
  try {
    console.log('データベースシート初期化開始:', spreadsheetId);
    
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    let sheet = spreadsheet.getSheetByName(DB_CONFIG.SHEET_NAME);
    
    // シートが存在しない場合は作成
    if (!sheet) {
      sheet = spreadsheet.insertSheet(DB_CONFIG.SHEET_NAME);
      console.log('Usersシートを新規作成');
    }
    
    // ヘッダー行が存在しない場合は設定
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, DB_CONFIG.HEADERS.length).setValues([DB_CONFIG.HEADERS]);
      console.log('ヘッダー行を設定:', DB_CONFIG.HEADERS);
    }
    
    console.log('✅ データベースシート初期化完了');
    return true;
    
  } catch (error) {
    console.error('データベースシート初期化エラー:', error.message);
    throw error;
  }
}

/**
 * 🧹 データベースクリーンアップ - 空のユーザーエントリを削除
 * @returns {Object} クリーンアップ結果
 */
function cleanupEmptyUsers() {
  try {
    console.log('🧹 データベースクリーンアップ開始...');
    
    const dbId = getSecureDatabaseId();
    const spreadsheet = SpreadsheetApp.openById(dbId);
    const sheet = spreadsheet.getSheetByName('Users');
    
    if (!sheet) {
      throw new Error('Usersシートが見つかりません');
    }
    
    const allData = sheet.getDataRange().getValues();
    const headers = allData[0];
    let deletedCount = 0;
    
    // 後ろから削除（インデックスのズレを防ぐ）
    for (let i = allData.length - 1; i > 0; i--) {
      const row = allData[i];
      const userEmail = row[1]; // emailIndex = 1
      
      // メールアドレスが空の行を削除
      if (!userEmail || userEmail === '') {
        sheet.deleteRow(i + 1); // シートの行番号は1ベース
        deletedCount++;
      }
    }
    
    console.log(`✅ クリーンアップ完了: ${deletedCount}件の空ユーザーを削除`);
    
    return {
      success: true,
      deletedCount,
      remainingUsers: sheet.getLastRow() - 1
    };
    
  } catch (error) {
    console.error('❌ クリーンアップエラー:', error.message);
    return { success: false, error: error.message };
  }
}

/**
 * 🔍 データベース診断関数 - 全ユーザーデータを表示
 * @returns {Object} データベースの全内容
 */
function debugShowAllUsers() {
  try {
    console.log('🔍 データベース診断開始...');
    
    const service = getSheetsService();
    const dbId = getSecureDatabaseId();
    const sheetName = 'Users';
    
    // スプレッドシートから直接データを取得
    const spreadsheet = SpreadsheetApp.openById(dbId);
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      console.error('❌ Usersシートが見つかりません');
      return { error: 'Usersシートが見つかりません' };
    }
    
    const allData = sheet.getDataRange().getValues();
    
    console.log('📊 データベース診断結果:', {
      spreadsheetId: dbId,
      sheetName: sheetName,
      totalRows: allData.length,
      headers: allData[0]
    });
    
    // 各ユーザーの詳細を表示
    for (let i = 1; i < allData.length; i++) {
      const row = allData[i];
      console.log(`ユーザー ${i}:`, {
        userId: row[0],
        userEmail: row[1],
        isActive: row[2],
        configJson: row[3] ? `${row[3].substring(0, 50)}...` : 'null',
        lastModified: row[4]
      });
    }
    
    // 削除ログも確認
    const deletionLogSheet = spreadsheet.getSheetByName('DeletionLogs');
    if (deletionLogSheet) {
      const deletionLogs = deletionLogSheet.getDataRange().getValues();
      console.log('🗑️ 削除ログ:', {
        totalDeletions: deletionLogs.length - 1,
        headers: deletionLogs[0]
      });
      
      for (let i = 1; i < Math.min(deletionLogs.length, 6); i++) { // 最新5件まで
        const log = deletionLogs[i];
        console.log(`削除ログ ${i}:`, {
          timestamp: log[0],
          executor: log[1],
          targetUserId: log[2],
          targetEmail: log[3],
          reason: log[4]
        });
      }
    }
    
    return {
      userCount: allData.length - 1,
      users: allData.slice(1).map(row => ({
        userId: row[0],
        email: row[1],
        isActive: row[2]
      }))
    };
    
  } catch (error) {
    console.error('❌ データベース診断エラー:', error.message);
    return { error: error.message };
  }
}