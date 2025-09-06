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
    'userId', // [0] UUID - 必須ID（検索用）
    'userEmail', // [1] メールアドレス - 必須認証（検索用）
    'isActive', // [2] アクティブ状態 - 必須フラグ（検索用）
    'configJson', // [3] 全設定統合 - メインデータ（JSON一括処理）
    'lastModified', // [4] 最終更新 - 監査用
  ]),

  // CLAUDE.md準拠：A:E範囲（5列のみ）
  RANGE: 'A:E',
  HEADER_RANGE: 'A1:E1',
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
      console.log('📊 createUser: configJSON中心型ユーザー作成開始', {
        userId: userData.userId,
        userEmail: userData.userEmail,
        timestamp: new Date().toISOString(),
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

        // Service Account使用
        const service = getSheetsServiceCached();
        if (!service) {
          throw new Error('Service Accountサービスが利用できません');
        }

        // ConfigManagerによる初期設定構築
        const configJson = ConfigManager.buildInitialConfig(userData);

        // CLAUDE.md準拠：5フィールドのみでデータ構築
        const newRow = [
          userData.userId,
          userData.userEmail,
          userData.isActive !== undefined ? userData.isActive : true, // Boolean値で設定
          JSON.stringify(configJson),
          new Date().toISOString(),
        ];

        console.log('📊 createUser: CLAUDE.md準拠5フィールド構築完了', {
          userId: userData.userId,
          configJsonSize: JSON.stringify(configJson).length,
          headers: DB_CONFIG.HEADERS,
          timestamp: new Date().toISOString(),
        });

        // Service Accountで新しい行を追加
        const appendResult = service.spreadsheets.values.append({
          spreadsheetId: dbId,
          range: `${sheetName}!A:E`,
          values: [newRow],
          valueInputOption: 'RAW',
          insertDataOption: 'INSERT_ROWS'
        });

        console.log('✅ createUser: configJSON中心型ユーザー作成完了', {
          userId: userData.userId,
          userEmail: userData.userEmail,
          configJsonFields: Object.keys(configJson).length,
          row: newRow,
        });

        return {
          success: true,
          userId: userData.userId,
          userEmail: userData.userEmail,
          configJson,
          timestamp: new Date().toISOString(),
        };
      } finally {
        lock.releaseLock();
      }
    } catch (error) {
      console.error('❌ createUser: configJSON中心型作成エラー:', {
        userId: userData.userId,
        error: error.message,
        stack: error.stack,
      });
      throw error;
    }
  },

  /**
   * 🚨 廃止予定：buildConfigJson（ConfigManagerに移行済み）
   * @deprecated ConfigManager.buildInitialConfigを使用
   */
  buildConfigJson_DEPRECATED(userData) {
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
        showNames: false, // CLAUDE.md準拠：心理的安全性重視
        showReactions: false,
      },
      createdAt: userData.createdAt || now,
      lastModified: userData.lastModified || now,
    };
  },

  /**
   * メールでユーザー検索（CLAUDE.md準拠版）
   * @param {string} email - メールアドレス
   * @returns {Object|null} ユーザー情報またはnull
   */
  // findUserByEmail - see implementation below (line 574)

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
    const cacheKey = `user_id_${  userId}`;

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

      // Service Accountでデータ取得
      const service = getSheetsServiceCached();
      if (!service) {
        throw new Error('Service Accountサービスが利用できません');
      }

      const batchGetResult = service.spreadsheets.values.batchGet({
        spreadsheetId: dbId,
        ranges: [`${sheetName}!A:E`]
      });
      
      const rows = batchGetResult.valueRanges[0].values || [];
      if (rows.length < 2) {
        console.log('findUserById: ユーザーデータが存在しません');
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
            CacheService.getScriptCache().put(
              cacheKey,
              JSON.stringify(userObj),
              DB_CONFIG.CACHE_TTL
            );
          } catch (cacheError) {
            console.warn('findUserById: キャッシュ保存エラー', cacheError.message);
          }

          console.log('✅ findUserById: configJSON中心型ユーザー発見', {
            userId,
            userEmail: userObj.userEmail,
            configFields: Object.keys(userObj.parsedConfig).length,
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

      console.log('findUserById: ユーザーが見つかりませんでした', { userId });
      return null;
    } catch (error) {
      console.error('❌ findUserById: configJSON中心型検索エラー:', {
        userId,
        error: error.message,
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
      lastModified: row[4] || '',
    };

    // configJsonをパース（無限再帰回避）
    try {
      userObj.parsedConfig = JSON.parse(userObj.configJson || '{}');
    } catch (e) {
      console.warn('configJson解析エラー:', {
        userId: userObj.userId,
        error: e.message,
      });
      userObj.parsedConfig = {};
    }

    // ✅ CLAUDE.md完全準拠：5フィールド構造を厳格実装
    // configJSON中心型 - 外側フィールド展開を完全削除
    // parsedConfigは参照専用、userObjへの展開は行わない（データ重複排除）

    return userObj;
  },

  /**
   * 🚨 廃止予定：updateUserConfig（ConfigManagerに移行済み）
   * @deprecated ConfigManager.saveConfigまたはConfigManager.updateConfigを使用
   */
  updateUserConfig_DEPRECATED(userId, configData) {
    try {
      console.log('🔥 updateUserConfig: configJSON重複回避更新開始', {
        userId,
        configFields: Object.keys(configData),
        timestamp: new Date().toISOString(),
      });

      // 現在のユーザーデータを取得
      const currentUser = this.findUserById(userId);
      if (!currentUser) {
        throw new Error('更新対象のユーザーが見つかりません');
      }

      // 🔥 重要：configDataをそのままJSONとして保存（マージなし）
      const dbUpdateData = {
        configJson: JSON.stringify(configData),
        lastModified: configData.lastModified || new Date().toISOString(),
      };

      // データベース更新実行
      this.updateUserInDatabase(userId, dbUpdateData);

      // キャッシュクリア
      this.clearUserCache(userId, currentUser.userEmail);

      console.log('✅ updateUserConfig: configJSON重複回避更新完了', {
        userId,
        configFields: Object.keys(configData),
        configSize: dbUpdateData.configJson.length,
      });

      return {
        success: true,
        userId,
        updatedConfig: configData,
        timestamp: dbUpdateData.lastModified,
      };
    } catch (error) {
      console.error('❌ updateUserConfig: configJSON重複回避更新エラー:', {
        userId,
        error: error.message,
        stack: error.stack,
      });
      throw error;
    }
  },

  /**
   * ユーザー更新（CLAUDE.md準拠版）
   * @param {string} userId - ユーザーID
   * @param {Object} updateData - 更新データ
   * @returns {Object} 更新結果
   */
  updateUser(userId, updateData) {
    try {
      console.log('📝 updateUser: configJSON中心型更新開始', {
        userId,
        updateFields: Object.keys(updateData),
        timestamp: new Date().toISOString(),
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
      Object.keys(updateData).forEach((key) => {
        if (key === 'userEmail' || key === 'isActive' || key === 'lastModified') {
          // 5フィールド構造の基本フィールドはそのまま
          return;
        }
        // その他はすべてconfigJsonに統合（統一データソース原則）
        updatedConfig[key] = updateData[key];
      });

      // 🚫 二重構造防止（第1層防御）: configJsonフィールドを絶対に含めない
      delete updatedConfig.configJson;
      delete updatedConfig.configJSON;

      // ネストした文字列形式のconfigJsonも検出して削除
      Object.keys(updatedConfig).forEach((key) => {
        if (key.toLowerCase() === 'configjson' || key === 'configJson') {
          console.warn(`⚠️ DB.updateUser: 危険なフィールド "${key}" を検出・削除`);
          delete updatedConfig[key];
        }
      });

      // lastModifiedを更新
      updatedConfig.lastModified = new Date().toISOString();

      // 🔥 CLAUDE.md準拠：完全configJSON中心型（重複フィールド削除）
      const dbUpdateData = {
        configJson: JSON.stringify(updatedConfig),
        lastModified: updatedConfig.lastModified,
      };

      // ⚡ DB基本フィールドの直接更新が必要な場合のみ追加
      if (updateData.userEmail !== undefined) {
        dbUpdateData.userEmail = updateData.userEmail;
      }
      if (updateData.isActive !== undefined) {
        dbUpdateData.isActive = updateData.isActive;
      }

      // データベース更新実行
      this.updateUserInDatabase(userId, dbUpdateData);

      // キャッシュクリア
      this.clearUserCache(userId, currentUser.userEmail);

      console.log('✅ updateUser: configJSON中心型更新完了', {
        userId,
        updatedFields: Object.keys(updateData),
        configSize: dbUpdateData.configJson.length,
      });

      return {
        success: true,
        userId,
        updatedConfig,
        timestamp: updatedConfig.lastModified,
      };
    } catch (error) {
      console.error('❌ updateUser: configJSON中心型更新エラー:', {
        userId,
        error: error.message,
        stack: error.stack,
      });
      throw error;
    }
  },

  /**
   * CLAUDE.md準拠：データベース物理更新（5フィールドのみ）
   * Service Account版
   */
  updateUserInDatabase(userId, dbUpdateData) {
    const dbId = getSecureDatabaseId();
    const sheetName = DB_CONFIG.SHEET_NAME;

    // Service Accountでデータ取得
    const service = getSheetsServiceCached();
    if (!service) {
      throw new Error('Service Accountサービスが利用できません');
    }

    const batchGetResult = service.spreadsheets.values.batchGet({
      spreadsheetId: dbId,
      ranges: [`${sheetName}!A:E`]
    });
    
    const values = batchGetResult.valueRanges[0].values || [];
    if (values.length === 0) {
      throw new Error('データベースが空です');
    }

    // ユーザーの行を特定
    let rowIndex = -1;
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === userId) {
        // userIdは1列目（インデックス0）
        rowIndex = i + 1; // 1-based index
        break;
      }
    }

    if (rowIndex === -1) {
      throw new Error('更新対象のユーザーが見つかりません');
    }

    // 🔥 CLAUDE.md準拠：5列更新（既存値保護版）
    const currentRow = values[rowIndex - 1]; // 0-based index for values array
    const updateRow = [
      userId, // 変更しない
      dbUpdateData.userEmail !== undefined ? dbUpdateData.userEmail : currentRow[1], // 既存userEmail保護
      dbUpdateData.isActive !== undefined ? dbUpdateData.isActive : currentRow[2], // 既存isActive保護
      dbUpdateData.configJson,
      dbUpdateData.lastModified,
    ];

    // Service Accountでデータ更新
    const updateResult = service.spreadsheets.values.update({
      spreadsheetId: dbId,
      range: `${sheetName}!A${rowIndex}:E${rowIndex}`,
      values: [updateRow],
      valueInputOption: 'RAW'
    });

    console.log('💾 CLAUDE.md準拠：5フィールド物理更新完了（Service Account版）', {
      userId,
      row: rowIndex,
      configJsonSize: dbUpdateData.configJson.length,
    });
  },

  /**
   * ユーザーキャッシュクリア
   */
  clearUserCache(userId, userEmail) {
    try {
      const cache = CacheService.getScriptCache();
      cache.remove(`user_id_${userId}`);
      if (userEmail) {
        cache.remove(`user_email_${  userEmail}`);
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
      console.log('📋 getAllUsers: configJSON中心型全ユーザー取得開始');

      const service = getSheetsService();
      const dbId = getSecureDatabaseId();
      const sheetName = DB_CONFIG.SHEET_NAME;

      // CLAUDE.md準拠：5列のみ取得
      const data = batchGetSheetsData(service, dbId, [`'${sheetName}'!${DB_CONFIG.RANGE}`]);

      if (!data.valueRanges || !data.valueRanges[0] || !data.valueRanges[0].values) {
        console.log('getAllUsers: ユーザーデータが見つかりません');
        return [];
      }

      const rows = data.valueRanges[0].values;
      if (rows.length < 2) {
        console.log('getAllUsers: データ行が存在しません');
        return [];
      }

      const headers = rows[0];
      const userRows = rows.slice(1);

      // 各行をユーザーオブジェクトに変換
      const users = userRows.map((row) => {
        return this.parseUserRow(headers, row);
      });

      console.log(`✅ getAllUsers: configJSON中心型で${users.length}件のユーザーデータを取得`);
      return users;
    } catch (error) {
      console.error('❌ getAllUsers: configJSON中心型取得エラー:', error.message);
      return [];
    }
  },

  /**
   * キャッシュなしでメールアドレスでユーザーを検索（ログイン専用）
   */
  findUserByEmail(email) {
    if (!email || typeof email !== 'string') {
      console.warn('findUserByEmail: 無効なメールアドレス', email);
      return null;
    }

    console.log('🔄 findUserByEmail: キャッシュをバイパスしてDB直接検索', { email });

    try {
      const service = getSheetsService();
      const dbId = getSecureDatabaseId();
      const sheetName = DB_CONFIG.SHEET_NAME;

      // CLAUDE.md準拠：5列のみ取得
      const data = batchGetSheetsData(service, dbId, [`'${sheetName}'!${DB_CONFIG.RANGE}`]);

      if (!data.valueRanges || !data.valueRanges[0] || !data.valueRanges[0].values) {
        console.warn('findUserByEmail: データベースからデータを取得できませんでした');
        return null;
      }

      const rows = data.valueRanges[0].values;
      if (rows.length < 2) {
        console.log('findUserByEmail: ユーザーデータがありません');
        return null;
      }

      const headers = rows[0];
      const userRows = rows.slice(1);

      // メールアドレス列のインデックスを取得（CLAUDE.md準拠：2列目）
      const emailIndex = 1;

      for (const row of userRows) {
        if (row[emailIndex] === email) {
          const user = this.parseUserRow(headers, row);

          console.log('✅ findUserByEmail: ユーザー発見（キャッシュバイパス）', {
            email,
            userId: user.userId,
            timestamp: new Date().toISOString(),
          });

          return user;
        }
      }

      console.log('findUserByEmail: ユーザーが見つかりませんでした', { email });
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
      const currentUserEmail = UserManager.getCurrentEmail();
      const props = PropertiesService.getScriptProperties();
      const adminEmail = props.getProperty('ADMIN_EMAIL');

      if (currentUserEmail !== adminEmail) {
        throw new Error('管理者権限が必要です');
      }

      // 3. 削除対象ユーザー情報取得（キャッシュバイパス）
      const targetUser = this.findUserById(targetUserId);
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
        executor: currentUserEmail,
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
        if (rows[i][0] === targetUserId) {
          // userId列（0番目）で判定
          targetRowIndex = i + 1; // Sheets APIは1ベース
          break;
        }
      }

      if (targetRowIndex === -1) {
        throw new Error('削除対象ユーザーがデータベースに見つかりません');
      }

      // 6. Service Accountでスプレッドシートから行削除
      const batchUpdateResult = service.spreadsheets.batchUpdate({
        spreadsheetId: dbId,
        requests: [{
          deleteDimension: {
            range: {
              sheetId: 0, // メインシートのID
              dimension: 'ROWS',
              startIndex: targetRowIndex - 1, // 0ベースに変換
              endIndex: targetRowIndex
            }
          }
        }]
      });

      // 7. 🔥 重要：キャッシュ完全クリア
      this.invalidateUserCache(targetUserId, targetUser.userEmail);

      // 8. 削除ログ記録
      this.logAccountDeletion(targetUserId, targetUser.userEmail, reason, currentUserEmail);

      console.log('✅ ユーザー削除完了', {
        targetUserId,
        targetEmail: targetUser.userEmail,
        rowIndex: targetRowIndex,
      });

      return {
        success: true,
        message: `ユーザー ${targetUser.userEmail} を正常に削除しました`,
        deletedUser: {
          userId: targetUserId,
          email: targetUser.userEmail,
        },
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
        `user_email_${  userEmail}`,
        `config_${userId}`,
        'all_users', // 全ユーザーリストキャッシュ
        'user_count', // ユーザー数キャッシュ
      ];

      cacheKeys.forEach((key) => {
        cache.remove(key);
      });

      console.log('🔥 キャッシュ完全クリア完了', {
        userId,
        userEmail,
        clearedKeys: cacheKeys.length,
      });
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

      // Service Accountでログ記録
      const service = getSheetsServiceCached();
      if (!service) {
        console.warn('Service Accountサービスが利用できないためログ記録をスキップ');
        return;
      }

      // ログエントリを追加
      const logEntry = [
        new Date().toISOString(),
        executorEmail,
        targetUserId,
        targetEmail,
        reason,
        'admin_deletion',
      ];

      try {
        const appendResult = service.spreadsheets.values.append({
          spreadsheetId: dbId,
          range: `${logSheetName}!A:F`,
          values: [logEntry],
          valueInputOption: 'RAW',
          insertDataOption: 'INSERT_ROWS'
        });
      } catch (error) {
        // シートが存在しない場合はヘッダー付きで作成
        if (error.message.includes('Unable to parse range')) {
          const headerRow = ['timestamp', 'executorEmail', 'targetUserId', 'targetEmail', 'reason', 'deleteType'];
          service.spreadsheets.values.append({
            spreadsheetId: dbId,
            range: `${logSheetName}!A1:F1`,
            values: [headerRow, logEntry],
            valueInputOption: 'RAW',
            insertDataOption: 'INSERT_ROWS'
          });
        } else {
          throw error;
        }
      }

      console.log('📝 削除ログ記録完了', { targetUserId, targetEmail });
    } catch (error) {
      console.warn('削除ログ記録エラー:', error.message);
    }
  },

  /**
   * 🧹 configJSON重複クリーンアップ（root cause fix）
   * ネストしたconfigJsonフィールドを正規化
   * @param {string} userId - 対象ユーザーID（省略時は全ユーザー）
   * @returns {Object} クリーンアップ結果
   */
  cleanupNestedConfigJson(userId = null) {
    try {
      console.log('🧹 cleanupNestedConfigJson: 重複configJSON修正開始', {
        targetUserId: userId || 'all_users',
        timestamp: new Date().toISOString(),
      });

      const users = userId ? [this.findUserById(userId)] : this.getAllUsers();
      const cleanupResults = {
        total: users.length,
        cleaned: 0,
        skipped: 0,
        errors: 0,
        details: [],
      };

      users.forEach((user) => {
        if (!user) return;

        try {
          const originalConfig = user.parsedConfig || {};
          let cleanedConfig = { ...originalConfig };
          let needsCleaning = false;

          // 🔥 重要：ネストしたconfigJsonフィールドを検出・修正
          if (cleanedConfig.configJson) {
            // configJsonフィールドが存在する場合、それを最上位に展開
            try {
              let nestedConfig;
              if (typeof cleanedConfig.configJson === 'string') {
                nestedConfig = JSON.parse(cleanedConfig.configJson);
              } else {
                nestedConfig = cleanedConfig.configJson;
              }

              // ネストされたconfigJsonを最上位にマージ
              cleanedConfig = { ...nestedConfig, ...cleanedConfig };

              // configJsonフィールド自体を削除
              delete cleanedConfig.configJson;
              needsCleaning = true;

              console.log(`🧹 ネストしたconfigJsonを修正: ${user.userEmail}`);
            } catch (parseError) {
              console.warn('configJson解析エラー:', parseError.message);
            }
          }

          // 🔥 その他の重複フィールドもクリーンアップ
          const duplicateFields = ['userId', 'userEmail', 'isActive', 'lastModified'];
          duplicateFields.forEach((field) => {
            if (cleanedConfig[field] !== undefined) {
              delete cleanedConfig[field];
              needsCleaning = true;
            }
          });

          if (needsCleaning) {
            // lastModifiedを更新
            cleanedConfig.lastModified = new Date().toISOString();

            // ConfigManager経由でクリーンなデータを保存
            ConfigManager.saveConfig(user.userId, cleanedConfig);

            cleanupResults.cleaned++;
            cleanupResults.details.push({
              userId: user.userId,
              email: user.userEmail,
              status: 'cleaned',
              removedFields: duplicateFields.filter((f) => originalConfig[f] !== undefined),
            });
          } else {
            cleanupResults.skipped++;
            cleanupResults.details.push({
              userId: user.userId,
              email: user.userEmail,
              status: 'skipped_no_issues',
            });
          }
        } catch (userError) {
          cleanupResults.errors++;
          cleanupResults.details.push({
            userId: user.userId,
            email: user.userEmail,
            status: 'error',
            error: userError.message,
          });
          console.error(`❌ ユーザークリーンアップエラー: ${user.userEmail}`, userError.message);
        }
      });

      console.log('✅ cleanupNestedConfigJson: 重複configJSON修正完了', cleanupResults);
      return {
        success: true,
        results: cleanupResults,
        timestamp: new Date().toISOString(),
      };
    } catch (error) {
      console.error('❌ cleanupNestedConfigJson: クリーンアップエラー', {
        error: error.message,
        timestamp: new Date().toISOString(),
      });
      return {
        success: false,
        error: error.message,
        timestamp: new Date().toISOString(),
      };
    }
  },
};

/**
 * CLAUDE.md準拠：グローバル関数（後方互換性）
 */
// updateUser global wrapper removed - use DB.updateUser() directly

// deleteUserAccountByAdmin global wrapper removed - use DB.deleteUserAccountByAdmin() directly

/**
 * 📊 データベースシートの初期化
 * @param {string} spreadsheetId - データベーススプレッドシートID
 */
function initializeDatabaseSheet(spreadsheetId) {
  try {
    console.log('データベースシート初期化開始:', spreadsheetId);

    // Service Accountでシート初期化
    const service = getSheetsServiceCached();
    if (!service) {
      throw new Error('Service Accountサービスが利用できません');
    }

    // シートの存在確認とヘッダー設定
    try {
      const batchGetResult = service.spreadsheets.values.batchGet({
        spreadsheetId,
        ranges: [`${DB_CONFIG.SHEET_NAME}!A1:E1`]
      });
      
      // ヘッダーがない場合は設定
      if (!batchGetResult.valueRanges[0].values || batchGetResult.valueRanges[0].values.length === 0) {
        console.log('Usersシートにヘッダーを設定');
        service.spreadsheets.values.update({
          spreadsheetId,
          range: `${DB_CONFIG.SHEET_NAME}!A1:E1`,
          values: [DB_CONFIG.HEADERS],
          valueInputOption: 'RAW'
        });
        console.log('ヘッダー行を設定:', DB_CONFIG.HEADERS);
      }
    } catch (error) {
      console.warn('シートの存在確認エラー:', error.message);
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
    
    // Service Accountでデータ取得
    const service = getSheetsServiceCached();
    if (!service) {
      throw new Error('Service Accountサービスが利用できません');
    }

    const batchGetResult = service.spreadsheets.values.batchGet({
      spreadsheetId: dbId,
      ranges: ['Users!A:E']
    });
    
    const rows = batchGetResult.valueRanges[0].values || [];
    if (rows.length < 2) {
      console.log('Usersシートにデータがありません');
      return { success: true, deletedCount: 0, message: '削除対象なし' };
    }

    const headers = rows[0];
    let deletedCount = 0;
    const deleteRequests = [];

    // 後ろから削除対象を特定（インデックスのズレを防ぐ）
    for (let i = rows.length - 1; i > 0; i--) {
      const row = rows[i];
      const userEmail = row[1]; // emailIndex = 1

      // メールアドレスが空の行を削除対象に
      if (!userEmail || userEmail === '') {
        deleteRequests.push({
          deleteDimension: {
            range: {
              sheetId: 0, // メインシートのID
              dimension: 'ROWS',
              startIndex: i, // 0ベース
              endIndex: i + 1
            }
          }
        });
        deletedCount++;
      }
    }

    // 一括削除実行
    if (deleteRequests.length > 0) {
      service.spreadsheets.batchUpdate({
        spreadsheetId: dbId,
        requests: deleteRequests
      });
    }

    console.log(`✅ クリーンアップ完了: ${deletedCount}件の空ユーザーを削除`);

    return {
      success: true,
      deletedCount,
      remainingUsers: rows.length - 1 - deletedCount,
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

    const service = getSheetsServiceCached();
    const dbId = getSecureDatabaseId();
    const sheetName = 'Users';

    // Service Accountでデータ取得
    if (!service) {
      console.error('❌ Service Accountサービスが利用できません');
      return { error: 'Service Accountサービスが利用できません' };
    }

    const batchGetResult = service.spreadsheets.values.batchGet({
      spreadsheetId: dbId,
      ranges: [`${sheetName}!A:E`]
    });
    
    const allData = batchGetResult.valueRanges[0].values || [];

    console.log('📊 データベース診断結果:', {
      spreadsheetId: dbId,
      sheetName,
      totalRows: allData.length,
      headers: allData[0],
    });

    // 各ユーザーの詳細を表示
    for (let i = 1; i < allData.length; i++) {
      const row = allData[i];
      console.log(`ユーザー ${i}:`, {
        userId: row[0],
        userEmail: row[1],
        isActive: row[2],
        configJson: row[3] ? `${row[3].substring(0, 50)}...` : 'null',
        lastModified: row[4],
      });
    }

    // 削除ログも確認
    const deletionLogSheet = spreadsheet.getSheetByName('DeletionLogs');
    if (deletionLogSheet) {
      const deletionLogs = deletionLogSheet.getDataRange().getValues();
      console.log('🗑️ 削除ログ:', {
        totalDeletions: deletionLogs.length - 1,
        headers: deletionLogs[0],
      });

      for (let i = 1; i < Math.min(deletionLogs.length, 6); i++) {
        // 最新5件まで
        const log = deletionLogs[i];
        console.log(`削除ログ ${i}:`, {
          timestamp: log[0],
          executor: log[1],
          targetUserId: log[2],
          targetEmail: log[3],
          reason: log[4],
        });
      }
    }

    return {
      userCount: allData.length - 1,
      users: allData.slice(1).map((row) => ({
        userId: row[0],
        email: row[1],
        isActive: row[2],
      })),
    };
  } catch (error) {
    console.error('❌ データベース診断エラー:', error.message);
    return { error: error.message };
  }
}

/**
 * 🚨 緊急修正用：ユーザーの基本フィールドを直接更新
 * @param {string} userId - ユーザーID
 * @param {Object} fields - 更新するフィールド {userEmail, isActive}
 */
function updateUserFields(userId, fields) {
  try {
    console.log('🚨 updateUserFields: 緊急修正開始', { userId, fields });

    if (!userId) {
      throw new Error('userIdが必要です');
    }

    const dbId = getSecureDatabaseId();
    const sheetName = DB_CONFIG.SHEET_NAME;
    const spreadsheet = SpreadsheetApp.openById(dbId);
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error('データベースシートが見つかりません');
    }

    const rows = sheet.getDataRange().getValues();
    const headers = rows[0];

    // ユーザー行を検索
    let targetRowIndex = -1;
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === userId) {
        // userId列で検索
        targetRowIndex = i + 1; // Sheetのインデックスは1ベース
        break;
      }
    }

    if (targetRowIndex === -1) {
      throw new Error('ユーザーが見つかりません');
    }

    // フィールドを更新
    if (fields.userEmail !== undefined) {
      const emailColIndex = headers.indexOf('userEmail');
      if (emailColIndex >= 0) {
        sheet.getRange(targetRowIndex, emailColIndex + 1).setValue(fields.userEmail);
        console.log('✅ userEmail更新完了:', fields.userEmail);
      }
    }

    if (fields.isActive !== undefined) {
      const activeColIndex = headers.indexOf('isActive');
      if (activeColIndex >= 0) {
        sheet.getRange(targetRowIndex, activeColIndex + 1).setValue(fields.isActive);
        console.log('✅ isActive更新完了:', fields.isActive);
      }
    }

    // lastModified更新
    const lastModifiedColIndex = headers.indexOf('lastModified');
    if (lastModifiedColIndex >= 0) {
      sheet.getRange(targetRowIndex, lastModifiedColIndex + 1).setValue(new Date().toISOString());
    }

    console.log('✅ updateUserFields: 緊急修正完了', { userId });
    return { success: true };
  } catch (error) {
    console.error('❌ updateUserFields: 緊急修正エラー:', error.message);
    throw error;
  }
}

/**
 * スプレッドシート情報取得（Sheets API使用）
 * @param {Object} service - Sheetsサービスオブジェクト
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {Object} スプレッドシート情報
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
    const {baseUrl} = service;
    const {accessToken} = service;

    // baseUrlが失われている場合の防御処理
    if (!baseUrl || typeof baseUrl !== 'string') {
      console.warn(
        '⚠️ baseUrlが見つかりません。デフォルトのGoogleSheetsAPIエンドポイントを使用します'
      );
      service.baseUrl = 'https://sheets.googleapis.com/v4/spreadsheets';
    }

    if (!accessToken || typeof accessToken !== 'string') {
      throw new Error(
        'アクセストークンが見つかりません。サービスオブジェクトが破損している可能性があります'
      );
    }

    // 安全なURL構築 - シート情報を含む基本的なメタデータを取得するために fields パラメータを追加
    const url = `${service.baseUrl}/${encodeURIComponent(spreadsheetId)}?fields=sheets.properties`;

    const response = UrlFetchApp.fetch(url, {
      headers: { Authorization: `Bearer ${accessToken}` },
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
        url: `${url.substring(0, 100)  }...`,
        spreadsheetId,
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
              `スプレッドシートへのアクセス権限がありません。サービスアカウント（${serviceAccountEmail}）をスプレッドシートの編集者として共有してください。`
            );
          }
        } catch (parseError) {
          console.warn('エラーレスポンスのJSON解析に失敗:', parseError.message);
        }
      }

      throw new Error(`Sheets API error: ${responseCode} - ${responseText}`);
    }

    let result;
    try {
      result = JSON.parse(responseText);
    } catch (parseError) {
      console.error('❌ JSON解析エラー:', parseError.message);
      console.error('❌ Response text:', responseText.substring(0, 200));
      throw new Error(`APIレスポンスのJSON解析に失敗: ${parseError.message}`);
    }

    // レスポンス構造の検証
    if (!result || typeof result !== 'object') {
      throw new Error(
        `無効なAPIレスポンス: オブジェクトが期待されましたが ${typeof result} を受信`
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
    throw new Error(`スプレッドシート情報取得に失敗しました: ${error.message}`);
  }
}
