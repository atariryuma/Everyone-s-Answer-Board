/**
 * PageBackend.gs - Page.html/page.js.html専用バックエンド関数
 * 2024年GAS V8 Best Practices準拠
 * CLAUDE.md準拠：統一データソース、セキュリティ、エラーハンドリング
 */

/**
 * Page専用設定定数
 */
const PAGE_CONFIG = Object.freeze({
  CACHE_TTL: CORE.TIMEOUTS.MEDIUM,
  MAX_SHEETS: 100,
  DEFAULT_SHEET_NAME: 'Sheet1',
});

/**
 * 管理者権限チェック（Page専用）
 * ボード所有者またはシステム管理者の権限を確認
 * @param {string} userId - ユーザーID（オプション、未指定時は現在のユーザー）
 * @returns {boolean} 管理者権限があるかどうか
 */
function checkAdmin(userId = null) {
  try {
    const currentUserEmail = User.email();
    console.log('checkAdmin: 管理者権限チェック開始', {
      userId: userId ? `${userId.substring(0, 8)}...` : 'null',
      currentUserEmail: currentUserEmail ? `${currentUserEmail.substring(0, 10)}...` : 'null',
    });

    // userIdが指定されていない場合は現在のユーザーのuserIdを取得
    let targetUserId = userId;
    if (!targetUserId) {
      const user = DB.findUserByEmail(currentUserEmail);
      if (!user) {
        console.warn('checkAdmin: 現在のユーザーがデータベースに見つかりません');
        return false;
      }
      targetUserId = user.userId;
    }

    // 入力検証
    if (!SecurityValidator.isValidUUID(targetUserId)) {
      console.error('checkAdmin: 無効なユーザーID');
      return false;
    }

    // App.getAccess().verifyAccess()でadmin権限をチェック
    const accessResult = App.getAccess().verifyAccess(targetUserId, 'admin', currentUserEmail);

    const isAdmin = accessResult.allowed;
    console.log('checkAdmin: 権限チェック結果', {
      allowed: isAdmin,
      userType: accessResult.userType,
      reason: accessResult.reason,
    });

    return isAdmin;
  } catch (error) {
    console.error('checkAdmin エラー:', {
      error: error.message,
      stack: error.stack,
      userId: userId,
    });
    return false;
  }
}

/**
 * 利用可能シート一覧取得（Page専用）
 * 現在のユーザーのスプレッドシートからシート一覧を取得
 * @param {string} userId - ユーザーID（オプション）
 * @returns {Array<Object>} シート情報の配列
 */
function getAvailableSheets(userId = null) {
  try {
    const currentUserEmail = User.email();
    console.log('getAvailableSheets: シート一覧取得開始', {
      userId: userId ? `${userId.substring(0, 8)}...` : 'null',
    });

    // userIdが指定されていない場合は現在のユーザーのuserIdを取得
    let targetUserId = userId;
    if (!targetUserId) {
      const user = DB.findUserByEmail(currentUserEmail);
      if (!user) {
        throw new Error('現在のユーザーがデータベースに見つかりません');
      }
      targetUserId = user.userId;
    }

    // 入力検証
    if (!SecurityValidator.isValidUUID(targetUserId)) {
      throw new Error('無効なユーザーIDです');
    }

    // アクセス権限確認
    const accessResult = App.getAccess().verifyAccess(targetUserId, 'view', currentUserEmail);
    if (!accessResult.allowed) {
      throw new Error(`アクセス権限がありません: ${accessResult.reason}`);
    }

    // ユーザー情報からスプレッドシートIDを取得（統一データソース）
    const userInfo = DB.findUserById(targetUserId);
    if (!userInfo || !userInfo.spreadsheetId) {
      throw new Error('ユーザーのスプレッドシート情報が見つかりません');
    }

    // AdminPanelBackend.gsのgetSheetList()を利用
    const sheetList = getSheetList(userInfo.spreadsheetId);

    console.log('getAvailableSheets: シート一覧取得完了', {
      spreadsheetId: `${userInfo.spreadsheetId.substring(0, 20)}...`,
      sheetCount: sheetList.length,
    });

    return sheetList;
  } catch (error) {
    console.error('getAvailableSheets エラー:', {
      error: error.message,
      stack: error.stack,
      userId: userId,
    });
    throw error;
  }
}

/**
 * アクティブシートクリア（Page専用）
 * ユーザーのアクティブシート状態をリセット
 * @param {string} userId - ユーザーID（オプション）
 * @returns {Object} 処理結果
 */
function clearActiveSheet(userId = null) {
  try {
    const currentUserEmail = User.email();
    console.log('clearActiveSheet: アクティブシートクリア開始', {
      userId: userId ? `${userId.substring(0, 8)}...` : 'null',
    });

    // userIdが指定されていない場合は現在のユーザーのuserIdを取得
    let targetUserId = userId;
    if (!targetUserId) {
      const user = DB.findUserByEmail(currentUserEmail);
      if (!user) {
        throw new Error('現在のユーザーがデータベースに見つかりません');
      }
      targetUserId = user.userId;
    }

    // 入力検証
    if (!SecurityValidator.isValidUUID(targetUserId)) {
      throw new Error('無効なユーザーIDです');
    }

    // アクセス権限確認（編集権限が必要）
    const accessResult = App.getAccess().verifyAccess(targetUserId, 'edit', currentUserEmail);
    if (!accessResult.allowed) {
      throw new Error(`編集権限がありません: ${accessResult.reason}`);
    }

    // キャッシュクリア（関連するキャッシュを削除）
    const cacheKeys = [
      `sheet_data_${targetUserId}`,
      `active_sheet_${targetUserId}`,
      `sheet_headers_${targetUserId}`,
      `reaction_data_${targetUserId}`,
    ];

    let clearedCount = 0;
    cacheKeys.forEach((key) => {
      try {
        CacheService.getScriptCache().remove(key);
        clearedCount++;
      } catch (e) {
        console.warn(`キャッシュクリア失敗: ${key}`, e.message);
      }
    });

    // ユーザーのアクティブシート情報をリセット（必要に応じて）
    // 現在のシステムではsheetNameは保持するため、特別な処理は不要

    const result = {
      success: true,
      clearedCacheCount: clearedCount,
      userId: targetUserId,
      timestamp: new Date().toISOString(),
    };

    console.log('clearActiveSheet: 完了', result);
    return result;
  } catch (error) {
    console.error('clearActiveSheet エラー:', {
      error: error.message,
      stack: error.stack,
      userId: userId,
    });

    return {
      success: false,
      error: error.message,
      timestamp: new Date().toISOString(),
    };
  }
}

/**
 * バッチリアクション処理（Page専用）
 * 複数のリアクション操作を効率的に実行
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @param {Array<Object>} batchOperations - バッチ操作の配列
 * @returns {Object} 処理結果
 */
function addReactionBatch(requestUserId, batchOperations) {
  try {
    const currentUserEmail = User.email();
    console.log('addReactionBatch: バッチリアクション処理開始', {
      userId: requestUserId ? requestUserId.substring(0, 8) + '...' : 'null',
      operationsCount: Array.isArray(batchOperations) ? batchOperations.length : 0,
    });

    // 入力検証
    if (!Array.isArray(batchOperations) || batchOperations.length === 0) {
      throw new Error('バッチ操作が無効です');
    }

    // バッチサイズ制限（安全性のため）
    const MAX_BATCH_SIZE = 20;
    if (batchOperations.length > MAX_BATCH_SIZE) {
      throw new Error(`バッチサイズが制限を超えています (最大${MAX_BATCH_SIZE}件)`);
    }

    // ユーザーID検証
    if (!SecurityValidator.isValidUUID(requestUserId)) {
      throw new Error('無効なユーザーIDです');
    }

    // アクセス権限確認
    const accessResult = App.getAccess().verifyAccess(requestUserId, 'view', currentUserEmail);
    if (!accessResult.allowed) {
      throw new Error(`アクセス権限がありません: ${accessResult.reason}`);
    }

    // ボードオーナーの情報をDBから取得
    const boardOwnerInfo = DB.findUserById(requestUserId);
    if (!boardOwnerInfo || !boardOwnerInfo.spreadsheetId) {
      throw new Error('無効なボードです');
    }

    // バッチ処理結果を格納
    const batchResults = [];
    const processedRows = new Set(); // 重複行の追跡

    // シート名を取得（統一データソース使用）
    const userConfig = App.getConfig().getUserConfig(requestUserId);
    const sheetName = userConfig?.sheetName || 'フォームの回答 1';

    console.log('addReactionBatch: 処理対象シート', {
      spreadsheetId: boardOwnerInfo.spreadsheetId.substring(0, 20) + '...',
      sheetName: sheetName,
    });

    // バッチ操作を順次処理
    for (let i = 0; i < batchOperations.length; i++) {
      const operation = batchOperations[i];

      try {
        // 入力検証
        if (!operation.rowIndex || !operation.reaction) {
          console.warn('addReactionBatch: 無効な操作をスキップ', operation);
          continue;
        }

        // 個別のリアクション処理（Core.gsのaddReaction関数を呼び出し）
        const result = addReaction(
          requestUserId,
          operation.rowIndex,
          operation.reaction,
          sheetName
        );

        if (result && result.success) {
          batchResults.push({
            rowIndex: operation.rowIndex,
            reaction: operation.reaction,
            status: 'success',
            timestamp: new Date().toISOString(),
          });
          processedRows.add(operation.rowIndex);
        } else {
          console.warn('addReactionBatch: リアクション処理失敗', operation, result?.message);
          batchResults.push({
            rowIndex: operation.rowIndex,
            reaction: operation.reaction,
            status: 'error',
            message: result?.message || 'リアクション処理失敗',
          });
        }
      } catch (operationError) {
        console.error('addReactionBatch: 個別操作エラー', operation, operationError.message);
        batchResults.push({
          rowIndex: operation.rowIndex,
          reaction: operation.reaction,
          status: 'error',
          message: operationError.message,
        });
      }
    }

    const successCount = batchResults.filter((r) => r.status === 'success').length;
    console.log('addReactionBatch: バッチリアクション処理完了', {
      total: batchOperations.length,
      processed: processedRows.size,
      success: successCount,
    });

    return {
      success: true,
      processedCount: batchOperations.length,
      successCount: successCount,
      timestamp: new Date().toISOString(),
      details: batchResults,
    };
  } catch (error) {
    console.error('addReactionBatch エラー:', {
      error: error.message,
      stack: error.stack,
      requestUserId: requestUserId,
    });

    return {
      success: false,
      error: error.message,
      fallbackToIndividual: true, // クライアント側が個別処理にフォールバック可能
      timestamp: new Date().toISOString(),
    };
  }
}

/**
 * ボードデータ更新・キャッシュクリア（Page専用）
 * ユーザーのキャッシュをクリアして最新データを取得
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @returns {Object} 処理結果
 */
function refreshBoardData(requestUserId) {
  try {
    const currentUserEmail = User.email();
    console.log('refreshBoardData: ボードデータ更新開始', {
      userId: requestUserId ? requestUserId.substring(0, 8) + '...' : 'null',
    });

    // ユーザーID検証
    if (!SecurityValidator.isValidUUID(requestUserId)) {
      throw new Error('無効なユーザーIDです');
    }

    // アクセス権限確認
    const accessResult = App.getAccess().verifyAccess(requestUserId, 'view', currentUserEmail);
    if (!accessResult.allowed) {
      throw new Error(`アクセス権限がありません: ${accessResult.reason}`);
    }

    // ユーザー情報取得
    const userInfo = DB.findUserById(requestUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    // キャッシュクリア（関連するキャッシュを削除）
    const cacheKeys = [
      `sheet_data_${requestUserId}`,
      `board_data_${requestUserId}`,
      `user_config_${requestUserId}`,
      `reaction_data_${requestUserId}`,
      `sheet_headers_${requestUserId}`,
    ];

    let clearedCount = 0;
    cacheKeys.forEach((key) => {
      try {
        CacheService.getScriptCache().remove(key);
        clearedCount++;
      } catch (e) {
        console.warn(`refreshBoardData: キャッシュクリア失敗: ${key}`, e.message);
      }
    });

    // ユーザーキャッシュの無効化（Core.gsの関数を利用）
    try {
      if (typeof invalidateUserCache === 'function') {
        invalidateUserCache(requestUserId, userInfo.userEmail, userInfo.spreadsheetId, false);
      }
    } catch (invalidateError) {
      console.warn('refreshBoardData: ユーザーキャッシュ無効化エラー:', invalidateError.message);
    }

    // 最新の設定情報を取得
    let latestConfig = {};
    try {
      if (typeof getAppConfig === 'function') {
        latestConfig = getAppConfig(requestUserId);
      } else {
        // フォールバック: App.getConfig()を使用
        latestConfig = App.getConfig().getUserConfig(requestUserId) || {};
      }
    } catch (configError) {
      console.warn('refreshBoardData: 設定取得エラー:', configError.message);
      latestConfig = { status: 'error', message: '設定の取得に失敗しました' };
    }

    const result = {
      success: true,
      clearedCacheCount: clearedCount,
      config: latestConfig,
      userId: requestUserId,
      timestamp: new Date().toISOString(),
      message: 'ボードデータを更新しました',
    };

    console.log('refreshBoardData: 完了', {
      clearedCount: clearedCount,
      hasConfig: !!latestConfig,
      userId: requestUserId,
    });

    return result;
  } catch (error) {
    console.error('refreshBoardData エラー:', {
      error: error.message,
      stack: error.stack,
      requestUserId: requestUserId,
    });

    return {
      success: false,
      status: 'error',
      error: error.message,
      message: 'ボードデータの再読み込みに失敗しました: ' + error.message,
      timestamp: new Date().toISOString(),
    };
  }
}

/**
 * データ件数取得（Page専用）
 * 指定されたフィルタ条件でのデータ件数を取得
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @param {string} classFilter - クラスフィルタ
 * @param {string} sortOrder - ソート順
 * @param {boolean} adminMode - 管理者モード
 * @returns {Object} データ件数情報
 */
function getDataCount(requestUserId, classFilter, sortOrder, adminMode) {
  try {
    const currentUserEmail = User.email();
    console.log('getDataCount: データ件数取得開始', {
      userId: requestUserId ? requestUserId.substring(0, 8) + '...' : 'null',
      classFilter: classFilter,
      adminMode: adminMode,
    });

    // ユーザーID検証
    if (!SecurityValidator.isValidUUID(requestUserId)) {
      throw new Error('無効なユーザーIDです');
    }

    // アクセス権限確認
    const accessResult = App.getAccess().verifyAccess(requestUserId, 'view', currentUserEmail);
    if (!accessResult.allowed) {
      throw new Error(`アクセス権限がありません: ${accessResult.reason}`);
    }

    // ユーザー情報取得
    const userInfo = DB.findUserById(requestUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    // 設定情報を取得
    let config = {};
    try {
      config = JSON.parse(userInfo.configJson || '{}');
    } catch (e) {
      console.warn('getDataCount: configJson解析エラー:', e.message);
      config = {};
    }

    // スプレッドシート設定確認（統一データソース使用）
    const spreadsheetId = userInfo.spreadsheetId;
    const sheetName = config.sheetName || userInfo.sheetName || 'フォームの回答 1';

    if (!spreadsheetId || !sheetName) {
      return {
        count: 0,
        lastUpdate: new Date().toISOString(),
        status: 'error',
        message: 'スプレッドシート設定が不完全です',
      };
    }

    // データ件数をカウント
    let count = 0;
    try {
      if (typeof countSheetRows === 'function') {
        count = countSheetRows(spreadsheetId, sheetName, classFilter);
      } else {
        // フォールバック: 直接スプレッドシートにアクセス
        const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
        const sheet = spreadsheet.getSheetByName(sheetName);
        if (sheet) {
          count = Math.max(0, sheet.getLastRow() - 1); // ヘッダー行を除く
        }
      }
    } catch (countError) {
      console.error('getDataCount: カウントエラー:', countError.message);
      count = 0;
    }

    const result = {
      count: count,
      lastUpdate: new Date().toISOString(),
      status: 'success',
      spreadsheetId: spreadsheetId,
      sheetName: sheetName,
      classFilter: classFilter || null,
      adminMode: adminMode || false,
    };

    console.log('getDataCount: 完了', {
      count: count,
      sheetName: sheetName,
      classFilter: classFilter,
    });

    return result;
  } catch (error) {
    console.error('getDataCount エラー:', {
      error: error.message,
      stack: error.stack,
      requestUserId: requestUserId,
    });

    return {
      count: 0,
      lastUpdate: new Date().toISOString(),
      status: 'error',
      message: error.message,
    };
  }
}
