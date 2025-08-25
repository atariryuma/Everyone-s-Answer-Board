/**
 * @fileoverview APIサービス - クライアントからのAPIリクエストをハンドリング
 */

/**
 * APIリクエストをハンドリング（Core.gsの代替）
 * @param {string} action - APIアクション
 * @param {Object} params - パラメータ
 * @returns {Object} APIレスポンス
 */
function handleCoreApiRequest(action, params) {
  try {
    // ユーザー認証
    const userId = params.userId || getUserId();
    const user = getUserInfo(userId);
    
    if (!user && !isPublicAction(action)) {
      return { success: false, error: '認証が必要です' };
    }
    
    // アクションをルーティング
    switch (action) {
      // データ取得系
      case 'getInitialData':
        return handleGetInitialData(userId, params);
      case 'getSheetData':
        return handleGetSheetData(userId, params);
      case 'getIncrementalData':
        return handleGetIncrementalData(userId, params);
      case 'getSheetsList':
        return handleGetSheetsList(userId);
      case 'getAppConfig':
        return handleGetAppConfig(userId);
      case 'getCurrentUserStatus':
        return handleGetCurrentUserStatus(userId);
      
      // データ更新系
      case 'saveSheetConfig':
        return handleSaveSheetConfig(userId, params);
      case 'updateReaction':
        return handleUpdateReaction(userId, params);
      case 'updateFormSettings':
        return handleUpdateFormSettings(userId, params);
      case 'saveSystemConfig':
        return handleSaveSystemConfig(userId, params);
      
      // フォーム作成系
      case 'createQuickStartForm':
        return handleCreateQuickStartForm(userId);
      case 'createCustomForm':
        return handleCreateCustomForm(userId, params);
      
      // 管理者機能
      case 'getAllUsers':
        return handleGetAllUsers();
      case 'updateUserStatus':
        return handleUpdateUserStatus(params);
      case 'deleteUser':
        return handleDeleteUser(params);
      
      default:
        return { success: false, error: `不明なアクション: ${action}` };
    }
  } catch (error) {
    logError(error, 'handleCoreApiRequest', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.SYSTEM, { action });
    return { success: false, error: error.message };
  }
}

/**
 * 公開アクションかチェック
 */
function isPublicAction(action) {
  const publicActions = ['getLoginStatus', 'login'];
  return publicActions.includes(action);
}

/**
 * 初期データ取得
 */
function handleGetInitialData(userId, params) {
  try {
    const userInfo = getUserInfo(userId);
    if (!userInfo || !userInfo.spreadsheetId) {
      return { success: false, setupRequired: true };
    }
    
    const sheetName = params.sheetName || getCurrentSheetName(userInfo.spreadsheetId);
    const config = getSheetConfig(userId, sheetName);
    const data = getSheetData(userId, sheetName, params.classFilter, params.sortOrder, params.adminMode);
    
    return {
      success: true,
      data: data,
      config: config,
      userInfo: {
        userId: userInfo.userId,
        email: userInfo.adminEmail,
        isAdmin: isDeployUser()
      }
    };
  } catch (error) {
    logError(error, 'handleGetInitialData', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.DATABASE);
    return { success: false, error: error.message };
  }
}

/**
 * シートデータ取得
 */
function handleGetSheetData(userId, params) {
  try {
    const data = getSheetData(
      userId,
      params.sheetName,
      params.classFilter,
      params.sortOrder,
      params.adminMode
    );
    return { success: true, data: data };
  } catch (error) {
    logError(error, 'handleGetSheetData', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.DATABASE);
    return { success: false, error: error.message };
  }
}

/**
 * 増分データ取得
 */
function handleGetIncrementalData(userId, params) {
  try {
    const data = getIncrementalData(
      userId,
      params.sheetName,
      params.sinceRowCount,
      params.classFilter,
      params.sortOrder,
      params.adminMode
    );
    return { success: true, data: data };
  } catch (error) {
    logError(error, 'handleGetIncrementalData', ERROR_SEVERITY.LOW, ERROR_CATEGORIES.DATABASE);
    return { success: false, error: error.message };
  }
}

/**
 * シートリスト取得
 */
function handleGetSheetsList(userId) {
  try {
    const userInfo = getUserInfo(userId);
    if (!userInfo || !userInfo.spreadsheetId) {
      return { success: false, sheets: [] };
    }
    
    const spreadsheet = Spreadsheet.openById(userInfo.spreadsheetId);
    const sheets = spreadsheet.getSheets().map(sheet => ({
      name: sheet.getName(),
      id: sheet.getSheetId()
    }));
    
    return { success: true, sheets: sheets };
  } catch (error) {
    logError(error, 'handleGetSheetsList', ERROR_SEVERITY.LOW, ERROR_CATEGORIES.DATABASE);
    return { success: false, sheets: [] };
  }
}

/**
 * アプリ設定取得
 */
function handleGetAppConfig(userId) {
  try {
    const userInfo = getUserInfo(userId);
    const config = userInfo ? JSON.parse(userInfo.configJson || '{}') : {};
    return { success: true, config: config };
  } catch (error) {
    logError(error, 'handleGetAppConfig', ERROR_SEVERITY.LOW, ERROR_CATEGORIES.DATABASE);
    return { success: false, config: {} };
  }
}

/**
 * 現在のユーザーステータス取得
 */
function handleGetCurrentUserStatus(userId) {
  try {
    const userInfo = getUserInfo(userId);
    return {
      success: true,
      status: {
        isActive: userInfo ? userInfo.isActive === 'true' : false,
        lastAccessed: userInfo ? userInfo.lastAccessedAt : null,
        userId: userId,
        email: userInfo ? userInfo.adminEmail : ''
      }
    };
  } catch (error) {
    logError(error, 'handleGetCurrentUserStatus', ERROR_SEVERITY.LOW, ERROR_CATEGORIES.DATABASE);
    return { success: false, status: {} };
  }
}

/**
 * シート設定保存
 */
function handleSaveSheetConfig(userId, params) {
  try {
    const userInfo = getUserInfo(userId);
    if (!userInfo) {
      return { success: false, error: 'ユーザーが見つかりません' };
    }
    
    const config = JSON.parse(userInfo.configJson || '{}');
    config.sheets = config.sheets || {};
    config.sheets[params.sheetName] = params.config;
    
    updateUser(userId, { configJson: JSON.stringify(config) });
    
    return { success: true };
  } catch (error) {
    logError(error, 'handleSaveSheetConfig', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.DATABASE);
    return { success: false, error: error.message };
  }
}

/**
 * リアクション更新
 */
function handleUpdateReaction(userId, params) {
  try {
    const result = updateReaction(
      userId,
      params.rowIndex,
      params.reactionType,
      params.sheetName
    );
    return { success: true, result: result };
  } catch (error) {
    logError(error, 'handleUpdateReaction', ERROR_SEVERITY.LOW, ERROR_CATEGORIES.DATABASE);
    return { success: false, error: error.message };
  }
}

/**
 * フォーム設定更新
 */
function handleUpdateFormSettings(userId, params) {
  try {
    const userInfo = getUserInfo(userId);
    if (!userInfo || !userInfo.spreadsheetId) {
      return { success: false, error: 'スプレッドシートが設定されていません' };
    }
    
    // フォーム設定を更新する処理
    const config = JSON.parse(userInfo.configJson || '{}');
    config.formTitle = params.title;
    config.formDescription = params.description;
    
    updateUser(userId, { configJson: JSON.stringify(config) });
    
    return { success: true };
  } catch (error) {
    logError(error, 'handleUpdateFormSettings', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.DATABASE);
    return { success: false, error: error.message };
  }
}

/**
 * システム設定保存
 */
function handleSaveSystemConfig(userId, params) {
  try {
    if (!isDeployUser()) {
      return { success: false, error: '管理者権限が必要です' };
    }
    
    Object.keys(params.config).forEach(key => {
      Properties.setScriptProperty(key, params.config[key]);
    });
    
    return { success: true };
  } catch (error) {
    logError(error, 'handleSaveSystemConfig', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.SYSTEM);
    return { success: false, error: error.message };
  }
}

/**
 * クイックスタートフォーム作成
 */
function handleCreateQuickStartForm(userId) {
  try {
    const userInfo = getUserInfo(userId);
    if (!userInfo) {
      return { success: false, error: 'ユーザーが見つかりません' };
    }
    
    // フォームとスプレッドシートを作成
    const result = createQuickStartForm(userInfo.adminEmail, userId);
    
    // ユーザー情報を更新
    updateUser(userId, {
      spreadsheetId: result.spreadsheetId,
      spreadsheetUrl: result.spreadsheetUrl
    });
    
    return {
      success: true,
      formUrl: result.formUrl,
      spreadsheetUrl: result.spreadsheetUrl
    };
  } catch (error) {
    logError(error, 'handleCreateQuickStartForm', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.SYSTEM);
    return { success: false, error: error.message };
  }
}

/**
 * カスタムフォーム作成
 */
function handleCreateCustomForm(userId, params) {
  try {
    const userInfo = getUserInfo(userId);
    if (!userInfo) {
      return { success: false, error: 'ユーザーが見つかりません' };
    }
    
    // カスタムフォームとスプレッドシートを作成
    const result = createCustomForm(userInfo.adminEmail, userId, params.config);
    
    // ユーザー情報を更新
    updateUser(userId, {
      spreadsheetId: result.spreadsheetId,
      spreadsheetUrl: result.spreadsheetUrl
    });
    
    return {
      success: true,
      formUrl: result.formUrl,
      spreadsheetUrl: result.spreadsheetUrl
    };
  } catch (error) {
    logError(error, 'handleCreateCustomForm', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.SYSTEM);
    return { success: false, error: error.message };
  }
}

/**
 * 全ユーザー取得（管理者用）
 */
function handleGetAllUsers() {
  try {
    if (!isDeployUser()) {
      return { success: false, error: '管理者権限が必要です' };
    }
    
    const users = getActiveUsers();
    return { success: true, users: users };
  } catch (error) {
    logError(error, 'handleGetAllUsers', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.DATABASE);
    return { success: false, users: [] };
  }
}

/**
 * ユーザーステータス更新（管理者用）
 */
function handleUpdateUserStatus(params) {
  try {
    if (!isDeployUser()) {
      return { success: false, error: '管理者権限が必要です' };
    }
    
    updateUser(params.targetUserId, { isActive: params.isActive ? 'true' : 'false' });
    
    return { success: true };
  } catch (error) {
    logError(error, 'handleUpdateUserStatus', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.DATABASE);
    return { success: false, error: error.message };
  }
}

/**
 * ユーザー削除（管理者用）
 */
function handleDeleteUser(params) {
  try {
    const userId = getUserId(); // 実行者のID
    if (!isDeployUser()) {
      return { success: false, error: '管理者権限が必要です' };
    }
    
    deleteUser(params.targetUserId);
    logAction(userId, 'DELETE_USER', { targetUserId: params.targetUserId });
    
    return { success: true };
  } catch (error) {
    logError(error, 'handleDeleteUser', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.DATABASE);
    return { success: false, error: error.message };
  }
}