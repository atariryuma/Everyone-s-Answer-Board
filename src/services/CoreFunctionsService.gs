/**
 * @fileoverview コア機能サービス - 既存Core.gsの重要機能を移植
 */

/**
 * 統合初期データ取得
 * @param {string} requestUserId - ユーザーID
 * @param {string} targetSheetName - シート名
 * @param {boolean} lightweightMode - 軽量モード
 * @returns {Object} 初期データ
 */
function getInitialData(requestUserId, targetSheetName, lightweightMode) {
  try {
    // ユーザー情報取得
    const currentUserId = requestUserId || getUserId();
    const userInfo = getUserInfo(currentUserId);
    
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    
    // 設定情報の解析
    const configJson = JSON.parse(userInfo.configJson || '{}');
    
    // セットアップステップの決定
    const setupStep = getSetupStep(userInfo, configJson);
    
    // シート一覧の取得
    const sheets = getSheetsList(currentUserId);
    
    // アプリURLの生成
    const appUrls = generateUserUrls(currentUserId);
    
    // 回答数の取得
    let answerCount = 0;
    let totalReactions = 0;
    if (configJson.publishedSpreadsheetId && configJson.publishedSheetName) {
      const responseData = getResponsesData(currentUserId, configJson.publishedSheetName);
      if (responseData && responseData.status === 'success') {
        answerCount = responseData.data ? responseData.data.length : 0;
        totalReactions = answerCount * 2;
      }
    }
    
    // レスポンスの構築
    const response = {
      userInfo: {
        userId: userInfo.userId,
        adminEmail: userInfo.adminEmail,
        isActive: userInfo.isActive,
        lastAccessedAt: userInfo.lastAccessedAt,
        spreadsheetId: userInfo.spreadsheetId,
        spreadsheetUrl: userInfo.spreadsheetUrl,
        configJson: userInfo.configJson
      },
      appUrls: appUrls,
      setupStep: setupStep,
      activeSheetName: configJson.publishedSheetName || null,
      webAppUrl: appUrls.webApp,
      isPublished: !!configJson.appPublished,
      answerCount: answerCount,
      totalReactions: totalReactions,
      config: {
        publishedSheetName: configJson.publishedSheetName || '',
        opinionHeader: configJson.opinionHeader || '',
        nameHeader: configJson.nameHeader || '',
        classHeader: configJson.classHeader || '',
        showNames: configJson.showNames || false,
        showCounts: configJson.showCounts !== undefined ? configJson.showCounts : false,
        displayMode: configJson.displayMode || 'anonymous',
        setupStatus: configJson.setupStatus || 'initial',
        isPublished: !!configJson.appPublished
      },
      allSheets: sheets,
      sheetNames: sheets,
      customFormInfo: configJson.formUrl ? {
        title: configJson.formTitle || 'カスタムフォーム',
        mainQuestion: configJson.mainQuestion || configJson.publishedSheetName || '質問',
        formUrl: configJson.formUrl
      } : null,
      _meta: {
        apiVersion: 'integrated_v1',
        executionTime: null,
        includedApis: ['getCurrentUserStatus', 'getUserId', 'getAppConfig']
      }
    };
    
    // シート詳細の取得（軽量モードでない場合）
    if (!lightweightMode && targetSheetName && userInfo.spreadsheetId) {
      try {
        const sheetDetails = getSheetDetails(currentUserId, targetSheetName);
        response.sheetDetails = sheetDetails;
      } catch (error) {
        warnLog('シート詳細取得エラー:', error.message);
      }
    }
    
    return response;
    
  } catch (error) {
    logError(error, 'getInitialData', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.SYSTEM);
    throw error;
  }
}

/**
 * 公開シートデータ取得
 * @param {string} requestUserId - ユーザーID
 * @param {string} classFilter - クラスフィルター
 * @param {string} sortOrder - ソート順
 * @param {boolean} adminMode - 管理者モード
 * @param {boolean} bypassCache - キャッシュバイパス
 * @returns {Object} シートデータ
 */
function getPublishedSheetData(requestUserId, classFilter, sortOrder, adminMode, bypassCache) {
  try {
    const userInfo = getUserInfo(requestUserId);
    if (!userInfo) {
      return { success: false, error: 'ユーザーが見つかりません' };
    }
    
    const config = JSON.parse(userInfo.configJson || '{}');
    const sheetName = config.publishedSheetName;
    
    if (!sheetName || !userInfo.spreadsheetId) {
      return { success: false, error: '公開シートが設定されていません' };
    }
    
    // キャッシュキーの生成
    const cacheKey = `sheet_data_${requestUserId}_${sheetName}_${classFilter}_${sortOrder}_${adminMode}`;
    
    // キャッシュチェック（バイパスでない場合）
    if (!bypassCache) {
      const cached = getCacheValue(cacheKey);
      if (cached) {
        return { success: true, data: cached, fromCache: true };
      }
    }
    
    // データ取得
    const data = getSheetData(requestUserId, sheetName, classFilter, sortOrder, adminMode);
    
    // キャッシュに保存
    if (data && data.length > 0) {
      setCacheValue(cacheKey, data, 60); // 1分間キャッシュ
    }
    
    return { success: true, data: data, fromCache: false };
    
  } catch (error) {
    logError(error, 'getPublishedSheetData', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.DATABASE);
    return { success: false, error: error.message };
  }
}

/**
 * ユーザーステータス取得
 * @param {string} requestUserId - ユーザーID
 * @returns {Object} ステータス情報
 */
function getCurrentUserStatus(requestUserId) {
  try {
    const userInfo = getUserInfo(requestUserId);
    const config = userInfo ? JSON.parse(userInfo.configJson || '{}') : {};
    
    return {
      success: true,
      status: {
        isActive: userInfo ? userInfo.isActive === 'true' : false,
        setupStatus: config.setupStatus || 'initial',
        appPublished: !!config.appPublished,
        formCreated: !!config.formCreated,
        spreadsheetId: userInfo ? userInfo.spreadsheetId : null,
        lastAccessed: userInfo ? userInfo.lastAccessedAt : null,
        userId: requestUserId,
        email: userInfo ? userInfo.adminEmail : ''
      }
    };
  } catch (error) {
    logError(error, 'getCurrentUserStatus', ERROR_SEVERITY.LOW, ERROR_CATEGORIES.DATABASE);
    return { success: false, status: {} };
  }
}

/**
 * セットアップステップを判定
 * @param {Object} userInfo - ユーザー情報
 * @param {Object} configJson - 設定JSON
 * @returns {number} セットアップステップ
 */
function getSetupStep(userInfo, configJson) {
  // ステップ1: 初期状態
  if (!userInfo || !userInfo.spreadsheetId) {
    return 1;
  }
  
  // ステップ2: スプレッドシート作成済み
  if (!configJson.formCreated) {
    return 2;
  }
  
  // ステップ3: フォーム作成済み
  if (!configJson.appPublished) {
    return 3;
  }
  
  // ステップ4: 公開済み
  return 4;
}

/**
 * シート一覧取得
 * @param {string} userId - ユーザーID
 * @returns {Array} シート名の配列
 */
function getSheetsList(userId) {
  try {
    const userInfo = getUserInfo(userId);
    if (!userInfo || !userInfo.spreadsheetId) {
      return [];
    }
    
    const spreadsheet = Spreadsheet.openById(userInfo.spreadsheetId);
    const sheets = spreadsheet.getSheets();
    
    return sheets.map(sheet => sheet.getName());
  } catch (error) {
    logError(error, 'getSheetsList', ERROR_SEVERITY.LOW, ERROR_CATEGORIES.DATABASE);
    return [];
  }
}

/**
 * ユーザー用URLを生成
 * @param {string} userId - ユーザーID
 * @returns {Object} URL情報
 */
function generateUserUrls(userId) {
  const baseUrl = getWebAppUrl();
  
  return {
    webApp: baseUrl,
    admin: `${baseUrl}?mode=admin&userId=${userId}`,
    setup: `${baseUrl}?mode=setup&userId=${userId}`,
    published: `${baseUrl}?userId=${userId}`
  };
}

/**
 * WebアプリのURLを取得
 * @returns {string} URL
 */
function getWebAppUrl() {
  try {
    return ScriptApp.getService().getUrl();
  } catch (error) {
    return 'https://script.google.com';
  }
}

/**
 * 回答データ取得
 * @param {string} userId - ユーザーID
 * @param {string} sheetName - シート名
 * @returns {Object} 回答データ
 */
function getResponsesData(userId, sheetName) {
  try {
    const data = getSheetData(userId, sheetName, null, 'newest', false);
    return {
      status: 'success',
      data: data
    };
  } catch (error) {
    logError(error, 'getResponsesData', ERROR_SEVERITY.LOW, ERROR_CATEGORIES.DATABASE);
    return {
      status: 'error',
      data: []
    };
  }
}

/**
 * シート詳細取得
 * @param {string} userId - ユーザーID
 * @param {string} sheetName - シート名
 * @returns {Object} シート詳細
 */
function getSheetDetails(userId, sheetName) {
  try {
    const userInfo = getUserInfo(userId);
    if (!userInfo || !userInfo.spreadsheetId) {
      return null;
    }
    
    const spreadsheet = Spreadsheet.openById(userInfo.spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      return null;
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowCount = sheet.getLastRow();
    
    return {
      name: sheetName,
      headers: headers,
      rowCount: rowCount,
      columnCount: headers.length
    };
  } catch (error) {
    logError(error, 'getSheetDetails', ERROR_SEVERITY.LOW, ERROR_CATEGORIES.DATABASE);
    return null;
  }
}

/**
 * ユーザーアクセス検証
 * @param {string} userId - ユーザーID
 * @returns {boolean} アクセス可能か
 */
function verifyUserAccess(userId) {
  const userInfo = getUserInfo(userId);
  if (!userInfo || userInfo.isActive !== 'true') {
    throw new Error('アクセスが拒否されました');
  }
  return true;
}

/**
 * シート設定を保存（新実装）
 * @param {string} userId - ユーザーID
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @param {Object} config - 設定
 * @param {Object} options - オプション
 * @returns {boolean} 成功フラグ
 */
function saveSheetConfig(userId, spreadsheetId, sheetName, config, options = {}) {
  try {
    const userInfo = getUserInfo(userId);
    if (!userInfo) {
      throw new Error('ユーザーが見つかりません');
    }
    
    const configJson = JSON.parse(userInfo.configJson || '{}');
    
    // シート設定を更新
    configJson[`sheet_${sheetName}`] = config;
    
    // 公開設定の更新
    if (options.setAsPublished) {
      configJson.publishedSheetName = sheetName;
      configJson.publishedSpreadsheetId = spreadsheetId;
      configJson.appPublished = true;
    }
    
    // データベースに保存
    updateUser(userId, {
      configJson: JSON.stringify(configJson)
    });
    
    // キャッシュクリア
    clearUserCache(userId);
    
    return true;
  } catch (error) {
    logError(error, 'saveSheetConfig', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.DATABASE);
    return false;
  }
}

/**
 * クイックスタートフォーム作成（UI経由）
 * @param {string} requestUserId - ユーザーID
 * @returns {Object} 作成結果
 */
function createQuickStartFormUI(requestUserId) {
  try {
    const userInfo = getUserInfo(requestUserId);
    if (!userInfo) {
      throw new Error('ユーザーが見つかりません');
    }
    
    // フォーム作成
    const result = createQuickStartForm(userInfo.adminEmail, requestUserId);
    
    // 設定を更新
    const configJson = JSON.parse(userInfo.configJson || '{}');
    configJson.formCreated = true;
    configJson.formUrl = result.formUrl;
    configJson.setupStatus = 'form_created';
    
    // ユーザー情報を更新
    updateUser(requestUserId, {
      spreadsheetId: result.spreadsheetId,
      spreadsheetUrl: result.spreadsheetUrl,
      configJson: JSON.stringify(configJson)
    });
    
    return {
      success: true,
      formUrl: result.formUrl,
      spreadsheetUrl: result.spreadsheetUrl,
      setupStep: 3
    };
  } catch (error) {
    logError(error, 'createQuickStartFormUI', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.SYSTEM);
    return {
      success: false,
      error: error.message
    };
  }
}