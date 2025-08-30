/**
 * @fileoverview AdminPanel バックエンド関数
 * 既存のシステムと統合する最小限の実装
 */

/**
 * 管理パネル用のサーバーサイド関数群
 * 既存のunifiedUserManager、database、Core関数を活用
 */

// =============================================================================
// データソース管理
// =============================================================================

/**
 * スプレッドシート一覧を取得
 * @returns {Array<Object>} スプレッドシート情報の配列
 */
function getSpreadsheetList() {
  try {
    console.log('getSpreadsheetList: スプレッドシート一覧取得開始');
    
    // Google Driveからスプレッドシートを検索
    const files = DriveApp.searchFiles(
      'mimeType="application/vnd.google-apps.spreadsheet" and trashed=false'
    );
    
    const spreadsheets = [];
    let count = 0;
    const maxResults = 50; // パフォーマンス制限
    
    while (files.hasNext() && count < maxResults) {
      const file = files.next();
      spreadsheets.push({
        id: file.getId(),
        name: file.getName(),
        lastModified: file.getLastUpdated().toISOString(),
        owner: file.getOwner().getEmail()
      });
      count++;
    }
    
    // 最終更新順でソート
    spreadsheets.sort((a, b) => new Date(b.lastModified) - new Date(a.lastModified));
    
    console.log(`getSpreadsheetList: ${spreadsheets.length}個のスプレッドシートを取得`);
    return spreadsheets;
    
  } catch (error) {
    console.error('getSpreadsheetList エラー:', error);
    throw new Error('スプレッドシート一覧の取得に失敗しました: ' + error.message);
  }
}

/**
 * 指定されたスプレッドシートのシート一覧を取得
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {Array<Object>} シート情報の配列
 */
function getSheetList(spreadsheetId) {
  try {
    console.log('getSheetList: シート一覧取得開始', spreadsheetId);
    
    if (!spreadsheetId) {
      throw new Error('スプレッドシートIDが指定されていません');
    }
    
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheets = spreadsheet.getSheets();
    
    const sheetList = sheets.map(sheet => ({
      name: sheet.getName(),
      index: sheet.getIndex(),
      rowCount: sheet.getLastRow(),
      columnCount: sheet.getLastColumn(),
      hidden: sheet.isSheetHidden()
    }));
    
    console.log(`getSheetList: ${sheetList.length}個のシートを取得`);
    return sheetList;
    
  } catch (error) {
    console.error('getSheetList エラー:', error);
    throw new Error('シート一覧の取得に失敗しました: ' + error.message);
  }
}

/**
 * データソースに接続し、列マッピングを自動検出
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} 接続結果と列マッピング情報
 */
function connectToDataSource(spreadsheetId, sheetName) {
  try {
    console.log('connectToDataSource: データソース接続開始', { spreadsheetId, sheetName });
    
    if (!spreadsheetId || !sheetName) {
      throw new Error('スプレッドシートIDとシート名が必要です');
    }
    
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error(`シート "${sheetName}" が見つかりません`);
    }
    
    // ヘッダー行を取得（最初の行を仮定）
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // 列マッピングを自動検出
    const columnMapping = detectColumnMapping(headerRow);
    
    // 設定を保存（既存のユーザー管理システムを活用）
    const currentUser = coreGetCurrentUserEmail();
    const userInfo = DB.findUserByEmail(currentUser);
    
    if (userInfo) {
      // ユーザーのスプレッドシート設定を更新
      updateUserSpreadsheetConfig(userInfo.userId, {
        spreadsheetId: spreadsheetId,
        sheetName: sheetName,
        columnMapping: columnMapping,
        lastConnected: new Date().toISOString()
      });
    }
    
    console.log('connectToDataSource: 接続成功', columnMapping);
    return {
      success: true,
      columnMapping: columnMapping,
      rowCount: sheet.getLastRow(),
      message: 'データソースに正常に接続されました'
    };
    
  } catch (error) {
    console.error('connectToDataSource エラー:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * ヘッダー行から列マッピングを自動検出
 * @param {Array<string>} headers - ヘッダー行の配列
 * @returns {Object} 検出された列マッピング
 */
function detectColumnMapping(headers) {
  const mapping = {
    question: null,
    answerer: null,
    reason: null,
    confidence: {}
  };
  
  // キーワードパターン
  const patterns = {
    question: ['質問', 'question', 'query', 'ask', 'Q'],
    answerer: ['回答者', '名前', 'name', 'answerer', 'user', 'ユーザー'],
    reason: ['理由', 'reason', 'comment', 'コメント', '備考', 'note']
  };
  
  headers.forEach((header, index) => {
    const headerLower = header.toString().toLowerCase();
    
    Object.keys(patterns).forEach(type => {
      const pattern = patterns[type];
      const matchCount = pattern.filter(p => 
        headerLower.includes(p.toLowerCase())
      ).length;
      
      if (matchCount > 0) {
        const confidence = Math.min(95, 60 + (matchCount * 15));
        if (!mapping[type] || confidence > (mapping.confidence[type] || 0)) {
          mapping[type] = index;
          mapping.confidence[type] = confidence;
        }
      }
    });
  });
  
  return mapping;
}

// =============================================================================
// システム監視・メトリクス
// =============================================================================


// =============================================================================
// 設定管理
// =============================================================================

/**
 * 現在の設定情報を取得
 * @returns {Object} 現在の設定情報
 */
function getCurrentConfig() {
  try {
    console.log('getCurrentConfig: 設定情報取得開始');
    
    const currentUser = coreGetCurrentUserEmail();
    const userInfo = DB.findUserByEmail(currentUser);
    
    if (!userInfo) {
      return {
        setupStatus: 'pending',
        spreadsheetId: null,
        sheetName: null,
        formCreated: false,
        appPublished: false,
        lastUpdated: null,
        user: currentUser
      };
    }
    
    // 既存のdetermineSetupStep関数を活用
    const configJson = getUserConfigJson(userInfo.userId);
    const setupStep = determineSetupStep(userInfo, configJson);
    
    const config = {
      setupStatus: getSetupStatusFromStep(setupStep),
      spreadsheetId: userInfo.spreadsheetId,
      sheetName: userInfo.sheetName,
      formCreated: configJson ? configJson.formCreated : false,
      appPublished: configJson ? configJson.appPublished : false,
      lastUpdated: userInfo.lastUpdated,
      setupStep: setupStep,
      user: currentUser,
      userId: userInfo.userId
    };
    
    console.log('getCurrentConfig: 設定情報取得完了', config);
    return config;
    
  } catch (error) {
    console.error('getCurrentConfig エラー:', error);
    return {
      setupStatus: 'error',
      error: error.message,
      lastUpdated: new Date().toISOString()
    };
  }
}

/**
 * セットアップステップからステータス文字列を取得
 * @param {number} step - セットアップステップ
 * @returns {string} ステータス文字列
 */
function getSetupStatusFromStep(step) {
  switch (step) {
    case 1: return 'pending';
    case 2: return 'configuring';
    case 3: return 'completed';
    default: return 'unknown';
  }
}

/**
 * 下書き設定を保存
 * @param {Object} config - 保存する設定
 * @returns {Object} 保存結果
 */
function saveDraftConfiguration(config) {
  try {
    console.log('saveDraftConfiguration: 下書き保存開始', config);
    
    const currentUser = coreGetCurrentUserEmail();
    let userInfo = DB.findUserByEmail(currentUser);
    
    if (!userInfo) {
      // 新規ユーザーの場合は作成（既存のシステムを活用）
      throw new Error('ユーザー情報が見つかりません。先にユーザー登録を行ってください。');
    }
    
    // 設定を更新
    const updateResult = updateUserSpreadsheetConfig(userInfo.userId, {
      spreadsheetId: config.spreadsheetId,
      sheetName: config.sheetName,
      displaySettings: {
        showNames: config.showNames,
        showReactions: config.showReactions
      },
      lastDraftSave: new Date().toISOString()
    });
    
    if (updateResult.success) {
      const updatedConfig = getCurrentConfig();
      return {
        success: true,
        config: updatedConfig,
        message: '下書きが保存されました'
      };
    } else {
      throw new Error(updateResult.error);
    }
    
  } catch (error) {
    console.error('saveDraftConfiguration エラー:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * アプリケーションを公開
 * @param {Object} config - 公開する設定
 * @returns {Object} 公開結果
 */
function publishApplication(config) {
  try {
    console.log('publishApplication: アプリ公開開始', config);
    
    const currentUser = coreGetCurrentUserEmail();
    const userInfo = DB.findUserByEmail(currentUser);
    
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    
    if (!config.spreadsheetId || !config.sheetName) {
      throw new Error('データソースが設定されていません');
    }
    
    // 既存の公開システムを活用（簡略化）
    const publishResult = executeAppPublish(userInfo.userId, {
      appName: config.appName,
      spreadsheetId: config.spreadsheetId,
      sheetName: config.sheetName,
      displaySettings: {
        showNames: config.showNames,
        showReactions: config.showReactions
      }
    });
    
    if (publishResult.success) {
      const updatedConfig = getCurrentConfig();
      return {
        success: true,
        config: updatedConfig,
        appUrl: publishResult.appUrl,
        message: 'アプリケーションが正常に公開されました'
      };
    } else {
      throw new Error(publishResult.error);
    }
    
  } catch (error) {
    console.error('publishApplication エラー:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

// =============================================================================
// ヘルパー関数
// =============================================================================

/**
 * ユーザーのスプレッドシート設定を更新
 * @param {string} userId - ユーザーID
 * @param {Object} config - 更新する設定
 * @returns {Object} 更新結果
 */
function updateUserSpreadsheetConfig(userId, config) {
  try {
    const props = PropertiesService.getScriptProperties();
    const configKey = `user_config_${userId}`;
    
    // 既存の設定を取得
    let existingConfig = {};
    try {
      const existingConfigStr = props.getProperty(configKey);
      if (existingConfigStr) {
        existingConfig = JSON.parse(existingConfigStr);
      }
    } catch (parseError) {
      console.warn('既存設定の解析エラー:', parseError);
    }
    
    // 設定をマージ
    const mergedConfig = {
      ...existingConfig,
      ...config,
      lastUpdated: new Date().toISOString()
    };
    
    // 設定を保存
    props.setProperty(configKey, JSON.stringify(mergedConfig));
    
    console.log('updateUserSpreadsheetConfig: 設定更新完了', { userId, config: mergedConfig });
    return {
      success: true,
      config: mergedConfig
    };
    
  } catch (error) {
    console.error('updateUserSpreadsheetConfig エラー:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * ユーザーの設定JSONを取得
 * @param {string} userId - ユーザーID
 * @returns {Object|null} 設定JSON
 */
function getUserConfigJson(userId) {
  try {
    const props = PropertiesService.getScriptProperties();
    const configKey = `user_config_${userId}`;
    const configStr = props.getProperty(configKey);
    
    if (!configStr) {
      return null;
    }
    
    return JSON.parse(configStr);
    
  } catch (error) {
    console.error('getUserConfigJson エラー:', error);
    return null;
  }
}

/**
 * アプリ公開を実行（簡略化実装）
 * @param {string} userId - ユーザーID
 * @param {Object} publishConfig - 公開設定
 * @returns {Object} 公開結果
 */
function executeAppPublish(userId, publishConfig) {
  try {
    // 既存のWebアプリURLを生成または取得
    const webAppUrl = getOrCreateWebAppUrl(userId, publishConfig.appName);
    
    // 公開状態をPropertiesServiceに記録
    const props = PropertiesService.getScriptProperties();
    props.setProperty(`${userId}_published`, 'true');
    props.setProperty(`${userId}_app_url`, webAppUrl);
    props.setProperty(`${userId}_publish_date`, new Date().toISOString());
    
    // ユーザー設定に公開情報を追加
    updateUserSpreadsheetConfig(userId, {
      appPublished: true,
      appName: publishConfig.appName,
      appUrl: webAppUrl,
      publishedAt: new Date().toISOString()
    });
    
    return {
      success: true,
      appUrl: webAppUrl,
      publishedAt: new Date().toISOString()
    };
    
  } catch (error) {
    console.error('executeAppPublish エラー:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * WebアプリURLを取得または作成
 * @param {string} userId - ユーザーID
 * @param {string} appName - アプリ名
 * @returns {string} WebアプリURL
 */
function getOrCreateWebAppUrl(userId, appName) {
  try {
    // 既存のWebアプリURLがあるかチェック
    const props = PropertiesService.getScriptProperties();
    const existingUrl = props.getProperty(`${userId}_app_url`);
    
    if (existingUrl) {
      return existingUrl;
    }
    
    // 新規WebアプリURLを生成（実際の実装ではScriptAppを使用）
    const scriptId = ScriptApp.getScriptId();
    const webAppUrl = `https://script.google.com/macros/s/${scriptId}/exec?userId=${userId}`;
    
    return webAppUrl;
    
  } catch (error) {
    console.error('getOrCreateWebAppUrl エラー:', error);
    // フォールバック
    return `https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec?userId=${userId}`;
  }
}

/**
 * 管理パネルのHTMLを返す
 * @returns {HtmlOutput} HTML出力
 */
function doGet() {
  try {
    return HtmlService.createTemplateFromFile('AdminPanel')
      .evaluate()
      .setTitle('StudyQuest 管理パネル')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    console.error('doGet エラー:', error);
    return HtmlService.createHtmlOutput(`
      <h1>エラー</h1>
      <p>管理パネルの読み込みに失敗しました: ${error.message}</p>
    `);
  }
}

/**
 * 現在のアクティブボード情報とURL生成
 * @returns {Object} ボード情報とURL
 */
function getCurrentBoardInfoAndUrls() {
  try {
    console.log('getCurrentBoardInfoAndUrls: ボード情報取得開始');
    
    // 現在ログイン中のユーザー情報を取得
    const userInfo = getCurrentUserInfo();
    if (!userInfo || !userInfo.userId) {
      console.warn('getCurrentBoardInfoAndUrls: ユーザー情報が見つかりません');
      return {
        isActive: false,
        questionText: 'アクティブなボードがありません',
        error: 'ユーザー情報が見つかりません'
      };
    }

    console.log('getCurrentBoardInfoAndUrls: ユーザー情報取得成功', {
      userId: userInfo.userId,
      hasSpreadsheetId: !!userInfo.spreadsheetId
    });

    // 現在のボードデータを取得
    let boardData = null;
    let questionText = '問題読み込み中...';
    
    if (userInfo.spreadsheetId) {
      try {
        boardData = getSheetData(userInfo.userId);
        questionText = boardData?.header || '問題文が設定されていません';
        console.log('getCurrentBoardInfoAndUrls: ボードデータ取得成功', {
          hasHeader: !!boardData?.header
        });
      } catch (error) {
        console.warn('getCurrentBoardInfoAndUrls: ボードデータ取得エラー:', error.message);
        questionText = '問題文の読み込みに失敗しました';
      }
    } else {
      questionText = 'スプレッドシートが設定されていません';
    }

    // URL生成
    const baseUrl = getWebAppUrlCached();
    if (!baseUrl) {
      console.error('getCurrentBoardInfoAndUrls: WebAppURL取得失敗');
      return {
        isActive: false,
        questionText: questionText,
        error: 'URLの生成に失敗しました'
      };
    }

    const viewUrl = `${baseUrl}?mode=view&userId=${encodeURIComponent(userInfo.userId)}`;
    const adminUrl = `${baseUrl}?mode=admin&userId=${encodeURIComponent(userInfo.userId)}`;

    const result = {
      isActive: !!userInfo.spreadsheetId,
      questionText: questionText,
      sheetName: userInfo.sheetName || 'シート名未設定',
      urls: {
        view: viewUrl,      // 閲覧者向け（共有用）
        admin: adminUrl     // 管理者向け
      },
      lastUpdated: new Date().toLocaleString('ja-JP'),
      ownerEmail: userInfo.adminEmail
    };

    console.log('getCurrentBoardInfoAndUrls: 成功', {
      isActive: result.isActive,
      hasQuestionText: !!result.questionText,
      viewUrl: result.urls.view
    });

    return result;

  } catch (error) {
    console.error('getCurrentBoardInfoAndUrls エラー:', error);
    return {
      isActive: false,
      questionText: 'エラーが発生しました',
      error: error.message
    };
  }
}

/**
 * システム管理者かどうかをチェック
 * @returns {boolean} システム管理者の場合はtrue
 */
function checkIsSystemAdmin() {
  try {
    console.log('checkIsSystemAdmin: システム管理者チェック開始');
    
    // 既存のisDeployUser関数を利用
    const isAdmin = Deploy.isUser();
    
    console.log('checkIsSystemAdmin: 結果', isAdmin);
    return isAdmin;
    
  } catch (error) {
    console.error('checkIsSystemAdmin エラー:', error);
    return false;
  }
}

console.log('AdminPanel.gs 読み込み完了');