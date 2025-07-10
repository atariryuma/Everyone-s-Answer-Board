/**
 * @fileoverview StudyQuest - Core Functions (最適化版)
 * 主要な業務ロジックとAPI エンドポイント
 */

// =================================================================
// メインロジック
// =================================================================

// doGetLegacy function removed - consolidated into main doGet in UltraOptimizedCore.gs

/**
 * 意見ヘッダーを安全に取得する関数（テンプレート変数の問題を回避）
 * @param {string} userId - ユーザーID
 * @param {string} sheetName - シート名
 * @returns {string} 意見ヘッダー
 */
function getOpinionHeaderSafely(userId, sheetName) {
  try {
    const userInfo = findUserById(userId);
    if (!userInfo) {
      return 'お題';
    }
    
    const config = JSON.parse(userInfo.configJson || '{}');
    const sheetConfigKey = 'sheet_' + (config.publishedSheetName || sheetName);
    const sheetConfig = config[sheetConfigKey] || {};
    
    const opinionHeader = sheetConfig.opinionHeader || config.publishedSheetName || 'お題';
    
    console.log('getOpinionHeaderSafely:', {
      userId: userId,
      sheetName: sheetName,
      opinionHeader: opinionHeader
    });
    
    return opinionHeader;
  } catch (e) {
    console.error('getOpinionHeaderSafely error:', e);
    return 'お題';
  }
}

/**
 * 新規ユーザーを登録する（データベース登録のみ）
 * フォーム作成はクイックスタートで実行される
 */
function registerNewUser(adminEmail) {
  var activeUser = Session.getActiveUser();
  if (adminEmail !== activeUser.getEmail()) {
    throw new Error('認証エラー: 操作を実行しているユーザーとメールアドレスが一致しません。');
  }

  // 既存ユーザーチェック（1ユーザー1行の原則）
  var existingUser = findUserByEmail(adminEmail);
  var userId, appUrls;
  
  if (existingUser) {
    // 既存ユーザーの場合は情報を更新
    userId = existingUser.userId;
    var existingConfig = JSON.parse(existingUser.configJson || '{}');
    
    // 設定をリセット（新規登録状態に戻す）
    var updatedConfig = {
      ...existingConfig,
      setupStatus: 'pending',
      lastRegistration: new Date().toISOString(),
      formCreated: false,
      appPublished: false
    };
    
    // 既存ユーザー情報を更新
    updateUser(userId, {
      lastAccessedAt: new Date().toISOString(),
      isActive: 'true',
      configJson: JSON.stringify(updatedConfig)
    });
    // キャッシュを無効化して最新状態を反映
    invalidateUserCache(userId, adminEmail, existingUser.spreadsheetId);
    
    debugLog('✅ 既存ユーザー情報を更新しました: ' + adminEmail);
    appUrls = generateAppUrls(userId);
    
    return {
      userId: userId,
      adminUrl: appUrls.adminUrl,
      viewUrl: appUrls.viewUrl,
      setupRequired: true,
      message: '既存ユーザーの情報を更新しました。クイックスタートで新しいフォームを作成してください。',
      isExistingUser: true
    };
  }

  // 新規ユーザーの場合
  userId = Utilities.getUuid();
  
  var initialConfig = {
    setupStatus: 'pending',
    createdAt: new Date().toISOString(),
    formCreated: false,
    appPublished: false
  };
  
  var userData = {
    userId: userId,
    adminEmail: adminEmail,
    spreadsheetId: '',
    spreadsheetUrl: '',
    createdAt: new Date().toISOString(),
    configJson: JSON.stringify(initialConfig),
    lastAccessedAt: new Date().toISOString(),
    isActive: 'true'
  };

  try {
    createUser(userData);
    debugLog('✅ データベースに新規ユーザーを登録しました: ' + adminEmail);
    // 生成されたユーザー情報のキャッシュをクリア
    invalidateUserCache(userId, adminEmail);
  } catch (e) {
    console.error('データベースへのユーザー登録に失敗: ' + e.message);
    throw new Error('ユーザー登録に失敗しました。システム管理者に連絡してください。');
  }

  // 成功レスポンスを返す
  appUrls = generateAppUrls(userId);
  return {
    userId: userId,
    adminUrl: appUrls.adminUrl,
    viewUrl: appUrls.viewUrl,
    setupRequired: true,
    message: 'ユーザー登録が完了しました！次にクイックスタートでフォームを作成してください。',
    isExistingUser: false
  };
}

/**
 * リアクションを追加/削除する
 * Page.htmlから呼び出される - フロントエンド期待形式に対応
 */
function addReaction(rowIndex, reactionKey, sheetName) {
  try {
    var reactingUserEmail = Session.getActiveUser().getEmail();
    var props = PropertiesService.getUserProperties();
    var ownerUserId = props.getProperty('CURRENT_USER_ID');

    if (!ownerUserId) {
      throw new Error('ボードのオーナー情報が見つかりません。');
    }

    // ボードオーナーの情報をDBから取得（キャッシュ利用）
    var boardOwnerInfo = findUserById(ownerUserId);
    if (!boardOwnerInfo) {
      throw new Error('無効なボードです。');
    }

    var result = processReaction(
      boardOwnerInfo.spreadsheetId,
      sheetName,
      rowIndex,
      reactionKey,
      reactingUserEmail
    );
    
    // Page.html期待形式に変換
    if (result && result.status === 'success') {
      // 更新後のリアクション情報を取得
      var updatedReactions = getRowReactions(boardOwnerInfo.spreadsheetId, sheetName, rowIndex, reactingUserEmail);
      
      return {
        status: "ok",
        reactions: updatedReactions
      };
    } else {
      throw new Error(result.message || 'リアクションの処理に失敗しました');
    }
  } catch (e) {
    console.error('addReaction エラー: ' + e.message);
    return {
      status: "error",
      message: e.message
    };
  }
}

// =================================================================
// データ取得関数
// =================================================================

/**
 * 公開されたシートのデータを取得
 * Page.htmlから呼び出される - フロントエンド期待形式に対応
 */
function getPublishedSheetData(classFilter, sortOrder) {
  try {
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    debugLog('getPublishedSheetData: userId=%s, classFilter=%s, sortOrder=%s', currentUserId, classFilter, sortOrder);
    
    if (!currentUserId) {
      throw new Error('ユーザーコンテキストが設定されていません');
    }
    
    var userInfo = getUserWithFallback(currentUserId);
    if (!userInfo) {
      handleMissingUser(currentUserId);
      // Fallback to active user email if property points to missing user
      var fallbackEmail = Session.getActiveUser().getEmail();
      var altUser = findUserByEmail(fallbackEmail);
      if (altUser) {
        currentUserId = altUser.userId;
        props.setProperty('CURRENT_USER_ID', currentUserId);
        userInfo = altUser;
      } else {
        throw new Error('ユーザー情報が見つかりません');
      }
    }
    debugLog('getPublishedSheetData: userInfo=%s', JSON.stringify(userInfo));
    
    var configJson = JSON.parse(userInfo.configJson || '{}');
    debugLog('getPublishedSheetData: configJson=%s', JSON.stringify(configJson));

    // 公開対象のスプレッドシートIDとシート名を取得
    var publishedSpreadsheetId = configJson.publishedSpreadsheetId;
    var publishedSheetName = configJson.publishedSheetName;

    if (!publishedSpreadsheetId || !publishedSheetName) {
      throw new Error('公開対象のスプレッドシートまたはシートが設定されていません。');
    }

    // シート固有の設定を取得 (sheetKey is based only on sheet name)
    var sheetKey = 'sheet_' + publishedSheetName;
    var sheetConfig = configJson[sheetKey] || {};
    debugLog('getPublishedSheetData: sheetConfig=%s', JSON.stringify(sheetConfig));
    
    // データ取得
    var sheetData = getSheetData(currentUserId, publishedSheetName, classFilter, sortOrder);
    debugLog('getPublishedSheetData: sheetData status=%s, totalCount=%s', sheetData.status, sheetData.totalCount);

    if (sheetData.status === 'error') {
      throw new Error(sheetData.message);
    }
    
    // Page.html期待形式に変換
    // 設定からヘッダー名を取得。未定義の場合のみデフォルト値を使用。
    var mainHeaderName = sheetConfig.opinionHeader || COLUMN_HEADERS.OPINION;
    var reasonHeaderName = sheetConfig.reasonHeader || COLUMN_HEADERS.REASON;
    var classHeaderName = sheetConfig.classHeader !== undefined ? sheetConfig.classHeader : COLUMN_HEADERS.CLASS;
    var nameHeaderName = sheetConfig.nameHeader !== undefined ? sheetConfig.nameHeader : COLUMN_HEADERS.NAME;
    debugLog('getPublishedSheetData: Configured Headers - mainHeaderName=%s, reasonHeaderName=%s, classHeaderName=%s, nameHeaderName=%s', mainHeaderName, reasonHeaderName, classHeaderName, nameHeaderName);

    // ヘッダーインデックスマップを取得（キャッシュされた実際のマッピング）
    var headerIndices = getHeaderIndices(publishedSpreadsheetId, publishedSheetName);
    debugLog('getPublishedSheetData: Available headerIndices=%s', JSON.stringify(headerIndices));
    
    // 動的列名のマッピング: 設定された名前と実際のヘッダーを照合
    var mappedIndices = mapConfigToActualHeaders({
      opinionHeader: mainHeaderName,
      reasonHeader: reasonHeaderName, 
      classHeader: classHeaderName,
      nameHeader: nameHeaderName
    }, headerIndices);
    debugLog('getPublishedSheetData: Mapped indices=%s', JSON.stringify(mappedIndices));

    var formattedData = sheetData.data.map(function(row, index) {
      // マッピングされたインデックスを使用してデータを取得
      var classIndex = mappedIndices.classHeader;
      var opinionIndex = mappedIndices.opinionHeader;
      var reasonIndex = mappedIndices.reasonHeader;
      var nameIndex = mappedIndices.nameHeader;

      debugLog('getPublishedSheetData: Row %s - classIndex=%s, opinionIndex=%s, reasonIndex=%s, nameIndex=%s', index, classIndex, opinionIndex, reasonIndex, nameIndex);
      debugLog('getPublishedSheetData: Row data length=%s, first few values=%s', row.originalData ? row.originalData.length : 'undefined', row.originalData ? JSON.stringify(row.originalData.slice(0, 5)) : 'undefined');
      
      return {
        rowIndex: row.rowNumber || (index + 2), // 実際の行番号
        name: (sheetData.displayMode === DISPLAY_MODES.NAMED && nameIndex !== undefined && row.originalData && row.originalData[nameIndex]) ? row.originalData[nameIndex] : '',
        class: (classIndex !== undefined && row.originalData && row.originalData[classIndex]) ? row.originalData[classIndex] : '',
        opinion: (opinionIndex !== undefined && row.originalData && row.originalData[opinionIndex]) ? row.originalData[opinionIndex] : '',
        reason: (reasonIndex !== undefined && row.originalData && row.originalData[reasonIndex]) ? row.originalData[reasonIndex] : '',
        reactions: {
          UNDERSTAND: { count: row.understandCount || 0, reacted: false },
          LIKE: { count: row.likeCount || 0, reacted: false },
          CURIOUS: { count: row.curiousCount || 0, reacted: false }
        },
        highlight: row.isHighlighted || false
      };
    });
    debugLog('getPublishedSheetData: formattedData length=%s', formattedData.length);

    // ★★★ここからが修正箇所★★★

    // ボードのタイトルを、保存された設定（opinionHeader）から取得する
    // もし設定がなければ、デフォルトでシート名を使う
    const headerTitle = sheetConfig.opinionHeader || publishedSheetName;
    
    // ...（データ取得とフォーマット処理は変更なし）...

    // 最終的に返すオブジェクトの header プロパティに、取得したタイトルを設定
    var result = {
      header: headerTitle,
      sheetName: publishedSheetName, // targetSheetからpublishedSheetNameに変更
      showCounts: configJson.showCounts !== false,
      displayMode: sheetData.displayMode || DISPLAY_MODES.ANONYMOUS,
      data: formattedData,
      rows: formattedData // 後方互換性のため
    };
    debugLog('getPublishedSheetData: Returning result=%s', JSON.stringify(result));
    return result;
    
  } catch (e) {
    console.error('公開シートデータ取得エラー: ' + e.message);
    return {
      status: 'error',
      message: 'データの取得に失敗しました: ' + e.message,
      data: [],
      rows: []
    };
  }
}

/**
 * アプリ設定を取得（最適化版）
 */
function getAppConfig() {
  try {
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    
    if (!currentUserId) {
      // コンテキストが設定されていない場合、現在のユーザーで検索
      var activeUser = Session.getActiveUser().getEmail();
      var userInfo = findUserByEmail(activeUser);
      if (userInfo) {
        currentUserId = userInfo.userId;
        props.setProperty('CURRENT_USER_ID', currentUserId);
      } else {
        throw new Error('ユーザー情報が見つかりません');
      }
    }
    
    var userInfo = getUserWithFallback(currentUserId);
    if (!userInfo) {
      handleMissingUser(currentUserId);
      var fallbackEmail = Session.getActiveUser().getEmail();
      var altUser = findUserByEmail(fallbackEmail);
      if (altUser) {
        currentUserId = altUser.userId;
        props.setProperty('CURRENT_USER_ID', currentUserId);
        userInfo = altUser;
      } else {
        throw new Error('ユーザー情報が見つかりません');
      }
    }
    
    var configJson = JSON.parse(userInfo.configJson || '{}');

    // --- Auto-healing for inconsistent setup states ---
    var needsUpdate = false;
    if (configJson.formUrl && !configJson.formCreated) {
      configJson.formCreated = true;
      needsUpdate = true;
    }
    if (configJson.formCreated && configJson.setupStatus !== 'completed') {
      configJson.setupStatus = 'completed';
      needsUpdate = true;
    }
    if (configJson.publishedSheetName && !configJson.appPublished) {
      configJson.appPublished = true;
      needsUpdate = true;
    }
    if (needsUpdate) {
      try {
        updateUser(currentUserId, { configJson: JSON.stringify(configJson) });
      } catch (updateErr) {
        console.warn('Config auto-heal failed: ' + updateErr.message);
      }
    }

    var sheets = getSheetsList(currentUserId);
    var appUrls = generateAppUrls(currentUserId);
    
    // 回答数を取得
    var answerCount = 0;
    var totalReactions = 0;
    try {
      if (configJson.publishedSpreadsheetId && configJson.publishedSheetName) {
        var responseData = getResponsesData(currentUserId, configJson.publishedSheetName);
        if (responseData.status === 'success') {
          answerCount = responseData.data.length;
          // リアクション数の概算計算（詳細実装は後回し）
          totalReactions = answerCount * 2; // 暫定値
        }
      }
    } catch (countError) {
      console.warn('回答数の取得に失敗: ' + countError.message);
    }
    
    return {
      status: 'success',
      userId: currentUserId,
      adminEmail: userInfo.adminEmail,
      publishedSpreadsheetId: configJson.publishedSpreadsheetId || '',
      publishedSheetName: configJson.publishedSheetName || '',
      displayMode: configJson.displayMode || DISPLAY_MODES.ANONYMOUS,
      isPublished: configJson.appPublished || false,
      availableSheets: sheets,
      allSheets: sheets, // AdminPanel.htmlで使用される
      spreadsheetUrl: userInfo.spreadsheetUrl,
      formUrl: configJson.formUrl || '',
      editFormUrl: configJson.editFormUrl || '',
      webAppUrl: appUrls.webAppUrl,
      adminUrl: appUrls.adminUrl,
      viewUrl: appUrls.viewUrl,
      activeSheetName: configJson.publishedSheetName || '',
      appUrls: appUrls,
      // データベース詳細情報
      userInfo: {
        userId: currentUserId,
        adminEmail: userInfo.adminEmail,
        spreadsheetId: userInfo.spreadsheetId || '',
        spreadsheetUrl: userInfo.spreadsheetUrl || '',
        createdAt: userInfo.createdAt || '',
        lastAccessedAt: userInfo.lastAccessedAt || '',
        isActive: userInfo.isActive || 'false',
        configJson: userInfo.configJson || '{}'
      },
      // 統計情報
      answerCount: answerCount,
      totalReactions: totalReactions,
      // システム状態
      systemStatus: {
        setupStatus: configJson.setupStatus || 'unknown',
        formCreated: configJson.formCreated || false,
        appPublished: configJson.appPublished || false,
        lastUpdated: new Date().toISOString()
      },
      // ユーザーのセットアップ段階を判定
      setupStep: (userInfo && userInfo.spreadsheetId) ? (configJson.publishedSheetName ? 3 : 2) : 1
    };
  } catch (e) {
    console.error('アプリ設定取得エラー: ' + e.message);
    return {
      status: 'error',
      message: '設定の取得に失敗しました: ' + e.message
    };
  }
}

/**
 * シート設定を保存する
 * AdminPanel.htmlから呼び出される
 * @param {string} spreadsheetId - 設定対象のスプレッドシートID
 * @param {string} sheetName - 設定対象のシート名
 * @param {object} config - 保存するシート固有の設定
 */
function saveSheetConfig(spreadsheetId, sheetName, config) {
  try {
    if (!spreadsheetId || typeof spreadsheetId !== 'string') {
      throw new Error('無効なspreadsheetIdです: ' + spreadsheetId);
    }
    if (!sheetName || typeof sheetName !== 'string') {
      throw new Error('無効なsheetNameです: ' + sheetName);
    }
    if (!config || typeof config !== 'object') {
      throw new Error('無効なconfigです: ' + config);
    }
    
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    if (!currentUserId) {
      throw new Error('ユーザーコンテキストが設定されていません');
    }

    var userInfo = getUserWithFallback(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    var configJson = JSON.parse(userInfo.configJson || '{}');

    // save using sheet-specific key expected by getConfig
    var sheetKey = 'sheet_' + sheetName;
    configJson[sheetKey] = config;

    updateUser(currentUserId, { configJson: JSON.stringify(configJson) });
    debugLog('✅ シート設定を保存しました: %s', sheetKey);
    return { status: 'success', message: 'シート設定を保存しました。' };
  } catch (e) {
    console.error('シート設定保存エラー: ' + e.message);
    return { status: 'error', message: 'シート設定の保存に失敗しました: ' + e.message };
  }
}

/**
 * 表示するシートを切り替える
 * AdminPanel.htmlから呼び出される
 * @param {string} spreadsheetId - 公開対象のスプレッドシートID
 * @param {string} sheetName - 公開対象のシート名
 */
function switchToSheet(spreadsheetId, sheetName) {
  try {
    if (!spreadsheetId || typeof spreadsheetId !== 'string') {
      throw new Error('無効なspreadsheetIdです: ' + spreadsheetId);
    }
    if (!sheetName || typeof sheetName !== 'string') {
      throw new Error('無効なsheetNameです: ' + sheetName);
    }
    
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    if (!currentUserId) {
      throw new Error('ユーザーコンテキストが設定されていません');
    }

    var userInfo = getUserWithFallback(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    var configJson = JSON.parse(userInfo.configJson || '{}');

    configJson.publishedSpreadsheetId = spreadsheetId;
    configJson.publishedSheetName = sheetName;
    configJson.appPublished = true; // シートを切り替えたら公開状態にする

    updateUser(currentUserId, { configJson: JSON.stringify(configJson) });
    debugLog('✅ 表示シートを切り替えました: %s - %s', spreadsheetId, sheetName);
    return { status: 'success', message: '表示シートを切り替えました。' };
  } catch (e) {
    console.error('シート切り替えエラー: ' + e.message);
    return { status: 'error', message: '表示シートの切り替えに失敗しました: ' + e.message };
  }
}

// =================================================================
// セットアップ関数
// =================================================================

/**
 * アプリケーションの初期セットアップ（管理者が手動で実行）
 */
function setupApplication(credsJson, dbId) {
  try {
    JSON.parse(credsJson);
    if (typeof dbId !== 'string' || dbId.length !== 44) {
      throw new Error('無効なスプレッドシートIDです。IDは44文字の文字列である必要があります。');
    }

    var props = PropertiesService.getScriptProperties();
    props.setProperty(SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS, credsJson);
    props.setProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID, dbId);

    // データベースシートの初期化
    initializeDatabaseSheet(dbId);

    console.log('✅ セットアップが正常に完了しました。');
    return { status: 'success', message: 'セットアップが正常に完了しました。' };
  } catch (e) {
    console.error('セットアップエラー:', e);
    throw new Error('セットアップに失敗しました: ' + e.message);
  }
}

/**
 * セットアップ状態をテストする
 */
function testSetup() {
  try {
    var props = PropertiesService.getScriptProperties();
    var dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
    var creds = props.getProperty(SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS);
    
    if (!dbId) {
      return { status: 'error', message: 'データベーススプレッドシートIDが設定されていません。' };
    }
    
    if (!creds) {
      return { status: 'error', message: 'サービスアカウント認証情報が設定されていません。' };
    }
    
    // データベースへの接続テスト
    try {
      var userInfo = findUserByEmail(Session.getActiveUser().getEmail());
      return { 
        status: 'success', 
        message: 'セットアップは正常に完了しています。システムは使用準備が整いました。',
        details: {
          databaseConnected: true,
          userCount: userInfo ? 'ユーザー登録済み' : '未登録',
          serviceAccountConfigured: true
        }
      };
    } catch (dbError) {
      return { 
        status: 'warning', 
        message: '設定は保存されていますが、データベースアクセスに問題があります。',
        details: { error: dbError.message }
      };
    }
    
  } catch (e) {
    console.error('セットアップテストエラー:', e);
    return { status: 'error', message: 'セットアップテストに失敗しました: ' + e.message };
  }
}

// =================================================================
// ヘルパー関数
// =================================================================

// include 関数は main.gs で定義されています


function getResponsesData(userId, sheetName) {
  var userInfo = findUserById(userId);
  if (!userInfo) {
    return { status: 'error', message: 'ユーザー情報が見つかりません' };
  }

  try {
    var service = getSheetsService();
    var spreadsheetId = userInfo.spreadsheetId;
    var range = "'" + (sheetName || 'フォームの回答 1') + "'!A:Z";
    
    var response = batchGetSheetsData(service, spreadsheetId, [range]);
    var values = response.valueRanges[0].values || [];
    
    if (values.length === 0) {
      return { status: 'success', data: [], headers: [] };
    }

    return {
      status: 'success',
      data: values.slice(1),
      headers: values[0]
    };
  } catch (e) {
    console.error('回答データの取得に失敗: ' + e.message);
    return { status: 'error', message: '回答データの取得に失敗しました: ' + e.message };
  }
}

// =================================================================
// HTML依存関数（UI連携）
// =================================================================

/**
 * 管理画面用のステータス情報を取得
 * AdminPanel.htmlから呼び出される
 */
function getStatus(forceRefresh = false) {
  if (forceRefresh) {
    // Force cache invalidation for current user
    try {
      var props = PropertiesService.getUserProperties();
      var currentUserId = props.getProperty('CURRENT_USER_ID');
      if (currentUserId) {
        var userInfo = findUserById(currentUserId);
        if (userInfo) {
          invalidateUserCache(currentUserId, userInfo.adminEmail, userInfo.spreadsheetId);
          console.log('強制リフレッシュ: ユーザーキャッシュを削除しました');
        }
      }
    } catch (e) {
      console.warn('強制リフレッシュ中にエラー:', e.message);
    }
  }
  
  return getAppConfig();
}

/**
 * アクティブなフォーム情報を取得
 * AdminPanel.htmlから呼び出される（パラメータなし）
 */
function getActiveFormInfo(userId) {
  try {
    var props = PropertiesService.getUserProperties();
    var currentUserId = userId || props.getProperty('CURRENT_USER_ID');
    
    if (!currentUserId) {
      // コンテキストが設定されていない場合、現在のユーザーで検索
      var activeUser = Session.getActiveUser().getEmail();
      var userInfo = findUserByEmail(activeUser);
      if (userInfo) {
        currentUserId = userInfo.userId;
        props.setProperty('CURRENT_USER_ID', currentUserId);
      } else {
        return { status: 'error', message: 'ユーザー情報が見つかりません' };
      }
    }
    
    var userInfo = getUserWithFallback(currentUserId);
    if (!userInfo) {
      handleMissingUser(currentUserId);
      var fallbackEmail = Session.getActiveUser().getEmail();
      var altUser = findUserByEmail(fallbackEmail);
      if (altUser) {
        currentUserId = altUser.userId;
        props.setProperty('CURRENT_USER_ID', currentUserId);
        userInfo = altUser;
      } else {
        return { status: 'error', message: 'ユーザー情報が見つかりません' };
      }
    }

    var configJson = JSON.parse(userInfo.configJson || '{}');
    
    // フォーム回答数を取得
    var answerCount = 0;
    try {
      if (configJson.publishedSpreadsheetId && configJson.publishedSheet) {
        var responseData = getResponsesData(currentUserId, configJson.publishedSheet);
        if (responseData.status === 'success') {
          answerCount = responseData.data.length;
        }
      }
    } catch (countError) {
      console.warn('回答数の取得に失敗: ' + countError.message);
    }
    
    return {
      status: 'success',
      formUrl: configJson.formUrl || '',
      editFormUrl: configJson.editFormUrl || '',
      editUrl: configJson.editFormUrl || '',  // AdminPanel.htmlが期待するフィールド名
      formId: extractFormIdFromUrl(configJson.formUrl || configJson.editFormUrl || ''),
      spreadsheetUrl: userInfo.spreadsheetUrl || '',
      answerCount: answerCount,
      isFormActive: !!(configJson.formUrl && configJson.formCreated)
    };
  } catch (e) {
    console.error('フォーム情報の取得に失敗: ' + e.message);
    return { status: 'error', message: 'フォーム情報の取得に失敗しました: ' + e.message };
  }
}

/**
 * ボードデータを再読み込み
 * AdminPanel.htmlから呼び出される
 */
function refreshBoardData(userId) {
  try {
    var props = PropertiesService.getUserProperties();
    var currentUserId = userId || props.getProperty('CURRENT_USER_ID');
    
    if (!currentUserId) {
      throw new Error('ユーザーコンテキストが設定されていません');
    }
    
    var userInfo = getUserWithFallback(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    
    // キャッシュをクリア
    invalidateUserCache(currentUserId, userInfo.adminEmail, userInfo.spreadsheetId);
    
    // 最新のステータスを取得
    return getAppConfig();
  } catch (e) {
    console.error('ボードデータの再読み込みに失敗: ' + e.message);
    return { status: 'error', message: 'ボードデータの再読み込みに失敗しました: ' + e.message };
  }
}

/**
 * フォームURLを追加
 * AdminPanel.htmlから呼び出される
 */
function addFormUrl(formUrl) {
  try {
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    
    if (!currentUserId) {
      throw new Error('ユーザーコンテキストが設定されていません');
    }
    
    if (!formUrl || typeof formUrl !== 'string') {
      throw new Error('無効なフォームURLです');
    }
    
    // フォームIDを抽出
    var formId = extractFormIdFromUrl(formUrl);
    if (!formId) {
      throw new Error('フォームURLからIDを抽出できませんでした');
    }
    
    var userInfo = getUserWithFallback(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    
    var configJson = JSON.parse(userInfo.configJson || '{}');
    configJson.formUrl = formUrl;
    configJson.formId = formId;
    configJson.formCreated = true;
    
    updateUser(currentUserId, {
      configJson: JSON.stringify(configJson)
    });
    
    return { 
      status: 'success', 
      message: 'フォームURLが正常に追加されました',
      formId: formId
    };
  } catch (e) {
    console.error('フォームURL追加エラー: ' + e.message);
    return { status: 'error', message: 'フォームURLの追加に失敗しました: ' + e.message };
  }
}

/**
 * 追加フォームを作成
 * AdminPanel.htmlから呼び出される
 */
function createAdditionalForm(title) {
  try {
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    
    if (!currentUserId) {
      throw new Error('ユーザーコンテキストが設定されていません');
    }
    
    var userInfo = getUserWithFallback(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    
    // 新しいフォームを作成
    var formTitle = title || 'StudyQuest 追加フォーム - ' + new Date().toLocaleDateString('ja-JP');
    var formAndSsInfo = createStudyQuestForm(userInfo.adminEmail, currentUserId, formTitle);
    
    return {
      status: 'success',
      message: '新しいフォームが正常に作成されました',
      formUrl: formAndSsInfo.formUrl,
      editFormUrl: formAndSsInfo.editFormUrl,
      spreadsheetUrl: formAndSsInfo.spreadsheetUrl,
      formTitle: formAndSsInfo.formTitle
    };
  } catch (e) {
    console.error('追加フォーム作成エラー: ' + e.message);
    return { status: 'error', message: '追加フォームの作成に失敗しました: ' + e.message };
  }
}

/**
 * フォーム設定を更新
 * AdminPanel.htmlから呼び出される
 */
function updateFormSettings(title, description) {
  try {
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    
    if (!currentUserId) {
      throw new Error('ユーザーコンテキストが設定されていません');
    }
    
    var userInfo = getUserWithFallback(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    
    var configJson = JSON.parse(userInfo.configJson || '{}');
    
    if (configJson.editFormUrl) {
      try {
        var formId = extractFormIdFromUrl(configJson.editFormUrl);
        var form = FormApp.openById(formId);
        
        if (title) {
          form.setTitle(title);
        }
        if (description) {
          form.setDescription(description);
        }
        
        return {
          status: 'success',
          message: 'フォーム設定が更新されました'
        };
      } catch (formError) {
        console.error('フォーム更新エラー: ' + formError.message);
        return { status: 'error', message: 'フォームの更新に失敗しました: ' + formError.message };
      }
    } else {
      return { status: 'error', message: 'フォームが見つかりません' };
    }
  } catch (e) {
    console.error('フォーム設定更新エラー: ' + e.message);
    return { status: 'error', message: 'フォーム設定の更新に失敗しました: ' + e.message };
  }
}

/**
 * システム設定を保存
 * AdminPanel.htmlから呼び出される
 */
function saveSystemConfig(config) {
  try {
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    
    if (!currentUserId) {
      throw new Error('ユーザーコンテキストが設定されていません');
    }
    
    var userInfo = getUserWithFallback(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    
    var configJson = JSON.parse(userInfo.configJson || '{}');
    
    // システム設定を更新
    configJson.systemConfig = {
      cacheEnabled: config.cacheEnabled,
      autosaveInterval: config.autosaveInterval,
      logLevel: config.logLevel,
      updatedAt: new Date().toISOString()
    };
    
    updateUser(currentUserId, {
      configJson: JSON.stringify(configJson)
    });
    
    return {
      status: 'success',
      message: 'システム設定が保存されました'
    };
  } catch (e) {
    console.error('システム設定保存エラー: ' + e.message);
    return { status: 'error', message: 'システム設定の保存に失敗しました: ' + e.message };
  }
}

/**
 * ハイライト状態の切り替え
 * Page.htmlから呼び出される - フロントエンド期待形式に対応
 */
function toggleHighlight(rowIndex, sheetName) {
  try {
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    
    if (!currentUserId) {
      throw new Error('ユーザーコンテキストが設定されていません');
    }
    
    var userInfo = getUserWithFallback(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    
    // 管理者権限チェック - 現在のユーザーがボードの所有者かどうかを確認
    var activeUserEmail = Session.getActiveUser().getEmail();
    if (activeUserEmail !== userInfo.adminEmail) {
      throw new Error('ハイライト機能は管理者のみ使用できます');
    }
    
    var result = processHighlightToggle(
      userInfo.spreadsheetId,
      sheetName || 'フォームの回答 1',
      rowIndex
    );
    
    // Page.html期待形式に変換
    if (result && result.status === 'success') {
      return {
        status: "ok",
        highlight: result.highlighted || false
      };
    } else {
      throw new Error(result.message || 'ハイライト切り替えに失敗しました');
    }
  } catch (e) {
    console.error('ハイライト切り替えエラー: ' + e.message);
    return { 
      status: "error", 
      message: e.message 
    };
  }
}



/**
 * 利用可能なシート一覧を取得
 * Page.htmlから呼び出される - フロントエンド期待形式に対応
 */
function getAvailableSheets() {
  try {
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    
    if (!currentUserId) {
      console.warn('getAvailableSheets: No current user ID set');
      return [];
    }
    
    var sheets = getSheetsList(currentUserId);
    
    if (!sheets || sheets.length === 0) {
      console.warn('getAvailableSheets: No sheets found for user:', currentUserId);
      return [];
    }
    
    // Page.html期待形式に変換: [{name: string}]
    return sheets.map(function(sheet) {
      return {
        name: typeof sheet === 'string' ? sheet : (sheet.name || sheet.title || sheet)
      };
    });
  } catch (e) {
    console.error('getAvailableSheets エラー: ' + e.message);
    console.error('Error details:', e.stack);
    return [];
  }
}

/**
 * クイックスタートセットアップ（完全版）
 * フォルダ作成、フォーム作成、スプレッドシート作成、ボード公開まで一括実行
 */
function quickStartSetup(userId) {
  try {
    debugLog('🚀 クイックスタートセットアップ開始: ' + userId);
    
    // ユーザー情報の取得
    var userInfo = findUserById(userId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    
    var configJson = JSON.parse(userInfo.configJson || '{}');
    var userEmail = userInfo.adminEmail;
    
    // 既にセットアップ済みかチェック
    if (configJson.formCreated && userInfo.spreadsheetId) {
      var appUrls = generateAppUrls(userId);
      return {
        status: 'already_completed',
        message: 'クイックスタートは既に完了しています。',
        webAppUrl: appUrls.webAppUrl,
        adminUrl: appUrls.adminUrl,
        viewUrl: appUrls.viewUrl,
        setupUrl: appUrls.setupUrl
      };
    }
    
    // ステップ1: ユーザー専用フォルダを作成
    debugLog('📁 ステップ1: フォルダ作成中...');
    var folder = createUserFolder(userEmail);
    
    // ステップ2: Googleフォームとスプレッドシートを作成
    debugLog('📝 ステップ2: フォーム作成中...');
    var formAndSsInfo = createStudyQuestForm(userEmail, userId);
    
    // 作成したファイルをフォルダに移動
    if (folder) {
      try {
        var formFile = DriveApp.getFileById(formAndSsInfo.formId);
        var ssFile = DriveApp.getFileById(formAndSsInfo.spreadsheetId);
        
        folder.addFile(formFile);
        folder.addFile(ssFile);
        
        // 元の場所から削除（Myドライブから移動）
        DriveApp.getRootFolder().removeFile(formFile);
        DriveApp.getRootFolder().removeFile(ssFile);
        
        debugLog('📁 ファイルをフォルダに移動しました: ' + folder.getName());
      } catch (moveError) {
        // フォルダ移動に失敗しても処理は継続
        console.warn('ファイル移動に失敗しましたが、処理を継続します: ' + moveError.message);
      }
    }
    
    // ステップ3: データベースを更新
    debugLog('💾 ステップ3: データベース更新中...');
    var updatedConfig = {
      ...configJson,
      setupStatus: 'completed',
      formCreated: true,
      formUrl: formAndSsInfo.viewFormUrl || formAndSsInfo.formUrl,
      editFormUrl: formAndSsInfo.editFormUrl,
      publishedSpreadsheetId: formAndSsInfo.spreadsheetId,
      publishedSheetName: formAndSsInfo.sheetName || 'フォームの回答 1',
      appPublished: true,
      folderId: folder ? folder.getId() : '',
      folderUrl: folder ? folder.getUrl() : '',
      completedAt: new Date().toISOString()
    };
    
    updateUser(userId, {
      spreadsheetId: formAndSsInfo.spreadsheetId,
      spreadsheetUrl: formAndSsInfo.spreadsheetUrl,
      configJson: JSON.stringify(updatedConfig)
    });
    // セットアップ完了後に関連キャッシュをクリア
    invalidateUserCache(userId, userEmail, formAndSsInfo.spreadsheetId);
    
    // ステップ4: 回答ボードを公開状態に設定
    debugLog('🌐 ステップ4: 回答ボード公開中...');
    
    debugLog('✅ クイックスタートセットアップ完了: ' + userId);
    
    var appUrls = generateAppUrls(userId);
    return {
      status: 'success',
      message: 'クイックスタートが完了しました！回答ボードをお楽しみください。',
      webAppUrl: appUrls.webAppUrl,
      adminUrl: appUrls.adminUrl,
      viewUrl: appUrls.viewUrl,
      setupUrl: appUrls.setupUrl,
      formUrl: updatedConfig.formUrl,
      editFormUrl: updatedConfig.editFormUrl,
      spreadsheetUrl: formAndSsInfo.spreadsheetUrl,
      folderUrl: updatedConfig.folderUrl
    };
    
  } catch (e) {
    console.error('❌ クイックスタートセットアップエラー: ' + e.message);
    
    // エラー時はセットアップ状態をリセット
    try {
      var currentConfig = JSON.parse(userInfo.configJson || '{}');
      currentConfig.setupStatus = 'error';
      currentConfig.lastError = e.message;
      currentConfig.errorAt = new Date().toISOString();
      
      updateUser(userId, {
        configJson: JSON.stringify(currentConfig)
      });
      invalidateUserCache(userId, userEmail);
    } catch (updateError) {
      console.error('エラー状態の更新に失敗: ' + updateError.message);
    }
    
    return {
      status: 'error',
      message: 'クイックスタートセットアップに失敗しました: ' + e.message,
      webAppUrl: '',
      adminUrl: '',
      viewUrl: '',
      setupUrl: ''
    };
  }
}

/**
 * ユーザー専用フォルダを作成
 */
function createUserFolder(userEmail) {
  try {
    var rootFolderName = "StudyQuest - ユーザーデータ";
    var userFolderName = "StudyQuest - " + userEmail + " - ファイル";
    
    // ルートフォルダを検索または作成
    var rootFolder;
    var folders = DriveApp.getFoldersByName(rootFolderName);
    if (folders.hasNext()) {
      rootFolder = folders.next();
    } else {
      rootFolder = DriveApp.createFolder(rootFolderName);
    }
    
    // ユーザー専用フォルダを作成
    var userFolders = rootFolder.getFoldersByName(userFolderName);
    if (userFolders.hasNext()) {
      return userFolders.next(); // 既存フォルダを返す
    } else {
      return rootFolder.createFolder(userFolderName);
    }
    
  } catch (e) {
    console.error('フォルダ作成エラー: ' + e.message);
    return null; // フォルダ作成に失敗してもnullを返して処理を継続
  }
}

/**
 * ハイライト切り替え処理
 */
function processHighlightToggle(spreadsheetId, sheetName, rowIndex) {
  try {
    var service = getSheetsService();
    var headerIndices = getHeaderIndices(spreadsheetId, sheetName);
    var highlightColumnIndex = headerIndices[COLUMN_HEADERS.HIGHLIGHT];
    
    if (highlightColumnIndex === undefined) {
      throw new Error('ハイライト列が見つかりません');
    }
    
    // 現在の値を取得
    var range = "'" + sheetName + "'!" + String.fromCharCode(65 + highlightColumnIndex) + rowIndex;
    var response = batchGetSheetsData(service, spreadsheetId, [range]);
    var isHighlighted = false;
    if (response && response.valueRanges && response.valueRanges[0] && 
        response.valueRanges[0].values && response.valueRanges[0].values[0] &&
        response.valueRanges[0].values[0][0]) {
      isHighlighted = response.valueRanges[0].values[0][0] === 'true';
    }
    
    // 値を切り替え
    var newValue = isHighlighted ? 'false' : 'true';
    updateSheetsData(service, spreadsheetId, range, [[newValue]]);
    
    return {
      status: 'success',
      highlighted: !isHighlighted,
      message: isHighlighted ? 'ハイライトを解除しました' : 'ハイライトしました'
    };
  } catch (e) {
    console.error('ハイライト処理エラー: ' + e.message);
    return { status: 'error', message: e.message };
  }
}

// =================================================================
// 互換性関数（後方互換性のため）
// =================================================================

function getWebAppUrl() {
  return getWebAppUrlCached();
}


function getHeaderIndices(spreadsheetId, sheetName) {
  debugLog('getHeaderIndices received in core.gs: spreadsheetId=%s, sheetName=%s', spreadsheetId, sheetName);
  return getHeadersCached(spreadsheetId, sheetName);
}

function clearAllCaches() {
  cacheManager.clearExpired();
  clearServiceAccountTokenCache();
}

function debugLog() {
  if (DEBUG && typeof console !== 'undefined' && console.log) {
    console.log.apply(console, arguments);
  }
}

/**
 * GoogleフォームURLからフォームIDを抽出
 */
function extractFormIdFromUrl(url) {
  if (!url) return '';
  
  try {
    // Regular expression to extract form ID from Google Forms URLs
    var formIdMatch = url.match(/\/forms\/d\/([a-zA-Z0-9-_]+)/);
    if (formIdMatch && formIdMatch[1]) {
      return formIdMatch[1];
    }
    
    // Alternative pattern for e/ URLs  
    var eFormIdMatch = url.match(/\/forms\/d\/e\/([a-zA-Z0-9-_]+)/);
    if (eFormIdMatch && eFormIdMatch[1]) {
      return eFormIdMatch[1];
    }
    
    return '';
  } catch (e) {
    console.warn('フォームID抽出エラー: ' + e.message);
    return '';
  }
}

// =================================================================
// リアクション処理関数
// =================================================================

/**
 * リアクション処理
 */
function processReaction(spreadsheetId, sheetName, rowIndex, reactionKey, reactingUserEmail) {
  try {
    // LockServiceを使って競合を防ぐ
    var lock = LockService.getScriptLock();
    try {
      lock.waitLock(10000);
      
      var service = getSheetsService();
      var headerIndices = getHeaderIndices(spreadsheetId, sheetName);
      
      var reactionColumnName = COLUMN_HEADERS[reactionKey];
      var reactionColumnIndex = headerIndices[reactionColumnName];
      
      if (reactionColumnIndex === undefined) {
        throw new Error('リアクション列が見つかりません: ' + reactionColumnName);
      }
      
      // 現在のリアクション文字列を取得
      var cellRange = "'" + sheetName + "'!" + String.fromCharCode(65 + reactionColumnIndex) + rowIndex;
      var response = batchGetSheetsData(service, spreadsheetId, [cellRange]);
      var currentReactionString = '';
      if (response && response.valueRanges && response.valueRanges[0] && 
          response.valueRanges[0].values && response.valueRanges[0].values[0] &&
          response.valueRanges[0].values[0][0]) {
        currentReactionString = response.valueRanges[0].values[0][0];
      }
      
      // リアクションの追加/削除処理
      var currentReactions = parseReactionString(currentReactionString);
      var userIndex = currentReactions.indexOf(reactingUserEmail);
      
      if (userIndex >= 0) {
        // 既にリアクション済み → 削除
        currentReactions.splice(userIndex, 1);
      } else {
        // 未リアクション → 追加
        currentReactions.push(reactingUserEmail);
      }
      
      // 更新された値を書き戻す
      var updatedReactionString = currentReactions.join(', ');
      updateSheetsData(service, spreadsheetId, cellRange, [[updatedReactionString]]);
      
      debugLog('リアクション更新完了: ' + reactingUserEmail + ' → ' + reactionKey + ' (' + (userIndex >= 0 ? '削除' : '追加') + ')');
      
      return { 
        status: 'success', 
        message: 'リアクションを更新しました。',
        action: userIndex >= 0 ? 'removed' : 'added',
        count: currentReactions.length
      };
      
    } finally {
      lock.releaseLock();
    }
    
  } catch (e) {
    console.error('リアクション処理エラー: ' + e.message);
    return { 
      status: 'error', 
      message: 'リアクションの処理に失敗しました: ' + e.message 
    };
  }
}

// =================================================================
// フォーム作成機能
// =================================================================

/**
 * フォーム作成ファクトリ
 * @param {Object} options - フォーム作成オプション
 * @param {string} options.userEmail - ユーザーのメールアドレス
 * @param {string} options.userId - ユーザーID
 * @param {string} options.questions - 質問タイプ（'default'など）
 * @param {string} options.formDescription - フォームの説明
 * @returns {Object} フォーム作成結果
 */
function createFormFactory(options) {
  try {
    var userEmail = options.userEmail;
    var userId = options.userId;
    var formDescription = options.formDescription || '🌟 みんなの回答ボード - デジタル・シティズンシップを育む対話の場 🌟\n\nここは、あなたの考えを安心して表現し、多様な意見を尊重し合う学びの場です。';
    
    // タイムスタンプ生成
    var now = new Date();
    var dateTimeString = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年MM月dd日 HH:mm:ss');
    
    // フォームタイトル生成
    var formTitle = '📝 みんなの回答ボード - ' + userEmail + ' - ' + dateTimeString;
    
    // フォーム作成
    var form = FormApp.create(formTitle);
    form.setDescription(formDescription);
    
    // 基本的な質問を追加
    addUnifiedQuestions(form, options.questions || 'default', {});
    
    // スプレッドシート作成
    var spreadsheetResult = createLinkedSpreadsheet(userEmail, form, dateTimeString);
    
    return {
      formId: form.getId(),
      formUrl: form.getPublishedUrl(),
      spreadsheetId: spreadsheetResult.spreadsheetId,
      spreadsheetUrl: spreadsheetResult.spreadsheetUrl,
      sheetName: spreadsheetResult.sheetName
    };
    
  } catch (error) {
    console.error('createFormFactory エラー:', error.message);
    throw new Error('フォーム作成ファクトリでエラーが発生しました: ' + error.message);
  }
}

/**
 * 統一された質問をフォームに追加
 * @param {GoogleAppsScript.Forms.Form} form - フォームオブジェクト
 * @param {string} questionType - 質問タイプ
 * @param {Object} customConfig - カスタム設定
 */
function addUnifiedQuestions(form, questionType, customConfig) {
  try {
    var config = getQuestionConfig(questionType, customConfig);
    
    // メールアドレス収集を有効化
    form.setCollectEmail(true);
    
    // クラス質問
    var classItem = form.addTextItem();
    classItem.setTitle(config.classQuestion.title);
    classItem.setHelpText(config.classQuestion.helpText);
    classItem.setRequired(true);
    
    // メイン質問
    var mainItem = form.addParagraphTextItem();
    mainItem.setTitle(config.mainQuestion.title);
    mainItem.setHelpText(config.mainQuestion.helpText);
    mainItem.setRequired(true);
    
    // 理由質問
    var reasonItem = form.addParagraphTextItem();
    reasonItem.setTitle(config.reasonQuestion.title);
    reasonItem.setHelpText(config.reasonQuestion.helpText);
    reasonItem.setRequired(false);
    
    // 名前質問
    var nameItem = form.addTextItem();
    nameItem.setTitle(config.nameQuestion.title);
    nameItem.setHelpText(config.nameQuestion.helpText);
    nameItem.setRequired(false);
    
    console.log('フォームに統一質問を追加しました: ' + questionType);
    
  } catch (error) {
    console.error('addUnifiedQuestions エラー:', error.message);
    throw error;
  }
}

/**
 * 質問設定を取得
 * @param {string} questionType - 質問タイプ
 * @param {Object} customConfig - カスタム設定
 * @returns {Object} 質問設定
 */
function getQuestionConfig(questionType, customConfig) {
  var defaultConfig = {
    classQuestion: {
      title: 'あなたのクラス・学年',
      helpText: 'あなたが所属するクラスや学年を教えてください（例：6年1組、中学3年A組、高校2年など）'
    },
    mainQuestion: {
      title: '今日のテーマについて、あなたの考えや意見を聞かせてください',
      helpText: 'この質問には正解や間違いはありません。あなた自身の考えを自由に、安心して表現してください。多様な意見こそが、みんなの学びを豊かにします。'
    },
    reasonQuestion: {
      title: 'そう考える理由や体験があれば教えてください（任意）',
      helpText: 'あなたがそのように考える背景や、関連する体験・エピソードがあれば共有してください。理由を書くことで、あなたの考えがより相手に伝わりやすくなります。'
    },
    nameQuestion: {
      title: 'ニックネーム・呼び名（任意）',
      helpText: 'みんなの回答ボードで表示される名前です。本名でも、普段呼ばれているニックネームでも、空白でも構いません。あなたが安心できる形で参加してください。'
    }
  };
  
  // カスタム設定をマージ
  if (customConfig && typeof customConfig === 'object') {
    for (var key in customConfig) {
      if (defaultConfig[key]) {
        Object.assign(defaultConfig[key], customConfig[key]);
      }
    }
  }
  
  return defaultConfig;
}

/**
 * リンクされたスプレッドシートを作成
 * @param {string} userEmail - ユーザーメール
 * @param {GoogleAppsScript.Forms.Form} form - フォーム
 * @param {string} dateTimeString - 日時文字列
 * @returns {Object} スプレッドシート情報
 */
function createLinkedSpreadsheet(userEmail, form, dateTimeString) {
  try {
    // スプレッドシート名を設定
    var spreadsheetName = userEmail + ' - 回答データ - ' + dateTimeString;
    
    // 新しいスプレッドシートを作成
    var spreadsheetObj = SpreadsheetApp.create(spreadsheetName);
    var spreadsheetId = spreadsheetObj.getId();
    
    // フォームの回答先としてスプレッドシートを設定
    form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheetId);
    
    // スプレッドシート名を再設定（フォーム設定後に名前が変わる場合があるため）
    spreadsheetObj.rename(spreadsheetName);
    
    // シート名を取得（通常は「フォームの回答 1」）
    var sheets = spreadsheetObj.getSheets();
    var sheetName = sheets[0].getName();
    
    return {
      spreadsheetId: spreadsheetId,
      spreadsheetUrl: spreadsheetObj.getUrl(),
      sheetName: sheetName
    };
    
  } catch (error) {
    console.error('createLinkedSpreadsheet エラー:', error.message);
    throw new Error('スプレッドシート作成に失敗しました: ' + error.message);
  }
}

/**
 * フォーム作成
 */
function createStudyQuestForm(userEmail, userId) {
  try {
    // パフォーマンス測定開始
    var profiler = (typeof globalProfiler !== 'undefined') ? globalProfiler : {
      start: function() {},
      end: function() {}
    };
    profiler.start('createForm');
    
    // 共通ファクトリを使用してフォーム作成
    var formResult = createFormFactory({
      userEmail: userEmail,
      userId: userId,
      questions: 'default',
      formDescription: `🌟 みんなの回答ボードへようこそ！🌟

ここは、あなたの大切な考えや意見を安心して表現できる場所です。

📚 このプラットフォームは「デジタル・シティズンシップ」の理念に基づいて設計されています：
• 多様な意見を尊重し合う
• 建設的で思いやりのある対話を心がける
• オンライン空間でも相手への敬意を忘れない
• 自分の考えを誠実に表現する

🔒 あなたのプライバシーを大切にします：
• 回答は匿名で表示されます
• 個人情報は適切に保護されます
• 安心して本音を共有してください

あなたの声が、みんなの学びを豊かにします。一緒に素晴らしい対話の場を作りましょう！`
    });
    
    // カスタマイズされた設定を追加
    var form = FormApp.openById(formResult.formId);
    
    // Email収集タイプの設定（可能な場合）
    try {
      if (typeof form.setEmailCollectionType === 'function') {
        form.setEmailCollectionType(FormApp.EmailCollectionType.VERIFIED);
      }
    } catch (undocumentedError) {
      // 失敗しても続行
      console.warn('Email collection type setting failed:', undocumentedError.message);
    }
    
    // 確認メッセージの設定（回答ボードURLを含む）
    var appUrls = generateAppUrls(userId);
    var boardUrl = appUrls.viewUrl || (appUrls.webAppUrl + '?userId=' + encodeURIComponent(userId || ''));
    
    var confirmationMessage = `🎉 素晴らしい！あなたの声が届きました！

✨ あなたの考えを共有してくれて、ありがとうございます。
あなたの意見は、クラスのみんなにとって大切な学びの材料になります。

🤝 デジタル・シティズンシップを実践しよう：
• 他の人の意見も尊重しながら読んでみましょう
• 違う考えに出会ったら、「なるほど！」で反応してみよう
• 良いと思った意見には「いいね！」で応援しよう
• もっと知りたいことがあれば「もっと知りたい！」で示そう

🌈 多様な意見こそが、みんなの成長につながります。
お友達の色々な考えも見てみましょう。きっと新しい発見がありますよ！

📋 みんなの回答ボード:${boardUrl}`;
      
    if (form.getPublishedUrl()) {
      confirmationMessage += `

✏️ 回答を修正したい場合は、こちらから編集できます:${form.getPublishedUrl()}`;
    }
    
    form.setConfirmationMessage(confirmationMessage);
    
    // サービスアカウントをスプレッドシートに追加
    addServiceAccountToSpreadsheet(formResult.spreadsheetId);
    
    // リアクション列をスプレッドシートに追加
    addReactionColumnsToSpreadsheet(formResult.spreadsheetId, formResult.sheetName);
    
    profiler.end('createForm');
    return formResult;
    
  } catch (e) {
    console.error('createStudyQuestFormエラー: ' + e.message);
    throw new Error('フォームの作成に失敗しました: ' + e.message);
  }
}

/**
 * サービスアカウントをスプレッドシートに追加
 */
function addServiceAccountToSpreadsheet(spreadsheetId) {
  try {
    var props = PropertiesService.getScriptProperties();
    var serviceAccountCreds = JSON.parse(props.getProperty(SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS));
    var serviceAccountEmail = serviceAccountCreds.client_email;
    
    if (serviceAccountEmail) {
      var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      spreadsheet.addEditor(serviceAccountEmail);
      console.log('サービスアカウント (' + serviceAccountEmail + ') をスプレッドシートの編集者として追加しました。');
    }
  } catch (e) {
    console.error('サービスアカウントの追加に失敗: ' + e.message);
    // エラーでも処理は継続
  }
}

/**
 * スプレッドシートにリアクション列を追加
 */
function addReactionColumnsToSpreadsheet(spreadsheetId, sheetName) {
  try {
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.getSheets()[0];
    
    var additionalHeaders = [
      COLUMN_HEADERS.UNDERSTAND,
      COLUMN_HEADERS.LIKE,
      COLUMN_HEADERS.CURIOUS,
      COLUMN_HEADERS.HIGHLIGHT
    ];
    
    // 効率的にヘッダー情報を取得
    var currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var startCol = currentHeaders.length + 1;
    
    // バッチでヘッダーを追加
    sheet.getRange(1, startCol, 1, additionalHeaders.length).setValues([additionalHeaders]);
    
    // スタイリングを一括適用
    var allHeadersRange = sheet.getRange(1, 1, 1, currentHeaders.length + additionalHeaders.length);
    allHeadersRange.setFontWeight('bold').setBackground('#E3F2FD');
    
    // 自動リサイズ（エラーが出ても続行）
    try {
      sheet.autoResizeColumns(1, allHeadersRange.getNumColumns());
    } catch (resizeError) {
      console.warn('Auto-resize failed:', resizeError.message);
    }
    
    console.log('リアクション列を追加しました: ' + sheetName);
  }
  catch (e) {
    console.error('リアクション列追加エラー: ' + e.message);
    // エラーでも処理は継続
  }
}

/**
 * スプレッドシートの共有設定をチェックする
 * @param {string} spreadsheetId - チェックするスプレッドシートのID
 * @returns {object} status ('success' or 'error') と message
 */
function checkSpreadsheetSharingPermission(spreadsheetId) {
  try {
    var service = getDriveService(); // Drive APIサービスを取得
    var file = service.files.get({
      fileId: spreadsheetId,
      fields: 'permissions'
    });

    var isRestricted = true;
    var publicAccessFound = false;

    if (file.permissions) {
      for (var i = 0; i < file.permissions.length; i++) {
        var permission = file.permissions[i];
        // typeが'anyone'または'domain'でroleが'reader'以上の場合は公開状態
        if (permission.type === 'anyone' || (permission.type === 'domain' && permission.role !== 'reader')) {
          publicAccessFound = true;
          isRestricted = false;
          break;
        }
      }
    }

    if (publicAccessFound) {
      return {
        status: 'error',
        message: 'スプレッドシートの共有設定が「制限付き」になっていません。セキュリティのため、「制限付き」に設定してください。'
      };
    } else {
      return {
        status: 'success',
        message: 'スプレッドシートの共有設定は「制限付き」です。'
      };
    }
  } catch (e) {
    console.error('スプレッドシート共有設定チェックエラー: ' + e.message);
    return {
      status: 'error',
      message: 'スプレッドシートの共有設定の確認中にエラーが発生しました: ' + e.message
    };
  }
}

/**
 * シートデータ取得
 */
function getSheetData(userId, sheetName, classFilter, sortMode) {
  try {
    var userInfo = findUserById(userId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    
    var spreadsheetId = userInfo.spreadsheetId;
    var service = getSheetsService();
    
    // フォーム回答データのみを取得（名簿機能は使用しない）
    var ranges = [sheetName + '!A:Z'];
    
    var responses = batchGetSheetsData(service, spreadsheetId, ranges);
    var sheetData = responses.valueRanges[0].values || [];
    
    // 名簿機能は使用せず、空の配列を設定
    var rosterData = [];
    
    
    if (sheetData.length === 0) {
      return {
        status: 'success',
        data: [],
        headers: [],
        totalCount: 0
      };
    }
    
    var headers = sheetData[0];
    var dataRows = sheetData.slice(1);
    
    // ヘッダーインデックスを取得（キャッシュ利用）
    var headerIndices = getHeaderIndices(spreadsheetId, sheetName);
    
    // 名簿マップを作成（キャッシュ利用）
    var rosterMap = buildRosterMap(rosterData);
    
    // 表示モードを取得
    var configJson = JSON.parse(userInfo.configJson || '{}');
    var displayMode = configJson.displayMode || DISPLAY_MODES.ANONYMOUS;
    
    // データを処理
    var processedData = dataRows.map(function(row, index) {
      return processRowData(row, headers, headerIndices, rosterMap, displayMode, index + 2);
    });
    
    // フィルタリング
    var filteredData = processedData;
    if (classFilter && classFilter !== 'all') {
      var classIndex = headerIndices[COLUMN_HEADERS.CLASS];
      if (classIndex !== undefined) {
        filteredData = processedData.filter(function(row) {
          return row.originalData[classIndex] === classFilter;
        });
      }
    }
    
    // ソート適用
    var sortedData = applySortMode(filteredData, sortMode || 'newest');
    
    return {
      status: 'success',
      data: sortedData,
      headers: headers,
      totalCount: sortedData.length,
      displayMode: displayMode
    };
    
  } catch (e) {
    console.error('シートデータ取得エラー: ' + e.message);
    return {
      status: 'error',
      message: 'データの取得に失敗しました: ' + e.message,
      data: [],
      headers: []
    };
  }
}

/**
 * シート一覧取得
 */
function getSheetsList(userId) {
  try {
    debugLog('getSheetsList: Start for userId:', userId);
    var userInfo = findUserById(userId);
    if (!userInfo) {
      debugLog('getSheetsList: User not found:', userId);
      return [];
    }
    
    debugLog('getSheetsList: UserInfo spreadsheetId:', userInfo.spreadsheetId);
    if (!userInfo.spreadsheetId) {
      debugLog('getSheetsList: No spreadsheet ID for user:', userId);
      return [];
    }
    
    var service = getSheetsService();
    debugLog('getSheetsList: SheetsService obtained.');
    var spreadsheet = getSpreadsheetsData(service, userInfo.spreadsheetId);
    
    debugLog('getSheetsList: Spreadsheet data obtained:', JSON.stringify(spreadsheet));
    if (!spreadsheet || !spreadsheet.sheets) {
      debugLog('getSheetsList: Invalid spreadsheet data or no sheets found:', spreadsheet);
      return [];
    }
    
    var sheets = spreadsheet.sheets.map(function(sheet) {
      return {
        name: sheet.properties.title,
        id: sheet.properties.sheetId
      };
    });
    debugLog('getSheetsList: Returning sheets:', JSON.stringify(sheets));
    return sheets;
  } catch (e) {
    console.error('シート一覧取得エラー: ' + e.message);
    console.error('Error details:', e.stack);
    return [];
  }
}

/**
 * 名簿マップを構築（名簿機能無効化のため空のマップを返す）
 * @param {array} rosterData - 名簿データ（使用されません）
 * @returns {object} 空の名簿マップ
 */
function buildRosterMap(rosterData) {
  // 名簿機能は使用しないため、常に空のマップを返す
  // 名前はフォーム入力から直接取得
  return {};
}

/**
 * 行データを処理（スコア計算、名前変換など）
 */
function processRowData(row, headers, headerIndices, rosterMap, displayMode, rowNumber) {
  var processedRow = {
    rowNumber: rowNumber,
    originalData: row,
    score: 0,
    likeCount: 0,
    understandCount: 0,
    curiousCount: 0,
    isHighlighted: false
  };
  
  // リアクションカウント計算
  REACTION_KEYS.forEach(function(reactionKey) {
    var columnName = COLUMN_HEADERS[reactionKey];
    var columnIndex = headerIndices[columnName];
    
    if (columnIndex !== undefined && row[columnIndex]) {
      var reactions = parseReactionString(row[columnIndex]);
      var count = reactions.length;
      
      switch (reactionKey) {
        case 'LIKE':
          processedRow.likeCount = count;
          break;
        case 'UNDERSTAND':
          processedRow.understandCount = count;
          break;
        case 'CURIOUS':
          processedRow.curiousCount = count;
          break;
      }
    }
  });
  
  // ハイライト状態チェック
  var highlightIndex = headerIndices[COLUMN_HEADERS.HIGHLIGHT];
  if (highlightIndex !== undefined && row[highlightIndex]) {
    processedRow.isHighlighted = row[highlightIndex].toString().toLowerCase() === 'true';
  }
  
  // スコア計算
  processedRow.score = calculateRowScore(processedRow);
  
  // 名前の表示処理（フォーム入力の名前を使用）
  var nameIndex = headerIndices[COLUMN_HEADERS.NAME];
  if (nameIndex !== undefined && row[nameIndex] && displayMode === DISPLAY_MODES.NAMED) {
    processedRow.displayName = row[nameIndex];
  } else if (displayMode === DISPLAY_MODES.NAMED) {
    // 名前入力がない場合のフォールバック
    processedRow.displayName = '匿名';
  }
  
  return processedRow;
}

/**
 * 行のスコアを計算
 */
function calculateRowScore(rowData) {
  var baseScore = 1.0;
  
  // いいね！による加算
  var likeBonus = rowData.likeCount * SCORING_CONFIG.LIKE_MULTIPLIER_FACTOR;
  
  // その他のリアクションも軽微な加算
  var reactionBonus = (rowData.understandCount + rowData.curiousCount) * 0.01;
  
  // ハイライトによる大幅加算
  var highlightBonus = rowData.isHighlighted ? 0.5 : 0;
  
  // ランダム要素（同じスコアの項目をランダムに並べるため）
  var randomFactor = Math.random() * SCORING_CONFIG.RANDOM_SCORE_FACTOR;
  
  return baseScore + likeBonus + reactionBonus + highlightBonus + randomFactor;
}

/**
 * データにソートを適用
 */
function applySortMode(data, sortMode) {
  switch (sortMode) {
    case 'score':
      return data.sort(function(a, b) { return b.score - a.score; });
    case 'newest':
      return data.reverse();
    case 'oldest':
      return data; // 元の順序（古い順）
    case 'random':
      return shuffleArray(data.slice()); // コピーをシャッフル
    case 'likes':
      return data.sort(function(a, b) { return b.likeCount - a.likeCount; });
    default:
      return data;
  }
}

/**
 * 配列をシャッフル（Fisher-Yates shuffle）
 */
function shuffleArray(array) {
  for (var i = array.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var temp = array[i];
    array[i] = array[j];
    array[j] = temp;
  }
  return array;
}

/**
 * リアクション文字列をパース
 */
function parseReactionString(val) {
  if (!val) return [];
  return val.toString().split(',').map(function(s) { return s.trim(); }).filter(Boolean);
}

/**
 * ヘルパー関数：ヘッダー配列から指定した名前のインデックスを取得
 * COLUMN_HEADERSと統一された方式を使用
 */
function getHeaderIndex(headers, headerName) {
  if (!headers || !headerName) return -1;
  return headers.indexOf(headerName);
}

/**
 * COLUMN_HEADERSキーから適切なヘッダー名を取得
 * @param {string} columnKey - COLUMN_HEADERSのキー（例：'OPINION', 'CLASS'）
 * @returns {string} ヘッダー名
 */
function getColumnHeaderName(columnKey) {
  return COLUMN_HEADERS[columnKey] || '';
}

/**
 * 統一されたヘッダーインデックス取得関数
 * @param {array} headers - ヘッダー配列
 * @param {string} columnKey - COLUMN_HEADERSのキー
 * @returns {number} インデックス（見つからない場合は-1）
 */
function getColumnIndex(headers, columnKey) {
  var headerName = getColumnHeaderName(columnKey);
  return getHeaderIndex(headers, headerName);
}

/**
 * 設定された列名と実際のスプレッドシートヘッダーをマッピング
 * @param {Object} configHeaders - 設定されたヘッダー名
 * @param {Object} actualHeaderIndices - 実際のヘッダーインデックスマップ
 * @returns {Object} マッピングされたインデックス
 */
function mapConfigToActualHeaders(configHeaders, actualHeaderIndices) {
  var mappedIndices = {};
  var availableHeaders = Object.keys(actualHeaderIndices);
  debugLog('mapConfigToActualHeaders: Available headers in spreadsheet: %s', JSON.stringify(availableHeaders));
  
  // 各設定ヘッダーでマッピングを試行
  for (var configKey in configHeaders) {
    var configHeaderName = configHeaders[configKey];
    var mappedIndex = undefined;
    
    debugLog('mapConfigToActualHeaders: Trying to map %s = "%s"', configKey, configHeaderName);
    
    if (configHeaderName && actualHeaderIndices.hasOwnProperty(configHeaderName)) {
      // 完全一致
      mappedIndex = actualHeaderIndices[configHeaderName];
      debugLog('mapConfigToActualHeaders: Exact match found for %s: index %s', configKey, mappedIndex);
    } else if (configHeaderName) {
      // 部分一致で検索（大文字小文字を区別しない）
      var normalizedConfigName = configHeaderName.toLowerCase().trim();
      
      for (var actualHeader in actualHeaderIndices) {
        var normalizedActualHeader = actualHeader.toLowerCase().trim();
        if (normalizedActualHeader === normalizedConfigName) {
          mappedIndex = actualHeaderIndices[actualHeader];
          debugLog('mapConfigToActualHeaders: Case-insensitive match found for %s: "%s" -> index %s', configKey, actualHeader, mappedIndex);
          break;
        }
      }
      
      // 部分一致検索
      if (mappedIndex === undefined) {
        for (var actualHeader in actualHeaderIndices) {
          var normalizedActualHeader = actualHeader.toLowerCase().trim();
          if (normalizedActualHeader.includes(normalizedConfigName) || normalizedConfigName.includes(normalizedActualHeader)) {
            mappedIndex = actualHeaderIndices[actualHeader];
            debugLog('mapConfigToActualHeaders: Partial match found for %s: "%s" -> index %s', configKey, actualHeader, mappedIndex);
            break;
          }
        }
      }
    }
    
    mappedIndices[configKey] = mappedIndex;
    
    if (mappedIndex === undefined) {
      debugLog('mapConfigToActualHeaders: WARNING - No match found for %s = "%s"', configKey, configHeaderName);
    }
  }
  
  debugLog('mapConfigToActualHeaders: Final mapping result: %s', JSON.stringify(mappedIndices));
  return mappedIndices;
}

/**
 * 特定の行のリアクション情報を取得
 */
function getRowReactions(spreadsheetId, sheetName, rowIndex, userEmail) {
  try {
    var service = getSheetsService();
    var headerIndices = getHeaderIndices(spreadsheetId, sheetName);
    
    var reactionData = {
      UNDERSTAND: { count: 0, reacted: false },
      LIKE: { count: 0, reacted: false },
      CURIOUS: { count: 0, reacted: false }
    };
    
    // 各リアクション列からデータを取得
    ['UNDERSTAND', 'LIKE', 'CURIOUS'].forEach(function(reactionKey) {
      var columnName = COLUMN_HEADERS[reactionKey];
      var columnIndex = headerIndices[columnName];
      
      if (columnIndex !== undefined) {
        var range = sheetName + '!' + String.fromCharCode(65 + columnIndex) + rowIndex;
        try {
          var response = batchGetSheetsData(service, spreadsheetId, [range]);
          var cellValue = '';
          if (response && response.valueRanges && response.valueRanges[0] && 
              response.valueRanges[0].values && response.valueRanges[0].values[0] &&
              response.valueRanges[0].values[0][0]) {
            cellValue = response.valueRanges[0].values[0][0];
          }
          
          if (cellValue) {
            var reactions = parseReactionString(cellValue);
            reactionData[reactionKey].count = reactions.length;
            reactionData[reactionKey].reacted = reactions.indexOf(userEmail) !== -1;
          }
        } catch (cellError) {
          console.warn('リアクション取得エラー(' + reactionKey + '): ' + cellError.message);
        }
      }
    });
    
    return reactionData;
  } catch (e) {
    console.error('getRowReactions エラー: ' + e.message);
    return {
      UNDERSTAND: { count: 0, reacted: false },
      LIKE: { count: 0, reacted: false },
      CURIOUS: { count: 0, reacted: false }
    };
  }
}

// =================================================================
// 追加のコアファンクション（Code.gsから移行）
// =================================================================



/**
 * 回答ボードのデータを強制的に再読み込み
 */
function refreshBoardData() {
  try {
    cacheManager.clearExpired(); // 全キャッシュをクリア
    debugLog('回答ボードのデータ強制再読み込みをトリガーしました。');
    return { status: 'success', message: '回答ボードのデータを更新しました。' };
  } catch (e) {
    console.error('回答ボードのデータ再読み込みエラー: ' + e.message);
    return { status: 'error', message: '回答ボードのデータ更新に失敗しました: ' + e.message };
  }
}



// =================================================================
// ユーティリティ関数
// =================================================================


/**
 * メールアドレスからドメインを抽出
 * @param {string} email - メールアドレス
 * @returns {string} ドメイン部分
 */
function getEmailDomain(email) {
  return email.split('@')[1] || '';
}

/**
 * Drive APIサービスを取得
 * @returns {object} Drive APIサービス
 */
function getDriveService() {
  var accessToken = getServiceAccountTokenCached();
  return {
    accessToken: accessToken,
    baseUrl: 'https://www.googleapis.com/drive/v3',
    files: {
      get: function(params) {
        var url = this.baseUrl + '/files/' + params.fileId + '?fields=' + encodeURIComponent(params.fields);
        var response = UrlFetchApp.fetch(url, {
          headers: { 'Authorization': 'Bearer ' + this.accessToken }
        });
        return JSON.parse(response.getContentText());
      }
    }
  };
}
