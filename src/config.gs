/**
 * config.gs - 軽量化版
 * 新サービスアカウントアーキテクチャで必要最小限の関数のみ
 */

const CONFIG_SHEET_NAME = 'Config';

/**
 * 現在のユーザーのスプレッドシートを取得
 * AdminPanel.htmlから呼び出される
 */
function getCurrentSpreadsheet() {
  try {
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    
    if (!currentUserId) {
      throw new Error('ユーザーIDが設定されていません');
    }
    
    var userInfo = findUserById(currentUserId);
    if (!userInfo || !userInfo.spreadsheetId) {
      throw new Error('ユーザー情報またはスプレッドシートIDが見つかりません');
    }
    
    return SpreadsheetApp.openById(userInfo.spreadsheetId);
  } catch (e) {
    console.error('getCurrentSpreadsheet エラー: ' + e.message);
    throw new Error('スプレッドシートの取得に失敗しました: ' + e.message);
  }
}

/**
 * アクティブなスプレッドシートのURLを取得
 * AdminPanel.htmlから呼び出される
 * @returns {string} スプレッドシートURL
 */
function openActiveSpreadsheet() {
  try {
    var ss = getCurrentSpreadsheet();
    return ss.getUrl();
  } catch (e) {
    console.error('openActiveSpreadsheet エラー: ' + e.message);
    throw new Error('スプレッドシートのURL取得に失敗しました: ' + e.message);
  }
}

/**
 * 設定取得関数（AdminPanel.htmlとの互換性のため）
 * 実際のシートヘッダーに基づいた設定を返す
 * @param {string} sheetName - シート名（AdminPanelから渡される、オプション）
 */
function getConfig(sheetName) {
  try {
    var spreadsheet = getCurrentSpreadsheet();
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    console.log('getConfig: userId=%s, sheetName=%s', currentUserId, sheetName);

    // デフォルト設定（COLUMN_HEADERSと統一）
    var config = {
      questionHeader: COLUMN_HEADERS.TIMESTAMP,
      answerHeader: COLUMN_HEADERS.OPINION,
      reasonHeader: COLUMN_HEADERS.REASON,
      nameHeader: COLUMN_HEADERS.NAME,
      classHeader: COLUMN_HEADERS.CLASS,
      rosterSheetName: '名簿',
      mainHeader: COLUMN_HEADERS.OPINION,
      rHeader: COLUMN_HEADERS.REASON,
      opinionHeader: COLUMN_HEADERS.OPINION,
      timestampHeader: COLUMN_HEADERS.TIMESTAMP,
      emailHeader: COLUMN_HEADERS.EMAIL
    };

    // シート固有の設定を取得
    if (sheetName) {
      console.log('設定を取得中: シート名 = ' + sheetName);
      
      // 保存済みの設定があるかチェック
      var hasExistingConfig = false;
      try {
        if (currentUserId) {
          var userInfo = findUserById(currentUserId);
          if (userInfo && userInfo.configJson) {
            var configJson = JSON.parse(userInfo.configJson);
            console.log('getConfig: Loaded configJson: %s', userInfo.configJson);
            var sheetConfigKey = 'sheet_' + sheetName;
            
            if (configJson[sheetConfigKey]) {
              hasExistingConfig = true;
              var savedConfig = configJson[sheetConfigKey];
              
              // 保存された設定を適用
              config.mainHeader = savedConfig.mainHeader !== undefined ? savedConfig.mainHeader : config.mainHeader;
              config.rHeader = savedConfig.rHeader !== undefined ? savedConfig.rHeader : config.rHeader;
              config.nameHeader = savedConfig.nameHeader !== undefined ? savedConfig.nameHeader : config.nameHeader;
              config.classHeader = savedConfig.classHeader !== undefined ? savedConfig.classHeader : config.classHeader;
              config.answerHeader = savedConfig.mainHeader !== undefined ? savedConfig.mainHeader : config.answerHeader;
              config.opinionHeader = savedConfig.mainHeader !== undefined ? savedConfig.mainHeader : config.opinionHeader;
              config.reasonHeader = savedConfig.rHeader !== undefined ? savedConfig.rHeader : config.reasonHeader;
              
              console.log('保存済み設定を適用しました:', savedConfig);
            }
          }
        }
      } catch (configCheckError) {
        console.warn('設定チェックでエラー:', configCheckError.message);
      }
      
      try {
        var sheet = spreadsheet.getSheetByName(sheetName);
        if (sheet && sheet.getLastRow() > 0) {
          var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
          var availableHeaders = headers.filter(function(h) { return h && h.toString().trim() !== ''; });
          
          // 利用可能なヘッダー情報を追加
          config.availableHeaders = availableHeaders;
          config.sheetName = sheetName;
          
          // 保存済み設定がない場合、自動マッピングを適用
          if (!hasExistingConfig && availableHeaders.length > 0) {
            console.log('保存済み設定がないため、自動マッピングを実行します');
            
            // より高度な自動マッピングを使用
            var autoMapping = autoMapSheetHeaders(sheetName);
            if (autoMapping) {
              config.mainHeader = autoMapping.mainHeader || config.mainHeader;
              config.rHeader = autoMapping.rHeader || config.rHeader;
              config.nameHeader = autoMapping.nameHeader || config.nameHeader;
              config.classHeader = autoMapping.classHeader || config.classHeader;
              config.answerHeader = autoMapping.mainHeader || config.answerHeader;
              config.opinionHeader = autoMapping.mainHeader || config.opinionHeader;
              config.reasonHeader = autoMapping.rHeader || config.reasonHeader;
              
              console.log('自動マッピングを適用しました:', autoMapping);
              
              // 自動マッピングが成功した場合、設定を保存
              try {
                var saveResult = saveSheetConfig(sheetName, {
                  mainHeader: autoMapping.mainHeader,
                  rHeader: autoMapping.rHeader,
                  nameHeader: autoMapping.nameHeader,
                  classHeader: autoMapping.classHeader
                });
                console.log('自動マッピング設定を保存しました:', saveResult);
              } catch (saveError) {
                console.warn('自動マッピング設定の保存でエラー:', saveError.message);
              }
            } else {
              // autoMapSheetHeaders が失敗した場合、従来のマッピングを使用
              console.log('autoMapSheetHeaders失敗、従来のマッピングを使用');
              availableHeaders.forEach(function(header) {
                var headerLower = header.toString().toLowerCase();
                if (headerLower.includes('回答') || headerLower.includes('意見') || headerLower.includes('answer')) {
                  config.mainHeader = header;
                  config.answerHeader = header;
                  config.opinionHeader = header;
                }
                if (headerLower.includes('理由') || headerLower.includes('reason')) {
                  config.rHeader = header;
                  config.reasonHeader = header;
                }
                if (headerLower.includes('名前') || headerLower.includes('name')) {
                  config.nameHeader = header;
                }
                if (headerLower.includes('クラス') || headerLower.includes('class')) {
                  config.classHeader = header;
                }
              });
            }
          }
          
          console.log('シート設定を取得しました:', {
            sheetName: sheetName,
            availableHeaders: availableHeaders.length,
            mainHeader: config.mainHeader,
            hasExistingConfig: hasExistingConfig
          });
        }
      } catch (sheetError) {
        console.warn('シート固有設定の取得でエラー:', sheetError.message);
        config.availableHeaders = [];
        config.sheetName = sheetName || '';
      }
    }

    // Configシートからの追加設定を読み込み
    var configSheet = spreadsheet.getSheetByName(CONFIG_SHEET_NAME);
    if (configSheet) {
      try {
        var data = configSheet.getDataRange().getValues();
        for (var i = 0; i < data.length; i++) {
          var key = data[i][0];
          var value = data[i][1];
          if (key && value !== undefined) {
            config[key] = value;
          }
        }
      } catch (configSheetError) {
        console.warn('Configシートの読み込みでエラー:', configSheetError.message);
      }
    }
    
    console.log('getConfig: Returning config: %s', JSON.stringify(config));
    return config;
  } catch (error) {
    console.error('getConfig error:', error.message);
    // エラー時のデフォルト設定
    return {
      questionHeader: 'タイムスタンプ',
      answerHeader: '回答',
      reasonHeader: '理由',
      nameHeader: '名前',
      classHeader: 'クラス',
      rosterSheetName: '名簿',
      mainHeader: '回答',
      rHeader: '理由',
      opinionHeader: '回答',
      timestampHeader: 'タイムスタンプ',
      emailHeader: 'メールアドレス',
      availableHeaders: [],
      sheetName: sheetName || ''
    };
  }
}""

/**
 * 名簿シート名を取得
 * @returns {string} 名簿シート名
 */
function getRosterSheetName() {
  try {
    var cfg = getConfig();
    return cfg.rosterSheetName || '名簿';
  } catch (e) {
    console.warn('getRosterSheetName error: ' + e.message);
    return '名簿';
  }
}

/**
 * シート固有の設定を保存
 * @param {string} sheetName - シート名
 * @param {object} cfg - 設定オブジェクト
 */
function saveSheetConfig(sheetName, cfg) {
  try {
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    console.log('saveSheetConfig: userId=%s, sheetName=%s, config=%s', currentUserId, sheetName, JSON.stringify(cfg));
    
    if (!currentUserId) {
      throw new Error('ユーザーコンテキストが設定されていません。');
    }
    
    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません。');
    }
    
    // 現在の設定を取得
    var configJson = JSON.parse(userInfo.configJson || '{}');
    
    // シート固有の設定を保存
    var sheetConfigKey = 'sheet_' + sheetName;
    configJson[sheetConfigKey] = {
      mainHeader: cfg.mainHeader || cfg.answerHeader || '',
      rHeader: cfg.rHeader || cfg.reasonHeader || '',
      nameHeader: cfg.nameHeader || '',
      classHeader: cfg.classHeader || '',
      savedAt: new Date().toISOString()
    };
    
    // データベースに更新
    var updatedUser = updateUser(currentUserId, {
      configJson: JSON.stringify(configJson)
    });
    
    // ユーザー情報のキャッシュをクリア
    cacheManager.remove('user_' + currentUserId);
    
    console.log('シート設定を保存しました:', {
      sheetName: sheetName,
      userId: currentUserId,
      config: configJson[sheetConfigKey]
    });
    
    return `シート「${sheetName}」の設定を保存しました`;
  } catch (error) {
    console.error('saveSheetConfig エラー:', error);
    return `設定の保存中にエラーが発生しました: ${error.message}`;
  }
}

/**
 * シートヘッダーを取得
 * AdminPanel.htmlから呼び出される
 */
function getSheetHeaders(sheetName) {
  try {
    var spreadsheet = getCurrentSpreadsheet();
    
    // シート名が指定されていない場合、利用可能なシートを表示
    if (!sheetName) {
      var allSheets = spreadsheet.getSheets();
      var sheetNames = allSheets.map(function(sheet) { return sheet.getName(); });
      console.log('利用可能なシート:', sheetNames);
      throw new Error('シート名が指定されていません。利用可能なシート: ' + sheetNames.join(', '));
    }
    
    var sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      // シートが見つからない場合、利用可能なシートを表示
      var allSheets = spreadsheet.getSheets();
      var sheetNames = allSheets.map(function(sheet) { return sheet.getName(); });
      console.log('要求されたシート「' + sheetName + '」が見つかりません');
      console.log('利用可能なシート:', sheetNames);
      
      // デフォルトシートを試行
      if (sheetNames.length > 0) {
        var firstSheet = allSheets[0];
        console.log('デフォルトシート「' + firstSheet.getName() + '」を使用します');
        sheet = firstSheet;
      } else {
        throw new Error('シートが見つかりません: ' + sheetName + '. 利用可能なシート: ' + sheetNames.join(', '));
      }
    }
    
    var lastColumn = sheet.getLastColumn();
    if (lastColumn === 0) {
      return [];
    }
    var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    return headers.filter(String); // 空のヘッダーを除外
  } catch (e) {
    console.error('getSheetHeaders エラー: ' + e.message);
    throw new Error('シートヘッダーの取得に失敗しました: ' + e.message);
  }
}

/**
 * スプレッドシートの列名から自動的にconfig設定を推測する
 */
function autoMapSheetHeaders(sheetName) {
  try {
    var headers = getSheetHeaders(sheetName);
    if (!headers || headers.length === 0) {
      return null;
    }
    
    var mapping = {
      mainHeader: '',
      rHeader: '',
      nameHeader: '',
      classHeader: ''
    };
    
    // 優先度付きマッピングルール
    var mappingRules = {
      mainHeader: [
        // 高優先度
        ['今日のテーマについて、あなたの考えや意見を聞かせてください', 'あなたの回答・意見', '回答・意見', '回答内容', '意見・回答'],
        // 中優先度  
        ['回答', '意見', 'コメント', '内容', '質問', '答え', 'answer', 'opinion', 'comment'],
        // 低優先度
        ['テキスト', 'text', '記述', '入力']
      ],
      rHeader: [
        // 高優先度
        ['そう考える理由や体験があれば教えてください（任意）', '理由・根拠', '根拠・理由', '理由や根拠'],
        // 中優先度
        ['理由', '根拠', '説明', '詳細', 'reason', '理由説明'],
        // 低優先度
        ['なぜ', 'why', '背景']
      ],
      nameHeader: [
        // 高優先度
        ['名前', '氏名', 'name'],
        // 中優先度
        ['ニックネーム', '呼び名', '表示名', 'nickname'],
        // 低優先度
        ['ユーザー', 'user', '投稿者']
      ],
      classHeader: [
        // 高優先度
        ['クラス名', 'クラス', '組', 'class'],
        // 中優先度
        ['学年', 'グループ', 'チーム', 'group', 'team'],
        // 低優先度
        ['所属', '部門', 'section']
      ]
    };
    
    // 各マッピングルールを適用
    Object.keys(mappingRules).forEach(function(mappingKey) {
      if (mapping[mappingKey]) return; // 既にマッピング済み
      
      var rules = mappingRules[mappingKey];
      
      // 優先度順にチェック
      for (var priority = 0; priority < rules.length && !mapping[mappingKey]; priority++) {
        var patterns = rules[priority];
        
        for (var i = 0; i < headers.length && !mapping[mappingKey]; i++) {
          var header = headers[i];
          var lowerHeader = header.toLowerCase();
          
          // 完全一致または部分一致をチェック
          for (var j = 0; j < patterns.length; j++) {
            var pattern = patterns[j].toLowerCase();
            if (lowerHeader === pattern || lowerHeader.includes(pattern)) {
              mapping[mappingKey] = header;
              debugLog('マッピング成功: ' + mappingKey + ' = "' + header + '" (パターン: "' + pattern + '", 優先度: ' + priority + ')');
              break;
            }
          }
        }
      }
    });
    
    debugLog('自動マッピング結果 for ' + sheetName + ': ' + JSON.stringify(mapping));
    return mapping;
    
  } catch (e) {
    console.error('autoMapSheetHeaders エラー: ' + e.message);
    return null;
  }
}

/**
 * スプレッドシートURLを追加してシート検出を実行
 * Unpublished.htmlから呼び出される
 */
function addSpreadsheetUrl(url) {
  try {
    var spreadsheetId = url.match(/\/d\/([a-zA-Z0-9_-]+)/)[1];
    if (!spreadsheetId) {
      throw new Error('無効なスプレッドシートURLです。');
    }
    
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    
    if (!currentUserId) {
      throw new Error('ユーザーコンテキストが設定されていません。');
    }
    
    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません。');
    }
    
    // サービスアカウントをスプレッドシートに追加
    addServiceAccountToSpreadsheet(spreadsheetId);
    
    // ユーザー情報にスプレッドシートIDとURLを更新
    updateUser(currentUserId, {
      spreadsheetId: spreadsheetId,
      spreadsheetUrl: url
    });
    
    // シートリストを即座に取得
    var sheets = [];
    try {
      sheets = getSheetsList(currentUserId);
      console.log('シート検出完了:', {
        spreadsheetId: spreadsheetId,
        sheetCount: sheets.length,
        sheets: sheets.map(s => s.name)
      });
    } catch (sheetError) {
      console.warn('シートリスト取得でエラー:', sheetError.message);
      // シートリスト取得に失敗してもスプレッドシート追加は成功とする
    }
    
    // キャッシュをクリアして最新の状態を確保
    cacheManager.clearByPattern(spreadsheetId);
    
    return { 
      status: 'success', 
      message: 'スプレッドシートが正常に追加されました。', 
      sheets: sheets,
      spreadsheetId: spreadsheetId,
      autoSelectFirst: sheets.length > 0 ? sheets[0].name : null,
      needsRefresh: true // UI側でのリフレッシュが必要
    };
  } catch (e) {
    console.error('addSpreadsheetUrl エラー: ' + e.message);
    throw new Error('スプレッドシートの追加に失敗しました: ' + e.message);
  }
}

/**
 * シートを切り替える
 * AdminPanel.htmlから呼び出される
 */
function switchToSheet(sheetName) {
  try {
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    
    if (!currentUserId) {
      throw new Error('ユーザーコンテキストが設定されていません。');
    }
    
    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません。');
    }
    
    // シートの存在確認
    var spreadsheet = getCurrentSpreadsheet();
    var sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`指定されたシート「${sheetName}」が見つかりません。`);
    }
    
    var configJson = JSON.parse(userInfo.configJson || '{}');
    configJson.publishedSheet = sheetName;
    configJson.appPublished = true;
    configJson.lastSheetSwitch = new Date().toISOString();
    
    // シート固有の設定があるかチェック
    var sheetConfigKey = 'sheet_' + sheetName;
    if (!configJson[sheetConfigKey]) {
      console.log(`シート「${sheetName}」の列設定がありません。デフォルト設定を使用します。`);
    }
    
    updateUser(currentUserId, {
      configJson: JSON.stringify(configJson)
    });
    
    // キャッシュをクリア（シート切り替え時に最新情報を取得するため）
    cacheManager.clearByPattern(userInfo.spreadsheetId);
    
    console.log('シート切り替え完了:', {
      userId: currentUserId,
      sheetName: sheetName,
      hasSheetConfig: !!configJson[sheetConfigKey]
    });
    
    return 'シートが正常に切り替わり、公開されました。';
  } catch (e) {
    console.error('switchToSheet エラー: ' + e.message);
    throw new Error('シートの切り替えに失敗しました: ' + e.message);
  }
}

/**
 * アクティブシートをクリア（公開停止）
 * AdminPanel.htmlから呼び出される
 */
function clearActiveSheet() {
  try {
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    
    if (!currentUserId) {
      throw new Error('ユーザーコンテキストが設定されていません。');
    }
    
    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません。');
    }
    
    var configJson = JSON.parse(userInfo.configJson || '{}');
    configJson.publishedSheet = ''; // アクティブシートをクリア
    configJson.appPublished = false; // 公開停止
    
    updateUser(currentUserId, {
      configJson: JSON.stringify(configJson)
    });
    
    return '回答ボードの公開を停止しました。';
  } catch (e) {
    console.error('clearActiveSheet エラー: ' + e.message);
    throw new Error('回答ボードの公開停止に失敗しました: ' + e.message);
  }
}

/**
 * 表示オプションを設定
 * AdminPanel.htmlから呼び出される
 */
function setDisplayOptions(options) {
  try {
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    
    if (!currentUserId) {
      throw new Error('ユーザーコンテキストが設定されていません。');
    }
    
    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません。');
    }
    
    var configJson = JSON.parse(userInfo.configJson || '{}');
    configJson.showNames = options.showNames;
    configJson.showCounts = options.showCounts;
    
    updateUser(currentUserId, {
      configJson: JSON.stringify(configJson)
    });
    
    return '表示オプションを保存しました。';
  } catch (e) {
    console.error('setDisplayOptions エラー: ' + e.message);
    throw new Error('表示オプションの保存に失敗しました: ' + e.message);
  }
}

/**
 * 管理者権限チェック
 * Page.htmlから呼び出される
 */
function checkAdmin() {
  debugLog('checkAdmin function called.');
  try {
    var activeUserEmail = Session.getActiveUser().getEmail();
    var props = PropertiesService.getScriptProperties();
    var dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
    var service = getSheetsService();
    var sheetName = DB_SHEET_CONFIG.SHEET_NAME;
    
    var response = service.spreadsheets.values.get(dbId, sheetName + '!A:H');
    var data = (response && response.values) ? response.values : [];
    
    if (data.length === 0) {
      debugLog('checkAdmin: No data found in Users sheet or sheet is empty.');
      return false;
    }
    
    var headers = data[0];
    var emailIndex = headers.indexOf('adminEmail');
    var isActiveIndex = headers.indexOf('isActive');
    
    if (emailIndex === -1 || isActiveIndex === -1) {
      console.warn('checkAdmin: Missing required headers in Users sheet.');
      return false;
    }

    for (var i = 1; i < data.length; i++) {
      if (data[i][emailIndex] === activeUserEmail && data[i][isActiveIndex] === 'true') {
        return true;
      }
    }
    return false;
  } catch (e) {
    console.error('checkAdmin エラー: ' + e.message);
    return false;
  }
}

/**
 * 管理者からボードを作成
 * AdminPanel.htmlから呼び出される
 */
function createBoardFromAdmin() {
  try {
    var activeUserEmail = Session.getActiveUser().getEmail();
    var userId = Utilities.getUuid(); // 新しいユーザーIDを生成
    
    // フォームとスプレッドシートを作成
    var formAndSsInfo = createStudyQuestForm(activeUserEmail, userId);
    
    // 中央データベースにユーザー情報を登録
    var initialConfig = {
      formUrl: formAndSsInfo.viewFormUrl || formAndSsInfo.formUrl,
      editFormUrl: formAndSsInfo.editFormUrl,
      createdAt: new Date().toISOString(),
      publishedSheet: formAndSsInfo.sheetName, // 作成時に公開シートを設定
      appPublished: true // 作成時に公開状態にする
    };
    
    var userData = {
      userId: userId,
      adminEmail: activeUserEmail,
      spreadsheetId: formAndSsInfo.spreadsheetId,
      spreadsheetUrl: formAndSsInfo.spreadsheetUrl,
      createdAt: new Date().toISOString(),
      configJson: JSON.stringify(initialConfig),
      lastAccessedAt: new Date().toISOString(),
      isActive: 'true'
    };
    
    createUser(userData);
    
    // 成功レスポンスを返す
    var appUrls = generateAppUrls(userId);
    return {
      status: 'success',
      message: '新しいボードが作成され、公開されました！',
      formUrl: formAndSsInfo.formUrl,
      spreadsheetUrl: formAndSsInfo.spreadsheetUrl,
      adminUrl: appUrls.adminUrl,
      viewUrl: appUrls.viewUrl,
      formTitle: formAndSsInfo.formTitle // フォームタイトルも返す
    };
  } catch (e) {
    console.error('createBoardFromAdmin エラー: ' + e.message);
    throw new Error('ボードの作成に失敗しました: ' + e.message);
  }
}

/**
 * 既存ボード情報を取得
 * Registration.htmlから呼び出される
 */
function getExistingBoard() {
  try {
    var activeUserEmail = Session.getActiveUser().getEmail();
    var userInfo = findUserByEmail(activeUserEmail);
    
    if (userInfo && userInfo.isActive === 'true') {
      var appUrls = generateAppUrls(userInfo.userId);
      return {
        status: 'existing_user',
        userId: userInfo.userId,
        adminUrl: appUrls.adminUrl,
        viewUrl: appUrls.viewUrl
      };
    } else if (userInfo && userInfo.isActive === 'false') {
      return {
        status: 'setup_required',
        userId: userInfo.userId
      };
    } else {
      return {
        status: 'new_user'
      };
    }
  } catch (e) {
    console.error('getExistingBoard エラー: ' + e.message);
    return { status: 'error', message: '既存ボード情報の取得に失敗しました: ' + e.message };
  }
}

/**
 * ユーザー認証を検証
 * Registration.htmlから呼び出される
 */
function verifyUserAuthentication() {
  try {
    var email = Session.getActiveUser().getEmail();
    if (email) {
      return { authenticated: true, email: email };
    } else {
      return { authenticated: false, email: null };
    }
  } catch (e) {
    console.error('verifyUserAuthentication エラー: ' + e.message);
    return { authenticated: false, email: null, error: e.message };
  }
}

/**
 * デバッグ用：現在のスプレッドシートとシートの状態を確認
 */
function debugCurrentSpreadsheetState() {
  try {
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    
    if (!currentUserId) {
      return { error: 'ユーザーIDが設定されていません' };
    }
    
    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      return { error: 'ユーザー情報が見つかりません' };
    }
    
    var configJson = JSON.parse(userInfo.configJson || '{}');
    var spreadsheet = SpreadsheetApp.openById(userInfo.spreadsheetId);
    var sheets = spreadsheet.getSheets();
    var sheetNames = sheets.map(function(sheet) { return sheet.getName(); });
    
    return {
      userId: currentUserId,
      spreadsheetId: userInfo.spreadsheetId,
      spreadsheetUrl: userInfo.spreadsheetUrl,
      publishedSheet: configJson.publishedSheet,
      availableSheets: sheetNames,
      sheetCount: sheets.length,
      hasPublishedSheetInSpreadsheet: sheetNames.indexOf(configJson.publishedSheet) !== -1
    };
    
  } catch (e) {
    console.error('debugCurrentSpreadsheetState エラー: ' + e.message);
    return { error: e.message };
  }
}
