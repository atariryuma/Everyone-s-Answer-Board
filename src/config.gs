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
function getConfig(sheetName, forceRefresh = false) {
  try {
    var spreadsheet = getCurrentSpreadsheet();
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    console.log('getConfig: userId=%s, sheetName=%s, forceRefresh=%s', currentUserId, sheetName, forceRefresh);
    
    // forceRefreshが指定された場合、キャッシュを無効化
    if (forceRefresh && currentUserId) {
      try {
        var userInfo = findUserById(currentUserId);
        if (userInfo) {
          invalidateUserCache(currentUserId, userInfo.adminEmail, userInfo.spreadsheetId);
          console.log('getConfig: 強制リフレッシュによりキャッシュを削除しました');
        }
      } catch (e) {
        console.warn('getConfig: 強制リフレッシュ中にエラー:', e.message);
      }
    }

    // デフォルト設定（データベース設定キーと統一）
    var config = {
      // データベース設定キー（プライマリ）
      mainHeader: COLUMN_HEADERS.OPINION,
      rHeader: COLUMN_HEADERS.REASON,
      nameHeader: COLUMN_HEADERS.NAME,
      classHeader: COLUMN_HEADERS.CLASS,
      
      // 後方互換性のための旧キー
      questionHeader: COLUMN_HEADERS.TIMESTAMP,
      answerHeader: COLUMN_HEADERS.OPINION,
      reasonHeader: COLUMN_HEADERS.REASON,
      opinionHeader: COLUMN_HEADERS.OPINION,
      timestampHeader: COLUMN_HEADERS.TIMESTAMP,
      emailHeader: COLUMN_HEADERS.EMAIL,
      rosterSheetName: '名簿'
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
              
              // 保存された設定を適用（データベース設定キーを優先）
              config.mainHeader = savedConfig.mainHeader !== undefined ? savedConfig.mainHeader : config.mainHeader;
              config.rHeader = savedConfig.rHeader !== undefined ? savedConfig.rHeader : config.rHeader;
              config.nameHeader = savedConfig.nameHeader !== undefined ? savedConfig.nameHeader : config.nameHeader;
              config.classHeader = savedConfig.classHeader !== undefined ? savedConfig.classHeader : config.classHeader;
              
              // 統一された変数名を使用
              config.opinionHeader = config.mainHeader;
              config.timestampHeader = config.questionHeader || 'タイムスタンプ';
              config.emailHeader = config.emailHeader || 'メールアドレス';
              
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
          
          // 保存済み設定がない場合、自動マッピングを適用（自動保存なし）
          if (!hasExistingConfig && availableHeaders.length > 0) {
            console.log('保存済み設定がないため、自動マッピングを実行します（自動保存なし）');
            
            // より高度な自動マッピングを使用
            var autoMapping = autoMapSheetHeaders(sheetName);
            if (autoMapping) {
              config.mainHeader = autoMapping.mainHeader || config.mainHeader;
              config.rHeader = autoMapping.rHeader || config.rHeader;
              config.nameHeader = autoMapping.nameHeader || config.nameHeader;
              config.classHeader = autoMapping.classHeader || config.classHeader;
              
              // 統一された変数名を使用
              config.opinionHeader = config.mainHeader;
              config.timestampHeader = config.questionHeader || 'タイムスタンプ';
              config.emailHeader = config.emailHeader || 'メールアドレス';
              
              console.log('自動マッピングを適用しました（一時的）:', autoMapping);
              
              // 自動マッピングの結果を保存しない - ユーザーが明示的に保存する必要がある
              console.log('自動マッピングの結果は一時的に適用されました。明示的な保存が必要です。');
            } else {
              // autoMapSheetHeaders が失敗した場合、従来のマッピングを使用
              console.log('autoMapSheetHeaders失敗、従来のマッピングを使用');
              availableHeaders.forEach(function(header) {
                var headerLower = header.toString().toLowerCase();
                if (headerLower.includes('回答') || headerLower.includes('意見') || headerLower.includes('answer')) {
                  config.mainHeader = header;
                }
                if (headerLower.includes('理由') || headerLower.includes('reason')) {
                  config.rHeader = header;
                }
                if (headerLower.includes('名前') || headerLower.includes('name')) {
                  config.nameHeader = header;
                }
                if (headerLower.includes('クラス') || headerLower.includes('class')) {
                  config.classHeader = header;
                }
              });
              
              // 統一された変数名を使用
              config.opinionHeader = config.mainHeader;
              config.timestampHeader = config.questionHeader || 'タイムスタンプ';
              config.emailHeader = config.emailHeader || 'メールアドレス';
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
    
    // 最終的な統一された変数名での返却前処理
    var finalConfig = {
      sheetName: config.sheetName || sheetName || '',
      opinionHeader: config.opinionHeader || config.mainHeader || '回答',
      timestampHeader: config.timestampHeader || config.questionHeader || 'タイムスタンプ',
      emailHeader: config.emailHeader || 'メールアドレス',
      reasonHeader: config.reasonHeader || config.rHeader || '理由',
      nameHeader: config.nameHeader || '名前',
      classHeader: config.classHeader || 'クラス',
      showNames: config.showNames !== undefined ? config.showNames : false,
      showCounts: config.showCounts !== undefined ? config.showCounts : false,
      availableHeaders: config.availableHeaders || [],
      rosterSheetName: config.rosterSheetName || '名簿'
    };
    
    console.log('getConfig: Returning unified config: %s', JSON.stringify(finalConfig));
    return finalConfig;
  } catch (error) {
    console.error('getConfig error:', error.message);
    // エラー時のデフォルト設定（統一された変数名）
    return {
      sheetName: sheetName || '',
      opinionHeader: '回答',
      timestampHeader: 'タイムスタンプ',
      emailHeader: 'メールアドレス',
      reasonHeader: '理由',
      nameHeader: '名前',
      classHeader: 'クラス',
      showNames: false,
      showCounts: false,
      availableHeaders: [],
      rosterSheetName: '名簿'
    };
  }
}

/**
 * シート設定の保存、アクティブ化、最新ステータスの取得を一つのトランザクションで実行する統合関数
 * @param {string} spreadsheetId - 対象のスプレッドシートID
 * @param {string} sheetName - 対象のシート名
 * @param {object} config - 保存する設定オブジェクト
 * @returns {object} { success: boolean, message: string, status: object } 形式のオブジェクト
 */
function saveAndActivateSheet(spreadsheetId, sheetName, config) {
  // ★★★ここから追加：引数の検証処理★★★
  if (typeof spreadsheetId !== 'string' || !spreadsheetId) {
    throw new Error('無効なspreadsheetIdです。スプレッドシートIDは必須です。');
  }
  if (typeof sheetName !== 'string' || !sheetName) {
    throw new Error('無効なsheetNameです。シート名は必須です。');
  }
  if (typeof config !== 'object' || config === null) {
    throw new Error('無効なconfigオブジェクトです。設定オブジェクトは必須です。');
  }
  // ★★★ここまで追加★★★

  try {
    console.log('saveAndActivateSheet開始: sheetName=%s', sheetName);

    // 1. 関連キャッシュを事前に削除
    const props = PropertiesService.getUserProperties();
    const currentUserId = props.getProperty('CURRENT_USER_ID');
    if (currentUserId) {
      const userInfo = findUserById(currentUserId);
      if (userInfo) {
        invalidateUserCache(currentUserId, userInfo.adminEmail, spreadsheetId);
        console.log('saveAndActivateSheet: 事前キャッシュクリア完了');
      }
    }

    // 2. 設定を保存
    saveSheetConfig(spreadsheetId, sheetName, config);
    console.log('saveAndActivateSheet: 設定保存完了');

    // 3. シートをアクティブ化
    switchToSheet(spreadsheetId, sheetName);
    console.log('saveAndActivateSheet: シート切り替え完了');

    // 4. 表示オプションを設定
    const displayOptions = {
      showNames: !!config.showNames, // booleanに変換
      showCounts: config.showCounts !== undefined ? !!config.showCounts : true // booleanに変換、デフォルトtrue
    };
    setDisplayOptions(displayOptions);
    console.log('saveAndActivateSheet: 表示オプション設定完了');

    // 5. 最新のステータスをキャッシュを無視して取得
    const finalStatus = getStatus(true);
    console.log('saveAndActivateSheet: 統合処理完了');

    return {
      success: true,
      message: '設定の保存・適用が完了しました',
      status: finalStatus
    };

  } catch (error) {
    console.error('saveAndActivateSheetで致命的なエラー:', error.message, error.stack);
    // クライアントには分かりやすいエラーメッセージを返す
    throw new Error('設定の保存・適用中にサーバーエラーが発生しました: ' + error.message);
  }
}

/**
 * シート固有の設定を保存
 * @param {string} sheetName - シート名
 * @param {object} cfg - 設定オブジェクト
 */
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
 * 指定されたシートの列ヘッダーを自動的に推測し、マッピング結果を返す
 * @param {string} spreadsheetId - 対象のスプレッドシートID（現在未使用だが将来のために残す）
 * @param {string} sheetName - 対象のシート名
 * @returns {object} 推測されたヘッダーのマッピングオブジェクト
 */
function getGuessedHeaders(spreadsheetId, sheetName) {
  try {
    // 注意: spreadsheetIdは現在使われていません。autoMapSheetHeadersが内部で
    // セッションに基づいたアクティブなスプレッドシートを取得するためです。
    console.log('getGuessedHeaders: spreadsheetId=%s, sheetName=%s', spreadsheetId, sheetName);

    const guessedConfig = autoMapSheetHeaders(sheetName);
    if (!guessedConfig) {
      throw new Error('自動マッピングに失敗しました。対象シートにヘッダー行が存在するか確認してください。');
    }

    console.log('自動判定を実行しました: sheet=%s, result=%s', sheetName, JSON.stringify(guessedConfig));
    return guessedConfig;

  } catch (error) {
    console.error('getGuessedHeadersでエラー:', error.message, error.stack);
    // クライアントにスタックトレース全体ではなく、分かりやすいメッセージを返す
    throw new Error('列の自動判定に失敗しました: ' + error.message);
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
