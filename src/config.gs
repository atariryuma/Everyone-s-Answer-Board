/**
 * config.gs - 軽量化版
 * 新サービスアカウントアーキテクチャで必要最小限の関数のみ
 */

const CONFIG_SHEET_NAME = 'Config';

/**
 * 現在のユーザーのスプレッドシートを取得
 * View_AdminPanel.htmlから呼び出される
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
 * View_AdminPanel.htmlから呼び出される
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
 * 現在のスクリプト実行ユーザーのユニークIDを取得する。
 * まずプロパティストアを確認し、なければセッション情報から取得・保存する。
 * @returns {string} 現在のユーザーのユニークID
 */
function getUserId() {
  const props = PropertiesService.getUserProperties();
  let userId = props.getProperty('CURRENT_USER_ID');
  
  if (!userId) {
    // アクティブユーザーのメールアドレスからIDを生成（より安定した方法）
    const email = Session.getActiveUser().getEmail();
    if (email) {
      // メールアドレスをハッシュ化してIDとするなど、ユニークなIDを生成
      // ここでは簡易的にメールアドレスをそのまま使っていますが、
      // 必要に応じてUtilities.computeDigest等でハッシュ化してください。
      userId = email; 
      props.setProperty('CURRENT_USER_ID', userId);
    } else {
      // それでも取得できない場合はエラー
      throw new Error('ユーザーIDを取得できませんでした。');
    }
  }
  return userId;
}

// 他の関数も同様に、存在することを確認
function getUserInfo() {
  const userId = getUserId();
  // findUserById関数を呼び出すなど、ユーザー情報を取得するロジック
  return findUserById(userId); 
}

function getSheetHeaders(spreadsheetId, sheetName) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

/**
 * 指定されたシートの設定情報を取得する、新しい推奨関数。
 * 保存された設定があればそれを最優先し、なければ自動マッピングを試みる。
 * @param {string} sheetName - 設定を取得するシート名
 * @param {boolean} forceRefresh - キャッシュを無視して強制的に再取得するかどうか
 * @returns {object} 統一された設定オブジェクト
 */
function getConfig(sheetName, forceRefresh = false) {
  const userId = getUserId(); // ★修正点: 関数の最初にIDを取得
  const userCache = CacheService.getUserCache();
  const cacheKey = 'config_v3_' + userId + '_' + sheetName; 
  
  if (!forceRefresh) {
    const cached = userCache.get(cacheKey);
    if (cached) {
      console.log('設定キャッシュヒット: %s', cacheKey);
      return JSON.parse(cached);
    }
  }

  try {
    console.log('設定を取得中: sheetName=%s, userId=%s', sheetName, userId);
    const userInfo = getUserInfo(); // 依存関係を明確化
    const headers = getSheetHeaders(userInfo.spreadsheetId, sheetName);

    // 1. 返却する設定オブジェクトの器を準備
    let finalConfig = {
      sheetName: sheetName,
      opinionHeader: '',
      reasonHeader: '',
      nameHeader: '',
      classHeader: '',
      showNames: false,
      showCounts: true, 
      availableHeaders: headers || [],
      hasExistingConfig: false
    };

    // 2. 保存済みの設定があるか確認
    const configJson = userInfo.configJson ? JSON.parse(userInfo.configJson) : {};
    const sheetConfigKey = 'sheet_' + sheetName;
    const savedSheetConfig = configJson[sheetConfigKey];

    if (savedSheetConfig && Object.keys(savedSheetConfig).length > 0) {
      // 3.【最優先】保存済みの設定を適用する
      console.log('保存済み設定を適用します:', JSON.stringify(savedSheetConfig));
      finalConfig.hasExistingConfig = true;
      finalConfig.opinionHeader = savedSheetConfig.opinionHeader || '';
      finalConfig.reasonHeader = savedSheetConfig.reasonHeader || '';
      finalConfig.nameHeader = savedSheetConfig.nameHeader || '';
      finalConfig.classHeader = savedSheetConfig.classHeader || '';
      finalConfig.showNames = savedSheetConfig.showNames || false;
      finalConfig.showCounts = savedSheetConfig.showCounts !== undefined ? savedSheetConfig.showCounts : true;

    } else if (headers && headers.length > 0) {
      // 4.【保存設定がない場合のみ】新しい自動マッピングを実行
      console.log('保存済み設定がないため、自動マッピングを実行します。');
      const guessedConfig = autoMapHeaders(headers); 
      finalConfig.opinionHeader = guessedConfig.opinionHeader || '';
      finalConfig.reasonHeader = guessedConfig.reasonHeader || '';
      finalConfig.nameHeader = guessedConfig.nameHeader || '';
      finalConfig.classHeader = guessedConfig.classHeader || '';
    }

    // 5. 最終的な設定をキャッシュに保存
    userCache.put(cacheKey, JSON.stringify(finalConfig), 3600); // 1時間キャッシュ

    console.log('getConfig: 最終設定を返します: %s', JSON.stringify(finalConfig));
    return finalConfig;

  } catch (error) {
    console.error('getConfigでエラー:', error.message, error.stack);
    throw new Error('シート設定の取得中にエラーが発生しました: ' + error.message);
  }
}

/**
 * 列ヘッダーのリストから、各項目に最もふさわしい列名を推測する（改良版）
 * @param {Array<string>} headers - スプレッドシートのヘッダーリスト
 * @returns {object} 推測されたマッピング結果
 */
function autoMapHeaders(headers) {
  const mappingRules = {
    opinionHeader: ['今日のテーマ', 'あなたの考え', '意見', 'answer', 'response', 'opinion', '投稿'],
    reasonHeader:  ['そう考える理由', '理由', '詳細', '説明', 'reason'],
    nameHeader:    ['ニックネーム', '名前', '氏名', 'name'],
    classHeader:   ['あなたのクラス', 'クラス', '学年', '組', 'class']
  };

  const remainingHeaders = [...headers];
  const result = {};

  Object.keys(mappingRules).forEach(key => {
    const keywords = mappingRules[key];
    let bestMatch = '';

    for (const keyword of keywords) {
      const matchIndex = remainingHeaders.findIndex(header => 
        header && String(header).toLowerCase().includes(keyword.toLowerCase())
      );
      if (matchIndex !== -1) {
        bestMatch = remainingHeaders[matchIndex];
        remainingHeaders.splice(matchIndex, 1); 
        break;
      }
    }
    result[key] = bestMatch;
  });

  if (!result.opinionHeader) {
    const nonMetaHeaders = remainingHeaders.filter(h => {
      const hStr = String(h || '').toLowerCase();
      return !hStr.includes('タイムスタンプ') && !hStr.includes('メールアドレス');
    });
    if (nonMetaHeaders.length > 0) {
      result.opinionHeader = nonMetaHeaders[0];
    }
  }

  console.log('自動判定結果:', JSON.stringify(result));
  return result;
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

    return finalStatus;

  } catch (error) {
    console.error('saveAndActivateSheetで致命的なエラー:', error.message, error.stack);
    // クライアントには分かりやすいエラーメッセージを返す
    throw new Error('設定の保存・適用中にサーバーエラーが発生しました: ' + error.message);
  }
}

/**
 * 設定のみを保存する（公開はしない）
 * @param {string} sheetName - シート名
 * @param {object} config - 設定オブジェクト
 * @returns {object} 保存完了メッセージ
 */
function saveConfig(sheetName, config) {
  try {
    if (typeof sheetName !== 'string' || !sheetName) {
      throw new Error('無効なsheetNameです。シート名は必須です。');
    }
    if (typeof config !== 'object' || config === null) {
      throw new Error('無効なconfigオブジェクトです。設定オブジェクトは必須です。');
    }

    console.log('saveConfig開始: sheetName=%s', sheetName);

    // ユーザー情報を取得
    const userInfo = getUserInfo();
    if (!userInfo || !userInfo.spreadsheetId) {
      throw new Error('ユーザーのスプレッドシート情報が見つかりません。');
    }

    // 設定を保存
    saveSheetConfig(userInfo.spreadsheetId, sheetName, config);
    console.log('saveConfig: 設定保存完了');

    return {
      success: true,
      message: '設定が正常に保存されました。'
    };

  } catch (error) {
    console.error('saveConfigで致命的なエラー:', error.message, error.stack);
    throw new Error('設定の保存中にサーバーエラーが発生しました: ' + error.message);
  }
}

/**
 * シートを公開状態にする（設定は既に保存済みであることを前提）
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {object} 最新のステータス
 */
function activateSheet(spreadsheetId, sheetName) {
  try {
    if (typeof spreadsheetId !== 'string' || !spreadsheetId) {
      throw new Error('無効なspreadsheetIdです。スプレッドシートIDは必須です。');
    }
    if (typeof sheetName !== 'string' || !sheetName) {
      throw new Error('無効なsheetNameです。シート名は必須です。');
    }

    console.log('activateSheet開始: sheetName=%s', sheetName);

    // 関連キャッシュをクリア
    const props = PropertiesService.getUserProperties();
    const currentUserId = props.getProperty('CURRENT_USER_ID');
    if (currentUserId) {
      const userInfo = findUserById(currentUserId);
      if (userInfo) {
        invalidateUserCache(currentUserId, userInfo.adminEmail, spreadsheetId);
        console.log('activateSheet: キャッシュクリア完了');
      }
    }

    // シートをアクティブ化
    switchToSheet(spreadsheetId, sheetName);
    console.log('activateSheet: シート切り替え完了');

    // 最新のステータスを取得
    const finalStatus = getStatus(true);
    console.log('activateSheet: 公開処理完了');

    return finalStatus;

  } catch (error) {
    console.error('activateSheetで致命的なエラー:', error.message, error.stack);
    throw new Error('シートの公開中にサーバーエラーが発生しました: ' + error.message);
  }
}

/**
 * シート固有の設定を保存
 * @param {string} sheetName - シート名
 * @param {object} cfg - 設定オブジェクト
 */
/**
 * シートヘッダーを取得
 * View_AdminPanel.htmlから呼び出される
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
 * View_TeacherLanding.htmlから呼び出される
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
 * View_AdminPanel.htmlから呼び出される
 */
/**
 * アクティブシートをクリア（公開停止）
 * View_AdminPanel.htmlから呼び出される
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
 * View_AdminPanel.htmlから呼び出される
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
 * View_Board.htmlから呼び出される
 */
function checkAdmin() {
  debugLog('checkAdmin function called.');
  try {
    var activeUserEmail = Session.getActiveUser().getEmail();
    var props = PropertiesService.getScriptProperties();
    var dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
    var service = getSheetsService();
    var sheetName = DB_SHEET_CONFIG.SHEET_NAME;
    
    var response = batchGetSheetsData(service, dbId, [sheetName + '!A:H']);
    var data = (response && response.valueRanges && response.valueRanges[0] && response.valueRanges[0].values) ? response.valueRanges[0].values : [];
    
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
 * View_AdminPanel.htmlから呼び出される
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
 * View_Registration.htmlから呼び出される
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
 * View_Registration.htmlから呼び出される
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
 * 指定されたシートのヘッダーを自動でマッピングし、その結果を返す。
 * この関数は View_AdminPanel.html から直接呼び出されることを想定しています。
 * @param {string} sheetName - 対象のスプレッドシート名。
 * @returns {object} 推測されたヘッダーのマッピング結果。
 */
function getGuessedHeaders(sheetName) {
  try {
    const spreadsheet = getCurrentSpreadsheet();
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`シート '${sheetName}' が見つかりません。`);
    }
    
    // ヘッダー行（1行目）を取得
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // 既存の堅牢なロジックを呼び出して結果を返す
    const mappingResult = autoMapHeaders(headers);
    
    // 既存の設定も取得
    let existingConfig = null;
    try {
      existingConfig = getConfig(sheetName);
    } catch (configError) {
      console.warn('既存設定の取得に失敗しましたが、処理を続行します: ' + configError.toString());
    }
    
    // 結果にallHeadersと既存設定も含めて返す
    const result = {
      ...mappingResult,
      allHeaders: headers || [],
      existingConfig: existingConfig
    };
    
    console.log('自動判定を実行しました: sheet=%s, result=%s', sheetName, JSON.stringify(result));
    return result;

  } catch (e) {
    console.error('ヘッダーの自動判定に失敗しました: ' + e.toString());
    // エラーが発生した場合は、クライアント側で処理できるようエラー情報を返す
    return { error: e.message };
  }
}

/**
 * 指定されたシートのヘッダー情報と既存設定をまとめて取得します。
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {object} { allHeaders: Array<string>, guessedConfig: object, existingConfig: object }
 */
function getSheetDetails(spreadsheetId, sheetName) {
  try {
    if (!spreadsheetId || !sheetName) {
      throw new Error('spreadsheetIdとsheetNameは必須です');
    }

    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error('シートが見つかりません: ' + sheetName);
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] || [];
    const guessed = autoMapHeaders(headers);

    let existing = {};
    try {
      existing = getConfig(sheetName, true) || {};
    } catch (e) {
      console.warn('getConfig failed in getSheetDetails:', e.message);
    }

    return {
      allHeaders: headers,
      guessedConfig: guessed,
      existingConfig: existing
    };

  } catch (error) {
    console.error('getSheetDetails error:', error.message);
    throw new Error('シート情報の取得に失敗しました: ' + error.message);
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
