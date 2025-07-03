/**
 * @fileoverview StudyQuest - みんなの回答ボード (最適化・非推奨版)
 * 注意: このファイルは最適化のため分割されました。
 * 新しいファイル構造:
 * - Core.gs: 主要な業務ロジック
 * - AuthManager.gs: 認証管理
 * - DatabaseManager.gs: データベース操作
 * - CacheManager.gs: キャッシュ管理
 * - ReactionManager.gs: リアクション処理
 * - UrlManager.gs: URL管理
 * - DataProcessor.gs: データ処理
 * 
 * 後方互換性のため一部の関数は残されていますが、
 * 新しいクラスベースの実装を推奨します。
 */

// =================================================================
// グローバル設定
// =================================================================

// スクリプトプロパティに保存されるキー
var SCRIPT_PROPS_KEYS = {
  SERVICE_ACCOUNT_CREDS: 'SERVICE_ACCOUNT_CREDS',
  DATABASE_SPREADSHEET_ID: 'DATABASE_SPREADSHEET_ID'
};

// ユーザー専用フォルダ管理の設定
var USER_FOLDER_CONFIG = {
  ROOT_FOLDER_NAME: "StudyQuest - ユーザーデータ",
  FOLDER_NAME_PATTERN: "StudyQuest - {email} - ファイル"
};

// 中央データベースのシート設定
var DB_SHEET_CONFIG = {
  SHEET_NAME: 'Users',
  HEADERS: [
    'userId', 'adminEmail', 'spreadsheetId', 'spreadsheetUrl',
    'createdAt', 'configJson', 'lastAccessedAt', 'isActive'
  ]
};

// 監査ログシート設定
var LOG_SHEET_CONFIG = {
  SHEET_NAME: 'Logs',
  HEADERS: ['timestamp', 'userId', 'action', 'details']
};

// 回答ボードのスプレッドシートの列ヘッダー
var COLUMN_HEADERS = {
  TIMESTAMP: 'タイムスタンプ',
  EMAIL: 'メールアドレス',
  CLASS: 'クラス',
  OPINION: '回答',
  REASON: '理由',
  NAME: '名前',
  UNDERSTAND: 'なるほど！',
  LIKE: 'いいね！',
  CURIOUS: 'もっと知りたい！',
  HIGHLIGHT: 'ハイライト'
};

var REACTION_KEYS = ["UNDERSTAND", "LIKE", "CURIOUS"];
var EMAIL_REGEX = /^[^\n@]+@[^\n@]+\.[^\n@]+$/;
var DEBUG = true;

// アプリケーション設定用の追加定数
var APP_PROPERTIES = {
  PUBLISHED_SHEET: 'PUBLISHED_SHEET',
  DISPLAY_MODE: 'DISPLAY_MODE', // 'anonymous' or 'named'
  APP_PUBLISHED: 'APP_PUBLISHED'
};

var DISPLAY_MODES = {
  ANONYMOUS: 'anonymous',
  NAMED: 'named'
};

var SCORING_CONFIG = {
  LIKE_MULTIPLIER_FACTOR: 0.05, // 1いいね！ごとにスコアが5%増加
  RANDOM_SCORE_FACTOR: 0.1 // ランダム要素の重み
};

var ROSTER_CONFIG = {
  // デフォルト値のみ定義し、実際のシート名は getRosterSheetName() で取得
  SHEET_NAME: '名簿',
  EMAIL_COLUMN: 'メールアドレス',
  NAME_COLUMN: '名前',
  CLASS_COLUMN: 'クラス'
};

var TIME_CONSTANTS = {
  LOCK_WAIT_MS: 10000,
  CACHE_TTL: 300000 // 5分
};

function debugLog() {
  if (DEBUG && typeof console !== 'undefined' && console.log) {
    console.log.apply(console, arguments);
  }
}

// =================================================================
// セットアップ＆管理用関数
// =================================================================

// Deprecated: Use setupApplication() in Core.gs instead

// Deprecated: Use initializeDatabaseSheetOptimized() in DatabaseManager.gs instead

// =================================================================
// サービスアカウント認証 & Sheets API ラッパー
// =================================================================

// Deleted: Use getServiceAccountTokenCached() in AuthManager.gs instead

/**
 * Google Sheets API v4 のための認証済みサービスオブジェクトを返す。
 * @returns {object} Sheets APIサービス
 */
function getSheetsService() {
  var accessToken = getServiceAccountTokenCached();
  
  return {
    spreadsheets: {
      get: function(spreadsheetId) {
        var url = 'https://sheets.googleapis.com/v4/spreadsheets/' + spreadsheetId;
        var response = UrlFetchApp.fetch(url, {
          headers: { 'Authorization': 'Bearer ' + accessToken }
        });
        return JSON.parse(response.getContentText());
      },
      values: {
        get: function(spreadsheetId, range) {
          var url = 'https://sheets.googleapis.com/v4/spreadsheets/' + spreadsheetId + '/values/' + range;
          var response = UrlFetchApp.fetch(url, {
            headers: { 'Authorization': 'Bearer ' + accessToken }
          });
          return JSON.parse(response.getContentText());
        },
        update: function(spreadsheetId, range, body, params) {
          var url = 'https://sheets.googleapis.com/v4/spreadsheets/' + spreadsheetId + '/values/' + range + '?valueInputOption=' + params.valueInputOption;
          var response = UrlFetchApp.fetch(url, {
            method: 'put',
            contentType: 'application/json',
            headers: { 'Authorization': 'Bearer ' + accessToken },
            payload: JSON.stringify(body)
          });
          return JSON.parse(response.getContentText());
        },
        append: function(spreadsheetId, range, body, params) {
          var url = 'https://sheets.googleapis.com/v4/spreadsheets/' + spreadsheetId + '/values/' + range + ':append?valueInputOption=' + params.valueInputOption + '&insertDataOption=INSERT_ROWS';
          var response = UrlFetchApp.fetch(url, {
            method: 'post',
            contentType: 'application/json',
            headers: { 'Authorization': 'Bearer ' + accessToken },
            payload: JSON.stringify(body)
          });
          return JSON.parse(response.getContentText());
        },
        batchUpdate: function(spreadsheetId, body) {
          var url = 'https://sheets.googleapis.com/v4/spreadsheets/' + spreadsheetId + '/values:batchUpdate';
          var response = UrlFetchApp.fetch(url, {
            method: 'post',
            contentType: 'application/json',
            headers: { 'Authorization': 'Bearer ' + accessToken },
            payload: JSON.stringify(body)
          });
          return JSON.parse(response.getContentText());
        }
      },
      batchUpdate: function(spreadsheetId, body) {
        var url = 'https://sheets.googleapis.com/v4/spreadsheets/' + spreadsheetId + ':batchUpdate';
        var response = UrlFetchApp.fetch(url, {
          method: 'post',
          contentType: 'application/json',
          headers: { 'Authorization': 'Bearer ' + accessToken },
          payload: JSON.stringify(body)
        });
        return JSON.parse(response.getContentText());
      }
    }
  };
}

// =================================================================
// データベース操作関数 (サービスアカウント経由)
// =================================================================

// Deprecated: Use findUserByIdOptimized() in Core.gs instead

// Deprecated: Use findUserByEmailOptimized() in Core.gs instead

// Deprecated: Use createUserOptimized() in DatabaseManager.gs instead

// Deprecated: Use updateUserOptimized() in DatabaseManager.gs instead

// =================================================================
// メインロジック
// =================================================================

// Deleted: Use doGet() in UltraOptimizedCore.gs instead



// =================================================================
// ヘルパー関数
// =================================================================

// Deleted: Use getWebAppUrlCached() in UrlManager.gs instead

// =================================================================
// 共通ファクトリ関数 - 重複削減とコード効率化
// =================================================================

/**
 * 共通フォーム作成ファクトリ
 * createStudyQuestFormとquickStartSetupの重複を統一
 */
function createFormFactory(options) {
  var userEmail = options.userEmail;
  var userId = options.userId;
  var formTitle = options.formTitle || null;
  var formDescription = options.formDescription || null;
  var questions = options.questions || 'default';
  var linkedSpreadsheet = options.linkedSpreadsheet || null;
  
  try {
    var now = new Date();
    var dateTimeString = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
    var finalTitle = formTitle || 'StudyQuest - みんなの回答ボード - ' + userEmail.split('@')[0] + ' - ' + dateTimeString;
    
    var form = FormApp.create(finalTitle);
    
    // 共通設定
    form.setCollectEmail(true);
    form.setRequireLogin(true);
    form.setLimitOneResponsePerUser(true);
    form.setAllowResponseEdits(true);
    
    // 説明設定
    if (formDescription) {
      form.setDescription(formDescription);
    }
    
    // 質問設定
    addUnifiedQuestions(form, questions);
    
    // スプレッドシート連携
    var spreadsheetInfo;
    if (linkedSpreadsheet) {
      // 既存のスプレッドシートに連携
      form.setDestination(FormApp.DestinationType.SPREADSHEET, linkedSpreadsheet);
      spreadsheetInfo = {
        spreadsheetId: linkedSpreadsheet,
        sheetName: 'フォームの回答 1' // デフォルトシート名
      };
    } else {
      // 新しいスプレッドシートを作成
      spreadsheetInfo = createLinkedSpreadsheet(userEmail, form, dateTimeString);
    }
    
    return {
      formId: form.getId(),
      formUrl: form.getPublishedUrl(),
      editFormUrl: form.getEditUrl(),
      
      spreadsheetId: spreadsheetInfo.spreadsheetId,
      spreadsheetUrl: spreadsheetInfo.spreadsheetUrl,
      sheetName: spreadsheetInfo.sheetName || 'フォームの回答 1'
    };
    
  } catch (e) {
    console.error('フォーム作成ファクトリエラー: ' + e.message);
    throw new Error('フォームの作成に失敗しました: ' + e.message);
  }
}

/**
 * フォーム質問項目追加（後方互換性のため）
 * @deprecated addUnifiedQuestionsを使用してください
 */
// Deleted: Use addUnifiedQuestions() directly instead

/**
 * 統一された質問設定関数
 * デフォルト、シンプル、カスタマイズされた質問を統一的に管理
 */
function addUnifiedQuestions(form, questionType, customConfig) {
  questionType = questionType || 'default';
  customConfig = customConfig || {};
  
  var questions = getQuestionConfig(questionType, customConfig);
  
  questions.forEach(function(questionData) {
    var item;
    
    // 質問タイプに応じてアイテムを作成
    switch (questionData.type) {
      case 'text':
        item = form.addTextItem();
        break;
      case 'paragraph':
        item = form.addParagraphTextItem();
        break;
      case 'multipleChoice':
        item = form.addMultipleChoiceItem();
        if (questionData.choices) {
          item.setChoiceValues(questionData.choices);
        }
        break;
      case 'scale':
        item = form.addScaleItem();
        if (questionData.lowerBound && questionData.upperBound) {
          item.setBounds(questionData.lowerBound, questionData.upperBound);
        }
        break;
      default:
        item = form.addTextItem();
    }
    
    // 共通設定を適用
    item.setTitle(questionData.title);
    if (questionData.helpText) {
      item.setHelpText(questionData.helpText);
    }
    item.setRequired(questionData.required || false);
    
    // バリデーション設定（テキストアイテムの場合）
    if (questionData.type === 'text' && questionData.validation) {
      var validation = FormApp.createTextValidation()
        .requireTextMatchesPattern(questionData.validation.pattern)
        .setHelpText(questionData.validation.helpText)
        .build();
      item.setValidation(validation);
    }
  });
}

/**
 * 質問設定を取得
 */
function getQuestionConfig(questionType, customConfig) {
  switch (questionType) {
    case 'simple':
      return [
        {
          type: 'text',
          title: 'あなたのクラス',
          helpText: '例: 6-1, A組など',
          required: true
        },
        {
          type: 'text',
          title: 'あなたの名前',
          helpText: 'ニックネーム可（表示設定により匿名になる場合があります）',
          required: true
        },
        {
          type: 'paragraph',
          title: 'あなたの回答・意見',
          helpText: '質問に対するあなたの考えや意見を自由に書いてください',
          required: true
        },
        {
          type: 'paragraph',
          title: '理由・根拠',
          helpText: 'その回答になった理由や根拠があれば書いてください',
          required: false
        }
      ];
    
    case 'default':
    default:
      return [
        {
          type: 'text',
          title: 'クラス名',
          helpText: 'あなたのクラスを入力してください（例: 6-1, A組）',
          required: true,
          validation: {
            pattern: '^[A-Za-z0-9]+-[A-Za-z0-9]+$',
            helpText: '【重要】クラス名は決められた形式で入力してください。\n\n✅ 正しい例：\n• 6年1組 → 6-1\n• 5年2組 → 5-2\n• 中1年A組 → 1-A\n• 中3年B組 → 3-B\n\n❌ 間違いの例：6年1組、6-1組、６－１\n\n※ 半角英数字とハイフン（-）のみ使用可能です'
          }
        },
        {
          type: 'text',
          title: '名前',
          helpText: 'あなたの名前を入力してください',
          required: true
        },
        {
          type: 'paragraph',
          title: 'あなたの回答・意見',
          helpText: '質問に対するあなたの考えや意見を自由に書いてください',
          required: true
        },
        {
          type: 'paragraph',
          title: '理由・根拠',
          helpText: 'その回答になった理由や根拠があれば書いてください',
          required: false
        }
      ];
  }
}

/**
 * デフォルト質問設定（後方互換性のため）
 * @deprecated addUnifiedQuestionsを使用してください
 */
// Deleted: Use addUnifiedQuestions() directly instead

/**
 * シンプル質問設定（後方互換性のため）
 * @deprecated addUnifiedQuestionsを使用してください
 */
// Deleted: Use addUnifiedQuestions() directly instead

/**
 * 新しいスプレッドシート作成と連携
 */
function createLinkedSpreadsheet(userEmail, form, dateTimeString) {
  var spreadsheetTitle = 'StudyQuest - みんなの回答ボード - ' + userEmail.split('@')[0] + ' - ' + dateTimeString;
  var spreadsheet = SpreadsheetApp.create(spreadsheetTitle);
  
  form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheet.getId());
  
  // フォーム連携後に実際に作成されたシート名を取得
  var actualSheetName = 'フォームの回答 1'; // デフォルト値
  try {
    // 少し待ってからシート名を取得（Google Formsがシートを作成するまで待機）
    Utilities.sleep(2000);
    var sheets = spreadsheet.getSheets();
    if (sheets.length > 0) {
      // 最初のシートまたは回答シートを取得
      for (var i = 0; i < sheets.length; i++) {
        var currentSheetName = sheets[i].getName();
        if (currentSheetName.indexOf('フォームの回答') !== -1 || 
            currentSheetName.indexOf('Form Responses') !== -1) {
          actualSheetName = currentSheetName;
          break;
        }
      }
      // 回答シートが見つからない場合は最初のシートを使用
      if (actualSheetName === 'フォームの回答 1' && sheets.length > 0) {
        actualSheetName = sheets[0].getName();
      }
    }
    debugLog('実際に作成されたシート名: ' + actualSheetName);
  } catch (e) {
    console.warn('シート名取得エラー（デフォルトを使用）: ' + e.message);
  }
  
  return {
    spreadsheetId: spreadsheet.getId(),
    spreadsheetUrl: spreadsheet.getUrl(),
    sheetName: actualSheetName
  };
}

function createStudyQuestForm(userEmail, userId) {
  try {
    // 共通ファクトリを使用してフォーム作成
    var formResult = createFormFactory({
      userEmail: userEmail,
      userId: userId,
      questions: 'default',
      formDescription: 'このフォームは「みんなの回答ボード」で表示されます。デジタル・シティズンシップの観点から、オンライン空間での責任ある行動と建設的な対話を育むことを目的としています。回答内容は匿名で表示されます。'
    });
    
    // カスタマイズされた設定を追加
    var form = FormApp.openById(formResult.formId);
    
    // Email収集タイプの設定（可能な場合）
    try {
      if (typeof form.setEmailCollectionType === 'function') {
        form.setEmailCollectionType(FormApp.EmailCollectionType.VERIFIED);
      }
    } catch (undocumentedError) {
      // ignore
    }
    
    // 確認メッセージの設定
    // 注意: Googleフォームの確認メッセージはHTMLをサポートせず、自動リダイレクトはできません。
    // ユーザーを特定のページにリダイレクトするには、カスタムWebアプリをフォームの送信先として使用する必要があります。
    var appUrls = generateAppUrls(userId); // userIdを使ってURLを生成
    var confirmationMessage = appUrls.viewUrl
      ? '🎉 回答ありがとうございます！\n\nあなたの大切な意見が届きました。\nみんなの回答ボードで、お友達の色々な考えも見てみましょう。\n新しい発見があるかもしれませんね！\n\n' + appUrls.viewUrl
      : '🎉 回答ありがとうございます！\n\nあなたの大切な意見が届きました。';
    form.setConfirmationMessage(confirmationMessage);
    
    // バリデーションは統一された質問設定で処理済み
    
    // サービスアカウントをスプレッドシートに追加
    addServiceAccountToSpreadsheet(formResult.spreadsheetId);
    
    // リアクション列をスプレッドシートに追加
    addReactionColumnsToSpreadsheet(formResult.spreadsheetId, formResult.sheetName);
    
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
      debugLog('サービスアカウント (' + serviceAccountEmail + ') をスプレッドシートの編集者として追加しました。');
    }
  } catch (e) {
    console.error('サービスアカウントの追加に失敗: ' + e.message);
  }
}

/**
 * スプレッドシートにリアクション列を追加
 */
function addReactionColumnsToSpreadsheet(spreadsheetId, sheetName) {
  try {
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.getSheets()[0];
    
    var existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var headersToAdd = [];

    var reactionHeaders = [
      COLUMN_HEADERS.UNDERSTAND,
      COLUMN_HEADERS.LIKE,
      COLUMN_HEADERS.CURIOUS,
      COLUMN_HEADERS.HIGHLIGHT
    ];

    reactionHeaders.forEach(function(header) {
      if (existingHeaders.indexOf(header) === -1) {
        headersToAdd.push(header);
      }
    });
    
    if (headersToAdd.length > 0) {
      var startCol = existingHeaders.length + 1;
      sheet.getRange(1, startCol, 1, headersToAdd.length).setValues([headersToAdd]);
      
      var allHeadersRange = sheet.getRange(1, 1, 1, existingHeaders.length + headersToAdd.length);
      allHeadersRange.setFontWeight('bold').setBackground('#E3F2FD');
      
      try {
        sheet.autoResizeColumns(1, allHeadersRange.getNumColumns());
      } catch (e) {
        console.warn('Auto-resize failed:', e);
      }
      debugLog('リアクション列を追加しました: ' + sheetName);
    } else {
      debugLog('リアクション列は既に存在します: ' + sheetName);
    }
  } catch (e) {
    console.error('リアクション列追加エラー: ' + e.message);
  }
}

function isValidEmail(email) {
  if (!email || typeof email !== 'string') {
    return false;
  }
  return EMAIL_REGEX.test(email);
}

function getEmailDomain(email) {
  return (email || '').toString().split('@').pop().toLowerCase();
}

function safeGetUserEmail() {
  try {
    var email = Session.getActiveUser().getEmail();
    if (!email || typeof email !== 'string' || !isValidEmail(email)) {
      throw new Error('Invalid user email');
    }
    return email;
  } catch (e) {
    console.error('Failed to get user email:', e);
    throw new Error('認証が必要です。Googleアカウントでログインしてください。');
  }
}

function parseReactionString(val) {
  if (!val) return [];
  return val
    .toString()
    .split(',')
    .map(function(s) { return s.trim(); })
    .filter(Boolean);
}

// =================================================================
// HTML テンプレート用ヘルパー関数
// =================================================================

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getActiveFormInfo(userId) {
  var userInfo = findUserById(userId);
  if (!userInfo) {
    return { status: 'error', message: 'ユーザー情報が見つかりません' };
  }

  try {
    var configJson = JSON.parse(userInfo.configJson || '{}');
    return {
      status: 'success',
      formUrl: configJson.formUrl || '',
      editFormUrl: configJson.editFormUrl || '',
      spreadsheetUrl: userInfo.spreadsheetUrl || ''
    };
  } catch (e) {
    console.error('設定情報の取得に失敗: ' + e.message);
    return { status: 'error', message: '設定情報の取得に失敗しました' };
  }
}

function getResponsesData(userId, sheetName) {
  var userInfo = findUserById(userId);
  if (!userInfo) {
    return { status: 'error', message: 'ユーザー情報が見つかりません' };
  }

  try {
    var service = getSheetsService();
    var spreadsheetId = userInfo.spreadsheetId;
    var range = (sheetName || 'フォームの回答 1') + '!A:Z';
    
    var response = service.spreadsheets.values.get(spreadsheetId, range);
    var values = response.values || [];
    
    if (values.length === 0) {
      return { status: 'success', data: [], headers: [] };
    }

    return {
      status: 'success',
      data: values.slice(1), // ヘッダーを除く
      headers: values[0]
    };
  } catch (e) {
    console.error('回答データの取得に失敗: ' + e.message);
    return { status: 'error', message: '回答データの取得に失敗しました: ' + e.message };
  }
}

// =================================================================
// キャッシュ管理関数
// =================================================================

/**
 * ユーザーキャッシュ（メモリ内）
 */
var USER_INFO_CACHE = new Map();
var HEADER_CACHE = new Map();
var ROSTER_CACHE = new Map();
var CACHE_TIMESTAMPS = new Map();

/**
 * キャッシュされたユーザー情報を取得
 */
function getCachedUserInfo(userId) {
  var cacheKey = 'user_' + userId;
  var cached = USER_INFO_CACHE.get(cacheKey);
  var timestamp = CACHE_TIMESTAMPS.get(cacheKey);
  
  if (cached && timestamp && (Date.now() - timestamp) < TIME_CONSTANTS.CACHE_TTL) {
    debugLog('Memory cache hit for user: ' + userId);
    return cached;
  }
  
  return null;
}

/**
 * ユーザー情報をキャッシュに保存
 */
function setCachedUserInfo(userId, userInfo) {
  var cacheKey = 'user_' + userId;
  USER_INFO_CACHE.set(cacheKey, userInfo);
  CACHE_TIMESTAMPS.set(cacheKey, Date.now());
}

/**
 * ヘッダーインデックスをキャッシュから取得
 */
function getAndCacheHeaderIndices(spreadsheetId, sheetName, headerRow) {
  var cacheKey = spreadsheetId + '_' + sheetName + '_headers';
  var cached = HEADER_CACHE.get(cacheKey);
  var timestamp = CACHE_TIMESTAMPS.get(cacheKey);
  
  if (cached && timestamp && (Date.now() - timestamp) < TIME_CONSTANTS.CACHE_TTL) {
    debugLog('Header cache hit for: ' + sheetName);
    return cached;
  }
  
  try {
    var service = getSheetsService();
    var range = sheetName + '!' + (headerRow || 1) + ':' + (headerRow || 1);
    var response = service.spreadsheets.values.get(spreadsheetId, range);
    var headers = response.values ? response.values[0] : [];
    
    var indices = findHeaderIndices(headers, Object.values(COLUMN_HEADERS));
    
    HEADER_CACHE.set(cacheKey, indices);
    CACHE_TIMESTAMPS.set(cacheKey, Date.now());
    
    return indices;
  } catch (e) {
    console.error('ヘッダー取得エラー: ' + e.message);
    return {};
  }
}

/**
 * ヘッダーインデックスを検索
 */
function findHeaderIndices(headers, requiredHeaders) {
  var indices = {};
  
  requiredHeaders.forEach(function(header) {
    var index = headers.indexOf(header);
    if (index !== -1) {
      indices[header] = index;
    }
  });
  
  return indices;
}

/**
 * 名簿データをキャッシュから取得（名前とクラスのマッピング）
 */
function getRosterMap(spreadsheetId) {
  var cacheKey = spreadsheetId + '_roster';
  var cached = ROSTER_CACHE.get(cacheKey);
  var timestamp = CACHE_TIMESTAMPS.get(cacheKey);
  
  if (cached && timestamp && (Date.now() - timestamp) < TIME_CONSTANTS.CACHE_TTL) {
    debugLog('Roster cache hit');
    return cached;
  }
  
  try {
    var service = getSheetsService();
    var range = getRosterSheetName() + '!A:Z';
    var response = service.spreadsheets.values.get(spreadsheetId, range);
    var values = response.values || [];
    
    if (values.length === 0) {
      return {};
    }
    
    var headers = values[0];
    var emailIndex = headers.indexOf(ROSTER_CONFIG.EMAIL_COLUMN);
    var nameIndex = headers.indexOf(ROSTER_CONFIG.NAME_COLUMN);
    var classIndex = headers.indexOf(ROSTER_CONFIG.CLASS_COLUMN);
    
    var rosterMap = {};
    
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      if (emailIndex !== -1 && row[emailIndex]) {
        rosterMap[row[emailIndex]] = {
          name: nameIndex !== -1 ? row[nameIndex] : '',
          class: classIndex !== -1 ? row[classIndex] : ''
        };
      }
    }
    
    ROSTER_CACHE.set(cacheKey, rosterMap);
    CACHE_TIMESTAMPS.set(cacheKey, Date.now());
    
    return rosterMap;
  } catch (e) {
    console.error('名簿データ取得エラー: ' + e.message);
    return {};
  }
}

/**
 * キャッシュをクリア
 */
function clearAllCaches() {
  USER_INFO_CACHE.clear();
  HEADER_CACHE.clear();
  ROSTER_CACHE.clear();
  CACHE_TIMESTAMPS.clear();
  debugLog('全キャッシュをクリアしました');
}

function clearRosterCache() {
  var keysToDelete = [];
  ROSTER_CACHE.forEach(function(value, key) {
    if (key.includes('_roster')) {
      keysToDelete.push(key);
    }
  });
  
  keysToDelete.forEach(function(key) {
    ROSTER_CACHE.delete(key);
    CACHE_TIMESTAMPS.delete(key);
  });
  
  debugLog('名簿キャッシュをクリアしました');
}

// =================================================================
// データ処理とソート機能
// =================================================================

/**
 * 公開されたシートのデータを取得（Page.htmlから呼び出される）
 */
function getPublishedSheetDataLegacy(classFilter, sortMode) {
  try {
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    
    if (!currentUserId) {
      throw new Error('ユーザーコンテキストが設定されていません');
    }
    
    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    
    // 設定から公開シートを取得
    var configJson = JSON.parse(userInfo.configJson || '{}');
    var publishedSheet = configJson.publishedSheet || 'フォームの回答 1';
    
    return getSheetData(currentUserId, publishedSheet, classFilter, sortMode);
  } catch (e) {
    console.error('公開シートデータ取得エラー: ' + e.message);
    return {
      status: 'error',
      message: 'データの取得に失敗しました: ' + e.message,
      data: [],
      headers: []
    };
  }
}

/**
 * 指定されたシートのデータを取得し、フィルタリング・ソートを適用
 */
function getSheetData(userId, sheetName, classFilter, sortMode) {
  try {
    var userInfo = findUserById(userId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    
    var spreadsheetId = userInfo.spreadsheetId;
    var service = getSheetsService();
    var range = sheetName + '!A:Z';
    
    var response = service.spreadsheets.values.get(spreadsheetId, range);
    var values = response.values || [];
    
    if (values.length === 0) {
      return {
        status: 'success',
        data: [],
        headers: [],
        totalCount: 0
      };
    }
    
    var headers = values[0];
    var dataRows = values.slice(1);
    
    // ヘッダーインデックスを取得
    var headerIndices = getAndCacheHeaderIndices(spreadsheetId, sheetName);
    
    // 名簿データを取得（名前表示用）
    var rosterMap = getRosterMap(spreadsheetId);
    
    // 表示モードを取得
    var configJson = JSON.parse(userInfo.configJson || '{}');
    var displayMode = configJson.displayMode || DISPLAY_MODES.ANONYMOUS;
    
    // データを処理
    var processedData = dataRows.map(function(row, index) {
      return processRowData(row, headers, headerIndices, rosterMap, displayMode, index + 2); // +2 for header row and 1-based indexing
    });
    
    // クラスフィルタを適用
    if (classFilter && classFilter !== 'all') {
      var classIndex = headerIndices[COLUMN_HEADERS.CLASS];
      if (classIndex !== undefined) {
        processedData = processedData.filter(function(row) {
          return row.originalData[classIndex] === classFilter;
        });
      }
    }
    
    // ソートを適用
    processedData = applySortMode(processedData, sortMode || 'newest');
    
    return {
      status: 'success',
      data: processedData,
      headers: headers,
      totalCount: processedData.length,
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
  
  // 各リアクションのカウントを計算
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
  
  // ハイライト状態をチェック
  var highlightIndex = headerIndices[COLUMN_HEADERS.HIGHLIGHT];
  if (highlightIndex !== undefined && row[highlightIndex]) {
    processedRow.isHighlighted = row[highlightIndex].toString().toLowerCase() === 'true';
  }
  
  // スコアを計算
  processedRow.score = calculateRowScore(processedRow);
  
  // 名前の表示処理
  var emailIndex = headerIndices[COLUMN_HEADERS.EMAIL];
  if (emailIndex !== undefined && row[emailIndex] && displayMode === DISPLAY_MODES.NAMED) {
    var email = row[emailIndex];
    var rosterInfo = rosterMap[email];
    if (rosterInfo && rosterInfo.name) {
      // 名簿に名前がある場合は名前を表示
      var nameIndex = headerIndices[COLUMN_HEADERS.NAME];
      if (nameIndex !== undefined) {
        processedRow.displayName = rosterInfo.name;
      }
    }
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
      return data.reverse(); // 最新が上に来るように
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

// =================================================================
// 管理機能
// =================================================================

/**
 * スプレッドシートのメニューを作成（onOpen）
 */
function onOpen() {
  try {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('📋 みんなの回答ボード')
      .addItem('📊 管理パネルを開く', 'showAdminSidebar')
      .addSeparator()
      .addItem('🔄 キャッシュをクリア', 'clearAllCaches')
      .addItem('📝 名簿キャッシュをクリア', 'clearRosterCache')
      .addToUi();
  } catch (e) {
    console.error('メニュー作成エラー: ' + e.message);
  }
}

/**
 * 管理サイドバーを表示
 */
function showAdminSidebar() {
  try {
    var template = HtmlService.createTemplateFromFile('AdminSidebar');
    var html = template.evaluate()
      .setTitle('みんなの回答ボード - 管理パネル')
      .setWidth(400);
    
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) {
    console.error('管理サイドバー表示エラー: ' + e.message);
    SpreadsheetApp.getUi().alert('管理パネルの表示に失敗しました: ' + e.message);
  }
}

/**
 * 統合されたアプリケーション設定を取得
 * 管理者設定、ステータス情報、URL情報を統一的に提供
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
    
    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    
    var configJson = JSON.parse(userInfo.configJson || '{}');
    var sheets = getSheets(currentUserId);
    var appUrls = generateAppUrls(currentUserId); // 拡張されたURL生成を使用
    
    return {
      status: 'success',
      userId: currentUserId,
      adminEmail: userInfo.adminEmail,
      publishedSheet: configJson.publishedSheet || '',
      displayMode: configJson.displayMode || DISPLAY_MODES.ANONYMOUS,
      isPublished: configJson.appPublished || false,
      availableSheets: sheets,
      spreadsheetUrl: userInfo.spreadsheetUrl,
      formUrl: configJson.formUrl || '',
      editFormUrl: configJson.editFormUrl || '',
      webAppUrl: appUrls.webAppUrl,
      adminUrl: appUrls.adminUrl,
      viewUrl: appUrls.viewUrl,
      activeSheetName: configJson.publishedSheet || '',
      appUrls: appUrls // 全URL情報をオブジェクトとして提供
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
 * 管理者設定を取得（後方互換性のため）
 * @deprecated getAppConfigを使用してください
 */
// Deleted: Use getAppConfig() directly instead

/**
 * 管理画面用のステータス情報を取得（レガシー実装）
 * @deprecated Core.gsのgetStatus()を使用してください
 */
function getStatusLegacy() {
  return getAppConfig();
}

/**
 * 利用可能なシート一覧を取得
 */
function getSheets(userId) {
  try {
    var userInfo = findUserById(userId);
    if (!userInfo) {
      return [];
    }
    
    var service = getSheetsService();
    var spreadsheet = service.spreadsheets.get(userInfo.spreadsheetId);
    
    return spreadsheet.sheets.map(function(sheet) {
      return {
        name: sheet.properties.title,
        id: sheet.properties.sheetId
      };
    });
  } catch (e) {
    console.error('シート一覧取得エラー: ' + e.message);
    return [];
  }
}

/**
 * アプリを公開
 */
function publishApp(sheetName) {
  try {
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    
    if (!currentUserId) {
      throw new Error('ユーザーコンテキストが設定されていません');
    }
    
    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    
    var configJson = JSON.parse(userInfo.configJson || '{}');
    configJson.publishedSheet = sheetName;
    configJson.appPublished = true;
    configJson.publishedAt = new Date().toISOString();
    
    updateUserInDb(currentUserId, {
      configJson: JSON.stringify(configJson)
    });

    auditLog('PUBLISH', currentUserId, { sheet: sheetName });
    
    debugLog('アプリを公開しました: ' + sheetName);
    
    return {
      status: 'success',
      message: 'アプリが公開されました',
      publishedSheet: sheetName
    };
  } catch (e) {
    console.error('アプリ公開エラー: ' + e.message);
    return {
      status: 'error',
      message: 'アプリの公開に失敗しました: ' + e.message
    };
  }
}

/**
 * アプリの公開を停止
 */
function unpublishApp() {
  try {
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    
    if (!currentUserId) {
      throw new Error('ユーザーコンテキストが設定されていません');
    }
    
    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    
    var configJson = JSON.parse(userInfo.configJson || '{}');
    configJson.appPublished = false;
    configJson.publishedAt = '';
    
    updateUserInDb(currentUserId, {
      configJson: JSON.stringify(configJson)
    });

    auditLog('UNPUBLISH', currentUserId, {});
    
    debugLog('アプリの公開を停止しました');
    
    return {
      status: 'success',
      message: 'アプリの公開を停止しました'
    };
  } catch (e) {
    console.error('アプリ公開停止エラー: ' + e.message);
    return {
      status: 'error',
      message: 'アプリの公開停止に失敗しました: ' + e.message
    };
  }
}

/**
 * 表示モードを保存
 */
function saveDisplayMode(mode) {
  try {
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    
    if (!currentUserId) {
      throw new Error('ユーザーコンテキストが設定されていません');
    }
    
    if (!Object.values(DISPLAY_MODES).includes(mode)) {
      throw new Error('無効な表示モードです: ' + mode);
    }
    
    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    
    var configJson = JSON.parse(userInfo.configJson || '{}');
    configJson.displayMode = mode;
    
    updateUserInDb(currentUserId, {
      configJson: JSON.stringify(configJson)
    });
    
    debugLog('表示モードを保存しました: ' + mode);
    
    return {
      status: 'success',
      message: '表示モードを保存しました',
      displayMode: mode
    };
  } catch (e) {
    console.error('表示モード保存エラー: ' + e.message);
    return {
      status: 'error',
      message: '表示モードの保存に失敗しました: ' + e.message
    };
  }
}

/**
 * アプリ設定を取得
 */
function getAppSettings() {
  try {
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    
    if (!currentUserId) {
      return {
        status: 'error',
        message: 'ユーザーコンテキストが設定されていません'
      };
    }
    
    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      return {
        status: 'error',
        message: 'ユーザー情報が見つかりません'
      };
    }
    
    var configJson = JSON.parse(userInfo.configJson || '{}');
    
    return {
      status: 'success',
      publishedSheet: configJson.publishedSheet || '',
      displayMode: configJson.displayMode || DISPLAY_MODES.ANONYMOUS,
      isPublished: configJson.appPublished || false
    };
  } catch (e) {
    console.error('アプリ設定取得エラー: ' + e.message);
    return {
      status: 'error',
      message: 'アプリ設定の取得に失敗しました: ' + e.message
    };
  }
}

/**
 * ハイライト状態を切り替え
 */
function toggleHighlight(rowIndex) {
  try {
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    
    if (!currentUserId) {
      throw new Error('ユーザーコンテキストが設定されていません');
    }
    
    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    
    var configJson = JSON.parse(userInfo.configJson || '{}');
    var sheetName = configJson.publishedSheet || 'フォームの回答 1';
    
    var service = getSheetsService();
    var spreadsheetId = userInfo.spreadsheetId;
    
    // ヘッダーを取得してハイライト列を特定
    var headerResponse = service.spreadsheets.values.get(spreadsheetId, sheetName + '!1:1');
    var headers = headerResponse.values ? headerResponse.values[0] : [];
    var highlightIndex = headers.indexOf(COLUMN_HEADERS.HIGHLIGHT);
    
    if (highlightIndex === -1) {
      throw new Error('ハイライト列が見つかりません');
    }
    
    // 現在の値を取得
    var cellRange = sheetName + '!' + String.fromCharCode(65 + highlightIndex) + rowIndex;
    var currentResponse = service.spreadsheets.values.get(spreadsheetId, cellRange);
    var currentValue = currentResponse.values && currentResponse.values[0] ? currentResponse.values[0][0] : '';
    
    // 値を切り替え
    var newValue = (currentValue.toString().toLowerCase() === 'true') ? 'false' : 'true';
    
    service.spreadsheets.values.update(
      spreadsheetId,
      cellRange,
      { values: [[newValue]] },
      { valueInputOption: 'RAW' }
    );
    
    debugLog('ハイライト状態を切り替えました: 行' + rowIndex + ' → ' + newValue);
    
    return {
      status: 'success',
      message: 'ハイライト状態を更新しました',
      isHighlighted: newValue === 'true',
      rowIndex: rowIndex
    };
  } catch (e) {
    console.error('ハイライト切り替えエラー: ' + e.message);
    return {
      status: 'error',
      message: 'ハイライトの切り替えに失敗しました: ' + e.message
    };
  }
}

/**
 * WebアプリのURLを取得
 */
function getWebAppUrl() {
  try {
    var url = ScriptApp.getService().getUrl();
    if (!url) {
      console.warn('ScriptApp.getService().getUrl()がnullまたは空文字を返しました');
      return '';
    }
    
    // URLの正規化（末尾のスラッシュを削除）
    if (url.endsWith('/')) {
      url = url.slice(0, -1);
    }
    
    return url;
  } catch (e) {
    console.error('WebアプリURL取得エラー: ' + e.message);
    // フォールバック処理（可能な限りURLを推測）
    try {
      var scriptId = ScriptApp.getScriptId();
      if (scriptId) {
        return 'https://script.google.com/macros/s/' + scriptId + '/exec';
      }
    } catch (fallbackError) {
      console.error('フォールバックURL生成エラー: ' + fallbackError.message);
    }
    return '';
  }
}

/**
 * アプリケーション用のURL群を生成
 * 基本WebアプリURL、管理画面URL、ビューURLなどを統一的に生成
 */
function generateAppUrls(userId) {
  try {
    var webAppUrl = getWebAppUrl();
    
    if (!webAppUrl) {
      return {
        webAppUrl: '',
        adminUrl: '',
        viewUrl: '',
        setupUrl: '',
        status: 'error',
        message: 'WebアプリURLが取得できませんでした'
      };
    }
    
    var adminUrl = webAppUrl + '?userId=' + userId + '&mode=admin';
    var viewUrl = webAppUrl + '?userId=' + userId;
    var setupUrl = webAppUrl + '?setup=true';
    
    return {
      webAppUrl: webAppUrl,
      adminUrl: adminUrl,
      viewUrl: viewUrl,
      setupUrl: setupUrl,
      status: 'success'
    };
  } catch (e) {
    console.error('URL生成エラー: ' + e.message);
    return {
      webAppUrl: '',
      adminUrl: '',
      viewUrl: '',
      setupUrl: '',
      status: 'error',
      message: 'URLの生成に失敗しました: ' + e.message
    };
  }
}

/**
 * デプロイ・ユーザー・ドメイン情報を取得（AdminPanel.htmlとRegistration.htmlから呼び出される）
 */
function getDeployUserDomainInfo() {
  try {
    var webAppUrl = getWebAppUrl();
    var activeUser = Session.getActiveUser().getEmail();
    var currentDomain = getEmailDomain(activeUser);
    
    // デプロイドメインを特定（URLから推測またはハードコード）
    var deployDomain = 'naha-okinawa.ed.jp'; // 実際のデプロイドメインに合わせて調整
    if (webAppUrl && webAppUrl.includes('/a/macros/')) {
      // Google Workspace環境の場合、URLからドメインを抽出
      var match = webAppUrl.match(/\/a\/macros\/([^\/]+)\//);
      if (match && match[1]) {
        deployDomain = match[1];
      }
    }
    
    // ドメイン一致の確認
    var isDomainMatch = currentDomain === deployDomain;
    
    var props = PropertiesService.getUserProperties();
    var currentUserId = props.getProperty('CURRENT_USER_ID');
    var userInfo = null;
    
    if (currentUserId) {
      userInfo = findUserById(currentUserId);
    }
    
    return {
      status: 'success',
      webAppUrl: webAppUrl,
      activeUser: activeUser,
      domain: currentDomain,
      currentDomain: currentDomain,
      deployDomain: deployDomain,
      isDomainMatch: isDomainMatch,
      userId: currentUserId,
      userInfo: userInfo,
      deploymentTimestamp: new Date().toISOString()
    };
  } catch (e) {
    console.error('デプロイ情報取得エラー: ' + e.message);
    return {
      status: 'error',
      message: 'デプロイ情報の取得に失敗しました: ' + e.message
    };
  }
}

// =================================================================
// 互換性とエラーハンドリング
// =================================================================


// Deleted: Function was unused


/**
 * クイックスタートセットアップ
 * Registration.htmlから呼び出される
 */
function quickStartSetup(userId) {
  try {
    debugLog('クイックスタートセットアップ開始: ' + userId);
    
    // ユーザー情報の取得
    var userInfo = findUserById(userId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    
    var configJson = JSON.parse(userInfo.configJson || '{}');
    var userEmail = userInfo.adminEmail;
    var spreadsheetId = userInfo.spreadsheetId;
    var spreadsheetUrl = userInfo.spreadsheetUrl;

    // 1. Googleフォームの作成（既に作成済みの場合はスキップ）
    var formUrl = configJson.formUrl;
    var editFormUrl = configJson.editFormUrl;
    var sheetName = 'フォームの回答 1'; // Default sheet name for form responses

    if (!formUrl) {
      var formAndSsInfo = createStudyQuestForm(userEmail, userId);
      formUrl = formAndSsInfo.formUrl;
      editFormUrl = formAndSsInfo.editFormUrl;
      spreadsheetId = formAndSsInfo.spreadsheetId;
      spreadsheetUrl = formAndSsInfo.spreadsheetUrl;
      sheetName = formAndSsInfo.sheetName;

      // Update user info with new form/spreadsheet details
      updateUserInDb(userId, {
        spreadsheetId: spreadsheetId,
        spreadsheetUrl: formAndSsInfo.spreadsheetUrl,
        configJson: JSON.stringify({
          ...configJson,
          formUrl: formUrl,
          editFormUrl: editFormUrl,
          publishedSheet: sheetName, // Set initial published sheet
          appPublished: true // Publish app on quick start
        })
      });
    }

    // 2. Configシートの作成と初期化
    createAndInitializeConfigSheet(spreadsheetId);

    var appUrls = generateAppUrlsOptimized(userId);
    debugLog('クイックスタートセットアップ完了: ' + userId);
    return {
      status: 'success',
      message: 'クイックスタートセットアップが完了しました。',
      adminUrl: appUrls.adminUrl,
      viewUrl: appUrls.viewUrl,
      formUrl: formUrl,
      spreadsheetUrl: spreadsheetUrl
    };

  } catch (e) {
    console.error('クイックスタートセットアップエラー: ' + e.message);
    return { status: 'error', message: 'クイックスタートセットアップに失敗しました: ' + e.message };
  }
}

/**
 * ユーザー認証の状態を確認
 * Registration.htmlから呼び出される
 */
function verifyUserAuthentication() {
  try {
    var activeUser = Session.getActiveUser();
    var userEmail = activeUser.getEmail();
    
    if (!userEmail) {
      return {
        authenticated: false,
        message: 'Googleアカウントでログインしてください'
      };
    }
    
    var domain = getEmailDomain(userEmail);
    
    return {
      authenticated: true,
      email: userEmail,
      domain: domain || 'unknown'
    };
    
  } catch (e) {
    console.error('認証確認エラー: ' + e.message);
    return {
      authenticated: false,
      message: '認証の確認中にエラーが発生しました'
    };
  }
}

/**
 * 現在のユーザーの既存回答ボード情報を取得
 * Registration.htmlから呼び出される
 */
function getExistingBoard() {
  try {
    // サービスアカウント設定の確認
    var props = PropertiesService.getScriptProperties();
    var serviceAccountCreds = props.getProperty(SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS);
    var databaseSpreadsheetId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
    
    if (!serviceAccountCreds || !databaseSpreadsheetId) {
      // セットアップが未完了の場合
      return {
        status: 'setup_required',
        message: 'サービスアカウントのセットアップが必要です'
      };
    }
    
    // 現在のユーザーのメールアドレスを取得
    var activeUser = Session.getActiveUser();
    var userEmail = activeUser.getEmail();
    
    if (!userEmail) {
      return {
        status: 'auth_required',
        message: 'Googleアカウントでログインしてください'
      };
    }
    
    // 既存ユーザーの検索
    var existingUser = findUserByEmail(userEmail);
    
    if (existingUser) {
      // 統一されたURL群を生成
      var appUrls = generateAppUrls(existingUser.userId);
      
      return {
        status: 'existing_user',
        userId: existingUser.userId,
        userInfo: existingUser,
        adminUrl: appUrls.adminUrl,
        viewUrl: appUrls.viewUrl
      };
    } else {
      return {
        status: 'new_user',
        message: '新規ユーザーとして登録できます'
      };
    }
    
  } catch (e) {
    console.error('既存ボード確認エラー: ' + e.message);
    return {
      status: 'error',
      message: 'アカウント確認中にエラーが発生しました'
    };
  }
}

/**
 * アクションをログシートに記録する
 * @param {string} action - 実行した操作
 * @param {string} userId - 対象ユーザーID
 * @param {object} [details] - 追加情報
 */
function auditLog(action, userId, details) {
  try {
    var props = PropertiesService.getScriptProperties();
    var dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
    if (!dbId) return;

    var service = getSheetsService();
    var range = LOG_SHEET_CONFIG.SHEET_NAME + '!A:D';
    service.spreadsheets.values.append(
      dbId,
      range,
      { values: [[new Date().toISOString(), userId, action, JSON.stringify(details || {})]] },
      { valueInputOption: 'USER_ENTERED' }
    );
  } catch (e) {
    console.error('auditLog error: ' + e.message);
  }
}

/**
 * 公開から6時間経過していれば自動的に非公開にする
 * @param {Object} userInfo
 * @param {string} userId
 * @returns {HtmlOutput|null}
 */
function checkAutoUnpublish(userInfo, userId) {
  try {
    var configJson = JSON.parse(userInfo.configJson || '{}');
    if (configJson.appPublished && configJson.publishedAt) {
      var published = new Date(configJson.publishedAt);
      if (Date.now() - published.getTime() > 6 * 60 * 60 * 1000) {
        configJson.appPublished = false;
        configJson.publishedAt = '';
        updateUserInDb(userId, { configJson: JSON.stringify(configJson) });
        auditLog('AUTO_UNPUBLISH', userId, { reason: 'timeout' });

        var template = HtmlService.createTemplateFromFile('Unpublished');
        template.message = 'この回答ボードは、公開から6時間が経過したため、安全のため自動的に非公開になりました。再度利用する場合は、管理者にご連絡ください。';
        return template.evaluate().setTitle('公開期間が終了しました');
      }
    }
  } catch (e) {
    console.error('auto unpublish check failed: ' + e.message);
  }
  return null;
}

/**
 * 回答ボードのデータを強制的に再読み込みするための関数。
 * 管理パネルから呼び出され、キャッシュをクリアし、クライアントサイドに再読み込みを促す。
 * @returns {object} 成功ステータスとメッセージ
 */
function refreshBoardData() {
  try {
    clearAllCache(); // 全キャッシュをクリア
    debugLog('回答ボードのデータ強制再読み込みをトリガーしました。');
    return { status: 'success', message: '回答ボードのデータを更新しました。' };
  } catch (e) {
    console.error('回答ボードのデータ再読み込みエラー: ' + e.message);
    return { status: 'error', message: '回答ボードのデータ更新に失敗しました: ' + e.message };
  }
}