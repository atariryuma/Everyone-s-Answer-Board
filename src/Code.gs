/**
 * @fileoverview StudyQuest - みんなの回答ボード
 * ★【管理者メニュー実装版】スプレッドシートのカスタムメニューからアプリを管理する機能を実装。
 */

// =================================================================
// 定数定義
// =================================================================

// マルチテナント用ユーザーデータベース設定
const USER_DB_CONFIG = {
  SHEET_NAME: 'Users',
  HEADERS: [
    'userId',
    'adminEmail',
    'spreadsheetId', 
    'spreadsheetUrl',
    'createdAt',
    'accessToken',
    'configJson',
    'lastAccessedAt',
    'isActive'
  ]
};
const COLUMN_HEADERS = {
  EMAIL: 'メールアドレス',
  CLASS: 'クラスを選択してください。',
  OPINION: 'これまでの学んだことや、経験したことから、根からとり入れた水は、植物のからだのどこを通るのか予想しましょう。',
  REASON: '予想したわけを書きましょう。',
  UNDERSTAND: 'なるほど！',
  LIKE: 'いいね！',
  CURIOUS: 'もっと知りたい！',
  HIGHLIGHT: 'ハイライト',
  TIMESTAMP: 'タイムスタンプ',
  NAME: '名前'
};
const ROSTER_CONFIG = {
  SHEET_NAME: 'roster',
  PROPERTY_NAME: 'ROSTER_SHEET_NAME',
  CACHE_KEY: 'roster_name_map_v3',
  HEADER_LAST_NAME: '姓',
  HEADER_FIRST_NAME: '名',
  HEADER_NICKNAME: 'ニックネーム',
  HEADER_EMAIL: 'Googleアカウント'
};
const SCORING_CONFIG = {
  LIKE_MULTIPLIER_FACTOR: 0.05 // 1いいね！ごとにスコアが5%増加
};
const APP_PROPERTIES = {
  ACTIVE_SHEET: 'ACTIVE_SHEET_NAME',
  IS_PUBLISHED: 'IS_PUBLISHED',
  SHOW_DETAILS: 'SHOW_DETAILS'
};

const TIME_CONSTANTS = {
  LOCK_WAIT_MS: 10000
};

const REACTION_KEYS = ["UNDERSTAND","LIKE","CURIOUS"];
var getConfig;
var handleError;
if (typeof global !== 'undefined' && global.getConfig) {
  getConfig = global.getConfig;
}

function getCurrentSpreadsheet() {
  const props = PropertiesService.getUserProperties();
  const spreadsheetId = props.getProperty('CURRENT_SPREADSHEET_ID');
  
  if (!spreadsheetId) {
    // 従来の動作（単一テナント時）
    return SpreadsheetApp.getActiveSpreadsheet();
  }
  
  // マルチテナント時は指定されたスプレッドシートを使用
  return SpreadsheetApp.openById(spreadsheetId);
}

function safeGetUserEmail() {
  try {
    return Session.getActiveUser().getEmail();
  } catch (e) {
    return '';
  }
}

function getAdminEmails() {
  const props = PropertiesService.getScriptProperties();
  const str = props.getProperty('ADMIN_EMAILS') || '';
  return str.split(',').map(e => e.trim()).filter(Boolean);
}

function isUserAdmin(email) {
  const userEmail = email || safeGetUserEmail();
  return getAdminEmails().includes(userEmail);
}

function checkAdmin() {
  const userProps = PropertiesService.getUserProperties();
  const userId = userProps.getProperty('CURRENT_USER_ID');
  
  if (userId) {
    // マルチテナントモード
    const userInfo = getUserInfo(userId);
    return userInfo && userInfo.adminEmail === safeGetUserEmail();
  }
  
  // 従来のモード
  return isUserAdmin();
}

function parseReactionString(val) {
  if (!val) return [];
  return val
    .toString()
    .split(',')
    .map(s => s.trim())
    .filter(Boolean);
}

function hashTimestamp(ts) {
  var str = String(ts);
  var hash = 0;
  for (var i = 0; i < str.length; i++) {
    hash = ((hash << 5) - hash) + str.charCodeAt(i);
    hash |= 0;
  }
  return hash;
}


// =================================================================
// スプレッドシートのカスタムメニュー
// =================================================================
/**
 * スプレッドシートを開いた時に「アプリ管理」メニューを追加します。
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('アプリ管理')
      .addItem('管理パネルを開く', 'showAdminSidebar')
      .addSeparator()
      .addItem('設定シートを追加', 'createConfigSheet')
      .addToUi();
}

/**
 * 管理用サイドバーを表示します。
 */
function showAdminSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('SheetSelector')
      .setTitle('管理パネル');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * 管理パネルの初期化に必要なデータを取得します。
 * @returns {object} - 現在の公開状態とシートのリスト
 */
function getAdminSettings() {
  const props = PropertiesService.getScriptProperties();
  const userProps = PropertiesService.getUserProperties();
  const userId = userProps.getProperty('CURRENT_USER_ID');
  
  let adminEmails = [];
  let appSettings = {};
  
  if (userId) {
    // マルチテナントモード
    const userInfo = getUserInfo(userId);
    if (userInfo) {
      adminEmails = [userInfo.adminEmail];
      appSettings = getAppSettingsForUser();
    }
  } else {
    // 従来のモード
    adminEmails = getAdminEmails();
    appSettings = getAppSettingsForUser();
  }
  
  const allSheets = getSheets();
  const currentUserEmail = safeGetUserEmail();
  
  return {
    allSheets: allSheets,
    currentUserEmail: currentUserEmail,
    deployId: props.getProperty('DEPLOY_ID'),
    webAppUrl: getWebAppUrl(),
    adminEmails: adminEmails,
    isUserAdmin: adminEmails.includes(currentUserEmail),
    isPublished: appSettings.isPublished,
    activeSheetName: appSettings.activeSheetName,
    showDetails: appSettings.showDetails
  };
}

/**
 * 指定されたシートでアプリを公開します。
 * @param {string} sheetName - 公開するシート名。
 */
function publishApp(sheetName) {
  const props = PropertiesService.getUserProperties();
  const userId = props.getProperty('CURRENT_USER_ID');
  
  if (!userId) {
    throw new Error('ユーザー情報が見つかりません。');
  }
  
  const userInfo = getUserInfo(userId);
  if (userInfo.adminEmail !== Session.getActiveUser().getEmail()) {
    throw new Error('権限がありません。');
  }
  
  if (!sheetName) {
    throw new Error('シート名が指定されていません。');
  }
  
  // スプレッドシートの準備
  const spreadsheet = SpreadsheetApp.openById(userInfo.spreadsheetId);
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`シート「${sheetName}」が見つかりません。`);
  }
  
  prepareSheetForBoard(sheetName);
  
  // ユーザー設定を更新
  updateUserConfig(userId, {
    isPublished: true,
    sheetName: sheetName
  });
  
  return `「${sheetName}」を公開しました。`;
}

/**
 * アプリの公開を終了します。
 */
function unpublishApp() {
  if (!checkAdmin()) {
    throw new Error('権限がありません。');
  }
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty(APP_PROPERTIES.IS_PUBLISHED, 'false');
  return 'アプリを非公開にしました。';
}

function setShowDetails(flag) {
  if (!checkAdmin()) {
    throw new Error('権限がありません。');
  }
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty(APP_PROPERTIES.SHOW_DETAILS, String(flag));
  return `詳細表示を${flag ? '有効' : '無効'}にしました。`;
}



// =================================================================
// GAS Webアプリケーションのエントリーポイント
// =================================================================
function doGet(e) {
  e = e || {};
  const params = e.parameter || {};
  const userId = params.userId;
  const mode = params.mode;
  
  // ユーザーIDがない場合は登録画面を表示
  if (!userId) {
    return HtmlService.createTemplateFromFile('Registration')
      .evaluate()
      .setTitle('StudyQuest - 新規登録')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  
  // ユーザー情報を取得
  const userInfo = getUserInfo(userId);
  if (!userInfo) {
    return HtmlService.createHtmlOutput('無効なユーザーIDです。')
      .setTitle('エラー');
  }
  
  // 現在のコンテキストを設定
  PropertiesService.getUserProperties().setProperties({
    CURRENT_USER_ID: userId,
    CURRENT_SPREADSHEET_ID: userInfo.spreadsheetId
  });
  
  // 管理モードの場合
  if (mode === 'admin') {
    const template = HtmlService.createTemplateFromFile('SheetSelector');
    template.userId = userId;
    template.userInfo = userInfo;
    return template.evaluate()
      .setTitle('StudyQuest - 管理パネル')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  }
  
  // 通常表示モード
  const config = userInfo.configJson || {};
  
  if (!config.isPublished) {
    const template = HtmlService.createTemplateFromFile('Unpublished');
    template.userEmail = userInfo.adminEmail;
    template.userId = userId;
    return template.evaluate().setTitle('公開終了');
  }
  
  if (!config.sheetName) {
    return HtmlService.createHtmlOutput('エラー: 表示するシートが設定されていません。')
      .setTitle('エラー');
  }
  
  // 既存のPage.htmlを使用
  const template = HtmlService.createTemplateFromFile('Page');
  const configFn = (typeof global !== 'undefined' && global.getConfig)
      ? global.getConfig
      : (typeof getConfig === 'function' ? getConfig : null);
  const mapping = configFn ? configFn(config.sheetName) : {};
  
  Object.assign(template, {
    showAdminFeatures: false,
    showHighlightToggle: false,
    isAdminUser: userInfo.adminEmail === Session.getActiveUser().getEmail(),
    showCounts: config.showDetails || false,
    displayMode: config.showDetails ? 'named' : 'anonymous',
    sheetName: config.sheetName,
    mapping: mapping,
    userId: userId
  });
  
  template.userEmail = Session.getActiveUser().getEmail();
  
  return template.evaluate()
      .setTitle('StudyQuest - みんなのかいとうボード')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


// =================================================================
// クライアントサイドから呼び出される関数
// =================================================================

/**
 * サーバー側で設定されたシートのデータを取得します。
 */

function getPublishedSheetData(classFilter, sortBy) {
  const props = PropertiesService.getUserProperties();
  const spreadsheetId = props.getProperty('CURRENT_SPREADSHEET_ID');
  const userId = props.getProperty('CURRENT_USER_ID');
  
  if (!spreadsheetId || !userId) {
    throw new Error('ユーザー情報が見つかりません。');
  }
  
  const userInfo = getUserInfo(userId);
  if (!userInfo) {
    throw new Error('無効なユーザーです。');
  }
  
  const config = userInfo.configJson || {};
  const sheetName = config.sheetName;
  
  if (!sheetName) {
    throw new Error('表示するシートが設定されていません。');
  }
  
  // ユーザーのスプレッドシートを開く
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    throw new Error(`シート「${sheetName}」が見つかりません。`);
  }
  
  // 既存のgetSheetData関数のロジックを使用（スプレッドシートを指定）
  const order = sortBy || 'newest';
  const data = getSheetDataForSpreadsheet(spreadsheet, sheetName, classFilter, order);
  
  return {
    sheetName: sheetName,
    header: data.header,
    rows: data.rows,
    showDetails: config.showDetails || false
  };
}


// =================================================================
// 内部処理関数
// =================================================================
function getSheets() {
  try {
    const allSheets = getCurrentSpreadsheet().getSheets();
    const visibleSheets = allSheets.filter(sheet => !sheet.isSheetHidden());
    const filtered = visibleSheets.filter(sheet => {
      const name = sheet.getName();
      return name !== 'Config' && name !== ROSTER_CONFIG.SHEET_NAME;
    });
    return filtered.map(sheet => sheet.getName());
  } catch (error) {
    handleError('getSheets', error);
  }
}

function getSheetHeaders(sheetName) {
  const sheet = getCurrentSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`シート '${sheetName}' が見つかりません。`);
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headerRow.map(v => (v !== null && v !== undefined) ? String(v) : '');
}

function getSheetData(sheetName, classFilter, sortBy) {
  try {
    const sheet = getCurrentSpreadsheet().getSheetByName(sheetName);
    if (!sheet) throw new Error(`指定されたシート「${sheetName}」が見つかりません。`);

    const allValues = sheet.getDataRange().getValues();
    if (allValues.length < 1) return { header: "シートにデータがありません", rows: [] };

    const cfgFunc = (typeof global !== 'undefined' && global.getConfig)
      ? global.getConfig
      : (typeof getConfig === 'function' ? getConfig : null);
    const cfg = cfgFunc ? cfgFunc(sheetName) : {};
    const answerHeader = cfg.answerHeader || COLUMN_HEADERS.OPINION;
    const reasonHeader = cfg.reasonHeader || COLUMN_HEADERS.REASON;
    const classHeader = cfg.classHeader || COLUMN_HEADERS.CLASS;
    const nameHeader = cfg.nameHeader || '';

    const required = [
      COLUMN_HEADERS.EMAIL,
      classHeader,
      answerHeader,
      reasonHeader,
      COLUMN_HEADERS.TIMESTAMP,
      COLUMN_HEADERS.UNDERSTAND,
      COLUMN_HEADERS.LIKE,
      COLUMN_HEADERS.CURIOUS,
      COLUMN_HEADERS.HIGHLIGHT
    ];
    if (nameHeader) required.push(nameHeader);

    const headerIndices = findHeaderIndices(allValues[0], required);
    const dataRows = allValues.slice(1);
    const userEmail = safeGetUserEmail();
    const isAdmin = isUserAdmin(userEmail);
    const emailToNameMap = {};

    const rows = dataRows.map((row, index) => {
      if (classFilter && classFilter !== 'すべて') {
        const className = row[headerIndices[classHeader]];
        if (className !== classFilter) return null;
      }
      const email = row[headerIndices[COLUMN_HEADERS.EMAIL]];
      const opinion = row[headerIndices[answerHeader]];

      if (email && opinion) {
        const understandArr = parseReactionString(row[headerIndices[COLUMN_HEADERS.UNDERSTAND]]);
        const likeArr = parseReactionString(row[headerIndices[COLUMN_HEADERS.LIKE]]);
        const curiousArr = parseReactionString(row[headerIndices[COLUMN_HEADERS.CURIOUS]]);
        const reason = row[headerIndices[reasonHeader]] || '';
        const highlightVal = row[headerIndices[COLUMN_HEADERS.HIGHLIGHT]];
        const timestampRaw = row[headerIndices[COLUMN_HEADERS.TIMESTAMP]];
        const timestamp = timestampRaw ? new Date(timestampRaw).toISOString() : '';
        const likes = likeArr.length;
        const baseScore = reason.length;
        const likeMultiplier = 1 + (likes * SCORING_CONFIG.LIKE_MULTIPLIER_FACTOR);
        const totalScore = baseScore * likeMultiplier;
        let name = emailToNameMap[email] || email.split('@')[0];
        if (nameHeader && row[headerIndices[nameHeader]]) {
          name = row[headerIndices[nameHeader]];
        }
        return {
          rowIndex: index + 2,
          timestamp: timestamp,
          name: isAdmin ? name : '',
          class: row[headerIndices[classHeader]] || '未分類',
          opinion: opinion,
          reason: reason,
          reactions: {
            UNDERSTAND: { count: understandArr.length, reacted: userEmail ? understandArr.includes(userEmail) : false },
            LIKE: { count: likeArr.length, reacted: userEmail ? likeArr.includes(userEmail) : false },
            CURIOUS: { count: curiousArr.length, reacted: userEmail ? curiousArr.includes(userEmail) : false }
          },
          highlight: highlightVal === true || String(highlightVal).toLowerCase() === 'true',
          score: totalScore
        };
      }
      return null;
    }).filter(Boolean);

    if (sortBy === 'newest') {
      rows.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    } else if (sortBy === 'random') {
      rows.sort((a, b) => hashTimestamp(a.timestamp) - hashTimestamp(b.timestamp));
    } else {
      rows.sort((a, b) => b.score - a.score);
    }
    const header = cfg.questionHeader || answerHeader;
    return { header: header, rows: rows };
  } catch(e) {
    handleError(`getSheetData for ${sheetName}`, e);
  }
}

function getSheetDataForSpreadsheet(spreadsheet, sheetName, classFilter, sortBy) {
  try {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) throw new Error(`指定されたシート「${sheetName}」が見つかりません。`);

    const allValues = sheet.getDataRange().getValues();
    if (allValues.length < 1) return { header: "シートにデータがありません", rows: [] };

    const cfgFunc = (typeof global !== 'undefined' && global.getConfig)
      ? global.getConfig
      : (typeof getConfig === 'function' ? getConfig : null);
    const cfg = cfgFunc ? cfgFunc(sheetName) : {};
    const answerHeader = cfg.answerHeader || COLUMN_HEADERS.OPINION;
    const reasonHeader = cfg.reasonHeader || COLUMN_HEADERS.REASON;
    const classHeader = cfg.classHeader || COLUMN_HEADERS.CLASS;
    const nameHeader = cfg.nameHeader || '';

    const required = [
      COLUMN_HEADERS.EMAIL,
      classHeader,
      answerHeader,
      reasonHeader,
      COLUMN_HEADERS.TIMESTAMP,
      COLUMN_HEADERS.UNDERSTAND,
      COLUMN_HEADERS.LIKE,
      COLUMN_HEADERS.CURIOUS,
      COLUMN_HEADERS.HIGHLIGHT
    ];
    if (nameHeader) required.push(nameHeader);

    const headerIndices = findHeaderIndices(allValues[0], required);
    const dataRows = allValues.slice(1);
    const userEmail = safeGetUserEmail();
    const isAdmin = isUserAdmin(userEmail);
    const emailToNameMap = {};

    const rows = dataRows.map((row, index) => {
      if (classFilter && classFilter !== 'すべて') {
        const className = row[headerIndices[classHeader]];
        if (className !== classFilter) return null;
      }
      const email = row[headerIndices[COLUMN_HEADERS.EMAIL]];
      const opinion = row[headerIndices[answerHeader]];

      if (email && opinion) {
        const understandArr = parseReactionString(row[headerIndices[COLUMN_HEADERS.UNDERSTAND]]);
        const likeArr = parseReactionString(row[headerIndices[COLUMN_HEADERS.LIKE]]);
        const curiousArr = parseReactionString(row[headerIndices[COLUMN_HEADERS.CURIOUS]]);
        const reason = row[headerIndices[reasonHeader]] || '';
        const highlightVal = row[headerIndices[COLUMN_HEADERS.HIGHLIGHT]];
        const timestampRaw = row[headerIndices[COLUMN_HEADERS.TIMESTAMP]];
        const timestamp = timestampRaw ? new Date(timestampRaw).toISOString() : '';
        const likes = likeArr.length;
        const baseScore = reason.length;
        const likeMultiplier = 1 + (likes * SCORING_CONFIG.LIKE_MULTIPLIER_FACTOR);
        const totalScore = baseScore * likeMultiplier;
        let name = emailToNameMap[email] || email.split('@')[0];
        if (nameHeader && row[headerIndices[nameHeader]]) {
          name = row[headerIndices[nameHeader]];
        }
        return {
          rowIndex: index + 2,
          timestamp: timestamp,
          name: isAdmin ? name : '',
          class: row[headerIndices[classHeader]] || '未分類',
          opinion: opinion,
          reason: reason,
          reactions: {
            UNDERSTAND: { count: understandArr.length, reacted: userEmail ? understandArr.includes(userEmail) : false },
            LIKE: { count: likeArr.length, reacted: userEmail ? likeArr.includes(userEmail) : false },
            CURIOUS: { count: curiousArr.length, reacted: userEmail ? curiousArr.includes(userEmail) : false }
          },
          highlight: highlightVal === true || String(highlightVal).toLowerCase() === 'true',
          score: totalScore
        };
      }
      return null;
    }).filter(Boolean);

    if (sortBy === 'newest') {
      rows.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    } else if (sortBy === 'random') {
      rows.sort((a, b) => hashTimestamp(a.timestamp) - hashTimestamp(b.timestamp));
    } else {
      rows.sort((a, b) => b.score - a.score);
    }
    const header = cfg.questionHeader || answerHeader;
    return { header: header, rows: rows };
  } catch(e) {
    handleError(`getSheetDataForSpreadsheet for ${sheetName}`, e);
  }
}

function buildBoardData(sheetName) {
  const cfgFunc = (typeof global !== 'undefined' && global.getConfig) ? global.getConfig : getConfig;
  const cfg = cfgFunc ? cfgFunc(sheetName) : {};
  const sheet = getCurrentSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`シート '${sheetName}' が見つかりません。`);
  const values = sheet.getDataRange().getValues();
  const headers = values.shift();
  const index = {};
  headers.forEach((h,i)=>{ if(h) index[h]=i; });
  const entries = values.map(row => {
    const email = index[COLUMN_HEADERS.EMAIL] !== undefined ? row[index[COLUMN_HEADERS.EMAIL]] : '';
    return {
      answer: row[index[cfg.answerHeader]],
      reason: cfg.reasonHeader ? row[index[cfg.reasonHeader]] : null,
      name: cfg.nameHeader ? row[index[cfg.nameHeader]] : (email ? email.split('@')[0] : ''),
      class: cfg.classHeader ? row[index[cfg.classHeader]] : undefined
    };
  });
  return { header: cfg.questionHeader, entries };
}

function addReaction(rowIndex, reactionKey, sheetName) {
  if (!rowIndex || !reactionKey || !COLUMN_HEADERS[reactionKey]) {
    return { status: 'error', message: '無効なパラメータです。' };
  }
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(TIME_CONSTANTS.LOCK_WAIT_MS)) {
      return { status: 'error', message: '他のユーザーが操作中です。しばらく待ってから再試行してください。' };
    }
    const userEmail = safeGetUserEmail();
    if (!userEmail) {
      return { status: 'error', message: 'ログインしていないため、操作できません。' };
    }
    
    // マルチテナント対応の修正
    const spreadsheet = getCurrentSpreadsheet();
    const settings = getAppSettingsForUser();
    const targetSheet = sheetName || settings.activeSheetName;
    const sheet = spreadsheet.getSheetByName(targetSheet);
    
    if (!sheet) throw new Error(`シート '${targetSheet}' が見つかりません。`);

  const reactionHeaders = REACTION_KEYS.map(k => COLUMN_HEADERS[k]);
  const headerIndices = getHeaderIndices(targetSheet);
  const startCol = headerIndices[reactionHeaders[0]] + 1;
  const reactionRange = sheet.getRange(rowIndex, startCol, 1, REACTION_KEYS.length);
  const values = reactionRange.getValues()[0];
  const lists = {};
  REACTION_KEYS.forEach((k, idx) => {
    lists[k] = { arr: parseReactionString(values[idx]) };
  });

    const wasReacted = lists[reactionKey].arr.includes(userEmail);
    Object.keys(lists).forEach(key => {
      const idx = lists[key].arr.indexOf(userEmail);
      if (idx > -1) lists[key].arr.splice(idx, 1);
    });
    if (!wasReacted) {
      lists[reactionKey].arr.push(userEmail);
    }
    reactionRange.setValues([
      REACTION_KEYS.map(k => lists[k].arr.join(','))
    ]);

    const reactions = REACTION_KEYS.reduce((obj, k) => {
      obj[k] = {
        count: lists[k].arr.length,
        reacted: lists[k].arr.includes(userEmail)
      };
      return obj;
    }, {});
    return { status: 'ok', reactions: reactions };
  } catch (error) {
    return handleError('addReaction', error, true);
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function toggleHighlight(rowIndex, sheetName) {
  if (!checkAdmin()) {
    return { status: 'error', message: '権限がありません。' };
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(TIME_CONSTANTS.LOCK_WAIT_MS);
  try {
    const settings = getAppSettingsForUser();
    const targetSheet = sheetName || settings.activeSheetName;
    const sheet = getCurrentSpreadsheet().getSheetByName(targetSheet);
    if (!sheet) throw new Error(`シート '${targetSheet}' が見つかりません。`);

    const headerIndices = getHeaderIndices(targetSheet);
    const colIndex = headerIndices[COLUMN_HEADERS.HIGHLIGHT] + 1;

    const cell = sheet.getRange(rowIndex, colIndex);
    const current = !!cell.getValue();
    const newValue = !current;
    cell.setValue(newValue);
    return { status: 'ok', highlight: newValue };
  } catch (error) {
    return handleError('toggleHighlight', error, true);
  } finally {
    lock.releaseLock();
  }
}


function getAppSettings() {
  const properties = PropertiesService.getScriptProperties() || {};
  const getProp = typeof properties.getProperty === 'function' ? (k) => properties.getProperty(k) : () => null;
  const published = getProp(APP_PROPERTIES.IS_PUBLISHED);
  const sheet = getProp(APP_PROPERTIES.ACTIVE_SHEET);
  const showDetailsProp = getProp(APP_PROPERTIES.SHOW_DETAILS);
  let activeName = sheet;
  if (!activeName) {
    try {
      const sheets = getSheets();
      activeName = sheets[0] || '';
    } catch (error) {
      console.error('getAppSettings Error:', error);
      activeName = '';
    }
  }
  return {
    isPublished: published === null ? true : published === 'true',
    activeSheetName: activeName,
    showDetails: showDetailsProp === 'true'
  };
}

function getAppSettingsForUser() {
  const props = PropertiesService.getUserProperties();
  const userId = props.getProperty('CURRENT_USER_ID');
  
  if (!userId) {
    // 従来の動作
    return getAppSettings();
  }
  
  // マルチテナント時
  const userInfo = getUserInfo(userId);
  if (!userInfo || !userInfo.configJson) {
    return {
      isPublished: false,
      activeSheetName: '',
      showDetails: false
    };
  }
  
  return {
    isPublished: userInfo.configJson.isPublished || false,
    activeSheetName: userInfo.configJson.sheetName || '',
    showDetails: userInfo.configJson.showDetails || false
  };
}

function getHeaderIndices(sheetName) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `headers_${sheetName}`;
  const cached = cache.get(cacheKey);
  if (cached) { return JSON.parse(cached); }
  const sheet = getCurrentSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`シート '${sheetName}' が見つかりません。`);
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const indices = findHeaderIndices(headerRow, [
    COLUMN_HEADERS.TIMESTAMP,
    COLUMN_HEADERS.UNDERSTAND,
    COLUMN_HEADERS.LIKE,
    COLUMN_HEADERS.CURIOUS,
    COLUMN_HEADERS.HIGHLIGHT
  ]);
  cache.put(cacheKey, JSON.stringify(indices), 21600);
  return indices;
}


function prepareSheetForBoard(sheetName) {
  const sheet = getCurrentSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`シート '${sheetName}' が見つかりません。`);
  const headers = getSheetHeaders(sheetName);
  let lastCol = headers.length;
  const required = [
    COLUMN_HEADERS.TIMESTAMP,
    COLUMN_HEADERS.UNDERSTAND,
    COLUMN_HEADERS.LIKE,
    COLUMN_HEADERS.CURIOUS,
    COLUMN_HEADERS.HIGHLIGHT
  ];
  required.forEach(h => {
    if (headers.indexOf(h) === -1) {
      sheet.insertColumnAfter(lastCol);
      lastCol++;
      sheet.getRange(1, lastCol).setValue(h);
      headers.push(h);
    }
  });
}

function findHeaderIndices(sheetHeaders, requiredHeaders) {
  const indices = {};
  const normalized = sheetHeaders.map(h => (typeof h === 'string' ? h.replace(/\s+/g, '') : h));
  const missingHeaders = [];
  requiredHeaders.forEach(name => {
    const idx = normalized.indexOf(name.replace(/\s+/g, ''));
    if (idx !== -1) { indices[name] = idx; } else { missingHeaders.push(name); }
  });
  if (missingHeaders.length > 0) {
    throw new Error(`必須ヘッダーが見つかりません: [${missingHeaders.join(', ')}]`);
  }
  return indices;
}

function saveWebAppUrl(url) {
  const props = PropertiesService.getScriptProperties();
  props.setProperties({ WEB_APP_URL: (url || '').trim() });
}

function getWebAppUrl() {
  const props = PropertiesService.getScriptProperties();
  return (props.getProperty('WEB_APP_URL') || '').trim();
}

function saveDeployId(id) {
  const props = PropertiesService.getScriptProperties();
  props.setProperties({ DEPLOY_ID: (id || '').trim() });
}

function createTemplateSheet(name) {
  if (!checkAdmin()) {
    throw new Error('権限がありません。');
  }
  const ss = getCurrentSpreadsheet();
  const sheetName = name || 'New Q&A';
  if (ss.getSheetByName(sheetName)) {
    throw new Error(`シート '${sheetName}' は既に存在します。`);
  }
  const sheet = ss.insertSheet(sheetName);
  const headers = [
    COLUMN_HEADERS.EMAIL,
    COLUMN_HEADERS.CLASS,
    '質問',
    '回答',
    '理由',
    COLUMN_HEADERS.NAME,
    'クラス'
  ];
  const req = [
    COLUMN_HEADERS.UNDERSTAND,
    COLUMN_HEADERS.LIKE,
    COLUMN_HEADERS.CURIOUS,
    COLUMN_HEADERS.HIGHLIGHT
  ];
  const all = headers.concat(req);
  sheet.getRange(1, 1, 1, all.length).setValues([all]);
  return sheet.getName();
}

// =================================================================
// マルチテナント管理機能
// =================================================================

/**
 * 新規ユーザーを登録
 */
function registerNewUser(spreadsheetUrl) {
  const userEmail = Session.getActiveUser().getEmail();
  if (!userEmail) {
    throw new Error('ユーザー認証に失敗しました。');
  }
  
  // URLからスプレッドシートIDを抽出
  const spreadsheetId = extractSpreadsheetId(spreadsheetUrl);
  
  // アクセス権限を確認
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    // StudyQuest用の初期設定を行う
    prepareSpreadsheetForStudyQuest(ss);
  } catch (e) {
    throw new Error('スプレッドシートにアクセスできません。編集権限があることを確認してください。');
  }
  
  // 既存ユーザーかチェック
  const existingUser = findUserByEmailAndSpreadsheet(userEmail, spreadsheetId);
  if (existingUser) {
    return {
      adminUrl: `${getWebAppUrl()}?userId=${existingUser.userId}&mode=admin`,
      viewUrl: `${getWebAppUrl()}?userId=${existingUser.userId}`,
      userId: existingUser.userId,
      message: '既存の登録が見つかりました。'
    };
  }
  
  // 新規ユーザーID生成
  const userId = Utilities.getUuid();
  const accessToken = generateAccessToken();
  
  // ユーザー情報を保存
  const userDb = getUserDatabase();
  userDb.appendRow([
    userId,
    userEmail,
    spreadsheetId,
    spreadsheetUrl,
    new Date(),
    accessToken,
    JSON.stringify({
      sheetName: '',
      showDetails: false,
      isPublished: false
    }),
    new Date(),
    true
  ]);
  
  // 管理者として登録
  updateAdminEmails(spreadsheetId, userEmail);
  
  return {
    adminUrl: `${getWebAppUrl()}?userId=${userId}&mode=admin`,
    viewUrl: `${getWebAppUrl()}?userId=${userId}`,
    userId: userId,
    message: '新規登録が完了しました。'
  };
}

/**
 * スプレッドシートをStudyQuest用に初期設定
 */
function prepareSpreadsheetForStudyQuest(spreadsheet) {
  // Configシートがなければ作成
  let configSheet = spreadsheet.getSheetByName('Config');
  if (!configSheet) {
    configSheet = spreadsheet.insertSheet('Config');
    const headers = ['表示シート名','問題文ヘッダー','回答ヘッダー','理由ヘッダー','名前列ヘッダー','クラス列ヘッダー'];
    configSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  
  // rosterシートがなければ作成
  let rosterSheet = spreadsheet.getSheetByName('roster');
  if (!rosterSheet) {
    rosterSheet = spreadsheet.insertSheet('roster');
    const headers = ['学年','組','番号','姓','名','Googleアカウント','ニックネーム'];
    rosterSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

/**
 * ユーザー情報を取得
 */
function getUserInfo(userId) {
  if (!userId) return null;
  
  const userDb = getUserDatabase();
  const data = userDb.getDataRange().getValues();
  const headers = data[0];
  const userIdIndex = headers.indexOf('userId');
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][userIdIndex] === userId) {
      const userInfo = {};
      headers.forEach((header, index) => {
        if (header === 'configJson') {
          try {
            userInfo[header] = JSON.parse(data[i][index]);
          } catch (e) {
            userInfo[header] = {};
          }
        } else {
          userInfo[header] = data[i][index];
        }
      });
      
      // 最終アクセス日時を更新
      userDb.getRange(i + 1, headers.indexOf('lastAccessedAt') + 1).setValue(new Date());
      
      return userInfo;
    }
  }
  
  return null;
}

/**
 * ユーザー設定を更新
 */
function updateUserConfig(userId, config) {
  const userDb = getUserDatabase();
  const data = userDb.getDataRange().getValues();
  const headers = data[0];
  const userIdIndex = headers.indexOf('userId');
  const configIndex = headers.indexOf('configJson');
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][userIdIndex] === userId) {
      const currentConfig = JSON.parse(data[i][configIndex] || '{}');
      const newConfig = Object.assign({}, currentConfig, config);
      userDb.getRange(i + 1, configIndex + 1).setValue(JSON.stringify(newConfig));
      return newConfig;
    }
  }
  
  throw new Error('ユーザーが見つかりません。');
}

/**
 * ユーザーデータベースを取得（なければ作成）
 */
function getUserDatabase() {
  const dbId = PropertiesService.getScriptProperties().getProperty('USER_DB_ID');
  
  if (dbId) {
    try {
      return SpreadsheetApp.openById(dbId).getSheetByName(USER_DB_CONFIG.SHEET_NAME);
    } catch (e) {
      // データベースが削除されている場合は再作成
    }
  }
  
  // 新規作成
  const newDb = SpreadsheetApp.create('StudyQuest_UserDatabase');
  const sheet = newDb.getActiveSheet();
  sheet.setName(USER_DB_CONFIG.SHEET_NAME);
  sheet.getRange(1, 1, 1, USER_DB_CONFIG.HEADERS.length)
    .setValues([USER_DB_CONFIG.HEADERS]);

  // Share database so users in the same domain can read it
  try {
    DriveApp.getFileById(newDb.getId())
      .setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {}
  
  PropertiesService.getScriptProperties().setProperty('USER_DB_ID', newDb.getId());
  
  return sheet;
}

/**
 * スプレッドシートIDを抽出
 */
function extractSpreadsheetId(url) {
  const patterns = [
    /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/,
    /[?&]id=([a-zA-Z0-9-_]+)/,
    /^([a-zA-Z0-9-_]+)$/
  ];
  
  for (const pattern of patterns) {
    const match = url.match(pattern);
    if (match) return match[1];
  }
  
  throw new Error('有効なGoogleスプレッドシートのURLまたはIDを入力してください。');
}

/**
 * アクセストークン生成
 */
function generateAccessToken() {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    Utilities.getUuid() + new Date().getTime()
  );
  return Utilities.base64Encode(bytes);
}

/**
 * 既存ユーザーを検索
 */
function findUserByEmailAndSpreadsheet(email, spreadsheetId) {
  const userDb = getUserDatabase();
  const data = userDb.getDataRange().getValues();
  const headers = data[0];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][headers.indexOf('adminEmail')] === email &&
        data[i][headers.indexOf('spreadsheetId')] === spreadsheetId) {
      const userInfo = {};
      headers.forEach((header, index) => {
        userInfo[header] = data[i][index];
      });
      return userInfo;
    }
  }
  
  return null;
}

/**
 * 管理者メールアドレスを更新
 */
function updateAdminEmails(spreadsheetId, email) {
  // スプレッドシートのプロパティに管理者メールを保存
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const docProps = PropertiesService.getDocumentProperties();
    const currentAdmins = docProps.getProperty('ADMIN_EMAILS') || '';
    const adminList = currentAdmins.split(',').filter(Boolean);
    
    if (!adminList.includes(email)) {
      adminList.push(email);
      docProps.setProperty('ADMIN_EMAILS', adminList.join(','));
    }
  } catch (e) {
    console.error('管理者メール更新エラー:', e);
  }
}

/**
 * 現在のユーザーのメールアドレスを取得
 */
function getActiveUserEmail() {
  return Session.getActiveUser().getEmail();
}

if (typeof module !== 'undefined') {
  handleError = require('./ErrorHandling.gs').handleError;
  const { getConfig, saveSheetConfig, createConfigSheet } = require('./config.gs');
  module.exports = {
    COLUMN_HEADERS,
    getAdminSettings,
    publishApp,
    unpublishApp,
    doGet,
    getConfig,
    buildBoardData,
    getPublishedSheetData,
    getSheets,
    getSheetHeaders,
    getSheetData,
    getSheetDataForSpreadsheet,
    getHeaderIndices,
    addReaction,
    toggleHighlight,
    setShowDetails,
    saveWebAppUrl,
    getWebAppUrl,
    saveDeployId,
    findHeaderIndices,
    parseReactionString,
    checkAdmin,
    handleError,
    saveSheetConfig,
    createConfigSheet,
    createTemplateSheet,
    prepareSheetForBoard,
    registerNewUser,
    getUserInfo,
    updateUserConfig,
    getUserDatabase,
    extractSpreadsheetId,
    generateAccessToken,
    findUserByEmailAndSpreadsheet,
    updateAdminEmails,
    getActiveUserEmail,
    prepareSpreadsheetForStudyQuest
  };
}
