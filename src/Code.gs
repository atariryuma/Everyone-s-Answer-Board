/**
 * @fileoverview StudyQuest - みんなの回答ボード
 * ★【管理者メニュー実装版】スプレッドシートのカスタムメニューからアプリを管理する機能を実装。
 */

// =================================================================
// 定数定義
// =================================================================
const COLUMN_HEADERS = {
  EMAIL: 'メールアドレス',
  CLASS: 'クラスを選択してください。',
  OPINION: 'これまでの学んだことや、経験したことから、根からとり入れた水は、植物のからだのどこを通るのか予想しましょう。',
  REASON: '予想したわけを書きましょう。',
  UNDERSTAND: 'なるほど！',
  LIKE: 'いいね！',
  CURIOUS: 'もっと知りたい！',
  HIGHLIGHT: 'ハイライト'
};
const ROSTER_CONFIG = {
  SHEET_NAME: 'sheet 1',
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
if (typeof global !== 'undefined' && global.getConfig) {
  getConfig = global.getConfig;
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
      .addItem('名簿キャッシュをリセット', 'clearRosterCache')
      .addItem('Configシートを作成', 'createConfigSheet')
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
  const allSheets = getSheets();
  const currentUserEmail = safeGetUserEmail();
  const adminEmails = getAdminEmails();
  const appSettings = getAppSettings();
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
  if (!checkAdmin()) {
    throw new Error('権限がありません。');
  }
  if (!sheetName) {
    throw new Error('シート名が指定されていません。');
  }
  prepareSheetForBoard(sheetName);
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty(APP_PROPERTIES.IS_PUBLISHED, 'true');
  properties.setProperty(APP_PROPERTIES.ACTIVE_SHEET, sheetName);
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
  const settings = getAppSettings();
  
  if (!settings.isPublished) {
    const userEmail = safeGetUserEmail();
    const template = HtmlService.createTemplateFromFile('Unpublished');
    template.userEmail = userEmail;
    return template.evaluate().setTitle('公開終了');
  }
  
  if (!settings.activeSheetName) {
    return HtmlService.createHtmlOutput('エラー: 表示するシートが設定されていません。スプレッドシートの「アプリ管理」メニューから設定してください。').setTitle('エラー');
  }

  const sheetName = (e && e.parameter && e.parameter.sheetName) ? e.parameter.sheetName : settings.activeSheetName;
  const mapping = (typeof getConfig === 'function') ? getConfig(sheetName) : {};
  const userEmail = safeGetUserEmail();

  const template = HtmlService.createTemplateFromFile('Page');
  const admin = isUserAdmin(userEmail);
  Object.assign(template, {
    showAdminFeatures: false,
    showHighlightToggle: false,
    isAdminUser: admin,
    showCounts: settings.showDetails,
    displayMode: settings.showDetails ? 'named' : 'anonymous',
    sheetName: sheetName,
    mapping: mapping
  });
  template.userEmail = userEmail;
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
  const settings = getAppSettings();
  const sheetName = settings.activeSheetName;

  if (!sheetName) {
    throw new Error('表示するシートが設定されていません。');
  }

  const order = sortBy || 'newest';
  const data = getSheetData(sheetName, classFilter, order);

  // ★改善: フロントエンドでシート名を表示できるよう、レスポンスに含める
  return {
    sheetName: sheetName,
    header: data.header,
    rows: data.rows,
    showDetails: settings.showDetails
  };
}


// =================================================================
// 内部処理関数
// =================================================================
function getSheets() {
  try {
    const allSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
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
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`シート '${sheetName}' が見つかりません。`);
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headerRow.map(v => (v !== null && v !== undefined) ? String(v) : '');
}

function getSheetData(sheetName, classFilter, sortBy) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) throw new Error(`指定されたシート「${sheetName}」が見つかりません。`);

    const allValues = sheet.getDataRange().getValues();
    if (allValues.length < 1) return { header: "シートにデータがありません", rows: [] };

    const cfgFunc = (typeof getConfig === 'function') ? getConfig : null;
    const cfg = cfgFunc ? cfgFunc(sheetName) : {};
    const answerHeader = cfg.answerHeader || COLUMN_HEADERS.OPINION;
    const reasonHeader = cfg.reasonHeader || COLUMN_HEADERS.REASON;
    const classHeader = cfg.classHeader || COLUMN_HEADERS.CLASS;
    const nameHeader = (cfg.nameMode === '同一シート') ? (cfg.nameHeader || '') : '';

    const required = [
      COLUMN_HEADERS.EMAIL,
      classHeader,
      answerHeader,
      reasonHeader,
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
    const emailToNameMap = getRosterMap();

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
      rows.sort((a, b) => b.rowIndex - a.rowIndex);
    } else if (sortBy === 'random') {
      rows.sort(() => Math.random() - 0.5);
    } else {
      rows.sort((a, b) => b.score - a.score);
    }
    const header = cfg.questionHeader || answerHeader;
    return { header: header, rows: rows };
  } catch(e) {
    handleError(`getSheetData for ${sheetName}`, e);
  }
}

function buildBoardData(sheetName) {
  const cfgFunc = (typeof global !== 'undefined' && global.getConfig) ? global.getConfig : getConfig;
  const cfg = cfgFunc ? cfgFunc(sheetName) : {};
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`シート '${sheetName}' が見つかりません。`);
  const values = sheet.getDataRange().getValues();
  const headers = values.shift();
  const index = {};
  headers.forEach((h,i)=>{ if(h) index[h]=i; });
  const entries = values.map(row => {
    const email = index[COLUMN_HEADERS.EMAIL] !== undefined ? row[index[COLUMN_HEADERS.EMAIL]] : '';
    const nameFromRoster = cfg.nameMode === '別シート' ? (getRosterMap()[email] || undefined) : undefined;
    return {
      answer: row[index[cfg.answerHeader]],
      reason: cfg.reasonHeader ? row[index[cfg.reasonHeader]] : null,
      name: cfg.nameMode === '同一シート' && cfg.nameHeader ? row[index[cfg.nameHeader]] : nameFromRoster,
      class: cfg.classHeader ? row[index[cfg.classHeader]] : undefined
    };
  });
  return { header: cfg.questionHeader, entries };
}

function addReaction(rowIndex, reactionKey) {
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
    const settings = getAppSettings();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(settings.activeSheetName);
    if (!sheet) throw new Error(`シート '${settings.activeSheetName}' が見つかりません。`);

  const reactionHeaders = REACTION_KEYS.map(k => COLUMN_HEADERS[k]);
  const headerIndices = getHeaderIndices(settings.activeSheetName);
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

function toggleHighlight(rowIndex) {
  if (!checkAdmin()) {
    return { status: 'error', message: '権限がありません。' };
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(TIME_CONSTANTS.LOCK_WAIT_MS);
  try {
    const sheetName = getAppSettings().activeSheetName;
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) throw new Error(`シート '${sheetName}' が見つかりません。`);

    const headerIndices = getHeaderIndices(sheetName);
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

function getRosterMap() {
  const cache = CacheService.getScriptCache();
  const cachedMap = cache.get(ROSTER_CONFIG.CACHE_KEY);
  if (cachedMap) { return JSON.parse(cachedMap); }
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ROSTER_CONFIG.SHEET_NAME);
  if (!sheet) { console.error(`名簿シート「${ROSTER_CONFIG.SHEET_NAME}」が見つかりません。`); return {}; }
  const rosterValues = sheet.getDataRange().getValues();
  const rosterHeaders = rosterValues.shift();
  const lastNameIndex = rosterHeaders.indexOf(ROSTER_CONFIG.HEADER_LAST_NAME);
  const firstNameIndex = rosterHeaders.indexOf(ROSTER_CONFIG.HEADER_FIRST_NAME);
  const nicknameIndex = rosterHeaders.indexOf(ROSTER_CONFIG.HEADER_NICKNAME);
  const emailIndex = rosterHeaders.indexOf(ROSTER_CONFIG.HEADER_EMAIL);
  if (lastNameIndex === -1 || firstNameIndex === -1 || emailIndex === -1) { throw new Error(`名簿シート「${ROSTER_CONFIG.SHEET_NAME}」に必要な列が見つかりません。`); }
  const nameMap = {};
  rosterValues.forEach(row => {
    const email = row[emailIndex];
    const lastName = row[lastNameIndex];
    const firstName = row[firstNameIndex];
    const nickname = (nicknameIndex !== -1 && row[nicknameIndex]) ? row[nicknameIndex] : ''; 
    if (email && lastName && firstName) {
      const fullName = `${lastName} ${firstName}`;
      nameMap[email] = nickname ? `${fullName} (${nickname})` : fullName;
    }
  });
  cache.put(ROSTER_CONFIG.CACHE_KEY, JSON.stringify(nameMap), 21600);
  return nameMap;
}

function getHeaderIndices(sheetName) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `headers_${sheetName}`;
  const cached = cache.get(cacheKey);
  if (cached) { return JSON.parse(cached); }
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`シート '${sheetName}' が見つかりません。`);
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const indices = findHeaderIndices(headerRow, Object.values(COLUMN_HEADERS));
  cache.put(cacheKey, JSON.stringify(indices), 21600);
  return indices;
}

function getAndCacheHeaderIndices(sheetName, headerRow) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `headers_${sheetName}`;
  const cachedHeaders = cache.get(cacheKey);
  if (cachedHeaders) { return JSON.parse(cachedHeaders); }
  const indices = findHeaderIndices(headerRow, Object.values(COLUMN_HEADERS));
  cache.put(cacheKey, JSON.stringify(indices), 21600);
  return indices;
}

function prepareSheetForBoard(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`シート '${sheetName}' が見つかりません。`);
  const headers = getSheetHeaders(sheetName);
  let lastCol = headers.length;
  const required = [
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

function clearRosterCache() {
  const cache = CacheService.getScriptCache();
  const cacheKey = ROSTER_CONFIG.CACHE_KEY;
  cache.remove(cacheKey);
  console.log(`名簿キャッシュ（キー: ${cacheKey}）を削除しました。`);
  try {
    SpreadsheetApp.getUi().alert('名簿のキャッシュをリセットしました。');
  } catch (e) { /* no-op */ }
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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
    '名前',
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

if (typeof module !== 'undefined') {
  const { handleError } = require('./ErrorHandling.gs');
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
    getRosterMap,
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
    prepareSheetForBoard
  };
}
