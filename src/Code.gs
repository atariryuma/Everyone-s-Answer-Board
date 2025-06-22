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

if (typeof global !== 'undefined') {
  global.COLUMN_HEADERS = COLUMN_HEADERS;
  global.ROSTER_CONFIG = ROSTER_CONFIG;
  global.SCORING_CONFIG = SCORING_CONFIG;
  global.APP_PROPERTIES = APP_PROPERTIES;
  global.TIME_CONSTANTS = TIME_CONSTANTS;
}

var safeGetUserEmail, getAdminEmails, isUserAdmin, checkAdmin;
var getAdminSettings, publishApp, unpublishApp, setShowDetails;
var getSheets, getAppSettings;
var saveWebAppUrl, getWebAppUrl, saveDeployId;
var parseReactionString, addReaction, toggleHighlight;
var getConfig;
if (typeof global !== 'undefined' && global.getConfig) {
  getConfig = global.getConfig;
}



// =================================================================
// スプレッドシートのカスタムメニュー
// =================================================================
/**
 * スプレッドシートを開いた時に「アプリ管理」メニューを追加します。
 */

/**
 * 管理パネルの初期化に必要なデータを取得します。
 * @returns {object} - 現在の公開状態とシートのリスト
 */



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




function getRosterMap() {
  const cache = CacheService.getScriptCache();
  const cachedMap = cache.get(ROSTER_CONFIG.CACHE_KEY);
  if (cachedMap) { return JSON.parse(cachedMap); }
  const props = PropertiesService.getScriptProperties();
  const rosterSheetName = props && typeof props.getProperty === 'function'
    ? (props.getProperty(ROSTER_CONFIG.PROPERTY_NAME) || ROSTER_CONFIG.SHEET_NAME)
    : ROSTER_CONFIG.SHEET_NAME;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(rosterSheetName);
  if (!sheet) { console.error(`名簿シート「${rosterSheetName}」が見つかりません。`); return {}; }
  const rosterValues = sheet.getDataRange().getValues();
  const rosterHeaders = rosterValues.shift();
  const lastNameIndex = rosterHeaders.indexOf(ROSTER_CONFIG.HEADER_LAST_NAME);
  const firstNameIndex = rosterHeaders.indexOf(ROSTER_CONFIG.HEADER_FIRST_NAME);
  const nicknameIndex = rosterHeaders.indexOf(ROSTER_CONFIG.HEADER_NICKNAME);
  const emailIndex = rosterHeaders.indexOf(ROSTER_CONFIG.HEADER_EMAIL);
  if (lastNameIndex === -1 || firstNameIndex === -1 || emailIndex === -1) { throw new Error(`名簿シート「${rosterSheetName}」に必要な列が見つかりません。`); }
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
  const admin = require('./admin.gs');
  const reactions = require('./reactions.gs');

  ({
    safeGetUserEmail,
    getAdminEmails,
    isUserAdmin,
    checkAdmin,
    getSheets,
    getAppSettings,
    getAdminSettings,
    publishApp,
    unpublishApp,
    setShowDetails,
    saveWebAppUrl,
    getWebAppUrl,
    saveDeployId
  } = admin);

  ({ parseReactionString, addReaction, toggleHighlight } = reactions);

  module.exports = {
    COLUMN_HEADERS,
    TIME_CONSTANTS,
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
