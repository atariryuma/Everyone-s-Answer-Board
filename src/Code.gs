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
  IS_PUBLISHED: 'IS_PUBLISHED'
};

const TIME_CONSTANTS = {
  LOCK_WAIT_MS: 10000
};

const REACTION_KEYS = ["UNDERSTAND","LIKE","CURIOUS"];

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
    adminEmails: adminEmails,
    isUserAdmin: adminEmails.includes(currentUserEmail),
    isPublished: appSettings.isPublished,
    activeSheetName: appSettings.activeSheetName
  };
}

/**
 * 指定されたシートでアプリを公開します。
 * @param {string} sheetName - 公開するシート名。
 */
function publishApp(sheetName) {
  if (!sheetName) {
    throw new Error('シート名が指定されていません。');
  }
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty(APP_PROPERTIES.IS_PUBLISHED, 'true');
  properties.setProperty(APP_PROPERTIES.ACTIVE_SHEET, sheetName);
  return `「${sheetName}」を公開しました。`;
}

/**
 * アプリの公開を終了します。
 */
function unpublishApp() {
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty(APP_PROPERTIES.IS_PUBLISHED, 'false');
  return 'アプリを非公開にしました。';
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

  const userEmail = safeGetUserEmail();

  const template = HtmlService.createTemplateFromFile('Page');
  const admin = isUserAdmin(userEmail);
  Object.assign(template, {
    showAdminFeatures: false,
    showHighlightToggle: false,
    isAdminUser: admin
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
    rows: data.rows
  };
}


// =================================================================
// 内部処理関数
// =================================================================
function getSheets() {
  try {
    const allSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    const visibleSheets = allSheets.filter(sheet => !sheet.isSheetHidden());
    return visibleSheets.map(sheet => sheet.getName());
  } catch (error) {
    console.error('getSheets Error:', error);
    throw new Error('シート一覧の取得に失敗しました。');
  }
}

function getSheetData(sheetName, classFilter, sortBy) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) throw new Error(`指定されたシート「${sheetName}」が見つかりません。`);

    const allValues = sheet.getDataRange().getValues();
    if (allValues.length < 1) return { header: "シートにデータがありません", rows: [] };
    
    const userEmail = safeGetUserEmail();
    const headerIndices = getAndCacheHeaderIndices(sheetName, allValues[0]);
    const dataRows = allValues.slice(1);
    const emailToNameMap = getRosterMap();

    const rows = dataRows.map((row, index) => {
      if (classFilter && classFilter !== 'すべて') {
        const className = row[headerIndices[COLUMN_HEADERS.CLASS]];
        if (className !== classFilter) return null;
      }
      const email = row[headerIndices[COLUMN_HEADERS.EMAIL]];
      const opinion = row[headerIndices[COLUMN_HEADERS.OPINION]];

      if (email && opinion) {
        const understandArr = parseReactionString(row[headerIndices[COLUMN_HEADERS.UNDERSTAND]]);
        const likeArr = parseReactionString(row[headerIndices[COLUMN_HEADERS.LIKE]]);
        const curiousArr = parseReactionString(row[headerIndices[COLUMN_HEADERS.CURIOUS]]);
        const reason = row[headerIndices[COLUMN_HEADERS.REASON]] || '';
        const highlightVal = row[headerIndices[COLUMN_HEADERS.HIGHLIGHT]];
        const likes = likeArr.length;
        const baseScore = reason.length;
        const likeMultiplier = 1 + (likes * SCORING_CONFIG.LIKE_MULTIPLIER_FACTOR);
        const totalScore = baseScore * likeMultiplier;
        const name = emailToNameMap[email] || email.split('@')[0];
        return {
          rowIndex: index + 2,
          name: name,
          class: row[headerIndices[COLUMN_HEADERS.CLASS]] || '未分類',
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
    return { header: COLUMN_HEADERS.OPINION, rows: rows };
  } catch(e) {
    console.error(`getSheetData Error for sheet "${sheetName}":`, e);
    throw new Error(`データの取得中にエラーが発生しました: ${e.message}`);
  }
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

  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const reactionHeaders = REACTION_KEYS.map(k => COLUMN_HEADERS[k]);
  const headerIndices = getAndCacheHeaderIndices(settings.activeSheetName, headerRow);
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
    console.error('addReaction Error:', error);
    return { status: 'error', message: `エラーが発生しました: ${error.message}` };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function toggleHighlight(rowIndex) {
  const lock = LockService.getScriptLock();
  lock.waitLock(TIME_CONSTANTS.LOCK_WAIT_MS);
  try {
    const sheetName = getAppSettings().activeSheetName;
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) throw new Error(`シート '${sheetName}' が見つかりません。`);

    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const headerIndices = getAndCacheHeaderIndices(sheetName, headerRow);
    const colIndex = headerIndices[COLUMN_HEADERS.HIGHLIGHT] + 1;

    const cell = sheet.getRange(rowIndex, colIndex);
    const current = !!cell.getValue();
    const newValue = !current;
    cell.setValue(newValue);
    return { status: 'ok', highlight: newValue };
  } catch (error) {
    console.error('toggleHighlight Error:', error);
    return { status: 'error', message: `エラーが発生しました: ${error.message}` };
  } finally {
    lock.releaseLock();
  }
}

function addLike(rowIndex) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const userEmail = safeGetUserEmail();
    if (!userEmail) return { status: 'error', message: 'ログインしていないため、操作できません。' };
    const settings = getAppSettings();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(settings.activeSheetName);
    if (!sheet) throw new Error(`シート '${settings.activeSheetName}' が見つかりません。`);

    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const headerIndices = getAndCacheHeaderIndices(settings.activeSheetName, headerRow);
    const likesColIndex = headerIndices[COLUMN_HEADERS.LIKE] + 1;

    const likeCell = sheet.getRange(rowIndex, likesColIndex);
    const likersString = likeCell.getValue().toString();
    let likers = likersString ? likersString.split(',').filter(Boolean) : [];
    const userIndex = likers.indexOf(userEmail);
    if (userIndex > -1) { likers.splice(userIndex, 1); } else { likers.push(userEmail); }
    likeCell.setValue(likers.join(','));
    return { status: 'ok', newScore: likers.length };
  } catch (error) {
    console.error('addLike Error:', error);
    return { status: 'error', message: `エラーが発生しました: ${error.message}` };
  } finally {
    lock.releaseLock();
  }
}

function getAppSettings() {
  const properties = PropertiesService.getScriptProperties() || {};
  const getProp = typeof properties.getProperty === 'function' ? (k) => properties.getProperty(k) : () => null;
  const published = getProp(APP_PROPERTIES.IS_PUBLISHED);
  const sheet = getProp(APP_PROPERTIES.ACTIVE_SHEET);
  return {
    isPublished: published === null ? true : published === 'true',
    activeSheetName: sheet || 'Sheet1'
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

function getAndCacheHeaderIndices(sheetName, headerRow) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `headers_${sheetName}`;
  const cachedHeaders = cache.get(cacheKey);
  if (cachedHeaders) { return JSON.parse(cachedHeaders); }
  const indices = findHeaderIndices(headerRow, Object.values(COLUMN_HEADERS));
  cache.put(cacheKey, JSON.stringify(indices), 21600);
  return indices;
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

function saveDeployId(id) {
  const props = PropertiesService.getScriptProperties();
  props.setProperties({ DEPLOY_ID: (id || '').trim() });
}

if (typeof module !== 'undefined') {
  module.exports = {
    COLUMN_HEADERS,
    getAdminSettings,
    publishApp,
    unpublishApp,
    doGet,
    getPublishedSheetData,
    getSheets,
    getSheetData,
    getRosterMap,
    addReaction,
    addLike,
    toggleHighlight,
    saveWebAppUrl,
    saveDeployId,
    findHeaderIndices,
    parseReactionString,
    checkAdmin
  };
}
