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
  HIGHLIGHT: 'Highlight'
};
const CACHE_KEYS = {
  ROSTER: 'roster_name_map_v3',
  HEADERS: (sheetName) => `headers_${sheetName}_v1`
};
const ROSTER_CONFIG = {
  SHEET_NAME: 'sheet 1',
  HEADER_LAST_NAME: '姓',
  HEADER_FIRST_NAME: '名',
  HEADER_NICKNAME: 'ニックネーム',
  HEADER_EMAIL: 'Googleアカウント'
};
const SCORING_CONFIG = {
  LIKE_MULTIPLIER_FACTOR: 0.05 // 各リアクション1件ごとにスコアが5%増加
};
const APP_PROPERTIES = {
  ACTIVE_SHEET: 'ACTIVE_SHEET_NAME',
  IS_PUBLISHED: 'IS_PUBLISHED',
  DISPLAY_MODE: 'DISPLAY_MODE',
  WEB_APP_URL: 'WEB_APP_URL',
  ADMIN_EMAILS: 'ADMIN_EMAILS'
};


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

  // スプレッドシートを開いたら自動で管理パネルを表示
  showAdminSidebar();
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
  const properties = PropertiesService.getScriptProperties();
  const allSheets = getSheets(); // 既存の関数を再利用
  const adminEmailsRaw = properties.getProperty(APP_PROPERTIES.ADMIN_EMAILS) || '';
  const adminEmails = adminEmailsRaw ? adminEmailsRaw.split(',').map(e => e.trim()).filter(Boolean) : [];
  let currentUser = '';
  try {
   currentUser = Session.getActiveUser().getEmail();
  } catch (e) {}
  return {
    isPublished: properties.getProperty(APP_PROPERTIES.IS_PUBLISHED) === 'true',
    activeSheetName: properties.getProperty(APP_PROPERTIES.ACTIVE_SHEET),
    allSheets: allSheets,
    displayMode: properties.getProperty(APP_PROPERTIES.DISPLAY_MODE) || 'anonymous',
    adminEmails: adminEmails,
    currentUserEmail: currentUser
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
  logDebug(`Published sheet: ${sheetName}`);
  return `「${sheetName}」を公開しました。`;
}

/**
 * アプリの公開を終了します。
 */
function unpublishApp() {
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty(APP_PROPERTIES.IS_PUBLISHED, 'false');
  properties.deleteProperty(APP_PROPERTIES.ACTIVE_SHEET);
  logDebug('Unpublished app');
  return 'アプリを非公開にしました。';
}

/**
 * 表示モードを保存します。
 * @param {string} mode - 'anonymous' または 'named'
 */
function saveDisplayMode(mode) {
  const properties = PropertiesService.getScriptProperties();
  const value = mode === 'named' ? 'named' : 'anonymous';
  properties.setProperty(APP_PROPERTIES.DISPLAY_MODE, value);
  logDebug(`Display mode set to ${value}`);
  return `表示モードを${value === 'named' ? '記名' : '匿名'}に設定しました。`;
}

/**
* 管理者メールアドレスを保存します。
* @param {string|Array} emails - カンマ区切りのメールアドレス文字列または配列
*/
function saveAdminEmails(emails) {
 const properties = PropertiesService.getScriptProperties();
 let value;
 if (Array.isArray(emails)) {
   value = emails.map(e => e.trim()).filter(Boolean).join(',');
 } else {
   value = (emails || '').split(',').map(e => e.trim()).filter(Boolean).join(',');
 }
 properties.setProperty(APP_PROPERTIES.ADMIN_EMAILS, value);
 logDebug(`Admin emails updated: ${value}`);
 return '管理者メールアドレスを更新しました。';
}

function getAdminEmails() {
 const str = PropertiesService.getScriptProperties()
     .getProperty(APP_PROPERTIES.ADMIN_EMAILS) || '';
 return str.split(',').map(e => e.trim()).filter(Boolean);
}



// =================================================================
// GAS Webアプリケーションのエントリーポイント
// =================================================================
function doGet(e) {
  const settings = getAppSettings();
  let userEmail;
  try {
    userEmail = Session.getActiveUser().getEmail();
  } catch (e) {
    userEmail = '匿名ユーザー';
  }

  const adminEmails = getAdminEmails();
  const isAdmin =
      e && e.parameter && e.parameter.admin === '1' &&
      adminEmails.includes(userEmail);


  if (isAdmin && e && e.parameter && e.parameter.groups === '1') {
    const t = HtmlService.createTemplateFromFile('OpinionGroups');
    return t.evaluate()
            .setTitle('意見のグループ化')
            .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  if (!settings.isPublished) {
    const template = HtmlService.createTemplateFromFile('Unpublished');
    template.userEmail = userEmail;
    template.isAdmin = isAdmin;
    return template.evaluate().setTitle('公開終了');
  }
  
  if (!settings.activeSheetName) {
    return HtmlService.createHtmlOutput('エラー: 表示するシートが設定されていません。スプレッドシートの「アプリ管理」メニューから設定してください。').setTitle('エラー');
  }

  // ★変更: Page.html にもメールアドレスを渡す
  const template = HtmlService.createTemplateFromFile('Page');
  template.userEmail = userEmail; // この行を追加
  template.isAdmin = isAdmin;
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

function getPublishedSheetData(classFilter, sortMode, isAdmin) {
  sortMode = sortMode || 'newest';
  const settings = getAppSettings();
  const sheetName = settings.activeSheetName;

  if (!sheetName) {
    throw new Error('表示するシートが設定されていません。');
  }

  // 既存のgetSheetDataロジックを再利用
  const data = getSheetData(sheetName, classFilter, sortMode, !!isAdmin);

  // ★改善: フロントエンドでシート名を表示できるよう、レスポンスに含める
  return {
    sheetName: sheetName,
    header: data.header,
    rows: data.rows
  };
}


function groupSimilarOpinions() {
  const settings = getAppSettings();
  const sheetName = settings.activeSheetName;
  if (!sheetName) throw new Error('表示するシートが設定されていません。');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`シート '${sheetName}' が見つかりません。`);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return '';
  const idx = findHeaderIndices(values[0], [COLUMN_HEADERS.OPINION, COLUMN_HEADERS.REASON]);
  const texts = values.slice(1).map(r => {
    const op = r[idx[COLUMN_HEADERS.OPINION]];
    const reason = r[idx[COLUMN_HEADERS.REASON]];
    return `意見: ${op} 理由: ${reason}`;
  }).filter(Boolean);

  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('GEMINI APIキーが設定されていません。');

  const prompt = '次の意見と理由を類似内容ごとにグループ化し、多数派と少数派を含めて要約してください。\n' + texts.join('\n');
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent?key=' + apiKey;
  const payload = {
    contents: [{ parts: [{ text: prompt }] }]
  };
  try {
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    const json = JSON.parse(res.getContentText() || '{}');
    return json.candidates && json.candidates[0] && json.candidates[0].content && json.candidates[0].content.parts ? json.candidates[0].content.parts[0].text : '';
  } catch (e) {
    console.error('groupSimilarOpinions error', e);
    return '';
  }
}

function getWebAppUrl() {
  try {
    return ScriptApp.getService().getUrl();
  } catch (e) {
    return '';
  }
}

function saveWebAppUrl(url) {
  if (!url) {
    try {
      url = ScriptApp.getService().getUrl();
    } catch (e) {
      url = '';
    }
  }
  PropertiesService.getScriptProperties().setProperty(APP_PROPERTIES.WEB_APP_URL, url);
}

function getWebAppUrlFromProps() {
  return PropertiesService.getScriptProperties().getProperty(APP_PROPERTIES.WEB_APP_URL) || '';
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

function getSheetData(sheetName, classFilter, sortMode, isAdmin) {
  sortMode = sortMode || 'newest';
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) throw new Error(`指定されたシート「${sheetName}」が見つかりません。`);

    const allValues = sheet.getDataRange().getValues();
    if (allValues.length < 1) return { header: "シートにデータがありません", rows: [] };
    
  const userEmail = Session.getActiveUser().getEmail();
  const headerIndices = getAndCacheHeaderIndices(sheetName, allValues[0]);
  const dataRows = allValues.slice(1);
  const emailToNameMap = getRosterMap();
  let displayMode = PropertiesService.getScriptProperties()
      .getProperty(APP_PROPERTIES.DISPLAY_MODE) || 'anonymous';
  if (!isAdmin) {
    displayMode = 'anonymous';
  }

    const filteredRows = dataRows.filter(row => {
      if (!classFilter || classFilter === 'すべて') return true;
      const className = row[headerIndices[COLUMN_HEADERS.CLASS]];
      return className === classFilter;
    });

    const rows = filteredRows.map((row, i) => {
      const email = row[headerIndices[COLUMN_HEADERS.EMAIL]];
      const opinion = row[headerIndices[COLUMN_HEADERS.OPINION]];

      if (email && opinion) {
        const understandStr = row[headerIndices[COLUMN_HEADERS.UNDERSTAND]] || '';
        const likeStr = row[headerIndices[COLUMN_HEADERS.LIKE]] || '';
        const curiousStr = row[headerIndices[COLUMN_HEADERS.CURIOUS]] || '';
        const understandArr = understandStr ? understandStr.toString().split(',').filter(Boolean) : [];
        const likeArr = likeStr ? likeStr.toString().split(',').filter(Boolean) : [];
        const curiousArr = curiousStr ? curiousStr.toString().split(',').filter(Boolean) : [];
        const reason = row[headerIndices[COLUMN_HEADERS.REASON]] || '';
        const highlightVal = row[headerIndices[COLUMN_HEADERS.HIGHLIGHT]];
        const highlight = String(highlightVal).toLowerCase() === 'true';

        const reactions = {
          UNDERSTAND: { count: understandArr.length, reacted: userEmail ? understandArr.includes(userEmail) : false },
          LIKE: { count: likeArr.length, reacted: userEmail ? likeArr.includes(userEmail) : false },
          CURIOUS: { count: curiousArr.length, reacted: userEmail ? curiousArr.includes(userEmail) : false },
        };

        const totalReactions = reactions.UNDERSTAND.count + reactions.LIKE.count + reactions.CURIOUS.count;
        const baseScore = reason.length;
        const likeMultiplier = 1 + (totalReactions * SCORING_CONFIG.LIKE_MULTIPLIER_FACTOR);
        const totalScore = baseScore * likeMultiplier;
        const actualName = emailToNameMap[email] || email.split('@')[0];
        const name = displayMode === 'named' ? actualName : '匿名';
        return {
          rowIndex: i + 2,
          name: name,
          class: row[headerIndices[COLUMN_HEADERS.CLASS]] || '未分類',
          opinion: opinion,
          reason: reason,
          reactions: reactions,
          highlight: highlight,
          score: totalScore
        };
      }
      return null;
    }).filter(Boolean);

    switch (sortMode) {
      case 'newest':
        rows.sort((a, b) => b.rowIndex - a.rowIndex);
        break;
      case 'random':
        rows.sort(() => Math.random() - 0.5);
        break;
      default:
        rows.sort((a, b) => b.score - a.score);
        break;
    }
    return { header: COLUMN_HEADERS.OPINION, rows: rows };
  } catch(e) {
    console.error(`getSheetData Error for sheet "${sheetName}":`, e);
    throw new Error(`データの取得中にエラーが発生しました: ${e.message}`);
  }
}

function addReaction(rowIndex, reactionKey) {
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(10000)) {
      return { status: 'error', message: '他のユーザーが操作中です。しばらく待ってから再試行してください。' };
    }
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) return { status: 'error', message: 'ログインしていないため、操作できません。' };
    const settings = getAppSettings();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(settings.activeSheetName);
    if (!sheet) throw new Error(`シート '${settings.activeSheetName}' が見つかりません。`);

    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (!COLUMN_HEADERS[reactionKey]) {
      throw new Error(`Unknown reaction: ${reactionKey}`);
    }
    const headerIndices = findHeaderIndices(headerRow, [COLUMN_HEADERS[reactionKey]]);
    const colIndex = headerIndices[COLUMN_HEADERS[reactionKey]] + 1;

    const cell = sheet.getRange(rowIndex, colIndex);
    const likersString = cell.getValue().toString();
    let likers = likersString ? likersString.split(',').filter(Boolean) : [];
    const userIndex = likers.indexOf(userEmail);
    if (userIndex > -1) { likers.splice(userIndex, 1); } else { likers.push(userEmail); }
    cell.setValue(likers.join(','));
    return { status: 'ok', reaction: reactionKey, newScore: likers.length };
  } catch (error) {
    console.error('addReaction Error:', error);
    return { status: 'error', message: `エラーが発生しました: ${error.message}` };
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {
      // already released
    }
  }
}

function toggleHighlight(rowIndex) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const settings = getAppSettings();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(settings.activeSheetName);
    if (!sheet) throw new Error(`シート '${settings.activeSheetName}' が見つかりません。`);

    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const headerIndices = findHeaderIndices(headerRow, [COLUMN_HEADERS.HIGHLIGHT]);
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

function getAppSettings() {
  const properties = PropertiesService.getScriptProperties();
  return {
    isPublished: properties.getProperty(APP_PROPERTIES.IS_PUBLISHED) === 'true',
    activeSheetName: properties.getProperty(APP_PROPERTIES.ACTIVE_SHEET)
  };
}

function getRosterMap() {
  const cache = CacheService.getScriptCache();
  const cachedMap = cache.get(CACHE_KEYS.ROSTER);
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
  cache.put(CACHE_KEYS.ROSTER, JSON.stringify(nameMap), 21600);
  return nameMap;
}

function getAndCacheHeaderIndices(sheetName, headerRow) {
  const cache = CacheService.getScriptCache();
  const cacheKey = CACHE_KEYS.HEADERS(sheetName);
  const cachedHeaders = cache.get(cacheKey);
  if (cachedHeaders) { return JSON.parse(cachedHeaders); }
  const indices = findHeaderIndices(headerRow, Object.values(COLUMN_HEADERS));
  cache.put(cacheKey, JSON.stringify(indices), 21600);
  return indices;
}

function findHeaderIndices(sheetHeaders, requiredHeaders) {
  const indices = {};
  const normalize = h => (typeof h === 'string' ? h.replace(/\s+/g, '') : h);
  const normalizedHeaders = sheetHeaders.map(normalize);
  const missingHeaders = [];
  requiredHeaders.forEach(headerName => {
    const index = normalizedHeaders.indexOf(normalize(headerName));
    if (index !== -1) {
      indices[headerName] = index;
    } else {
      missingHeaders.push(headerName);
    }
  });
  if (missingHeaders.length > 0) {
    throw new Error(`必須ヘッダーが見つかりません: [${missingHeaders.join(', ')}]`);
  }
  return indices;
}

function logDebug(message) {
  if (typeof PropertiesService === 'undefined') return;
  try {
    const props = PropertiesService.getScriptProperties();
    const raw = props.getProperty('DEBUG_LOG') || '[]';
    const logs = JSON.parse(raw);
    logs.push(`${new Date().toISOString()} ${message}`);
    while (logs.length > 200) logs.shift();
    props.setProperty('DEBUG_LOG', JSON.stringify(logs));
  } catch (e) {}
}


// Export for Jest testing
if (typeof module !== 'undefined') {
  module.exports = {
    COLUMN_HEADERS,
    findHeaderIndices,
    getSheetData,
    addReaction,
    toggleHighlight,
    groupSimilarOpinions,
    logDebug,
    getWebAppUrl,
    saveWebAppUrl,
    getWebAppUrlFromProps,
  };
}

function clearRosterCache() {
  const cache = CacheService.getScriptCache();
  const cacheKey = CACHE_KEYS.ROSTER;
  cache.remove(cacheKey);
  console.log(`名簿キャッシュ（キー: ${cacheKey}）を削除しました。`);
  logDebug('Roster cache cleared');
  try {
    SpreadsheetApp.getUi().alert('名簿のキャッシュをリセットしました。');
  } catch (e) { /* no-op */ }
}

