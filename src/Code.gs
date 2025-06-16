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
const HEADERS_CACHE_TTL = 1800; // seconds
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
  DEPLOY_ID: 'DEPLOY_ID',
  ADMIN_EMAILS: 'ADMIN_EMAILS',
  REACTION_COUNT_ENABLED: 'REACTION_COUNT_ENABLED',
  SCORE_SORT_ENABLED: 'SCORE_SORT_ENABLED'
};

/**
 * Save multiple script properties at once.
 * If a value is null or undefined the property is removed.
 * @param {Object<string, any>} settings Key-value pairs to store.
 */
function saveSettings(settings) {
  const props = PropertiesService.getScriptProperties();
  Object.keys(settings || {}).forEach(key => {
    const value = settings[key];
    if (value === null || value === undefined) {
      props.deleteProperty(key);
    } else {
      props.setProperty(key, String(value));
    }
  });
}

/**
 * 管理パネルの初期化に必要なデータを取得します。
 * @returns {object} - 現在の公開状態とシートのリスト
 */
function getAdminSettings() {
  if (!isUserAdmin()) {
    throw new Error('管理者のみ実行できます。');
  }
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
    currentUserEmail: currentUser,
    deployId: properties.getProperty(APP_PROPERTIES.DEPLOY_ID) || '',
    reactionCountEnabled: properties.getProperty(APP_PROPERTIES.REACTION_COUNT_ENABLED) === 'true',
    scoreSortEnabled: properties.getProperty(APP_PROPERTIES.SCORE_SORT_ENABLED) === 'true'
  };
}

/**
 * 指定されたシートでアプリを公開します。
 * @param {string} sheetName - 公開するシート名。
 */
function publishApp(sheetName) {
  if (!isUserAdmin()) {
    throw new Error('管理者のみ実行できます。');
  }
  if (!sheetName) {
    throw new Error('シート名が指定されていません。');
  }
  saveSettings({
    [APP_PROPERTIES.IS_PUBLISHED]: 'true',
    [APP_PROPERTIES.ACTIVE_SHEET]: sheetName
  });
  return `「${sheetName}」を公開しました。`;
}

/**
 * アプリの公開を終了します。
 */
function unpublishApp() {
  if (!isUserAdmin()) {
    throw new Error('管理者のみ実行できます。');
  }
  saveSettings({
    [APP_PROPERTIES.IS_PUBLISHED]: 'false',
    [APP_PROPERTIES.ACTIVE_SHEET]: null
  });
  return 'アプリを非公開にしました。';
}

/**
* 管理者メールアドレスを保存します。
* @param {string|Array} emails - カンマ区切りのメールアドレス文字列または配列
*/
function saveAdminEmails(emails) {
 let value;
  if (Array.isArray(emails)) {
    value = emails.map(e => e.trim()).filter(Boolean).join(',');
  } else {
    value = (emails || '').split(',').map(e => e.trim()).filter(Boolean).join(',');
  }
 saveSettings({ [APP_PROPERTIES.ADMIN_EMAILS]: value });
  return '管理者メールアドレスを更新しました。';
}

function saveReactionCountSetting(enabled) {
  const value = enabled ? 'true' : 'false';
  saveSettings({ [APP_PROPERTIES.REACTION_COUNT_ENABLED]: value });
  return `リアクション数表示を${enabled ? '有効' : '無効'}にしました。`;
}

function saveScoreSortSetting(enabled) {
  const value = enabled ? 'true' : 'false';
  saveSettings({ [APP_PROPERTIES.SCORE_SORT_ENABLED]: value });
  return `スコア順ソートを${enabled ? '有効' : '無効'}にしました。`;
}

function saveDisplayMode(mode) {
  const value = mode === 'named' ? 'named' : 'anonymous';
  saveSettings({ [APP_PROPERTIES.DISPLAY_MODE]: value });
  return `表示モードを${value === 'named' ? '実名' : '匿名'}にしました。`;
}

function saveDeployId(id) {
  const value = id ? String(id).trim() : '';
  saveSettings({ [APP_PROPERTIES.DEPLOY_ID]: value });
  return 'デプロイIDを更新しました。';
}

function getAdminEmails() {
 const str = PropertiesService.getScriptProperties()
     .getProperty(APP_PROPERTIES.ADMIN_EMAILS) || '';
 return str.split(',').map(e => e.trim()).filter(Boolean);
}

function isUserAdmin() {
  const admins = getAdminEmails();
  let email = '';
  try {
    email = Session.getActiveUser().getEmail();
  } catch (e) {}
  return admins.includes(email);
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
  const userIsAdmin = adminEmails.includes(userEmail);
  const view = e && e.parameter && e.parameter.view;
  const forceAdmin = e && e.parameter && e.parameter.mode === 'admin';
  const isAdmin = userIsAdmin && forceAdmin;

  if (!settings.isPublished && !(userIsAdmin && view === 'board')) {
    const template = HtmlService.createTemplateFromFile('Unpublished');
    template.userEmail = userEmail;
    template.isAdmin = userIsAdmin;
    return template.evaluate().setTitle('公開終了');
  }

  if (!settings.activeSheetName) {
    return HtmlService.createHtmlOutput('エラー: 表示するシートが設定されていません。スプレッドシートの「アプリ管理」メニューから設定してください。').setTitle('エラー');
  }

  // ★変更: Page.html にもメールアドレスを渡す
  const template = HtmlService.createTemplateFromFile('Page');
  template.userEmail = userEmail; // この行を追加
  template.isAdmin = isAdmin;
  template.showCounts = settings.reactionCountEnabled;
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

function getPublishedSheetData(classFilter, sortMode, adminOverride) {
  sortMode = sortMode || 'newest';
  const settings = getAppSettings();
  const sheetName = settings.activeSheetName;

  if (!sheetName) {
    throw new Error('表示するシートが設定されていません。');
  }

  // 既存のgetSheetDataロジックを再利用
  const data = getSheetData(sheetName, classFilter, sortMode, adminOverride);

  // ★改善: フロントエンドでシート名を表示できるよう、レスポンスに含める
  return {
    sheetName: sheetName,
    header: data.header,
    rows: data.rows
  };
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
  saveSettings({ [APP_PROPERTIES.WEB_APP_URL]: url });
}

function getWebAppUrlFromProps() {
  return PropertiesService.getScriptProperties().getProperty(APP_PROPERTIES.WEB_APP_URL) || '';
}

function getSheetUpdates(classFilter, sortMode, clientHashesJson, adminOverride) {
  const settings = getAppSettings();
  const sheetName = settings.activeSheetName;
  if (!sheetName) {
    throw new Error('表示するシートが設定されていません。');
  }
  const data = getSheetData(sheetName, classFilter, sortMode, adminOverride);
  const clientMap = clientHashesJson ? JSON.parse(clientHashesJson) : {};
  const changedRows = [];
  const newMap = {};
  data.rows.forEach(row => {
    const hash = Utilities.base64Encode(
      Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(row))
    );
    newMap[row.rowIndex] = hash;
    if (clientMap[String(row.rowIndex)] !== hash) {
      changedRows.push(Object.assign({ hash: hash }, row));
    }
  });
  const removedRows = Object.keys(clientMap)
    .filter(k => !newMap[k])
    .map(k => parseInt(k, 10));
  return {
    sheetName: sheetName,
    header: data.header,
    changedRows: changedRows,
    removedRows: removedRows,
    hashMap: newMap,
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

function getSheetData(sheetName, classFilter, sortMode, adminOverride) {
  sortMode = sortMode || 'newest';
  const isAdmin = typeof adminOverride === 'boolean' ? adminOverride : isUserAdmin();
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) throw new Error(`指定されたシート「${sheetName}」が見つかりません。`);

    const allValues = sheet.getDataRange().getValues();
    if (allValues.length < 1) return { header: "シートにデータがありません", rows: [] };
    
  const userEmail = Session.getActiveUser().getEmail();
  const headerIndices = getAndCacheHeaderIndices(sheetName, allValues[0]);
  const dataRows = allValues.slice(1);

  let displayMode = PropertiesService.getScriptProperties()
      .getProperty(APP_PROPERTIES.DISPLAY_MODE) || 'anonymous';
  if (!isAdmin) {
    displayMode = 'anonymous';
  }

  const emailToNameMap = displayMode === 'named' ? getRosterMap() : {};

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
        const actualName =
          displayMode === 'named' && emailToNameMap[email]
            ? emailToNameMap[email]
            : email.split('@')[0];
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
    const reactionHeaders = [COLUMN_HEADERS.UNDERSTAND, COLUMN_HEADERS.LIKE, COLUMN_HEADERS.CURIOUS];
    const headerIndices = findHeaderIndices(headerRow, reactionHeaders);
    const colIndices = {
      UNDERSTAND: headerIndices[COLUMN_HEADERS.UNDERSTAND] + 1,
      LIKE: headerIndices[COLUMN_HEADERS.LIKE] + 1,
      CURIOUS: headerIndices[COLUMN_HEADERS.CURIOUS] + 1,
    };

    const selectedHeader = COLUMN_HEADERS[reactionKey];
    const selectedCol = colIndices[reactionKey];

    // Retrieve current lists for all reactions
    const lists = {};
    Object.keys(colIndices).forEach(key => {
      const cell = sheet.getRange(rowIndex, colIndices[key]);
      const str = cell.getValue().toString();
      lists[key] = { cell: cell, arr: str ? str.split(',').filter(Boolean) : [] };
    });

    const wasReacted = lists[reactionKey].arr.includes(userEmail);

    // Remove user from all reactions
    Object.keys(lists).forEach(key => {
      const idx = lists[key].arr.indexOf(userEmail);
      if (idx > -1) lists[key].arr.splice(idx, 1);
    });

    // If user wasn't toggling off the same reaction, add them to selected
    if (!wasReacted) {
      lists[reactionKey].arr.push(userEmail);
    }

    // Write back all values
    Object.keys(lists).forEach(key => {
      lists[key].cell.setValue(lists[key].arr.join(','));
    });

    const reactions = {};
    Object.keys(lists).forEach(key => {
      reactions[key] = {
        count: lists[key].arr.length,
        reacted: lists[key].arr.includes(userEmail)
      };
    });
    return { status: 'ok', reactions: reactions };
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
  if (!isUserAdmin()) {
    throw new Error('管理者のみ実行できます。');
  }
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
    activeSheetName: properties.getProperty(APP_PROPERTIES.ACTIVE_SHEET),
    reactionCountEnabled: properties.getProperty(APP_PROPERTIES.REACTION_COUNT_ENABLED) === 'true',
    scoreSortEnabled: properties.getProperty(APP_PROPERTIES.SCORE_SORT_ENABLED) === 'true'
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
  cache.put(cacheKey, JSON.stringify(indices), HEADERS_CACHE_TTL);
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

// Export for Jest testing
if (typeof module !== 'undefined') {
  module.exports = {
    COLUMN_HEADERS,
    findHeaderIndices,
    getSheetData,
    getAdminSettings,
    doGet,
    addReaction,
    toggleHighlight,
    saveReactionCountSetting,
    saveScoreSortSetting,
    getAppSettings,
  getWebAppUrl,
  saveWebAppUrl,
  getWebAppUrlFromProps,
  getSheetUpdates,
  saveDisplayMode,
  saveDeployId,
  };
}
