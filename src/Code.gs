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
  LIKES: 'いいね！'
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
  const properties = PropertiesService.getScriptProperties();
  const allSheets = getSheets(); // 既存の関数を再利用
  return {
    isPublished: properties.getProperty(APP_PROPERTIES.IS_PUBLISHED) === 'true',
    activeSheetName: properties.getProperty(APP_PROPERTIES.ACTIVE_SHEET),
    allSheets: allSheets
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
function doGet() {
  const settings = getAppSettings();
  
  if (!settings.isPublished) {
    const userEmail = Session.getActiveUser().getEmail();
    const template = HtmlService.createTemplateFromFile('Unpublished');
    template.userEmail = userEmail;
    return template.evaluate().setTitle('公開終了');
  }
  
  if (!settings.activeSheetName) {
    return HtmlService.createHtmlOutput('エラー: 表示するシートが設定されていません。スプレッドシートの「アプリ管理」メニューから設定してください。').setTitle('エラー');
  }

  // ★変更: Page.html にもメールアドレスを渡す
  const template = HtmlService.createTemplateFromFile('Page');
  template.userEmail = Session.getActiveUser().getEmail(); // この行を追加
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

function getPublishedSheetData(classFilter) {
  const settings = getAppSettings();
  const sheetName = settings.activeSheetName;
  
  if (!sheetName) {
    throw new Error('表示するシートが設定されていません。');
  }
  
  // 既存のgetSheetDataロジックを再利用
  const data = getSheetData(sheetName, classFilter);

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

function getSheetData(sheetName, classFilter) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) throw new Error(`指定されたシート「${sheetName}」が見つかりません。`);

    const allValues = sheet.getDataRange().getValues();
    if (allValues.length < 1) return { header: "シートにデータがありません", rows: [] };
    
    const userEmail = Session.getActiveUser().getEmail();
    const headerIndices = getAndCacheHeaderIndices(sheetName, allValues[0]);
    const dataRows = allValues.slice(1);
    const emailToNameMap = getRosterMap();

    const filteredRows = dataRows.filter(row => {
      if (!classFilter || classFilter === 'すべて') return true;
      const className = row[headerIndices[COLUMN_HEADERS.CLASS]];
      return className === classFilter;
    });

    const rows = filteredRows.map((row) => {
      const email = row[headerIndices[COLUMN_HEADERS.EMAIL]];
      const opinion = row[headerIndices[COLUMN_HEADERS.OPINION]];

      if (email && opinion) {
        const likersString = row[headerIndices[COLUMN_HEADERS.LIKES]] || '';
        const likers = likersString ? likersString.toString().split(',').filter(Boolean) : [];
        const reason = row[headerIndices[COLUMN_HEADERS.REASON]] || '';
        const likes = likers.length;
        const baseScore = reason.length;
        const likeMultiplier = 1 + (likes * SCORING_CONFIG.LIKE_MULTIPLIER_FACTOR);
        const totalScore = baseScore * likeMultiplier;
        return {
          rowIndex: dataRows.indexOf(row) + 2,
          name: emailToNameMap[email] || email.split('@')[0],
          class: row[headerIndices[COLUMN_HEADERS.CLASS]] || '未分類',
          opinion: opinion,
          reason: reason,
          likes: likes,
          hasLiked: userEmail ? likers.includes(userEmail) : false,
          score: totalScore
        };
      }
      return null;
    }).filter(Boolean);

    rows.sort((a, b) => b.score - a.score);
    return { header: COLUMN_HEADERS.OPINION, rows: rows };
  } catch(e) {
    console.error(`getSheetData Error for sheet "${sheetName}":`, e);
    throw new Error(`データの取得中にエラーが発生しました: ${e.message}`);
  }
}

function addLike(rowIndex) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) return { status: 'error', message: 'ログインしていないため、操作できません。' };
    const settings = getAppSettings();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(settings.activeSheetName);
    if (!sheet) throw new Error(`シート '${settings.activeSheetName}' が見つかりません。`);

    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const headerIndices = findHeaderIndices(headerRow, [COLUMN_HEADERS.LIKES]);
    const likesColIndex = headerIndices[COLUMN_HEADERS.LIKES] + 1;

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
  const properties = PropertiesService.getScriptProperties();
  return {
    isPublished: properties.getProperty(APP_PROPERTIES.IS_PUBLISHED) === 'true',
    activeSheetName: properties.getProperty(APP_PROPERTIES.ACTIVE_SHEET)
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
  const trimmedSheetHeaders = sheetHeaders.map(h => (typeof h === 'string' ? h.trim() : h));
  const missingHeaders = [];
  requiredHeaders.forEach(headerName => {
    const index = trimmedSheetHeaders.indexOf(headerName);
    if (index !== -1) { indices[headerName] = index; } else { missingHeaders.push(headerName); }
  });
  if (missingHeaders.length > 0) { throw new Error(`必須ヘッダーが見つかりません: [${missingHeaders.join(', ')}]`); }
  return indices;
}

// Export for Jest testing
if (typeof module !== 'undefined') {
  module.exports = { findHeaderIndices };
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
