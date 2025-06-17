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
  REACTION_COUNT_ENABLED: 'REACTION_COUNT_ENABLED',
  SCORE_SORT_ENABLED: 'SCORE_SORT_ENABLED',
  PUBLISH_TIMESTAMP: 'PUBLISH_TIMESTAMP'
};

/**
 * Safely get the active user's email with fallback
 * @returns {string} User email or empty string
 */
function safeGetUserEmail() {
  try {
    return Session.getActiveUser().getEmail() || '';
  } catch (error) {
    console.warn('Failed to get user email:', error);
    return '';
  }
}

// Legacy compatibility
this.getActiveUserEmail = safeGetUserEmail;

/**
 * Batch save script properties with null/undefined cleanup
 * @param {Object<string, any>} settings Key-value pairs to store
 */
function saveSettings(settings) {
  if (!settings || typeof settings !== 'object') return;
  
  const props = PropertiesService.getScriptProperties();
  const batch = {};
  const toDelete = [];
  
  Object.entries(settings).forEach(([key, value]) => {
    if (value === null || value === undefined) {
      toDelete.push(key);
    } else {
      batch[key] = String(value);
    }
  });
  
  // Batch operations for better performance
  if (Object.keys(batch).length > 0) props.setProperties(batch);
  toDelete.forEach(key => props.deleteProperty(key));
}

/**
 * 管理パネルの初期化に必要なデータを取得します。
 * @returns {object} - 現在の公開状態とシートのリスト
 */
function getAdminSettings() {
  const properties = PropertiesService.getScriptProperties();
  const allSheets = getSheets(); // 既存の関数を再利用
  let currentUser = '';
  try {
   currentUser = getActiveUserEmail();
  } catch (e) {}
  return {
    isPublished: properties.getProperty(APP_PROPERTIES.IS_PUBLISHED) === 'true',
    activeSheetName: properties.getProperty(APP_PROPERTIES.ACTIVE_SHEET),
    allSheets: allSheets,
    displayMode: properties.getProperty(APP_PROPERTIES.DISPLAY_MODE) || 'anonymous',
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
  if (!sheetName) {
    throw new Error('シート名が指定されていません。');
  }
  saveSettings({
    [APP_PROPERTIES.IS_PUBLISHED]: 'true',
    [APP_PROPERTIES.ACTIVE_SHEET]: sheetName,
    [APP_PROPERTIES.PUBLISH_TIMESTAMP]: Date.now()
  });
  return `「${sheetName}」を公開しました。`;
}

/**
 * アプリの公開を終了します。
 */
function unpublishApp(force) {
  saveSettings({
    [APP_PROPERTIES.IS_PUBLISHED]: 'false',
    [APP_PROPERTIES.ACTIVE_SHEET]: null,
    [APP_PROPERTIES.PUBLISH_TIMESTAMP]: null
  });
  return 'アプリを非公開にしました。';
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


function checkPublishExpiry() {
  const props = PropertiesService.getScriptProperties();
  const ts = parseInt(props.getProperty(APP_PROPERTIES.PUBLISH_TIMESTAMP) || '0', 10);
  const isPublished = props.getProperty(APP_PROPERTIES.IS_PUBLISHED) === 'true';
  if (isPublished && ts) {
    const age = Date.now() - ts;
    if (age > 6 * 60 * 60 * 1000) {
      try {
        unpublishApp(true);
      } catch (e) {
        console.error('Auto unpublish failed:', e);
      }
    }
  }
}

// =================================================================
// GAS Webアプリケーションのエントリーポイント
// =================================================================
function doGet(e) {
  checkPublishExpiry();
  const settings = getAppSettings();
  let userEmail;
  try {
    userEmail = getActiveUserEmail();
  } catch (e) {
    userEmail = '匿名ユーザー';
  }

  if (!settings.isPublished) {
    const template = HtmlService.createTemplateFromFile('Unpublished');
    template.userEmail = userEmail;
    return template.evaluate().setTitle('公開終了');
  }

  if (!settings.activeSheetName) {
    return HtmlService.createHtmlOutput('エラー: 表示するシートが設定されていません。スプレッドシートの「アプリ管理」メニューから設定してください。').setTitle('エラー');
  }

  const template = HtmlService.createTemplateFromFile('Page');
  template.userEmail = userEmail;
  template.showCounts = settings.reactionCountEnabled;
  template.scoreSortEnabled = settings.scoreSortEnabled;
  return template.evaluate()
      .setTitle('StudyQuest - みんなのかいとうボード')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// =================================================================
// クライアントサイドから呼び出される関数
// =================================================================

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

function getSheetUpdates(classFilter, sortMode, clientHashesJson) {
  const settings = getAppSettings();
  const sheetName = settings.activeSheetName;
  if (!sheetName) {
    throw new Error('表示するシートが設定されていません。');
  }
  const data = getSheetData(sheetName, classFilter, sortMode);
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

/**
 * Process and format sheet data with filtering and sorting
 * @param {string} sheetName Target sheet name
 * @param {string} classFilter Class filter ('すべて' for all)
 * @param {string} sortMode Sort mode ('newest', 'oldest', 'score')
 * @returns {Object} Processed sheet data
 */
function getSheetData(sheetName, classFilter = '', sortMode = 'newest') {
  if (!sheetName) {
    throw new Error('シート名が指定されていません');
  }
  
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`指定されたシート「${sheetName}」が見つかりません。`);
    }

    const allValues = sheet.getDataRange().getValues();
    if (allValues.length < 1) {
      return { header: "シートにデータがありません", rows: [] };
    }
    
    const userEmail = safeGetUserEmail();
    const headerIndices = getAndCacheHeaderIndices(sheetName, allValues[0]);
    const dataRows = allValues.slice(1);
    const displayMode = getDisplayMode();
    const emailToNameMap = displayMode === 'named' ? getRosterMap() : {};

    const filteredRows = filterRowsByClass(dataRows, classFilter, headerIndices);
    const rows = filteredRows
      .map((row, i) => processRowData(row, i, headerIndices, userEmail, emailToNameMap))
      .filter(Boolean);
    
    const sortedRows = sortRows(rows, sortMode);

    return {
      header: COLUMN_HEADERS.OPINION,
      rows: sortedRows
    };
  } catch (error) {
    console.error(`getSheetData Error for sheet "${sheetName}":`, error);
    throw new Error(`データの取得中にエラーが発生しました: ${error.message}`);
  }
}

/**
 * Add or toggle reaction with optimized batch operations
 */
function addReaction(rowIndex, reactionKey) {
  // Input validation
  if (!rowIndex || !reactionKey || !COLUMN_HEADERS[reactionKey]) {
    return { status: 'error', message: '無効なパラメータです。' };
  }
  
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(10000)) {
      return { status: 'error', message: '他のユーザーが操作中です。しばらく待ってから再試行してください。' };
    }
    
    const userEmail = safeGetUserEmail();
    if (!userEmail) {
      return { status: 'error', message: 'ログインしていないため、操作できません。' };
    }
    
    const settings = getAppSettings();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(settings.activeSheetName);
    if (!sheet) {
      throw new Error(`シート '${settings.activeSheetName}' が見つかりません。`);
    }

    // Batch header lookup with caching
    const headerIndices = getAndCacheHeaderIndices(settings.activeSheetName, 
      sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    );
    
    const reactionHeaders = [COLUMN_HEADERS.UNDERSTAND, COLUMN_HEADERS.LIKE, COLUMN_HEADERS.CURIOUS];
    const colIndices = reactionHeaders.reduce((acc, header) => {
      acc[header] = headerIndices[header] + 1;
      return acc;
    }, {});

    // Batch read all reaction columns at once
    const reactionCols = Object.values(colIndices);
    const reactionRange = sheet.getRange(rowIndex, Math.min(...reactionCols), 1, Math.max(...reactionCols) - Math.min(...reactionCols) + 1);
    const reactionValues = reactionRange.getValues()[0];
    
    // Process current reaction lists
    const lists = {};
    Object.entries(colIndices).forEach(([header, colIndex]) => {
      const valueIndex = colIndex - Math.min(...reactionCols);
      const str = reactionValues[valueIndex] ? reactionValues[valueIndex].toString() : '';
      lists[header] = {
        colIndex,
        arr: str ? str.split(',').filter(Boolean) : []
      };
    });

    const targetHeader = COLUMN_HEADERS[reactionKey];
    const wasReacted = lists[targetHeader].arr.includes(userEmail);

    // Remove user from all reactions
    Object.keys(lists).forEach(key => {
      const idx = lists[key].arr.indexOf(userEmail);
      if (idx > -1) lists[key].arr.splice(idx, 1);
    });

    // If user wasn't toggling off the same reaction, add them to selected
    if (!wasReacted) {
      lists[targetHeader].arr.push(userEmail);
    }

    // Batch write all values for better performance
    const updateData = [];
    Object.entries(lists).forEach(([header, data]) => {
      updateData.push([data.arr.join(',')]);
    });
    
    if (updateData.length > 0) {
      const writeRange = sheet.getRange(rowIndex, Math.min(...reactionCols), 1, updateData.length);
      writeRange.setValues([updateData.map(item => item[0])]);
    }

    // Prepare response
    const reactions = {};
    Object.entries(lists).forEach(([header, data]) => {
      // Convert header back to reaction key
      const reactionKeyForHeader = Object.keys(COLUMN_HEADERS).find(key => COLUMN_HEADERS[key] === header);
      if (reactionKeyForHeader) {
        reactions[reactionKeyForHeader] = {
          count: data.arr.length,
          reacted: data.arr.includes(userEmail)
        };
      }
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

/**
 * Get roster mapping with improved caching strategy
 */
function getRosterMap() {
  const cache = CacheService.getScriptCache();
  const cachedMap = cache.get(CACHE_KEYS.ROSTER);
  if (cachedMap) { 
    try {
      return JSON.parse(cachedMap);
    } catch (error) {
      console.warn('Failed to parse cached roster:', error);
      cache.remove(CACHE_KEYS.ROSTER);
    }
  }
  
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ROSTER_CONFIG.SHEET_NAME);
    if (!sheet) { 
      console.warn(`名簿シート「${ROSTER_CONFIG.SHEET_NAME}」が見つかりません。`); 
      return {}; 
    }
    
    const rosterValues = sheet.getDataRange().getValues();
    if (rosterValues.length < 2) return {}; // No data rows
    
    const rosterHeaders = rosterValues[0];
    const headerMap = {
      lastNameIndex: rosterHeaders.indexOf(ROSTER_CONFIG.HEADER_LAST_NAME),
      firstNameIndex: rosterHeaders.indexOf(ROSTER_CONFIG.HEADER_FIRST_NAME),
      nicknameIndex: rosterHeaders.indexOf(ROSTER_CONFIG.HEADER_NICKNAME),
      emailIndex: rosterHeaders.indexOf(ROSTER_CONFIG.HEADER_EMAIL)
    };
    
    if (headerMap.lastNameIndex === -1 || headerMap.firstNameIndex === -1 || headerMap.emailIndex === -1) { 
      throw new Error(`名簿シート「${ROSTER_CONFIG.SHEET_NAME}」に必要な列が見つかりません。`); 
    }
    
    const nameMap = {};
    // Process in batches for large datasets
    const dataRows = rosterValues.slice(1);
    dataRows.forEach(row => {
      const email = row[headerMap.emailIndex];
      const lastName = row[headerMap.lastNameIndex];
      const firstName = row[headerMap.firstNameIndex];
      const nickname = (headerMap.nicknameIndex !== -1 && row[headerMap.nicknameIndex]) ? row[headerMap.nicknameIndex] : ''; 
      
      if (email && lastName && firstName) {
        const fullName = `${lastName} ${firstName}`;
        nameMap[email] = nickname ? `${fullName} (${nickname})` : fullName;
      }
    });
    
    // Cache for 6 hours with error handling
    try {
      cache.put(CACHE_KEYS.ROSTER, JSON.stringify(nameMap), 21600);
    } catch (error) {
      console.warn('Failed to cache roster map:', error);
    }
    
    return nameMap;
  } catch (error) {
    console.error('Error getting roster map:', error);
    return {};
  }
}

// =================================================================
// HELPER FUNCTIONS FOR DATA PROCESSING
// =================================================================

/**
 * Get display mode with fallback
 */
function getDisplayMode() {
  return PropertiesService.getScriptProperties()
    .getProperty(APP_PROPERTIES.DISPLAY_MODE) || 'anonymous';
}

/**
 * Filter rows by class with null safety
 */
function filterRowsByClass(dataRows, classFilter, headerIndices) {
  if (!classFilter || classFilter === 'すべて') return dataRows;
  
  return dataRows.filter(row => {
    const className = row[headerIndices[COLUMN_HEADERS.CLASS]];
    return className === classFilter;
  });
}

/**
 * Parse reaction string to array
 */
function parseReactionString(reactionStr) {
  return reactionStr ? reactionStr.toString().split(',').filter(Boolean) : [];
}

/**
 * Calculate score based on reactions
 */
function calculateScore(reactions) {
  const totalReactions = reactions.UNDERSTAND.count + reactions.LIKE.count + reactions.CURIOUS.count;
  const baseScore = 0; // Can be extended with additional logic
  const likeMultiplier = 1 + (totalReactions * SCORING_CONFIG.LIKE_MULTIPLIER_FACTOR);
  return baseScore * likeMultiplier;
}

/**
 * Process single row data
 */
function processRowData(row, index, headerIndices, userEmail, emailToNameMap) {
  const email = row[headerIndices[COLUMN_HEADERS.EMAIL]];
  const opinion = row[headerIndices[COLUMN_HEADERS.OPINION]];
  
  if (!email || !opinion) return null;
  
  const understandArr = parseReactionString(row[headerIndices[COLUMN_HEADERS.UNDERSTAND]]);
  const likeArr = parseReactionString(row[headerIndices[COLUMN_HEADERS.LIKE]]);
  const curiousArr = parseReactionString(row[headerIndices[COLUMN_HEADERS.CURIOUS]]);
  
  const reactions = {
    UNDERSTAND: { 
      count: understandArr.length, 
      reacted: userEmail ? understandArr.includes(userEmail) : false 
    },
    LIKE: { 
      count: likeArr.length, 
      reacted: userEmail ? likeArr.includes(userEmail) : false 
    },
    CURIOUS: { 
      count: curiousArr.length, 
      reacted: userEmail ? curiousArr.includes(userEmail) : false 
    }
  };
  
  const displayMode = getDisplayMode();
  const actualName = displayMode === 'named' && emailToNameMap[email] 
    ? emailToNameMap[email] 
    : email.split('@')[0];
  const name = displayMode === 'named' ? actualName : '匿名';
  
  const reason = row[headerIndices[COLUMN_HEADERS.REASON]] || '';
  const highlight = String(row[headerIndices[COLUMN_HEADERS.HIGHLIGHT]]).toLowerCase() === 'true';
  
  const totalReactions = reactions.UNDERSTAND.count + reactions.LIKE.count + reactions.CURIOUS.count;
  const baseScore = reason.length;
  const likeMultiplier = 1 + (totalReactions * SCORING_CONFIG.LIKE_MULTIPLIER_FACTOR);
  const totalScore = baseScore * likeMultiplier;
  
  return {
    rowIndex: index + 2,
    name: name,
    class: row[headerIndices[COLUMN_HEADERS.CLASS]] || '未分類',
    opinion: opinion,
    reason: reason,
    reactions: reactions,
    highlight: highlight,
    score: totalScore
  };
}

/**
 * Sort rows based on mode
 */
function sortRows(rows, sortMode) {
  switch (sortMode) {
    case 'newest':
      return rows.sort((a, b) => b.rowIndex - a.rowIndex);
    case 'random':
      return rows.sort(() => Math.random() - 0.5);
    default:
      return rows.sort((a, b) => b.score - a.score);
  }
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
    saveWebAppUrl,
    getSheetUpdates,
    saveDisplayMode,
    saveDeployId,
    getActiveUserEmail: this.getActiveUserEmail,
  };
}
