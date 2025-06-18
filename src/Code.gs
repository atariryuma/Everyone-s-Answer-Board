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
// Time and performance constants
const TIME_CONSTANTS = {
  HEADERS_CACHE_TTL: 1800, // seconds (30 minutes)
  ROSTER_CACHE_TTL: 21600, // seconds (6 hours)
  PUBLISH_EXPIRY_MS: 6 * 60 * 60 * 1000, // 6 hours in milliseconds
  LOCK_WAIT_MS: 10000, // 10 seconds
  POLLING_INTERVAL_MS: 15000 // 15 seconds
};

// UI constants
const UI_CONSTANTS = {
  SIDEBAR_WIDTH: 400,
  MAX_DISPLAY_LINES: 5,
  ANIMATION_DURATION_MS: 300
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
  DEPLOY_ID: 'DEPLOY_ID',
  ADMIN_EMAILS: 'ADMIN_EMAILS', // 管理者メールアドレスのリスト（カンマ区切り）
  WEB_APP_URL: 'WEB_APP_URL'
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
 * Check if current user is an admin
 * @returns {boolean} True if user is admin
 */
function isUserAdmin() {
  const userEmail = safeGetUserEmail();
  if (!userEmail) return false;
  
  const adminEmails = PropertiesService.getScriptProperties().getProperty(APP_PROPERTIES.ADMIN_EMAILS);
  if (!adminEmails) return false;
  
  const adminList = adminEmails.split(',').map(email => email.trim().toLowerCase());
  return adminList.includes(userEmail.toLowerCase());
}

/**
 * Get admin emails list
 * @returns {Array<string>} List of admin emails
 */
function getAdminEmails() {
  const adminEmails = PropertiesService.getScriptProperties().getProperty(APP_PROPERTIES.ADMIN_EMAILS);
  return adminEmails ? adminEmails.split(',').map(email => email.trim()) : [];
}

/**
 * Set admin emails list
 * @param {Array<string>} emails - List of admin emails
 */
function setAdminEmails(emails) {
  const emailString = Array.isArray(emails) ? emails.join(',') : '';
  saveSettings({ [APP_PROPERTIES.ADMIN_EMAILS]: emailString });
}

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
 * Get admin panel data - simplified for new structure
 * @returns {object} - Admin panel data
 */
function getAdminSettings() {
  const properties = PropertiesService.getScriptProperties();
  const allSheets = getSheets();
  let currentUser = '';
  try {
   currentUser = getActiveUserEmail();
  } catch (e) {}
  
  return {
    allSheets: allSheets,
    currentUserEmail: currentUser,
    deployId: properties.getProperty(APP_PROPERTIES.DEPLOY_ID) || '',
    adminEmails: getAdminEmails(),
    isUserAdmin: isUserAdmin()
  };
}

/**
 * Generate admin URL with authentication
 * @returns {string} Admin URL
 */
function generateAdminUrl() {
  const deployId = PropertiesService.getScriptProperties().getProperty(APP_PROPERTIES.DEPLOY_ID);
  if (!deployId) return '';
  
  const baseUrl = generateWebAppUrl(deployId);
  return `${baseUrl}?admin=true`;
}

function saveDeployId(id) {
  const value = id ? String(id).trim() : '';
  saveSettings({ [APP_PROPERTIES.DEPLOY_ID]: value });
  return 'デプロイIDを更新しました。';
}

/**
 * Save the published Web App URL to script properties
 * @param {string} url - Web app URL
 * @returns {string} Status message
 */
function saveWebAppUrl(url) {
  const value = url ? String(url).trim() : '';
  saveSettings({ [APP_PROPERTIES.WEB_APP_URL]: value });
  return 'Web App URL を保存しました。';
}



// =================================================================
// GAS Webアプリケーションのエントリーポイント
// =================================================================
/**
 * Default entry point - Student view (anonymous, no counts, no admin features)
 * @param {Object} e - URL parameters
 * @returns {HtmlOutput} Student interface
 */
function doGet(e) {
  // Check if admin mode is requested
  const isAdminMode = e.parameter.admin === 'true' || isUserAdmin();
  
  if (isAdminMode) {
    return doGetAdmin(e);
  }
  
  // Default student view
  const template = HtmlService.createTemplateFromFile('Page');
  template.showCounts = false; // 生徒用：リアクション数非表示
  template.showAdminFeatures = false; // 生徒用：管理機能非表示
  template.showHighlightToggle = false; // 生徒用：ハイライト切り替えなし
  template.showScoreSort = false; // 生徒用：スコア順表示なし
  template.showPublishControls = false; // 生徒用：公開終了ボタンなし
  template.displayMode = 'anonymous'; // 生徒用：匿名表示
  template.isAdminUser = isUserAdmin(); // 管理者判定

  return template.evaluate()
      .setTitle('StudyQuest - みんなのかいとうボード')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Admin entry point - Full featured view with access control
 * @param {Object} e - URL parameters  
 * @returns {HtmlOutput} Admin interface or access denied
 */
function doGetAdmin(e) {
  // Check admin access
  if (!isUserAdmin()) {
    return HtmlService.createHtmlOutput(`
      <div style="text-align: center; padding: 50px; font-family: Arial, sans-serif;">
        <h2>アクセス拒否</h2>
        <p>管理者権限が必要です。</p>
        <p>アクセス権限がない場合は、管理者にお問い合わせください。</p>
      </div>
    `).setTitle('アクセス拒否');
  }
  
  // Admin view with all features
  const template = HtmlService.createTemplateFromFile('Page');
  template.showCounts = true; // 管理者用：リアクション数表示
  template.showAdminFeatures = true; // 管理者用：管理機能表示
  template.showHighlightToggle = true; // 管理者用：ハイライト切り替えあり
  template.showScoreSort = true; // 管理者用：スコア順表示あり
  template.showPublishControls = true; // 管理者用：公開終了ボタンあり
  template.displayMode = 'named'; // 管理者用：実名表示
  template.isAdminUser = true; // 管理者判定
  
  let userEmail;
  try {
    userEmail = getActiveUserEmail();
  } catch (e) {
    userEmail = '管理者';
  }
  template.userEmail = userEmail;
  
  return template.evaluate()
      .setTitle('StudyQuest - 管理者ビュー')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// =================================================================
// クライアントサイドから呼び出される関数
// =================================================================

/**
 * Generate Web App URL from deploy ID
 * @param {string} deployId - Google Apps Script deploy ID
 * @returns {string} Complete Web App URL
 */
function generateWebAppUrl(deployId) {
  if (!deployId) {
    console.warn('Deploy ID is empty');
    return '';
  }
  
  // Google Apps Script Web App URL format
  const baseUrl = 'https://script.google.com/macros/s';
  return `${baseUrl}/${deployId}/exec`;
}

/**
 * Get sheet data for display - simplified version
 * @param {string} sheetName - Sheet name to get data from  
 * @param {string} classFilter - Class filter
 * @param {string} sortMode - Sort mode
 * @param {boolean} isAdminMode - Whether admin features are enabled
 * @returns {Object} Sheet data
 */
function getSheetDataForDisplay(sheetName, classFilter = '', sortMode = 'newest', isAdminMode = false) {
  // Use first available sheet if no sheet name provided
  if (!sheetName) {
    const sheets = getSheets();
    sheetName = sheets.length > 0 ? sheets[0] : null;
  }
  
  if (!sheetName) {
    throw new Error('利用可能なシートがありません。');
  }
  
  return getSheetData(sheetName, classFilter, sortMode, isAdminMode);
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
 * @param {boolean} isAdminMode Whether admin features are enabled
 * @returns {Object} Processed sheet data
 */
function getSheetData(sheetName, classFilter = '', sortMode = 'newest', isAdminMode = false) {
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
    // Admin mode: named display, Student mode: anonymous
    const displayMode = isAdminMode ? 'named' : 'anonymous';
    const emailToNameMap = displayMode === 'named' ? getRosterMap() : {};

    const filteredRows = filterRowsByClass(dataRows, classFilter, headerIndices);
    const rows = filteredRows
      .map((row, i) => processRowData(row, i, headerIndices, userEmail, emailToNameMap, displayMode))
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
 * @param {number} rowIndex - Target row index (1-based)
 * @param {string} reactionKey - Reaction type key ('UNDERSTAND', 'LIKE', 'CURIOUS')
 * @returns {Object} Response object with status and updated reaction data
 * @returns {string} returns.status - 'ok' or 'error'
 * @returns {string} [returns.message] - Error message if status is 'error'
 * @returns {Object<string, {count: number, reacted: boolean}>} [returns.reactions] - Updated reaction data
 */
function addReaction(rowIndex, reactionKey) {
  // Input validation
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
    
    // Use first available sheet
    const sheets = getSheets();
    if (sheets.length === 0) {
      throw new Error('利用可能なシートがありません。');
    }
    
    const sheetName = sheets[0];
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`シート '${sheetName}' が見つかりません。`);
    }

    // Batch header lookup with caching
    const headerIndices = getAndCacheHeaderIndices(sheetName, 
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

/**
 * Toggle highlight status for a specific row
 * @param {number} rowIndex - Target row index (1-based)
 * @returns {Object} Response object with status and highlight state
 * @returns {string} returns.status - 'ok' or 'error'
 * @returns {string} [returns.message] - Error message if status is 'error'
 * @returns {boolean} [returns.highlight] - New highlight state
 */
function toggleHighlight(rowIndex) {
  const lock = LockService.getScriptLock();
  lock.waitLock(TIME_CONSTANTS.LOCK_WAIT_MS);
  try {
    // Use first available sheet
    const sheets = getSheets();
    if (sheets.length === 0) {
      throw new Error('利用可能なシートがありません。');
    }
    
    const sheetName = sheets[0];
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) throw new Error(`シート '${sheetName}' が見つかりません。`);

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
      cache.put(CACHE_KEYS.ROSTER, JSON.stringify(nameMap), TIME_CONSTANTS.ROSTER_CACHE_TTL);
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
 * Parse reaction string to array with enhanced error handling
 * @param {any} reactionStr - Reaction string (can be null/undefined)
 * @returns {Array<string>} Array of reaction emails
 */
function parseReactionString(reactionStr) {
  if (!reactionStr) return [];
  
  try {
    return String(reactionStr).split(',').filter(Boolean).map(s => s.trim());
  } catch (error) {
    console.warn('Failed to parse reaction string:', reactionStr, error);
    return [];
  }
}

/**
 * Calculate score based on reactions with null safety
 * @param {Object} reactions - Reaction data object
 * @returns {number} Calculated score
 */
function calculateScore(reactions) {
  if (!reactions || typeof reactions !== 'object') return 0;
  
  const understandCount = reactions.UNDERSTAND?.count ?? 0;
  const likeCount = reactions.LIKE?.count ?? 0;
  const curiousCount = reactions.CURIOUS?.count ?? 0;
  
  const totalReactions = understandCount + likeCount + curiousCount;
  const baseScore = 0; // Can be extended with additional logic
  const likeMultiplier = 1 + (totalReactions * SCORING_CONFIG.LIKE_MULTIPLIER_FACTOR);
  
  return Math.max(0, baseScore * likeMultiplier);
}

/**
 * Process single row data with enhanced validation
 * @param {Array} row - Row data array
 * @param {number} index - Row index
 * @param {Object} headerIndices - Header to index mapping
 * @param {string} userEmail - Current user email
 * @param {Object} emailToNameMap - Email to name mapping
 * @param {string} displayMode - Display mode ('named' or 'anonymous')
 * @returns {Object|null} Processed row data or null if invalid
 */
function processRowData(row, index, headerIndices, userEmail, emailToNameMap, displayMode = 'anonymous') {
  // Enhanced input validation
  if (!Array.isArray(row) || !headerIndices || typeof headerIndices !== 'object') {
    console.warn('Invalid input to processRowData:', { row: !!row, headerIndices: !!headerIndices });
    return null;
  }
  
  const emailIndex = headerIndices[COLUMN_HEADERS.EMAIL];
  const opinionIndex = headerIndices[COLUMN_HEADERS.OPINION];
  
  if (emailIndex === undefined || opinionIndex === undefined) {
    console.warn('Missing required header indices');
    return null;
  }
  
  const email = row[emailIndex];
  const opinion = row[opinionIndex];
  
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
  
  let name;
  if (displayMode === 'named') {
    name = emailToNameMap[email] || email.split('@')[0];
  } else {
    name = '匿名';
  }
  
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
  cache.put(cacheKey, JSON.stringify(indices), TIME_CONSTANTS.HEADERS_CACHE_TTL);
  return indices;
}

/**
 * Find header indices with optimized normalization
 * @param {Array<string>} sheetHeaders - Sheet header row
 * @param {Array<string>} requiredHeaders - Required header names
 * @returns {Object<string, number>} Header name to index mapping
 */
function findHeaderIndices(sheetHeaders, requiredHeaders) {
  const indices = {};
  const missingHeaders = [];
  
  // Normalize function with memoization-like behavior
  const normalize = h => (typeof h === 'string' ? h.replace(/\s+/g, '') : String(h));
  
  // Pre-normalize all headers once for O(n) instead of O(n*m)
  const normalizedSheetHeaders = sheetHeaders.map(normalize);
  const normalizedRequiredHeaders = requiredHeaders.map(normalize);
  
  // Create a lookup map for faster searches
  const headerMap = new Map();
  normalizedSheetHeaders.forEach((normalizedHeader, index) => {
    headerMap.set(normalizedHeader, index);
  });
  
  // Find indices using Map lookup (O(1) average case)
  requiredHeaders.forEach((originalHeader, reqIndex) => {
    const normalizedRequired = normalizedRequiredHeaders[reqIndex];
    const index = headerMap.get(normalizedRequired);
    
    if (index !== undefined) {
      indices[originalHeader] = index;
    } else {
      missingHeaders.push(originalHeader);
    }
  });
  
  if (missingHeaders.length > 0) {
    throw new Error(`必須ヘッダーが見つかりません: [${missingHeaders.join(', ')}]`);
  }
  
  return indices;
}

// =================================================================
// SIDEBAR AND UI FUNCTIONS
// =================================================================

/**
 * Show the sheet selector sidebar in Google Sheets
 */
function showSheetSelector() {
  try {
    const htmlOutput = HtmlService.createTemplateFromFile('SheetSelector')
      .evaluate()
      .setTitle('StudyQuest 管理パネル')
      .setWidth(UI_CONSTANTS.SIDEBAR_WIDTH);
    
    SpreadsheetApp.getUi().showSidebar(htmlOutput);
  } catch (error) {
    console.error('Failed to show sheet selector:', error);
    SpreadsheetApp.getUi().alert('サイドバーの表示に失敗しました: ' + error.message);
  }
}

/**
 * Create custom menu in Google Sheets
 */
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('StudyQuest 管理')
      .addItem('📊 管理パネルを開く', 'showSheetSelector')
      .addSeparator()
      .addItem('🚀 アプリを開く', 'openPublishedApp')
      .addToUi();
  } catch (error) {
    console.error('Failed to create menu:', error);
  }
}

/**
 * Open the published app in a new tab using DEPLOY_ID
 */
function openPublishedApp() {
  try {
    const deployId = PropertiesService.getScriptProperties().getProperty(APP_PROPERTIES.DEPLOY_ID);
    
    if (!deployId) {
      SpreadsheetApp.getUi().alert('デプロイIDが設定されていません。管理パネルからデプロイIDを設定してください。');
      return;
    }
    
    const url = generateWebAppUrl(deployId);
    
    if (url) {
      const htmlOutput = HtmlService.createHtmlOutput(`
        <script>
          window.open('${url}', '_blank');
          google.script.host.close();
        </script>
      `).setWidth(1).setHeight(1);
      
      SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'アプリを開いています...');
    } else {
      SpreadsheetApp.getUi().alert('URLの生成に失敗しました。デプロイIDを確認してください。');
    }
  } catch (error) {
    console.error('Failed to open app:', error);
    SpreadsheetApp.getUi().alert('アプリの起動に失敗しました: ' + error.message);
  }
}

// Export for Jest testing
if (typeof module !== 'undefined') {
  module.exports = {
    COLUMN_HEADERS,
    findHeaderIndices,
    getSheetData,
    getSheetDataForDisplay,
    getAdminSettings,
    doGet,
    doGetAdmin,
    addReaction,
    toggleHighlight,
    generateWebAppUrl,
    generateAdminUrl,
    saveDeployId,
    saveWebAppUrl,
    showSheetSelector,
    onOpen,
    openPublishedApp,
    isUserAdmin,
    getAdminEmails,
    setAdminEmails,
    getActiveUserEmail: this.getActiveUserEmail,
  };
}
