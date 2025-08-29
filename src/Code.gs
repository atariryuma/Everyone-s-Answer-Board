/**
 * @fileoverview StudyQuest - みんなの回答ボード
 * スプレッドシートのカスタムメニューを使用しない独立版。
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

const TIME_CONSTANTS = {
  LOCK_WAIT_MS: 10000
};

var getConfig;
var handleError;

// Default error handler
if (!handleError) {
  handleError = function(context, error, returnErrorObj = false) {
    console.error(`Error in ${context}:`, error);
    if (returnErrorObj) {
      return { status: 'error', message: error.message || 'エラーが発生しました。' };
    }
    throw error;
  };
}

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
    const email = Session.getActiveUser().getEmail();
    if (!email || typeof email !== 'string' || !isValidEmail(email)) {
      throw new Error('Invalid user email');
    }
    return email;
  } catch (e) {
    console.error('Failed to get user email:', e);
    // In test environment, return a default test email
    if (typeof global !== 'undefined' && global.mockUserEmail) {
      return global.mockUserEmail;
    }
    throw new Error('認証が必要です。Googleアカウントでログインしてください。');
  }
}

function isValidEmail(email) {
  if (!email || typeof email !== 'string') {
    return false;
  }
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

function getEmailDomain(email) {
  return (email || '').toString().split('@').pop().toLowerCase();
}

function isSameDomain(emailA, emailB) {
  if (!emailA || !emailB || typeof emailA !== 'string' || typeof emailB !== 'string') {
    return false;
  }
  if (!isValidEmail(emailA) || !isValidEmail(emailB)) {
    return false;
  }
  const domainA = getEmailDomain(emailA);
  const domainB = getEmailDomain(emailB);
  return domainA && domainB && domainA === domainB;
}

function getAdminEmails(spreadsheetId) {
  const props = PropertiesService.getScriptProperties();
  let adminEmails = [];
  
  if (spreadsheetId) {
    const str = props.getProperty(`ADMIN_EMAILS_${spreadsheetId}`);
    if (str) adminEmails = str.split(',').map(e => e.trim()).filter(Boolean);
  }
  
  if (adminEmails.length === 0) {
    const str = props.getProperty('ADMIN_EMAILS') || '';
    adminEmails = str.split(',').map(e => e.trim()).filter(Boolean);
  }
  
  // 管理者が設定されていない場合、スプレッドシートの所有者を管理者とする
  if (adminEmails.length === 0) {
    try {
      const ss = getCurrentSpreadsheet();
      const owner = ss.getOwner();
      if (owner) {
        const ownerEmail = owner.getEmail();
        if (ownerEmail) {
          console.log('管理者未設定のため、所有者を管理者に設定:', ownerEmail);
          adminEmails = [ownerEmail];
        }
      }
    } catch (e) {
      console.error('所有者情報の取得に失敗:', e);
    }
  }
  
  return adminEmails;
}

function isUserAdmin(email) {
  const userProps = PropertiesService.getUserProperties();
  const userId = userProps.getProperty('CURRENT_USER_ID');
  const userEmail = email || safeGetUserEmail();

  console.log('isUserAdmin check:', { 
    email: email, 
    userEmail: userEmail, 
    userId: userId 
  });

  if (userId) {
    const userInfo = getUserInfo(userId);
    const spreadsheetId = userInfo && userInfo.spreadsheetId;
    const admins = getAdminEmails(spreadsheetId);
    console.log('Multi-tenant admin check:', { 
      spreadsheetId: spreadsheetId, 
      admins: admins, 
      userEmail: userEmail, 
      isAdmin: admins.includes(userEmail) 
    });
    return admins.includes(userEmail);
  }

  const admins = getAdminEmails();
  console.log('Single-tenant admin check:', { 
    admins: admins, 
    userEmail: userEmail, 
    isAdmin: admins.includes(userEmail) 
  });
  return admins.includes(userEmail);
}

function checkAdmin() {
  return isUserAdmin();
}

function publishApp(sheetName) {
  if (!checkAdmin()) {
    throw new Error('権限がありません。');
  }
  const props = PropertiesService.getScriptProperties();
  props.setProperty('IS_PUBLISHED', 'true');
  props.setProperty('PUBLISHED_SHEET_NAME', sheetName);
  return `「${sheetName}」を公開しました。`;
}

function unpublishApp() {
  if (!checkAdmin()) {
    throw new Error('権限がありません。');
  }
  const props = PropertiesService.getScriptProperties();
  props.setProperty('IS_PUBLISHED', 'false');
  if (props.deleteProperty) props.deleteProperty('PUBLISHED_SHEET_NAME');
  return 'アプリを非公開にしました。';
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


/**
 * 管理パネルの初期化に必要なデータを取得します。
 * @returns {object} - 現在のアクティブシート情報とシートのリスト
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
      adminEmails = getAdminEmails(userInfo.spreadsheetId);
      appSettings = getAppSettingsForUser();
    }
  } else {
    // 従来のモード
    adminEmails = getAdminEmails();
    appSettings = getAppSettingsForUser();
  }
  
  const allSheets = getSheets();
  const currentUserEmail = safeGetUserEmail();
  
  // 新しいシステム：常に利用可能な回答ボード
  // activeSheetNameがあれば利用可能、なければシート選択が必要
  const isPublished = props.getProperty('IS_PUBLISHED') === 'true';
  const activeSheetName = appSettings.activeSheetName || props.getProperty('ACTIVE_SHEET_NAME') || '';
  const isAvailable = !!activeSheetName;
  
  return {
    allSheets: allSheets,
    currentUserEmail: currentUserEmail,
    deployId: props.getProperty('DEPLOY_ID'),
    webAppUrl: getWebAppUrl(),
    adminEmails: adminEmails,
    isUserAdmin: adminEmails.includes(currentUserEmail),
    activeSheetName: activeSheetName,
    showNames: appSettings.showNames,
    showCounts: appSettings.showCounts,
    isPublished: isPublished
  };
}

/**
 * アクティブシートを指定されたシートに変更します。
 * @param {string} sheetName - アクティブにするシート名。
 */
function switchActiveSheet(sheetName) {
  const props = PropertiesService.getUserProperties();
  const userId = props.getProperty('CURRENT_USER_ID');
  
  if (!userId) {
    throw new Error('ユーザー情報が見つかりません。');
  }
  
  const userInfo = getUserInfo(userId);
  
  if (!userInfo || userInfo.adminEmail !== Session.getActiveUser().getEmail()) {
    throw new Error('権限がありません。');
  }
  
  if (!sheetName) {
    throw new Error('シート名が指定されていません。');
  }
  
  // スプレッドシートの存在確認
  const spreadsheet = SpreadsheetApp.openById(userInfo.spreadsheetId);
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`シート「${sheetName}」が見つかりません。`);
  }

  prepareSheetForBoard(sheetName);
  clearHeaderCache(sheetName);

  try {
    updateUserConfig(userId, {
      activeSheetName: sheetName
    });
  } catch (error) {
    throw new Error(`設定の更新に失敗しました: ${error.message}`);
  }

  auditLog('ACTIVE_SHEET_CHANGED', userId, { sheetName, userEmail: userInfo.adminEmail });
  return `アクティブシートを「${sheetName}」に変更しました。`;
}

/**
 * アクティブシートの選択をクリアします。
 */
function clearActiveSheet() {
  const props = PropertiesService.getUserProperties();
  const userId = props.getProperty('CURRENT_USER_ID');
  
  if (!userId) {
    throw new Error('ユーザー情報が見つかりません。');
  }
  
  const userInfo = getUserInfo(userId);
  if (!userInfo || userInfo.adminEmail !== Session.getActiveUser().getEmail()) {
    throw new Error('権限がありません。');
  }
  
  // アクティブシートをクリア
  try {
    updateUserConfig(userId, {
      activeSheetName: ''
    });
  } catch (error) {
    console.error('Failed to update user config:', error);
    throw new Error(`設定の更新に失敗しました: ${error.message}`);
  }
  
  auditLog('ACTIVE_SHEET_CLEARED', userId, { userEmail: userInfo.adminEmail });
  
  return 'アクティブシートをクリアしました。';
}


function setShowDetails(flag) {
  const props = PropertiesService.getUserProperties();
  const userId = props.getProperty('CURRENT_USER_ID');
  
  if (!userId) {
    throw new Error('ユーザー情報が見つかりません。');
  }
  
  const userInfo = getUserInfo(userId);
  if (!userInfo || userInfo.adminEmail !== Session.getActiveUser().getEmail()) {
    throw new Error('権限がありません。');
  }
  
  // ユーザー設定を更新
  try {
    updateUserConfig(userId, {
      showNames: Boolean(flag),
      showCounts: Boolean(flag),
      showDetails: Boolean(flag)
    });
  } catch (error) {
    console.error('Failed to update user config:', error);
    throw new Error(`設定の更新に失敗しました: ${error.message}`);
  }

  auditLog('SHOW_DETAILS_UPDATED', userId, { showNames: Boolean(flag), showCounts: Boolean(flag), userEmail: userInfo.adminEmail });

  return `詳細表示を${flag ? '有効' : '無効'}にしました。`;
}

function setDisplayOptions(options) {
  const props = PropertiesService.getUserProperties();
  const userId = props.getProperty('CURRENT_USER_ID');

  if (!userId) {
    throw new Error('ユーザー情報が見つかりません。');
  }

  const userInfo = getUserInfo(userId);
  if (!userInfo || userInfo.adminEmail !== Session.getActiveUser().getEmail()) {
    throw new Error('権限がありません。');
  }

  try {
    updateUserConfig(userId, {
      showNames: Boolean(options.showNames),
      showCounts: Boolean(options.showCounts)
    });
  } catch (error) {
    console.error('Failed to update user config:', error);
    throw new Error(`設定の更新に失敗しました: ${error.message}`);
  }

  auditLog('DISPLAY_OPTIONS_UPDATED', userId, { options, userEmail: userInfo.adminEmail });

  return '表示設定を更新しました。';
}



// =================================================================
// GAS Webアプリケーションのエントリーポイント
// =================================================================
function validateUserId(userId) {
  if (!userId || typeof userId !== 'string') {
    throw new Error('ユーザーIDが指定されていません。');
  }
  if (!/^[a-zA-Z0-9-_]{10,}$/.test(userId)) {
    throw new Error('無効なユーザーIDの形式です。');
  }
  return userId;
}



// =================================================================
// クライアントサイドから呼び出される関数
// =================================================================

/**
 * 指定されたシートまたはアクティブシートのデータを取得します。
 * @param {string} requestedSheetName - 取得するシート名（省略時はアクティブシート）
 * @param {string} classFilter - クラスフィルター
 * @param {string} sortBy - ソート順
 */
function getPublishedSheetData(requestedSheetName, classFilter, sortBy) {
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
  
  // 下位互換性を保ちながら、アクティブシート名を取得
  const activeSheetName = config.activeSheetName || config.sheetName || '';
  
  // リクエストされたシート名、またはアクティブシートを使用
  const targetSheetName = requestedSheetName || activeSheetName;
  
  if (!targetSheetName) {
    throw new Error('表示するシートが設定されていません。');
  }
  
  // ユーザーのスプレッドシートを開く
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheet = spreadsheet.getSheetByName(targetSheetName);
  
  if (!sheet) {
    throw new Error(`シート「${targetSheetName}」が見つかりません。`);
  }
  
  // 既存のgetSheetData関数のロジックを使用（スプレッドシートを指定）
  const order = sortBy || 'newest';
  const data = getSheetDataForSpreadsheet(spreadsheet, targetSheetName, classFilter, order);
  
  const cfgShowNames = typeof config.showNames !== 'undefined' ? config.showNames : (config.showDetails || false);
  const cfgShowCounts = typeof config.showCounts !== 'undefined' ? config.showCounts : (config.showDetails || false);

  return {
    sheetName: targetSheetName,
    header: data.header,
    rows: data.rows,
    showNames: cfgShowNames,
    showCounts: cfgShowCounts,
    isActiveSheet: targetSheetName === activeSheetName
  };
}

/**
 * 利用可能なシート一覧とアクティブシート情報を取得します。
 */
function getAvailableSheets() {
  const props = PropertiesService.getUserProperties();
  const userId = props.getProperty('CURRENT_USER_ID');
  
  if (!userId) {
    throw new Error('ユーザー情報が見つかりません。');
  }
  
  const userInfo = getUserInfo(userId);
  const config = userInfo.configJson || {};
  const activeSheetName = config.activeSheetName || config.sheetName || '';
  
  const allSheets = getSheets();
  
  return {
    sheets: allSheets.map(sheetName => ({
      name: sheetName,
      isActive: sheetName === activeSheetName
    })),
    activeSheetName: activeSheetName
  };
}

/**
 * 新しいスプレッドシートURLを追加し、アクティブシートとして設定します。
 * @param {string} spreadsheetUrl - 追加するスプレッドシートのURL
 */
function addSpreadsheetUrl(spreadsheetUrl) {
  console.log('addSpreadsheetUrl called with:', spreadsheetUrl);
  
  const props = PropertiesService.getUserProperties();
  const userId = props.getProperty('CURRENT_USER_ID');
  
  console.log('Current user ID:', userId);
  
  if (!userId) {
    throw new Error('ユーザー情報が見つかりません。登録プロセスを再実行してください。');
  }
  
  const userInfo = getUserInfo(userId);
  const currentUserEmail = Session.getActiveUser().getEmail();
  
  console.log('User info:', userInfo);
  console.log('Current user email:', currentUserEmail);
  
  if (!userInfo) {
    throw new Error('ユーザー情報が見つかりません。登録プロセスを再実行してください。');
  }
  
  if (userInfo.adminEmail !== currentUserEmail) {
    throw new Error('権限がありません。管理者アカウントでログインしてください。');
  }
  
  // URLからスプレッドシートIDを抽出
  let spreadsheetId;
  const urlPattern = /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/;
  const match = spreadsheetUrl.match(urlPattern);
  
  if (match) {
    spreadsheetId = match[1];
  } else {
    throw new Error('有効なGoogleスプレッドシートのURLではありません。');
  }
  
  // スプレッドシートにアクセスできるかテスト
  try {
    console.log('Attempting to open spreadsheet with ID:', spreadsheetId);
    const testSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheets = testSpreadsheet.getSheets();
    
    console.log('Spreadsheet name:', testSpreadsheet.getName());
    console.log('Number of sheets:', sheets.length);
    
    if (sheets.length === 0) {
      throw new Error('スプレッドシートにシートが見つかりません。');
    }
    
    // ユーザー設定を更新：新しいスプレッドシートIDを設定
    const userDb = getDatabase().getSheetByName(USER_DB_CONFIG.SHEET_NAME);
    const data = userDb.getDataRange().getValues();
    const headers = data[0];
    const userIdIndex = headers.indexOf('userId');
    const spreadsheetIdIndex = headers.indexOf('spreadsheetId');
    const spreadsheetUrlIndex = headers.indexOf('spreadsheetUrl');
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][userIdIndex] === userId) {
        // スプレッドシートIDとURLを更新
        userDb.getRange(i + 1, spreadsheetIdIndex + 1).setValue(spreadsheetId);
        userDb.getRange(i + 1, spreadsheetUrlIndex + 1).setValue(spreadsheetUrl);
        break;
      }
    }
    
    // ユーザープロパティも更新
    props.setProperty('CURRENT_SPREADSHEET_ID', spreadsheetId);
    
    // 最初のシートをアクティブシートとして設定
    const firstSheetName = sheets[0].getName();
    updateUserConfig(userId, {
      activeSheetName: firstSheetName
    });
    
    // 新しいシートの基本設定を自動作成
    try {
      // シートのヘッダーを取得して推測
      const firstSheet = testSpreadsheet.getSheetByName(firstSheetName);
      if (firstSheet && firstSheet.getLastRow() > 0) {
        const headers = firstSheet.getRange(1, 1, 1, firstSheet.getLastColumn()).getValues()[0];
        const guessedConfig = guessHeadersFromArray(headers);
        
        // 基本設定を保存（少なくとも1つのヘッダーが推測できた場合）
        if (guessedConfig.questionHeader || guessedConfig.answerHeader) {
          saveSheetConfig(firstSheetName, guessedConfig);
          console.log('Auto-created config for new sheet:', firstSheetName, guessedConfig);
        }
      }
    } catch (configError) {
      console.warn('Failed to auto-create config for new sheet:', configError.message);
      // 設定作成に失敗してもスプレッドシート追加は成功とする
    }
    
    auditLog('SPREADSHEET_ADDED', userId, { 
      spreadsheetId, 
      spreadsheetUrl, 
      firstSheetName,
      userEmail: userInfo.adminEmail 
    });
    
    console.log('Successfully added spreadsheet:', testSpreadsheet.getName());
    console.log('Active sheet set to:', firstSheetName);
    
    return {
      success: true,
      message: `スプレッドシート「${testSpreadsheet.getName()}」が追加され、シート「${firstSheetName}」がアクティブになりました。`,
      spreadsheetId: spreadsheetId,
      firstSheetName: firstSheetName
    };
    
  } catch (error) {
    console.error('Failed to access spreadsheet:', error);
    throw new Error(`スプレッドシートにアクセスできません。URLが正しいか、共有設定を確認してください。エラー: ${error.message}`);
  }
}

/**
 * AdminPanel用のステータス取得関数
 */
function getStatus() {
  try {
    const settings = getAdminSettings();
    const props = PropertiesService.getUserProperties();
    const userId = props.getProperty('CURRENT_USER_ID');
    
    if (!userId) {
      return {
        activeSheetName: '',
        allSheets: [],
        answerCount: 0,
        totalReactions: 0,
        webAppUrl: getWebAppUrl(),
        showNames: false,
        showCounts: false
      };
    }
    
    const userInfo = getUserInfo(userId);
    const config = userInfo ? userInfo.configJson || {} : {};
    const activeSheetName = config.activeSheetName || '';
    
    // Get available sheets
    let allSheets = [];
    try {
      allSheets = getSheets() || [];
    } catch (error) {
      console.warn('Failed to get sheets:', error);
    }
    
    // Get answer count if active sheet exists
    let answerCount = 0;
    let totalReactions = 0;
    
    if (activeSheetName) {
      try {
        const sheetData = getPublishedSheetData(activeSheetName);
        answerCount = sheetData.rows ? sheetData.rows.length : 0;
        totalReactions = sheetData.rows ? 
          sheetData.rows.reduce((sum, row) => {
            const reactions = row.reactions || {};
            return sum + (reactions.UNDERSTAND?.count || 0) + 
                       (reactions.LIKE?.count || 0) + 
                       (reactions.CURIOUS?.count || 0);
          }, 0) : 0;
      } catch (error) {
        console.warn('Failed to get sheet data for status:', error);
      }
    }
    
    return {
      activeSheetName: activeSheetName,
      allSheets: allSheets,
      answerCount: answerCount,
      totalReactions: totalReactions,
      webAppUrl: getWebAppUrl(),
      showNames: typeof config.showNames !== 'undefined' ? config.showNames : (config.showDetails || false),
      showCounts: typeof config.showCounts !== 'undefined' ? config.showCounts : (config.showDetails || false)
    };
  } catch (error) {
    console.error('getStatus error:', error);
    throw new Error('ステータスの取得に失敗しました: ' + error.message);
  }
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
      return name !== 'Config';
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

/**
 * ヘッダー配列から設定項目を推測する関数
 *
 * Keywords cover common variants (e.g. コメント vs 回答) and include
 * both full/half-width characters. Header strings are normalized by
 * trimming/collapsing whitespace and converting full-width ASCII before
 * searching.
 *
 * @param {Array} headers - ヘッダー名の配列
 * @returns {Object} 推測された設定オブジェクト
 */
function guessHeadersFromArray(headers) {
  // Normalize header strings: convert full-width ASCII, collapse whitespace and lowercase
  const normalize = (str) => String(str || '')
    .replace(/[\uFF01-\uFF5E]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0))
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();

  const normalizedHeaders = headers.map(h => ({ original: String(h), normalized: normalize(h) }));

  // Find the first header that contains any of the provided keywords
  const find = (keys) => {
    const keyNorm = keys.map(k => normalize(k));
    const found = normalizedHeaders.find(h => keyNorm.some(k => h.normalized.includes(k)));
    return found ? found.original : '';
  };
  
  console.log('Available headers:', headers);
  
  // Googleフォーム特有のヘッダー構造に対応
  const isGoogleForm = normalizedHeaders.some(h =>
    h.normalized.includes(normalize('タイムスタンプ')) ||
    h.normalized.includes(normalize('メールアドレス'))
  );
  
  let question = '';
  let answer = '';
  let reason = '';
  let name = '';
  let classHeader = '';
  
  if (isGoogleForm) {
    console.log('Detected Google Form response sheet');
    
    // Googleフォームの一般的な構造: タイムスタンプ, メールアドレス, [質問1], [質問2], ...
    const nonMetaHeaders = normalizedHeaders.filter(h => {
      const hStr = h.normalized;
      return !hStr.includes(normalize('タイムスタンプ')) &&
             !hStr.includes('timestamp') &&
             !hStr.includes(normalize('メールアドレス')) &&
             !hStr.includes('email');
    });
    
    console.log('Non-meta headers:', nonMetaHeaders.map(h => h.original));
    
    // より柔軟な推測ロジック
    for (let i = 0; i < nonMetaHeaders.length; i++) {
      const { original: header, normalized: hStr } = nonMetaHeaders[i];
      
      // 質問を含む長いテキストを質問ヘッダーとして推測
      if (!question && (hStr.includes('だろうか') || hStr.includes('ですか') || hStr.includes('でしょうか') || hStr.length > 20)) {
        question = header;
        // 同じ内容が複数列にある場合、回答用として2番目を使用
        if (i + 1 < nonMetaHeaders.length && nonMetaHeaders[i + 1].original === header) {
          answer = header;
          continue;
        }
      }
      
      // 回答・意見に関するヘッダー
      if (!answer && (hStr.includes('回答') || hStr.includes('答え') || hStr.includes('意見') || hStr.includes('考え'))) {
        answer = header;
      }
      
      // 理由に関するヘッダー
      if (!reason && (hStr.includes('理由') || hStr.includes('詳細') || hStr.includes('説明'))) {
        reason = header;
      }
      
      // 名前に関するヘッダー
      if (!name && (hStr.includes('名前') || hStr.includes('氏名') || hStr.includes('学生'))) {
        name = header;
      }
      
      // クラスに関するヘッダー
      if (!classHeader && (hStr.includes('クラス') || hStr.includes('組') || hStr.includes('学級'))) {
        classHeader = header;
      }
    }
    
    // フォールバック: まだ回答が見つからない場合
    if (!answer && nonMetaHeaders.length > 0) {
      // 最初の非メタヘッダーを回答として使用
      answer = nonMetaHeaders[0].original;
    }
    
  } else {
    // 通常のシート用の推測ロジック
    // Keyword lists include common variants and full/half-width characters
    question = find(['質問', '質問内容', '問題', '問い', 'お題', 'question', 'question text', 'q', 'Ｑ']);
    answer = find(['回答', '解答', '答え', '返答', 'answer', 'a', 'Ａ', '意見', 'コメント', 'opinion', 'response']);
    reason = find(['理由', '訳', 'わけ', '根拠', '詳細', '説明', 'comment', 'why', 'reason', 'detail']);
    name = find(['名前', '氏名', 'name', '学生', 'student', 'ペンネーム', 'ニックネーム', 'author']);
    classHeader = find(['クラス', 'class', 'classroom', '組', '学級', '班', 'グループ']);
  }
  
  const guessed = {
    questionHeader: question,
    answerHeader: answer,
    reasonHeader: reason,
    nameHeader: name,
    classHeader: classHeader
  };
  
  console.log('Guessed headers:', guessed);
  
  // 最終フォールバック: 何も推測できない場合
  if (!question && !answer && headers.length > 0) {
    console.log('No specific headers found, using positional mapping');
    
    // タイムスタンプとメールを除外して最初の列を回答として使用
    const usableHeaders = normalizedHeaders
      .filter(h => {
        const hStr = h.normalized;
        return !hStr.includes(normalize('タイムスタンプ')) &&
               !hStr.includes('timestamp') &&
               !hStr.includes(normalize('メールアドレス')) &&
               !hStr.includes('email') &&
               h.normalized !== '';
      })
      .map(h => h.original);
    
    if (usableHeaders.length > 0) {
      guessed.answerHeader = usableHeaders[0];
      if (usableHeaders.length > 1) guessed.reasonHeader = usableHeaders[1];
      if (usableHeaders.length > 2) guessed.nameHeader = usableHeaders[2];
      if (usableHeaders.length > 3) guessed.classHeader = usableHeaders[3];
    }
  }
  
  return guessed;
}

/**
 * Minimal board data builder used in tests
 */
function buildBoardData(sheetName) {
  const sheet = getCurrentSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`指定されたシート「${sheetName}」が見つかりません。`);

  const values = sheet.getDataRange().getValues();
  if (values.length < 1) return { header: '', entries: [] };

  const cfgFunc = (typeof global !== 'undefined' && global.getConfig)
    ? global.getConfig
    : (typeof getConfig === 'function' ? getConfig : null);
  const cfg = cfgFunc ? cfgFunc(sheetName) : {};

  const answerHeader = cfg.answerHeader || cfg.questionHeader || COLUMN_HEADERS.OPINION;
  const reasonHeader = cfg.reasonHeader || COLUMN_HEADERS.REASON;
  const classHeader = cfg.classHeader || COLUMN_HEADERS.CLASS;
  const nameHeader = cfg.nameHeader || '';

  const required = [answerHeader, reasonHeader, classHeader, COLUMN_HEADERS.EMAIL];
  if (nameHeader) required.push(nameHeader);

  const headerIndices = findHeaderIndices(values[0], required);

  const entries = values.slice(1).map(row => ({
    answer: row[headerIndices[answerHeader]] || '',
    reason: row[headerIndices[reasonHeader]] || '',
    name: nameHeader ? row[headerIndices[nameHeader]] || '' : '',
    class: row[headerIndices[classHeader]] || ''
  }));

  return { header: cfg.questionHeader || answerHeader, entries: entries };
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
    
    let cfg = {};
    try {
      cfg = cfgFunc ? cfgFunc(sheetName) : {};
    } catch (configError) {
      console.log(`Config not found for sheet ${sheetName}, creating default config:`, configError.message);
      
      // ヘッダーから自動推測して基本設定を作成
      const headers = allValues[0] || [];
      const guessedConfig = guessHeadersFromArray(headers);
      
      // 少なくとも1つのヘッダーが推測できた場合、デフォルト設定を保存
      if (guessedConfig.questionHeader || guessedConfig.answerHeader) {
        try {
          saveSheetConfig(sheetName, guessedConfig);
          cfg = guessedConfig;
          console.log('Auto-created and saved config for sheet:', sheetName, cfg);
        } catch (saveError) {
          console.warn('Failed to save auto-created config:', saveError.message);
          cfg = guessedConfig; // 保存に失敗してもメモリ上のconfigは使用
        }
      }
    }
    
    const answerHeader = cfg.answerHeader || cfg.questionHeader || COLUMN_HEADERS.OPINION;
    const reasonHeader = cfg.reasonHeader || COLUMN_HEADERS.REASON;
    const classHeaderGuess = cfg.classHeader || COLUMN_HEADERS.CLASS;
    const nameHeader = cfg.nameHeader || '';

    const sheetHeaders = allValues[0].map(h => String(h || '').replace(/\s+/g,''));
    const normalized = header => String(header || '').replace(/\s+/g,'');

    const required = [
      COLUMN_HEADERS.EMAIL,
      answerHeader,
      reasonHeader,
      COLUMN_HEADERS.TIMESTAMP,
      COLUMN_HEADERS.UNDERSTAND,
      COLUMN_HEADERS.LIKE,
      COLUMN_HEADERS.CURIOUS,
      COLUMN_HEADERS.HIGHLIGHT
    ];
    if (nameHeader && sheetHeaders.includes(normalized(nameHeader))) {
      required.push(nameHeader);
    }

    if (sheetHeaders.includes(normalized(classHeaderGuess))) {
      required.splice(1,0,classHeaderGuess);
    }

    const headerIndices = findHeaderIndices(allValues[0], required);

    const classHeader = sheetHeaders.includes(normalized(classHeaderGuess)) ? classHeaderGuess : '';
    const dataRows = allValues.slice(1);
    const userEmail = safeGetUserEmail();
    const isAdmin = isUserAdmin(userEmail);

    const rows = dataRows.map((row, index) => {
      if (classHeader && classFilter && classFilter !== 'すべて') {
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
        let name = '';
        if (nameHeader && row[headerIndices[nameHeader]]) {
          name = row[headerIndices[nameHeader]];
        } else if (email) {
          name = email.split('@')[0];
        }
        return {
          rowIndex: index + 2,
          timestamp: timestamp,
          name: name, // 常に名前を含める（クライアント側で制御）
          email: email, // 常にemailを含める（クライアント側で制御）
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






function getAppSettingsForUser() {
  const props = PropertiesService.getUserProperties();
  const userId = props.getProperty('CURRENT_USER_ID');
  
  if (!userId) {
    // 従来の動作のフォールバック
    return {
      activeSheetName: '',
      showNames: false,
      showCounts: false
    };
  }
  
  // マルチテナント時
  const userInfo = getUserInfo(userId);
  if (!userInfo || !userInfo.configJson) {
    return {
      activeSheetName: '',
      showNames: false,
      showCounts: false
    };
  }
  
  // 下位互換性：古いsheetNameを新しいactiveSheetNameにマッピング
  const activeSheetName = userInfo.configJson.activeSheetName || userInfo.configJson.sheetName || '';
  
  return {
    activeSheetName: activeSheetName,
    showNames: typeof userInfo.configJson.showNames !== 'undefined' ? userInfo.configJson.showNames : (userInfo.configJson.showDetails || false),
    showCounts: typeof userInfo.configJson.showCounts !== 'undefined' ? userInfo.configJson.showCounts : (userInfo.configJson.showDetails || false)
  };
}


function clearHeaderCache(sheetName) {
  const cache = (typeof CacheService !== 'undefined') ? CacheService.getScriptCache() : null;
  if (cache) {
    try {
      cache.remove(`headers_${sheetName}`);
    } catch (e) {
      console.error('Cache removal failed:', e);
    }
  }
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
  clearHeaderCache(sheetName);
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

function getUrlOrigin(url) {
  const match = String(url).match(/^(https?:\/\/[^/]+)/);
  return match ? match[1] : '';
}

function convertPreviewUrl(url, deployId) {
  if (!url) return url;
  if (/script\.googleusercontent\.com/.test(url) && deployId) {
    const query = url.split('?')[1] || '';
    const base = `https://script.google.com/macros/s/${deployId}/exec`;
    return query ? `${base}?${query}` : base;
  }
  return url;
}

function getWebAppUrl() {
  const props = PropertiesService.getScriptProperties();
  // deployIdを最初に取得しており、見通しが良い
  const deployId = props.getProperty('DEPLOY_ID'); 
  let stored = (props.getProperty('WEB_APP_URL') || '').trim();

  // ★★★ このブロックが重要 ★★★
  // 保存済みのURLが/dev形式の場合、正規の/exec形式に変換して保存し直す自己修正機能。
  // これにより、プロパティに保存されるURLの形式が一貫する。
  if (stored) {
    const converted = convertPreviewUrl(stored, deployId);
    if (converted !== stored) {
      props.setProperties({ WEB_APP_URL: converted.trim() });
      stored = converted.trim();
    }
  }

  let current = '';
  try {
    if (typeof ScriptApp !== 'undefined') {
      current = ScriptApp.getService().getUrl();
    }
  } catch (e) {
    current = '';
  }

  if (current) {
    current = convertPreviewUrl(current, deployId);
    const currOrigin = getUrlOrigin(current);
    const storedOrigin = getUrlOrigin(stored);
    // 現在のURLと保存されているURLのドメインが異なる場合に更新するロジック
    if (!stored || (storedOrigin && currOrigin && currOrigin !== storedOrigin)) {
      props.setProperties({ WEB_APP_URL: current.trim() });
      stored = current.trim();
    }
  }
  return stored || current || '';
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
function checkRateLimit(action, userEmail) {
  try {
    const key = `rateLimit_${action}_${userEmail}`;
    const cache = (typeof CacheService !== 'undefined') ? CacheService.getScriptCache() : null;
    const attempts = cache ? parseInt(cache.get(key) || '0') : 0;
    
    if (attempts > 10) { // 10 attempts per hour
      throw new Error('レート制限に達しました。しばらく待ってから再試行してください。');
    }
    
    if (cache) {
      cache.put(key, String(attempts + 1), 3600);
    }
  } catch (error) {
    // Cache service error - continue without rate limiting
    console.warn('Rate limiting failed:', error);
  }
}

function validateSpreadsheetUrl(url) {
  if (!url || typeof url !== 'string') {
    throw new Error('無効なスプレッドシートURLです。');
  }
  
  const sanitizedUrl = url.trim();
  if (sanitizedUrl.length > 2048) {
    throw new Error('URLが長すぎます。');
  }
  
  return sanitizedUrl;
}

function createStudyQuestForm(userEmail) {
  try {
    // FormAppの利用可能性をチェック
    if (typeof FormApp === 'undefined') {
      throw new Error('Google Forms API is not available');
    }
    
    // 新しいGoogleフォームを作成
    const form = FormApp.create(`StudyQuest - 回答フォーム - ${userEmail.split('@')[0]}`);
    
    // フォームの説明を設定
    form.setDescription('StudyQuestで使用する回答フォームです。質問に対する回答を入力してください。');
    
    // 必要な項目を追加
    form.addTextItem()
        .setTitle('クラス')
        .setRequired(true);
    
    form.addTextItem()
        .setTitle('名前')
        .setRequired(true);
    
    form.addParagraphTextItem()
        .setTitle('回答')
        .setHelpText('質問に対するあなたの回答を入力してください')
        .setRequired(true);
    
    form.addParagraphTextItem()
        .setTitle('理由')
        .setHelpText('その回答を選んだ理由を教えてください')
        .setRequired(false);
    
    // フォームの回答先スプレッドシートを作成
    const spreadsheet = SpreadsheetApp.create(`StudyQuest - 回答データ - ${userEmail.split('@')[0]}`);
    form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheet.getId());
    const sheet = spreadsheet.getSheets()[0];
    
    // スプレッドシートに追加の列を準備
    const existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const additionalHeaders = [
      COLUMN_HEADERS.UNDERSTAND,
      COLUMN_HEADERS.LIKE,
      COLUMN_HEADERS.CURIOUS,
      COLUMN_HEADERS.HIGHLIGHT
    ];
    
    // 既存のヘッダーの後に追加の列を挿入
    const startCol = existingHeaders.length + 1;
    sheet.getRange(1, startCol, 1, additionalHeaders.length).setValues([additionalHeaders]);
    
    // ヘッダー行のフォーマット
    const allHeadersRange = sheet.getRange(1, 1, 1, existingHeaders.length + additionalHeaders.length);
    allHeadersRange.setFontWeight('bold');
    allHeadersRange.setBackground('#E3F2FD');
    
    // 列幅を調整
    try {
      sheet.autoResizeColumns(1, allHeadersRange.getNumColumns());
    } catch (e) {
      console.warn('Auto-resize failed:', e);
    }
    
    // Configシートも作成
    prepareSpreadsheetForStudyQuest(spreadsheet);
    
    return {
      formId: form.getId(),
      formUrl: form.getPublishedUrl(),
      spreadsheetId: spreadsheet.getId(),
      spreadsheetUrl: spreadsheet.getUrl(),
      editFormUrl: form.getEditUrl()
    };
  } catch (error) {
    console.error('Failed to create form and spreadsheet:', error);
    
    // FormAppが利用できない場合はスプレッドシートのみ作成
    if (error.message.includes('Google Forms API is not available') || error.message.includes('FormApp')) {
      console.warn('FormApp not available, creating spreadsheet only');
      return createStudyQuestSpreadsheetFallback(userEmail);
    }
    
    // 権限エラーの場合
    if (error.message.includes('permission') || error.message.includes('Permission')) {
      throw new Error('Googleフォーム作成の権限がありません。管理者にお問い合わせください。');
    }
    
    // その他のエラー詳細を含める
    throw new Error(`Googleフォームとスプレッドシートの作成に失敗しました。詳細: ${error.message}`);
  }
}

function createStudyQuestSpreadsheetFallback(userEmail) {
  try {
    // 新しいスプレッドシートを作成（フォームなし）
    const spreadsheet = SpreadsheetApp.create(`StudyQuest - 回答データ - ${userEmail.split('@')[0]}`);
    const sheet = spreadsheet.getActiveSheet();
    sheet.setName('回答データ');
    
    // ヘッダー行を設定
    const headers = [
      COLUMN_HEADERS.TIMESTAMP,
      COLUMN_HEADERS.EMAIL,
      COLUMN_HEADERS.CLASS,
      '質問',
      COLUMN_HEADERS.OPINION,
      COLUMN_HEADERS.REASON,
      COLUMN_HEADERS.NAME,
      COLUMN_HEADERS.UNDERSTAND,
      COLUMN_HEADERS.LIKE,
      COLUMN_HEADERS.CURIOUS,
      COLUMN_HEADERS.HIGHLIGHT
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // ヘッダー行のフォーマット
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#E3F2FD');
    
    // 列幅を調整
    try {
      sheet.autoResizeColumns(1, headers.length);
    } catch (e) {
      console.warn('Auto-resize failed:', e);
    }
    
    // Configシートも作成
    prepareSpreadsheetForStudyQuest(spreadsheet);
    
    return {
      spreadsheetId: spreadsheet.getId(),
      spreadsheetUrl: spreadsheet.getUrl(),
      formUrl: null, // フォームは作成されていない
      editFormUrl: null
    };
  } catch (error) {
    console.error('Failed to create fallback spreadsheet:', error);
    throw new Error('スプレッドシートの作成に失敗しました。');
  }
}

function prepareSpreadsheetForStudyQuest(spreadsheet) {
  // Configシートがなければ作成
  let configSheet = spreadsheet.getSheetByName('Config');
  if (!configSheet) {
    configSheet = spreadsheet.insertSheet('Config');
    const headers = ['表示シート名','問題文ヘッダー','回答ヘッダー','理由ヘッダー','名前列ヘッダー','クラス列ヘッダー'];
    configSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // 自動作成の場合は初期設定を追加
    const sheets = spreadsheet.getSheets().filter(s => s.getName() !== 'Config');
    if (sheets.length > 0) {
      const mainSheet = sheets[0];
      const sheetName = mainSheet.getName();
      
      // デフォルト設定を追加
      const defaultConfig = [
        sheetName,           // 表示シート名
        '質問',             // 問題文ヘッダー
        COLUMN_HEADERS.OPINION, // 回答ヘッダー
        COLUMN_HEADERS.REASON,  // 理由ヘッダー
        COLUMN_HEADERS.NAME,    // 名前列ヘッダー
        COLUMN_HEADERS.CLASS    // クラス列ヘッダー
      ];
      
      configSheet.getRange(2, 1, 1, headers.length).setValues([defaultConfig]);
    }
  }
}

function canAccessUserData(currentUser, userId) {
  // Users can always access their own data
  // Admins can access data for users in their domain
  try {
    const userInfo = getUserInfoInternal(userId);
    if (!userInfo) return false;
    
    // Check if current user is admin for this user's data
    return isSameDomain(currentUser, userInfo.adminEmail);
  } catch (error) {
    console.error('Access check failed:', error);
    return false;
  }
}

function filterSensitiveUserData(userInfo, currentUser) {
  if (!userInfo) return null;
  
  // Remove sensitive fields for non-admin users
  const filtered = { ...userInfo };
  
  // Always remove access token
  delete filtered.accessToken;
  
  // Only admins can see full user data
  if (!isSameDomain(currentUser, userInfo.adminEmail)) {
    delete filtered.spreadsheetId;
    delete filtered.spreadsheetUrl;
  }
  
  return filtered;
}

function getUserInfoInternal(userId) {
  if (!userId) return null;

  // Try cache first for performance
  const cache = (typeof CacheService !== 'undefined') ? CacheService.getScriptCache() : null;
  const cacheKey = `userInfo_${userId}`;
  
  try {
    if (cache) {
      const cached = cache.get(cacheKey);
      if (cached) {
        return JSON.parse(cached);
      }
    }
  } catch (e) {
    console.error('User cache retrieval failed:', e);
  }
  
  const userDb = getDatabase().getSheetByName(USER_DB_CONFIG.SHEET_NAME);
  const data = userDb.getDataRange().getValues();
  const headers = data[0];
  const userIdIndex = headers.indexOf('userId');
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][userIdIndex] === userId) {
      const userInfo = {};
      headers.forEach((header, index) => {
        if (header === 'configJson') {
          try {
            userInfo[header] = JSON.parse(data[i][index] || '{}');
          } catch (e) {
            console.error('Invalid JSON in configJson:', e);
            userInfo[header] = {};
          }
        } else {
          userInfo[header] = data[i][index];
        }
      });
      
      // Cache for 30 minutes
      try {
        if (cache) {
          cache.put(cacheKey, JSON.stringify(userInfo), 1800);
        }
      } catch (e) {
        console.error('User cache storage failed:', e);
      }
      
      return userInfo;
    }
  }
  
  return null;
}

/**
 * ユーザー情報を取得（セキュリティチェック付き）
 */
function getUserInfo(userId) {
  if (!userId) return null;
  
  try {
    const currentUser = safeGetUserEmail();
    
    // Authorization check
    if (!canAccessUserData(currentUser, userId)) {
      auditLog('UNAUTHORIZED_USER_ACCESS', userId, { currentUser });
      throw new Error('このユーザー情報にアクセスする権限がありません。');
    }
    
    const userInfo = getUserInfoInternal(userId);
    if (!userInfo) return null;
    
    // Update last accessed time
    const userDb = getDatabase().getSheetByName(USER_DB_CONFIG.SHEET_NAME);
    const data = userDb.getDataRange().getValues();
    const headers = data[0];
    const userIdIndex = headers.indexOf('userId');
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][userIdIndex] === userId) {
        userDb.getRange(i + 1, headers.indexOf('lastAccessedAt') + 1).setValue(new Date());
        break;
      }
    }
    
    return filterSensitiveUserData(userInfo, currentUser);
  } catch (error) {
    console.error('getUserInfo error:', error);
    throw error;
  }
}

/**
 * ユーザー設定を更新（セキュリティチェック付き）
 */
function updateUserConfig(userId, config) {
  // Use LockService only in production environment
  const lock = (typeof LockService !== 'undefined') ? LockService.getScriptLock() : null;
  try {
    if (lock && !lock.tryLock(10000)) {
      throw new Error('設定更新処理でロックを取得できませんでした。しばらく待ってから再試行してください。');
    }
    
    const currentUser = safeGetUserEmail();
    
    // Authorization check
    if (!canAccessUserData(currentUser, userId)) {
      auditLog('UNAUTHORIZED_CONFIG_UPDATE', userId, { currentUser, config });
      throw new Error('この設定を更新する権限がありません。');
    }
    
    // Validate config data
    if (!config || typeof config !== 'object') {
      throw new Error('無効な設定データです。');
    }
    
    const userDb = getDatabase().getSheetByName(USER_DB_CONFIG.SHEET_NAME);
    const data = userDb.getDataRange().getValues();
    const headers = data[0];
    const userIdIndex = headers.indexOf('userId');
    const configIndex = headers.indexOf('configJson');
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][userIdIndex] === userId) {
        const currentConfig = JSON.parse(data[i][configIndex] || '{}');
        
        // Sanitize config to prevent malicious data
        const sanitizedConfig = sanitizeConfigData(config);
        const newConfig = Object.assign({}, currentConfig, sanitizedConfig);
        
        userDb.getRange(i + 1, configIndex + 1).setValue(JSON.stringify(newConfig));
        
        // キャッシュをクリアして最新のデータが取得されるようにする
        const cache = (typeof CacheService !== 'undefined') ? CacheService.getScriptCache() : null;
        if (cache) {
          const cacheKey = `userInfo_${userId}`;
          cache.remove(cacheKey);
        }
        
        auditLog('CONFIG_UPDATED', userId, { currentUser, updatedFields: Object.keys(sanitizedConfig), newConfig });
        return newConfig;
      }
    }
    
    throw new Error('ユーザーが見つかりません。');
  } finally {
    if (lock) {
      lock.releaseLock();
    }
  }
}

function sanitizeConfigData(config) {
  // 新しいシステム：isPublishedを削除し、常に利用可能な回答ボード
  const allowedFields = ['activeSheetName', 'showNames', 'showCounts', 'showDetails'];
  const sanitized = {};

  for (const field of allowedFields) {
    if (config.hasOwnProperty(field)) {
      if (field === 'showDetails' || field === 'showNames' || field === 'showCounts') {
        sanitized[field] = Boolean(config[field]);
      } else if (field === 'activeSheetName') {
        const sheetName = config[field];
        if (typeof sheetName === 'string' && sheetName.length <= 100) {
          sanitized[field] = sheetName.trim();
        }
      }
    }
  }

  // showDetailsの下位互換
  if (sanitized.hasOwnProperty('showDetails')) {
    if (!sanitized.hasOwnProperty('showNames')) {
      sanitized.showNames = sanitized.showDetails;
    }
    if (!sanitized.hasOwnProperty('showCounts')) {
      sanitized.showCounts = sanitized.showDetails;
    }
  }

  // 下位互換性のためのマッピング
  if (config.hasOwnProperty('sheetName')) {
    const sheetName = config['sheetName'];
    if (typeof sheetName === 'string' && sheetName.length <= 100) {
      sanitized['activeSheetName'] = sheetName.trim();
    }
  }
  
  return sanitized;
}

// ===============================================================
// Database initialization and access
// ===============================================================

// 管理者が一度だけ実行
function setup() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('DATABASE_ID')) {
    console.log('Setup already done.');
    return;
  }
  const db = SpreadsheetApp.create('StudyQuest_UserDatabase');
  props.setProperty('DATABASE_ID', db.getId());
  const sheet = db.getSheets()[0];
  sheet.setName(USER_DB_CONFIG.SHEET_NAME);
  sheet.appendRow(USER_DB_CONFIG.HEADERS);
  console.log('Database created. ID: ' + db.getId());
}

// DBアクセスは必ずこの関数を経由させる
function getDatabase() {
  const dbId = PropertiesService.getScriptProperties().getProperty('DATABASE_ID');
  if (!dbId) {
    throw new Error('Database not initialized. Please run setup().');
  }
  return SpreadsheetApp.openById(dbId);
}

/**
 * ユーザーデータベースを取得（なければ作成）
 */
function getUserDatabase() {
  const db = getDatabase();
  const sheet = db.getSheetByName(USER_DB_CONFIG.SHEET_NAME);
  if (!sheet) {
    throw new Error('Database sheet not found. Please run setup().');
  }
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
  // Utilities.getRandomBytes is not available, use alternative approach
  const uuid = Utilities.getUuid();
  const timestamp = new Date().getTime();
  const randomString = uuid + timestamp + Math.random().toString(36);
  
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    randomString
  );
  
  return Utilities.base64EncodeWebSafe(bytes);
}

function auditLog(action, userId, details) {
  try {
    const timestamp = new Date();
    const currentUser = (typeof Session !== 'undefined' && Session.getActiveUser) ? Session.getActiveUser().getEmail() : '';

    // Use cache for temporary audit logging (since we can't create sheets dynamically in all environments)
    const cache = (typeof CacheService !== 'undefined') ? CacheService.getScriptCache() : null;
    const logEntry = {
      timestamp: timestamp.toISOString(),
      user: currentUser,
      action: action,
      userId: userId,
      details: details
    };
    
    // Store in cache with 6 hour expiration
    const uuid = (typeof Utilities !== 'undefined' && Utilities.getUuid) ? Utilities.getUuid() : Math.random().toString(36).slice(2);
    const cacheKey = `audit_${timestamp.getTime()}_${uuid}`;
    if (cache) {
      cache.put(cacheKey, JSON.stringify(logEntry), 21600);
    }
    
    // Also log to console for immediate debugging
    console.log(`AUDIT: ${action} by ${currentUser} for ${userId}`, details);
  } catch (error) {
    console.error('Audit logging failed:', error);
  }
}

/**
 * 既存ユーザーを検索
 */
function findUserByEmailAndSpreadsheet(email, spreadsheetId) {
  const userDb = getDatabase().getSheetByName(USER_DB_CONFIG.SHEET_NAME);
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
 * Find existing user by email only (used for single board per user).
 */

/**
 * 管理者メールアドレスを更新
 */
function updateAdminEmails(spreadsheetId, email) {
  try {
    const scriptProps = PropertiesService.getScriptProperties();
    const key = `ADMIN_EMAILS_${spreadsheetId}`;
    const current = scriptProps.getProperty(key) || '';
    const list = current.split(',').filter(Boolean);

    if (!list.includes(email)) {
      list.push(email);
      scriptProps.setProperty(key, list.join(','));
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

/**
 * Return existing board info for the current user if available.
 */
function getExistingBoard() {
  const email = safeGetUserEmail();
  const user = findUserByEmail(email);
  if (!user) return null;
  return {
    adminUrl: `${getWebAppUrl()}?userId=${user.userId}&mode=admin`,
    viewUrl: `${getWebAppUrl()}?userId=${user.userId}`,
    spreadsheetUrl: user.spreadsheetUrl,
    userId: user.userId
  };
}

if (typeof module !== 'undefined') {
  handleError = require('./ErrorHandling.gs').handleError;
  const { getConfig, saveSheetConfig, createConfigSheet } = require('./config.gs');
  module.exports = {
    COLUMN_HEADERS,
    getAdminSettings,
    doGet,
    buildBoardData,
    getConfig,
    getPublishedSheetData,
    getSheets,
    getSheetHeaders,
    getSheetData,
    getSheetDataForSpreadsheet,
    getHeaderIndices,
    addReaction,
    toggleHighlight,
    setShowDetails,
    setDisplayOptions,
    saveWebAppUrl,
    getWebAppUrl,
    saveDeployId,
    findHeaderIndices,
    parseReactionString,
    checkAdmin,
    isUserAdmin,
    handleError,
    saveSheetConfig,
    createConfigSheet,
    createTemplateSheet,
    prepareSheetForBoard,
    registerNewUser,
    getUserInfo,
    getUserInfoInternal,
    updateUserConfig,
    getUserDatabase,
    setup,
    getDatabase,
    extractSpreadsheetId,
    generateAccessToken,
    findUserByEmailAndSpreadsheet,
    findUserByEmail,
    getExistingBoard,
    updateAdminEmails,
    publishApp,
    unpublishApp,
    getActiveUserEmail,
    prepareSpreadsheetForStudyQuest,
    isSameDomain,
    getEmailDomain,
    getUrlOrigin,
    convertPreviewUrl,
    buildBoardData,
    guessHeadersFromArray,
    clearHeaderCache
  };
}
