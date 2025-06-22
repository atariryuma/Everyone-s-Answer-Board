var admin;
// When running in Node, load dependencies with require. In Apps Script the
// referenced functions are available globally, so create a simple wrapper
// object to access them.
if (typeof module !== 'undefined') {
  admin = require('./admin.gs');
  var handleError = require('./ErrorHandling.gs').handleError;
} else {
  admin = {
    safeGetUserEmail: safeGetUserEmail,
    getAppSettings: getAppSettings,
    checkAdmin: checkAdmin
  };
}
const REACTION_KEYS = ['UNDERSTAND','LIKE','CURIOUS'];

function parseReactionString(val) {
  if (!val) return [];
  return val
    .toString()
    .split(',')
    .map(s => s.trim())
    .filter(Boolean);
}

function addReaction(rowIndex, reactionKey) {
  rowIndex = Number(rowIndex);
  if (typeof module !== 'undefined') {
    var { COLUMN_HEADERS, getHeaderIndices, TIME_CONSTANTS } = require('./Code.gs');
  }
  if (!rowIndex || !reactionKey || !COLUMN_HEADERS[reactionKey]) {
    return { status: 'error', message: '無効なパラメータです。' };
  }
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(TIME_CONSTANTS.LOCK_WAIT_MS)) {
      return { status: 'error', message: '他のユーザーが操作中です。しばらく待ってから再試行してください。' };
    }
    const userEmail = admin.safeGetUserEmail();
    if (!userEmail) {
      return { status: 'error', message: 'ログインしていないため、操作できません。' };
    }
    const settings = admin.getAppSettings();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(settings.activeSheetName);
    if (!sheet) throw new Error(`シート '${settings.activeSheetName}' が見つかりません。`);

    const headerIndices = getHeaderIndices(settings.activeSheetName);
    const lists = {};
    REACTION_KEYS.forEach(k => {
      const col = headerIndices[COLUMN_HEADERS[k]] + 1;
      const val = sheet.getRange(rowIndex, col).getValue();
      lists[k] = { arr: parseReactionString(val), col: col };
    });

    const wasReacted = lists[reactionKey].arr.includes(userEmail);
    Object.keys(lists).forEach(key => {
      const idx = lists[key].arr.indexOf(userEmail);
      if (idx > -1) lists[key].arr.splice(idx, 1);
    });
    if (!wasReacted) {
      lists[reactionKey].arr.push(userEmail);
    }
    REACTION_KEYS.forEach(k => {
      sheet.getRange(rowIndex, lists[k].col).setValue(lists[k].arr.join(','));
    });
    if (typeof SpreadsheetApp !== 'undefined' && typeof SpreadsheetApp.flush === 'function') {
      SpreadsheetApp.flush();
    }

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
  rowIndex = Number(rowIndex);
  if (typeof module !== 'undefined') {
    var { COLUMN_HEADERS, getHeaderIndices, TIME_CONSTANTS } = require('./Code.gs');
  }
  if (!admin.checkAdmin()) {
    return { status: 'error', message: '権限がありません。' };
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(TIME_CONSTANTS.LOCK_WAIT_MS);
  try {
    const sheetName = admin.getAppSettings().activeSheetName;
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

if (typeof module !== 'undefined') {
  module.exports = { parseReactionString, addReaction, toggleHighlight };
}
