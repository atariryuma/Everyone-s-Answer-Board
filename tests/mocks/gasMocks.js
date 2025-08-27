/**
 * Google Apps Script API モック
 * JestでGAS APIをテストするためのモック実装
 */

// SpreadsheetApp モック
global.SpreadsheetApp = {
  getActiveSpreadsheet: jest.fn(() => ({
    getSheetByName: jest.fn((name) => ({
      getRange: jest.fn((row, col, numRows, numCols) => ({
        getValues: jest.fn(() => Array(numRows || 1).fill(Array(numCols || 1).fill(''))),
        setValues: jest.fn(),
        getValue: jest.fn(() => ''),
        setValue: jest.fn(),
        clear: jest.fn(),
      })),
      getLastRow: jest.fn(() => 10),
      getLastColumn: jest.fn(() => 5),
      appendRow: jest.fn(),
      deleteRow: jest.fn(),
      insertRowAfter: jest.fn(),
      getName: jest.fn(() => 'Sheet1'),
    })),
    getSheets: jest.fn(() => []),
    insertSheet: jest.fn(),
    deleteSheet: jest.fn(),
  })),
  getActiveSheet: jest.fn(() => ({
    getRange: jest.fn(() => ({
      getValues: jest.fn(() => [[]]),
      setValues: jest.fn(),
    })),
    getLastRow: jest.fn(() => 10),
    getLastColumn: jest.fn(() => 5),
  })),
  create: jest.fn((_name) => ({
    getName: jest.fn(() => _name),
  })),
  openById: jest.fn((_id) => ({
    getId: jest.fn(() => _id),
  })),
};

// PropertiesService モック
global.PropertiesService = {
  getScriptProperties: jest.fn(() => ({
    getProperty: jest.fn((_key) => null),
    setProperty: jest.fn(),
    deleteProperty: jest.fn(),
    getProperties: jest.fn(() => ({})),
    setProperties: jest.fn(),
    deleteAllProperties: jest.fn(),
  })),
  getUserProperties: jest.fn(() => ({
    getProperty: jest.fn((_key) => null),
    setProperty: jest.fn(),
    deleteProperty: jest.fn(),
    getProperties: jest.fn(() => ({})),
    setProperties: jest.fn(),
    deleteAllProperties: jest.fn(),
  })),
  getDocumentProperties: jest.fn(() => ({
    getProperty: jest.fn((_key) => null),
    setProperty: jest.fn(),
    deleteProperty: jest.fn(),
    getProperties: jest.fn(() => ({})),
    setProperties: jest.fn(),
    deleteAllProperties: jest.fn(),
  })),
};

// CacheService モック
global.CacheService = {
  getScriptCache: jest.fn(() => ({
    get: jest.fn((_key) => null),
    put: jest.fn(),
    remove: jest.fn(),
    getAll: jest.fn(() => ({})),
    putAll: jest.fn(),
    removeAll: jest.fn(),
  })),
  getUserCache: jest.fn(() => ({
    get: jest.fn((_key) => null),
    put: jest.fn(),
    remove: jest.fn(),
    getAll: jest.fn(() => ({})),
    putAll: jest.fn(),
    removeAll: jest.fn(),
  })),
  getDocumentCache: jest.fn(() => ({
    get: jest.fn((_key) => null),
    put: jest.fn(),
    remove: jest.fn(),
    getAll: jest.fn(() => ({})),
    putAll: jest.fn(),
    removeAll: jest.fn(),
  })),
};

// UrlFetchApp モック
global.UrlFetchApp = {
  fetch: jest.fn((_url, _params) => ({
    getResponseCode: jest.fn(() => 200),
    getContentText: jest.fn(() => JSON.stringify({ success: true })),
    getBlob: jest.fn(),
    getHeaders: jest.fn(() => ({})),
    getAllHeaders: jest.fn(() => ({})),
  })),
};

// HtmlService モック
global.HtmlService = {
  createTemplateFromFile: jest.fn((filename) => ({
    evaluate: jest.fn(() => ({
      getContent: jest.fn(() => `<html>${filename}</html>`),
      setTitle: jest.fn(),
      setSandboxMode: jest.fn(),
      setWidth: jest.fn(),
      setHeight: jest.fn(),
      addMetaTag: jest.fn(),
    })),
    filename,
  })),
  createHtmlOutput: jest.fn((html) => ({
    getContent: jest.fn(() => html),
    setTitle: jest.fn(),
    setSandboxMode: jest.fn(),
    setWidth: jest.fn(),
    setHeight: jest.fn(),
    addMetaTag: jest.fn(),
  })),
  SandboxMode: {
    IFRAME: 'IFRAME',
    NATIVE: 'NATIVE',
    EMULATED: 'EMULATED',
  },
};

// Utilities モック
global.Utilities = {
  formatDate: jest.fn((_date, _timeZone, _format) => 'formatted date'),
  formatString: jest.fn((template, ..._args) => template),
  getUuid: jest.fn(() => 'mock-uuid-1234'),
  base64Encode: jest.fn((data) => Buffer.from(data).toString('base64')),
  base64Decode: jest.fn((data) => Buffer.from(data, 'base64').toString()),
  newBlob: jest.fn((data) => ({ getBytes: () => data })),
  parseCsv: jest.fn((_csv) => [['col1', 'col2']]),
  sleep: jest.fn(),
  jsonParse: jest.fn((str) => JSON.parse(str)),
  jsonStringify: jest.fn((obj) => JSON.stringify(obj)),
};

// Logger モック
global.Logger = {
  log: jest.fn(console.log),
  clear: jest.fn(),
  getLog: jest.fn(() => ''),
};

// ScriptApp モック
global.ScriptApp = {
  newTrigger: jest.fn((_functionName) => ({
    timeBased: jest.fn(() => ({
      everyMinutes: jest.fn(() => ({ create: jest.fn() })),
      everyHours: jest.fn(() => ({ create: jest.fn() })),
      everyDays: jest.fn(() => ({ create: jest.fn() })),
      atHour: jest.fn(() => ({ create: jest.fn() })),
    })),
    forSpreadsheet: jest.fn(() => ({
      onEdit: jest.fn(() => ({ create: jest.fn() })),
      onChange: jest.fn(() => ({ create: jest.fn() })),
      onFormSubmit: jest.fn(() => ({ create: jest.fn() })),
    })),
  })),
  deleteTrigger: jest.fn(),
  getProjectTriggers: jest.fn(() => []),
  getAuthorizationInfo: jest.fn(() => ({
    getAuthorizationStatus: jest.fn(),
    getAuthorizationUrl: jest.fn(),
  })),
};

// DriveApp モック
global.DriveApp = {
  getFileById: jest.fn((_id) => ({
    getName: jest.fn(() => 'file.txt'),
    getBlob: jest.fn(),
    getId: jest.fn(() => _id),
    getUrl: jest.fn(() => `https://drive.google.com/file/d/${_id}`),
    setSharing: jest.fn(),
    makeCopy: jest.fn(),
  })),
  getFolderById: jest.fn((id) => ({
    getName: jest.fn(() => 'folder'),
    createFile: jest.fn(),
    createFolder: jest.fn(),
    getFiles: jest.fn(() => ({
      hasNext: jest.fn(() => false),
      next: jest.fn(),
    })),
  })),
  createFile: jest.fn((_blob) => ({
    getId: jest.fn(() => 'new-file-id'),
  })),
  createFolder: jest.fn((name) => ({
    getId: jest.fn(() => 'new-folder-id'),
    getName: jest.fn(() => name),
  })),
};

// Session モック
global.Session = {
  getActiveUser: jest.fn(() => ({
    getEmail: jest.fn(() => 'user@example.com'),
  })),
  getEffectiveUser: jest.fn(() => ({
    getEmail: jest.fn(() => 'user@example.com'),
  })),
  getTemporaryActiveUserKey: jest.fn(() => 'temp-key'),
  getTimeZone: jest.fn(() => 'Asia/Tokyo'),
  getScriptTimeZone: jest.fn(() => 'Asia/Tokyo'),
};

// LockService モック
global.LockService = {
  getScriptLock: jest.fn(() => ({
    tryLock: jest.fn(() => true),
    releaseLock: jest.fn(),
    hasLock: jest.fn(() => true),
    waitLock: jest.fn(),
  })),
  getUserLock: jest.fn(() => ({
    tryLock: jest.fn(() => true),
    releaseLock: jest.fn(),
    hasLock: jest.fn(() => true),
    waitLock: jest.fn(),
  })),
  getDocumentLock: jest.fn(() => ({
    tryLock: jest.fn(() => true),
    releaseLock: jest.fn(),
    hasLock: jest.fn(() => true),
    waitLock: jest.fn(),
  })),
};

// console モック（GAS環境用）
if (!global.console) {
  global.console = {
    log: jest.fn(),
    error: jest.fn(),
    warn: jest.fn(),
    info: jest.fn(),
    debug: jest.fn(),
  };
}

module.exports = {
  SpreadsheetApp: global.SpreadsheetApp,
  PropertiesService: global.PropertiesService,
  CacheService: global.CacheService,
  UrlFetchApp: global.UrlFetchApp,
  HtmlService: global.HtmlService,
  Utilities: global.Utilities,
  Logger: global.Logger,
  ScriptApp: global.ScriptApp,
  DriveApp: global.DriveApp,
  Session: global.Session,
  LockService: global.LockService,
  console: global.console,
};