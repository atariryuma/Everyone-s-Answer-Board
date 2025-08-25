/**
 * @fileoverview GAS APIラッパー - Google Apps Scriptの全サービスへの統一アクセス層
 */

/**
 * PropertiesServiceラッパー
 */
const Properties = {
  /**
   * スクリプトプロパティを取得
   * @returns {GoogleAppsScript.Properties.Properties}
   */
  getScriptProperties() {
    try {
      return PropertiesService.getScriptProperties();
    } catch (error) {
      logError(error, 'Properties.getScriptProperties', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.SYSTEM);
      throw error;
    }
  },
  
  /**
   * ユーザープロパティを取得
   * @returns {GoogleAppsScript.Properties.Properties}
   */
  getUserProperties() {
    try {
      return PropertiesService.getUserProperties();
    } catch (error) {
      logError(error, 'Properties.getUserProperties', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.SYSTEM);
      throw error;
    }
  },
  
  /**
   * スクリプトプロパティの値を取得
   * @param {string} key - プロパティキー
   * @returns {string|null}
   */
  getScriptProperty(key) {
    try {
      return PropertiesService.getScriptProperties().getProperty(key);
    } catch (error) {
      logError(error, 'Properties.getScriptProperty', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.SYSTEM, { key });
      return null;
    }
  },
  
  /**
   * スクリプトプロパティの値を設定
   * @param {string} key - プロパティキー
   * @param {string} value - 値
   */
  setScriptProperty(key, value) {
    try {
      PropertiesService.getScriptProperties().setProperty(key, value);
    } catch (error) {
      logError(error, 'Properties.setScriptProperty', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.SYSTEM, { key });
      throw error;
    }
  }
};

/**
 * SpreadsheetAppラッパー
 */
const Spreadsheet = {
  /**
   * スプレッドシートをIDで開く
   * @param {string} id - スプレッドシートID
   * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet}
   */
  openById(id) {
    try {
      return SpreadsheetApp.openById(id);
    } catch (error) {
      logError(error, 'Spreadsheet.openById', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.DATABASE, { spreadsheetId: id });
      throw error;
    }
  },
  
  /**
   * 新しいスプレッドシートを作成
   * @param {string} name - スプレッドシート名
   * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet}
   */
  create(name) {
    try {
      return SpreadsheetApp.create(name);
    } catch (error) {
      logError(error, 'Spreadsheet.create', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.DATABASE, { name });
      throw error;
    }
  },
  
  /**
   * アクティブなスプレッドシートを取得
   * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet}
   */
  getActive() {
    try {
      return SpreadsheetApp.getActiveSpreadsheet();
    } catch (error) {
      logError(error, 'Spreadsheet.getActive', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.DATABASE);
      return null;
    }
  }
};

/**
 * DriveAppラッパー
 */
const Drive = {
  /**
   * ファイルをIDで取得
   * @param {string} id - ファイルID
   * @returns {GoogleAppsScript.Drive.File}
   */
  getFileById(id) {
    try {
      return DriveApp.getFileById(id);
    } catch (error) {
      logError(error, 'Drive.getFileById', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.SYSTEM, { fileId: id });
      throw error;
    }
  },
  
  /**
   * フォルダをIDで取得
   * @param {string} id - フォルダID
   * @returns {GoogleAppsScript.Drive.Folder}
   */
  getFolderById(id) {
    try {
      return DriveApp.getFolderById(id);
    } catch (error) {
      logError(error, 'Drive.getFolderById', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.SYSTEM, { folderId: id });
      throw error;
    }
  },
  
  /**
   * ファイルを作成
   * @param {string} name - ファイル名
   * @param {string} content - ファイル内容
   * @returns {GoogleAppsScript.Drive.File}
   */
  createFile(name, content) {
    try {
      return DriveApp.createFile(name, content);
    } catch (error) {
      logError(error, 'Drive.createFile', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.SYSTEM, { name });
      throw error;
    }
  }
};

/**
 * HtmlServiceラッパー
 */
const Html = {
  /**
   * HTMLファイルからアウトプットを作成
   * @param {string} filename - ファイル名
   * @returns {GoogleAppsScript.HTML.HtmlOutput}
   */
  createHtmlOutputFromFile(filename) {
    try {
      return HtmlService.createHtmlOutputFromFile(filename);
    } catch (error) {
      logError(error, 'Html.createHtmlOutputFromFile', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.SYSTEM, { filename });
      throw error;
    }
  },
  
  /**
   * HTMLテンプレートファイルから作成
   * @param {string} filename - ファイル名
   * @returns {GoogleAppsScript.HTML.HtmlTemplate}
   */
  createTemplateFromFile(filename) {
    try {
      return HtmlService.createTemplateFromFile(filename);
    } catch (error) {
      logError(error, 'Html.createTemplateFromFile', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.SYSTEM, { filename });
      throw error;
    }
  },
  
  /**
   * HTML文字列からアウトプットを作成
   * @param {string} html - HTML文字列
   * @returns {GoogleAppsScript.HTML.HtmlOutput}
   */
  createHtmlOutput(html) {
    try {
      return HtmlService.createHtmlOutput(html);
    } catch (error) {
      logError(error, 'Html.createHtmlOutput', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.SYSTEM);
      throw error;
    }
  }
};

/**
 * CacheServiceラッパー（既存のCacheService.gsと統合）
 */
const Cache = {
  /**
   * スクリプトキャッシュを取得
   * @returns {GoogleAppsScript.Cache.Cache}
   */
  getScriptCache() {
    try {
      return CacheService.getScriptCache();
    } catch (error) {
      logError(error, 'Cache.getScriptCache', ERROR_SEVERITY.LOW, ERROR_CATEGORIES.CACHE);
      return null;
    }
  },
  
  /**
   * ユーザーキャッシュを取得
   * @returns {GoogleAppsScript.Cache.Cache}
   */
  getUserCache() {
    try {
      return CacheService.getUserCache();
    } catch (error) {
      logError(error, 'Cache.getUserCache', ERROR_SEVERITY.LOW, ERROR_CATEGORIES.CACHE);
      return null;
    }
  }
};

/**
 * LockServiceラッパー
 */
const Lock = {
  /**
   * スクリプトロックを取得
   * @returns {GoogleAppsScript.Lock.Lock}
   */
  getScriptLock() {
    try {
      return LockService.getScriptLock();
    } catch (error) {
      logError(error, 'Lock.getScriptLock', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.SYSTEM);
      throw error;
    }
  },
  
  /**
   * ユーザーロックを取得
   * @returns {GoogleAppsScript.Lock.Lock}
   */
  getUserLock() {
    try {
      return LockService.getUserLock();
    } catch (error) {
      logError(error, 'Lock.getUserLock', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.SYSTEM);
      throw error;
    }
  },
  
  /**
   * ロックを取得して実行
   * @param {Function} func - 実行する関数
   * @param {number} [timeout=10000] - タイムアウト（ミリ秒）
   * @returns {*} 関数の戻り値
   */
  withLock(func, timeout = 10000) {
    const lock = Lock.getScriptLock();
    try {
      lock.waitLock(timeout);
      return func();
    } catch (error) {
      logError(error, 'Lock.withLock', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.SYSTEM);
      throw error;
    } finally {
      lock.releaseLock();
    }
  }
};

/**
 * UrlFetchAppラッパー
 */
const Http = {
  /**
   * HTTPリクエストを送信
   * @param {string} url - URL
   * @param {Object} [params={}] - パラメータ
   * @returns {GoogleAppsScript.URL_Fetch.HTTPResponse}
   */
  fetch(url, params = {}) {
    try {
      // デフォルトオプションを設定
      const options = {
        muteHttpExceptions: true,
        ...params
      };
      
      return UrlFetchApp.fetch(url, options);
    } catch (error) {
      logError(error, 'Http.fetch', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.NETWORK, { url });
      throw error;
    }
  },
  
  /**
   * 複数のHTTPリクエストを並列実行
   * @param {Array<Object>} requests - リクエスト配列
   * @returns {Array<GoogleAppsScript.URL_Fetch.HTTPResponse>}
   */
  fetchAll(requests) {
    try {
      return UrlFetchApp.fetchAll(requests);
    } catch (error) {
      logError(error, 'Http.fetchAll', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.NETWORK);
      throw error;
    }
  }
};

/**
 * Utilitiesラッパー
 */
const Utils = {
  /**
   * UUIDを生成
   * @returns {string}
   */
  getUuid() {
    return Utilities.getUuid();
  },
  
  /**
   * Base64エンコード
   * @param {string} data - データ
   * @returns {string}
   */
  base64Encode(data) {
    return Utilities.base64Encode(data);
  },
  
  /**
   * Base64デコード
   * @param {string} data - エンコードされたデータ
   * @returns {string}
   */
  base64Decode(data) {
    return Utilities.base64Decode(data);
  },
  
  /**
   * スリープ
   * @param {number} milliseconds - ミリ秒
   */
  sleep(milliseconds) {
    Utilities.sleep(milliseconds);
  },
  
  /**
   * 日付をフォーマット
   * @param {Date} date - 日付
   * @param {string} timeZone - タイムゾーン
   * @param {string} format - フォーマット
   * @returns {string}
   */
  formatDate(date, timeZone, format) {
    return Utilities.formatDate(date, timeZone, format);
  }
};

/**
 * Sessionラッパー
 */
const SessionService = {
  /**
   * アクティブユーザーを取得
   * @returns {GoogleAppsScript.Base.User}
   */
  getActiveUser() {
    try {
      return Session.getActiveUser();
    } catch (error) {
      logError(error, 'SessionService.getActiveUser', ERROR_SEVERITY.LOW, ERROR_CATEGORIES.AUTHENTICATION);
      return null;
    }
  },
  
  /**
   * アクティブユーザーのメールを取得
   * @returns {string}
   */
  getActiveUserEmail() {
    try {
      return Session.getActiveUser().getEmail();
    } catch (error) {
      logError(error, 'SessionService.getActiveUserEmail', ERROR_SEVERITY.LOW, ERROR_CATEGORIES.AUTHENTICATION);
      return '';
    }
  },
  
  /**
   * 一時的なアクティブユーザーキーを取得
   * @returns {string}
   */
  getTemporaryActiveUserKey() {
    try {
      return Session.getTemporaryActiveUserKey();
    } catch (error) {
      logError(error, 'SessionService.getTemporaryActiveUserKey', ERROR_SEVERITY.LOW, ERROR_CATEGORIES.AUTHENTICATION);
      return '';
    }
  }
};