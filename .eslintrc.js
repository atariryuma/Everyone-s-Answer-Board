module.exports = {
  env: {
    es6: true,
    node: true,
    jest: true
  },
  extends: [
    'eslint:recommended'
  ],
  parserOptions: {
    ecmaVersion: 2020,
    sourceType: 'module'
  },
  rules: {
    // GAS環境では未使用関数がHTMLから動的に呼ばれるため無効化
    'no-unused-vars': 'off',

    // テストファイルでの未定義グローバル変数を許可
    'no-undef': ['error', {
      'typeof': false
    }]
  },
  overrides: [
    // Google Apps Script ファイル用設定
    {
      files: ['src/**/*.gs'],
      env: {
        es6: true,
        browser: false,
        node: false,
        googleappsscript: true
      },
      parserOptions: {
        ecmaVersion: 2020,
        sourceType: 'script'  // GASはモジュール非対応
      },
      globals: {
        // GAS API globals
        'SpreadsheetApp': 'readonly',
        'PropertiesService': 'readonly', 
        'CacheService': 'readonly',
        'Session': 'readonly',
        'Utilities': 'readonly',
        'ScriptApp': 'readonly',
        'FormApp': 'readonly',
        'HtmlService': 'readonly',
        'ContentService': 'readonly',
        'UrlFetchApp': 'readonly',
        'DriveApp': 'readonly',
        'GmailApp': 'readonly',
        'CalendarApp': 'readonly',
        'DocumentApp': 'readonly',
        'Logger': 'readonly',
        'Browser': 'readonly',
        
        // Custom application globals (プロジェクト固有)
        'CORE': 'readonly',
        'DB_CONFIG': 'readonly', 
        'PROPS_KEYS': 'readonly',
        'SYSTEM_CONSTANTS': 'readonly',
        'DB': 'writable',
        'App': 'writable',
        'Services': 'writable',
        'UserManager': 'writable',
        'ConfigManager': 'writable',
        'SystemManager': 'writable',
        'SecurityValidator': 'writable',
        'cacheManager': 'writable',
        'getSheetsService': 'writable',
        'getSheetsServiceCached': 'writable',
        'getServiceAccountTokenCached': 'writable',
        'getSecureDatabaseId': 'writable',
        'getSecureServiceAccountCreds': 'writable',
        'batchGetSheetsData': 'writable',
        'updateSheetsData': 'writable',
        'initializeDatabaseSheet': 'writable',
        'getWebAppUrl': 'writable',
        'resolveColumnConflicts': 'writable',
        'getSpreadsheetColumnIndices': 'writable',
        'updateUserLastAccess': 'writable',
        'executeGetPublishedSheetData': 'writable',
        'invalidateUserCache': 'writable',
        'createResponse': 'writable',
        'validateAccess': 'writable',
        'validateAccessAPI': 'writable',
        'verifyAccess': 'writable',
        'buildUserAdminUrl': 'writable',
        'logError': 'writable',
        'debugLog': 'writable',
        'warnLog': 'writable', 
        'infoLog': 'writable',
        'ULog': 'writable',
        'UError': 'writable',
        'UValidate': 'writable',
        'UExecute': 'writable',
        'UData': 'writable',
        'UnifiedValidation': 'writable',
        'resilientExecutor': 'writable',
        'resilientUrlFetch': 'writable',
        'resilientSpreadsheetOperation': 'writable',
        'resilientCacheOperation': 'writable',
        'getCurrentUserEmail': 'writable',
        'findUserById': 'writable',
        'checkLoginStatus': 'writable',
        'updateLoginStatus': 'writable',
        'clearTimeout': 'writable',
        'LockService': 'readonly',
        'WebApp': 'readonly',
        'RequestGate': 'writable'
      },
      rules: {
        // GAS特有のルール調整 - GASでは未使用関数が後でHTMLから呼ばれる可能性があるため無効化
        'no-unused-vars': 'off',
        'no-undef': 'warn',
        'no-var': 'warn',
        'prefer-const': 'warn',
        'no-console': 'off'
      }
    },
    {
      files: ['tests/**/*.js'],
      env: {
        node: true,
        jest: true,
        browser: true  // setTimeout などのWebAPI用
      },
      globals: {
        // Browser/Node globals
        'setTimeout': 'readonly',
        'clearTimeout': 'readonly',
        'setInterval': 'readonly',
        'clearInterval': 'readonly',
        
        // GAS API globals
        'SpreadsheetApp': 'readonly',
        'PropertiesService': 'readonly',
        'CacheService': 'readonly',
        'Session': 'readonly',
        'Utilities': 'readonly',
        'ScriptApp': 'readonly',
        'FormApp': 'readonly',
        'HtmlService': 'readonly',
        'ContentService': 'readonly',
        'UrlFetchApp': 'readonly',
        
        // Custom application globals
        'logError': 'readonly',
        'debugLog': 'readonly',
        'warnLog': 'readonly',
        'infoLog': 'readonly',
        'getCacheValue': 'readonly',
        'setCacheValue': 'readonly',
        'removeCacheValue': 'readonly',
        'clearCacheValue': 'readonly',
        'clearCacheByPattern': 'readonly',
        'getUserById': 'readonly',
        'findUserById': 'readonly',
        'createUser': 'readonly',
        'updateUser': 'readonly',
        'handleCoreApiRequest': 'readonly',
        'clearExecutionUserInfoCache': 'readonly',
        'doGet': 'readonly',
        'getCurrentUserStatus': 'readonly'
      }
    }
  ]
};