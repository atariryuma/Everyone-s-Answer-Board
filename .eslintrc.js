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
    // 未使用変数の警告
    'no-unused-vars': ['warn', { 
      argsIgnorePattern: '^_',
      varsIgnorePattern: '^_'
    }],
    
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
        'WebApp': 'readonly'
      },
      rules: {
        // GAS特有のルール調整
        'no-unused-vars': ['warn', { 
          argsIgnorePattern: '^_',
          varsIgnorePattern: '^_'
        }],
        'no-undef': 'warn',
        'no-var': 'warn',          // GAS V8でもconstを推奨
        'prefer-const': 'warn',    // 再代入しない変数はconst
        'no-console': 'off'        // console.logはGASで使用可能
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