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
    {
      files: ['tests/**/*.js'],
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