const js = require('@eslint/js');
const prettierConfig = require('eslint-config-prettier');

module.exports = [
  js.configs.recommended,
  prettierConfig,
  {
    files: ['src/**/*.{js,gs}', 'tests/**/*.js'],
    languageOptions: {
      ecmaVersion: 2021,
      sourceType: 'script',
      globals: {
        // GAS API globals
        SpreadsheetApp: 'readonly',
        DriveApp: 'readonly',
        DocumentApp: 'readonly',
        FormApp: 'readonly',
        SlidesApp: 'readonly',
        Logger: 'readonly',
        UrlFetchApp: 'readonly',
        PropertiesService: 'readonly',
        CacheService: 'readonly',
        Utilities: 'readonly',
        HtmlService: 'readonly',
        ContentService: 'readonly',
        ScriptApp: 'readonly',
        MailApp: 'readonly',
        GmailApp: 'readonly',
        Session: 'readonly',
        LockService: 'readonly',
        Browser: 'readonly',
        CalendarApp: 'readonly',
        CardService: 'readonly',
        Charts: 'readonly',
        ContactsApp: 'readonly',
        DataStudioApp: 'readonly',
        GroupsApp: 'readonly',
        LanguageApp: 'readonly',
        MapsApp: 'readonly',
        SitesApp: 'readonly',
        console: 'readonly',
        google: 'readonly',

        // Node.js globals for tests
        process: 'readonly',
        Buffer: 'readonly',
        global: 'readonly',
        __dirname: 'readonly',
        __filename: 'readonly',
        exports: 'writable',
        module: 'writable',
        require: 'readonly'
      }
    },
    rules: {
      'no-var': 'error',
      'prefer-const': 'error',
      'prefer-arrow-callback': 'warn',
      'prefer-template': 'warn',
      'object-shorthand': 'warn',
      'prefer-destructuring': 'warn',
      'no-console': 'off',
      'no-unused-vars': ['warn', { argsIgnorePattern: '^_' }]
    }
  },
  {
    files: ['tests/**/*.js'],
    languageOptions: {
      globals: {
        // Jest globals
        jest: 'readonly',
        describe: 'readonly',
        test: 'readonly',
        it: 'readonly',
        expect: 'readonly',
        beforeEach: 'readonly',
        afterEach: 'readonly',
        beforeAll: 'readonly',
        afterAll: 'readonly'
      }
    }
  }
];