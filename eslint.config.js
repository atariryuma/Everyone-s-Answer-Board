/**
 * ESLint flat config — GAS backend code.
 *
 * Why: HTML テンプレート内 JS は ESLint で扱いにくいので、まずは src/*.js
 *   （GAS バックエンド）に限定して走らせる。フロント JS は scripts/lint.js の
 *   カスタムルール + ブラウザ DevTools + Cloud Logging（reportClientError 経由）
 *   でカバー。
 *
 * 既存負債を吸収する移行期間として、規則は控えめにする：
 *   - error: バグの可能性が高いもの (no-undef, no-unused-vars, no-redeclare)
 *   - warn:  スタイル・推奨 (prefer-const, no-var, eqeqeq)
 *   - off:   GAS では問題ない / 既存に大量にあるもの
 */
const globals = require('globals');

module.exports = [
  {
    ignores: ['node_modules/**', '.clasp/**', 'docs/**', 'scripts/**', 'tests/**'],
  },
  {
    files: ['src/*.js'],
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: 'script',
      globals: {
        // GAS V8 ランタイム ービス（appsscript.json の oauthScopes に紐づく）
        SpreadsheetApp: 'readonly',
        DriveApp: 'readonly',
        FormApp: 'readonly',
        DocumentApp: 'readonly',
        ScriptApp: 'readonly',
        UrlFetchApp: 'readonly',
        Utilities: 'readonly',
        Session: 'readonly',
        PropertiesService: 'readonly',
        CacheService: 'readonly',
        LockService: 'readonly',
        HtmlService: 'readonly',
        ContentService: 'readonly',
        Logger: 'readonly',
        Browser: 'readonly',
        MailApp: 'readonly',
        GmailApp: 'readonly',
        CalendarApp: 'readonly',
        // ES2022 globals (URL は GAS V8 で利用可能)
        ...globals.browser,
        ...globals.es2022,
        // 単一グローバルスコープ — 他ファイルで定義された関数 / 定数
        // これらは src/ 内で `/* global ... */` で逐次宣言されているが、
        // 全部列挙すると膨大なので no-undef は warn に降格 して許容する。
      },
    },
    rules: {
      // バグ予防（error）
      'no-dupe-keys': 'error',
      'no-dupe-args': 'error',
      'no-duplicate-case': 'error',
      'no-unreachable': 'error',
      'no-unsafe-negation': 'error',
      'no-irregular-whitespace': 'error',
      'use-isnan': 'error',
      'no-cond-assign': 'error',
      'no-self-assign': 'error',
      'no-self-compare': 'error',
      'no-unsafe-finally': 'error',

      // GAS 単一スコープに合わせた緩和
      // Why: 全関数が他ファイルから呼ばれうるため「未使用」は判定不可能。
      //      また /* global X */ と function X() の併記はファイル間依存の宣言慣習で
      //      no-redeclare が誤検出する。両者とも warn に降格して情報として残す。
      'no-redeclare': 'warn',
      'no-unused-vars': ['warn', {
        args: 'none',
        vars: 'local',                // file-local var/let のみ追跡、function 宣言は無視
        varsIgnorePattern: '^_',
        caughtErrors: 'none',
      }],
      'no-undef': 'off',
      'no-empty': ['warn', { allowEmptyCatch: true }],

      // スタイル（warn）
      'prefer-const': 'warn',
      'no-var': 'warn',
      'eqeqeq': ['warn', 'smart'],

      // 明示的に off
      'no-console': 'off',
      'no-prototype-builtins': 'off',
    },
  },
];
