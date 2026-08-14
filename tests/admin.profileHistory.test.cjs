/**
 * Option B: profileHistory append + sanitize tests.
 *
 * Why: 教師が profile 切替 → history append → 生徒の wire に乗る、という data flow
 *   全体の根拠は __appendProfileHistory_ と sanitizeProfileHistory が正しく動くこと。
 *   重複防止 / 削除済 entry の除去 / cap 50 件 / 直前と同名は no-op の挙動を回帰防止する。
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('path');
const vm = require('vm');
const { gasResponseStubs } = require('./_helpers.cjs');

function loadAdminContext(overrides = {}) {
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    ...gasResponseStubs(),
    getCurrentEmail: () => 'admin@example.com',
    isAdministrator: () => true,
    findUserByEmail: () => ({ userId: 'u1', userEmail: 'admin@example.com' }),
    findUserById: () => ({ userId: 'u1', userEmail: 'admin@example.com', isActive: true }),
    getUserConfig: () => ({ success: true, config: {} }),
    saveUserConfig: () => ({ success: true }),
    requireAdmin: () => ({ email: 'admin@example.com', isAdmin: true }),
    getConfigOrDefault: () => ({}),
    getBatchedAdminAuth: () => ({ success: true, email: 'admin@example.com', isAdmin: true }),
    DEFAULT_DISPLAY_SETTINGS: { showNames: false, showReactions: true, theme: 'default', pageSize: 20 },
    PropertiesService: {
      getScriptProperties: () => ({ getProperty: () => null, setProperty: () => {}, getProperties: () => ({}) })
    },
    getCachedProperty: () => null,
    ...overrides
  };
  vm.createContext(context);
  const source = fs.readFileSync(path.resolve(__dirname, '../src/AdminApis.js'), 'utf8');
  vm.runInContext(source, context, { filename: 'AdminApis.js' });
  return context;
}

function loadConfigContext(overrides = {}) {
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    SYSTEM_LIMITS: { PREVIEW_LENGTH: 100, DEFAULT_PAGE_SIZE: 20, MAX_PAGE_SIZE: 100 },
    DEFAULT_DISPLAY_SETTINGS: { showNames: false, showReactions: true, theme: 'default', pageSize: 20 },
    validateConfig: () => ({ isValid: true, sanitized: {}, errors: [] }),
    VALIDATOR_BOARD_MODES: ['auto', 'board', 'numberline', 'matrix', 'wordcloud', 'pie'],
    CacheService: { getScriptCache: () => ({ get: () => null, put: () => {}, remove: () => {}, removeAll: () => {} }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
    getCachedProperty: () => null,
    getWebAppUrl: () => 'https://example.com/exec',
    ...overrides
  };
  vm.createContext(context);
  const source = fs.readFileSync(path.resolve(__dirname, '../src/ConfigService.js'), 'utf8');
  vm.runInContext(source, context, { filename: 'ConfigService.js' });
  return context;
}

