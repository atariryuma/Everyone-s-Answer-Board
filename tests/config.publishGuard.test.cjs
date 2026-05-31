const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('path');
const vm = require('vm');

/**
 * Tests for the publish-state single-source-of-truth guard in saveUserConfig (v2865).
 *
 * Invariant (CLAUDE.md): config.isPublished / publishedAt may only be written by the 4
 *   lifecycle functions (publishApp / republishMyBoard / unpublishBoard /
 *   toggleUserBoardStatus), which pass options.__allowPublishStateWrite. Any other save
 *   must NOT be able to change publish state — it is forced back to the persisted value.
 *
 * Why a guard (not a strip): removing the field would flip isUserBoardPublished to false
 *   (= silent unpublish / data loss). The guard overwrites with the *persisted* value.
 */

function loadConfigContext(persistedConfig) {
  let saved = null; // captured configJson string
  const user = {
    userId: 'u1',
    userEmail: 'owner@example.com',
    spreadsheetId: 'sheet1',
    configJson: persistedConfig == null ? null : JSON.stringify(persistedConfig)
  };
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    CacheService: {
      getScriptCache: () => ({ removeAll: () => {}, remove: () => {}, get: () => null, put: () => {} })
    },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null, setProperty: () => {} }) },
    SYSTEM_LIMITS: { DEFAULT_PAGE_SIZE: 20, MAX_PAGE_SIZE: 100, CONFIG_JSON_MAX_CHARS: 32768, RADIX_DECIMAL: 10 },
    DEFAULT_DISPLAY_SETTINGS: { showNames: false, showReactions: true, theme: 'default', pageSize: 20 },
    Utilities: { getUuid: () => 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee' },
    logError_: () => {},
    getCachedProperty: () => null,
    getCurrentEmail: () => 'owner@example.com',
    isAdministrator: () => false,
    findUserById: () => user,
    findUserByEmail: () => user,
    findUserBySpreadsheetId: () => null,
    validateConfigUserId: () => true,
    // permissive validator: pass every field through (mirrors validators.js isPublished passthrough)
    validateConfig: (cfg) => ({ isValid: true, errors: [], warnings: [], sanitized: { ...cfg } }),
    updateUser: (_id, patch) => { saved = patch.configJson; return { success: true }; },
    URL,
    __getSaved: () => (saved ? JSON.parse(saved) : null)
  };
  vm.createContext(context);
  const source = fs.readFileSync(path.resolve(__dirname, '../src/ConfigService.js'), 'utf8');
  vm.runInContext(source, context, { filename: 'ConfigService.js' });
  return context;
}

test('guard: non-sentinel save CANNOT unpublish a published board (forces persisted true)', () => {
  const ctx = loadConfigContext({ isPublished: true, publishedAt: '2020-01-01T00:00:00.000Z', spreadsheetId: 'sheet1' });
  const res = ctx.saveUserConfig('u1', { isPublished: false, publishedAt: null, spreadsheetId: 'sheet1' });
  assert.equal(res.success, true);
  const saved = ctx.__getSaved();
  assert.equal(saved.isPublished, true, 'published state must be preserved');
  assert.equal(saved.publishedAt, '2020-01-01T00:00:00.000Z', 'publishedAt must reflect persisted state');
});

test('guard: non-sentinel save CANNOT publish an unpublished board (forces persisted false)', () => {
  const ctx = loadConfigContext({ isPublished: false, publishedAt: null, spreadsheetId: 'sheet1' });
  const res = ctx.saveUserConfig('u1', { isPublished: true, publishedAt: '2026-05-31T00:00:00.000Z', spreadsheetId: 'sheet1' });
  assert.equal(res.success, true);
  const saved = ctx.__getSaved();
  assert.equal(saved.isPublished, false, 'unpublished state must be preserved');
  assert.equal(saved.publishedAt, null, 'publishedAt must be null when unpublished');
});

test('guard: __allowPublishStateWrite honors the new published state', () => {
  const ctx = loadConfigContext({ isPublished: false, publishedAt: null, spreadsheetId: 'sheet1' });
  const res = ctx.saveUserConfig('u1',
    { isPublished: true, publishedAt: '2026-05-31T00:00:00.000Z', spreadsheetId: 'sheet1' },
    { __allowPublishStateWrite: true });
  assert.equal(res.success, true);
  const saved = ctx.__getSaved();
  assert.equal(saved.isPublished, true, 'lifecycle write must publish');
  assert.equal(saved.publishedAt, '2026-05-31T00:00:00.000Z');
});

test('guard: __allowPublishStateWrite honors unpublish (publishedAt null)', () => {
  const ctx = loadConfigContext({ isPublished: true, publishedAt: '2020-01-01T00:00:00.000Z', spreadsheetId: 'sheet1' });
  const res = ctx.saveUserConfig('u1',
    { isPublished: false, publishedAt: null, spreadsheetId: 'sheet1' },
    { __allowPublishStateWrite: true });
  assert.equal(res.success, true);
  const saved = ctx.__getSaved();
  assert.equal(saved.isPublished, false, 'lifecycle write must unpublish');
  assert.equal(saved.publishedAt, null);
});

test('guard: brand-new config (no persisted) defaults to unpublished on non-sentinel save', () => {
  const ctx = loadConfigContext(null);
  const res = ctx.saveUserConfig('u1', { isPublished: true, spreadsheetId: 'sheet1' });
  assert.equal(res.success, true);
  const saved = ctx.__getSaved();
  assert.equal(saved.isPublished, false, 'new board must start unpublished without a lifecycle write');
  assert.equal(saved.publishedAt, null);
});

test('guard: non-publish fields still save normally on non-sentinel save', () => {
  const ctx = loadConfigContext({ isPublished: true, publishedAt: '2020-01-01T00:00:00.000Z', spreadsheetId: 'sheet1' });
  const res = ctx.saveUserConfig('u1', { isPublished: false, spreadsheetId: 'sheet1', sheetName: 'NewSheet' });
  assert.equal(res.success, true);
  const saved = ctx.__getSaved();
  assert.equal(saved.sheetName, 'NewSheet', 'ordinary fields must persist');
  assert.equal(saved.isPublished, true, 'but publish state is still protected');
});
