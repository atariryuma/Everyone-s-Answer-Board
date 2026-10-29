/**
 * updateUser_: 内容が同じ書き込みは DB に書かず USER_CACHE_VERSION も上げない (2026-10 監査)。
 *
 * Why: 書き込みのたびに clearDatabaseUserCache が version を上げ、児童 30 人分の user cache が
 *   一斉に落ちて次の poll が users シートを読み直していた。lastAccessedAt / etag しか違わない
 *   保存 (同じ内容の再保存、同じ patch の再適用) で version を上げる理由は無い。
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { gasResponseStubs } = require('./_helpers.cjs');

const DB_SOURCE = fs.readFileSync(path.resolve(__dirname, '../src/DatabaseCore.js'), 'utf8');
const HEADERS = ['userId', 'userEmail', 'googleId', 'isActive', 'configJson', 'lastModified', 'createdAt'];

function makeUsersSheet(rows) {
  const data = [HEADERS.slice(), ...rows.map((r) => r.slice())];
  const calls = { setValues: [] };
  return {
    calls,
    _data: data,
    getName: () => 'users',
    getLastRow: () => data.length,
    getLastColumn: () => HEADERS.length,
    getDataRange: () => ({ getValues: () => data.map((r) => r.slice()) }),
    getRange: (row, col, numRows, numCols) => ({
      getValues: () => [(data[row - 1] || []).slice(col - 1, col - 1 + (numCols || 1))],
      setValues: (vs) => {
        calls.setValues.push({ row, col, values: vs });
        for (let j = 0; j < vs[0].length; j++) data[row - 1][col - 1 + j] = vs[0][j];
      }
    })
  };
}

function loadCtx(sheet) {
  const store = new Map();
  const props = {};
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    ...gasResponseStubs(),
    CacheService: { getScriptCache: () => ({
      get: (k) => (store.has(k) ? store.get(k) : null), put: (k, v) => { store.set(k, v); }, remove: (k) => { store.delete(k); }
    }) },
    CACHE_DURATION: { SHORT: 10, MEDIUM: 30, LONG: 300, DATABASE_LONG: 600, USER_INDIVIDUAL: 900 },
    PropertiesService: { getScriptProperties: () => ({
      getProperty: (k) => (k in props ? props[k] : null), setProperty: (k, v) => { props[k] = v; }
    }) },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
    SpreadsheetApp: { openById: () => { throw new Error('not stubbed'); } },
    UrlFetchApp: { fetch: () => { throw new Error('not stubbed'); } },
    Utilities: { sleep: () => {} },
    getCurrentEmail: () => 'owner@example.com',
    isAdministrator: () => false,
    getCachedProperty_: (k) => (k === 'USER_CACHE_VERSION' ? (props.USER_CACHE_VERSION || '0') : null),
    clearPropertyCache_: () => {},
    getUserConfig: () => ({ success: true, config: {} }),
    executeWithRetry: (fn) => fn(),
    validateEmail: (e) => ({ isValid: /.+@.+/.test(String(e || '')), sanitized: e, errors: [] }),
    safeJsonParse_: (t, fb) => { try { return JSON.parse(t); } catch (_) { return fb === undefined ? null : fb; } },
    sameEmail_: (a, b) => String(a || '').toLowerCase().trim() === String(b || '').toLowerCase().trim(),
    simpleHash: (o) => Object.keys(o || {}).sort().map((k) => `${k}:${o[k]}`).join('|'),
    saveToCacheWithSizeCheck: (k, v) => { store.set(k, JSON.stringify(v)); return true; },
    acquireCacheLock_: () => 'tok',
    releaseCacheLock_: () => {}
  };
  vm.createContext(context);
  vm.runInContext(DB_SOURCE, context, { filename: 'DatabaseCore.js' });
  context.openDatabase = () => ({ getSheetByName: (n) => (n === 'users' ? sheet : null) });
  return { ctx: context, props };
}

const PERSISTED = { spreadsheetId: 'ss1', isPublished: true, displaySettings: { showNames: false, boardMode: 'board' },
  profiles: [{ id: 'p1', spreadsheetId: 'ss1' }], lastAccessedAt: '2026-10-01T00:00:00.000Z', etag: 'old_etag' };
const row = () => ['u1', 'owner@example.com', '', true, JSON.stringify(PERSISTED), '2026-10-01T00:00:00.000Z', ''];

test('updateUser: lastAccessedAt / etag だけが違う configJson は書かず、version も上げない', () => {
  const sheet = makeUsersSheet([row()]);
  const { ctx, props } = loadCtx(sheet);
  // key 順も変えて渡す (順序の違いは差分ではない)
  const same = { etag: 'new_etag', lastAccessedAt: '2026-10-16T00:00:00.000Z', profiles: [{ spreadsheetId: 'ss1', id: 'p1' }],
    displaySettings: { boardMode: 'board', showNames: false }, isPublished: true, spreadsheetId: 'ss1' };
  const res = ctx.updateUser_('u1', { configJson: JSON.stringify(same) }, { requestingUser: 'owner@example.com' });
  assert.equal(res.success, true);
  assert.equal(res.unchanged, true);
  assert.equal(res.currentEtag, 'old_etag', '永続済の etag を返す (client が次の保存で使う)');
  assert.equal(res.currentConfig.etag, 'old_etag');
  assert.equal(sheet.calls.setValues.length, 0, 'シートに書かない');
  assert.equal(props.USER_CACHE_VERSION, undefined, 'USER_CACHE_VERSION を上げない');
});

test('updateUser: 内容が違えば従来どおり書いて version を上げる', () => {
  const sheet = makeUsersSheet([row()]);
  const { ctx, props } = loadCtx(sheet);
  const changed = { ...PERSISTED, displaySettings: { showNames: true, boardMode: 'board' }, etag: 'new_etag' };
  const res = ctx.updateUser_('u1', { configJson: JSON.stringify(changed) }, { requestingUser: 'owner@example.com' });
  assert.equal(res.success, true);
  assert.ok(!res.unchanged);
  assert.equal(sheet.calls.setValues.length, 1);
  assert.equal(props.USER_CACHE_VERSION, '1');
});

test('updateUser: 配列の中身の違い (profile 追加) は差分として書く', () => {
  const sheet = makeUsersSheet([row()]);
  const { ctx } = loadCtx(sheet);
  const changed = { ...PERSISTED, profiles: [{ id: 'p1', spreadsheetId: 'ss1' }, { id: 'p2', spreadsheetId: 'ss2' }] };
  const res = ctx.updateUser_('u1', { configJson: JSON.stringify(changed) }, { requestingUser: 'owner@example.com' });
  assert.equal(res.success, true);
  assert.equal(sheet.calls.setValues.length, 1);
});

test('updateUser: configJson 以外の列 (isActive) も同じ値なら書かない / 違えば書く', () => {
  const sheet = makeUsersSheet([row()]);
  const { ctx } = loadCtx(sheet);
  assert.equal(ctx.updateUser_('u1', { isActive: true }, { requestingUser: 'owner@example.com' }).unchanged, true);
  assert.equal(sheet.calls.setValues.length, 0);
  const res = ctx.updateUser_('u1', { isActive: false }, { requestingUser: 'owner@example.com' });
  assert.equal(res.success, true);
  assert.ok(!res.unchanged);
  assert.equal(sheet.calls.setValues.length, 1);
});

test('updateUser: etag の CAS (expectedEtag) は no-op 判定より先に効く', () => {
  const sheet = makeUsersSheet([row()]);
  const { ctx } = loadCtx(sheet);
  const res = ctx.updateUser_('u1', { configJson: JSON.stringify(PERSISTED) },
    { requestingUser: 'owner@example.com', expectedEtag: 'stale_etag' });
  assert.equal(res.success, false);
  assert.equal(res.error, 'etag_mismatch');
});
