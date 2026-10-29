const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function createMockCache(initial = {}) {
  const store = new Map(Object.entries(initial));
  return {
    get: (k) => (store.has(k) ? store.get(k) : null),
    put: (k, v) => { store.set(k, v); },
    remove: (k) => { store.delete(k); },
    _store: store
  };
}

function loadDatabaseContext(overrides = {}) {
  const cache = overrides.cache || createMockCache();
  const propsStore = overrides.propsStore || {};

  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    CacheService: { getScriptCache: () => cache },
    CACHE_DURATION: {
      SHORT: 10, MEDIUM: 30, FORM_DATA: 30, LONG: 300,
      DATABASE_LONG: 600, USER_INDIVIDUAL: 900, EXTRA_LONG: 3600
    },
    TIMEOUT_MS: { DEFAULT: 5000 },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (k) => (k in propsStore ? propsStore[k] : null),
        setProperty: (k, v) => { propsStore[k] = v; },
        getProperties: () => ({ ...propsStore })
      })
    },
    SpreadsheetApp: overrides.SpreadsheetApp || {
      openById: () => { throw new Error('SpreadsheetApp.openById not stubbed'); }
    },
    Session: {
      getActiveUser: () => ({ getEmail: () => overrides.currentEmail || 'actor@example.com' })
    },
    // Real validateEmail returns { isValid, sanitized, errors }; stub matches that shape.
    validateEmail: (email) => ({
      isValid: typeof email === 'string' && /.+@.+/.test(email),
      sanitized: email,
      errors: []
    }),
    getCurrentEmail: overrides.getCurrentEmail || (() => overrides.currentEmail || 'actor@example.com'),
    isAdministrator: overrides.isAdministrator || (() => false),
    getUserConfig: () => ({ success: true, config: {} }),
    executeWithRetry: (fn) => fn(),
    getCachedProperty: (k) => (k in propsStore ? propsStore[k] : null),
    clearPropertyCache: () => {},
    simpleHash: (s) => String(s).length,
    saveToCacheWithSizeCheck: (k, data) => { cache.put(k, JSON.stringify(data)); return true; },
    safeJsonParse_: (text, fallback) => {
      const fb = (fallback === undefined ? null : fallback);
      if (text == null || text === '') return fb;
      if (typeof text === 'object') return text;
      try { return JSON.parse(text); } catch (_) { return fb; }
    },
    logError_: () => {},
    ...overrides
  };

  vm.createContext(context);
  const source = fs.readFileSync(path.resolve(__dirname, '../src/DatabaseCore.js'), 'utf8');
  vm.runInContext(source, context, { filename: 'DatabaseCore.js' });
  context._cache = cache;
  context._propsStore = propsStore;
  return context;
}

// 先生として実行する配置: スプレッドシートは SpreadsheetApp.openById で直接開く。
//   開かせてよいか (checkSpreadsheetAccess_) だけをここで確かめる。
const fakeSs = (id) => ({ _id: id, getId: () => id, getSheetByName: (n) => ({ name: n, ssId: id }) });
const openByIdRecorder = () => {
  const opened = [];
  return { opened, SpreadsheetApp: { openById: (id) => { opened.push(id); return fakeSs(id); } } };
};
// ボードの持ち主の検索と公開状態を差し替える (DatabaseCore の関数はグローバルなので後から上書きできる)。
function withBoards(ctx, boards) {
  ctx.findUserBySpreadsheetId = (id) => boards[id] || null;
  ctx.isUserBoardPublished = (user) => Boolean(user && user.published);
  return ctx;
}

// =====================================================================
// openSpreadsheet
// =====================================================================

test('openSpreadsheet: 不正な id は null', () => {
  const ctx = loadDatabaseContext();
  assert.equal(ctx.openSpreadsheet(''), null);
  assert.equal(ctx.openSpreadsheet(null), null);
});

test('openSpreadsheet: DB は誰の呼び出しでも直接開く', () => {
  const { opened, SpreadsheetApp } = openByIdRecorder();
  const ctx = loadDatabaseContext({ SpreadsheetApp, propsStore: { DATABASE_SPREADSHEET_ID: 'db-id' } });
  const access = ctx.openSpreadsheet('db-id');
  assert.ok(access);
  assert.equal(access.spreadsheet.getId(), 'db-id');
  assert.deepEqual(opened, ['db-id']);
});

test('openSpreadsheet: 先生 (管理者) はどのボードにも入っていないシートも開ける (作ったばかりのフォームの回答先)', () => {
  const { opened, SpreadsheetApp } = openByIdRecorder();
  const ctx = withBoards(loadDatabaseContext({ SpreadsheetApp, isAdministrator: () => true, currentEmail: 'teacher@example.com' }), {});
  assert.ok(ctx.openSpreadsheet('new-form-ss'));
  assert.deepEqual(opened, ['new-form-ss']);
});

test('openSpreadsheet: 児童は公開中のボードのシートだけ開ける', () => {
  const { opened, SpreadsheetApp } = openByIdRecorder();
  const ctx = withBoards(loadDatabaseContext({ SpreadsheetApp, currentEmail: 'child@example.com' }), {
    'board-ss': { userId: 'u1', published: true },
    'closed-ss': { userId: 'u1', published: false }
  });
  assert.ok(ctx.openSpreadsheet('board-ss'));
  assert.equal(ctx.openSpreadsheet('closed-ss'), null, '非公開のボード');
  assert.equal(ctx.openSpreadsheet('other-ss'), null, 'どのボードにも入っていない');
  assert.deepEqual(opened, ['board-ss'], '断ったシートは開かない');
});

test('openSpreadsheet: openById が失敗したら null', () => {
  const ctx = loadDatabaseContext({
    SpreadsheetApp: { openById: () => { throw new Error('not found'); } },
    propsStore: { DATABASE_SPREADSHEET_ID: 'db-id' }
  });
  assert.equal(ctx.openSpreadsheet('db-id'), null);
});

test('openSpreadsheet: getSheet はシート名で引き、名前が無い・失敗したときは null', () => {
  const ss = { getSheetByName: (n) => { if (n === 'boom') throw new Error('x'); return { name: n }; } };
  const ctx = loadDatabaseContext({ SpreadsheetApp: { openById: () => ss }, propsStore: { DATABASE_SPREADSHEET_ID: 'db-id' } });
  const access = ctx.openSpreadsheet('db-id');
  assert.equal(access.getSheet('users').name, 'users');
  assert.equal(access.getSheet(''), null);
  assert.equal(access.getSheet('boom'), null);
});

// =====================================================================
// checkSpreadsheetAccess_ の cache と、公開状態が変わったときの古さの解消
// =====================================================================

test('checkSpreadsheetAccess_: 許可は覚えるが、公開を終えたら invalidateAccessDecisionCache_ で古くなる', () => {
  const boards = { 'board-ss': { userId: 'u1', published: true } };
  const ctx = withBoards(loadDatabaseContext({ currentEmail: 'child@example.com' }), boards);
  assert.equal(ctx.checkSpreadsheetAccess_('board-ss').allowed, true);
  boards['board-ss'].published = false;
  assert.equal(ctx.checkSpreadsheetAccess_('board-ss').allowed, true, '60 秒は覚えている');
  ctx.invalidateAccessDecisionCache_('board-ss');
  assert.equal(ctx.checkSpreadsheetAccess_('board-ss').allowed, false, '版を上げたら判定し直す');
});

// =====================================================================
// openDatabase
// =====================================================================

test('openDatabase: DATABASE_SPREADSHEET_ID が無ければ null', () => {
  const ctx = loadDatabaseContext();
  assert.equal(ctx.openDatabase(), null);
});

test('openDatabase: 先生のスプレッドシートを直接開く / 開けなければ null', () => {
  const { opened, SpreadsheetApp } = openByIdRecorder();
  const ctx = loadDatabaseContext({ SpreadsheetApp, propsStore: { DATABASE_SPREADSHEET_ID: 'db-id' } });
  assert.equal(ctx.openDatabase().getId(), 'db-id');
  assert.deepEqual(opened, ['db-id']);
  const broken = loadDatabaseContext({
    SpreadsheetApp: { openById: () => { throw new Error('gone'); } },
    propsStore: { DATABASE_SPREADSHEET_ID: 'db-id' }
  });
  assert.equal(broken.openDatabase(), null);
});
