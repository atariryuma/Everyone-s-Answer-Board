/**
 * 先生専用の配置 (isSingleTenant_) のテスト。
 *
 * Why: 2026-10-29、児童は「開いた人として実行」の Apps Script を開けなくなった (管理側の設定)。
 *   先生ごとに「先生として実行」で配置し、SA pool を使わず先生の権限で直接開く。
 *   ここで守るのは 2 つ: (1) 開く経路だけが変わり、「誰が何を見てよいか」は変わらない、
 *   (2) 管理者と DB は配置時に焼き込んだ値だけで決まり、最初に開いた人では決まらない。
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const read = (f) => fs.readFileSync(path.resolve(__dirname, '../src', f), 'utf8');

// ---------------------------------------------------------------------
// openSpreadsheet: 開く経路
// ---------------------------------------------------------------------

function loadDatabase({ singleTenant, currentEmail, published }) {
  const opened = [];
  const ss = { id: 'board', getSheetByName: () => null };
  const props = { DATABASE_SPREADSHEET_ID: 'db-id' };   // SERVICE_ACCOUNT_CREDS は無い
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    CacheService: { getScriptCache: () => ({ get: () => null, put: () => {}, remove: () => {} }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: (k) => props[k] || null, setProperty: () => {} }) },
    SpreadsheetApp: { openById: (id) => { opened.push(id); return ss; } },
    UrlFetchApp: { fetch: () => { throw new Error('先生専用では Sheets API (SA) を呼ばない'); } },
    Utilities: { sleep: () => {} },
    getCurrentEmail: () => currentEmail,
    isAdministrator: () => false,
    getCachedProperty: (k) => props[k] || null,
    logError_: () => {},
    safeJsonParse_: (t, fb) => { try { return JSON.parse(t); } catch (_) { return fb; } }
  };
  vm.createContext(context);
  vm.runInContext(read('DatabaseCore.js'), context, { filename: 'DatabaseCore.js' });
  // source の関数宣言で上書きされないよう、読み込み後に差し替える。
  context.isSingleTenant_ = () => singleTenant;
  context.findUserBySpreadsheetId = () => ({
    userId: 'u1', userEmail: 'teacher@example.com', configJson: JSON.stringify({ isPublished: published })
  });
  return { context, opened, ss };
}

test('先生専用: 公開中のボードを児童が読むときも、SA ではなく openById で直接開く', () => {
  const { context, opened, ss } = loadDatabase({ singleTenant: true, currentEmail: 'child@example.com', published: true });
  const access = context.openSpreadsheet('board-id', { useServiceAccount: true });
  assert.ok(access, '開けるはず');
  assert.equal(access.spreadsheet, ss);
  assert.deepEqual(opened, ['board-id']);
});

test('先生専用: 非公開のボードは児童には今までどおり開けない (判定は変わらない)', () => {
  const { context, opened } = loadDatabase({ singleTenant: true, currentEmail: 'child@example.com', published: false });
  assert.equal(context.openSpreadsheet('board-id', { useServiceAccount: true }), null);
  assert.deepEqual(opened, [], '判定で断られたら openById すら呼ばない');
});

test('先生専用: DB も直接開く', () => {
  const { context, opened } = loadDatabase({ singleTenant: true, currentEmail: 'child@example.com', published: true });
  assert.ok(context.openSpreadsheet('db-id', { useServiceAccount: true }));
  assert.deepEqual(opened, ['db-id']);
});

test('本体 (みんなで 1 つ): 児童の読みは今までどおり SA 経路 (SA が無ければ開けない)', () => {
  const { context, opened } = loadDatabase({ singleTenant: false, currentEmail: 'child@example.com', published: true });
  assert.equal(context.openSpreadsheet('board-id', { useServiceAccount: true }), null);
  assert.deepEqual(opened, [], '児童の読みで deploy 主の openById に落ちてはいけない');
});

// ---------------------------------------------------------------------
// ensureSingleTenantBootstrap_: 初回の準備
// ---------------------------------------------------------------------

function loadHelpers({ tenancy, props = {} }) {
  const ensured = [];
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    PROPERTY_CACHE_TTL: 0,
    PropertiesService: { getScriptProperties: () => ({
      getProperty: (k) => (k in props ? props[k] : null),
      setProperty: (k, v) => { props[k] = v; }
    }) },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
    SpreadsheetApp: { openById: (id) => ({ id }) },
    __ensureSheetWithHeaders_: (ss, name) => { ensured.push(ss.id + ':' + name); },
    USERS_SHEET_HEADERS: [], LESSONS_SHEET_HEADERS: [], LESSON_RESPONSES_SHEET_HEADERS: []
  };
  vm.createContext(context);
  // Tenancy.js の代わり (配置時に差し替えられる値)。
  vm.runInContext(`const DEPLOY_TENANCY = ${JSON.stringify(tenancy)};
    const DEPLOY_ADMIN_EMAIL = 'Teacher@Example.com';
    const DEPLOY_DATABASE_SPREADSHEET_ID = 'teacher-ss';`, context);
  vm.runInContext(read('helpers.js'), context, { filename: 'helpers.js' });
  return { context, props, ensured };
}

test('初回の準備: 焼き込んだ管理者と DB でプロパティとシートをそろえる', () => {
  const { context, props, ensured } = loadHelpers({ tenancy: 'single' });
  context.ensureSingleTenantBootstrap_();
  assert.equal(props.ADMIN_EMAIL, 'teacher@example.com', '小文字にそろえる');
  assert.equal(props.DATABASE_SPREADSHEET_ID, 'teacher-ss');
  assert.equal(props.TENANCY, 'single');
  assert.deepEqual(ensured, ['teacher-ss:users', 'teacher-ss:lessons', 'teacher-ss:lesson_responses']);
  assert.equal(context.isSingleTenant_(), true);
});

test('初回の準備: 既にあるプロパティは上書きしない (先生が後から変えた値を守る)', () => {
  const { context, props, ensured } = loadHelpers({
    tenancy: 'single', props: { ADMIN_EMAIL: 'other@example.com', DATABASE_SPREADSHEET_ID: 'other-ss' }
  });
  context.ensureSingleTenantBootstrap_();
  assert.equal(props.ADMIN_EMAIL, 'other@example.com');
  assert.equal(props.DATABASE_SPREADSHEET_ID, 'other-ss');
  assert.deepEqual(ensured, []);
});

test('本体 (DEPLOY_TENANCY が空): 何もしない', () => {
  const { context, props, ensured } = loadHelpers({ tenancy: '' });
  context.ensureSingleTenantBootstrap_();
  assert.deepEqual(props, {});
  assert.deepEqual(ensured, []);
  assert.equal(context.isSingleTenant_(), false);
});
