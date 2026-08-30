/**
 * canAccessTargetUser — 「第三者がこのユーザーのボードを読めるか」の唯一の関門。
 *
 * Why ここを pin するか: findUserById → applyUserAccessControl → canAccessTargetUser が
 *   viewer の cross-user read (getPublishedSheetData / getNotificationUpdate /
 *   リアクション) すべての根にある。ここの結論が doGet の結論とズレると、
 *   「画面では止まっているのに API 経由では読める」状態になる。
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const DB_SOURCE = fs.readFileSync(path.resolve(__dirname, '../src/DatabaseCore.js'), 'utf8');
const DB_SCRIPT = new vm.Script(DB_SOURCE, { filename: 'DatabaseCore.js' });

function loadCtx(overrides = {}) {
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    logError_: () => {},
    CacheService: { getScriptCache: () => ({ get: () => null, put: () => {}, remove: () => {} }) },
    CACHE_DURATION: { SHORT: 10, MEDIUM: 30, LONG: 300, DATABASE_LONG: 600, USER_INDIVIDUAL: 900 },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null, setProperty: () => {} }) },
    SpreadsheetApp: { openById: () => { throw new Error('not stubbed'); } },
    UrlFetchApp: { fetch: () => { throw new Error('not stubbed'); } },
    Utilities: { sleep: () => {} },
    Session: { getActiveUser: () => ({ getEmail: () => 'viewer@example.com' }) },
    getCurrentEmail: () => 'viewer@example.com',
    isAdministrator: (e) => e === 'admin@example.com',
    getCachedProperty: () => null,
    getUserConfig: () => ({ success: true, config: {} }),
    executeWithRetry: (fn) => fn(),
    validateEmail: (e) => ({ isValid: true, sanitized: e, errors: [] }),
    safeJsonParse_: (t, fb) => { try { return JSON.parse(t); } catch (_) { return fb === undefined ? null : fb; } },
    // DataApis.js 側の実装 (GAS は単一グローバルスコープなので実行時は解決される)
    sameEmail_: (a, b) => String(a || '').toLowerCase().trim() === String(b || '').toLowerCase().trim()
  };
  Object.assign(context, overrides);
  vm.createContext(context);
  DB_SCRIPT.runInContext(context);
  return context;
}

// 公開/非公開 + 有効/無効の組み合わせで target user を作る。
function target({ published = true, isActive = true, email = 'owner@example.com' } = {}) {
  return {
    userId: 'u1',
    userEmail: email,
    isActive,
    configJson: JSON.stringify({ isPublished: published })
  };
}

const viewerCtx = { requestingUser: 'viewer@example.com', allowPublishedRead: true };

test('canAccessTargetUser: 公開中のボードは第三者も読める', () => {
  const ctx = loadCtx();
  assert.equal(ctx.canAccessTargetUser(target({ published: true }), viewerCtx), true);
});

test('canAccessTargetUser: 非公開のボードは第三者から読めない', () => {
  const ctx = loadCtx();
  assert.equal(ctx.canAccessTargetUser(target({ published: false }), viewerCtx), false);
});

test('canAccessTargetUser: 無効化されたユーザーのボードは公開中でも第三者から読めない', () => {
  const ctx = loadCtx();
  // doGet (main.js) は isActive=false を「公開の強制終了」として扱う。
  //   データ層がこれを見ていないと、画面では止まるのに API 経由では読めてしまう。
  assert.equal(ctx.canAccessTargetUser(target({ published: true, isActive: false }), viewerCtx), false);
});

test('canAccessTargetUser: 無効化されていても本人は読める', () => {
  const ctx = loadCtx();
  assert.equal(
    ctx.canAccessTargetUser(target({ published: true, isActive: false }),
      { requestingUser: 'owner@example.com', allowPublishedRead: true }),
    true
  );
});

test('canAccessTargetUser: 無効化されていても管理者は読める', () => {
  const ctx = loadCtx();
  assert.equal(
    ctx.canAccessTargetUser(target({ published: true, isActive: false }),
      { requestingUser: 'admin@example.com', allowPublishedRead: true }),
    true
  );
});

test('canAccessTargetUser: isActive 未設定の既存行は巻き込まない', () => {
  const ctx = loadCtx();
  const legacy = { userId: 'u1', userEmail: 'owner@example.com', configJson: JSON.stringify({ isPublished: true }) };
  assert.equal(ctx.canAccessTargetUser(legacy, viewerCtx), true);
});

test('canAccessTargetUser: allowPublishedRead なしなら第三者は読めない (公開中でも)', () => {
  const ctx = loadCtx();
  assert.equal(
    ctx.canAccessTargetUser(target({ published: true }), { requestingUser: 'viewer@example.com' }),
    false
  );
});

test('canAccessTargetUser: 本人はメールの大文字小文字が違っても読める', () => {
  const ctx = loadCtx();
  assert.equal(
    ctx.canAccessTargetUser(target({ published: false, email: 'Owner@Example.com' }),
      { requestingUser: 'owner@example.com', allowPublishedRead: true }),
    true
  );
});
