/**
 * users シートの read を増やさないための約束 (2026-09-08 の 429 storm 対策)。
 *
 * Why: users に居ない児童の findUserByEmail が毎回 users シートを直接読み、
 *   validateServiceAccountUsage の skipCache が users 一覧の cache まで飛ばし、
 *   updateUser が同じ行を 2 回読んで 429 で落ちていた。どれも「読まなくてよい read」。
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { gasResponseStubs } = require('./_helpers.cjs');

const DB_SOURCE = fs.readFileSync(path.resolve(__dirname, '../src/DatabaseCore.js'), 'utf8');
const DB_SCRIPT = new vm.Script(DB_SOURCE, { filename: 'DatabaseCore.js' });

const HEADERS = ['userId', 'userEmail', 'googleId', 'isActive', 'configJson', 'lastModified', 'createdAt'];

function makeUsersSheet(rows) {
  const data = [HEADERS.slice(), ...rows.map((r) => r.slice())];
  const calls = { getDataRange: 0, getRangeValues: 0, getLastRow: 0, setValues: [] };
  return {
    calls,
    _data: data,
    getName: () => 'users',
    getLastRow: () => { calls.getLastRow++; return data.length; },
    getLastColumn: () => HEADERS.length,
    getDataRange: () => ({ getValues: () => { calls.getDataRange++; return data.map((r) => r.slice()); } }),
    getRange: (row, col, numRows, numCols) => ({
      getValues: () => {
        calls.getRangeValues++;
        const out = [];
        for (let i = 0; i < (numRows || 1); i++) out.push((data[row - 1 + i] || []).slice(col - 1, col - 1 + (numCols || 1)));
        return out;
      },
      setValues: (vs) => {
        calls.setValues.push({ row, col, values: vs });
        for (let j = 0; j < vs[0].length; j++) {
          if (!data[row - 1]) data[row - 1] = [];
          data[row - 1][col - 1 + j] = vs[0][j];
        }
      }
    })
  };
}

function loadCtx(sheet, overrides = {}) {
  const store = new Map();
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    ...gasResponseStubs(),
    CacheService: {
      getScriptCache: () => ({
        get: (k) => store.has(k) ? store.get(k) : null,
        put: (k, v) => { store.set(k, v); },
        remove: (k) => { store.delete(k); }
      })
    },
    CACHE_DURATION: { SHORT: 10, MEDIUM: 30, LONG: 300, DATABASE_LONG: 600, USER_INDIVIDUAL: 900 },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null, setProperty: () => {} }) },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
    SpreadsheetApp: { openById: () => { throw new Error('not stubbed'); } },
    UrlFetchApp: { fetch: () => { throw new Error('not stubbed'); } },
    Utilities: { sleep: () => {} },
    getCurrentEmail: () => 'child@example.com',
    isAdministrator: (e) => e === 'admin@example.com',
    getCachedProperty: (k) => (k === 'USER_CACHE_VERSION' ? '0' : null),
    clearPropertyCache: () => {},
    getUserConfig: () => ({ success: true, config: {} }),
    executeWithRetry: (fn) => fn(),
    validateEmail: (e) => ({ isValid: /.+@.+/.test(String(e || '')), sanitized: e, errors: [] }),
    safeJsonParse_: (t, fb) => { try { return JSON.parse(t); } catch (_) { return fb === undefined ? null : fb; } },
    sameEmail_: (a, b) => String(a || '').toLowerCase().trim() === String(b || '').toLowerCase().trim(),
    simpleHash: (o) => Object.keys(o || {}).sort().map((k) => `${k}:${o[k]}`).join('|'),
    saveToCacheWithSizeCheck: (k, v) => { store.set(k, JSON.stringify(v)); return true; }
  };
  Object.assign(context, overrides);
  vm.createContext(context);
  DB_SCRIPT.runInContext(context);
  // openDatabase は DatabaseCore.js 自身が定義するので、script 実行後に差し替える。
  context.openDatabase = () => ({ getSheetByName: (n) => (n === 'users' ? sheet : null) });
  return { ctx: context, store };
}

const OWNER_ROW = ['u1', 'owner@example.com', '', true, JSON.stringify({ isPublished: true, spreadsheetId: 'ss1' }), '', ''];

test('findUserByEmail: users 一覧 (cache) に居ない児童は、users シートを直接読み直さずに null', () => {
  const sheet = makeUsersSheet([OWNER_ROW]);
  const { ctx } = loadCtx(sheet);
  // 1 回目: 一覧を読んで cache に載せる
  assert.equal(ctx.findUserByEmail('child@example.com', { requestingUser: 'child@example.com' }), null);
  const after = sheet.calls.getDataRange;
  assert.equal(after, 1, '一覧の読みは 1 回');
  assert.equal(sheet.calls.getLastRow, 0, '寸法 + TextFinder の直接読みをしない');
  assert.equal(sheet.calls.getRangeValues, 0);
  // 2 回目以降は cache から (何も読まない)
  assert.equal(ctx.findUserByEmail('child@example.com', { requestingUser: 'child@example.com' }), null);
  assert.equal(sheet.calls.getDataRange, after);
});

test('findUserBySpreadsheetId(skipCache): SS→user の対応は読み直すが、users 一覧は cache を使う', () => {
  const sheet = makeUsersSheet([OWNER_ROW]);
  const { ctx } = loadCtx(sheet);
  assert.equal(ctx.findUserBySpreadsheetId('ss1', { skipCache: true }).userId, 'u1');
  assert.equal(sheet.calls.getDataRange, 1);
  // validateServiceAccountUsage が児童ごと・SS ごとに 60 秒おきに呼ぶ経路。毎回 users を読まない。
  assert.equal(ctx.findUserBySpreadsheetId('ss1', { skipCache: true }).userId, 'u1');
  assert.equal(ctx.findUserBySpreadsheetId('ss1', { skipCache: true }).userId, 'u1');
  assert.equal(sheet.calls.getDataRange, 1, 'users 一覧は cache から');
});

test('updateUser: 同じ行を 2 回読まず、既読の行を下敷きに書く (429 で [] が返っても落ちない)', () => {
  // values API は末尾の空セルを省くので、lastModified / createdAt が無い短い行にする。
  const short = ['u1', 'owner@example.com', '', true, JSON.stringify({ isPublished: true })];
  const sheet = makeUsersSheet([short]);
  const { ctx } = loadCtx(sheet, { getCurrentEmail: () => 'owner@example.com' });
  const res = ctx.updateUser('u1', { configJson: JSON.stringify({ isPublished: false }) }, { requestingUser: 'owner@example.com' });
  assert.equal(res.success, true, res.message);
  assert.equal(sheet.calls.getRangeValues, 0, '行の再読みをしない');
  assert.equal(sheet.calls.setValues.length, 1);
  const w = sheet.calls.setValues[0];
  assert.equal(w.col, 5, 'configJson 列から');
  assert.equal(w.values[0].length, 2, 'configJson + lastModified の 2 列');
  assert.equal(JSON.parse(w.values[0][0]).isPublished, false);
  assert.match(w.values[0][1], /^\d{4}-\d{2}-\d{2}T/, 'lastModified は ISO 時刻');
});

test('updateUser: 更新しない列が範囲に挟まるときも既読の値を保つ', () => {
  const row = ['u1', 'owner@example.com', 'gid', true, JSON.stringify({ a: 1 }), 'old', 'c'];
  const sheet = makeUsersSheet([row]);
  const { ctx } = loadCtx(sheet, { getCurrentEmail: () => 'owner@example.com' });
  // userEmail (2) と configJson (5) を更新 → 範囲は 2..6、googleId / isActive は既存値のまま
  const res = ctx.updateUser('u1', { userEmail: 'new@example.com', configJson: '{"a":2}' }, { requestingUser: 'owner@example.com' });
  assert.equal(res.success, true, res.message);
  const w = sheet.calls.setValues[0];
  assert.equal(w.col, 2);
  assert.equal(JSON.stringify(w.values[0].slice(0, 4)), JSON.stringify(['new@example.com', 'gid', true, '{"a":2}']));
  assert.equal(sheet.calls.getRangeValues, 0);
});
