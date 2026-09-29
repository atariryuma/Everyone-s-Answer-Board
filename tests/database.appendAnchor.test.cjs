/**
 * SA proxy の :append は「A1 を含むテーブル」の直下・A 列に固定し、書いた位置を応答で確かめる。
 *
 * Why: Sheets API の :append は範囲内で「最後のテーブル」を探し、その直下・その先頭列から書く。
 *   範囲をシート名だけにしていたため、表の外の孤立セル (切替直後の 👍 が空行に書いたもの)
 *   を最後のテーブルと見なし、児童の回答が J 列から積まれてアプリに見えなくなった (2026-09-29)。
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const DB_SOURCE = fs.readFileSync(path.resolve(__dirname, '../src/DatabaseCore.js'), 'utf8');
const DB_SCRIPT = new vm.Script(DB_SOURCE, { filename: 'DatabaseCore.js' });

function loadCtx(fetchImpl) {
  const calls = [];
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    logError_: () => {},
    CacheService: { getScriptCache: () => ({ get: () => null, put: () => {}, remove: () => {} }) },
    CACHE_DURATION: { SHORT: 10, MEDIUM: 30, LONG: 300, DATABASE_LONG: 600, USER_INDIVIDUAL: 900 },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null, setProperty: () => {} }) },
    SpreadsheetApp: { openById: () => { throw new Error('not stubbed'); } },
    UrlFetchApp: { fetch: (url, opts) => { calls.push({ url, opts }); return fetchImpl(url, opts); } },
    Utilities: { sleep: () => {} },
    Session: { getActiveUser: () => ({ getEmail: () => 'x@example.com' }) },
    getCurrentEmail: () => 'x@example.com',
    isAdministrator: () => false,
    getCachedProperty: () => null,
    executeWithRetry: (fn) => fn(),
    safeJsonParse_: (t, fb) => { try { return JSON.parse(t); } catch (_) { return fb === undefined ? null : fb; } },
    sameEmail_: (a, b) => String(a || '').toLowerCase().trim() === String(b || '').toLowerCase().trim()
  };
  vm.createContext(context);
  DB_SCRIPT.runInContext(context);
  return { ctx: context, calls };
}

const ok = (updatedRange) => () => ({
  getResponseCode: () => 200,
  getContentText: () => JSON.stringify({ updates: { updatedRange } })
});

function proxy(ctx) {
  return ctx.createServiceAccountSheetProxy('ss1', 'phase4', 'tok', {}, 'sa@example.com', () => ({ token: 'tok', saEmail: 'sa@example.com' }));
}

test('appendRow: 範囲は「シート名!A1」で、A 列に書けたことを応答で確かめる', () => {
  const { ctx, calls } = loadCtx(ok("'phase4'!A45:H45"));
  proxy(ctx).appendRow(['t', 'a@example.com', '6年4組', 'A', 4, 2, '理由', '']);
  assert.equal(calls.length, 1);
  assert.match(calls[0].url, /\/values\/phase4!A1:append\?/);
});

test('appendRow: 追記位置が A 列でなければ失敗にする (「送りました」を出さない)', () => {
  const { ctx } = loadCtx(ok("'phase4'!J68:Q68"));
  assert.throws(() => proxy(ctx).appendRow(['t', 'a@example.com']), /A 列ではありません/);
});

test('appendRows: 同じ固定と検証。開始行は応答から取る', () => {
  const { ctx, calls } = loadCtx(ok("'lesson_responses'!A5:I6"));
  const res = ctx.createServiceAccountSheetProxy('db', 'lesson_responses', 'tok', {}, 'sa@example.com', () => ({ token: 'tok', saEmail: 'sa@example.com' }))
    .appendRows([[1], [2]]);
  assert.match(calls[0].url, /\/values\/lesson_responses!A1:append\?/);
  assert.equal(res.startRow, 5);
  assert.equal(res.rowCount, 2);
});

test('appendRows: A 列以外に書かれたら失敗にする (アーカイブのポインタを壊さない)', () => {
  const { ctx } = loadCtx(ok("'lesson_responses'!C9:K10"));
  assert.throws(() => proxy(ctx).appendRows([[1]]), /A 列ではありません/);
});
