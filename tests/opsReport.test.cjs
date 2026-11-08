/**
 * ops-monitor への報告 (src/OpsReport.js) のテスト。
 *
 * Why: logError_ の context やエラーの文面には閲覧者のメールや回答が入ることがある。
 *   外へ出すのは「処理名・Error の name・ファイルと行」だけであることを固定する。
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const read = (f) => fs.readFileSync(path.resolve(__dirname, '../src', f), 'utf8');

function load({ url = 'https://example.test/report', token = 'ops_test' } = {}) {
  const sent = [];
  const store = {};
  const context = {
    UrlFetchApp: { fetch: (u, opt) => { sent.push({ url: u, opt }); } },
    CacheService: { getScriptCache: () => ({ get: (k) => store[k] || null, put: (k, v) => { store[k] = v; } }) },
    Utilities: { base64EncodeWebSafe: (s) => Buffer.from(s).toString('base64url') }
  };
  vm.createContext(context);
  vm.runInContext(`const DEPLOY_OPS_REPORT_URL = ${JSON.stringify(url)};
    const DEPLOY_OPS_REPORT_TOKEN = ${JSON.stringify(token)};`, context);
  vm.runInContext(read('OpsReport.js'), context, { filename: 'OpsReport.js' });
  return { context, sent };
}

test('送るのは where / kind / file / line だけで、文面は送らない', () => {
  const { context, sent } = load();
  const err = new TypeError('taro@example.com の回答が読めない');
  err.stack = 'TypeError: taro@example.com の回答が読めない\n    at getData_ (DataApis:120:7)\n    at doGet (main:10:3)';
  context.opsReport_('getData_', err);
  assert.equal(sent.length, 1);
  const body = JSON.parse(sent[0].opt.payload);
  assert.deepEqual(body, { where: 'getData_', kind: 'TypeError', file: 'DataApis', line: 120 });
  assert.ok(!sent[0].opt.payload.includes('example.com'));
  assert.equal(sent[0].opt.headers.Authorization, 'Bearer ops_test');
});

test('同じ場所は間引く', () => {
  const { context, sent } = load();
  const err = new Error('x');
  err.stack = 'Error: x\n    at f (main:1:1)';
  context.opsReport_('f', err);
  context.opsReport_('f', err);
  assert.equal(sent.length, 1);
});

test('鍵が空なら送らない (リポジトリの既定)', () => {
  const { context, sent } = load({ token: '' });
  context.opsReport_('f', new Error('x'));
  assert.equal(sent.length, 0);
});

test('受け口が受け付けない処理名は unknown にし、送信の失敗で投げない', () => {
  const { context, sent } = load();
  context.opsReport_('名前 にスペース', 'ただの文字列');
  assert.equal(JSON.parse(sent[0].opt.payload).where, 'unknown');
  context.UrlFetchApp.fetch = () => { throw new Error('network'); };
  assert.doesNotThrow(() => context.opsReport_('g', new Error('y')));
});
