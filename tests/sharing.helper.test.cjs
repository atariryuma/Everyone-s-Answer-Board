const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

// v2782+ refactor: domain-wide sharing は廃止。 sharing helper は SA pool 全員を editor に
// 追加する単一の責務に統一された。 旧 `setupDomainWideSharing` のテストは削除。

function loadSharingHelperContext(overrides = {}) {
  const editorAdds = [];
  const ss = overrides.spreadsheet || {
    addEditor: (email) => { editorAdds.push(email); return ss; }
  };

  const pool = overrides.pool || [
    { client_email: 'sa1@project.iam.gserviceaccount.com', private_key: 'k1', type: 'service_account' }
  ];

  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    SpreadsheetApp: { openById: () => ss },
    getAllServiceAccounts_: () => pool,
    parseServiceAccountCredsSoft_: () => null,
    invalidateSaCache_: () => {},
    logError_: () => {},
    getCachedProperty: () => null,
    ...overrides
  };

  vm.createContext(context);
  const source = fs.readFileSync(path.resolve(__dirname, '../src/SharingHelper.js'), 'utf8');
  vm.runInContext(source, context, { filename: 'SharingHelper.js' });
  return { context, editorAdds, spreadsheet: ss };
}

// vm.createContext は別 realm を作るため、 context 内で生成された Array は test realm の
// Array.prototype と異なる。 deepStrictEqual はこれを reject する。 比較前に [...arr] で
// test realm の Array に詰め替えて回避する。
const cloneArr = (arr) => [...(arr || [])];

test('addServiceAccountsAsEditors: adds all pool SAs as editors', () => {
  const { context, editorAdds } = loadSharingHelperContext({
    pool: [
      { client_email: 'sa1@example.iam', private_key: 'k', type: 'service_account' },
      { client_email: 'sa2@example.iam', private_key: 'k', type: 'service_account' },
      { client_email: 'sa3@example.iam', private_key: 'k', type: 'service_account' }
    ]
  });

  const result = context.addServiceAccountsAsEditors('ss-1');

  assert.equal(result.success, true);
  assert.deepEqual(cloneArr(result.added).sort(), ['sa1@example.iam', 'sa2@example.iam', 'sa3@example.iam']);
  assert.equal(result.errors.length, 0);
  assert.deepEqual(editorAdds.slice().sort(), ['sa1@example.iam', 'sa2@example.iam', 'sa3@example.iam']);
});

test('addServiceAccountsAsEditors: continues after individual SA failure (fail-soft)', () => {
  const ss = {
    addEditor: (email) => {
      if (email === 'sa-broken@example.iam') throw new Error('not authorized');
      return ss;
    }
  };
  const { context } = loadSharingHelperContext({
    spreadsheet: ss,
    pool: [
      { client_email: 'sa-ok@example.iam', private_key: 'k', type: 'service_account' },
      { client_email: 'sa-broken@example.iam', private_key: 'k', type: 'service_account' },
      { client_email: 'sa-ok2@example.iam', private_key: 'k', type: 'service_account' }
    ]
  });

  const result = context.addServiceAccountsAsEditors('ss-1');

  assert.equal(result.success, true);
  assert.deepEqual(cloneArr(result.added).sort(), ['sa-ok2@example.iam', 'sa-ok@example.iam']);
  assert.equal(result.errors.length, 1);
  assert.match(result.errors[0], /sa-broken@example\.iam: not authorized/);
});

test('addServiceAccountsAsEditors: returns error when pool is empty', () => {
  const { context } = loadSharingHelperContext({ pool: [] });
  const result = context.addServiceAccountsAsEditors('ss-1');
  assert.equal(result.success, false);
  assert.match(result.errors.join(' '), /SA pool not configured/);
});

test('addServiceAccountAsEditor (legacy wrapper): adds only primary SA', () => {
  const { context, editorAdds } = loadSharingHelperContext({
    pool: [
      { client_email: 'primary@example.iam', private_key: 'k', type: 'service_account' },
      { client_email: 'secondary@example.iam', private_key: 'k', type: 'service_account' }
    ]
  });

  const result = context.addServiceAccountAsEditor('ss-1');

  assert.equal(result.success, true);
  assert.equal(result.added, true);
  assert.equal(result.saEmail, 'primary@example.iam');
  assert.deepEqual(editorAdds, ['primary@example.iam']);
});

test('applySpreadsheetSharingDefaults: returns saEmails (no domainShared field)', () => {
  const { context } = loadSharingHelperContext({
    pool: [
      { client_email: 'sa1@example.iam', private_key: 'k', type: 'service_account' },
      { client_email: 'sa2@example.iam', private_key: 'k', type: 'service_account' }
    ]
  });

  const result = context.applySpreadsheetSharingDefaults('ss-1');

  assert.equal(result.saAdded, true);
  assert.deepEqual(cloneArr(result.saEmails).sort(), ['sa1@example.iam', 'sa2@example.iam']);
  assert.equal(result.errors.length, 0);
  // v2782: 旧 domainShared field は削除済み
  assert.equal(result.domainShared, undefined);
});

test('applySpreadsheetSharingDefaults: ownerEmail is accepted for backward compat (ignored)', () => {
  const { context } = loadSharingHelperContext();
  // 旧 API: 第2引数 ownerEmail。 v2782+ は無視される (throw しないこと)。
  const result = context.applySpreadsheetSharingDefaults('ss-1', 'teacher@school.edu');
  assert.equal(result.saAdded, true);
});

test('applySpreadsheetSharingDefaults: surfaces SA add errors', () => {
  const ss = {
    addEditor: () => { throw new Error('Drive 429'); }
  };
  const { context } = loadSharingHelperContext({
    spreadsheet: ss,
    pool: [{ client_email: 'sa1@example.iam', private_key: 'k', type: 'service_account' }]
  });

  const result = context.applySpreadsheetSharingDefaults('ss-1');

  assert.equal(result.saAdded, false);
  assert.equal(result.saEmails.length, 0);
  assert.ok(result.errors.length > 0);
  assert.match(result.errors[0], /SA: sa1@example\.iam: Drive 429/);
});
