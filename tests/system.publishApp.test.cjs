const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function loadSystemControllerContext(overrides = {}) {
  const context = {
    console: {
      log: () => {},
      warn: () => {},
      error: () => {}
    },
    getCurrentEmail: () => 'teacher@example.com',
    findUserByEmail: () => ({ userId: 'user-1', userEmail: 'teacher@example.com' }),
    getUserConfig: () => ({
      success: true,
      config: {
        etag: 'etag-1',
        displaySettings: { showNames: false, showReactions: false, theme: 'default', pageSize: 20 },
        formUrl: 'https://docs.google.com/forms/d/form-id/viewform',
        formTitle: 'Current Form'
      }
    }),
    saveUserConfig: () => ({ success: true, etag: 'etag-2', config: {} }),
    sanitizeDisplaySettings: (settings) => ({
      showNames: Boolean(settings.showNames),
      showReactions: Boolean(settings.showReactions),
      theme: String(settings.theme || 'default'),
      pageSize: Number(settings.pageSize || 20)
    }),
    sanitizeMapping: (mapping) => {
      const sanitized = {};
      ['answer', 'reason', 'class', 'name', 'timestamp', 'email'].forEach((key) => {
        const value = mapping[key];
        if (typeof value === 'number' && value >= 0) sanitized[key] = value;
      });
      return sanitized;
    }
  };

  Object.assign(context, overrides);
  vm.createContext(context);
  const source = fs.readFileSync(path.resolve(__dirname, '../src/SystemController.js'), 'utf8');
  vm.runInContext(source, context, { filename: 'SystemController.js' });
  return context;
}

test('publishApp: non-object config is rejected', () => {
  const context = loadSystemControllerContext();
  const result = context.publishApp(null);
  assert.equal(result.success, false);
  assert.equal(result.message, '公開設定が必要です');
});

test('publishApp: requires etag when current config has etag', () => {
  let saveCalled = false;
  const context = loadSystemControllerContext({
    saveUserConfig: () => {
      saveCalled = true;
      return { success: true };
    }
  });

  const result = context.publishApp({
    spreadsheetId: 'spreadsheet-id',
    sheetName: 'フォームの回答 1',
    columnMapping: { answer: 2 }
  });

  assert.equal(result.success, false);
  assert.equal(result.error, 'etag_required');
  assert.equal(saveCalled, false);
});

test('publishApp: allowlist fields only and save with publish flags', () => {
  let savedPayload = null;

  const context = loadSystemControllerContext({
    saveUserConfig: (_userId, config) => {
      savedPayload = config;
      return {
        success: true,
        etag: 'etag-2',
        config
      };
    }
  });

  const result = context.publishApp({
    spreadsheetId: 'spreadsheet-id',
    sheetName: 'フォームの回答 1',
    columnMapping: { answer: 2, adminOnlyColumn: 999 },
    displaySettings: { showNames: true, showReactions: true, theme: 'x', pageSize: 50 },
    formUrl: 'https://docs.google.com/forms/d/new-form/viewform',
    formTitle: 'New Form',
    etag: 'etag-1',
    adminOnly: true,
    setupStatus: 'hacked'
  });

  assert.equal(result.success, true);
  assert.ok(savedPayload);
  assert.equal(savedPayload.adminOnly, undefined);
  assert.equal(savedPayload.setupStatus, 'completed');
  assert.equal(savedPayload.isPublished, true);
  assert.deepEqual(savedPayload.columnMapping, { answer: 2 });
  assert.deepEqual(savedPayload.displaySettings, {
    showNames: true,
    showReactions: true,
    theme: 'x',
    pageSize: 50
  });
});

test('publishApp: answer column is required', () => {
  const context = loadSystemControllerContext();
  const result = context.publishApp({
    spreadsheetId: 'spreadsheet-id',
    sheetName: 'フォームの回答 1',
    columnMapping: { reason: 3 },
    etag: 'etag-1'
  });

  assert.equal(result.success, false);
  assert.match(result.message, /answer/);
});
