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
    DEFAULT_DISPLAY_SETTINGS: { showNames: false, showReactions: false, theme: 'default', pageSize: 20 },
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

  context.getConfigOrDefault = (userId, user) => {
    const result = context.getUserConfig(userId, user);
    return result.success ? result.config : {};
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

// =====================================================================
// sanitizePublishPayload_
// =====================================================================

test('sanitizePublishPayload_: returns empty config for non-object input', () => {
  const ctx = loadSystemControllerContext();
  for (const bad of [null, undefined, [], 42, 'string', true]) {
    const { config, ignoredKeys } = ctx.sanitizePublishPayload_(bad);
    assert.equal(typeof config, 'object');
    assert.ok(!Array.isArray(config));
    assert.deepEqual([...ignoredKeys], []);
  }
});

test('sanitizePublishPayload_: trims spreadsheetId and sheetName strings', () => {
  const ctx = loadSystemControllerContext();
  const { config } = ctx.sanitizePublishPayload_({
    spreadsheetId: '  ss-123  ',
    sheetName: '  Sheet1\n'
  });
  assert.equal(config.spreadsheetId, 'ss-123');
  assert.equal(config.sheetName, 'Sheet1');
});

test('sanitizePublishPayload_: drops non-string spreadsheetId and sheetName', () => {
  const ctx = loadSystemControllerContext();
  const { config } = ctx.sanitizePublishPayload_({
    spreadsheetId: 123,
    sheetName: ['Sheet1']
  });
  assert.equal(config.spreadsheetId, undefined);
  assert.equal(config.sheetName, undefined);
});

test('sanitizePublishPayload_: preserves currentConfig.displaySettings when request omits them', () => {
  const ctx = loadSystemControllerContext();
  const { config } = ctx.sanitizePublishPayload_(
    { spreadsheetId: 'ss' },
    {
      displaySettings: { showNames: true, showReactions: false, theme: 'dark', pageSize: 30 }
    }
  );
  assert.ok(config.displaySettings);
  assert.equal(config.displaySettings.showNames, true);
});

test('sanitizePublishPayload_: incoming displaySettings override currentConfig', () => {
  const ctx = loadSystemControllerContext();
  const { config } = ctx.sanitizePublishPayload_(
    { displaySettings: { showNames: false, showReactions: true, theme: 'light', pageSize: 10 } },
    { displaySettings: { showNames: true, showReactions: true, theme: 'dark', pageSize: 30 } }
  );
  assert.equal(config.displaySettings.showNames, false);
  assert.equal(config.displaySettings.theme, 'light');
});

test('sanitizePublishPayload_: array/non-object displaySettings fall back to currentConfig', () => {
  const ctx = loadSystemControllerContext();
  const currentConfig = { displaySettings: { showNames: true, showReactions: false, theme: 'a', pageSize: 20 } };
  // Array is not a plain object
  const resultA = ctx.sanitizePublishPayload_({ displaySettings: ['x'] }, currentConfig);
  assert.equal(resultA.config.displaySettings.showNames, true);
  // null is not a plain object
  const resultB = ctx.sanitizePublishPayload_({ displaySettings: null }, currentConfig);
  assert.equal(resultB.config.displaySettings.showNames, true);
});

test('sanitizePublishPayload_: formUrl falls back to currentConfig when request lacks it', () => {
  const ctx = loadSystemControllerContext();
  const { config } = ctx.sanitizePublishPayload_(
    {},
    { formUrl: 'https://docs.google.com/forms/existing' }
  );
  assert.equal(config.formUrl, 'https://docs.google.com/forms/existing');
});

test('sanitizePublishPayload_: whitespace-only formUrl is treated as absent', () => {
  const ctx = loadSystemControllerContext();
  const { config } = ctx.sanitizePublishPayload_(
    { formUrl: '   ' },
    { formUrl: 'https://docs.google.com/forms/existing' }
  );
  assert.equal(config.formUrl, 'https://docs.google.com/forms/existing');
});

test('sanitizePublishPayload_: truncates long formTitle to 4 * PREVIEW_LENGTH', () => {
  const ctx = loadSystemControllerContext();
  const { config } = ctx.sanitizePublishPayload_({ formTitle: 'x'.repeat(1000) });
  assert.ok(config.formTitle.length <= 200); // PREVIEW_LENGTH=50, *4=200
});

test('sanitizePublishPayload_: showDetails coerced only from boolean', () => {
  const ctx = loadSystemControllerContext();
  assert.equal(ctx.sanitizePublishPayload_({ showDetails: true }).config.showDetails, true);
  assert.equal(ctx.sanitizePublishPayload_({ showDetails: false }).config.showDetails, false);
  assert.equal(ctx.sanitizePublishPayload_({ showDetails: 'true' }).config.showDetails, undefined);
  assert.equal(ctx.sanitizePublishPayload_({ showDetails: 1 }).config.showDetails, undefined);
});

test('sanitizePublishPayload_: showDetails falls back to currentConfig when request has non-boolean', () => {
  const ctx = loadSystemControllerContext();
  const { config } = ctx.sanitizePublishPayload_(
    { showDetails: 'yes' },
    { showDetails: false }
  );
  assert.equal(config.showDetails, false);
});

test('sanitizePublishPayload_: etag must be string 1..255 chars', () => {
  const ctx = loadSystemControllerContext();
  assert.equal(ctx.sanitizePublishPayload_({ etag: 'abc' }).config.etag, 'abc');
  assert.equal(ctx.sanitizePublishPayload_({ etag: 'x'.repeat(256) }).config.etag, undefined);
  assert.equal(ctx.sanitizePublishPayload_({ etag: '' }).config.etag, undefined);
  assert.equal(ctx.sanitizePublishPayload_({ etag: 123 }).config.etag, undefined);
});

test('sanitizePublishPayload_: returns ignoredKeys list for unknown fields', () => {
  const ctx = loadSystemControllerContext();
  const { ignoredKeys } = ctx.sanitizePublishPayload_({
    spreadsheetId: 'ss',
    adminOnly: true,
    secretField: 'x',
    setupStatus: 'hacked'
  });
  assert.ok(ignoredKeys.includes('adminOnly'));
  assert.ok(ignoredKeys.includes('secretField'));
  assert.ok(ignoredKeys.includes('setupStatus'));
  assert.ok(!ignoredKeys.includes('spreadsheetId'));
});

// =====================================================================
// isValidWebAppUrl_
// =====================================================================

test('isValidWebAppUrl_: accepts production /exec URL', () => {
  const ctx = loadSystemControllerContext();
  assert.equal(ctx.isValidWebAppUrl_('https://script.google.com/macros/s/abc/exec'), true);
});

test('isValidWebAppUrl_: accepts dev /dev URL', () => {
  const ctx = loadSystemControllerContext();
  assert.equal(ctx.isValidWebAppUrl_('https://script.google.com/macros/s/abc/dev'), true);
});

test('isValidWebAppUrl_: rejects OAuth cold-start userCodeAppPanel URL', () => {
  const ctx = loadSystemControllerContext();
  assert.equal(
    ctx.isValidWebAppUrl_('https://script.google.com/macros/s/abc/exec?userCodeAppPanel=x'),
    false
  );
});

test('isValidWebAppUrl_: rejects createOAuthDialog URL', () => {
  const ctx = loadSystemControllerContext();
  assert.equal(
    ctx.isValidWebAppUrl_('https://script.google.com/createOAuthDialog?x=1'),
    false
  );
});

test('isValidWebAppUrl_: rejects URL without /exec or /dev', () => {
  const ctx = loadSystemControllerContext();
  assert.equal(ctx.isValidWebAppUrl_('https://script.google.com/macros/s/abc'), false);
  assert.equal(ctx.isValidWebAppUrl_('https://example.com/some/path'), false);
});

test('isValidWebAppUrl_: rejects null, undefined, empty, non-string', () => {
  const ctx = loadSystemControllerContext();
  assert.equal(ctx.isValidWebAppUrl_(null), false);
  assert.equal(ctx.isValidWebAppUrl_(undefined), false);
  assert.equal(ctx.isValidWebAppUrl_(''), false);
  assert.equal(ctx.isValidWebAppUrl_(42), false);
});
