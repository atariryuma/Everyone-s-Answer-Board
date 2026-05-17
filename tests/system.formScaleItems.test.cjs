/**
 * __extractFormScaleItems_ regression tests (v2780).
 *
 * Why: 軸ラベル自動補完の中核。FormApp.openByUrl → getItems(SCALE) → asScaleItem
 *   の chain が壊れたら本番で「軸ラベル placeholder/自動補完が出ない」regression。
 *   FormApp 不在環境 (= node test) では vm context で stub する。
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function makeScaleItem(title, left, right, lower, upper) {
  return {
    getTitle: () => title,
    asScaleItem: () => ({
      getLeftLabel: () => left,
      getRightLabel: () => right,
      getLowerBound: () => lower,
      getUpperBound: () => upper
    })
  };
}

function loadCtx(formAppOverrides) {
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    FormApp: formAppOverrides,
    // 残りの SystemController.js が parse-time に必要とする shim
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null, setProperty: () => {} }) },
    CacheService: { getScriptCache: () => ({ get: () => null, put: () => {}, remove: () => {} }) },
    Session: { getActiveUser: () => ({ getEmail: () => 't@example.com' }) },
    SpreadsheetApp: { openById: () => null },
    DriveApp: {},
    ScriptApp: { getService: () => ({ getUrl: () => '' }) },
    UrlFetchApp: { fetch: () => { throw new Error('not stubbed'); } },
    Utilities: { sleep: () => {}, getUuid: () => 'uuid' },
    CACHE_DURATION: { SHORT: 10, MEDIUM: 30, LONG: 300 },
    DEFAULT_DISPLAY_SETTINGS: {},
    SYSTEM_LIMITS: { CONFIG_JSON_MAX_CHARS: 32000 },
    getCurrentEmail: () => null,
    isAdministrator: () => false,
    getCachedProperty: () => null,
    findUserByEmail: () => null,
    saveToCacheWithSizeCheck: () => true,
    openSpreadsheet: () => null,
    detectFormConnection: () => ({}),
    createErrorResponse: (msg) => ({ success: false, message: msg }),
    createSuccessResponse: (msg, data) => ({ success: true, message: msg, data })
  };
  vm.createContext(context);
  const src = fs.readFileSync(path.resolve(__dirname, '../src/SystemController.js'), 'utf8');
  vm.runInContext(src, context, { filename: 'SystemController.js' });
  return context;
}

test('__extractFormScaleItems_: returns null when formUrl is missing', () => {
  const ctx = loadCtx({ openByUrl: () => null, ItemType: { SCALE: 'SCALE' } });
  assert.equal(ctx.__extractFormScaleItems_(''), null);
  assert.equal(ctx.__extractFormScaleItems_(null), null);
});

test('__extractFormScaleItems_: returns null when FormApp is unavailable', () => {
  const ctx = loadCtx(undefined);  // FormApp is undefined
  assert.equal(ctx.__extractFormScaleItems_('https://docs.google.com/forms/d/X/viewform'), null);
});

test('__extractFormScaleItems_: extracts labels and bounds for matrix form (2 scale items)', () => {
  const items = [
    makeScaleItem('X 軸：効率↔倫理', '効率優先', '倫理優先', 1, 5),
    makeScaleItem('Y 軸：個人↔集団', '個人重視', '集団重視', 1, 5)
  ];
  const ctx = loadCtx({
    openByUrl: () => ({ getItems: () => items }),
    ItemType: { SCALE: 'SCALE' }
  });
  const result = ctx.__extractFormScaleItems_('https://docs.google.com/forms/d/X/viewform');
  assert.equal(result.length, 2);
  assert.equal(result[0].leftLabel, '効率優先');
  assert.equal(result[0].rightLabel, '倫理優先');
  assert.equal(result[0].lowerBound, 1);
  assert.equal(result[0].upperBound, 5);
  assert.equal(result[1].leftLabel, '個人重視');
  assert.equal(result[1].rightLabel, '集団重視');
});

test('__extractFormScaleItems_: returns single item for numberline form', () => {
  const items = [makeScaleItem('立場', 'そう思わない', 'とてもそう思う', 1, 5)];
  const ctx = loadCtx({
    openByUrl: () => ({ getItems: () => items }),
    ItemType: { SCALE: 'SCALE' }
  });
  const result = ctx.__extractFormScaleItems_('https://docs.google.com/forms/d/X/viewform');
  assert.equal(result.length, 1);
  assert.equal(result[0].leftLabel, 'そう思わない');
  assert.equal(result[0].rightLabel, 'とてもそう思う');
});

test('__extractFormScaleItems_: returns empty array for form without scale items', () => {
  const ctx = loadCtx({
    openByUrl: () => ({ getItems: () => [] }),
    ItemType: { SCALE: 'SCALE' }
  });
  const result = ctx.__extractFormScaleItems_('https://docs.google.com/forms/d/X/viewform');
  assert.deepEqual(result, []);
});

test('__extractFormScaleItems_: returns null on FormApp.openByUrl error (権限 etc.)', () => {
  const ctx = loadCtx({
    openByUrl: () => { throw new Error('Permission denied'); },
    ItemType: { SCALE: 'SCALE' }
  });
  const result = ctx.__extractFormScaleItems_('https://docs.google.com/forms/d/X/viewform');
  assert.equal(result, null);
});

test('__extractFormScaleItems_: handles null label strings gracefully', () => {
  const items = [
    {
      getTitle: () => 'X 軸',
      asScaleItem: () => ({
        getLeftLabel: () => null,
        getRightLabel: () => null,
        getLowerBound: () => 1,
        getUpperBound: () => 5
      })
    }
  ];
  const ctx = loadCtx({
    openByUrl: () => ({ getItems: () => items }),
    ItemType: { SCALE: 'SCALE' }
  });
  const result = ctx.__extractFormScaleItems_('https://docs.google.com/forms/d/X/viewform');
  assert.equal(result[0].leftLabel, '');
  assert.equal(result[0].rightLabel, '');
});
