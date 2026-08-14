const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { gasResponseStubs } = require('./_helpers.cjs');

function loadDataApisContext(overrides = {}) {
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    URL,
    ...gasResponseStubs(),
    getCurrentEmail: () => 'actor@example.com',
    findUserByEmail: () => null,
    findUserById: () => null,
    findPublishedBoardOwner: () => null,
    findUserBySpreadsheetId: () => null,
    getUserConfig: () => ({ success: true, config: {} }),
    getConfigOrDefault: () => ({}),
    saveUserConfig: () => ({ success: true }),
    openSpreadsheet: () => null,
    SpreadsheetApp: { openById: () => { throw new Error('not stubbed'); } },
    UrlFetchApp: { fetch: () => { throw new Error('not stubbed'); } },
    ScriptApp: { getService: () => ({ getUrl: () => 'https://script.google.com/x' }) },
    Session: { getActiveUser: () => ({ getEmail: () => 'actor@example.com' }) },
    CacheService: { getScriptCache: () => ({ get: () => null, put: () => {}, remove: () => {} }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null, setProperty: () => {} }) },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
    CACHE_DURATION: { SHORT: 10, MEDIUM: 30, LONG: 300, DATABASE_LONG: 600, FORM_DATA: 30 },
    DEFAULT_DISPLAY_SETTINGS: { showNames: false, showReactions: true, theme: 'default', pageSize: 20 },
    getBatchedAdminAuth: () => ({ success: true, authenticated: true, email: 'viewer@example.com', isAdmin: false }),
    getCachedProperty: () => null,
    saveToCacheWithSizeCheck: () => true,
    emailToShortHash: (email) => (email && typeof email === 'string') ? `h_${email.slice(0, 4)}` : null,
    validateEmail: (e) => ({ isValid: typeof e === 'string' && /.+@.+/.test(e) }),
    validateUrl: () => ({ isValid: true }),
    validateSpreadsheetId: () => true,
    normalizeHeader: (h) => String(h || '').toLowerCase().trim(),
    extractReactions: () => ({}),
    extractHighlight: () => false,
    formatTimestamp: (v) => String(v || ''),
    getQuestionText: () => '',
    createDataServiceErrorResponse: (message) => ({ success: false, message }),
    isAdministrator: () => false,
    // v2773: read-time cross-ref filter — default stub は sanitizer 本物と同じ規約
    //   (profiles[] に無い name は drop)。テストで orphan を含めたい場合は overrides で
    //   pass-through stub を渡す (v2772 既存テスト互換)。
    sanitizeProfileHistory: (history, profiles) => {
      if (!Array.isArray(history)) return [];
      const validNames = (Array.isArray(profiles) && profiles.length > 0)
        ? new Set(profiles.filter(p => p && p.name).map(p => p.name))
        : null;
      const out = [];
      for (const h of history) {
        if (!h || typeof h !== 'object' || !h.name) continue;
        if (validNames && !validNames.has(h.name)) continue;
        out.push({ name: h.name, activatedAt: h.activatedAt || '' });
      }
      return out;
    },
    ...overrides
  };
  vm.createContext(context);
  const source = fs.readFileSync(path.resolve(__dirname, '../src/DataApis.js'), 'utf8');
  vm.runInContext(source, context, { filename: 'DataApis.js' });
  return context;
}

// =====================================================================
// buildSafePublishedDataResult: boardMode auto-detection
// Why: フロントが boardMode を一意に受け取れる契約を保証。
//      auto → numericX 有無で numberline/board、numericY もあれば matrix。
// =====================================================================

test('boardMode auto: returns "board" when no numeric columns mapped', () => {
  const ctx = loadDataApisContext();
  const result = ctx.buildSafePublishedDataResult(
    { data: [{ id: 'row_2', answer: 'はい' }], header: 'q', sheetName: 's' },
    { displaySettings: { boardMode: 'auto' }, columnMapping: { answer: 1 } },
    { isAdmin: false, isOwnBoard: false }
  );
  assert.equal(result.displaySettings.boardMode, 'board');
});

test('boardMode auto: returns "numberline" when numericX only mapped', () => {
  const ctx = loadDataApisContext();
  const result = ctx.buildSafePublishedDataResult(
    { data: [{ id: 'row_2', numericX: 3 }], header: 'q', sheetName: 's' },
    { displaySettings: { boardMode: 'auto' }, columnMapping: { numericX: 4 } },
    { isAdmin: false, isOwnBoard: false }
  );
  assert.equal(result.displaySettings.boardMode, 'numberline');
});

test('boardMode auto: returns "matrix" when both numericX and numericY mapped', () => {
  const ctx = loadDataApisContext();
  const result = ctx.buildSafePublishedDataResult(
    { data: [{ id: 'row_2', numericX: 3, numericY: 4 }], header: 'q', sheetName: 's' },
    { displaySettings: { boardMode: 'auto' }, columnMapping: { numericX: 4, numericY: 5 } },
    { isAdmin: false, isOwnBoard: false }
  );
  assert.equal(result.displaySettings.boardMode, 'matrix');
});

test('boardMode auto: missing boardMode is treated as auto', () => {
  const ctx = loadDataApisContext();
  const result = ctx.buildSafePublishedDataResult(
    { data: [], header: 'q', sheetName: 's' },
    { displaySettings: {}, columnMapping: { numericX: 4 } },
    { isAdmin: false, isOwnBoard: false }
  );
  assert.equal(result.displaySettings.boardMode, 'numberline');
});

test('boardMode: explicit non-auto value is preserved', () => {
  const ctx = loadDataApisContext();
  const result = ctx.buildSafePublishedDataResult(
    { data: [], header: 'q', sheetName: 's' },
    { displaySettings: { boardMode: 'board' }, columnMapping: { numericX: 4, numericY: 5 } },
    { isAdmin: false, isOwnBoard: false }
  );
  // numericX/Y は揃っているが explicit 'board' を尊重
  assert.equal(result.displaySettings.boardMode, 'board');
});

// =====================================================================
// emailHash inclusion
// =====================================================================

test('emailHash: included for student viewer even when name/email stripped', () => {
  const ctx = loadDataApisContext();
  const result = ctx.buildSafePublishedDataResult(
    { data: [{ id: 'row_2', email: 'student@x.com', name: '田中', answer: 'はい' }], header: '', sheetName: '' },
    { displaySettings: { showNames: false, boardMode: 'numberline' }, columnMapping: { numericX: 4 } },
    { isAdmin: false, isOwnBoard: false }
  );
  const row = result.data[0];
  assert.equal(row.email, undefined, 'email should be stripped');
  assert.equal(row.name, undefined, 'name should be stripped');
  assert.equal(row.emailHash, 'h_stud', 'emailHash should be present for swing tracking');
});

test('emailHash: also included for owner viewer (does not duplicate, just adds)', () => {
  const ctx = loadDataApisContext();
  const result = ctx.buildSafePublishedDataResult(
    { data: [{ id: 'row_2', email: 'me@x.com', name: '本人', answer: 'はい' }], header: '', sheetName: '' },
    { displaySettings: { showNames: false, boardMode: 'board' }, columnMapping: {} },
    { isAdmin: false, isOwnBoard: true }
  );
  const row = result.data[0];
  assert.equal(row.email, 'me@x.com', 'owner sees raw email');
  assert.equal(row.name, '本人');
  assert.equal(row.emailHash, 'h_me@x');
});

test('emailHash: null when email missing', () => {
  const ctx = loadDataApisContext();
  const result = ctx.buildSafePublishedDataResult(
    { data: [{ id: 'row_2', answer: 'はい' }], header: '', sheetName: '' },
    { displaySettings: {}, columnMapping: {} },
    {}
  );
  // 元データに email 無し → cleaned.emailHash 未付与
  assert.equal(result.data[0].emailHash, undefined);
});

// =====================================================================
// axisConfig
// =====================================================================

test('axisConfig: passed through from config', () => {
  const ctx = loadDataApisContext();
  const result = ctx.buildSafePublishedDataResult(
    { data: [], header: '', sheetName: '' },
    {
      displaySettings: { boardMode: 'numberline' },
      columnMapping: { numericX: 4 },
      xAxisLabels: { min: 'そのまま', max: '直す' },
      allowResubmit: true
    },
    {}
  );
  assert.ok(result.axisConfig);
  assert.equal(result.axisConfig.xAxisLabels.min, 'そのまま');
  assert.equal(result.axisConfig.xAxisLabels.max, '直す');
  assert.equal(result.axisConfig.allowResubmit, true);
  assert.equal(result.axisConfig.defaultMin, 1);
  assert.equal(result.axisConfig.defaultMax, 5);
});

test('axisConfig: missing optional fields yield null safely', () => {
  const ctx = loadDataApisContext();
  const result = ctx.buildSafePublishedDataResult(
    { data: [], header: '', sheetName: '' },
    { displaySettings: {}, columnMapping: {} },
    {}
  );
  assert.equal(result.axisConfig.xAxisLabels, null);
  assert.equal(result.axisConfig.yAxisLabels, null);
  assert.equal(result.axisConfig.matrixQuadrantLabels, null);
  assert.equal(result.axisConfig.allowResubmit, false);
});

// =====================================================================
// formMeta — 生徒の「📝 回答フォーム」ボタン即時更新用
// Why: profile 切替で active config.formUrl が変わる → 5秒以内に生徒のボタンも追従。
// =====================================================================

test('formMeta: returns formUrl + formTitle from active config', () => {
  const ctx = loadDataApisContext();
  const result = ctx.buildSafePublishedDataResult(
    { data: [], header: '', sheetName: '' },
    {
      displaySettings: {},
      columnMapping: {},
      formUrl: 'https://docs.google.com/forms/d/abc/viewform',
      formTitle: '導入アンケート'
    },
    {}
  );
  assert.ok(result.formMeta);
  assert.equal(result.formMeta.formUrl, 'https://docs.google.com/forms/d/abc/viewform');
  assert.equal(result.formMeta.formTitle, '導入アンケート');
});

test('formMeta: returns empty strings when config has no formUrl', () => {
  const ctx = loadDataApisContext();
  const result = ctx.buildSafePublishedDataResult(
    { data: [], header: '', sheetName: '' },
    { displaySettings: {}, columnMapping: {} },
    {}
  );
  assert.ok(result.formMeta);
  assert.equal(result.formMeta.formUrl, '');
  assert.equal(result.formMeta.formTitle, '');
});

test('formMeta: rejects non-string formUrl (defensive)', () => {
  const ctx = loadDataApisContext();
  const result = ctx.buildSafePublishedDataResult(
    { data: [], header: '', sheetName: '' },
    { displaySettings: {}, columnMapping: {}, formUrl: { malicious: 'object' } },
    {}
  );
  assert.equal(result.formMeta.formUrl, '');
});


// =====================================================================
// viewerIsTeacher + profiles gating
// Why: profiles の wire 露出は所有者/管理者のみ。student に漏らさない。
// =====================================================================

test('viewerIsTeacher: true when isOwnBoard', () => {
  const ctx = loadDataApisContext();
  const result = ctx.buildSafePublishedDataResult(
    { data: [], header: '', sheetName: '' },
    { displaySettings: {}, columnMapping: {} },
    { isOwnBoard: true }
  );
  assert.equal(result.viewerIsTeacher, true);
});

test('viewerIsTeacher: true when isAdmin', () => {
  const ctx = loadDataApisContext();
  const result = ctx.buildSafePublishedDataResult(
    { data: [], header: '', sheetName: '' },
    { displaySettings: {}, columnMapping: {} },
    { isAdmin: true }
  );
  assert.equal(result.viewerIsTeacher, true);
});

test('viewerIsTeacher: false for regular viewer (student)', () => {
  const ctx = loadDataApisContext();
  const result = ctx.buildSafePublishedDataResult(
    { data: [], header: '', sheetName: '' },
    { displaySettings: {}, columnMapping: {} },
    {}
  );
  assert.equal(result.viewerIsTeacher, false);
});

