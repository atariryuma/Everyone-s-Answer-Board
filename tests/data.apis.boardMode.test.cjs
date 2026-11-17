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
    findUserByEmail_: () => null,
    findUserById_: () => null,
    findPublishedBoardOwner_: () => null,
    findUserBySpreadsheetId_: () => null,
    getUserConfig: () => ({ success: true, config: {} }),
    getConfigOrDefault: () => ({}),
    saveUserConfig_: () => ({ success: true }),
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
    getCachedProperty_: () => null,
    saveToCacheWithSizeCheck_: () => true,
    boardHashScope_: () => 'scope',
    emailToShortHash_: (email) => (email && typeof email === 'string') ? `h_${email.slice(0, 4)}` : null,
    validateEmail: (e) => ({ isValid: typeof e === 'string' && /.+@.+/.test(e) }),
    validateUrl: () => ({ isValid: true }),
    validateSpreadsheetId: () => true,
    normalizeHeader: (h) => String(h || '').toLowerCase().trim(),
    extractReactions: () => ({}),
    extractHighlight: () => false,
    formatTimestamp: (v) => String(v || ''),
    getQuestionText_: () => '',
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
// buildSafePublishedDataResult_: boardMode auto-detection
// Why: フロントが boardMode を一意に受け取れる契約を保証。
//      auto → numericX 有無で numberline/board、numericY もあれば matrix。
// =====================================================================

test('boardMode auto: returns "board" when no numeric columns mapped', () => {
  const ctx = loadDataApisContext();
  const result = ctx.buildSafePublishedDataResult_(
    { data: [{ id: 'row_2', answer: 'はい' }], header: 'q', sheetName: 's' },
    { displaySettings: { boardMode: 'auto' }, columnMapping: { answer: 1 } },
    { isAdmin: false, isOwnBoard: false }
  );
  assert.equal(result.displaySettings.boardMode, 'board');
});

test('boardMode auto: returns "numberline" when numericX only mapped', () => {
  const ctx = loadDataApisContext();
  const result = ctx.buildSafePublishedDataResult_(
    { data: [{ id: 'row_2', numericX: 3 }], header: 'q', sheetName: 's' },
    { displaySettings: { boardMode: 'auto' }, columnMapping: { numericX: 4 } },
    { isAdmin: false, isOwnBoard: false }
  );
  assert.equal(result.displaySettings.boardMode, 'numberline');
});

test('boardMode auto: returns "matrix" when both numericX and numericY mapped', () => {
  const ctx = loadDataApisContext();
  const result = ctx.buildSafePublishedDataResult_(
    { data: [{ id: 'row_2', numericX: 3, numericY: 4 }], header: 'q', sheetName: 's' },
    { displaySettings: { boardMode: 'auto' }, columnMapping: { numericX: 4, numericY: 5 } },
    { isAdmin: false, isOwnBoard: false }
  );
  assert.equal(result.displaySettings.boardMode, 'matrix');
});

test('boardMode auto: missing boardMode is treated as auto', () => {
  const ctx = loadDataApisContext();
  const result = ctx.buildSafePublishedDataResult_(
    { data: [], header: 'q', sheetName: 's' },
    { displaySettings: {}, columnMapping: { numericX: 4 } },
    { isAdmin: false, isOwnBoard: false }
  );
  assert.equal(result.displaySettings.boardMode, 'numberline');
});

test('boardMode: explicit non-auto value is preserved', () => {
  const ctx = loadDataApisContext();
  const result = ctx.buildSafePublishedDataResult_(
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
  const result = ctx.buildSafePublishedDataResult_(
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
  const result = ctx.buildSafePublishedDataResult_(
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
  const result = ctx.buildSafePublishedDataResult_(
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
  const result = ctx.buildSafePublishedDataResult_(
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
  const result = ctx.buildSafePublishedDataResult_(
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
  const result = ctx.buildSafePublishedDataResult_(
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
  const result = ctx.buildSafePublishedDataResult_(
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
  const result = ctx.buildSafePublishedDataResult_(
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
  const result = ctx.buildSafePublishedDataResult_(
    { data: [], header: '', sheetName: '' },
    { displaySettings: {}, columnMapping: {} },
    { isOwnBoard: true }
  );
  assert.equal(result.viewerIsTeacher, true);
});

test('viewerIsTeacher: true when isAdmin', () => {
  const ctx = loadDataApisContext();
  const result = ctx.buildSafePublishedDataResult_(
    { data: [], header: '', sheetName: '' },
    { displaySettings: {}, columnMapping: {} },
    { isAdmin: true }
  );
  assert.equal(result.viewerIsTeacher, true);
});

test('viewerIsTeacher: false for regular viewer (student)', () => {
  const ctx = loadDataApisContext();
  const result = ctx.buildSafePublishedDataResult_(
    { data: [], header: '', sheetName: '' },
    { displaySettings: {}, columnMapping: {} },
    {}
  );
  assert.equal(result.viewerIsTeacher, false);
});


// =====================================================================
// 「ほかの児童に見せない」(hidden) の回答
// Why: 行は閲覧者で共有する cache に全件入っている。ボードへ返す出口はここ 1 つなので、
//   ここで外れなければ、隠した回答がほかの児童にも電子黒板にも出る。
// =====================================================================

const HIDDEN_ROWS = () => ([
  { id: 'row_2', email: 'a@example.com', name: 'あ', answer: 'A', hidden: true },
  { id: 'row_3', email: 'b@example.com', name: 'い', answer: 'B', hidden: false },
  { id: 'row_4', email: 'c@example.com', name: 'う', answer: 'C' }
]);
const hiddenCfg = { displaySettings: { boardMode: 'board' }, columnMapping: { answer: 1 } };

test('hidden: ほかの児童には返さない', () => {
  const ctx = loadDataApisContext();
  const result = ctx.buildSafePublishedDataResult_(
    { data: HIDDEN_ROWS(), header: 'q', sheetName: 's' }, hiddenCfg,
    { isAdmin: false, isOwnBoard: false, email: 'b@example.com' });
  assert.deepEqual(Array.from(result.data, r => r.id), ['row_3', 'row_4']);
});

test('hidden: 本人には今までどおり返し、印は載せない (大文字小文字の違いは無視)', () => {
  const ctx = loadDataApisContext();
  const result = ctx.buildSafePublishedDataResult_(
    { data: HIDDEN_ROWS(), header: 'q', sheetName: 's' }, hiddenCfg,
    { isAdmin: false, isOwnBoard: false, email: 'A@Example.com' });
  assert.deepEqual(Array.from(result.data, r => r.id), ['row_2', 'row_3', 'row_4']);
  assert.equal(result.data.some(r => 'hidden' in r), false);
});

test('hidden: 先生のボード (電子黒板に映る) にも返さない。メールが分からない閲覧者にも返さない', () => {
  const ctx = loadDataApisContext();
  const teacher = ctx.buildSafePublishedDataResult_(
    { data: HIDDEN_ROWS(), header: 'q', sheetName: 's' }, hiddenCfg,
    { isAdmin: true, isOwnBoard: true, email: 'teacher@example.com' });
  assert.deepEqual(Array.from(teacher.data, r => r.id), ['row_3', 'row_4']);
  const anon = ctx.buildSafePublishedDataResult_(
    { data: HIDDEN_ROWS(), header: 'q', sheetName: 's' }, hiddenCfg, { isAdmin: false, isOwnBoard: true });
  assert.deepEqual(Array.from(anon.data, r => r.id), ['row_3', 'row_4']);
});

test('hidden: 授業の記録を取るとき (includeHidden) は全件を返す。印は載せない', () => {
  const ctx = loadDataApisContext();
  const result = ctx.buildSafePublishedDataResult_(
    { data: HIDDEN_ROWS(), header: 'q', sheetName: 's' }, hiddenCfg,
    { isAdmin: true, isOwnBoard: true, email: 'teacher@example.com', includeHidden: true });
  assert.deepEqual(Array.from(result.data, r => r.id), ['row_2', 'row_3', 'row_4']);
  assert.equal(result.data.some(r => 'hidden' in r), false);
});

// 児童の読み込みは、新しい版を読んでいる間、直前の結果 (押す前の行) を返すことがある。
//   押したばかりの回答を覚えておき、直前の結果にも印を付け直す。
test('hidden: 押したばかりの回答は、直前の結果 (印の無い古い行) にも印を付け直す。戻したら外れる', () => {
  const store = new Map();
  const ctx = loadDataApisContext({
    CacheService: { getScriptCache: () => ({
      get: (k) => (store.has(k) ? store.get(k) : null), put: (k, v) => { store.set(k, v); }, remove: (k) => { store.delete(k); },
      getAll: (keys) => { const o = {}; keys.forEach((k) => { if (store.has(k)) o[k] = store.get(k); }); return o; }
    }) }
  });
  const staleRows = () => ({ success: true, stale: true, data: HIDDEN_ROWS().map(r => Object.assign({}, r, { hidden: false })) });
  const marked = (uid, sheet) => Array.from(ctx.markRecentlyHiddenRows_(staleRows(), uid, sheet).data.filter(r => r.hidden), r => r.id);

  assert.deepEqual(marked('u1', 'phase1'), []);

  // 2 つを続けて押しても、どちらも覚えている (回答ごとに別の key なので、書き戻しで消し合わない)
  ctx.noteRecentlyHidden_('u1', 'phase1', 'B@example.com', true);
  ctx.noteRecentlyHidden_('u1', 'phase1', 'c@example.com', true);
  assert.deepEqual(marked('u1', 'phase1'), ['row_3', 'row_4']);
  // ほかのボード・次のフェーズのシートの回答には付けない
  assert.deepEqual(marked('u2', 'phase1'), []);
  assert.deepEqual(marked('u1', 'phase4'), []);

  ctx.noteRecentlyHidden_('u1', 'phase1', 'b@example.com', false);
  assert.deepEqual(marked('u1', 'phase1'), ['row_4']);
});
