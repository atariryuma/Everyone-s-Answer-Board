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
    // DataApis.js は findPublishedBoardOwner 経由で公開ボード閲覧時の lookup を行うため、
    // 上書きを尊重しつつ findUserById にデリゲートする。
    findPublishedBoardOwner: (userId, viewerEmail, extra = {}) =>
      (context.findUserById ? context.findUserById(userId, { ...extra, requestingUser: viewerEmail, allowPublishedRead: true }) : null),
    findUserBySpreadsheetId: () => null,
    getUserConfig: () => ({ success: true, config: {} }),
    getConfigOrDefault: () => ({}),
    saveUserConfig: () => ({ success: true }),
    openSpreadsheet: () => null,
    // 既定では openById 成功 = 呼び出し元が当該 SS へアクセス権を持つ (owner 想定)。
    //   アクセス権なしを検証するテストは openById を throw する stub で上書きする。
    SpreadsheetApp: { openById: () => ({ getSheetByName: () => null, getName: () => 'mock-ss' }) },
    UrlFetchApp: { fetch: () => { throw new Error('not stubbed'); } },
    ScriptApp: { getService: () => ({ getUrl: () => 'https://script.google.com/x' }) },
    Session: { getActiveUser: () => ({ getEmail: () => 'actor@example.com' }) },
    CacheService: { getScriptCache: () => ({ get: () => null, put: () => {}, remove: () => {} }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null, setProperty: () => {} }) },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
    CACHE_DURATION: { SHORT: 10, MEDIUM: 30, LONG: 300, DATABASE_LONG: 600, FORM_DATA: 30 },
    TIMEOUT_MS: { DEFAULT: 5000 },
    SYSTEM_LIMITS: {},
    DEFAULT_DISPLAY_SETTINGS: { showNames: false, showReactions: true, theme: 'default', pageSize: 20 },
    getBatchedAdminAuth: () => ({ success: true, authenticated: true, email: 'viewer@example.com', isAdmin: false }),
    getCachedProperty: () => null,
    saveToCacheWithSizeCheck: () => true,
    // emailHash は helpers.js 由来（M1/M2 可視化モードで揺らぎ追跡に使う仮名化ID）。
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
    ...overrides
  };

  vm.createContext(context);
  const source = fs.readFileSync(path.resolve(__dirname, '../src/DataApis.js'), 'utf8');
  vm.runInContext(source, context, { filename: 'DataApis.js' });
  return context;
}

// =====================================================================
// sameEmail_ (owner 判定の case-insensitive email 比較)
// =====================================================================

test('sameEmail_: case/whitespace を無視して一致 (owner 誤判定防止)', () => {
  const ctx = loadDataApisContext();
  assert.equal(ctx.sameEmail_('User@Naha-Okinawa.ed.jp', 'user@naha-okinawa.ed.jp'), true);
  assert.equal(ctx.sameEmail_('  teacher@x.jp ', 'teacher@x.jp'), true);
  assert.equal(ctx.sameEmail_('a@x.jp', 'b@x.jp'), false);
  assert.equal(ctx.sameEmail_(null, ''), true);     // 両方空相当
  assert.equal(ctx.sameEmail_('a@x.jp', null), false);
});

// =====================================================================
// isValidFormUrl
// =====================================================================

test('isValidFormUrl: accepts docs.google.com forms URL', () => {
  const ctx = loadDataApisContext();
  assert.equal(ctx.isValidFormUrl('https://docs.google.com/forms/d/e/1FAIpQLSd.../viewform'), true);
});

test('isValidFormUrl: accepts forms.gle short URL', () => {
  const ctx = loadDataApisContext();
  assert.equal(ctx.isValidFormUrl('https://forms.gle/abc123'), true);
});

test('isValidFormUrl: rejects http:// (requires https)', () => {
  const ctx = loadDataApisContext();
  assert.equal(ctx.isValidFormUrl('http://docs.google.com/forms/d/e/x/viewform'), false);
});

test('isValidFormUrl: rejects javascript: protocol', () => {
  const ctx = loadDataApisContext();
  assert.equal(ctx.isValidFormUrl('javascript:alert(1)'), false);
});

test('isValidFormUrl: rejects data: protocol', () => {
  const ctx = loadDataApisContext();
  assert.equal(ctx.isValidFormUrl('data:text/html,<script>alert(1)</script>'), false);
});

test('isValidFormUrl: rejects non-Google hosts', () => {
  const ctx = loadDataApisContext();
  assert.equal(ctx.isValidFormUrl('https://evil.example.com/forms'), false);
  assert.equal(ctx.isValidFormUrl('https://docs.google.com.evil.com/forms/x'), false);
});

test('isValidFormUrl: rejects docs.google.com without /forms/ or /viewform', () => {
  const ctx = loadDataApisContext();
  assert.equal(ctx.isValidFormUrl('https://docs.google.com/spreadsheets/d/xxx'), false);
  assert.equal(ctx.isValidFormUrl('https://docs.google.com/document/d/xxx'), false);
});

test('isValidFormUrl: rejects null, undefined, empty, non-string', () => {
  const ctx = loadDataApisContext();
  assert.equal(ctx.isValidFormUrl(null), false);
  assert.equal(ctx.isValidFormUrl(undefined), false);
  assert.equal(ctx.isValidFormUrl(''), false);
  assert.equal(ctx.isValidFormUrl(42), false);
  assert.equal(ctx.isValidFormUrl({}), false);
});

test('isValidFormUrl: trims whitespace before validation', () => {
  const ctx = loadDataApisContext();
  assert.equal(ctx.isValidFormUrl('  https://forms.gle/abc  '), true);
});

test('isValidFormUrl: rejects malformed URL strings', () => {
  const ctx = loadDataApisContext();
  assert.equal(ctx.isValidFormUrl('not a url'), false);
  assert.equal(ctx.isValidFormUrl('https://'), false);
});

// =====================================================================
// validateHeaderIntegrity
// =====================================================================

test('validateHeaderIntegrity: returns user-not-found without session or target', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => null,
    findUserById: () => null,
    findUserByEmail: () => null
  });
  const result = ctx.validateHeaderIntegrity(null);
  assert.equal(result.success, false);
});

test('validateHeaderIntegrity: returns error when spreadsheet config incomplete', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    findUserById: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    getConfigOrDefault: () => ({ spreadsheetId: '', sheetName: '' })
  });
  const result = ctx.validateHeaderIntegrity('u1');
  assert.equal(result.success, false);
  assert.match(result.error, /configuration incomplete/i);
});

test('validateHeaderIntegrity: returns error when target sheet not found', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    findUserById: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    getConfigOrDefault: () => ({ spreadsheetId: 'ss', sheetName: 'Missing' }),
    openSpreadsheet: () => ({ getSheet: () => null })
  });
  const result = ctx.validateHeaderIntegrity('u1');
  assert.equal(result.success, false);
  assert.match(result.error, /Sheet not found/);
});

test('validateHeaderIntegrity: returns valid=true for healthy headers', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    findUserById: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    getConfigOrDefault: () => ({ spreadsheetId: 'ss', sheetName: 'Sheet1' }),
    openSpreadsheet: () => ({ getSheet: () => ({ name: 'Sheet1' }) }),
    getSheetInfo: () => ({ headers: ['Q1', '理由', '名前', 'LIKE'], lastRow: 10, lastCol: 4 })
  });
  const result = ctx.validateHeaderIntegrity('u1');
  assert.equal(result.success, true);
  assert.equal(result.valid, true);
  assert.equal(result.headerCount, 4);
  assert.equal(result.emptyHeaderCount, 0);
});

test('validateHeaderIntegrity: reports empty-header count', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    findUserById: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    getConfigOrDefault: () => ({ spreadsheetId: 'ss', sheetName: 'Sheet1' }),
    openSpreadsheet: () => ({ getSheet: () => ({}) }),
    getSheetInfo: () => ({ headers: ['Q1', '', '  ', 'LIKE'], lastRow: 2, lastCol: 4 })
  });
  const result = ctx.validateHeaderIntegrity('u1');
  assert.equal(result.valid, true); // at least one non-empty header
  assert.equal(result.emptyHeaderCount, 2);
});

test('validateHeaderIntegrity: returns valid=false when all headers are empty', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    findUserById: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    getConfigOrDefault: () => ({ spreadsheetId: 'ss', sheetName: 'Sheet1' }),
    openSpreadsheet: () => ({ getSheet: () => ({}) }),
    getSheetInfo: () => ({ headers: ['', '  ', null], lastRow: 0, lastCol: 3 })
  });
  const result = ctx.validateHeaderIntegrity('u1');
  assert.equal(result.success, false);
  assert.equal(result.valid, false);
});

// =====================================================================
// getPublishedSheetData — authorization decisions
// =====================================================================

function ctxWithAuth(overrides = {}) {
  return loadDataApisContext({
    getBatchedAdminAuth: () => ({
      success: true,
      authenticated: true,
      email: 'viewer@example.com',
      isAdmin: false
    }),
    getUserSheetData: () => ({
      success: true,
      data: [{ id: 1 }],
      headers: ['Q1'],
      sheetName: 'Sheet1',
      header: '',
      showDetails: true
    }),
    getQuestionText: () => '',
    ...overrides
  });
}

test('getPublishedSheetData: rejects unauthenticated viewer', () => {
  const ctx = ctxWithAuth({
    getBatchedAdminAuth: () => ({ success: false, authenticated: false })
  });
  const result = ctx.getPublishedSheetData(null, 'newest', false, 'u1');
  assert.equal(result.success, false);
  assert.match(result.error, /Authentication/);
});

test('getPublishedSheetData: rejects when target user not found', () => {
  const ctx = ctxWithAuth({
    findUserById: () => null
  });
  const result = ctx.getPublishedSheetData(null, 'newest', false, 'ghost');
  assert.equal(result.success, false);
  assert.match(result.error, /Target user/);
});

test('getPublishedSheetData: blocks non-admin non-owner viewer from unpublished board', () => {
  let getUserSheetDataCalled = false;
  const ctx = ctxWithAuth({
    findUserById: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    getConfigOrDefault: () => ({ spreadsheetId: 'ss', sheetName: 'Sheet1', isPublished: false }),
    getUserSheetData: () => { getUserSheetDataCalled = true; return { success: true }; }
  });
  const result = ctx.getPublishedSheetData(null, 'newest', false, 'u1');
  assert.equal(result.success, false);
  assert.match(result.error, /未公開/);
  assert.equal(getUserSheetDataCalled, false, 'Must not fetch data for unauthorized viewer');
});

test('getPublishedSheetData: allows non-admin on a published board', () => {
  const ctx = ctxWithAuth({
    findUserById: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    getConfigOrDefault: () => ({ spreadsheetId: 'ss', sheetName: 'Sheet1', isPublished: true })
  });
  const result = ctx.getPublishedSheetData(null, 'newest', false, 'u1');
  assert.equal(result.success, true);
});

test('getPublishedSheetData: allows owner to view their own unpublished board', () => {
  const ctx = ctxWithAuth({
    getBatchedAdminAuth: () => ({
      success: true, authenticated: true,
      email: 'owner@example.com', isAdmin: false
    }),
    findUserById: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    getConfigOrDefault: () => ({ spreadsheetId: 'ss', sheetName: 'Sheet1', isPublished: false })
  });
  const result = ctx.getPublishedSheetData(null, 'newest', false, 'u1');
  assert.equal(result.success, true);
});

test('getPublishedSheetData: allows system admin regardless of publish state', () => {
  const ctx = ctxWithAuth({
    getBatchedAdminAuth: () => ({
      success: true, authenticated: true,
      email: 'admin@example.com', isAdmin: true
    }),
    findUserById: () => ({ userId: 'u1', userEmail: 'other@example.com' }),
    getConfigOrDefault: () => ({ spreadsheetId: 'ss', sheetName: 'Sheet1', isPublished: false })
  });
  const result = ctx.getPublishedSheetData(null, 'newest', false, 'u1');
  assert.equal(result.success, true);
});

test('getPublishedSheetData: forwards classFilter and sortOrder to getUserSheetData', () => {
  let capturedOptions = null;
  const ctx = ctxWithAuth({
    findUserById: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    getConfigOrDefault: () => ({ spreadsheetId: 'ss', sheetName: 'Sheet1', isPublished: true }),
    getUserSheetData: (_userId, options) => {
      capturedOptions = options;
      return { success: true, data: [], headers: [], sheetName: '' };
    }
  });
  ctx.getPublishedSheetData('3-1', 'oldest', false, 'u1');
  assert.equal(capturedOptions.classFilter, '3-1');
  assert.equal(capturedOptions.sortBy, 'oldest');
});

test('getPublishedSheetData: classFilter="すべて" is converted to undefined', () => {
  let capturedOptions = null;
  const ctx = ctxWithAuth({
    findUserById: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    getConfigOrDefault: () => ({ spreadsheetId: 'ss', sheetName: 'Sheet1', isPublished: true }),
    getUserSheetData: (_userId, options) => {
      capturedOptions = options;
      return { success: true, data: [], headers: [], sheetName: '' };
    }
  });
  ctx.getPublishedSheetData('すべて', 'newest', false, 'u1');
  assert.equal(capturedOptions.classFilter, undefined);
});

// =====================================================================
// Option B: getPublishedSheetDataForProfile (past-phase read-only endpoint)
// =====================================================================

function ctxWithProfileHistory(profileHistory, profiles, overrides = {}) {
  return ctxWithAuth({
    findUserById: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    getConfigOrDefault: () => ({
      spreadsheetId: 'current-ss',
      sheetName: 'Current',
      isPublished: true,
      activeProfile: '本時',
      profiles,
      profileHistory
    }),
    getUserSheetData: () => ({
      success: true, data: [{ id: 1 }], headers: [], sheetName: 'Past', header: ''
    }),
    ...overrides
  });
}

const SAMPLE_PROFILES = [
  { name: '導入', formTitle: 'イントロ', spreadsheetId: 'past-ss', sheetName: 'Intro', columnMapping: {} },
  { name: '本時', formTitle: '討論', spreadsheetId: 'current-ss', sheetName: 'Current', columnMapping: {} },
  { name: '未来', formTitle: '振り返り', spreadsheetId: 'future-ss', sheetName: 'Future', columnMapping: {} }
];
const SAMPLE_HISTORY = [
  { name: '導入', activatedAt: '2026-05-14T00:00:00Z' },
  { name: '本時', activatedAt: '2026-05-14T01:00:00Z' }
];

test('getPublishedSheetDataForProfile: rejects missing targetUserId', () => {
  const ctx = ctxWithProfileHistory(SAMPLE_HISTORY, SAMPLE_PROFILES);
  const result = ctx.getPublishedSheetDataForProfile('', '導入');
  assert.equal(result.success, false);
  assert.match(result.error, /targetUserId/);
});

test('getPublishedSheetDataForProfile: rejects missing profileName', () => {
  const ctx = ctxWithProfileHistory(SAMPLE_HISTORY, SAMPLE_PROFILES);
  const result = ctx.getPublishedSheetDataForProfile('u1', '');
  assert.equal(result.success, false);
  assert.match(result.error, /profileName/);
});

test('getPublishedSheetDataForProfile: students cannot view future profile (not in history)', () => {
  // viewer != owner & not admin → student
  const ctx = ctxWithProfileHistory(SAMPLE_HISTORY, SAMPLE_PROFILES, {
    getBatchedAdminAuth: () => ({
      success: true, authenticated: true,
      email: 'student@example.com', isAdmin: false
    })
  });
  const result = ctx.getPublishedSheetDataForProfile('u1', '未来');
  assert.equal(result.success, false);
  assert.match(result.error, /公開/);
});

test('getPublishedSheetDataForProfile: students can view past profile (in history)', () => {
  const ctx = ctxWithProfileHistory(SAMPLE_HISTORY, SAMPLE_PROFILES, {
    getBatchedAdminAuth: () => ({
      success: true, authenticated: true,
      email: 'student@example.com', isAdmin: false
    })
  });
  const result = ctx.getPublishedSheetDataForProfile('u1', '導入');
  assert.equal(result.success, true);
  assert.equal(result.viewingPastProfile, '導入');
});

test('getPublishedSheetDataForProfile: owner can view any profile (preview future)', () => {
  // owner: history gate is bypassed
  const ctx = ctxWithProfileHistory(SAMPLE_HISTORY, SAMPLE_PROFILES, {
    getBatchedAdminAuth: () => ({
      success: true, authenticated: true,
      email: 'owner@example.com', isAdmin: false  // owner of the board
    })
  });
  const result = ctx.getPublishedSheetDataForProfile('u1', '未来');
  assert.equal(result.success, true);
});

test('getPublishedSheetDataForProfile: rejects when board unpublished and viewer not privileged', () => {
  const ctx = ctxWithProfileHistory(SAMPLE_HISTORY, SAMPLE_PROFILES, {
    getBatchedAdminAuth: () => ({
      success: true, authenticated: true,
      email: 'student@example.com', isAdmin: false
    }),
    getConfigOrDefault: () => ({
      spreadsheetId: 'current-ss',
      sheetName: 'Current',
      isPublished: false,
      activeProfile: '本時',
      profiles: SAMPLE_PROFILES,
      profileHistory: SAMPLE_HISTORY
    })
  });
  const result = ctx.getPublishedSheetDataForProfile('u1', '導入');
  assert.equal(result.success, false);
  assert.match(result.error, /未公開/);
});

test('getPublishedSheetDataForProfile: synthesizes config from profile snapshot', () => {
  let captured = null;
  const ctx = ctxWithProfileHistory(SAMPLE_HISTORY, SAMPLE_PROFILES, {
    getUserSheetData: (_uid, _opts, _user, cfg) => {
      captured = cfg;
      return { success: true, data: [], sheetName: 'Intro', headers: [], header: '' };
    }
  });
  ctx.getPublishedSheetDataForProfile('u1', '導入');
  assert.ok(captured);
  // 過去 profile の SS を渡している（現在の SS ではない）
  assert.equal(captured.spreadsheetId, 'past-ss');
  assert.equal(captured.sheetName, 'Intro');
});

test('getPublishedSheetDataForProfile: rejects deleted profile (history hit, profiles miss)', () => {
  const profiles = [
    { name: '本時', spreadsheetId: 'current-ss', sheetName: 'Current' }
  ];
  const history = [
    { name: '導入', activatedAt: '2026-05-14T00:00:00Z' },  // 削除済
    { name: '本時', activatedAt: '2026-05-14T01:00:00Z' }
  ];
  const ctx = ctxWithProfileHistory(history, profiles);
  const result = ctx.getPublishedSheetDataForProfile('u1', '導入');
  assert.equal(result.success, false);
  assert.match(result.error, /見つかりません|削除済/);
});

// =====================================================================
// getNotificationUpdate
// =====================================================================

test('getNotificationUpdate: rejects invalid request (no email or targetUserId)', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => null
  });
  assert.equal(ctx.getNotificationUpdate('u1', {}).success, false);

  const ctx2 = loadDataApisContext({
    getCurrentEmail: () => 'a@x.com'
  });
  assert.equal(ctx2.getNotificationUpdate(null, {}).success, false);
});

test('getNotificationUpdate: rejects unknown target user', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'a@x.com',
    findUserById: () => null
  });
  const result = ctx.getNotificationUpdate('ghost', {});
  assert.equal(result.success, false);
});

test('getNotificationUpdate: blocks non-admin non-owner on unpublished board', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'viewer@example.com',
    findUserById: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    getConfigOrDefault: () => ({ isPublished: false }),
    isAdministrator: () => false
  });
  const result = ctx.getNotificationUpdate('u1', {});
  assert.equal(result.success, false);
  assert.match(result.message, /Access denied/);
});

test('getNotificationUpdate: reports hasNewContent=false when no items newer than lastUpdateTime', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'viewer@example.com',
    findUserById: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    getConfigOrDefault: () => ({ isPublished: true }),
    getUserSheetData: () => ({
      success: true,
      data: [
        { timestamp: '2026-04-18T10:00:00Z' },
        { timestamp: '2026-04-18T11:00:00Z' }
      ]
    })
  });
  const result = ctx.getNotificationUpdate('u1', {
    lastUpdateTime: '2026-04-19T00:00:00Z'
  });
  assert.equal(result.success, true);
  assert.equal(result.hasNewContent, false);
  assert.equal(result.newItemsCount, 0);
});

test('getNotificationUpdate: reports hasNewContent=true when items newer than lastUpdateTime', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'viewer@example.com',
    findUserById: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    getConfigOrDefault: () => ({ isPublished: true }),
    getUserSheetData: () => ({
      success: true,
      data: [
        { timestamp: '2026-04-18T10:00:00Z' }, // older
        { timestamp: '2026-04-19T12:00:00Z' }, // newer
        { timestamp: '2026-04-19T13:00:00Z' }  // newer
      ]
    })
  });
  const result = ctx.getNotificationUpdate('u1', {
    lastUpdateTime: '2026-04-19T00:00:00Z'
  });
  assert.equal(result.success, true);
  assert.equal(result.hasNewContent, true);
  assert.equal(result.newItemsCount, 2);
});

// =====================================================================
// processFormUrlInput — URL validation (doesn't require FormApp)
// =====================================================================

test('processFormUrlInput: rejects non-string input', () => {
  const ctx = loadDataApisContext();
  assert.equal(ctx.processFormUrlInput(null).success, false);
  assert.equal(ctx.processFormUrlInput(42).success, false);
  assert.equal(ctx.processFormUrlInput('').success, false);
});

test('processFormUrlInput: rejects URL without /forms/d/ or forms.gle/', () => {
  const ctx = loadDataApisContext();
  assert.equal(ctx.processFormUrlInput('https://docs.google.com/spreadsheets/d/abc').success, false);
  assert.equal(ctx.processFormUrlInput('https://example.com/path').success, false);
});

test('processFormUrlInput: accepts /forms/d/ and /forms.gle/ URLs', () => {
  const ctx = loadDataApisContext({
    FormApp: {
      openByUrl: () => { throw new Error('dummy - we only test URL gate'); }
    }
  });
  // These pass URL gate but fail at FormApp.openByUrl with a specific error
  const r1 = ctx.processFormUrlInput('https://docs.google.com/forms/d/abc/viewform');
  assert.equal(r1.success, false);
  assert.match(r1.error, /アクセスできません/); // Got past URL gate, failed at FormApp

  const r2 = ctx.processFormUrlInput('https://forms.gle/abc');
  assert.equal(r2.success, false);
  assert.match(r2.error, /アクセスできません/);
});

// =====================================================================
// saveConfig
// =====================================================================

test('saveConfig: returns auth error without session', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => null
  });
  const result = ctx.saveConfig({}, {});
  assert.equal(result.success, false);
  assert.match(result.message, /認証/);
});

test('saveConfig: returns user-not-found when email unknown', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'ghost@example.com',
    findUserByEmail: () => null
  });
  const result = ctx.saveConfig({}, {});
  assert.equal(result.success, false);
  assert.match(result.message, /ユーザー/);
});

test('saveConfig: delegates to saveUserConfig with isMainConfig by default', () => {
  let capturedOpts = null;
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    findUserByEmail: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    saveUserConfig: (_userId, _config, opts) => {
      capturedOpts = opts;
      return { success: true };
    }
  });
  ctx.saveConfig({ spreadsheetId: 'ss' });
  assert.equal(capturedOpts.isMainConfig, true);
  assert.equal(capturedOpts.isDraft, undefined);
});

test('saveConfig: uses isDraft when options.isDraft is truthy', () => {
  let capturedOpts = null;
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    findUserByEmail: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    saveUserConfig: (_userId, _config, opts) => {
      capturedOpts = opts;
      return { success: true };
    }
  });
  ctx.saveConfig({ spreadsheetId: 'ss' }, { isDraft: true });
  assert.equal(capturedOpts.isDraft, true);
  assert.equal(capturedOpts.isMainConfig, undefined);
});

test('saveConfig: forwards userId and config to saveUserConfig unchanged', () => {
  let capturedUserId = null;
  let capturedConfig = null;
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    findUserByEmail: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    saveUserConfig: (userId, config) => {
      capturedUserId = userId;
      capturedConfig = config;
      return { success: true };
    }
  });
  const input = { spreadsheetId: 'ss', sheetName: 'Sheet1', columnMapping: { answer: 2 } };
  ctx.saveConfig(input);
  assert.equal(capturedUserId, 'u1');
  assert.equal(capturedConfig.spreadsheetId, 'ss');
  assert.equal(capturedConfig.columnMapping.answer, 2);
});

test('saveConfig: propagates saveUserConfig error', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    findUserByEmail: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    saveUserConfig: () => ({ success: false, message: 'validation failed: invalid sheetName' })
  });
  const result = ctx.saveConfig({});
  assert.equal(result.success, false);
  assert.match(result.message, /validation failed/);
});

test('saveConfig: returns exception response when saveUserConfig throws', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    findUserByEmail: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    saveUserConfig: () => { throw new Error('storage down'); }
  });
  const result = ctx.saveConfig({});
  assert.equal(result.success, false);
  assert.match(result.message, /storage down/);
});

// =====================================================================
// getActiveFormInfo
// =====================================================================

test('getActiveFormInfo: returns auth error without session', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => null
  });
  const result = ctx.getActiveFormInfo(null);
  assert.equal(result.success, false);
  assert.equal(result.formUrl, null);
  assert.equal(result.formTitle, null);
});

test('getActiveFormInfo: resolves current user when no userId given', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    findUserByEmail: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    getConfigOrDefault: () => ({
      formUrl: 'https://docs.google.com/forms/d/abc/viewform',
      formTitle: 'クラスアンケート'
    })
  });
  const result = ctx.getActiveFormInfo(null);
  assert.equal(result.success, true);
  assert.equal(result.shouldShow, true);
  assert.equal(result.formTitle, 'クラスアンケート');
  assert.equal(result.isValidUrl, true);
});

test('getActiveFormInfo: returns user-not-found when self-lookup fails', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'ghost@example.com',
    findUserByEmail: () => null
  });
  const result = ctx.getActiveFormInfo(null);
  assert.equal(result.success, false);
  assert.match(result.message, /User not found/);
});

test('getActiveFormInfo: uses explicit userId over current user', () => {
  let findByEmailCalled = false;
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'viewer@example.com',
    findUserByEmail: () => { findByEmailCalled = true; return null; },
    getConfigOrDefault: (uid) => {
      assert.equal(uid, 'other-user-id');
      return { formUrl: 'https://forms.gle/xyz', formTitle: 'Other' };
    }
  });
  const result = ctx.getActiveFormInfo('other-user-id');
  assert.equal(result.success, true);
  assert.equal(findByEmailCalled, false, 'Must not look up current user when id given');
});

test('getActiveFormInfo: reports shouldShow=false when formUrl absent', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    findUserByEmail: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    getConfigOrDefault: () => ({})
  });
  const result = ctx.getActiveFormInfo(null);
  assert.equal(result.success, false);
  assert.equal(result.shouldShow, false);
  assert.equal(result.formUrl, null);
});

test('getActiveFormInfo: reports isValidUrl=false when formUrl is non-Google URL', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    findUserByEmail: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    getConfigOrDefault: () => ({ formUrl: 'https://evil.example.com/fake-form' })
  });
  const result = ctx.getActiveFormInfo(null);
  assert.equal(result.shouldShow, true); // hasFormUrl=true → shouldShow
  assert.equal(result.isValidUrl, false); // but isValidFormUrl rejects it
});

test('getActiveFormInfo: whitespace-only formUrl treated as absent', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    findUserByEmail: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    getConfigOrDefault: () => ({ formUrl: '   ' })
  });
  const result = ctx.getActiveFormInfo(null);
  assert.equal(result.shouldShow, false);
});

test('getActiveFormInfo: default formTitle when absent from config', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    findUserByEmail: () => ({ userId: 'u1', userEmail: 'owner@example.com' }),
    getConfigOrDefault: () => ({ formUrl: 'https://docs.google.com/forms/d/abc/viewform' })
  });
  const result = ctx.getActiveFormInfo(null);
  assert.equal(result.formTitle, 'フォーム');
});

// =====================================================================
// connectDataSource / processDataSourceOperations
//
// These call same-file functions (getColumnAnalysis, getFormInfoInternal),
// so we monkey-patch the loaded context after construction rather than
// via harness overrides — vm.runInContext's function hoisting would
// otherwise clobber any overrides we pass in.
// =====================================================================

test('connectDataSource: rejects unauthenticated user', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => null
  });
  const result = ctx.connectDataSource('ss-1', 'Sheet1');
  assert.equal(result.success, false);
  assert.match(result.error, /認証/);
});

test('connectDataSource: 非所有者は SA 共有適用前に拒否 (副作用なし)', () => {
  let sharingAttempted = false;
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'stranger@example.com',
    isAdministrator: () => false,
    SpreadsheetApp: { openById: () => { throw new Error('PERMISSION_DENIED'); } },
    applySpreadsheetSharingDefaults: () => { sharingAttempted = true; }
  });
  const result = ctx.connectDataSource('someone-elses-ss', 'Sheet1');
  assert.equal(result.success, false);
  assert.equal(sharingAttempted, false, '権限チェックで弾き SA 共有を適用しない');
});

test('connectDataSource: applies SA pool sharing defaults but tolerates failure', () => {
  // v2782+: domain-wide sharing was replaced by SA pool editor add.
  let sharingAttempted = false;
  let analysisCalled = false;
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    applySpreadsheetSharingDefaults: () => {
      sharingAttempted = true;
      throw new Error('sharing denied');
    }
  });
  ctx.getColumnAnalysis = () => {
    analysisCalled = true;
    return { success: true, mapping: {}, headers: [] };
  };
  const result = ctx.connectDataSource('ss-1', 'Sheet1');
  assert.equal(sharingAttempted, true);
  assert.equal(analysisCalled, true, 'Sharing failure must not block the rest of the flow');
  assert.equal(result.success, true);
});

test('connectDataSource: delegates to getColumnAnalysis when no batch operations', () => {
  let capturedArgs = null;
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    applySpreadsheetSharingDefaults: () => ({ saAdded: true, saEmails: [], errors: [] })
  });
  ctx.getColumnAnalysis = (ss, sheet) => {
    capturedArgs = { ss, sheet };
    return { success: true };
  };
  ctx.connectDataSource('my-ss', 'my-sheet');
  assert.equal(capturedArgs.ss, 'my-ss');
  assert.equal(capturedArgs.sheet, 'my-sheet');
});

test('connectDataSource: delegates to processDataSourceOperations when batch given', () => {
  let columnCallCount = 0;
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    applySpreadsheetSharingDefaults: () => ({ saAdded: true, saEmails: [], errors: [] })
  });
  ctx.getColumnAnalysis = () => {
    columnCallCount += 1;
    return { success: true, mapping: {}, headers: [] };
  };
  ctx.getFormInfo = () => ({ formData: { formUrl: 'https://docs.google.com/forms/d/x' } });
  const result = ctx.connectDataSource('ss-1', 'Sheet1', [
    { type: 'validateAccess' },
    { type: 'getFormInfo' },
    { type: 'connectDataSource' }
  ]);
  assert.equal(result.success, true);
  assert.equal(columnCallCount, 1, 'Column analysis should be cached across ops');
  assert.ok(result.batchResults.validation);
  assert.ok(result.batchResults.formInfo);
});

test('processDataSourceOperations: reports failure in validation branch', () => {
  const ctx = loadDataApisContext();
  ctx.getColumnAnalysis = () => ({ success: false, message: 'permission denied' });
  const result = ctx.processDataSourceOperations('ss-1', 'Sheet1', [
    { type: 'validateAccess' }
  ]);
  assert.equal(result.batchResults.validation.success, false);
  assert.match(result.batchResults.validation.details.connectionError, /permission denied/);
});

test('processDataSourceOperations: marks overall success=false when connect op fails', () => {
  const ctx = loadDataApisContext();
  ctx.getColumnAnalysis = () => ({
    success: false,
    message: 'header integrity issue',
    errorResponse: { message: 'header integrity issue' }
  });
  const result = ctx.processDataSourceOperations('ss-1', 'Sheet1', [
    { type: 'connectDataSource' }
  ]);
  assert.equal(result.success, false);
  assert.match(result.error, /header integrity issue/);
});

test('processDataSourceOperations: unknown op type is silently skipped', () => {
  const ctx = loadDataApisContext();
  ctx.getColumnAnalysis = () => ({ success: true });
  const result = ctx.processDataSourceOperations('ss-1', 'Sheet1', [
    { type: 'unknownOp' }
  ]);
  assert.equal(result.success, true);
  assert.deepEqual({ ...result.batchResults }, {});
});

test('processDataSourceOperations: handles empty operations array', () => {
  const ctx = loadDataApisContext();
  const result = ctx.processDataSourceOperations('ss-1', 'Sheet1', []);
  assert.equal(result.success, true);
  assert.deepEqual({ ...result.batchResults }, {});
});

// =====================================================================
// getColumnAnalysis
// =====================================================================

test('getColumnAnalysis: rejects unauthenticated user', () => {
  const ctx = loadDataApisContext({ getCurrentEmail: () => null });
  const result = ctx.getColumnAnalysis('ss-1', 'Sheet1');
  assert.equal(result.success, false);
  assert.match(result.error, /認証/);
});

test('getColumnAnalysis: 非所有者 (openById 失敗) かつ非adminは拒否', () => {
  let openSpreadsheetCalled = false;
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'stranger@example.com',
    isAdministrator: () => false,
    SpreadsheetApp: { openById: () => { throw new Error('PERMISSION_DENIED'); } },
    openSpreadsheet: () => { openSpreadsheetCalled = true; return null; }
  });
  const result = ctx.getColumnAnalysis('someone-elses-ss', 'Sheet1');
  assert.equal(result.success, false);
  assert.match(result.message || result.error || '', /権限/);
  assert.equal(openSpreadsheetCalled, false, '権限チェックで弾き SA pool 読み取りに到達しない');
});

test('getColumnAnalysis: admin は openById 不可でも許可 (cross-user 分析)', () => {
  const sheet = {
    getLastRow: () => 1,
    getLastColumn: () => 1,
    getRange: () => ({ getValues: () => [['Q1']] }),
    getDataRange: () => ({ getValues: () => [['Q1']] })
  };
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'admin@example.com',
    isAdministrator: () => true,
    SpreadsheetApp: { openById: () => { throw new Error('PERMISSION_DENIED'); } },
    openSpreadsheet: () => ({ getSheet: () => sheet })
  });
  ctx.performIntegratedColumnDiagnostics = () => ({ recommendedMapping: { answer: 0 }, confidence: {} });
  ctx.setupReactionAndHighlightColumns = () => ({ columnsAdded: [] });
  const result = ctx.getColumnAnalysis('any-ss', 'Sheet1');
  assert.equal(result.success, true);
});

test('getColumnAnalysis: returns error when openSpreadsheet returns null', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    openSpreadsheet: () => null
  });
  const result = ctx.getColumnAnalysis('ss-1', 'Sheet1');
  assert.equal(result.success, false);
  assert.match(result.error, /アクセスできません/);
});

test('getColumnAnalysis: returns error when openSpreadsheet throws', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    openSpreadsheet: () => { throw new Error('denied'); }
  });
  const result = ctx.getColumnAnalysis('ss-1', 'Sheet1');
  assert.equal(result.success, false);
  assert.match(result.error, /アクセス権がありません/);
});

test('getColumnAnalysis: returns sheet-not-found when getSheet returns null', () => {
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    openSpreadsheet: () => ({ getSheet: () => null })
  });
  const result = ctx.getColumnAnalysis('ss-1', 'Missing');
  assert.equal(result.success, false);
  assert.match(result.message, /Sheet not found/);
});

test('getColumnAnalysis: returns headers and mapping on success', () => {
  const headers = ['Q1', '理由', '名前'];
  const sheet = {
    getLastRow: () => 1,
    getLastColumn: () => headers.length,
    getRange: (row, col, numRows, numCols) => ({
      getValues: () => (numRows === 1 && row === 1 ? [headers] : [])
    }),
    getDataRange: () => ({ getValues: () => [headers] })
  };
  const ctx = loadDataApisContext({
    getCurrentEmail: () => 'owner@example.com',
    openSpreadsheet: () => ({ getSheet: () => sheet })
  });
  // Stub performIntegratedColumnDiagnostics via in-context patching to avoid
  // requiring the real ColumnMappingService in this test.
  ctx.performIntegratedColumnDiagnostics = () => ({
    recommendedMapping: { answer: 0, reason: 1, name: 2 },
    confidence: { answer: 90 }
  });
  // Suppress column-setup side effect path by forcing no columns added
  ctx.setupReactionAndHighlightColumns = () => ({ columnsAdded: [] });

  const result = ctx.getColumnAnalysis('ss-1', 'Sheet1');
  assert.equal(result.success, true);
  assert.equal(result.mapping.answer, 0);
  assert.ok(result.headers.length === 3);
});

// =====================================================================
// buildSafePublishedDataResult — privacy: identity fields must not leak
// =====================================================================

function buildRow(overrides = {}) {
  return {
    id: 'row_2',
    rowIndex: 2,
    answer: 'A',
    opinion: 'A',
    reason: 'because',
    class: '1-1',
    name: 'Student A',
    email: 'student@example.com',
    reactions: {},
    highlight: false,
    ...overrides
  };
}

test('buildSafePublishedDataResult: strips email/name for student viewer when showNames=false', () => {
  const ctx = loadDataApisContext();
  const out = ctx.buildSafePublishedDataResult(
    { data: [buildRow()], header: 'H', sheetName: 'S' },
    { displaySettings: { showNames: false } },
    { isAdmin: false, isOwnBoard: false }
  );
  assert.equal(out.data.length, 1);
  assert.equal(out.data[0].email, undefined);
  assert.equal(out.data[0].name, undefined);
  // Non-identity fields survive
  assert.equal(out.data[0].answer, 'A');
  assert.equal(out.data[0].reason, 'because');
});

test('buildSafePublishedDataResult: keeps email/name for admin viewer', () => {
  const ctx = loadDataApisContext();
  const out = ctx.buildSafePublishedDataResult(
    { data: [buildRow()], header: 'H', sheetName: 'S' },
    { displaySettings: { showNames: false } },
    { isAdmin: true, isOwnBoard: false }
  );
  assert.equal(out.data[0].email, 'student@example.com');
  assert.equal(out.data[0].name, 'Student A');
});

test('buildSafePublishedDataResult: keeps email/name for board owner', () => {
  const ctx = loadDataApisContext();
  const out = ctx.buildSafePublishedDataResult(
    { data: [buildRow()], header: 'H', sheetName: 'S' },
    { displaySettings: { showNames: false } },
    { isAdmin: false, isOwnBoard: true }
  );
  assert.equal(out.data[0].email, 'student@example.com');
  assert.equal(out.data[0].name, 'Student A');
});

test('buildSafePublishedDataResult: keeps email/name when showNames=true (named mode)', () => {
  const ctx = loadDataApisContext();
  const out = ctx.buildSafePublishedDataResult(
    { data: [buildRow()], header: 'H', sheetName: 'S' },
    { displaySettings: { showNames: true } },
    { isAdmin: false, isOwnBoard: false }
  );
  assert.equal(out.data[0].name, 'Student A');
  assert.equal(out.data[0].email, 'student@example.com');
});

test('buildSafePublishedDataResult: default viewerContext (omitted) treated as non-privileged', () => {
  const ctx = loadDataApisContext();
  const out = ctx.buildSafePublishedDataResult(
    { data: [buildRow()], header: 'H', sheetName: 'S' },
    { displaySettings: { showNames: false } }
  );
  assert.equal(out.data[0].email, undefined);
  assert.equal(out.data[0].name, undefined);
});

test('buildSafePublishedDataResult: converts Date timestamps to ISO strings', () => {
  const ctx = loadDataApisContext();
  // Why vm.runInContext: Date constructors differ across vm contexts, so
  //   `v instanceof Date` inside the source (which uses the context's Date)
  //   is false for a Date built out here. Build the Date inside the context.
  const ts = vm.runInContext("new Date('2026-04-22T10:00:00Z')", ctx);
  const out = ctx.buildSafePublishedDataResult(
    { data: [buildRow({ timestamp: ts })], header: 'H', sheetName: 'S' },
    { displaySettings: { showNames: true } },
    { isAdmin: true }
  );
  assert.equal(out.data[0].timestamp, '2026-04-22T10:00:00.000Z');
});

// =====================================================================
// setupReactionAndHighlightColumns — column detection is exact-match
// =====================================================================

test('setupReactionAndHighlightColumns: detects all four columns when present exactly', () => {
  const sheet = {
    getRange: () => ({ setValues: () => {}, setValue: () => {} })
  };
  const ctx = loadDataApisContext({
    openSpreadsheet: () => ({ spreadsheet: { getSheetByName: () => sheet } }),
    invalidateSheetHeadersCache: () => {}
  });
  const result = ctx.setupReactionAndHighlightColumns('ss-1', 'Sheet1',
    ['Q1', 'UNDERSTAND', 'LIKE', 'CURIOUS', 'HIGHLIGHT']
  );
  assert.equal(result.success, true);
  assert.equal(result.columnsAdded.length, 0);
  assert.equal(result.alreadyExists, 4);
});

test('setupReactionAndHighlightColumns: "UNDERSTANDING YOUR ANSWER" does NOT count as UNDERSTAND', () => {
  // Why regression: a previous substring check (header.includes("UNDERSTAND"))
  // treated a question header "UNDERSTANDING YOUR ANSWER" as if the UNDERSTAND
  // column already existed, and skipped adding it. processReactionDirect then
  // failed with "リアクション列が見つかりません" because it uses exact match.
  const addedValues = [];
  const sheetHeaders = ['Q1', 'UNDERSTANDING YOUR ANSWER'];
  const sheet = {
    // Why getLastColumn / getValues for header re-read: v2760 で lock 取得後に fresh re-read
    //   する double-add 防止ロジックが入った。test mock も headers を返す必要がある。
    getLastColumn: () => sheetHeaders.length,
    getRange: () => ({
      setValues: (vals) => { addedValues.push(...vals[0]); },
      setValue: (v) => { addedValues.push(v); },
      getValues: () => [sheetHeaders]
    })
  };
  const ctx = loadDataApisContext({
    openSpreadsheet: () => ({ spreadsheet: { getSheetByName: () => sheet } }),
    getSheetInfo: () => ({ lastCol: 2, lastRow: 1, headers: [] }),
    invalidateSheetHeadersCache: () => {},
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) }
  });
  const result = ctx.setupReactionAndHighlightColumns('ss-1', 'Sheet1', sheetHeaders);
  assert.equal(result.success, true, JSON.stringify(result));
  // All four reaction/highlight columns should be added because the header
  // "UNDERSTANDING YOUR ANSWER" is not a strict match for UNDERSTAND.
  const added = Array.from(addedValues);
  assert.deepEqual(added, ['UNDERSTAND', 'LIKE', 'CURIOUS', 'HIGHLIGHT']);
});

test('setupReactionAndHighlightColumns: detects lowercase/whitespace-padded headers', () => {
  const sheet = {
    getRange: () => ({ setValues: () => {}, setValue: () => {} })
  };
  const ctx = loadDataApisContext({
    SpreadsheetApp: {
      openById: () => ({ getSheetByName: () => sheet })
    },
    invalidateSheetHeadersCache: () => {}
  });
  const result = ctx.setupReactionAndHighlightColumns('ss-1', 'Sheet1',
    ['Q1', '  understand ', 'Like', 'CURIOUS', ' highlight']
  );
  assert.equal(result.columnsAdded.length, 0);
});
