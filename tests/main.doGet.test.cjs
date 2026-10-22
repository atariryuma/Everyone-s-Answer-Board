const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('path');
const vm = require('vm');

function createHtmlOutput(title) {
  return {
    _title: title || '',
    setTitle(t) { this._title = t; return this; },
    setXFrameOptionsMode() { return this; },
    addMetaTag() { return this; }
  };
}

function loadDoGetContext(overrides = {}) {
  const templates = [];
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    logError_: () => {},
    HtmlService: {
      createTemplateFromFile: (name) => {
        templates.push(name);
        return {
          evaluate: () => createHtmlOutput(name),
          title: null, message: null, hideLoginButton: null,
          userId: null, configJSON: null, mapping: null,
          spreadsheetId: null, sheetName: null, editorName: null,
          isEditor: null, boardOwnerName: null, isPublished: null,
          userEmail: null, isAdministrator: null, timestamp: null,
          questionText: null, formUrl: null
        };
      },
      XFrameOptionsMode: { ALLOWALL: 'ALLOWALL' }
    },
    Date: Date,
    JSON: JSON,
    getCurrentEmail: () => 'user@example.com',
    isAdministrator: (email) => email === 'admin@example.com',
    evaluateDomainRestriction: () => ({ allowed: true }),
    checkAppAccessRestriction: () => false,
    hasCoreSystemProps: () => true,
    getWebAppUrl: () => 'https://example.com/exec',
    getCachedProperty: (k) => k === 'ADMIN_EMAIL' ? 'admin@example.com' : null,
    shouldEnforceDomainRestrictions: () => true,
    validateDomainAccess: () => ({ allowed: true }),
    getBatchedAdminData: () => ({ success: false, error: 'test error' }),
    getBatchedViewerData: () => ({ success: false, error: 'test error' }),
    // main.js は DataApis.js の sameEmail_ を参照する (GAS は単一グローバルスコープ)
    sameEmail_: (a, b) => String(a || '').toLowerCase().trim() === String(b || '').toLowerCase().trim(),
    enhanceConfigWithDynamicUrls: (c) => c,
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (k) => k === 'ADMIN_EMAIL' ? 'admin@example.com' : 'value'
      })
    },
    ...overrides
  };
  vm.createContext(context);
  const source = fs.readFileSync(path.resolve(__dirname, '../src/main.js'), 'utf8');
  vm.runInContext(source, context, { filename: 'main.js' });
  return { context, templates };
}

function createGetEvent(params = {}) {
  return { parameter: params };
}

// --- Public modes (no auth required) ---

test('doGet: login mode loads LoginPage template', () => {
  const { context, templates } = loadDoGetContext();
  context.doGet(createGetEvent({ mode: 'login' }));
  assert.ok(templates.includes('LoginPage.html'));
});


// --- mode=admin error paths ---

test('doGet: admin without auth returns error page', () => {
  const { context, templates } = loadDoGetContext({ getCurrentEmail: () => null });
  context.doGet(createGetEvent({ mode: 'admin', userId: 'u1' }));
  assert.ok(templates.includes('ErrorBoundary.html'));
});

test('doGet: admin without userId returns error page', () => {
  const { context, templates } = loadDoGetContext();
  context.doGet(createGetEvent({ mode: 'admin' }));
  assert.ok(templates.includes('ErrorBoundary.html'));
});

test('doGet: admin with failed data fetch returns error page', () => {
  const { context, templates } = loadDoGetContext();
  context.doGet(createGetEvent({ mode: 'admin', userId: 'u1' }));
  assert.ok(templates.includes('ErrorBoundary.html'));
});

// --- mode=view error paths ---

test('doGet: view without auth returns error page', () => {
  const { context, templates } = loadDoGetContext({ getCurrentEmail: () => null });
  context.doGet(createGetEvent({ mode: 'view', userId: 'u1' }));
  assert.ok(templates.includes('ErrorBoundary.html'));
});

test('doGet: view without userId returns error page', () => {
  const { context, templates } = loadDoGetContext();
  context.doGet(createGetEvent({ mode: 'view' }));
  assert.ok(templates.includes('ErrorBoundary.html'));
});

test('doGet: view with failed data fetch returns error page', () => {
  const { context, templates } = loadDoGetContext();
  context.doGet(createGetEvent({ mode: 'view', userId: 'u1' }));
  assert.ok(templates.includes('ErrorBoundary.html'));
});

// --- mode=appSetup ---

test('doGet: appSetup without auth returns error page', () => {
  const { context, templates } = loadDoGetContext({ getCurrentEmail: () => null });
  context.doGet(createGetEvent({ mode: 'appSetup' }));
  assert.ok(templates.includes('ErrorBoundary.html'));
});

test('doGet: appSetup with non-admin returns error page', () => {
  const { context, templates } = loadDoGetContext();
  context.doGet(createGetEvent({ mode: 'appSetup' }));
  assert.ok(templates.includes('ErrorBoundary.html'));
});

// --- default mode ---

test('doGet: default mode returns AccessRestricted', () => {
  const { context, templates } = loadDoGetContext();
  context.doGet(createGetEvent({}));
  assert.ok(templates.includes('AccessRestricted.html'));
});

// --- domain restriction ---

test('doGet: domain restriction blocks all modes', () => {
  const { context, templates } = loadDoGetContext({
    shouldEnforceDomainRestrictions: () => true,
    validateDomainAccess: () => ({ allowed: false, reason: 'domain_mismatch' })
  });
  context.doGet(createGetEvent({ mode: 'login' }));
  assert.ok(templates.includes('AccessRestricted.html'));
});

// --- app disabled ---

test('doGet: disabled app blocks non-admin access', () => {
  const { context, templates } = loadDoGetContext({
    checkAppAccessRestriction: () => true
  });
  context.doGet(createGetEvent({ mode: 'login' }));
  assert.ok(templates.includes('AccessRestricted.html'));
});

// --- null event ---

test('doGet: null event defaults to main mode', () => {
  const { context, templates } = loadDoGetContext();
  context.doGet(null);
  assert.ok(templates.includes('AccessRestricted.html'));
});

// --- 未公開ボードの扱い (v2895) ---
//
// 公開前でも編集者は実物を確認できる必要がある (確認のために一度児童へ公開する、
// という運用を避ける)。児童には従来どおり「準備中」を返す。

// main.js 内で定義される関数 (getCurrentEmail / getBatchedViewerData) は context override では
// 上書きできないので、その依存を差し替える。
function viewerStubs(email, { isPublished = false, isActive = true, isAdmin = false, isCollaborator = false } = {}) {
  return {
    Session: { getActiveUser: () => ({ getEmail: () => email }) },
    isAdministrator: (e) => isAdmin && e === email,
    findPublishedBoardOwner: () => ({ userId: 'u1', userEmail: 'owner@example.com', isActive }),
    getConfigOrDefault: () => ({ isPublished, sheetName: 'Sheet1', displaySettings: {} }),
    isBoardCollaborator: () => isCollaborator,
    DEFAULT_DISPLAY_SETTINGS: {},
    VALIDATOR_BOARD_MODES: ['board']
  };
}

test('doGet view: 未公開ボードを児童が開くと「準備中」ページ', () => {
  const { context, templates } = loadDoGetContext(viewerStubs('child@example.com'));
  context.doGet(createGetEvent({ mode: 'view', userId: 'u1' }));
  assert.ok(templates.includes('Unpublished.html'));
  assert.ok(!templates.includes('Page.html'));
});

test('doGet view: 未公開でも所有者にはボードを描く (プレビュー)', () => {
  const { context, templates } = loadDoGetContext(viewerStubs('owner@example.com'));
  context.doGet(createGetEvent({ mode: 'view', userId: 'u1' }));
  assert.ok(templates.includes('Page.html'));
  assert.ok(!templates.includes('Unpublished.html'));
});

test('doGet view: 未公開でも管理者にはボードを描く', () => {
  const { context, templates } = loadDoGetContext(viewerStubs('admin@example.com', { isAdmin: true }));
  context.doGet(createGetEvent({ mode: 'view', userId: 'u1' }));
  assert.ok(templates.includes('Page.html'));
});

test('doGet view: 未公開でも共同教師にはボードを描く', () => {
  const { context, templates } = loadDoGetContext(viewerStubs('colleague@example.com', { isCollaborator: true }));
  context.doGet(createGetEvent({ mode: 'view', userId: 'u1' }));
  assert.ok(templates.includes('Page.html'));
});

test('doGet view: 無効化されたユーザーのボードは所有者以外に「準備中」', () => {
  const { context, templates } = loadDoGetContext(viewerStubs('child@example.com', { isPublished: true, isActive: false }));
  context.doGet(createGetEvent({ mode: 'view', userId: 'u1' }));
  assert.ok(templates.includes('Unpublished.html'));
});

// =====================================================================
// 許可が一部だけの利用者 (granular consent) に、許可し直すページを出す (2026-10-21 の児童 2 人)
// =====================================================================

function scriptAppStub(status, url) {
  return {
    AuthMode: { FULL: 'FULL' },
    AuthorizationStatus: { REQUIRED: 'REQUIRED', NOT_REQUIRED: 'NOT_REQUIRED' },
    getAuthorizationInfo: () => ({ getAuthorizationStatus: () => status, getAuthorizationUrl: () => url })
  };
}

function htmlServiceWithOutput(captured, templates) {
  return {
    createTemplateFromFile: (name) => { templates.push(name); return { evaluate: () => ({ setTitle() { return this; }, addMetaTag() { return this; }, setXFrameOptionsMode() { return this; } }) }; },
    createHtmlOutput: (html) => { captured.html = html; return { setTitle() { return this; } }; },
    XFrameOptionsMode: { ALLOWALL: 'ALLOWALL' }
  };
}

test('doGet: 許可が足りないと、ボードではなく「許可し直す」ページを出す', () => {
  const captured = {}; const templates = [];
  const { context } = loadDoGetContext({
    ScriptApp: scriptAppStub('REQUIRED', 'https://accounts.example/auth?x=1&y=2'),
    HtmlService: htmlServiceWithOutput(captured, templates)
  });
  context.doGet(createGetEvent({ mode: 'view', userId: 'u1' }));
  assert.match(captured.html || '', /許可し直す/);
  assert.match(captured.html, /https:\/\/accounts\.example\/auth\?x=1&amp;y=2/, 'URL は escape して埋める');
  assert.match(captured.html, /https:\/\/example\.com\/exec\?mode=view&amp;userId=u1/, '読み込み直すと同じボードへ');
  assert.equal(templates.length, 0, 'ボードのテンプレートは評価しない');
});

test('doGet: 許可がそろっていれば従来どおり (再許可ページを出さない)', () => {
  const captured = {}; const templates = [];
  const { context } = loadDoGetContext({
    ScriptApp: scriptAppStub('NOT_REQUIRED', ''),
    HtmlService: htmlServiceWithOutput(captured, templates)
  });
  context.doGet(createGetEvent({ mode: 'login' }));
  assert.equal(captured.html, undefined);
  assert.ok(templates.includes('LoginPage.html'));
});

test('getMissingAuthorization_: 判定が例外になったら「足りている」に倒す (全員を止めない)', () => {
  const { context } = loadDoGetContext({
    ScriptApp: { AuthMode: { FULL: 'FULL' }, AuthorizationStatus: { REQUIRED: 'REQUIRED' }, getAuthorizationInfo: () => { throw new Error('boom'); } }
  });
  assert.equal(context.getMissingAuthorization_(), null);
});

test('getMissingAuthorization_: ScriptApp.getAuthorizationInfo が無い環境でも null', () => {
  const { context } = loadDoGetContext({ ScriptApp: {} });
  assert.equal(context.getMissingAuthorization_(), null);
});
