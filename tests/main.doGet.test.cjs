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

test('doGet: manual mode loads TeacherManual template', () => {
  const { context, templates } = loadDoGetContext();
  context.doGet(createGetEvent({ mode: 'manual' }));
  assert.ok(templates.includes('TeacherManual.html'));
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
