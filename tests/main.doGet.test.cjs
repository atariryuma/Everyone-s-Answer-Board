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
  const templateObjects = [];
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    logError_: () => {},
    HtmlService: {
      createTemplateFromFile: (name) => {
        templates.push(name);
        const tpl = {
          evaluate: () => createHtmlOutput(name),
          title: null, message: null,
          userId: null, configJSON: null, mapping: null,
          spreadsheetId: null, sheetName: null, editorName: null,
          isEditor: null, boardOwnerName: null, isPublished: null,
          userEmail: null, timestamp: null,
          questionText: null, formUrl: null
        };
        templateObjects.push({ name, tpl });
        return tpl;
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
  return { context, templates, templateObjects };
}

function createGetEvent(params = {}) {
  return { parameter: params };
}

// --- 廃止した画面 (ログイン・初期設定・アプリ設定) ---
//   ルートごと消したので、古いブックマークの ?mode=login / setup / appSetup は
//   未知の mode と同じく main (トップ) と同じ振る舞いになる。

for (const mode of ['login', 'setup', 'appSetup']) {
  test(`doGet: mode=${mode} は廃止。トップ (main) と同じ案内を返す`, () => {
    const { context, templates } = loadDoGetContext();
    context.doGet(createGetEvent({ mode }));
    assert.deepEqual(templates, ['AccessRestricted.html']);
  });

  test(`doGet: mode=${mode} を管理者が開いても、廃止した画面は出ない`, () => {
    const { context, templates } = loadDoGetContext({
      Session: { getActiveUser: () => ({ getEmail: () => 'admin@example.com' }) }
    });
    context.doGet(createGetEvent({ mode }));
    assert.deepEqual(templates, ['AccessRestricted.html']);
  });
}

test('doGet: 削除したルートは handler table にも残っていない', () => {
  const { context } = loadDoGetContext();
  const modes = vm.runInContext('Object.keys(DO_GET_HANDLERS)', context);
  assert.deepEqual([...modes].sort(), ['admin', 'main', 'review', 'view']);
});

test('doGet: トップの案内は users を引かない (ボタンの出し分けが無くなったため)', () => {
  let lookups = 0;
  const { context, templates } = loadDoGetContext({
    // getCurrentEmail は main.js 内で定義し直されるので、元の Session を差し替える。
    Session: { getActiveUser: () => ({ getEmail: () => 'child@example.com' }) },
    findUserByEmail: () => { lookups++; return null; }
  });
  context.doGet(createGetEvent({}));
  assert.ok(templates.includes('AccessRestricted.html'));
  assert.equal(lookups, 0);
});

// --- 必須プロパティ未設定 ---

test('doGet: 必須プロパティがそろっていなければ、初期設定画面ではなく準備中の案内を返す', () => {
  const { context, templates, templateObjects } = loadDoGetContext({ hasCoreSystemProps: () => false });
  for (const mode of ['main', 'admin', 'view', 'setup']) {
    templates.length = 0;
    templateObjects.length = 0;
    context.doGet(createGetEvent({ mode, userId: 'u1' }));
    assert.deepEqual(templates, ['AccessRestricted.html'], `mode=${mode}`);
    assert.equal(templateObjects[0].tpl.message, 'このアプリはまだ準備できていません。管理者に連絡してください。');
  }
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
  context.doGet(createGetEvent({ mode: 'admin', userId: 'u1' }));
  assert.deepEqual(templates, ['AccessRestricted.html']);
});

// --- app disabled ---

test('doGet: disabled app blocks non-admin access', () => {
  const { context, templates, templateObjects } = loadDoGetContext({
    checkAppAccessRestriction: () => true
  });
  context.doGet(createGetEvent({ mode: 'view', userId: 'u1' }));
  assert.deepEqual(templates, ['AccessRestricted.html']);
  assert.equal(templateObjects[0].tpl.isAppDisabled, true);
});

test('doGet: 停止中は管理者も案内だけ (アプリ設定画面の例外は廃止。再開は Admin API の enableApp)', () => {
  const { context, templates } = loadDoGetContext({
    Session: { getActiveUser: () => ({ getEmail: () => 'admin@example.com' }) },
    checkAppAccessRestriction: () => true
  });
  for (const mode of ['appSetup', 'admin', 'main']) {
    templates.length = 0;
    context.doGet(createGetEvent({ mode, userId: 'u1' }));
    assert.deepEqual(templates, ['AccessRestricted.html'], `mode=${mode}`);
  }
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

// --- 先生専用の配置 (isSingleTenant_): トップを先生が開いたら管理パネルを直接開く ---

function loadSingleTenantMain({ singleTenant, email, loginOk = true }) {
  const loginCalls = [];
  const { context, templates } = loadDoGetContext({
    // getCurrentEmail は main.js 内で定義し直されるので、元の Session を差し替える。
    Session: { getActiveUser: () => ({ getEmail: () => email }) },
    isSingleTenant_: () => singleTenant,
    processLoginAction: () => {
      loginCalls.push(email);
      return loginOk
        ? { success: true, data: { userId: 'teacher-uid', email } }
        : { success: false, message: 'ユーザー登録に失敗しました' };
    },
    // getBatchedAdminData は main.js の実物が動くので、その下の data 層だけ用意する。
    findUserById: (uid) => ({ userId: uid, userEmail: email, isActive: true }),
    getConfigOrDefault: () => ({}),
    getQuestionText: () => '',
    VALIDATOR_BOARD_MODES: ['auto', 'matrix'],
    findUserByEmail: () => null
  });
  return { context, templates, loginCalls };
}

test('doGet 先生専用: 先生がトップを開くと、登録を済ませて管理パネルを返す (ログイン・アプリ設定の入口を出さない)', () => {
  const h = loadSingleTenantMain({ singleTenant: true, email: 'admin@example.com' });
  h.context.doGet(createGetEvent({}));
  assert.deepEqual(h.loginCalls, ['admin@example.com']);
  assert.ok(h.templates.includes('AdminPanel.html'), h.templates.join(','));
  assert.ok(!h.templates.includes('AccessRestricted.html'));
});

test('doGet 先生専用: 児童がトップを開いても登録はせず、本体と同じ案内のまま', () => {
  const h = loadSingleTenantMain({ singleTenant: true, email: 'child@example.com' });
  h.context.doGet(createGetEvent({}));
  assert.deepEqual(h.loginCalls, []);
  assert.ok(h.templates.includes('AccessRestricted.html'));
});

test('doGet 本体 (みんなで 1 つ): 管理者がトップを開いても今までどおりの画面', () => {
  const h = loadSingleTenantMain({ singleTenant: false, email: 'admin@example.com' });
  h.context.doGet(createGetEvent({}));
  assert.deepEqual(h.loginCalls, []);
  assert.ok(h.templates.includes('AccessRestricted.html'));
});

test('doGet 先生専用: 古いブックマークの ?mode=login を先生が開いても管理パネルを返す', () => {
  const h = loadSingleTenantMain({ singleTenant: true, email: 'admin@example.com' });
  h.context.doGet(createGetEvent({ mode: 'login' }));
  assert.deepEqual(h.loginCalls, ['admin@example.com']);
  assert.ok(h.templates.includes('AdminPanel.html'), h.templates.join(','));
});

test('doGet 先生専用: 登録に失敗したら案内の画面に戻す', () => {
  const h = loadSingleTenantMain({ singleTenant: true, email: 'admin@example.com', loginOk: false });
  h.context.doGet(createGetEvent({}));
  assert.ok(h.templates.includes('AccessRestricted.html'));
  assert.ok(!h.templates.includes('AdminPanel.html'));
});
