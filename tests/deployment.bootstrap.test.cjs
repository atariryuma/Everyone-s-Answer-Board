/**
 * 配置の初回の準備 (ensureDeploymentBootstrap_) のテスト。
 *
 * Why: 先生ごとに「先生として実行」で配置する (scripts/deploy-single.js が Deployment.js に
 *   その先生のメールと DB を焼き込む)。管理者と DB は焼き込んだ値だけで決まり、最初に開いた人
 *   では決まらない (児童が先に開いても児童が管理者にならない)。
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const read = (f) => fs.readFileSync(path.resolve(__dirname, '../src', f), 'utf8');

// ---------------------------------------------------------------------
// ensureDeploymentBootstrap_: 初回の準備
// ---------------------------------------------------------------------

function loadHelpers({ admin = 'Teacher@Example.com', db = 'teacher-ss', props = {} } = {}) {
  const ensured = [];
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    PROPERTY_CACHE_TTL: 0,
    PropertiesService: { getScriptProperties: () => ({
      getProperty: (k) => (k in props ? props[k] : null),
      setProperty: (k, v) => { props[k] = v; }
    }) },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
    SpreadsheetApp: { openById: (id) => ({ id }) },
    __ensureSheetWithHeaders_: (ss, name) => { ensured.push(ss.id + ':' + name); },
    USERS_SHEET_HEADERS: [], LESSONS_SHEET_HEADERS: [], LESSON_RESPONSES_SHEET_HEADERS: []
  };
  vm.createContext(context);
  // Deployment.js の代わり (配置時に差し替えられる値)。
  vm.runInContext(`const DEPLOY_ADMIN_EMAIL = ${JSON.stringify(admin)};
    const DEPLOY_DATABASE_SPREADSHEET_ID = ${JSON.stringify(db)};`, context);
  vm.runInContext(read('helpers.js'), context, { filename: 'helpers.js' });
  return { context, props, ensured };
}

test('初回の準備: 焼き込んだ管理者と DB でプロパティとシートをそろえる', () => {
  const { context, props, ensured } = loadHelpers();
  context.ensureDeploymentBootstrap_();
  assert.equal(props.ADMIN_EMAIL, 'teacher@example.com', '小文字にそろえる');
  assert.equal(props.DATABASE_SPREADSHEET_ID, 'teacher-ss');
  assert.deepEqual(ensured, ['teacher-ss:users', 'teacher-ss:lessons', 'teacher-ss:lesson_responses']);
});

test('初回の準備: 既にあるプロパティは上書きしない (先生が後から変えた値を守る)', () => {
  const { context, props, ensured } = loadHelpers({
    props: { ADMIN_EMAIL: 'other@example.com', DATABASE_SPREADSHEET_ID: 'other-ss' }
  });
  context.ensureDeploymentBootstrap_();
  assert.equal(props.ADMIN_EMAIL, 'other@example.com');
  assert.equal(props.DATABASE_SPREADSHEET_ID, 'other-ss');
  assert.deepEqual(ensured, []);
});

test('焼き込みが空 (リポジトリの Deployment.js のまま): 何もしない', () => {
  const { context, props, ensured } = loadHelpers({ admin: '', db: '' });
  context.ensureDeploymentBootstrap_();
  assert.deepEqual(props, {});
  assert.deepEqual(ensured, []);
});

test('getCachedProperty_: 管理者と DB は焼き込みから返し、プロパティを読まない (児童の確認ごとに読まない)', () => {
  const reads = [];
  const { context } = loadHelpers({ props: { ADMIN_EMAIL: 'old@example.com', OTHER: 'x' } });
  const orig = context.PropertiesService.getScriptProperties;
  context.PropertiesService.getScriptProperties = () => {
    const p = orig();
    return Object.assign({}, p, { getProperty: (k) => { reads.push(k); return p.getProperty(k); } });
  };
  assert.equal(context.getCachedProperty_('ADMIN_EMAIL'), 'teacher@example.com', '焼き込みが先に効く (小文字)');
  assert.equal(context.getCachedProperty_('DATABASE_SPREADSHEET_ID'), 'teacher-ss');
  assert.deepEqual(reads, []);
  assert.equal(context.getCachedProperty_('OTHER'), 'x', 'ほかの鍵はプロパティから');
});
