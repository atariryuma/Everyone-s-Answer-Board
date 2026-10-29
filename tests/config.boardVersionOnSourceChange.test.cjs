/**
 * saveUserConfig_: ボードの読む先 (シート / 列の割り当て) が変わったら、ボードの版を上げる。
 *
 * Why: ボードの読み取り cache の key はボードの版で、シートを含まない。フェーズ切替やフォームの
 *   切替で別のシートを指しても版が変わらず、最大 12 秒は前のシートの行が返っていた。それを読み込んだ
 *   児童の画面は、次の投稿が来るまで前のフェーズの回答を新しいフェーズの軸で映し続けた。
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function loadCtx(persisted, { unchanged = false } = {}) {
  const bumps = [];
  const user = { userId: 'u1', userEmail: 'owner@example.com', configJson: JSON.stringify(persisted) };
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    CacheService: { getScriptCache: () => ({ removeAll: () => {}, remove: () => {}, get: () => null, put: () => {} }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null, setProperty: () => {} }) },
    SYSTEM_LIMITS: { DEFAULT_PAGE_SIZE: 20, MAX_PAGE_SIZE: 100, CONFIG_JSON_MAX_CHARS: 32768, RADIX_DECIMAL: 10 },
    DEFAULT_DISPLAY_SETTINGS: { showNames: false, showReactions: true, theme: 'default', pageSize: 20 },
    Utilities: { getUuid: () => 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee' },
    logError_: () => {},
    getCachedProperty_: () => null,
    getCurrentEmail: () => 'owner@example.com',
    isAdministrator: () => false,
    findUserById_: () => user,
    validateConfig: (cfg) => ({ isValid: true, errors: [], warnings: [], sanitized: { ...cfg } }),
    updateUser_: () => (unchanged ? { success: true, unchanged: true, currentConfig: persisted } : { success: true }),
    bumpBoardDataVersion_: (id) => { bumps.push(id); }
  };
  vm.createContext(context);
  vm.runInContext(fs.readFileSync(path.resolve(__dirname, '../src/ConfigService.js'), 'utf8'), context, { filename: 'ConfigService.js' });
  return { ctx: context, bumps };
}

const BASE = { spreadsheetId: 'ss1', sheetName: '考える', columnMapping: { answer: 4, reason: 5 }, displaySettings: { showNames: false } };

test('シートが変わったら版を上げる (フェーズ切替で別の回答シートへ)', () => {
  const { ctx, bumps } = loadCtx(BASE);
  const res = ctx.saveUserConfig_('u1', { ...BASE, sheetName: 'もう一度考える' });
  assert.equal(res.success, true);
  assert.deepEqual(bumps, ['u1']);
});

test('スプレッドシートや列の割り当てが変わっても版を上げる', () => {
  const a = loadCtx(BASE);
  a.ctx.saveUserConfig_('u1', { ...BASE, spreadsheetId: 'ss2' });
  assert.deepEqual(a.bumps, ['u1']);
  const b = loadCtx(BASE);
  b.ctx.saveUserConfig_('u1', { ...BASE, columnMapping: { answer: 4, reason: 6 } });
  assert.deepEqual(b.bumps, ['u1']);
});

test('読む先が同じなら版を上げない (表示設定だけの保存で児童全員の cache を捨てない)', () => {
  const { ctx, bumps } = loadCtx(BASE);
  const res = ctx.saveUserConfig_('u1', { ...BASE, displaySettings: { showNames: true } });
  assert.equal(res.success, true);
  assert.deepEqual(bumps, []);
});

test('内容が同じで書かなかった保存では版を上げない', () => {
  const { ctx, bumps } = loadCtx(BASE, { unchanged: true });
  const res = ctx.saveUserConfig_('u1', { ...BASE, sheetName: 'もう一度考える' });
  assert.equal(res.unchanged, true);
  assert.deepEqual(bumps, []);
});
