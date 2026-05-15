/**
 * Option B: profileHistory append + sanitize tests.
 *
 * Why: 教師が profile 切替 → history append → 生徒の wire に乗る、という data flow
 *   全体の根拠は __appendProfileHistory_ と sanitizeProfileHistory が正しく動くこと。
 *   重複防止 / 削除済 entry の除去 / cap 50 件 / 直前と同名は no-op の挙動を回帰防止する。
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('path');
const vm = require('vm');
const { gasResponseStubs } = require('./_helpers.cjs');

function loadAdminContext(overrides = {}) {
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    ...gasResponseStubs(),
    getCurrentEmail: () => 'admin@example.com',
    isAdministrator: () => true,
    findUserByEmail: () => ({ userId: 'u1', userEmail: 'admin@example.com' }),
    findUserById: () => ({ userId: 'u1', userEmail: 'admin@example.com', isActive: true }),
    getUserConfig: () => ({ success: true, config: {} }),
    saveUserConfig: () => ({ success: true }),
    requireAdmin: () => ({ email: 'admin@example.com', isAdmin: true }),
    getConfigOrDefault: () => ({}),
    getBatchedAdminAuth: () => ({ success: true, email: 'admin@example.com', isAdmin: true }),
    DEFAULT_DISPLAY_SETTINGS: { showNames: false, showReactions: true, theme: 'default', pageSize: 20 },
    PropertiesService: {
      getScriptProperties: () => ({ getProperty: () => null, setProperty: () => {}, getProperties: () => ({}) })
    },
    getCachedProperty: () => null,
    ...overrides
  };
  vm.createContext(context);
  const source = fs.readFileSync(path.resolve(__dirname, '../src/AdminApis.js'), 'utf8');
  vm.runInContext(source, context, { filename: 'AdminApis.js' });
  return context;
}

function loadConfigContext(overrides = {}) {
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    SYSTEM_LIMITS: { PREVIEW_LENGTH: 100, DEFAULT_PAGE_SIZE: 20, MAX_PAGE_SIZE: 100 },
    DEFAULT_DISPLAY_SETTINGS: { showNames: false, showReactions: true, theme: 'default', pageSize: 20 },
    validateConfig: () => ({ isValid: true, sanitized: {}, errors: [] }),
    VALIDATOR_BOARD_MODES: ['auto', 'board', 'numberline', 'matrix', 'wordcloud', 'pie'],
    CacheService: { getScriptCache: () => ({ get: () => null, put: () => {}, remove: () => {}, removeAll: () => {} }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
    getCachedProperty: () => null,
    getWebAppUrl: () => 'https://example.com/exec',
    ...overrides
  };
  vm.createContext(context);
  const source = fs.readFileSync(path.resolve(__dirname, '../src/ConfigService.js'), 'utf8');
  vm.runInContext(source, context, { filename: 'ConfigService.js' });
  return context;
}

// =====================================================================
// __appendProfileHistory_ helper
// =====================================================================

test('__appendProfileHistory_: appends new entry with activatedAt', () => {
  const ctx = loadAdminContext();
  const result = ctx.__appendProfileHistory_([], '導入');
  assert.equal(result.length, 1);
  assert.equal(result[0].name, '導入');
  assert.ok(result[0].activatedAt);
  // ISO8601 format
  assert.match(result[0].activatedAt, /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/);
});

test('__appendProfileHistory_: returns new array (immutable)', () => {
  const ctx = loadAdminContext();
  const original = [{ name: '導入', activatedAt: '2026-05-14T00:00:00Z' }];
  const result = ctx.__appendProfileHistory_(original, '本時');
  assert.equal(original.length, 1);  // 元配列は mutate されない
  assert.equal(result.length, 2);
});

test('__appendProfileHistory_: no-op when last entry is same name', () => {
  const ctx = loadAdminContext();
  const original = [
    { name: '導入', activatedAt: '2026-05-14T00:00:00Z' },
    { name: '本時', activatedAt: '2026-05-14T01:00:00Z' }
  ];
  const result = ctx.__appendProfileHistory_(original, '本時');
  assert.equal(result.length, 2);  // 重複 append しない
});

test('__appendProfileHistory_: appends after gap (non-adjacent same name allowed)', () => {
  const ctx = loadAdminContext();
  const original = [
    { name: '導入', activatedAt: '2026-05-14T00:00:00Z' },
    { name: '本時', activatedAt: '2026-05-14T01:00:00Z' }
  ];
  // 「導入」に戻すケースは記録する（教師が振り返りを促す価値あり）
  const result = ctx.__appendProfileHistory_(original, '導入');
  assert.equal(result.length, 3);
  assert.equal(result[2].name, '導入');
});

test('__appendProfileHistory_: empty profileName is no-op', () => {
  const ctx = loadAdminContext();
  const result = ctx.__appendProfileHistory_([{ name: 'x', activatedAt: '2026-01-01' }], '');
  assert.equal(result.length, 1);
});

test('__appendProfileHistory_: null currentHistory treated as empty', () => {
  const ctx = loadAdminContext();
  const result = ctx.__appendProfileHistory_(null, '本時');
  assert.equal(result.length, 1);
  assert.equal(result[0].name, '本時');
});

// =====================================================================
// sanitizeProfileHistory (ConfigService)
// =====================================================================

test('sanitizeProfileHistory: drops invalid entries', () => {
  const ctx = loadConfigContext();
  const out = ctx.sanitizeProfileHistory([
    { name: '導入', activatedAt: '2026-05-14T00:00:00Z' },
    null,                              // 落とす
    { activatedAt: '2026' },           // name なし → 落とす
    'string',                          // 落とす
    { name: '本時', activatedAt: '' }  // activatedAt 空でも name あれば残す
  ]);
  assert.equal(out.length, 2);
  assert.equal(out[0].name, '導入');
  assert.equal(out[1].name, '本時');
});

test('sanitizeProfileHistory: caps to 50 entries (keeps last 50)', () => {
  const ctx = loadConfigContext();
  const input = [];
  for (let i = 0; i < 60; i++) {
    input.push({ name: 'p' + i, activatedAt: '2026-05-14T' + String(i % 24).padStart(2, '0') + ':00:00Z' });
  }
  const out = ctx.sanitizeProfileHistory(input);
  assert.equal(out.length, 50);
  // 末尾が新しい側として残るので、最後の entry は p59
  assert.equal(out[out.length - 1].name, 'p59');
  // 最初の entry は p10 (60-50=10)
  assert.equal(out[0].name, 'p10');
});

test('sanitizeProfileHistory: non-array returns empty array', () => {
  // Why: vm context boundary — host の [] と vm 内 [] は別 prototype。length で比較。
  const ctx = loadConfigContext();
  assert.equal(ctx.sanitizeProfileHistory(null).length, 0);
  assert.equal(ctx.sanitizeProfileHistory('x').length, 0);
  assert.equal(ctx.sanitizeProfileHistory({}).length, 0);
});

test('sanitizeProfileHistory: Date-string activatedAt is preserved', () => {
  // Why: vm-host instanceof Date は context 跨ぎで false なので、test では string ISO のみカバー。
  //   __appendProfileHistory_ は new Date().toISOString() で string にして渡すので、運用経路では問題なし。
  const ctx = loadConfigContext();
  const iso = '2026-05-14T10:00:00.000Z';
  const out = ctx.sanitizeProfileHistory([{ name: 'p', activatedAt: iso }]);
  assert.equal(out.length, 1);
  assert.equal(out[0].activatedAt, iso);
});

test('sanitizeProfileHistory: name truncated to 50 chars', () => {
  const ctx = loadConfigContext();
  const longName = 'x'.repeat(80);
  const out = ctx.sanitizeProfileHistory([{ name: longName, activatedAt: '2026' }]);
  assert.equal(out.length, 1);
  assert.equal(out[0].name.length, 50);
});

// =====================================================================
// validateAndSanitizeConfig: profileHistory roundtrip
// =====================================================================

test('validateAndSanitizeConfig: preserves profileHistory', () => {
  const ctx = loadConfigContext();
  const input = {
    profileHistory: [
      { name: '導入', activatedAt: '2026-05-14T00:00:00Z' },
      { name: '本時', activatedAt: '2026-05-14T01:00:00Z' }
    ]
  };
  const result = ctx.validateAndSanitizeConfig(input, 'u1');
  assert.equal(result.success, true);
  assert.equal(result.data.profileHistory.length, 2);
});

test('validateAndSanitizeConfig: empty profileHistory is removed from config', () => {
  const ctx = loadConfigContext();
  const result = ctx.validateAndSanitizeConfig({ profileHistory: [] }, 'u1');
  assert.equal(result.success, true);
  assert.equal('profileHistory' in result.data, false);
});
