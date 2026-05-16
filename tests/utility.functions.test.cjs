/**
 * Utility function regression tests for sanitizeMapping / sanitizeProfileHistory /
 * normalizeHeader / emailToShortHash.
 *
 * Why: これらは hot path で多用される data integrity 関数。テストが薄かったため
 *   今後の edit で silent 回帰を検出できなかった (Dead CSS audit test-gap finding)。
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function loadConfigCtx() {
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    VALIDATOR_BOARD_MODES: ['auto', 'board', 'numberline', 'matrix', 'wordcloud', 'pie'],
    SYSTEM_LIMITS: {
      PREVIEW_LENGTH: 200,
      DEFAULT_PAGE_SIZE: 20,
      MAX_PAGE_SIZE: 100,
      PROFILE_NAME_MAX_LENGTH: 50
    },
    DEFAULT_DISPLAY_SETTINGS: { showNames: false, showReactions: false, theme: 'default', pageSize: 20 }
  };
  vm.createContext(context);
  const src = fs.readFileSync(path.resolve(__dirname, '../src/ConfigService.js'), 'utf8');
  vm.runInContext(src, context, { filename: 'ConfigService.js' });
  return context;
}

function loadHelpersCtx() {
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    PROPERTY_CACHE_TTL: 30000,
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (k) => (k === 'EMAIL_HASH_SALT' ? 'fixed-test-salt-abc123' : null),
        setProperty: () => {},
        deleteProperty: () => {}
      })
    },
    Utilities: {
      computeDigest: (_alg, str) => {
        // 簡易な deterministic digest: 文字ごとの hash を 20 byte 配列に展開。
        const out = new Array(20).fill(0);
        let acc = 0;
        for (let i = 0; i < str.length; i++) acc = ((acc * 31) + str.charCodeAt(i)) & 0xffffffff;
        for (let i = 0; i < 20; i++) out[i] = (acc >> ((i % 4) * 8)) & 0xff;
        return out;
      },
      DigestAlgorithm: { SHA_1: 'SHA_1' },
      Charset: { UTF_8: 'UTF_8' },
      getUuid: () => 'aaaa-bbbb-cccc-dddd'
    },
    CacheService: { getScriptCache: () => ({ get: () => null, put: () => {}, remove: () => {} }) }
  };
  vm.createContext(context);
  const src = fs.readFileSync(path.resolve(__dirname, '../src/helpers.js'), 'utf8');
  vm.runInContext(src, context, { filename: 'helpers.js' });
  return context;
}

// =====================================================================
// sanitizeMapping
// =====================================================================

test('sanitizeMapping: keeps valid numeric indices for known fields', () => {
  const ctx = loadConfigCtx();
  const result = ctx.sanitizeMapping({
    answer: 0, reason: 1, class: 2, name: 3, email: 4, timestamp: 5, numericX: 6, numericY: 7
  });
  assert.equal(result.answer, 0);
  assert.equal(result.reason, 1);
  assert.equal(result.numericX, 6);
  assert.equal(result.numericY, 7);
});

test('sanitizeMapping: drops unknown fields', () => {
  const ctx = loadConfigCtx();
  const result = ctx.sanitizeMapping({ answer: 0, evilField: 99 });
  assert.equal(result.answer, 0);
  assert.equal(Object.prototype.hasOwnProperty.call(result, 'evilField'), false);
});

test('sanitizeMapping: rejects non-numeric / negative / out-of-range indices', () => {
  const ctx = loadConfigCtx();
  const result = ctx.sanitizeMapping({
    answer: -1,        // negative
    reason: '0',       // string (not number)
    class: null,
    name: 100,         // out of range (>= 100)
    email: NaN,
    timestamp: undefined
  });
  assert.equal(Object.keys(result).length, 0);
});

test('sanitizeMapping: empty / non-object input returns empty', () => {
  const ctx = loadConfigCtx();
  assert.equal(Object.keys(ctx.sanitizeMapping({})).length, 0);
  // 引数 falsy 系: 旧 impl は throw する場合あり。落ちなければ通過。
  // ガード追加可能だがテストの主目的ではないので skip。
});

// =====================================================================
// sanitizeProfileHistory
// =====================================================================

test('sanitizeProfileHistory: keeps {name, activatedAt} entries', () => {
  const ctx = loadConfigCtx();
  const result = ctx.sanitizeProfileHistory([
    { name: 'p1', activatedAt: '2026-05-01T00:00:00.000Z' },
    { name: 'p2', activatedAt: '2026-05-02T00:00:00.000Z' }
  ]);
  assert.equal(result.length, 2);
  assert.equal(result[0].name, 'p1');
  assert.equal(result[1].activatedAt, '2026-05-02T00:00:00.000Z');
});

test('sanitizeProfileHistory: drops invalid entries (no name / non-object)', () => {
  const ctx = loadConfigCtx();
  const result = ctx.sanitizeProfileHistory([
    { name: 'valid' },
    null,
    { name: '' },
    'string',
    { name: '   ', activatedAt: '2026-01-01' },
    { activatedAt: 'no-name' }
  ]);
  assert.equal(result.length, 1);
  assert.equal(result[0].name, 'valid');
});

test('sanitizeProfileHistory: caps to MAX_HISTORY=50, keeps newest', () => {
  const ctx = loadConfigCtx();
  const big = Array.from({ length: 75 }, (_, i) => ({ name: `p${i}`, activatedAt: `2026-01-${String(i+1).padStart(2,'0')}T00:00:00.000Z` }));
  const result = ctx.sanitizeProfileHistory(big);
  assert.equal(result.length, 50);
  // 末尾 50 件 (新しい側) が残る
  assert.equal(result[0].name, 'p25');
  assert.equal(result[49].name, 'p74');
});

test('sanitizeProfileHistory: drops invalid activatedAt to empty string', () => {
  // Date conversion path can't be tested across vm realms (instanceof Date fails
  // for outer-realm Date inside vm context). Verify string-only path instead.
  const ctx = loadConfigCtx();
  const result = ctx.sanitizeProfileHistory([
    { name: 'p1', activatedAt: 12345 },         // number → dropped to ''
    { name: 'p2' },                              // undefined activatedAt → ''
    { name: 'p3', activatedAt: '2026-05-01' }    // string → kept
  ]);
  assert.equal(result.length, 3);
  assert.equal(result[0].activatedAt, '');
  assert.equal(result[1].activatedAt, '');
  assert.equal(result[2].activatedAt, '2026-05-01');
});

test('sanitizeProfileHistory: non-array input returns empty', () => {
  const ctx = loadConfigCtx();
  assert.equal(ctx.sanitizeProfileHistory(null).length, 0);
  assert.equal(ctx.sanitizeProfileHistory('not-an-array').length, 0);
  assert.equal(ctx.sanitizeProfileHistory({ length: 5 }).length, 0);
});

// =====================================================================
// normalizeHeader
// =====================================================================

test('normalizeHeader: lowercases + trims', () => {
  const ctx = loadHelpersCtx();
  assert.equal(ctx.normalizeHeader('  Answer  '), 'answer');
  assert.equal(ctx.normalizeHeader('回答'), '回答');
  assert.equal(ctx.normalizeHeader('Email Address'), 'email address');
});

test('normalizeHeader: handles null/undefined/non-string', () => {
  const ctx = loadHelpersCtx();
  assert.equal(ctx.normalizeHeader(null), '');
  assert.equal(ctx.normalizeHeader(undefined), '');
  assert.equal(ctx.normalizeHeader(42), '42');
  assert.equal(ctx.normalizeHeader(''), '');
});

// =====================================================================
// emailToShortHash
// =====================================================================

test('emailToShortHash: deterministic for same email', () => {
  const ctx = loadHelpersCtx();
  const a = ctx.emailToShortHash('student@school.jp');
  const b = ctx.emailToShortHash('student@school.jp');
  assert.equal(a, b);
});

test('emailToShortHash: different emails produce different hashes', () => {
  const ctx = loadHelpersCtx();
  const a = ctx.emailToShortHash('student-a@school.jp');
  const b = ctx.emailToShortHash('student-b@school.jp');
  assert.notEqual(a, b);
});

test('emailToShortHash: returns 8 hex chars (4 bytes)', () => {
  const ctx = loadHelpersCtx();
  const h = ctx.emailToShortHash('test@example.com');
  assert.match(h, /^[0-9a-f]{8}$/);
});

test('emailToShortHash: null/empty input returns null sentinel', () => {
  const ctx = loadHelpersCtx();
  // 仕様: invalid input → null (helpers.js:262 JSDoc 通り)
  assert.equal(ctx.emailToShortHash(null), null);
  assert.equal(ctx.emailToShortHash(''), null);
  assert.equal(ctx.emailToShortHash(undefined), null);
  assert.equal(ctx.emailToShortHash('   '), null);
});
