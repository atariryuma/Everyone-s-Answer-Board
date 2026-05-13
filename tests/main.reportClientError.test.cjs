const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

/**
 * Tests for the reportClientError doPost action and handleClientErrorReport handler.
 *
 * Why: フロントエラーを Cloud Logging に流す唯一の入口なので、
 *      validation・rate-limit・truncation・xss-safe な logging を確実にする。
 */

function createContentServiceStub() {
  return {
    MimeType: { JSON: 'application/json' },
    createTextOutput(body) {
      return {
        body,
        mimeType: null,
        setMimeType(mimeType) {
          this.mimeType = mimeType;
          return this;
        }
      };
    }
  };
}

function loadMainContext(overrides = {}, errorSink) {
  // 引数の errorSink が console.error の呼び出しを蓄積するなら、ログ内容を検証可能。
  const logs = errorSink || { error: [], warn: [], log: [] };
  const context = {
    console: {
      log: (...args) => logs.log.push(args),
      warn: (...args) => logs.warn.push(args),
      error: (...args) => logs.error.push(args)
    },
    ContentService: createContentServiceStub(),
    createSuccessResponse: (message, data, extra) => ({ success: true, message, ...(data && { data }), ...extra }),
    createAuthError: () => ({ success: false, message: 'auth required' }),
    createUserNotFoundError: () => ({ success: false, message: 'user not found' }),
    createErrorResponse: (message, data, extra) => ({ success: false, message, error: message, ...extra }),
    createExceptionResponse: (error) => ({ success: false, message: error.message, error: error.message }),
    createAdminRequiredError: () => ({ success: false, message: 'admin required' }),
    Session: {
      getActiveUser: () => ({ getEmail: () => 'student@example.com' })
    },
    evaluateDomainRestriction: () => ({ allowed: true }),
    findUserByEmail: () => ({ userId: 'user-1' }),
    getUserSheetData: () => ({ rows: [] }),
    addReaction: () => ({ success: true }),
    toggleHighlight: () => ({ success: true }),
    publishApp: () => ({ success: true }),
    PropertiesService: {
      getScriptProperties: () => ({ getProperty: () => null, setProperty: () => {} })
    },
    CacheService: {
      // Rate-limit テスト用：put した内容を覚えておく簡易キャッシュ
      _store: new Map(),
      getScriptCache() {
        const store = this._store;
        return {
          get(key) { return store.get(key) || null; },
          put(key, value /*, ttl */) { store.set(key, value); }
        };
      }
    },
    Utilities: {
      DigestAlgorithm: { SHA_1: 'SHA_1' },
      // 簡易 hash（テスト目的、本物の SHA-1 でなくてよい）
      computeDigest(_alg, input) {
        const bytes = [];
        for (let i = 0; i < 8; i++) {
          let h = 0;
          for (let j = 0; j < String(input).length; j++) {
            h = (h * 31 + String(input).charCodeAt(j) + i) & 0xff;
          }
          bytes.push(h);
        }
        return bytes;
      }
    },
    getCachedProperty: () => null,
    dispatchAdminOperation: () => ({ success: false }),
    timingSafeEqual: (a, b) => a === b
  };
  Object.assign(context, overrides);
  if (!context.setCachedProperty) {
    context.setCachedProperty = (key, value) => {
      try { context.PropertiesService.getScriptProperties().setProperty(key, value); } catch (_) {}
    };
  }
  vm.createContext(context);
  const source = fs.readFileSync(path.resolve(__dirname, '../src/main.js'), 'utf8');
  vm.runInContext(source, context, { filename: 'main.js' });
  return { context, logs };
}

function parseResponse(output) {
  assert.ok(output);
  assert.equal(output.mimeType, 'application/json');
  return JSON.parse(output.body);
}

function postEvent(payload) {
  return { postData: { contents: JSON.stringify(payload), type: 'application/json' } };
}

// =====================================================================
// doPost dispatch
// =====================================================================

test('reportClientError: action is in allowlist (no UNKNOWN_ACTION)', () => {
  const { context } = loadMainContext();
  const res = parseResponse(context.doPost(postEvent({
    action: 'reportClientError',
    payload: { level: 'error', message: 'boom' }
  })));
  // success or RATE_LIMIT (cache works), but never UNKNOWN_ACTION
  assert.notEqual(res.error, 'UNKNOWN_ACTION');
  assert.equal(res.success, true);
});

test('reportClientError: rejects non-object payload', () => {
  const { context } = loadMainContext();
  const res = parseResponse(context.doPost(postEvent({
    action: 'reportClientError',
    payload: 'not an object'
  })));
  assert.equal(res.success, false);
  assert.equal(res.error, 'INVALID_PAYLOAD');
});

// =====================================================================
// handleClientErrorReport: structured logging
// =====================================================================

test('handleClientErrorReport: logs error-level to console.error', () => {
  const { context, logs } = loadMainContext();
  const res = context.handleClientErrorReport('user@example.com', {
    level: 'error',
    message: 'TypeError: x is null',
    context: 'renderBoard',
    stack: 'at foo:1\nat bar:2',
    source: 'page.js',
    line: 42,
    url: 'https://app/?mode=view'
  });
  assert.equal(res.success, true);
  assert.equal(logs.error.length, 1);
  const [summary, detail] = logs.error[0];
  assert.match(summary, /\[client\/error\]/);
  assert.match(summary, /renderBoard/);
  assert.match(summary, /TypeError/);
  assert.equal(detail.line, 42);
  // メアドはマスクされて記録される（local part も含めて完全仮名化: anon-XXXXXXXX）
  assert.match(detail.reporter, /^anon-[0-9a-f]{8}$/);
});

test('handleClientErrorReport: warn level → console.warn', () => {
  const { context, logs } = loadMainContext();
  const res = context.handleClientErrorReport('s@x.com', {
    level: 'warn', message: 'deprecated API'
  });
  assert.equal(res.success, true);
  assert.equal(logs.warn.length, 1);
  assert.match(logs.warn[0][0], /\[client\/warn\]/);
});

test('handleClientErrorReport: invalid level falls back to error', () => {
  const { context, logs } = loadMainContext();
  context.handleClientErrorReport('s@x.com', { level: 'panic', message: 'bad level' });
  // panic → falls back to error
  assert.equal(logs.error.length, 1);
});

// =====================================================================
// Truncation: protect Cloud Logging from huge payloads
// =====================================================================

test('handleClientErrorReport: truncates oversized message/stack/context', () => {
  const { context, logs } = loadMainContext();
  const longMsg = 'x'.repeat(2000);
  const longStack = 'y'.repeat(5000);
  const longCtx = 'z'.repeat(800);
  context.handleClientErrorReport('s@x.com', {
    level: 'error',
    message: longMsg,
    stack: longStack,
    context: longCtx
  });
  const [summary, detail] = logs.error[0];
  // summary includes message but truncated to ≤500
  assert.ok(summary.length < longMsg.length, 'summary should be much shorter than input');
  // stack capped at 2000 chars
  assert.ok(detail.stack && detail.stack.length <= 2000);
});

test('handleClientErrorReport: strips CR/LF/TAB from logged strings', () => {
  const { context, logs } = loadMainContext();
  context.handleClientErrorReport('s@x.com', {
    level: 'error',
    message: 'line1\nline2\rline3\t',
    context: 'a\nb'
  });
  const [summary] = logs.error[0];
  assert.equal(summary.includes('\n'), false);
  assert.equal(summary.includes('\r'), false);
  assert.equal(summary.includes('\t'), false);
});

// =====================================================================
// Rate limiting
// =====================================================================

test('handleClientErrorReport: rate-limits after 30 reports per email', () => {
  const { context, logs } = loadMainContext();
  for (let i = 0; i < 30; i++) {
    context.handleClientErrorReport('s@x.com', { level: 'error', message: `err ${i}` });
  }
  // 31 件目で RATE_LIMIT
  const blocked = context.handleClientErrorReport('s@x.com', { level: 'error', message: 'over limit' });
  assert.equal(blocked.success, false);
  assert.equal(blocked.error, 'RATE_LIMIT');
  // 1-30 はログに残るが 31 は残らない
  assert.equal(logs.error.length, 30);
});

test('handleClientErrorReport: rate limits independent per email', () => {
  const { context } = loadMainContext();
  for (let i = 0; i < 30; i++) {
    context.handleClientErrorReport('alice@x.com', { level: 'error', message: 'a' });
  }
  // bob は影響なし
  const bobRes = context.handleClientErrorReport('bob@x.com', { level: 'error', message: 'b' });
  assert.equal(bobRes.success, true);
});

// =====================================================================
// Defensive
// =====================================================================

test('handleClientErrorReport: cache failure still allows logging (best-effort)', () => {
  const { context, logs } = loadMainContext({
    CacheService: {
      getScriptCache() {
        return {
          get() { throw new Error('cache outage'); },
          put() { throw new Error('cache outage'); }
        };
      }
    }
  });
  const res = context.handleClientErrorReport('s@x.com', { level: 'error', message: 'msg' });
  assert.equal(res.success, true);
  assert.equal(logs.error.length, 1);
});

test('handleClientErrorReport: empty message still logs (degenerate case)', () => {
  const { context, logs } = loadMainContext();
  const res = context.handleClientErrorReport('s@x.com', { level: 'error' });
  assert.equal(res.success, true);
  // logs.error[0][0] = "[client/error] " (no context, empty message)
  assert.equal(logs.error.length, 1);
});

test('handleClientErrorReport: unknown email is allowed (sometimes called pre-auth)', () => {
  const { context } = loadMainContext();
  // pass '' as email — should still log (reporter shown as 'unknown')
  const res = context.handleClientErrorReport('', { level: 'error', message: 'x' });
  assert.equal(res.success, true);
});
