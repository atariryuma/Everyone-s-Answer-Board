/**
 * executeWithRetry / isRetryableError tests
 *
 * Why: これらは DatabaseCore.fetchSheetsAPIWithRetry や DataApis.getSheetNameFromGid 等の
 *      backbone。Naha GCP の 429 storm に対する load-bearing なリトライ戦略だが、
 *      従来 audit で「全テストで stub されており直接の挙動テストが無い」と CRITICAL
 *      指定されていた。本ファイルで:
 *        - retry の発生条件 (retryable error / max-retries / non-retryable early break)
 *        - is429Error の doubled backoff
 *        - Utilities.sleep のコール回数と引数
 *        - isRetryableError のパターン優先順位 (non-retryable check 先)
 *      を pin する。
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { gasResponseStubs } = require('./_helpers.cjs');

// vm.Script を 1 度だけ parse して使い回す（テスト数増えても parse コスト一定）。
const MAIN_SOURCE = fs.readFileSync(path.resolve(__dirname, '../src/main.js'), 'utf8');
const MAIN_SCRIPT = new vm.Script(MAIN_SOURCE, { filename: 'main.js' });

function loadMainCtx(overrides = {}) {
  const sleeps = []; // テストから検査するため context の外向け配列
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    ContentService: {
      MimeType: { JSON: 'application/json' },
      createTextOutput: (b) => ({ body: b, setMimeType: () => ({ body: b }) })
    },
    Session: { getActiveUser: () => ({ getEmail: () => 'a@example.com' }) },
    Utilities: { sleep: (ms) => sleeps.push(ms) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null, setProperty: () => {} }) },
    CacheService: { getScriptCache: () => ({ get: () => null, put: () => {}, remove: () => {} }) },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
    ...gasResponseStubs(),
    findUserByEmail: () => null,
    findUserById: () => null,
    findPublishedBoardOwner: () => null,
    getCachedProperty: () => null,
    setCachedProperty: () => {},
    isAdministrator: () => false,
    getConfigOrDefault: () => ({}),
    getUserSheetData: () => ({}),
    addReaction: () => ({}),
    toggleHighlight: () => ({}),
    enhanceConfigWithDynamicUrls: (c) => c,
    shouldEnforceDomainRestrictions: () => false,
    validateDomainAccess: () => ({ allowed: true }),
    dispatchAdminOperation: () => ({ success: true }),
    timingSafeEqual: (a, b) => a === b,
    hasCoreSystemProps: () => true,
    getQuestionText: () => '',
    getWebAppUrl: () => 'https://example.com',
    publishApp: () => ({ success: true }),
    __sleeps: sleeps
  };
  Object.assign(context, overrides);
  vm.createContext(context);
  MAIN_SCRIPT.runInContext(context);
  return context;
}

// =====================================================================
// isRetryableError — パターン分類
// =====================================================================

test('isRetryableError: timeout / network / quota / rate limit はリトライ', () => {
  const { isRetryableError } = loadMainCtx();
  assert.equal(isRetryableError('Connection timeout'), true);
  assert.equal(isRetryableError('Network unreachable'), true);
  assert.equal(isRetryableError('Quota exceeded (429)'), true);
  assert.equal(isRetryableError('Rate limit hit'), true);
  assert.equal(isRetryableError('Service unavailable'), true);
  assert.equal(isRetryableError('Internal error'), true);
  assert.equal(isRetryableError('Backend error 503'), true);
  assert.equal(isRetryableError('Connection reset'), true);
  assert.equal(isRetryableError('Socket hang up'), true);
});

test('isRetryableError: permission / not found / authentication failed はリトライしない', () => {
  const { isRetryableError } = loadMainCtx();
  assert.equal(isRetryableError('Permission denied'), false);
  assert.equal(isRetryableError('Resource not found'), false);
  assert.equal(isRetryableError('Not authorized'), false);
  assert.equal(isRetryableError('Invalid argument'), false);
  assert.equal(isRetryableError('Malformed request'), false);
  assert.equal(isRetryableError('Access denied'), false);
  assert.equal(isRetryableError('Authentication failed'), false);
});

test('isRetryableError: 非冪等 write (:append) の失敗はリトライしない (H3)', () => {
  // Why: :append は非冪等。5xx/network 断は commit 済か判定不能で、盲目 retry は重複行を生む。
  //   fetchSheetsAPIWithRetry が idempotent=false のとき "non-idempotent" を含むメッセージを
  //   throw し、それを non-retryable にすることで executeWithRetry の再実行を止める。
  const { isRetryableError } = loadMainCtx();
  assert.equal(isRetryableError('non-idempotent write failed (no retry): API returned 503: backend error'), false);
  assert.equal(isRetryableError('non-idempotent write aborted (no retry): connection reset'), false);
});

test('isRetryableError: 非string / falsy 入力は false', () => {
  const { isRetryableError } = loadMainCtx();
  assert.equal(isRetryableError(null), false);
  assert.equal(isRetryableError(undefined), false);
  assert.equal(isRetryableError(''), false);
  assert.equal(isRetryableError(42), false);
  assert.equal(isRetryableError({}), false);
});

test('isRetryableError: non-retryable が先勝ち — "invalid timeout" は false', () => {
  // Why: source の for ループは nonRetryable を先に走らせるため、両方のキーワードを
  //      含むメッセージは non-retryable 判定になる。これは契約。
  const { isRetryableError } = loadMainCtx();
  assert.equal(isRetryableError('invalid timeout value'), false,
    'non-retryable "invalid" should win over retryable "timeout"');
  assert.equal(isRetryableError('permission denied during network operation'), false);
});

test('isRetryableError: 未知のメッセージはデフォルトでリトライ (true)', () => {
  // Why: 安全側に倒す。未知エラーは一時的かもしれないので少なくとも 1 回はリトライさせる。
  const { isRetryableError } = loadMainCtx();
  assert.equal(isRetryableError('Some completely unknown error message'), true);
  assert.equal(isRetryableError('Foo bar baz'), true);
});

test('isRetryableError: 大文字小文字を区別しない', () => {
  const { isRetryableError } = loadMainCtx();
  assert.equal(isRetryableError('PERMISSION DENIED'), false);
  assert.equal(isRetryableError('Quota Exceeded'), true);
});

// =====================================================================
// evaluateDomainRestriction — fail-closed after setup (M6)
// =====================================================================

test('evaluateDomainRestriction: setup 未完了 (enforce 不要) は通す (fail-open)', () => {
  const { evaluateDomainRestriction } = loadMainCtx({
    shouldEnforceDomainRestrictions: () => false
  });
  const r = evaluateDomainRestriction('x@out.com');
  assert.equal(r.allowed, true);
  assert.equal(r.reason, 'setup_incomplete');
});

test('evaluateDomainRestriction: enforce 必要 + validateDomainAccess が throw → deny (fail-closed M6)', () => {
  const { evaluateDomainRestriction } = loadMainCtx({
    shouldEnforceDomainRestrictions: () => true,
    validateDomainAccess: () => { throw new Error('transient cache error'); }
  });
  const r = evaluateDomainRestriction('x@out.com');
  assert.equal(r.allowed, false, 'enforcement confirmed → exception must fail closed');
  assert.equal(r.reason, 'validation_error');
});

test('evaluateDomainRestriction: validator 不在 → 通す (コード未読込はランタイム事象でない)', () => {
  // validator が単一グローバルスコープに無い = アプリ全体未読込であり、 M6 (transient 例外の
  //   fail-closed) の対象外。 従来どおり fail-open。
  const { evaluateDomainRestriction } = loadMainCtx({
    shouldEnforceDomainRestrictions: () => true,
    validateDomainAccess: undefined
  });
  const r = evaluateDomainRestriction('x@out.com');
  assert.equal(r.allowed, true);
  assert.equal(r.reason, 'validator_missing');
});

test('evaluateDomainRestriction: enforcement 判定自体が throw → 状態不明なので通す (可用性優先)', () => {
  const { evaluateDomainRestriction } = loadMainCtx({
    shouldEnforceDomainRestrictions: () => { throw new Error('cannot determine setup state'); }
  });
  const r = evaluateDomainRestriction('x@out.com');
  assert.equal(r.allowed, true, 'unknown enforcement state stays fail-open to avoid setup lockout');
});

test('evaluateDomainRestriction: enforce 必要 + validateDomainAccess が allowed を返せばそれを尊重', () => {
  const { evaluateDomainRestriction } = loadMainCtx({
    shouldEnforceDomainRestrictions: () => true,
    validateDomainAccess: (email) => ({ allowed: /@in\.com$/.test(email), reason: 'checked' })
  });
  assert.equal(evaluateDomainRestriction('teacher@in.com').allowed, true);
  assert.equal(evaluateDomainRestriction('x@out.com').allowed, false);
});

// =====================================================================
// executeWithRetry — リトライ戦略
// =====================================================================

test('executeWithRetry: 初回成功時は operation を 1 回だけ呼び結果を返す', () => {
  const ctx = loadMainCtx();
  let calls = 0;
  const result = ctx.executeWithRetry(() => { calls++; return 'ok'; });
  assert.equal(result, 'ok');
  assert.equal(calls, 1);
  assert.equal(ctx.__sleeps.length, 0, 'no sleep on first-try success');
});

test('executeWithRetry: retryable error → 再試行して 2 回目で成功', () => {
  const ctx = loadMainCtx();
  let calls = 0;
  const result = ctx.executeWithRetry(() => {
    calls++;
    if (calls === 1) throw new Error('Network timeout');
    return 'recovered';
  }, { operationName: 'TestOp' });
  assert.equal(result, 'recovered');
  assert.equal(calls, 2);
  assert.equal(ctx.__sleeps.length, 1, 'one sleep before retry');
});

test('executeWithRetry: non-retryable error → 即座に throw、リトライしない', () => {
  const ctx = loadMainCtx();
  let calls = 0;
  assert.throws(
    () => ctx.executeWithRetry(() => { calls++; throw new Error('Permission denied'); }),
    /Permission denied/
  );
  assert.equal(calls, 1, 'non-retryable error must not trigger retry');
  assert.equal(ctx.__sleeps.length, 0);
});

test('executeWithRetry: maxRetries 到達 → lastError を throw', () => {
  const ctx = loadMainCtx();
  let calls = 0;
  assert.throws(
    () => ctx.executeWithRetry(
      () => { calls++; throw new Error('Network timeout'); },
      { maxRetries: 3 }
    ),
    /Network timeout/
  );
  assert.equal(calls, 3, 'should attempt exactly maxRetries times');
  assert.equal(ctx.__sleeps.length, 2, '2 sleeps between 3 attempts');
});

test('executeWithRetry: 429 error は initialDelay * 2 で開始 (doubled backoff)', () => {
  const ctx = loadMainCtx();
  let calls = 0;
  ctx.executeWithRetry(() => {
    calls++;
    if (calls < 3) throw new Error('Quota exceeded (429)');
    return 'ok';
  }, { initialDelay: 500, maxRetries: 5 });

  // 1回目失敗 → retry 1 で sleep。baseDelay = 500*2 = 1000, pow(2,0) = 1 → 1000ms
  // 2回目失敗 → retry 2 で sleep。baseDelay = 1000, pow(2,1) = 2 → 2000ms
  assert.equal(ctx.__sleeps.length, 2);
  assert.equal(ctx.__sleeps[0], 1000, 'first retry: doubled initialDelay');
  assert.equal(ctx.__sleeps[1], 2000, 'second retry: exponential backoff');
});

test('executeWithRetry: 通常エラーは initialDelay で開始 (no doubling)', () => {
  const ctx = loadMainCtx();
  let calls = 0;
  ctx.executeWithRetry(() => {
    calls++;
    if (calls < 3) throw new Error('Network connection reset');
    return 'ok';
  }, { initialDelay: 500, maxRetries: 5 });

  assert.equal(ctx.__sleeps[0], 500, 'non-429 first retry: initialDelay raw');
  assert.equal(ctx.__sleeps[1], 1000, 'non-429 second retry: 2x');
});

test('executeWithRetry: maxDelay を超える backoff は cap される', () => {
  const ctx = loadMainCtx();
  let calls = 0;
  ctx.executeWithRetry(() => {
    calls++;
    if (calls < 5) throw new Error('Network timeout');
    return 'ok';
  }, { initialDelay: 500, maxDelay: 1500, maxRetries: 6 });

  // retry 1: 500, retry 2: 1000, retry 3: 2000 → cap 1500, retry 4: 4000 → cap 1500
  assert.deepEqual(Array.from(ctx.__sleeps), [500, 1000, 1500, 1500]);
});

test('executeWithRetry: operation が string を throw しても lastError として保持', () => {
  // Why: throw 'string' は JS で許可されており、error.message が undefined になる。
  //      executeWithRetry は String(error) でフォールバックする実装になっているはず。
  const ctx = loadMainCtx();
  let calls = 0;
  assert.throws(
    () => ctx.executeWithRetry(() => {
      calls++;
      throw 'string error: timeout';
    }, { maxRetries: 2 }),
    (thrown) => {
      // thrown は文字列 or Error — 実装は throw lastError をそのまま伝搬する
      return thrown === 'string error: timeout' || (thrown.message && thrown.message.includes('timeout'));
    }
  );
  assert.equal(calls, 2);
});

test('executeWithRetry: maxRetries=1 は再試行ゼロ (1 回試して失敗で即 throw)', () => {
  const ctx = loadMainCtx();
  let calls = 0;
  assert.throws(
    () => ctx.executeWithRetry(
      () => { calls++; throw new Error('Network timeout'); },
      { maxRetries: 1 }
    ),
    /Network timeout/
  );
  assert.equal(calls, 1);
  assert.equal(ctx.__sleeps.length, 0);
});

test('executeWithRetry: operation の戻り値が undefined / null でも正常 return', () => {
  // Why: 過去に retry コードで「成功 = truthy」と誤判定するバグがあった。
  //      undefined や null を返す operation も第1試行で return する契約を pin。
  const ctx = loadMainCtx();
  assert.equal(ctx.executeWithRetry(() => undefined), undefined);
  assert.equal(ctx.executeWithRetry(() => null), null);
  assert.equal(ctx.executeWithRetry(() => 0), 0);
  assert.equal(ctx.executeWithRetry(() => false), false);
});

test('executeWithRetry: operationName がログ出力に使われる (副次効果テスト)', () => {
  const warnLogs = [];
  const errorLogs = [];
  const ctx = loadMainCtx({
    console: { log: () => {}, warn: (...a) => warnLogs.push(a.join(' ')), error: (...a) => errorLogs.push(a.join(' ')) }
  });
  assert.throws(
    () => ctx.executeWithRetry(
      () => { throw new Error('Network timeout'); },
      { maxRetries: 3, operationName: 'MyCustomOp' }
    ),
    /Network timeout/
  );
  const allLogs = [...warnLogs, ...errorLogs].join('\n');
  assert.match(allLogs, /MyCustomOp/);
});
