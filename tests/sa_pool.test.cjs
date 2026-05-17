// v2782+: SA pool (round-robin + cooldown + 2 段 token cache + per-SA verify cache) のテスト。
// DatabaseCore.js を vm.createContext で読み込み、 GAS API を最小 stub で再現する。

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function createMockCache(initial = {}) {
  const store = new Map(Object.entries(initial));
  return {
    get: (k) => (store.has(k) ? store.get(k) : null),
    put: (k, v) => { store.set(k, v); },
    remove: (k) => { store.delete(k); },
    _store: store
  };
}

function makeSaJson(email, idx) {
  return JSON.stringify({
    type: 'service_account',
    client_email: email,
    private_key: '-----BEGIN PRIVATE KEY-----\nFAKE-' + idx + '\n-----END PRIVATE KEY-----\n',
    project_id: 'proj-' + idx,
    private_key_id: ('abcd1234' + idx).repeat(2).slice(0, 40)
  });
}

function loadCtx(overrides = {}) {
  const cache = overrides.cache || createMockCache();
  const propsStore = overrides.propsStore || {};

  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    CacheService: { getScriptCache: () => cache },
    CACHE_DURATION: {
      SHORT: 10, MEDIUM: 30, FORM_DATA: 30, LONG: 300,
      DATABASE_LONG: 600, USER_INDIVIDUAL: 900, EXTRA_LONG: 3600
    },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (k) => (k in propsStore ? propsStore[k] : null),
        setProperty: (k, v) => { propsStore[k] = v; },
        deleteProperty: (k) => { delete propsStore[k]; },
        getProperties: () => ({ ...propsStore })
      })
    },
    SpreadsheetApp: {
      openById: () => { throw new Error('not stubbed'); }
    },
    DriveApp: overrides.DriveApp || {
      getFileById: () => ({ addEditor: () => {} })
    },
    UrlFetchApp: overrides.UrlFetchApp || {
      fetch: () => { throw new Error('UrlFetchApp not stubbed'); }
    },
    Utilities: {
      base64EncodeWebSafe: (s) => Buffer.from(s).toString('base64').replace(/\+/g, '-').replace(/\//g, '_'),
      computeRsaSha256Signature: () => new Uint8Array([1, 2, 3, 4]),
      sleep: () => {}
    },
    Session: { getActiveUser: () => ({ getEmail: () => 'actor@example.com' }) },
    validateEmail: (email) => ({ isValid: typeof email === 'string' && /.+@.+/.test(email), sanitized: email, errors: [] }),
    getCurrentEmail: () => 'actor@example.com',
    isAdministrator: () => false,
    getUserConfig: () => ({ success: true, config: {} }),
    executeWithRetry: (fn) => fn(),
    getCachedProperty: (k) => (k in propsStore ? propsStore[k] : null),
    clearPropertyCache: () => {},
    simpleHash: (s) => String(s).length,
    saveToCacheWithSizeCheck: (k, data) => { cache.put(k, JSON.stringify(data)); return true; },
    logError_: () => {},
    ...overrides
  };
  vm.createContext(context);
  const source = fs.readFileSync(path.resolve(__dirname, '../src/DatabaseCore.js'), 'utf8');
  vm.runInContext(source, context, { filename: 'DatabaseCore.js' });
  context._cache = cache;
  context._propsStore = propsStore;
  return context;
}

// =====================================================================
// parseServiceAccountCredsSoft_
// =====================================================================

test('parseServiceAccountCredsSoft_: rejects null / empty', () => {
  const ctx = loadCtx();
  assert.equal(ctx.parseServiceAccountCredsSoft_(null), null);
  assert.equal(ctx.parseServiceAccountCredsSoft_(''), null);
  assert.equal(ctx.parseServiceAccountCredsSoft_('not json'), null);
});

test('parseServiceAccountCredsSoft_: rejects missing required fields', () => {
  const ctx = loadCtx();
  assert.equal(ctx.parseServiceAccountCredsSoft_(JSON.stringify({ type: 'service_account' })), null);
  assert.equal(ctx.parseServiceAccountCredsSoft_(JSON.stringify({ client_email: 'a@b.c' })), null);
});

test('parseServiceAccountCredsSoft_: rejects invalid private key', () => {
  const ctx = loadCtx();
  const sa = { type: 'service_account', client_email: 'sa@x.iam', private_key: 'not-a-key' };
  assert.equal(ctx.parseServiceAccountCredsSoft_(JSON.stringify(sa)), null);
});

test('parseServiceAccountCredsSoft_: accepts valid SA', () => {
  const ctx = loadCtx();
  const parsed = ctx.parseServiceAccountCredsSoft_(makeSaJson('sa1@x.iam', 1));
  assert.equal(parsed.client_email, 'sa1@x.iam');
  assert.equal(parsed.type, 'service_account');
});

// =====================================================================
// getAllServiceAccounts_
// =====================================================================

test('getAllServiceAccounts_: returns [] when none configured', () => {
  const ctx = loadCtx();
  assert.deepEqual([...ctx.getAllServiceAccounts_()], []);
});

test('getAllServiceAccounts_: reads CREDS + CREDS_2..N in order', () => {
  const ctx = loadCtx({
    propsStore: {
      SERVICE_ACCOUNT_CREDS: makeSaJson('sa1@x.iam', 1),
      SERVICE_ACCOUNT_CREDS_2: makeSaJson('sa2@x.iam', 2),
      SERVICE_ACCOUNT_CREDS_3: makeSaJson('sa3@x.iam', 3)
    }
  });
  const pool = ctx.getAllServiceAccounts_();
  assert.equal(pool.length, 3);
  assert.equal(pool[0].client_email, 'sa1@x.iam');
  assert.equal(pool[1].client_email, 'sa2@x.iam');
  assert.equal(pool[2].client_email, 'sa3@x.iam');
});

test('getAllServiceAccounts_: skips invalid slots silently', () => {
  const ctx = loadCtx({
    propsStore: {
      SERVICE_ACCOUNT_CREDS: makeSaJson('sa1@x.iam', 1),
      SERVICE_ACCOUNT_CREDS_2: 'invalid json',
      SERVICE_ACCOUNT_CREDS_3: makeSaJson('sa3@x.iam', 3)
    }
  });
  const pool = ctx.getAllServiceAccounts_();
  assert.equal(pool.length, 2);
  assert.deepEqual([...pool.map((s) => s.client_email)], ['sa1@x.iam', 'sa3@x.iam']);
});

// =====================================================================
// pickServiceAccount_ (round-robin)
// =====================================================================

test('pickServiceAccount_: returns null when pool is empty', () => {
  const ctx = loadCtx();
  assert.equal(ctx.pickServiceAccount_(), null);
});

test('pickServiceAccount_: returns the only SA when pool has 1', () => {
  const ctx = loadCtx({
    propsStore: { SERVICE_ACCOUNT_CREDS: makeSaJson('only@x.iam', 1) }
  });
  const picked = ctx.pickServiceAccount_();
  assert.equal(picked.client_email, 'only@x.iam');
});

test('pickServiceAccount_: round-robin distributes evenly across 3 SAs', () => {
  const ctx = loadCtx({
    propsStore: {
      SERVICE_ACCOUNT_CREDS: makeSaJson('sa1@x.iam', 1),
      SERVICE_ACCOUNT_CREDS_2: makeSaJson('sa2@x.iam', 2),
      SERVICE_ACCOUNT_CREDS_3: makeSaJson('sa3@x.iam', 3)
    }
  });
  const counts = { 'sa1@x.iam': 0, 'sa2@x.iam': 0, 'sa3@x.iam': 0 };
  for (let i = 0; i < 30; i++) counts[ctx.pickServiceAccount_().client_email]++;
  assert.equal(counts['sa1@x.iam'], 10);
  assert.equal(counts['sa2@x.iam'], 10);
  assert.equal(counts['sa3@x.iam'], 10);
});

test('pickServiceAccount_: skips cooling SAs', () => {
  const ctx = loadCtx({
    propsStore: {
      SERVICE_ACCOUNT_CREDS: makeSaJson('sa1@x.iam', 1),
      SERVICE_ACCOUNT_CREDS_2: makeSaJson('sa2@x.iam', 2),
      SERVICE_ACCOUNT_CREDS_3: makeSaJson('sa3@x.iam', 3)
    }
  });
  ctx.markServiceAccountCoolingDown_('sa1@x.iam');
  const picks = new Set();
  for (let i = 0; i < 20; i++) picks.add(ctx.pickServiceAccount_().client_email);
  assert.ok(!picks.has('sa1@x.iam'), 'sa1 should be excluded while cooling');
  assert.ok(picks.has('sa2@x.iam'));
  assert.ok(picks.has('sa3@x.iam'));
});

test('pickServiceAccount_: when all cooling, falls back to full pool', () => {
  const ctx = loadCtx({
    propsStore: {
      SERVICE_ACCOUNT_CREDS: makeSaJson('sa1@x.iam', 1),
      SERVICE_ACCOUNT_CREDS_2: makeSaJson('sa2@x.iam', 2)
    }
  });
  ctx.markServiceAccountCoolingDown_('sa1@x.iam');
  ctx.markServiceAccountCoolingDown_('sa2@x.iam');
  const picked = ctx.pickServiceAccount_();
  assert.ok(picked && picked.client_email, 'must still return some SA, not null');
});

test('recordSaPick_: increments per-SA counter for admin visibility', () => {
  const ctx = loadCtx({
    propsStore: { SERVICE_ACCOUNT_CREDS: makeSaJson('sa1@x.iam', 1) }
  });
  for (let i = 0; i < 5; i++) ctx.pickServiceAccount_();
  const usage = ctx.getServiceAccountUsage();
  assert.equal(usage.success, true);
  assert.equal(usage.pool.length, 1);
  assert.equal(usage.pool[0].clientEmail, 'sa1@x.iam');
  assert.equal(usage.pool[0].picks, 5);
});

// =====================================================================
// makeProxyAuthResolver_
// =====================================================================

test('makeProxyAuthResolver_: returns initial when not cooling', () => {
  const ctx = loadCtx({
    propsStore: {
      SERVICE_ACCOUNT_CREDS: makeSaJson('sa1@x.iam', 1),
      SERVICE_ACCOUNT_CREDS_2: makeSaJson('sa2@x.iam', 2)
    }
  });
  const resolver = ctx.makeProxyAuthResolver_('ss-x', 'tok1', 'sa1@x.iam');
  const auth = resolver();
  assert.equal(auth.token, 'tok1');
  assert.equal(auth.saEmail, 'sa1@x.iam');
});

test('makeProxyAuthResolver_: switches to fresh SA when initial is cooling', () => {
  const ctx = loadCtx({
    propsStore: {
      SERVICE_ACCOUNT_CREDS: makeSaJson('sa1@x.iam', 1),
      SERVICE_ACCOUNT_CREDS_2: makeSaJson('sa2@x.iam', 2)
    }
  });
  // Pre-cache sa2 token so resolver doesn't need JWT exchange.
  ctx._cache.put('sa_token:sa2@x.iam', 'tok2', 3000);
  ctx.markServiceAccountCoolingDown_('sa1@x.iam');
  const resolver = ctx.makeProxyAuthResolver_('ss-x', 'tok1', 'sa1@x.iam');
  const auth = resolver();
  assert.equal(auth.saEmail, 'sa2@x.iam', 'must swap to non-cooling SA');
  assert.equal(auth.token, 'tok2');
});

// =====================================================================
// addServiceAccountToPool / removeServiceAccountFromPool
// =====================================================================

test('addServiceAccountToPool: rejects invalid JSON', () => {
  const ctx = loadCtx();
  const result = ctx.addServiceAccountToPool('not json');
  assert.equal(result.success, false);
  assert.equal(result.error, 'INVALID_SA_JSON');
});

test('addServiceAccountToPool: rejects duplicate', () => {
  const ctx = loadCtx({
    propsStore: { SERVICE_ACCOUNT_CREDS: makeSaJson('sa1@x.iam', 1) }
  });
  const result = ctx.addServiceAccountToPool(makeSaJson('sa1@x.iam', 99));
  assert.equal(result.success, false);
  assert.equal(result.error, 'ALREADY_REGISTERED');
});

test('addServiceAccountToPool: writes to first empty slot', () => {
  const propsStore = {
    SERVICE_ACCOUNT_CREDS: makeSaJson('sa1@x.iam', 1)
  };
  const ctx = loadCtx({ propsStore });
  // Stub Drive + token exchange to avoid network failures.
  ctx.DriveApp = { getFileById: () => ({ addEditor: () => {} }) };
  ctx.UrlFetchApp = {
    fetch: () => ({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({ access_token: 'tok2' })
    })
  };

  const result = ctx.addServiceAccountToPool(makeSaJson('sa2@x.iam', 2));
  assert.equal(result.success, true);
  assert.equal(result.slot, 'SERVICE_ACCOUNT_CREDS_2');
  assert.equal(result.index, 2);
  assert.equal(result.clientEmail, 'sa2@x.iam');
  // Property store should have the new key.
  assert.ok(propsStore.SERVICE_ACCOUNT_CREDS_2);
});

test('removeServiceAccountFromPool: refuses primary', () => {
  const ctx = loadCtx({
    propsStore: { SERVICE_ACCOUNT_CREDS: makeSaJson('sa1@x.iam', 1) }
  });
  const result = ctx.removeServiceAccountFromPool('SERVICE_ACCOUNT_CREDS');
  assert.equal(result.success, false);
  assert.equal(result.error, 'PRIMARY_LOCKED');
});

test('removeServiceAccountFromPool: removes secondary', () => {
  const propsStore = {
    SERVICE_ACCOUNT_CREDS: makeSaJson('sa1@x.iam', 1),
    SERVICE_ACCOUNT_CREDS_2: makeSaJson('sa2@x.iam', 2)
  };
  const ctx = loadCtx({ propsStore });
  const result = ctx.removeServiceAccountFromPool('SERVICE_ACCOUNT_CREDS_2');
  assert.equal(result.success, true);
  assert.equal(result.removed, true);
  assert.equal(propsStore.SERVICE_ACCOUNT_CREDS_2, undefined);
});

test('removeServiceAccountFromPool: rejects unknown slot key format', () => {
  const ctx = loadCtx();
  const result = ctx.removeServiceAccountFromPool('NOT_A_VALID_KEY');
  assert.equal(result.success, false);
  assert.equal(result.error, 'INVALID_SLOT');
});

// =====================================================================
// listServiceAccountPool
// =====================================================================

test('listServiceAccountPool: returns slot meta without private_key', () => {
  const ctx = loadCtx({
    propsStore: {
      SERVICE_ACCOUNT_CREDS: makeSaJson('sa1@x.iam', 1),
      SERVICE_ACCOUNT_CREDS_2: makeSaJson('sa2@x.iam', 2)
    }
  });
  const result = ctx.listServiceAccountPool();
  assert.equal(result.success, true);
  assert.equal(result.poolSize, 2);
  assert.equal(result.slots.length, 2);
  assert.equal(result.slots[0].role, 'primary');
  assert.equal(result.slots[0].clientEmail, 'sa1@x.iam');
  assert.equal(result.slots[1].role, 'secondary');
  assert.equal(result.slots[1].clientEmail, 'sa2@x.iam');
  // private_key must not leak.
  assert.equal(result.slots[0].private_key, undefined);
});

// =====================================================================
// getServiceAccountAccessToken_ (2-tier cache)
// =====================================================================

test('getServiceAccountAccessToken_: ScriptCache hit avoids JWT exchange', () => {
  const ctx = loadCtx();
  ctx._cache.put('sa_token:sa1@x.iam', 'cached-token', 3000);
  let fetchCalled = false;
  ctx.UrlFetchApp = {
    fetch: () => { fetchCalled = true; return { getResponseCode: () => 200, getContentText: () => '{}' }; }
  };
  const sa = { client_email: 'sa1@x.iam', private_key: 'k' };
  assert.equal(ctx.getServiceAccountAccessToken_(sa), 'cached-token');
  assert.equal(fetchCalled, false, 'no JWT exchange when token is cached');
});

test('getServiceAccountAccessToken_: cache miss performs JWT exchange and caches', () => {
  const ctx = loadCtx();
  let fetchCount = 0;
  ctx.UrlFetchApp = {
    fetch: () => {
      fetchCount++;
      return { getResponseCode: () => 200, getContentText: () => JSON.stringify({ access_token: 'fresh-tok' }) };
    }
  };
  const sa = { client_email: 'sa1@x.iam', private_key: 'k' };
  assert.equal(ctx.getServiceAccountAccessToken_(sa), 'fresh-tok');
  // Second call should hit in-memory cache.
  assert.equal(ctx.getServiceAccountAccessToken_(sa), 'fresh-tok');
  assert.equal(fetchCount, 1, 'JWT exchange called once');
});
