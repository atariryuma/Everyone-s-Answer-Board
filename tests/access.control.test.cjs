// src/AccessControl.js (ボード共同編集者 collaborator 認可) のユニットテスト。
// vm.createContext で AccessControl.js を読み込み、GAS API (CacheService / UrlFetchApp /
// Utilities / SA pool helper / getUserConfig) を最小 stub で再現する。
//
// AccessControl.js の表面:
//   - isBoardCollaborator(targetUser, viewerEmail)  … 公開エントリ (guard + config 取得 + Drive 委譲)
//   - __hasEditorPermissionViaDrive_(fileId, emailNorm) … Drive permissions.list + cache (security 中核)
//   - __emailHash_(email) … SHA-256 cache key (衝突で他人の editor 判定を継承しないため)
//
// 修正済みバグ (v2855 collaborator 認可): isBoardCollaborator は getUserConfig() の
//   envelope {success, config:{spreadsheetId}} を .config で unwrap する。以前は envelope 直読みで
//   ssId が常に undefined になり editor でも false を返していた (fail-closed = 権限過剰でなく不足)。
//   corrupted config では fail-closed (認可を開かない)。

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const crypto = require('node:crypto');

function createMockCache(initial = {}) {
  const store = new Map(Object.entries(initial));
  const puts = [];
  return {
    get: (k) => (store.has(k) ? store.get(k) : null),
    put: (k, v, ttl) => { store.set(k, v); puts.push({ k, v, ttl }); },
    remove: (k) => { store.delete(k); },
    _store: store,
    _puts: puts
  };
}

// GAS Utilities.computeDigest 互換: SHA-256 を signed byte 配列で返す (AccessControl は & 0xff で正規化)。
function computeDigestSha256(str) {
  const buf = crypto.createHash('sha256').update(String(str), 'utf8').digest();
  return Array.from(buf).map((b) => (b > 127 ? b - 256 : b)); // GAS の signed byte を模す
}

// UrlFetchApp.fetch の HTTPResponse stub。
function mockResponse(code, body) {
  return { getResponseCode: () => code, getContentText: () => (typeof body === 'string' ? body : JSON.stringify(body)) };
}

function drivePermissionsBody(entries) {
  return { permissions: entries.map((e) => ({ emailAddress: e.email, role: e.role })) };
}

function loadCtx(overrides = {}) {
  const cache = overrides.cache || createMockCache();
  const fetchCalls = [];
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    CacheService: { getScriptCache: () => cache },
    pickServiceAccount_: overrides.pickServiceAccount_ || (() => ({ client_email: 'sa@x.iam.gserviceaccount.com' })),
    getServiceAccountAccessToken_: overrides.getServiceAccountAccessToken_ || (() => 'fake-token'),
    getUserConfig: overrides.getUserConfig || (() => ({ success: true, config: { spreadsheetId: 'SS_DEFAULT' } })),
    UrlFetchApp: {
      fetch: overrides.fetch || ((url, opts) => { fetchCalls.push({ url, opts }); return mockResponse(200, drivePermissionsBody([])); })
    },
    Utilities: {
      computeDigest: (_algo, str) => computeDigestSha256(str),
      DigestAlgorithm: { SHA_256: 'SHA_256' },
      Charset: { UTF_8: 'UTF_8' }
    },
    logError_: () => {}
  };
  vm.createContext(context);
  const source = fs.readFileSync(path.resolve(__dirname, '../src/AccessControl.js'), 'utf8');
  vm.runInContext(source, context, { filename: 'AccessControl.js' });
  context._cache = cache;
  context._fetchCalls = fetchCalls;
  return context;
}

// =====================================================================
// __emailHash_ : cache key の衝突耐性 (他人の editor 判定を継承しないための security fix)
// =====================================================================

test('__emailHash_: 決定的 — 同じメールは同じ hash', () => {
  const ctx = loadCtx();
  assert.equal(ctx.__emailHash_('teacher@naha.ed.jp'), ctx.__emailHash_('teacher@naha.ed.jp'));
});

test('__emailHash_: 異なるメールは異なる hash (衝突回避) かつ 64 桁 hex', () => {
  const ctx = loadCtx();
  const a = ctx.__emailHash_('a@x.jp');
  const b = ctx.__emailHash_('b@x.jp');
  assert.notEqual(a, b);
  assert.match(a, /^[0-9a-f]{64}$/); // SHA-256 = 32 byte = 64 hex
});

// =====================================================================
// __hasEditorPermissionViaDrive_ : Drive permissions.list + cache (security 中核)
// =====================================================================

test('__hasEditorPermissionViaDrive_: writer ロールの editor は true', () => {
  const ctx = loadCtx({
    fetch: () => mockResponse(200, drivePermissionsBody([
      { email: 'editor@x.jp', role: 'writer' },
      { email: 'other@x.jp', role: 'reader' }
    ]))
  });
  assert.equal(ctx.__hasEditorPermissionViaDrive_('SS1', 'editor@x.jp'), true);
});

test('__hasEditorPermissionViaDrive_: owner ロールも editor 扱いで true', () => {
  const ctx = loadCtx({
    fetch: () => mockResponse(200, drivePermissionsBody([{ email: 'boss@x.jp', role: 'owner' }]))
  });
  assert.equal(ctx.__hasEditorPermissionViaDrive_('SS1', 'boss@x.jp'), true);
});

test('__hasEditorPermissionViaDrive_: reader のみ / 不在は false', () => {
  const ctx = loadCtx({
    fetch: () => mockResponse(200, drivePermissionsBody([{ email: 'viewer@x.jp', role: 'reader' }]))
  });
  assert.equal(ctx.__hasEditorPermissionViaDrive_('SS1', 'viewer@x.jp'), false);
  assert.equal(ctx.__hasEditorPermissionViaDrive_('SS1', 'nobody@x.jp'), false);
});

test('__hasEditorPermissionViaDrive_: Drive が非200 のときは fail-closed で false', () => {
  const ctx = loadCtx({ fetch: () => mockResponse(403, 'Forbidden') });
  assert.equal(ctx.__hasEditorPermissionViaDrive_('SS1', 'editor@x.jp'), false);
});

test('__hasEditorPermissionViaDrive_: SA が引けない / token 無しは fail-closed で false', () => {
  const noSa = loadCtx({ pickServiceAccount_: () => null });
  assert.equal(noSa.__hasEditorPermissionViaDrive_('SS1', 'editor@x.jp'), false);
  const noToken = loadCtx({ getServiceAccountAccessToken_: () => null });
  assert.equal(noToken.__hasEditorPermissionViaDrive_('SS1', 'editor@x.jp'), false);
});

test('__hasEditorPermissionViaDrive_: fetch が例外を投げても fail-closed で false (認可を開かない)', () => {
  const ctx = loadCtx({ fetch: () => { throw new Error('network down'); } });
  assert.equal(ctx.__hasEditorPermissionViaDrive_('SS1', 'editor@x.jp'), false);
});

test('__hasEditorPermissionViaDrive_: cache hit ("1"/"0") は Drive を叩かず即返す', () => {
  const okCache = loadCtx({
    fetch: () => { throw new Error('should not fetch on cache hit'); }
  });
  // 該当 key を直接 seed する (key 生成は本物の __emailHash_ に合わせる)
  const key = 'collab:SS1:' + okCache.__emailHash_('editor@x.jp');
  okCache._cache._store.set(key, '1');
  assert.equal(okCache.__hasEditorPermissionViaDrive_('SS1', 'editor@x.jp'), true);
  assert.equal(okCache._fetchCalls.length, 0);

  const noCache = loadCtx({ fetch: () => { throw new Error('should not fetch'); } });
  const key2 = 'collab:SS1:' + noCache.__emailHash_('editor@x.jp');
  noCache._cache._store.set(key2, '0');
  assert.equal(noCache.__hasEditorPermissionViaDrive_('SS1', 'editor@x.jp'), false);
});

test('__hasEditorPermissionViaDrive_: 許可は 120s / 拒否は 600s で cache 書込み', () => {
  const okCtx = loadCtx({
    fetch: () => mockResponse(200, drivePermissionsBody([{ email: 'editor@x.jp', role: 'writer' }]))
  });
  okCtx.__hasEditorPermissionViaDrive_('SS1', 'editor@x.jp');
  const okPut = okCtx._cache._puts.at(-1);
  assert.equal(okPut.v, '1');
  assert.equal(okPut.ttl, 120); // 付与は短命 (revoke 後の stale grant 窓を短く)

  const noCtx = loadCtx({ fetch: () => mockResponse(200, drivePermissionsBody([])) });
  noCtx.__hasEditorPermissionViaDrive_('SS1', 'nobody@x.jp');
  const noPut = noCtx._cache._puts.at(-1);
  assert.equal(noPut.v, '0');
  assert.equal(noPut.ttl, 600); // 拒否は長命
});

test('__hasEditorPermissionViaDrive_: cache key は fileId + emailHash で分離 (別メールは別 key)', () => {
  const ctx = loadCtx();
  const kEditor = 'collab:SS1:' + ctx.__emailHash_('editor@x.jp');
  const kOther = 'collab:SS1:' + ctx.__emailHash_('other@x.jp');
  assert.notEqual(kEditor, kOther);
});

// =====================================================================
// isBoardCollaborator : guard 節 (null / owner / 正規化)
// =====================================================================

test('isBoardCollaborator: targetUser / viewerEmail が欠けると false', () => {
  const ctx = loadCtx();
  assert.equal(ctx.isBoardCollaborator(null, 'a@x.jp'), false);
  assert.equal(ctx.isBoardCollaborator({ userId: 'u1', userEmail: 'o@x.jp' }, ''), false);
  assert.equal(ctx.isBoardCollaborator({ userId: 'u1', userEmail: 'o@x.jp' }, null), false);
});

test('isBoardCollaborator: owner 自身は collaborator でない (case-insensitive で false)', () => {
  const ctx = loadCtx({ getUserConfig: () => { throw new Error('should not reach config for owner'); } });
  const target = { userId: 'u1', userEmail: 'Owner@Naha.ed.jp' };
  assert.equal(ctx.isBoardCollaborator(target, 'owner@naha.ed.jp'), false);
  assert.equal(ctx.isBoardCollaborator(target, 'OWNER@NAHA.ED.JP'), false);
});

test('isBoardCollaborator: getUserConfig が throw しても fail-closed で false', () => {
  const ctx = loadCtx({ getUserConfig: () => { throw new Error('config load boom'); } });
  assert.equal(ctx.isBoardCollaborator({ userId: 'u1', userEmail: 'o@x.jp' }, 'editor@x.jp'), false);
});

test('isBoardCollaborator: config に spreadsheetId が無ければ false', () => {
  const ctx = loadCtx({ getUserConfig: () => ({ success: true, config: {} }) });
  assert.equal(ctx.isBoardCollaborator({ userId: 'u1', userEmail: 'o@x.jp' }, 'editor@x.jp'), false);
});

// happy-path (v2855 collaborator 認可): getUserConfig は envelope {success, config:{spreadsheetId}}
// を返す。AccessControl は .config で unwrap し、ボード SS に Drive editor (writer/owner) 権限を
// 持つメールを collaborator として認可する。以前は envelope 直読みで ssId が undefined になり
// editor でも false だった (修正済み)。
test('isBoardCollaborator: envelope を unwrap し Drive editor を collaborator として true', () => {
  const ctx = loadCtx({
    // 本物の getUserConfig と同じ envelope 形状
    getUserConfig: () => ({ success: true, config: { spreadsheetId: 'SS1' } }),
    // editor として Drive 上は writer 権限を持つ
    fetch: () => mockResponse(200, drivePermissionsBody([{ email: 'editor@x.jp', role: 'writer' }]))
  });
  assert.equal(ctx.isBoardCollaborator({ userId: 'u1', userEmail: 'owner@x.jp' }, 'editor@x.jp'), true);
});

test('isBoardCollaborator: SS に editor 権限が無い viewer は false (unwrap 後も権限判定は Drive 依存)', () => {
  const ctx = loadCtx({
    getUserConfig: () => ({ success: true, config: { spreadsheetId: 'SS1' } }),
    // editor@x.jp のみ writer。viewer@x.jp は permission list に不在
    fetch: () => mockResponse(200, drivePermissionsBody([{ email: 'editor@x.jp', role: 'writer' }]))
  });
  assert.equal(ctx.isBoardCollaborator({ userId: 'u1', userEmail: 'owner@x.jp' }, 'viewer@x.jp'), false);
});

test('isBoardCollaborator: getUserConfig が success:false なら false (config 取得失敗)', () => {
  const ctx = loadCtx({
    getUserConfig: () => ({ success: false, config: { spreadsheetId: 'SS1' }, message: 'User not found' }),
    fetch: () => mockResponse(200, drivePermissionsBody([{ email: 'editor@x.jp', role: 'writer' }]))
  });
  assert.equal(ctx.isBoardCollaborator({ userId: 'u1', userEmail: 'owner@x.jp' }, 'editor@x.jp'), false);
});

// corrupted config (JSON parse 失敗で default に fallback) は認可判定では fail-closed。
// stale な spreadsheetId が残っていても collaborator を認可しない (破損 config を信頼しない)。
// AdminApis の publish 状態変更が corrupted を拒否するのと同じ扱い。
test('isBoardCollaborator: corrupted config は spreadsheetId が残っていても fail-closed で false', () => {
  const ctx = loadCtx({
    getUserConfig: () => ({ success: true, corrupted: true, config: { spreadsheetId: 'SS1' } }),
    fetch: () => mockResponse(200, drivePermissionsBody([{ email: 'editor@x.jp', role: 'writer' }]))
  });
  assert.equal(ctx.isBoardCollaborator({ userId: 'u1', userEmail: 'owner@x.jp' }, 'editor@x.jp'), false);
});
