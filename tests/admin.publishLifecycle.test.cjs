/**
 * Tests for the unified publish lifecycle in AdminApis.js
 *
 * The 4 official lifecycle ops (no other code should write isPublished directly):
 *   - publishApp (SystemController.js, tested in system.publishApp.test.cjs)
 *   - republishMyBoard           : owner re-publish
 *   - unpublishBoard             : owner / admin unpublish
 *   - toggleUserBoardStatus      : admin toggle of another user's board
 *
 * Invariants verified here:
 *   1. publishedAt reflects current state (now when published, null when not)
 *   2. etag mismatch is detected uniformly across all 3 helpers
 *   3. Authorization is enforced uniformly (own board or admin)
 *   4. Response shape is unified
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('path');
const vm = require('vm');
const { gasResponseStubs } = require('./_helpers.cjs');

function loadAdminContext(overrides = {}) {
  const savedConfigs = {};
  const users = overrides.users || [
    { userId: 'owner1', userEmail: 'owner@example.com', configJson: '{}' },
    { userId: 'other1', userEmail: 'other@example.com', configJson: '{}' }
  ];

  const baseContext = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    ...gasResponseStubs(),
    requireAdmin: () => ({ email: 'owner@example.com', isAdmin: true }),
    getCurrentEmail: () => 'owner@example.com',
    isAdministrator: () => true,
    getWebAppUrl: () => 'https://example.com/exec',
    findUserById: (id) => users.find(u => u.userId === id),
    findUserByEmail: (email) => users.find(u => u.userEmail === email),
    // Why getUserConfig (not getConfigOrDefault): __applyPublishStateChange uses
    //   getUserConfig envelope to detect 429/cache-miss and refuse destructive overwrite.
    //   (旧 getConfigOrDefault stub は v2757 で不要に。)
    getUserConfig: (userId) => ({
      success: true,
      config: savedConfigs[userId] || { isPublished: false, publishedAt: null }
    }),
    saveUserConfig: (userId, config) => {
      savedConfigs[userId] = { ...config, etag: 'etag_' + Date.now() + '_' + Math.random() };
      return { success: true, etag: savedConfigs[userId].etag, config: savedConfigs[userId] };
    },
    updateUser: () => ({ success: true }),
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: () => null,
        setProperty: () => {},
        deleteProperty: () => {}
      })
    },
    getCachedProperty: () => null,
    __savedConfigs: savedConfigs,
    ...overrides
  };
  // Apply overrides AFTER base so they can override functions
  Object.assign(baseContext, overrides);
  vm.createContext(baseContext);
  const source = fs.readFileSync(path.resolve(__dirname, '../src/AdminApis.js'), 'utf8');
  vm.runInContext(source, baseContext, { filename: 'AdminApis.js' });
  return baseContext;
}

// =====================================================================
// publishedAt セマンティクス: 常に現在の状態を反映
// =====================================================================

test('unpublishBoard: sets isPublished=false and publishedAt=null', () => {
  const ctx = loadAdminContext();
  // Seed as published
  ctx.__savedConfigs['owner1'] = { isPublished: true, publishedAt: '2020-01-01T00:00:00.000Z' };

  const result = ctx.unpublishBoard('owner1');
  assert.equal(result.success, true);
  assert.equal(result.isPublished, false);
  assert.equal(result.publishedAt, null);
  assert.equal(ctx.__savedConfigs['owner1'].isPublished, false);
  assert.equal(ctx.__savedConfigs['owner1'].publishedAt, null);
});

test('republishMyBoard: sets isPublished=true and publishedAt=now', () => {
  const ctx = loadAdminContext();
  ctx.__savedConfigs['owner1'] = { isPublished: false, publishedAt: null };

  const result = ctx.republishMyBoard();
  assert.equal(result.success, true);
  assert.equal(result.isPublished, true);
  assert.ok(result.publishedAt, 'publishedAt must be set');
  // ISO 8601 format check
  assert.match(result.publishedAt, /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/);
});

test('toggleUserBoardStatus: false → true → false alternates state correctly', () => {
  const ctx = loadAdminContext();
  ctx.__savedConfigs['other1'] = { isPublished: false, publishedAt: null };

  const r1 = ctx.toggleUserBoardStatus('other1');
  assert.equal(r1.success, true);
  assert.equal(r1.isPublished, true);
  assert.ok(r1.publishedAt);

  const r2 = ctx.toggleUserBoardStatus('other1');
  assert.equal(r2.success, true);
  assert.equal(r2.isPublished, false);
  assert.equal(r2.publishedAt, null);
});

// =====================================================================
// etag conflict detection は 3 関数すべてで一貫
// =====================================================================

test('unpublishBoard: rejects with etag_mismatch when sourceEtag differs', () => {
  const ctx = loadAdminContext();
  ctx.__savedConfigs['owner1'] = {
    isPublished: true,
    publishedAt: '2020-01-01T00:00:00.000Z',
    etag: 'current_etag_v2'
  };

  const result = ctx.unpublishBoard('owner1', { sourceEtag: 'old_etag_v1' });
  assert.equal(result.success, false);
  assert.equal(result.error, 'etag_mismatch');
  assert.equal(result.currentEtag, 'current_etag_v2');
  // 状態は変わっていないこと
  assert.equal(ctx.__savedConfigs['owner1'].isPublished, true);
});

test('republishMyBoard: rejects with etag_mismatch when sourceEtag differs', () => {
  const ctx = loadAdminContext();
  ctx.__savedConfigs['owner1'] = {
    isPublished: false,
    publishedAt: null,
    etag: 'current_etag_v2'
  };

  const result = ctx.republishMyBoard({ sourceEtag: 'old_etag_v1' });
  assert.equal(result.success, false);
  assert.equal(result.error, 'etag_mismatch');
});

test('toggleUserBoardStatus: rejects with etag_mismatch when sourceEtag differs', () => {
  const ctx = loadAdminContext();
  ctx.__savedConfigs['other1'] = {
    isPublished: false,
    publishedAt: null,
    etag: 'current_etag_v2'
  };

  const result = ctx.toggleUserBoardStatus('other1', { sourceEtag: 'old_etag_v1' });
  assert.equal(result.success, false);
  assert.equal(result.error, 'etag_mismatch');
});

test('sourceEtag 未指定でも saveUserConfig の etag_mismatch を surface する (RMW窓の並行更新)', () => {
  // sourceEtag を渡さない caller でも、 read(L158)→save の間に別 writer が config を
  //   変えていれば saveUserConfig が継承 etag で conflict を返す。 これを汎用エラーに
  //   埋もれさせず構造化して返すこと (publishApp と対称)。
  const ctx = loadAdminContext({
    saveUserConfig: () => ({
      success: false,
      error: 'etag_mismatch',
      message: 'Configuration has been modified by another user',
      currentConfig: { isPublished: true, etag: 'newer_etag' }
    })
  });
  ctx.__savedConfigs['owner1'] = { isPublished: true, publishedAt: null, etag: 'read_etag' };

  const result = ctx.unpublishBoard('owner1'); // sourceEtag 無し
  assert.equal(result.success, false);
  assert.equal(result.error, 'etag_mismatch');
  assert.equal(result.currentEtag, 'newer_etag');
});

test('unpublishBoard: succeeds when sourceEtag matches', () => {
  const ctx = loadAdminContext();
  ctx.__savedConfigs['owner1'] = {
    isPublished: true,
    publishedAt: '2020-01-01T00:00:00.000Z',
    etag: 'current_etag_v2'
  };

  const result = ctx.unpublishBoard('owner1', { sourceEtag: 'current_etag_v2' });
  assert.equal(result.success, true);
  assert.equal(result.isPublished, false);
});

test('unpublishBoard: succeeds when sourceEtag omitted (no concurrent edit detection)', () => {
  const ctx = loadAdminContext();
  ctx.__savedConfigs['owner1'] = {
    isPublished: true,
    publishedAt: '2020-01-01T00:00:00.000Z',
    etag: 'current_etag_v2'
  };

  // No sourceEtag → no etag check → succeed (back-compat with existing callers)
  const result = ctx.unpublishBoard('owner1');
  assert.equal(result.success, true);
});

// =====================================================================
// 認可: 自分のボード or 管理者のみ
// =====================================================================

test('unpublishBoard: rejects non-admin attempting to unpublish other user', () => {
  const ctx = loadAdminContext({
    isAdministrator: () => false,
    getCurrentEmail: () => 'student@example.com',
    findUserByEmail: () => null
  });
  ctx.__savedConfigs['owner1'] = { isPublished: true, publishedAt: '2020-01-01T00:00:00.000Z' };

  const result = ctx.unpublishBoard('owner1');
  assert.equal(result.success, false);
  assert.match(result.message, /権限/);
});

// v2855+: ボード SS の editor 共有者 (collaborator) は unpublish (緊急停止) のみ許可。
//   publish / toggle は引き続き owner / admin のみ。
test('unpublishBoard: collaborator (SS editor) can unpublish other user board', () => {
  const ctx = loadAdminContext({
    isAdministrator: () => false,
    getCurrentEmail: () => 'collab@example.com',
    findUserByEmail: () => null,
    isBoardCollaborator: (user, email) => email === 'collab@example.com'
  });
  ctx.__savedConfigs['owner1'] = { isPublished: true, publishedAt: '2020-01-01T00:00:00.000Z' };

  const result = ctx.unpublishBoard('owner1');
  assert.equal(result.success, true);
  assert.equal(result.isPublished, false);
  assert.equal(result.publishedAt, null);
});

test('republishMyBoard: collaborator CANNOT publish other user board (write op blocked)', () => {
  const ctx = loadAdminContext({
    isAdministrator: () => false,
    getCurrentEmail: () => 'collab@example.com',
    findUserByEmail: () => null,
    isBoardCollaborator: (user, email) => email === 'collab@example.com'
  });
  ctx.__savedConfigs['owner1'] = { isPublished: false, publishedAt: null };

  // republishMyBoard は自分のボードに対する操作なので、collaborator は other user に対しては
  // findUserByEmail (caller resolution) で fail → user not found
  const result = ctx.republishMyBoard();
  assert.equal(result.success, false);
});

test('toggleUserBoardStatus: rejects non-admin caller', () => {
  const ctx = loadAdminContext({
    isAdministrator: () => false,
    getCurrentEmail: () => 'other@example.com'
  });
  ctx.__savedConfigs['other1'] = { isPublished: false, publishedAt: null };

  const result = ctx.toggleUserBoardStatus('other1');
  assert.equal(result.success, false);
  assert.match(result.message, /管理者/);
});

test('toggleUserBoardStatus: requires targetUserId', () => {
  const ctx = loadAdminContext();
  const result = ctx.toggleUserBoardStatus(null);
  assert.equal(result.success, false);
  assert.match(result.message, /targetUserId/);
});

// =====================================================================
// レスポンス形の統一
// =====================================================================

test('all 3 lifecycle helpers return unified shape with publishedAt + isPublished + redirectUrl + etag', () => {
  const ctx = loadAdminContext();
  ctx.__savedConfigs['owner1'] = { isPublished: false, publishedAt: null };

  const checkShape = (r) => {
    assert.equal(r.success, true);
    assert.ok('isPublished' in r, 'has isPublished');
    assert.ok('boardPublished' in r, 'has boardPublished (back-compat)');
    assert.ok('publishedAt' in r, 'has publishedAt');
    assert.ok('redirectUrl' in r && r.redirectUrl.includes('?mode=view'), 'has redirectUrl');
    assert.ok('etag' in r, 'has etag field');
    assert.ok('userId' in r, 'has userId');
    assert.equal(r.boardPublished, r.isPublished, 'boardPublished === isPublished');
  };

  checkShape(ctx.republishMyBoard());
  checkShape(ctx.unpublishBoard('owner1'));
  checkShape(ctx.toggleUserBoardStatus('other1'));
});

// =====================================================================
// dispatcher 経由でも同じ動作になること
// =====================================================================

test('dispatch: unpublishBoard via dispatcher works', () => {
  const ctx = loadAdminContext();
  ctx.__savedConfigs['owner1'] = { isPublished: true, publishedAt: '2020-01-01T00:00:00.000Z' };

  const result = ctx.dispatchAdminOperation('unpublishBoard', { targetUserId: 'owner1' });
  assert.equal(result.success, true);
  assert.equal(result.isPublished, false);
});

test('dispatch: republishMyBoard via dispatcher works', () => {
  const ctx = loadAdminContext();
  ctx.__savedConfigs['owner1'] = { isPublished: false, publishedAt: null };

  const result = ctx.dispatchAdminOperation('republishMyBoard', {});
  assert.equal(result.success, true);
  assert.equal(result.isPublished, true);
});

test('dispatch: toggleUserBoard with etag forwards to lifecycle helper', () => {
  const ctx = loadAdminContext();
  ctx.__savedConfigs['other1'] = { isPublished: false, publishedAt: null, etag: 'cur' };

  // wrong etag rejected
  const r1 = ctx.dispatchAdminOperation('toggleUserBoard', { targetUserId: 'other1', etag: 'wrong' });
  assert.equal(r1.success, false);
  assert.equal(r1.error, 'etag_mismatch');

  // correct etag passes
  const r2 = ctx.dispatchAdminOperation('toggleUserBoard', { targetUserId: 'other1', etag: 'cur' });
  assert.equal(r2.success, true);
});
