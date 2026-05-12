const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('path');
const vm = require('vm');

/**
 * Tests for new adminApi operations: findUser, getUserConfig, exportConfigs,
 * setUserConfig, bulkSetUserConfig, runColumnAnalysis, previewBoard.
 *
 * Why: これらが CLI から本番 config を変更できる新しい入口なので、validation
 *      と sanitize の動作を確実にする。とくに setUserConfig の deep-merge と
 *      protected フィールドの拒否が要。
 */

function loadAdminContext(overrides = {}) {
  let savedConfigs = new Map(); // userId → config object (last saved)

  const baseUsers = [
    {
      userId: 'u1', userEmail: 'a@example.com', isActive: true,
      createdAt: '2026-01-01', lastModified: '2026-04-01',
      configJson: JSON.stringify({
        userId: 'u1',
        spreadsheetId: 'sheet1',
        isPublished: true,
        displaySettings: { showNames: false, boardMode: 'auto', pageSize: 20 },
        columnMapping: { answer: 3, reason: 4 }
      })
    },
    {
      userId: 'u2', userEmail: 'b@example.com', isActive: true,
      createdAt: '2026-01-02', lastModified: '2026-04-02',
      configJson: JSON.stringify({
        userId: 'u2',
        spreadsheetId: 'sheet2',
        isPublished: false,
        displaySettings: { showNames: true, boardMode: 'board' }
      })
    },
    {
      userId: 'u3', userEmail: 'c@example.com', isActive: false,
      createdAt: '2026-01-03', lastModified: '2026-04-03',
      configJson: null
    }
  ];

  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    createErrorResponse: (msg, data, extra) => ({ success: false, message: msg, error: msg, ...extra }),
    createSuccessResponse: (msg, data) => ({ success: true, message: msg, ...(data && { data }) }),
    createAdminRequiredError: () => ({ success: false, message: 'admin required' }),
    createAuthError: () => ({ success: false, message: 'auth required' }),
    createUserNotFoundError: () => ({ success: false, message: 'user not found' }),
    createExceptionResponse: (e) => ({ success: false, message: e.message }),
    getCurrentEmail: () => 'admin@example.com',
    isAdministrator: () => true,
    findUserByEmail: (email) => baseUsers.find(u => u.userEmail === email) || null,
    findUserById: (id) => baseUsers.find(u => u.userId === id) || null,
    getAllUsers: () => baseUsers,
    updateUser: () => ({ success: true }),
    // getUserConfig returns parsed config
    getUserConfig: (userId) => {
      const u = baseUsers.find(x => x.userId === userId);
      if (!u) return { success: false, message: 'not found' };
      try {
        return { success: true, config: u.configJson ? JSON.parse(u.configJson) : {} };
      } catch (_) {
        return { success: true, config: {} };
      }
    },
    saveUserConfig: (userId, config, options) => {
      savedConfigs.set(userId, { config, options });
      // 既存配列も更新（テスト中での連続呼び出しに対応）
      const u = baseUsers.find(x => x.userId === userId);
      if (u) u.configJson = JSON.stringify(config);
      return { success: true, message: 'saved' };
    },
    requireAdmin: () => ({ email: 'admin@example.com', isAdmin: true }),
    getConfigOrDefault: () => ({}),
    getBatchedAdminAuth: () => ({ success: true, email: 'admin@example.com', isAdmin: true }),
    testSystemDiagnosis: () => ({ success: true }),
    performAutoRepair: () => ({ success: true }),
    forceUrlSystemReset: () => ({ success: true }),
    getPerformanceMetrics: (c) => ({ success: true, category: c }),
    diagnosePerformance: () => ({ success: true }),
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: () => null, setProperty: () => {}, getProperties: () => ({}), deleteProperty: () => {}
      })
    },
    getCachedProperty: () => null,
    getColumnAnalysis: (spreadsheetId, sheetName) => ({
      success: true,
      headers: ['ts', 'name', 'ans', 'reason', 'rate'],
      mapping: { answer: 2, reason: 3, name: 1, numericX: 4 },
      confidence: { answer: 90, reason: 85, numericX: 92 },
      numericScaleCandidates: [{ index: 4, header: 'rate', min: 1, max: 5, confidence: 92 }]
    }),
    getPublishedSheetData: (cls, sort, admin, targetId) => ({
      success: true, header: 'Q', sheetName: 'S',
      displaySettings: { boardMode: 'numberline', showNames: false },
      axisConfig: { defaultMin: 1, defaultMax: 5 },
      data: [{ id: 'row_2', numericX: 3 }, { id: 'row_3', numericX: 4 }, { id: 'row_4', numericX: 5 }, { id: 'row_5', numericX: 2 }]
    }),
    // expose saved configs for assertions
    __savedConfigs: savedConfigs,
    ...overrides
  };
  vm.createContext(context);
  const source = fs.readFileSync(path.resolve(__dirname, '../src/AdminApis.js'), 'utf8');
  vm.runInContext(source, context, { filename: 'AdminApis.js' });
  return context;
}

// =====================================================================
// findUser
// =====================================================================

test('findUser: returns parsed config for matching email', () => {
  const ctx = loadAdminContext();
  const res = ctx.dispatchAdminOperation('findUser', { email: 'a@example.com' });
  assert.equal(res.success, true);
  assert.equal(res.data.user.userId, 'u1');
  assert.equal(res.data.user.userEmail, 'a@example.com');
  assert.equal(res.data.user.config.isPublished, true);
});

test('findUser: rejects missing email', () => {
  const ctx = loadAdminContext();
  const res = ctx.dispatchAdminOperation('findUser', {});
  assert.equal(res.success, false);
  assert.match(res.message, /email/);
});

test('findUser: user not found', () => {
  const ctx = loadAdminContext();
  const res = ctx.dispatchAdminOperation('findUser', { email: 'nope@x.com' });
  assert.equal(res.success, false);
  assert.match(res.message, /not found/i);
});

// =====================================================================
// getUserConfig
// =====================================================================

test('getUserConfig: returns parsed config', () => {
  const ctx = loadAdminContext();
  const res = ctx.dispatchAdminOperation('getUserConfig', { userId: 'u1' });
  assert.equal(res.success, true);
  assert.equal(res.data.config.spreadsheetId, 'sheet1');
  assert.equal(res.data.config.displaySettings.boardMode, 'auto');
});

test('getUserConfig: missing userId rejects', () => {
  const ctx = loadAdminContext();
  const res = ctx.dispatchAdminOperation('getUserConfig', {});
  assert.equal(res.success, false);
});

// =====================================================================
// exportConfigs
// =====================================================================

test('exportConfigs: returns all users with parsed config', () => {
  const ctx = loadAdminContext({
    // getAdminUsers が getAllUsers と異なる API シェイプを返すケースに対応するための stub
  });
  // getAdminUsers 内部関数も同 file 内なので、まずは現実の returns に近い形を直接 stub
  ctx.getAdminUsers = () => ({
    success: true,
    data: { users: ctx.getAllUsers() }
  });

  const res = ctx.dispatchAdminOperation('exportConfigs', {});
  assert.equal(res.success, true);
  assert.equal(res.data.count, 3);
  assert.equal(res.data.configs[0].userEmail, 'a@example.com');
  assert.equal(res.data.configs[0].config.isPublished, true);
});

// =====================================================================
// setUserConfig — deep merge + protected fields
// =====================================================================

test('setUserConfig: shallow-merges top-level fields', () => {
  const ctx = loadAdminContext();
  const res = ctx.dispatchAdminOperation('setUserConfig', {
    userId: 'u1',
    patch: { allowResubmit: true }
  });
  assert.equal(res.success, true);
  const saved = ctx.__savedConfigs.get('u1').config;
  assert.equal(saved.allowResubmit, true);
  // 既存フィールドは保持
  assert.equal(saved.spreadsheetId, 'sheet1');
  assert.equal(saved.displaySettings.boardMode, 'auto');
});

test('setUserConfig: deep-merges nested objects (preserves siblings)', () => {
  const ctx = loadAdminContext();
  ctx.dispatchAdminOperation('setUserConfig', {
    userId: 'u1',
    patch: { displaySettings: { boardMode: 'numberline' } }
  });
  const saved = ctx.__savedConfigs.get('u1').config;
  assert.equal(saved.displaySettings.boardMode, 'numberline');
  // showNames / pageSize はそのまま
  assert.equal(saved.displaySettings.showNames, false);
  assert.equal(saved.displaySettings.pageSize, 20);
});

test('setUserConfig: null in patch deletes the key', () => {
  const ctx = loadAdminContext();
  ctx.dispatchAdminOperation('setUserConfig', {
    userId: 'u1',
    patch: { columnMapping: { reason: null } }
  });
  const saved = ctx.__savedConfigs.get('u1').config;
  assert.equal(saved.columnMapping.reason, undefined);
  assert.equal(saved.columnMapping.answer, 3);
});

test('setUserConfig: strips protected fields from patch', () => {
  const ctx = loadAdminContext();
  ctx.dispatchAdminOperation('setUserConfig', {
    userId: 'u1',
    patch: { userId: 'INJECTED', userEmail: 'attacker@x.com', allowResubmit: true }
  });
  const saved = ctx.__savedConfigs.get('u1').config;
  assert.equal(saved.userId, 'u1');                    // 保護された値が維持
  assert.equal(saved.userEmail, undefined);            // 元 config に email は無いので未追加でOK
  assert.equal(saved.allowResubmit, true);             // 通常フィールドは適用
});

test('setUserConfig: rejects missing userId', () => {
  const ctx = loadAdminContext();
  const res = ctx.dispatchAdminOperation('setUserConfig', { patch: { allowResubmit: true } });
  assert.equal(res.success, false);
});

test('setUserConfig: rejects non-object patch', () => {
  const ctx = loadAdminContext();
  const r1 = ctx.dispatchAdminOperation('setUserConfig', { userId: 'u1', patch: 'string' });
  assert.equal(r1.success, false);
  const r2 = ctx.dispatchAdminOperation('setUserConfig', { userId: 'u1', patch: [] });
  assert.equal(r2.success, false);
});

test('setUserConfig: publish option triggers isPublish save mode', () => {
  const ctx = loadAdminContext();
  ctx.dispatchAdminOperation('setUserConfig', {
    userId: 'u1', patch: { isPublished: true }, publish: true
  });
  const opts = ctx.__savedConfigs.get('u1').options;
  assert.equal(opts.isPublish, true);
});

// =====================================================================
// bulkSetUserConfig
// =====================================================================

test('bulkSetUserConfig: dryRun returns matched count without writes', () => {
  const ctx = loadAdminContext();
  ctx.getAdminUsers = () => ({ success: true, data: { users: ctx.getAllUsers() } });
  const res = ctx.dispatchAdminOperation('bulkSetUserConfig', {
    patch: { allowResubmit: true },
    dryRun: true
  });
  assert.equal(res.success, true);
  assert.equal(res.data.matchedCount, 3);
  // 書き込みは発生しない
  assert.equal(ctx.__savedConfigs.size, 0);
});

test('bulkSetUserConfig: filter isPublished=false matches subset', () => {
  const ctx = loadAdminContext();
  ctx.getAdminUsers = () => ({ success: true, data: { users: ctx.getAllUsers() } });
  const res = ctx.dispatchAdminOperation('bulkSetUserConfig', {
    patch: { allowResubmit: true },
    filter: { isPublished: false },
    dryRun: true
  });
  // u2 (isPublished:false) と u3 (configJson:null → isPublished:undefined → false 扱い) がマッチ
  assert.equal(res.data.matchedCount, 2);
});

test('bulkSetUserConfig: applies patch to matched users', () => {
  const ctx = loadAdminContext();
  ctx.getAdminUsers = () => ({ success: true, data: { users: ctx.getAllUsers() } });
  const res = ctx.dispatchAdminOperation('bulkSetUserConfig', {
    patch: { allowResubmit: true },
    filter: { emailContains: 'a@' }
  });
  assert.equal(res.success, true);
  assert.equal(res.data.appliedTo, 1);
  assert.equal(ctx.__savedConfigs.size, 1);
});

test('bulkSetUserConfig: rejects non-object patch', () => {
  const ctx = loadAdminContext();
  const res = ctx.dispatchAdminOperation('bulkSetUserConfig', { patch: null });
  assert.equal(res.success, false);
});

// =====================================================================
// runColumnAnalysis
// =====================================================================

test('runColumnAnalysis: direct call with spreadsheetId', () => {
  const ctx = loadAdminContext();
  const res = ctx.dispatchAdminOperation('runColumnAnalysis', {
    spreadsheetId: 'sheet1',
    sheetName: 'フォームの回答 1'
  });
  assert.equal(res.success, true);
  assert.ok(res.mapping);
  assert.equal(res.mapping.numericX, 4);
  assert.equal(res.numericScaleCandidates.length, 1);
});

test('runColumnAnalysis: resolves spreadsheetId from userId', () => {
  const ctx = loadAdminContext();
  const res = ctx.dispatchAdminOperation('runColumnAnalysis', { userId: 'u1' });
  assert.equal(res.success, true);
  assert.ok(res.mapping);
});

test('runColumnAnalysis: rejects when neither userId nor spreadsheetId provided', () => {
  const ctx = loadAdminContext();
  const res = ctx.dispatchAdminOperation('runColumnAnalysis', {});
  assert.equal(res.success, false);
});

// =====================================================================
// previewBoard
// =====================================================================

test('previewBoard: returns trimmed result (sample only, no full data)', () => {
  const ctx = loadAdminContext();
  const res = ctx.dispatchAdminOperation('previewBoard', { userId: 'u1' });
  assert.equal(res.success, true);
  assert.equal(res.data.dataCount, 4);
  // sample は 3 件まで
  assert.equal(res.data.sampleData.length, 3);
  // axisConfig / displaySettings は維持
  assert.equal(res.data.displaySettings.boardMode, 'numberline');
});

test('previewBoard: requires userId', () => {
  const ctx = loadAdminContext();
  const res = ctx.dispatchAdminOperation('previewBoard', {});
  assert.equal(res.success, false);
});
