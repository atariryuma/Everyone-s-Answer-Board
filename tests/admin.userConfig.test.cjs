const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('path');
const vm = require('vm');
const { gasResponseStubs } = require('./_helpers.cjs');

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
    ...gasResponseStubs(),
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
    // Form-related stubs
    getForms: () => ({ success: true, forms: [{ id: 'f1', name: 'Form1', url: 'https://docs.google.com/forms/d/f1/viewform' }] }),
    createTemplateForm: (type) => ({
      success: true,
      formUrl: `https://docs.google.com/forms/d/new_${type || 'board'}/viewform`,
      formTitle: `Template ${type || 'board'}`,
      formId: `new_${type || 'board'}`,
      spreadsheetId: `sheet_${type || 'board'}`,
      sheetName: 'フォームの回答 1',
      wasCreated: true
    }),
    customizeForm: (id, schema) => {
      if (!id) return { success: false, error: 'formId required' };
      if (!schema || !Array.isArray(schema.questions)) return { success: false, error: 'schema.questions required' };
      return {
        success: true,
        formId: id,
        formUrl: 'https://docs.google.com/forms/d/' + id + '/viewform',
        editUrl: 'https://docs.google.com/forms/d/' + id + '/edit',
        title: schema.title || 'customized',
        itemCount: schema.questions.length
      };
    },
    processFormUrlInput: (url) => {
      if (!url || !url.includes('/forms/')) {
        return { success: false, error: 'Invalid form URL' };
      }
      return {
        success: true,
        formUrl: url,
        formTitle: 'Existing Form',
        formId: 'exf1',
        spreadsheetId: 'exsheet1',
        sheetName: 'フォームの回答 1',
        wasCreated: false
      };
    },
    isValidFormUrl: (url) => typeof url === 'string' && (url.includes('/forms/d/') || url.includes('forms.gle/')),
    FormApp: {
      openByUrl: (url) => {
        if (!url.includes('/forms/')) throw new Error('Unreachable form');
        return {
          getTitle: () => 'Existing Form',
          getId: () => 'exf1',
          getDestinationId: () => 'exsheet1'
        };
      }
    },
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
  const ctx = loadAdminContext();
  ctx.getAdminUsers = () => ({
    success: true,
    users: ctx.getAllUsers()
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

test('setUserConfig: preserves boardMode through sanitization pipeline', () => {
  // Why: 過去 validators.js:validateConfig が displaySettings を再構築する際に
  //      boardMode を許可リストに含めず、永遠に削除されていた regression。
  //      これを防ぐため、setUserConfig 経由で boardMode が wire まで生きることを検証。
  const ctx = loadAdminContext();
  ctx.dispatchAdminOperation('setUserConfig', {
    userId: 'u1', patch: { displaySettings: { boardMode: 'numberline' } }
  });
  const saved = ctx.__savedConfigs.get('u1').config;
  assert.equal(saved.displaySettings.boardMode, 'numberline');
});

// (boardMode invalid-value rejection is tested at the validators.js layer in
//  tests/validators.test.cjs, since saveUserConfig is stubbed at this layer.)

// =====================================================================
// bulkSetUserConfig
// =====================================================================

test('bulkSetUserConfig: dryRun returns matched count without writes', () => {
  const ctx = loadAdminContext();
  ctx.getAdminUsers = () => ({ success: true, users: ctx.getAllUsers() });
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
  ctx.getAdminUsers = () => ({ success: true, users: ctx.getAllUsers() });
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
  ctx.getAdminUsers = () => ({ success: true, users: ctx.getAllUsers() });
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

test('previewBoard: --limit raises the sample cap (default stays 3)', () => {
  const ctx = loadAdminContext();
  // limit 未指定は従来通り 3 件
  assert.equal(ctx.dispatchAdminOperation('previewBoard', { userId: 'u1' }).data.sampleData.length, 3);
  // limit 指定で dataCount(=4) まで取れる
  assert.equal(ctx.dispatchAdminOperation('previewBoard', { userId: 'u1', limit: 4 }).data.sampleData.length, 4);
  // limit 1 は 1 件
  assert.equal(ctx.dispatchAdminOperation('previewBoard', { userId: 'u1', limit: 1 }).data.sampleData.length, 1);
});

test('previewBoard: requires userId', () => {
  const ctx = loadAdminContext();
  const res = ctx.dispatchAdminOperation('previewBoard', {});
  assert.equal(res.success, false);
});

// =====================================================================
// Form operations (v3)
// =====================================================================

test('listMyForms: returns list of forms in caller Drive', () => {
  const ctx = loadAdminContext();
  const res = ctx.dispatchAdminOperation('listMyForms', {});
  assert.equal(res.success, true);
  assert.equal(res.forms.length, 1);
  assert.equal(res.forms[0].id, 'f1');
});

test('validateFormUrl: rejects missing URL', () => {
  const ctx = loadAdminContext();
  const res = ctx.dispatchAdminOperation('validateFormUrl', {});
  assert.equal(res.success, false);
});

test('validateFormUrl: detects non-form URL', () => {
  const ctx = loadAdminContext();
  const res = ctx.dispatchAdminOperation('validateFormUrl', {
    formUrl: 'https://example.com/not-a-form'
  });
  assert.equal(res.success, true);
  assert.equal(res.data.isValidFormUrl, false);
});

test('validateFormUrl: confirms a reachable Form URL', () => {
  const ctx = loadAdminContext();
  const res = ctx.dispatchAdminOperation('validateFormUrl', {
    formUrl: 'https://docs.google.com/forms/d/f1/viewform'
  });
  assert.equal(res.success, true);
  assert.equal(res.data.isValidFormUrl, true);
  assert.equal(res.data.reachable, true);
  assert.equal(res.data.formTitle, 'Existing Form');
});

test('connectForm: rejects missing formUrl', () => {
  const ctx = loadAdminContext();
  const res = ctx.dispatchAdminOperation('connectForm', { userId: 'u1' });
  assert.equal(res.success, false);
});

test('connectForm: links existing form to user config', () => {
  const ctx = loadAdminContext();
  const res = ctx.dispatchAdminOperation('connectForm', {
    userId: 'u2',
    formUrl: 'https://docs.google.com/forms/d/f1/viewform'
  });
  assert.equal(res.success, true);
  assert.equal(res.data.userId, 'u2');
  assert.equal(res.data.applied.formTitle, 'Existing Form');
  assert.equal(res.data.applied.spreadsheetId, 'exsheet1');
  // 実際に saveUserConfig が呼ばれて config に反映
  const saved = ctx.__savedConfigs.get('u2').config;
  assert.equal(saved.formUrl, 'https://docs.google.com/forms/d/f1/viewform');
  assert.equal(saved.spreadsheetId, 'exsheet1');
});

test('connectForm: surfaces error when processFormUrlInput fails', () => {
  const ctx = loadAdminContext();
  const res = ctx.dispatchAdminOperation('connectForm', {
    userId: 'u1',
    formUrl: 'not-a-form-url'
  });
  assert.equal(res.success, false);
});

test('createForm: creates board template and links to user', () => {
  const ctx = loadAdminContext();
  const res = ctx.dispatchAdminOperation('createForm', { userId: 'u1' });
  assert.equal(res.success, true);
  assert.equal(res.data.templateType, 'board');
  // 紐付け確認
  const saved = ctx.__savedConfigs.get('u1').config;
  assert.equal(saved.formTitle, 'Template board');
  assert.equal(saved.spreadsheetId, 'sheet_board');
});

test('createForm: numberline template also sets boardMode', () => {
  const ctx = loadAdminContext();
  const res = ctx.dispatchAdminOperation('createForm', {
    userId: 'u1', templateType: 'numberline'
  });
  assert.equal(res.success, true);
  assert.equal(res.data.templateType, 'numberline');
  const saved = ctx.__savedConfigs.get('u1').config;
  assert.equal(saved.displaySettings.boardMode, 'numberline');
});

test('createForm: matrix template also sets boardMode', () => {
  const ctx = loadAdminContext();
  ctx.dispatchAdminOperation('createForm', {
    userId: 'u1', templateType: 'matrix'
  });
  const saved = ctx.__savedConfigs.get('u1').config;
  assert.equal(saved.displaySettings.boardMode, 'matrix');
});

test('createForm: invalid templateType falls back to board', () => {
  const ctx = loadAdminContext();
  const res = ctx.dispatchAdminOperation('createForm', {
    userId: 'u1', templateType: 'INVALID'
  });
  assert.equal(res.success, true);
  assert.equal(res.data.templateType, 'board');
});

test('createForm / connectForm: rejects unknown userId', () => {
  const ctx = loadAdminContext();
  const r1 = ctx.dispatchAdminOperation('createForm', { userId: 'NOPE' });
  assert.equal(r1.success, false);
  const r2 = ctx.dispatchAdminOperation('connectForm', {
    userId: 'NOPE', formUrl: 'https://docs.google.com/forms/d/f1/viewform'
  });
  assert.equal(r2.success, false);
});

// =====================================================================
// Multi-board profiles (v4)
// =====================================================================

test('listProfiles: returns empty list when none saved', () => {
  const ctx = loadAdminContext();
  const res = ctx.dispatchAdminOperation('listProfiles', { userId: 'u1' });
  assert.equal(res.success, true);
  assert.equal(res.data.count, 0);
  assert.equal(res.data.activeProfile, null);
});

test('listProfiles: rejects missing userId', () => {
  const ctx = loadAdminContext();
  const res = ctx.dispatchAdminOperation('listProfiles', {});
  assert.equal(res.success, false);
});

test('saveProfile: rejects missing name', () => {
  const ctx = loadAdminContext();
  const res = ctx.dispatchAdminOperation('saveProfile', { userId: 'u1' });
  assert.equal(res.success, false);
});

test('saveProfile: snapshots current active config', () => {
  const ctx = loadAdminContext();
  // u1 の active は formUrl 等未設定だが columnMapping は { answer:3, reason:4 }
  const res = ctx.dispatchAdminOperation('saveProfile', {
    userId: 'u1', name: '導入アンケート'
  });
  assert.equal(res.success, true);
  const saved = ctx.__savedConfigs.get('u1').config;
  assert.ok(Array.isArray(saved.profiles));
  assert.equal(saved.profiles.length, 1);
  assert.equal(saved.profiles[0].name, '導入アンケート');
  assert.equal(saved.profiles[0].columnMapping.answer, 3);
});

test('saveProfile: same name overwrites existing profile', () => {
  const ctx = loadAdminContext();
  ctx.dispatchAdminOperation('saveProfile', { userId: 'u1', name: 'p1' });
  ctx.dispatchAdminOperation('saveProfile', { userId: 'u1', name: 'p1' });
  const saved = ctx.__savedConfigs.get('u1').config;
  // 同名 2 回保存しても profile は 1 件
  assert.equal(saved.profiles.length, 1);
});

test('saveProfile: snapshot override applies different data', () => {
  const ctx = loadAdminContext();
  const res = ctx.dispatchAdminOperation('saveProfile', {
    userId: 'u1',
    name: '振り返り',
    snapshot: {
      formUrl: 'https://...../forms/diff/viewform',
      spreadsheetId: 'diff-sheet',
      sheetName: '別シート',
      columnMapping: { answer: 1 },
      displaySettings: { boardMode: 'wordcloud' }
    }
  });
  assert.equal(res.success, true);
  const saved = ctx.__savedConfigs.get('u1').config;
  const p = saved.profiles[0];
  assert.equal(p.formUrl, 'https://...../forms/diff/viewform');
  assert.equal(p.spreadsheetId, 'diff-sheet');
  assert.equal(p.displaySettings.boardMode, 'wordcloud');
});

test('loadProfile: applies snapshot to active config + sets activeProfile', () => {
  const ctx = loadAdminContext();
  // 先に保存
  ctx.dispatchAdminOperation('saveProfile', {
    userId: 'u1',
    name: '本時の議論',
    snapshot: {
      formUrl: 'https://forms.example.com/main',
      spreadsheetId: 'main-sheet',
      sheetName: 'メイン',
      columnMapping: { answer: 5, reason: 6 },
      displaySettings: { boardMode: 'numberline' },
      xAxisLabels: { min: 'そのまま', max: '直す' }
    }
  });
  // active に適用
  const res = ctx.dispatchAdminOperation('loadProfile', {
    userId: 'u1', name: '本時の議論'
  });
  assert.equal(res.success, true);
  const saved = ctx.__savedConfigs.get('u1').config;
  assert.equal(saved.activeProfile, '本時の議論');
  assert.equal(saved.formUrl, 'https://forms.example.com/main');
  assert.equal(saved.spreadsheetId, 'main-sheet');
  assert.equal(saved.sheetName, 'メイン');
  assert.equal(saved.displaySettings.boardMode, 'numberline');
  assert.equal(saved.xAxisLabels.min, 'そのまま');
});

test('loadProfile: rejects missing name or unknown profile', () => {
  const ctx = loadAdminContext();
  const r1 = ctx.dispatchAdminOperation('loadProfile', { userId: 'u1' });
  assert.equal(r1.success, false);
  const r2 = ctx.dispatchAdminOperation('loadProfile', { userId: 'u1', name: 'NOPE' });
  assert.equal(r2.success, false);
});

test('deleteProfile: removes profile from list', () => {
  const ctx = loadAdminContext();
  ctx.dispatchAdminOperation('saveProfile', { userId: 'u1', name: 'p1' });
  ctx.dispatchAdminOperation('saveProfile', { userId: 'u1', name: 'p2' });
  const r = ctx.dispatchAdminOperation('deleteProfile', { userId: 'u1', name: 'p1' });
  assert.equal(r.success, true);
  const saved = ctx.__savedConfigs.get('u1').config;
  assert.equal(saved.profiles.length, 1);
  assert.equal(saved.profiles[0].name, 'p2');
});

test('deleteProfile: clears activeProfile when deleting active one', () => {
  const ctx = loadAdminContext();
  ctx.dispatchAdminOperation('saveProfile', { userId: 'u1', name: 'p1' });
  ctx.dispatchAdminOperation('loadProfile', { userId: 'u1', name: 'p1' });
  let saved = ctx.__savedConfigs.get('u1').config;
  assert.equal(saved.activeProfile, 'p1');
  ctx.dispatchAdminOperation('deleteProfile', { userId: 'u1', name: 'p1' });
  saved = ctx.__savedConfigs.get('u1').config;
  // activeProfile は削除（null か undefined）
  assert.ok(!saved.activeProfile);
});

test('deleteProfile: rejects unknown profile name', () => {
  const ctx = loadAdminContext();
  const res = ctx.dispatchAdminOperation('deleteProfile', { userId: 'u1', name: 'NOPE' });
  assert.equal(res.success, false);
});

test('deleteProfile: re-anchors active to remaining profile (orphan prevention)', () => {
  const ctx = loadAdminContext();
  ctx.dispatchAdminOperation('saveProfile', {
    userId: 'u1',
    name: 'p1',
    snapshot: {
      formUrl: 'https://forms.example.com/p1',
      formTitle: 'P1 タイトル',
      spreadsheetId: 'sheet-p1',
      sheetName: 'P1 シート',
      columnMapping: { answer: 4, reason: 5 },
      displaySettings: { boardMode: 'pie' }
    }
  });
  ctx.dispatchAdminOperation('saveProfile', {
    userId: 'u1',
    name: 'p2',
    snapshot: {
      formUrl: 'https://forms.example.com/p2',
      formTitle: 'P2 タイトル',
      spreadsheetId: 'sheet-p2',
      sheetName: 'P2 シート',
      columnMapping: { answer: 3 },
      displaySettings: { boardMode: 'wordcloud' }
    }
  });
  // active = p1
  ctx.dispatchAdminOperation('loadProfile', { userId: 'u1', name: 'p1' });
  let saved = ctx.__savedConfigs.get('u1').config;
  assert.equal(saved.activeProfile, 'p1');
  assert.equal(saved.spreadsheetId, 'sheet-p1');

  // active 削除: p2 に再 anchor されることを期待（旧 sheet-p1 が top-level に残らない）
  const r = ctx.dispatchAdminOperation('deleteProfile', { userId: 'u1', name: 'p1' });
  assert.equal(r.success, true);
  assert.equal(r.data.reAnchored, true);
  assert.equal(r.data.remainingActive, 'p2');

  saved = ctx.__savedConfigs.get('u1').config;
  assert.equal(saved.activeProfile, 'p2');
  assert.equal(saved.spreadsheetId, 'sheet-p2');
  assert.equal(saved.formUrl, 'https://forms.example.com/p2');
  assert.equal(saved.sheetName, 'P2 シート');
  assert.equal(saved.displaySettings.boardMode, 'wordcloud');
});

test('deleteProfile: clears top-level form fields when deleting last profile (full orphan clear)', () => {
  const ctx = loadAdminContext();
  ctx.dispatchAdminOperation('saveProfile', {
    userId: 'u1',
    name: 'only',
    snapshot: {
      formUrl: 'https://forms.example.com/only',
      spreadsheetId: 'sheet-only',
      sheetName: 'シート',
      columnMapping: { answer: 4 }
    }
  });
  ctx.dispatchAdminOperation('loadProfile', { userId: 'u1', name: 'only' });
  let saved = ctx.__savedConfigs.get('u1').config;
  assert.equal(saved.spreadsheetId, 'sheet-only');

  ctx.dispatchAdminOperation('deleteProfile', { userId: 'u1', name: 'only' });
  saved = ctx.__savedConfigs.get('u1').config;
  assert.ok(!saved.activeProfile);
  assert.equal(saved.spreadsheetId, '');
  assert.equal(saved.formUrl, '');
  assert.equal(saved.sheetName, '');
});

test('loadProfile: clears prior columnMapping when new profile has none (null fallback)', () => {
  const ctx = loadAdminContext();
  // matrix profile: columnMapping に numericX/Y を含む
  ctx.dispatchAdminOperation('saveProfile', {
    userId: 'u1',
    name: 'matrix',
    snapshot: {
      formUrl: 'https://forms.example.com/matrix',
      spreadsheetId: 'sheet-matrix',
      columnMapping: { answer: 4, reason: 6, numericX: 4, numericY: 5 },
      displaySettings: { boardMode: 'matrix' }
    }
  });
  // wordcloud profile: columnMapping を持たない（空）
  ctx.dispatchAdminOperation('saveProfile', {
    userId: 'u1',
    name: 'wc',
    snapshot: {
      formUrl: 'https://forms.example.com/wc',
      spreadsheetId: 'sheet-wc',
      displaySettings: { boardMode: 'wordcloud' }
    }
  });

  ctx.dispatchAdminOperation('loadProfile', { userId: 'u1', name: 'matrix' });
  let saved = ctx.__savedConfigs.get('u1').config;
  assert.equal(saved.columnMapping.numericX, 4);

  // wc に切替: numericX/Y が消えるべき (前 matrix の値を引きずらない)
  ctx.dispatchAdminOperation('loadProfile', { userId: 'u1', name: 'wc' });
  saved = ctx.__savedConfigs.get('u1').config;
  assert.equal(saved.activeProfile, 'wc');
  // columnMapping は空 (numericX が居残らない)
  assert.ok(!saved.columnMapping || saved.columnMapping.numericX === undefined);
});

test('listProfiles: after save shows profile summary', () => {
  const ctx = loadAdminContext();
  ctx.dispatchAdminOperation('saveProfile', {
    userId: 'u1',
    name: '導入',
    snapshot: { formTitle: 'タイトル A', displaySettings: { boardMode: 'pie' } }
  });
  ctx.dispatchAdminOperation('saveProfile', {
    userId: 'u1',
    name: '展開',
    snapshot: { formTitle: 'タイトル B', displaySettings: { boardMode: 'numberline' } }
  });
  const res = ctx.dispatchAdminOperation('listProfiles', { userId: 'u1' });
  assert.equal(res.success, true);
  assert.equal(res.data.count, 2);
  const byName = Object.fromEntries(res.data.profiles.map(p => [p.name, p]));
  assert.equal(byName['導入'].boardMode, 'pie');
  assert.equal(byName['展開'].boardMode, 'numberline');
});

// =====================================================================
// customizeForm
// =====================================================================

test('customizeForm: rejects missing formId/formUrl', () => {
  const ctx = loadAdminContext();
  const res = ctx.dispatchAdminOperation('customizeForm', { schema: { questions: [] } });
  assert.equal(res.success, false);
});

test('customizeForm: rejects missing schema', () => {
  const ctx = loadAdminContext();
  const res = ctx.dispatchAdminOperation('customizeForm', { formId: 'abc' });
  assert.equal(res.success, false);
});

test('customizeForm: success with valid schema', () => {
  const ctx = loadAdminContext();
  const res = ctx.dispatchAdminOperation('customizeForm', {
    formId: 'newForm123',
    schema: {
      title: '道徳: 議論',
      questions: [
        { type: 'list', title: 'クラス', choices: ['6年1組', '6年2組', '6年3組', '6年4組'] },
        { type: 'text', title: '名前' },
        { type: 'scale', title: 'あなたならどうする？', min: 1, max: 5, leftLabel: 'そのまま', rightLabel: '直す' },
        { type: 'paragraph', title: '理由' }
      ]
    }
  });
  assert.equal(res.success, true);
  assert.equal(res.itemCount, 4);
});

test('customizeForm: accepts formUrl alternative', () => {
  const ctx = loadAdminContext();
  const res = ctx.dispatchAdminOperation('customizeForm', {
    formUrl: 'https://docs.google.com/forms/d/xyz/edit',
    schema: { questions: [{ type: 'text', title: 'Q1' }] }
  });
  assert.equal(res.success, true);
});
