/**
 * createTemplateForm / createUserFolder / moveFileToFolder_ のテスト
 *
 * Why: 可視化モード（M1/M2）連動テンプレートと、Drive 単一親モデル準拠の
 *      ファイル移動、オーナー絞り込み付きフォルダ検索を保証する。
 *
 * FormApp / SpreadsheetApp / DriveApp / UrlFetchApp は GAS 専用 API なので
 * フェイクを差し込んで作成された Items の構造を assert する。
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

// =====================================================================
// Fake builder utilities
// =====================================================================

function makeFakeForm() {
  const items = [];
  let dest = null;
  let collectEmail = false;
  let limitOne = false;

  function makeItem(kind) {
    const item = { kind, title: '', required: false, helpText: '', choices: null, bounds: null, labels: null };
    return {
      setTitle: (t) => { item.title = t; return makeChain(item); },
      _item: item
    };
  }
  function makeChain(item) {
    return {
      setTitle: (t) => { item.title = t; return makeChain(item); },
      setRequired: (b) => { item.required = b; return makeChain(item); },
      setHelpText: (t) => { item.helpText = t; return makeChain(item); },
      setChoiceValues: (v) => { item.choices = v.slice(); return makeChain(item); },
      setBounds: (lo, hi) => { item.bounds = [lo, hi]; return makeChain(item); },
      setLabels: (lo, hi) => { item.labels = [lo, hi]; return makeChain(item); }
    };
  }

  const form = {
    _id: 'form_' + Math.random().toString(36).slice(2, 10),
    _items: items,
    getId: function () { return this._id; },
    getPublishedUrl: function () { return 'https://docs.google.com/forms/d/' + this._id + '/viewform'; },
    getEditUrl: function () { return 'https://docs.google.com/forms/d/' + this._id + '/edit'; },
    setLimitOneResponsePerUser: function (b) { limitOne = b; return this; },
    setCollectEmail: function (b) { collectEmail = b; return this; },
    setDestination: function (_type, ssId) { dest = ssId; return this; },
    addListItem: function () { const it = makeItem('list'); items.push(it._item); return makeChain(it._item); },
    addTextItem: function () { const it = makeItem('text'); items.push(it._item); return makeChain(it._item); },
    addParagraphTextItem: function () { const it = makeItem('paragraph'); items.push(it._item); return makeChain(it._item); },
    addMultipleChoiceItem: function () { const it = makeItem('mc'); items.push(it._item); return makeChain(it._item); },
    addScaleItem: function () { const it = makeItem('scale'); items.push(it._item); return makeChain(it._item); },
    // テスト側から参照する
    _state: () => ({ limitOne, collectEmail, dest, items })
  };
  return form;
}

function makeFakeSpreadsheet(title) {
  return {
    _id: 'ss_' + Math.random().toString(36).slice(2, 10),
    _name: title,
    getId: function () { return this._id; },
    getUrl: function () { return 'https://docs.google.com/spreadsheets/d/' + this._id + '/edit'; },
    getName: function () { return this._name; }
  };
}

function makeFakeFolder(name, ownerEmail) {
  return {
    _name: name,
    _owner: ownerEmail,
    _moved: [],
    _addedFiles: [],
    _removed: [],
    getName: function () { return this._name; },
    getUrl: function () { return 'https://drive.google.com/drive/folders/folder_' + name; },
    getOwner: function () { return { getEmail: () => this._owner }; },
    addFile: function (f) { this._addedFiles.push(f); return this; },
    removeFile: function (f) { this._removed.push(f); return this; }
  };
}

function makeFakeDriveFile(id, supportMoveTo) {
  const file = {
    _id: id,
    _movedTo: null,
    getId: () => id,
    getName: () => 'file_' + id
  };
  if (supportMoveTo) {
    file.moveTo = function (folder) { this._movedTo = folder; return this; };
  }
  return file;
}

// =====================================================================
// Context loader
// =====================================================================

function loadCtx(overrides = {}) {
  const defaultForm = makeFakeForm();
  const defaultSs = makeFakeSpreadsheet('Sheet');
  const myEmail = overrides.__myEmail || 'teacher@example.com';

  // Drive: getFoldersByName と createFolder, getFileById をフェイクする
  const driveState = overrides.__drive || {
    foldersByName: {},      // name -> [folder, folder, ...]
    rootFolder: { _removed: [], removeFile: function (f) { this._removed.push(f); } },
    files: {}                // id -> file object
  };

  const ctx = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    URL,
    createErrorResponse: (m) => ({ success: false, message: m }),
    createExceptionResponse: (e) => ({ success: false, message: e.message }),
    createSuccessResponse: (m, data) => ({ success: true, message: m, ...(data && { data }) }),
    getCurrentEmail: () => myEmail,
    findUserByEmail: () => null,
    findUserById: () => null,
    findPublishedBoardOwner: () => null,
    findUserBySpreadsheetId: () => null,
    getUserConfig: () => ({ success: true, config: {} }),
    getConfigOrDefault: () => ({}),
    saveUserConfig: () => ({ success: true }),
    openSpreadsheet: () => null,
    SpreadsheetApp: {
      create: (title) => {
        const ss = makeFakeSpreadsheet(title);
        driveState.files[ss.getId()] = makeFakeDriveFile(ss.getId(), true);
        return ss;
      },
      openById: () => { throw new Error('not stubbed'); }
    },
    FormApp: {
      create: (title) => {
        const f = makeFakeForm();
        f._title = title;
        f.getTitle = () => title;
        // ctx.__lastForm に保持してテストから検査
        ctx.__lastForm = f;
        driveState.files[f.getId()] = makeFakeDriveFile(f.getId(), true);
        return f;
      },
      DestinationType: { SPREADSHEET: 'SPREADSHEET' },
      openByUrl: () => { throw new Error('not stubbed'); }
    },
    DriveApp: {
      getFoldersByName: (name) => {
        const arr = driveState.foldersByName[name] || [];
        let idx = 0;
        return {
          hasNext: () => idx < arr.length,
          next: () => arr[idx++]
        };
      },
      createFolder: (name) => {
        const f = makeFakeFolder(name, myEmail);
        if (!driveState.foldersByName[name]) driveState.foldersByName[name] = [];
        driveState.foldersByName[name].push(f);
        return f;
      },
      getFileById: (id) => {
        if (!driveState.files[id]) driveState.files[id] = makeFakeDriveFile(id, true);
        return driveState.files[id];
      },
      getRootFolder: () => driveState.rootFolder
    },
    UrlFetchApp: {
      fetch: () => ({ getResponseCode: () => 200, getContentText: () => '{}' })
    },
    ScriptApp: {
      getOAuthToken: () => 'fake-token',
      getService: () => ({ getUrl: () => 'https://script.google.com/x' })
    },
    Session: {
      getActiveUser: () => ({ getEmail: () => myEmail }),
      getScriptTimeZone: () => 'Asia/Tokyo'
    },
    Utilities: {
      formatDate: () => '1/1 12:00'
    },
    CacheService: { getScriptCache: () => ({ get: () => null, put: () => {}, remove: () => {} }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null, setProperty: () => {} }) },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
    CACHE_DURATION: { SHORT: 10, MEDIUM: 30, LONG: 300, DATABASE_LONG: 600, FORM_DATA: 30 },
    TIMEOUT_MS: { DEFAULT: 5000 },
    SYSTEM_LIMITS: {},
    DEFAULT_DISPLAY_SETTINGS: { showNames: false, showReactions: true, theme: 'default', pageSize: 20 },
    getBatchedAdminAuth: () => ({ success: true, authenticated: true, email: myEmail, isAdmin: false }),
    getCachedProperty: () => null,
    saveToCacheWithSizeCheck: () => true,
    emailToShortHash: () => 'h_xxxx',
    validateEmail: (e) => ({ isValid: typeof e === 'string' && /.+@.+/.test(e) }),
    validateUrl: () => ({ isValid: true }),
    validateSpreadsheetId: () => true,
    normalizeHeader: (h) => String(h || '').toLowerCase().trim(),
    extractReactions: () => ({}),
    extractHighlight: () => false,
    formatTimestamp: (v) => String(v || ''),
    getQuestionText: () => '',
    createDataServiceErrorResponse: (m) => ({ success: false, message: m }),
    isAdministrator: () => false,
    __driveState: driveState
  };
  Object.assign(ctx, overrides);
  delete ctx.__myEmail;
  delete ctx.__drive;

  vm.createContext(ctx);
  const source = fs.readFileSync(path.resolve(__dirname, '../src/DataApis.js'), 'utf8');
  vm.runInContext(source, ctx, { filename: 'DataApis.js' });
  return ctx;
}

// =====================================================================
// createTemplateForm: auth gate
// =====================================================================

test('createTemplateForm: unauthenticated → early return 認証が必要です', () => {
  const ctx = loadCtx({ getCurrentEmail: () => null });
  const res = ctx.createTemplateForm('board');
  assert.equal(res.success, false);
  assert.match(res.error, /認証/);
});

// =====================================================================
// createTemplateForm: board (default)
// =====================================================================

test('createTemplateForm: default → board テンプレート、回答(選択肢) を含む', () => {
  const ctx = loadCtx();
  const res = ctx.createTemplateForm();
  assert.equal(res.success, true);
  assert.equal(res.templateType, 'board');
  assert.match(res.formTitle, /\[掲示板\]/);
  assert.match(res.message, /掲示板/);

  const state = ctx.__lastForm._state();
  assert.equal(state.limitOne, true);
  const kinds = state.items.map(i => i.kind);
  assert.deepEqual(kinds, ['list', 'text', 'mc', 'paragraph']);
  // 「回答」の選択肢が用意されている
  // Why Array.from: vm.runInContext 内で生成された Array は別 realm のため
  //                 deepStrictEqual が prototype 不一致で落ちる。プレーン化して比較。
  const mc = state.items.find(i => i.kind === 'mc');
  assert.equal(mc.title, '回答');
  assert.deepEqual(Array.from(mc.choices), ['賛成', '反対', 'どちらでもない']);
});

test('createTemplateForm: 未知の templateType は board にフォールバック', () => {
  const ctx = loadCtx();
  const res = ctx.createTemplateForm('nonexistent');
  assert.equal(res.templateType, 'board');
  const kinds = ctx.__lastForm._state().items.map(i => i.kind);
  assert.deepEqual(kinds, ['list', 'text', 'mc', 'paragraph']);
});

// =====================================================================
// createTemplateForm: numberline (M1)
// =====================================================================

test('createTemplateForm: numberline → 立場(scale 1-5) + 理由', () => {
  const ctx = loadCtx();
  const res = ctx.createTemplateForm('numberline');
  assert.equal(res.success, true);
  assert.equal(res.templateType, 'numberline');
  assert.match(res.formTitle, /\[数直線\]/);

  const items = ctx.__lastForm._state().items;
  const kinds = items.map(i => i.kind);
  assert.deepEqual(kinds, ['list', 'text', 'scale', 'paragraph']);
  const scale = items.find(i => i.kind === 'scale');
  assert.equal(scale.title, '立場');
  assert.deepEqual(Array.from(scale.bounds), [1, 5]);
  assert.ok(scale.labels && scale.labels.length === 2);
});

// =====================================================================
// createTemplateForm: matrix (M2)
// =====================================================================

test('createTemplateForm: matrix → X軸/Y軸 の 2 つの線形尺度 + 理由', () => {
  const ctx = loadCtx();
  const res = ctx.createTemplateForm('matrix');
  assert.equal(res.success, true);
  assert.equal(res.templateType, 'matrix');
  assert.match(res.formTitle, /\[マトリクス\]/);

  const items = ctx.__lastForm._state().items;
  const kinds = items.map(i => i.kind);
  assert.deepEqual(kinds, ['list', 'text', 'scale', 'scale', 'paragraph']);

  const scales = items.filter(i => i.kind === 'scale');
  assert.equal(scales.length, 2);
  assert.match(scales[0].title, /X軸/);
  assert.match(scales[1].title, /Y軸/);
  assert.deepEqual(Array.from(scales[0].bounds), [1, 5]);
  assert.deepEqual(Array.from(scales[1].bounds), [1, 5]);
});

// =====================================================================
// Drive: moveTo() preferred over addFile/removeFile
// =====================================================================

test('createTemplateForm: ファイル移動は moveTo() を使う（単一親モデル）', () => {
  const ctx = loadCtx();
  const res = ctx.createTemplateForm('board');
  assert.equal(res.success, true);

  // form/ss のファイルが moveTo されている
  const state = ctx.__driveState;
  const formFile = state.files[ctx.__lastForm.getId()];
  assert.ok(formFile._movedTo, 'form file should be moved via moveTo()');
  assert.equal(formFile._movedTo._name, 'みんなの回答ボード');

  // 旧 API (root.removeFile) は呼ばれていない
  assert.equal(state.rootFolder._removed.length, 0,
    'should not fall back to DriveApp.getRootFolder().removeFile()');
});

test('createTemplateForm: moveTo() 非対応環境では addFile + removeFile にフォールバック', () => {
  // moveTo を持たないファイルを返す Drive スタブ
  const driveState = {
    foldersByName: {},
    rootFolder: { _removed: [], removeFile: function (f) { this._removed.push(f); } },
    files: {}
  };
  const ctx = loadCtx({
    __drive: driveState,
    DriveApp: undefined // 後で上書き
  });
  // ↑ loadCtx で DriveApp は default のものが入る。
  //   moveTo 非対応版に差し替えるため、もう一度コンテキストを作り直す。
  const ctx2 = loadCtx({
    DriveApp: {
      getFoldersByName: () => ({ hasNext: () => false, next: () => null }),
      createFolder: (name) => {
        const f = makeFakeFolder(name, 'teacher@example.com');
        driveState.foldersByName[name] = [f];
        return f;
      },
      getFileById: (id) => {
        // moveTo を持たない
        if (!driveState.files[id]) driveState.files[id] = makeFakeDriveFile(id, false);
        return driveState.files[id];
      },
      getRootFolder: () => driveState.rootFolder
    }
  });
  const res = ctx2.createTemplateForm('board');
  assert.equal(res.success, true);

  // root.removeFile が呼ばれている（fallback path）
  assert.ok(driveState.rootFolder._removed.length >= 1,
    'fallback path should call DriveApp.getRootFolder().removeFile()');
});

// =====================================================================
// createUserFolder: owner filter
// =====================================================================

test('createUserFolder: 同名フォルダが他人所有なら無視して新規作成', () => {
  // 既存に他人所有の「みんなの回答ボード」がある状況
  const stranger = makeFakeFolder('みんなの回答ボード', 'someone-else@example.com');
  const driveState = {
    foldersByName: { 'みんなの回答ボード': [stranger] },
    rootFolder: { _removed: [], removeFile: () => {} },
    files: {}
  };
  const ctx = loadCtx({
    __drive: driveState,
    __myEmail: 'teacher@example.com',
    DriveApp: {
      getFoldersByName: (name) => {
        const arr = driveState.foldersByName[name] || [];
        let idx = 0;
        return { hasNext: () => idx < arr.length, next: () => arr[idx++] };
      },
      createFolder: (name) => {
        const f = makeFakeFolder(name, 'teacher@example.com');
        driveState.foldersByName[name].push(f);
        return f;
      },
      getFileById: (id) => driveState.files[id] || (driveState.files[id] = makeFakeDriveFile(id, true)),
      getRootFolder: () => driveState.rootFolder
    }
  });

  const folder = ctx.createUserFolder();
  assert.ok(folder);
  // 他人のフォルダではなく、新規作成（オーナーが自分）
  const owner = folder.getOwner().getEmail();
  assert.equal(owner, 'teacher@example.com');
  // 既存（他人所有）はそのまま残っている = 新規作成された
  assert.equal(driveState.foldersByName['みんなの回答ボード'].length, 2);
});

test('createUserFolder: 自分所有のフォルダがあれば再利用', () => {
  const mine = makeFakeFolder('みんなの回答ボード', 'teacher@example.com');
  const driveState = {
    foldersByName: { 'みんなの回答ボード': [mine] },
    rootFolder: { _removed: [], removeFile: () => {} },
    files: {}
  };
  let createCalled = false;
  const ctx = loadCtx({
    __drive: driveState,
    __myEmail: 'teacher@example.com',
    DriveApp: {
      getFoldersByName: (name) => {
        const arr = driveState.foldersByName[name] || [];
        let idx = 0;
        return { hasNext: () => idx < arr.length, next: () => arr[idx++] };
      },
      createFolder: () => { createCalled = true; return mine; },
      getFileById: (id) => driveState.files[id] || (driveState.files[id] = makeFakeDriveFile(id, true)),
      getRootFolder: () => driveState.rootFolder
    }
  });

  const folder = ctx.createUserFolder();
  assert.equal(folder, mine);
  assert.equal(createCalled, false, 'must not create a new folder when own folder exists');
});

test('createUserFolder: getOwner() が例外を投げても skip して次へ', () => {
  const broken = {
    _name: 'みんなの回答ボード',
    getName: function () { return this._name; },
    getOwner: function () { throw new Error('permission denied'); },
    getUrl: () => 'x',
    addFile: () => {},
    removeFile: () => {}
  };
  const mine = makeFakeFolder('みんなの回答ボード', 'teacher@example.com');
  const driveState = {
    foldersByName: { 'みんなの回答ボード': [broken, mine] },
    rootFolder: { _removed: [], removeFile: () => {} },
    files: {}
  };
  const ctx = loadCtx({
    __drive: driveState,
    __myEmail: 'teacher@example.com',
    DriveApp: {
      getFoldersByName: (name) => {
        const arr = driveState.foldersByName[name] || [];
        let idx = 0;
        return { hasNext: () => idx < arr.length, next: () => arr[idx++] };
      },
      createFolder: () => makeFakeFolder('みんなの回答ボード', 'teacher@example.com'),
      getFileById: (id) => driveState.files[id] || (driveState.files[id] = makeFakeDriveFile(id, true)),
      getRootFolder: () => driveState.rootFolder
    }
  });

  const folder = ctx.createUserFolder();
  assert.equal(folder, mine, 'should skip broken folder and pick the owned one');
});
