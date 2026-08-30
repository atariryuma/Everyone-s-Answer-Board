/**
 * 授業モード (native 入力) のテスト。
 *
 * Why: このモードの肝は「フェーズが画面の権能を決める」ことなので、UI で入力欄を
 *      隠しているかどうかではなく、サーバが投稿を受理/拒否する条件を pin する。
 *      画面が壊れても、議論中に投稿が通ることはない、という保証をここで作る。
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { gasResponseStubs } = require('./_helpers.cjs');

const LESSON_SOURCE = fs.readFileSync(path.resolve(__dirname, '../src/LessonService.js'), 'utf8');

// 汎用の in-memory シート。1-based row/col。
function createSheet(headers, name) {
  const data = [headers.slice()];
  return {
    _name: name || 'sheet',
    _data: data,
    getName: () => name || 'sheet',
    setName: (n) => { name = n; return this; },
    getLastRow: () => data.length,
    getLastColumn: () => headers.length,
    getRange: (row, col, numRows, numCols) => {
      const r = row - 1, c = col - 1;
      const rows = numRows || 1, cols = numCols || 1;
      return {
        getValues: () => {
          const out = [];
          for (let i = 0; i < rows; i++) {
            const src = data[r + i] || [];
            out.push(src.slice(c, c + cols));
          }
          return out;
        },
        setValues: (vs) => {
          for (let i = 0; i < vs.length; i++) {
            if (!data[r + i]) data[r + i] = [];
            for (let j = 0; j < vs[i].length; j++) data[r + i][c + j] = vs[i][j];
          }
        },
        setValue: (v) => { if (!data[r]) data[r] = []; data[r][c] = v; }
      };
    },
    appendRow: (row) => { data.push(row.slice()); },
    deleteRow: (i) => { data.splice(i - 1, 1); },
    createTextFinder: (query) => ({
      matchEntireCell: () => ({
        findNext: () => {
          for (let i = 1; i < data.length; i++) {
            for (let j = 0; j < data[i].length; j++) {
              if (data[i][j] === query) return { getRow: () => i + 1, getColumn: () => j + 1 };
            }
          }
          return null;
        },
        findAll: () => []
      })
    })
  };
}

const LESSONS_HEADERS = ['lessonId', 'userId', 'name', 'state', 'createdAt', 'startedAt', 'endedAt', 'schemaVersion', 'sizeBytes', 'etag', 'lessonJson'];
const RESPONSES_HEADERS = ['lessonId', 'phaseIndex', 'rowIndex', 'timestamp', 'class', 'answer', 'reason', 'numericX', 'numericY'];

function loadContext(overrides = {}) {
  const lessonsSheet = createSheet(LESSONS_HEADERS, 'lessons');
  const responsesSheet = createSheet(RESPONSES_HEADERS, 'lesson_responses');
  // native 回答スプレッドシート (授業に 1 つ、フェーズごとに 1 シート)。
  const nativeSheets = new Map();
  const sharingCalls = [];
  const editorsAdded = [];
  const cacheBumps = [];
  let uuidCounter = 0;
  let currentEmail = overrides.currentEmail || 'teacher@example.com';

  const nativeSs = {
    getId: () => 'native_ss_1',
    getSheets: () => Array.from(nativeSheets.values()),
    addEditor: (email) => { editorsAdded.push(email); },
    getSheetByName: (n) => nativeSheets.get(n) || null,
    insertSheet: (n) => {
      const s = createSheet([], n);
      s._data.length = 0;  // header は呼び出し側が appendRow する
      nativeSheets.set(n, s);
      return s;
    }
  };

  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    ...gasResponseStubs(),
    openDatabase: () => ({
      getSheetByName: (name) => name === 'lessons' ? lessonsSheet
        : name === 'lesson_responses' ? responsesSheet : null,
      getSheets: () => [{ getName: () => 'lessons' }, { getName: () => 'lesson_responses' }]
    }),
    SpreadsheetApp: {
      openById: (id) => id === 'native_ss_1' ? nativeSs
        : ({ getSheetByName: (name) => name === 'lessons' ? lessonsSheet : null }),
      create: () => nativeSs
    },
    // 児童は SS 直接権限を持たないので、本番では SA pool 経由で開かれる。
    // 本物の openSpreadsheet は { spreadsheet, auth, accessMode, getSheet(name) } を返す。
    //   偽物が中身を直接返していたため、getSheetByName 誤用が本番まで素通りした。
    openSpreadsheet: overrides.openSpreadsheet || (() => ({
      spreadsheet: nativeSs,
      auth: { isValid: true },
      accessMode: 'sa',
      getSheet: (name) => nativeSs.getSheetByName(name)
    })),
    applySpreadsheetSharingDefaults: (id) => { sharingCalls.push(id); return { saAdded: true }; },
    bumpBoardDataVersion_: (uid) => { cacheBumps.push(uid); },
    LESSONS_SHEET_HEADERS: LESSONS_HEADERS,
    LESSON_RESPONSES_SHEET_HEADERS: RESPONSES_HEADERS,
    deepClone: (v) => (v === null || v === undefined) ? v : JSON.parse(JSON.stringify(v)),
    getCachedProperty: (k) => k === 'DATABASE_SPREADSHEET_ID' ? 'db-id' : null,
    getCurrentEmail: () => currentEmail,
    isAdministrator: (email) => email === 'admin@example.com',
    findUserByEmail: (email) => (email === 'teacher@example.com')
      ? { userId: 'u1', userEmail: email } : { userId: 'stu', userEmail: email },
    findUserById: (id) => id === 'u1' ? { userId: 'u1', userEmail: 'teacher@example.com' } : null,
    createTemplateForm: () => ({ success: true, formId: 'f1', formUrl: 'https://forms.example/1', spreadsheetId: 'ss1', sheetName: 'フォームの回答 1' }),
    applyConfigPatch_: () => ({ success: true }),
    getPublishedSheetData: () => ({ success: true, data: [] }),
    getAllUsers: () => [],
    getConfigOrDefault: overrides.getConfigOrDefault || (() => ({})),
    ScriptApp: { getProjectTriggers: () => [], newTrigger: () => ({ timeBased: () => ({ everyDays: () => ({ atHour: () => ({ create: () => {} }) }) }) }) },
    Date: Date,
    FormApp: { openById: () => ({ setAcceptingResponses: () => {} }) },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
    Utilities: {
      getUuid: () => `uuid${++uuidCounter}aaaaaaaaaaaa`,
      computeDigest: (_a, s) => [s.length & 0xff, 0, 0, 0, 0, 0, 0, 0],
      DigestAlgorithm: { SHA_256: 'SHA_256' },
      Charset: { UTF_8: 'UTF_8' },
      newBlob: (s) => ({ getBytes: () => Buffer.from(String(s), 'utf8') })
    }
  };

  vm.createContext(context);
  vm.runInContext(LESSON_SOURCE, context, { filename: 'LessonService.js' });

  return {
    context, lessonsSheet, nativeSheets, sharingCalls, cacheBumps, editorsAdded,
    setEmail: (e) => { currentEmail = e; }
  };
}

// 授業を native テンプレで開始し、active にして返す。
function startNativeLesson(h) {
  const draft = h.context.createLessonDraft('u1', 'ロレンゾの友達', 'dialogue-reconsider-5phase');
  assert.equal(draft.success, true, 'draft 作成に失敗');
  const lessonId = draft.data.lesson.lessonId;
  h.context.updateLessonDraft('u1', lessonId, 'classes', ['6年1組']);
  const started = h.context.startLesson('u1', lessonId);
  assert.equal(started.success, true, `startLesson 失敗: ${started.message}`);
  return lessonId;
}

// =====================================================================
// テンプレート / draft
// =====================================================================

test('dialogue-reconsider-5phase: 5 フェーズ + 権能が定義される', () => {
  const h = loadContext();
  const res = h.context.createLessonDraft('u1', 'ロレンゾ', 'dialogue-reconsider-5phase');
  assert.equal(res.success, true);
  const phases = Array.from(res.data.lesson.lessonJson.phases);
  assert.deepEqual(phases.map(p => p.name),
    ['考える', '出会う', '議論する', 'もう一度考える', 'ふりかえる']);
  assert.deepEqual(phases.map(p => p.screenRole),
    ['input', 'browse', 'discuss', 'reinput', 'reflect']);
  assert.equal(res.data.lesson.lessonJson.inputMode, 'native');
});

test('従来テンプレは inputMode=form のまま (掲示板モードに影響しない)', () => {
  const h = loadContext();
  const res = h.context.createLessonDraft('u1', '通常授業', 'doutoku-3phase');
  assert.equal(res.data.lesson.lessonJson.inputMode, 'form');
});

// =====================================================================
// startLesson (native)
// =====================================================================

test('startLesson: native は Form を作らず、フェーズごとに回答シートを用意する', () => {
  const h = loadContext();
  const lessonId = startNativeLesson(h);

  // 5 フェーズ分のシートができ、Form は 1 つも作られない。
  assert.deepEqual(Array.from(h.nativeSheets.keys()).sort(),
    ['phase1', 'phase2', 'phase3', 'phase4', 'phase5']);
  const lessonJson = JSON.parse(h.lessonsSheet._data[1][10]);
  const phases = Array.from(lessonJson.phases);
  assert.equal(phases.every(p => !p.formId), true, 'native なのに Form が作られている');
  assert.equal(phases.every(p => p.spreadsheetId === 'native_ss_1'), true);
  assert.equal(phases[0].sheetName, 'phase1');
  assert.ok(lessonId);
});

test('startLesson: native 回答シートに SA pool 共有を当てる (児童の投稿経路が開く)', () => {
  const h = loadContext();
  startNativeLesson(h);
  assert.deepEqual(h.sharingCalls, ['native_ss_1']);
});

test('startLesson: 回答シートにボード所有者を editor として加える', () => {
  const h = loadContext();
  startNativeLesson(h);
  // owner は自分のボードを openById で開く経路を通るので、編集権が無いと
  //   管理者が代理で開始したときに教師自身のボードが開けなくなる。
  assert.deepEqual(Array.from(h.editorsAdded), ['teacher@example.com']);
});

test('startLesson: native シートは Form 回答シートと同じ列構成で作られる', () => {
  const h = loadContext();
  startNativeLesson(h);
  // vm context の Array なので Array.from で realm を揃えてから比較する。
  const headers = Array.from(h.nativeSheets.get('phase1')._data[0]);
  assert.deepEqual(headers,
    ['タイムスタンプ', 'メールアドレス', 'クラス', '名前', '横軸', '縦軸', '理由', '加わったこと']);
});

// =====================================================================
// フェーズごとの表示データ源
// =====================================================================

test('出会うフェーズは「考える」のシートを見る (自分のシートは空なので)', () => {
  const h = loadContext();
  const draft = h.context.createLessonDraft('u1', 'ロレンゾ', 'dialogue-reconsider-5phase');
  const lessonJson = draft.data.lesson.lessonJson;
  const phases = lessonJson.phases;
  // startLesson 相当のシート割当を手で入れる
  for (let i = 0; i < phases.length; i++) {
    phases[i].spreadsheetId = 'native_ss_1';
    phases[i].sheetName = 'phase' + (i + 1);
  }

  const browse = h.context.__buildPhaseConfigPatch_(phases[1], lessonJson, 'l1');
  assert.equal(browse.sheetName, 'phase1', '出会うフェーズが空のシートを指している');

  const discuss = h.context.__buildPhaseConfigPatch_(phases[2], lessonJson, 'l1');
  assert.equal(discuss.sheetName, 'phase1');
});

test('ふりかえりフェーズは「もう一度考える」のシートを見る (直近の入力フェーズ)', () => {
  const h = loadContext();
  const draft = h.context.createLessonDraft('u1', 'ロレンゾ', 'dialogue-reconsider-5phase');
  const lessonJson = draft.data.lesson.lessonJson;
  const phases = lessonJson.phases;
  for (let i = 0; i < phases.length; i++) {
    phases[i].spreadsheetId = 'native_ss_1';
    phases[i].sheetName = 'phase' + (i + 1);
  }
  const reflect = h.context.__buildPhaseConfigPatch_(phases[4], lessonJson, 'l1');
  assert.equal(reflect.sheetName, 'phase4');
});

test('入力フェーズは自分のシートを見る', () => {
  const h = loadContext();
  const draft = h.context.createLessonDraft('u1', 'ロレンゾ', 'dialogue-reconsider-5phase');
  const lessonJson = draft.data.lesson.lessonJson;
  const phases = lessonJson.phases;
  for (let i = 0; i < phases.length; i++) {
    phases[i].spreadsheetId = 'native_ss_1';
    phases[i].sheetName = 'phase' + (i + 1);
  }
  assert.equal(h.context.__buildPhaseConfigPatch_(phases[0], lessonJson, 'l1').sheetName, 'phase1');
  assert.equal(h.context.__buildPhaseConfigPatch_(phases[3], lessonJson, 'l1').sheetName, 'phase4');
});

test('もう一度考えるフェーズは「考える」と同じ軸を使う (● → ★ の比較が成立する条件)', () => {
  const h = loadContext();
  const draft = h.context.createLessonDraft('u1', 'ロレンゾ', 'dialogue-reconsider-5phase');
  const lessonJson = draft.data.lesson.lessonJson;
  const phases = lessonJson.phases;
  for (let i = 0; i < phases.length; i++) {
    phases[i].spreadsheetId = 'native_ss_1';
    phases[i].sheetName = 'phase' + (i + 1);
  }
  // 教師は最初の入力フェーズにだけ軸を設定する
  phases[0].templateOptions = {
    xLow: '自首を勧める', xHigh: '逃がす', yLow: '迷いあり', yHigh: '迷いなし'
  };
  const reinput = h.context.__buildPhaseConfigPatch_(phases[3], lessonJson, 'l1');
  assert.deepEqual(Object.assign({}, reinput.xAxisLabels), { min: '自首を勧める', max: '逃がす' });
  assert.deepEqual(Object.assign({}, reinput.yAxisLabels), { min: '迷いあり', max: '迷いなし' });
  // データ源は自分のシート (入力フェーズなので)
  assert.equal(reinput.sheetName, 'phase4');
});

test('phase 個別に軸を設定すればそちらが優先される', () => {
  const h = loadContext();
  const draft = h.context.createLessonDraft('u1', 'ロレンゾ', 'dialogue-reconsider-5phase');
  const lessonJson = draft.data.lesson.lessonJson;
  const phases = lessonJson.phases;
  for (let i = 0; i < phases.length; i++) {
    phases[i].spreadsheetId = 'native_ss_1';
    phases[i].sheetName = 'phase' + (i + 1);
  }
  phases[0].templateOptions = { xLow: 'A', xHigh: 'B', yLow: 'C', yHigh: 'D' };
  phases[3].templateOptions = { xLow: 'X', xHigh: 'Y', yLow: 'Z', yHigh: 'W' };
  const reinput = h.context.__buildPhaseConfigPatch_(phases[3], lessonJson, 'l1');
  assert.deepEqual(Object.assign({}, reinput.xAxisLabels), { min: 'X', max: 'Y' });
});

test('出会うフェーズは「考える」の軸ラベルを引き継ぐ (データと軸の意味を一致させる)', () => {
  const h = loadContext();
  const draft = h.context.createLessonDraft('u1', 'ロレンゾ', 'dialogue-reconsider-5phase');
  const lessonJson = draft.data.lesson.lessonJson;
  const phases = lessonJson.phases;
  for (let i = 0; i < phases.length; i++) {
    phases[i].spreadsheetId = 'native_ss_1';
    phases[i].sheetName = 'phase' + (i + 1);
  }
  // 教師は「考える」にだけ軸を設定した (出会う は既定のまま)
  phases[0].templateOptions = {
    xLow: '自首を勧める', xHigh: '逃がす', yLow: '迷いあり', yHigh: '迷いなし'
  };
  const browse = h.context.__buildPhaseConfigPatch_(phases[1], lessonJson, 'l1');
  assert.deepEqual(Object.assign({}, browse.xAxisLabels), { min: '自首を勧める', max: '逃がす' });
  assert.deepEqual(Object.assign({}, browse.yAxisLabels), { min: '迷いあり', max: '迷いなし' });
});

test('軸ラベルは native phase でも config に載る (ロレンゾの軸が board に届く)', () => {
  const h = loadContext();
  const draft = h.context.createLessonDraft('u1', 'ロレンゾ', 'dialogue-reconsider-5phase');
  const lessonJson = draft.data.lesson.lessonJson;
  const phases = lessonJson.phases;
  phases[0].spreadsheetId = 'native_ss_1';
  phases[0].sheetName = 'phase1';
  phases[0].templateOptions = {
    xLow: '自首を勧める', xHigh: '逃がす', yLow: '迷いあり', yHigh: '迷いなし'
  };
  const patch = h.context.__buildPhaseConfigPatch_(phases[0], lessonJson, 'l1');
  assert.deepEqual(Object.assign({}, patch.xAxisLabels), { min: '自首を勧める', max: '逃がす' });
  assert.deepEqual(Object.assign({}, patch.yAxisLabels), { min: '迷いあり', max: '迷いなし' });
  assert.equal(patch.displaySettings.boardMode, 'matrix');
});

// =====================================================================
// フェーズの権能 (submitLessonAnswer)
// =====================================================================

function submitAs(h, email, payload) {
  h.setEmail(email);
  return h.context.submitLessonAnswer('u1', payload);
}

// フェーズ送りは教師の操作。直前に児童として投稿していることが多いので、
//   email を教師に戻してから呼ぶ (本番でも owner 認証が要るのと同じ)。
function advanceAsTeacher(h, lessonId, targetIndex) {
  h.setEmail('teacher@example.com');
  const res = h.context.advanceLessonPhase('u1', lessonId, null, targetIndex);
  assert.equal(res.success, true, `advanceLessonPhase 失敗: ${res.message}`);
  return res;
}

// active な授業を指す config を返す helper。
function withActiveLesson(lessonId) {
  return () => ({ activeLessonId: lessonId });
}

test('submitLessonAnswer: input フェーズでは投稿できる', () => {
  const h = loadContext();
  const lessonId = startNativeLesson(h);
  h.context.getConfigOrDefault = withActiveLesson(lessonId);

  const res = submitAs(h, 'child1@example.com', {
    lessonId, phaseIndex: 0, numericX: 4, numericY: 2,
    reason: '友達だから助けたい', class: '6年1組', name: 'はるき'
  });
  assert.equal(res.success, true, res.message);
  const rows = h.nativeSheets.get('phase1')._data;
  assert.equal(rows.length, 2);
  assert.equal(rows[1][1], 'child1@example.com');
  assert.equal(rows[1][4], 4);
  assert.equal(rows[1][5], 2);
  assert.equal(rows[1][6], '友達だから助けたい');
});

test('submitLessonAnswer: 議論フェーズは投稿を拒否する (アプリが黙る時間を構造で守る)', () => {
  const h = loadContext();
  const lessonId = startNativeLesson(h);
  h.context.getConfigOrDefault = withActiveLesson(lessonId);
  // 議論フェーズ (index 2) へ
  advanceAsTeacher(h, lessonId, 2);

  const res = submitAs(h, 'child1@example.com', {
    lessonId, phaseIndex: 2, numericX: 3, numericY: 3, reason: 'あとから書く'
  });
  assert.equal(res.success, false);
  assert.match(res.message, /考えを送る時間ではありません/);
});

test('submitLessonAnswer: browse フェーズも拒否する (出会う時間は読むだけ)', () => {
  const h = loadContext();
  const lessonId = startNativeLesson(h);
  h.context.getConfigOrDefault = withActiveLesson(lessonId);
  advanceAsTeacher(h, lessonId, 1);

  const res = submitAs(h, 'child1@example.com', {
    lessonId, phaseIndex: 1, numericX: 3, numericY: 3, reason: 'x'
  });
  assert.equal(res.success, false);
});

test('submitLessonAnswer: フェーズがズレた投稿は PHASE_CHANGED で弾く', () => {
  const h = loadContext();
  const lessonId = startNativeLesson(h);
  h.context.getConfigOrDefault = withActiveLesson(lessonId);
  advanceAsTeacher(h, lessonId, 3);  // reinput へ

  // 児童の画面はまだ phase 0 のつもりで送ってくる
  const res = submitAs(h, 'child1@example.com', {
    lessonId, phaseIndex: 0, numericX: 3, numericY: 3, reason: 'x'
  });
  assert.equal(res.success, false);
  assert.match(res.message, /PHASE_CHANGED/);
});

test('submitLessonAnswer: 同一フェーズ内の再投稿は行を更新する (置き直しを許す)', () => {
  const h = loadContext();
  const lessonId = startNativeLesson(h);
  h.context.getConfigOrDefault = withActiveLesson(lessonId);

  submitAs(h, 'child1@example.com', { lessonId, phaseIndex: 0, numericX: 4, numericY: 2, reason: '最初' });
  const res2 = submitAs(h, 'child1@example.com', { lessonId, phaseIndex: 0, numericX: 2, numericY: 4, reason: 'やっぱりこう' });

  assert.equal(res2.success, true);
  assert.equal(res2.data.updated, true);
  const rows = h.nativeSheets.get('phase1')._data;
  assert.equal(rows.length, 2, '行が増えている (更新でなく追記された)');
  assert.equal(rows[1][6], 'やっぱりこう');
});

test('submitLessonAnswer: 別の児童は別行になる', () => {
  const h = loadContext();
  const lessonId = startNativeLesson(h);
  h.context.getConfigOrDefault = withActiveLesson(lessonId);

  submitAs(h, 'child1@example.com', { lessonId, phaseIndex: 0, numericX: 4, numericY: 2, reason: 'A' });
  submitAs(h, 'child2@example.com', { lessonId, phaseIndex: 0, numericX: 1, numericY: 5, reason: 'B' });
  assert.equal(h.nativeSheets.get('phase1')._data.length, 3);
});

test('submitLessonAnswer: 児童の同定は payload ではなく Session の email で行う', () => {
  const h = loadContext();
  const lessonId = startNativeLesson(h);
  h.context.getConfigOrDefault = withActiveLesson(lessonId);

  // 他人になりすまそうとする payload を送っても、記録されるのは実際の Session email。
  const res = submitAs(h, 'child1@example.com', {
    lessonId, phaseIndex: 0, numericX: 3, numericY: 3, reason: 'x',
    email: 'victim@example.com', userEmail: 'victim@example.com'
  });
  assert.equal(res.success, true);
  assert.equal(h.nativeSheets.get('phase1')._data[1][1], 'child1@example.com');
});

test('submitLessonAnswer: 理由が空なら拒否 (座標だけの投稿を作らない)', () => {
  const h = loadContext();
  const lessonId = startNativeLesson(h);
  h.context.getConfigOrDefault = withActiveLesson(lessonId);

  const res = submitAs(h, 'child1@example.com', { lessonId, phaseIndex: 0, numericX: 3, numericY: 3, reason: '   ' });
  assert.equal(res.success, false);
  assert.match(res.message, /理由/);
});

test('submitLessonAnswer: 尺度の範囲外は拒否', () => {
  const h = loadContext();
  const lessonId = startNativeLesson(h);
  h.context.getConfigOrDefault = withActiveLesson(lessonId);

  assert.equal(submitAs(h, 'c@example.com', { lessonId, phaseIndex: 0, numericX: 9, numericY: 3, reason: 'x' }).success, false);
  assert.equal(submitAs(h, 'c@example.com', { lessonId, phaseIndex: 0, numericX: 3, numericY: 0, reason: 'x' }).success, false);
  assert.equal(submitAs(h, 'c@example.com', { lessonId, phaseIndex: 0, numericX: 'あ', numericY: 3, reason: 'x' }).success, false);
});

test('submitLessonAnswer: 長すぎるテキストは切り詰める', () => {
  const h = loadContext();
  const lessonId = startNativeLesson(h);
  h.context.getConfigOrDefault = withActiveLesson(lessonId);

  submitAs(h, 'c@example.com', { lessonId, phaseIndex: 0, numericX: 3, numericY: 3, reason: 'あ'.repeat(900) });
  assert.equal(h.nativeSheets.get('phase1')._data[1][6].length, 500);
});

test('submitLessonAnswer: 投稿で board cache を落とす (他の児童の画面に届く)', () => {
  const h = loadContext();
  const lessonId = startNativeLesson(h);
  h.context.getConfigOrDefault = withActiveLesson(lessonId);

  submitAs(h, 'c@example.com', { lessonId, phaseIndex: 0, numericX: 3, numericY: 3, reason: 'x' });
  assert.deepEqual(h.cacheBumps, ['u1']);
});

test('submitLessonAnswer: 授業が active でなければ拒否', () => {
  const h = loadContext();
  const lessonId = startNativeLesson(h);
  h.context.getConfigOrDefault = withActiveLesson(lessonId);
  h.context.endLesson('u1', lessonId);

  const res = submitAs(h, 'c@example.com', { lessonId, phaseIndex: 0, numericX: 3, numericY: 3, reason: 'x' });
  assert.equal(res.success, false);
});

// =====================================================================
// フェーズ配信 (__getViewerLessonPhase_)
// =====================================================================

test('__getViewerLessonPhase_: 児童には現在フェーズだけを返す (未来の問いは見せない)', () => {
  const h = loadContext();
  const lessonId = startNativeLesson(h);
  h.context.getConfigOrDefault = withActiveLesson(lessonId);

  const info = h.context.__getViewerLessonPhase_('u1');
  assert.equal(info.phaseIndex, 0);
  assert.equal(info.screenRole, 'input');
  assert.equal(info.phaseName, '考える');
  assert.equal(info.phaseCount, 5);
  // 全フェーズの一覧は含まれない
  assert.equal(info.phases, undefined);
});

test('__getViewerLessonPhase_: フェーズを進めると screenRole が変わる', () => {
  const h = loadContext();
  const lessonId = startNativeLesson(h);
  h.context.getConfigOrDefault = withActiveLesson(lessonId);

  advanceAsTeacher(h, lessonId, 2);
  assert.equal(h.context.__getViewerLessonPhase_('u1').screenRole, 'discuss');
  advanceAsTeacher(h, lessonId, 4);
  assert.equal(h.context.__getViewerLessonPhase_('u1').screenRole, 'reflect');
});

test('__getViewerLessonPhase_: Form 経由の授業では null (掲示板モードは無関係)', () => {
  const h = loadContext();
  const draft = h.context.createLessonDraft('u1', '通常', 'doutoku-3phase');
  const lessonId = draft.data.lesson.lessonId;
  h.context.updateLessonDraft('u1', lessonId, 'classes', ['6年1組']);
  h.context.startLesson('u1', lessonId);
  h.context.getConfigOrDefault = withActiveLesson(lessonId);

  assert.equal(h.context.__getViewerLessonPhase_('u1'), null);
});

test('__getViewerLessonPhase_: 授業をしていなければ null', () => {
  const h = loadContext();
  h.context.getConfigOrDefault = () => ({});
  assert.equal(h.context.__getViewerLessonPhase_('u1'), null);
});

// =====================================================================
// 航跡 (getMyLessonTrajectory)
// =====================================================================

test('getMyLessonTrajectory: 入力フェーズだけが航跡の点になる', () => {
  const h = loadContext();
  const lessonId = startNativeLesson(h);
  h.context.getConfigOrDefault = withActiveLesson(lessonId);

  submitAs(h, 'child1@example.com', { lessonId, phaseIndex: 0, numericX: 4, numericY: 2, reason: '友達だから助けたい' });
  advanceAsTeacher(h, lessonId, 3);
  submitAs(h, 'child1@example.com', {
    lessonId, phaseIndex: 3, numericX: 4, numericY: 4,
    reason: 'やっぱり逃がす', addedInsight: '友情には厳しさも含まれる'
  });

  h.setEmail('child1@example.com');
  const res = h.context.getMyLessonTrajectory('u1');
  assert.equal(res.success, true);
  const phases = Array.from(res.data.phases);
  // browse / discuss は投稿が存在しないので点にならない
  assert.deepEqual(phases.map(p => p.phaseIndex), [0, 3]);
  assert.equal(phases[0].numericX, 4);
  assert.equal(phases[0].numericY, 2);
  assert.equal(phases[1].numericY, 4);
  assert.equal(phases[1].addedInsight, '友情には厳しさも含まれる');
});

test('getMyLessonTrajectory: 位置が変わらなくても記述の変化は残る', () => {
  const h = loadContext();
  const lessonId = startNativeLesson(h);
  h.context.getConfigOrDefault = withActiveLesson(lessonId);

  submitAs(h, 'child1@example.com', { lessonId, phaseIndex: 0, numericX: 4, numericY: 2, reason: '友達だから' });
  advanceAsTeacher(h, lessonId, 3);
  submitAs(h, 'child1@example.com', {
    lessonId, phaseIndex: 3, numericX: 4, numericY: 2,
    reason: 'やっぱり逃がす。でもロレンゾのこれからを考えて決めたい',
    addedInsight: '友情には厳しさも含まれる'
  });

  h.setEmail('child1@example.com');
  const phases = Array.from(h.context.getMyLessonTrajectory('u1').data.phases);
  assert.equal(phases[0].numericX, phases[1].numericX);
  assert.equal(phases[0].numericY, phases[1].numericY);
  assert.notEqual(phases[0].reason, phases[1].reason);
  assert.ok(phases[1].addedInsight);
});

test('getMyLessonTrajectory: 他人の航跡は返らない (自分の分だけ)', () => {
  const h = loadContext();
  const lessonId = startNativeLesson(h);
  h.context.getConfigOrDefault = withActiveLesson(lessonId);

  submitAs(h, 'child1@example.com', { lessonId, phaseIndex: 0, numericX: 1, numericY: 1, reason: 'A' });
  submitAs(h, 'child2@example.com', { lessonId, phaseIndex: 0, numericX: 5, numericY: 5, reason: 'B' });

  h.setEmail('child2@example.com');
  const phases = Array.from(h.context.getMyLessonTrajectory('u1').data.phases);
  assert.equal(phases.length, 1);
  assert.equal(phases[0].numericX, 5);
  assert.equal(phases[0].reason, 'B');
});

// =====================================================================
// 移行由来の授業の Form 復旧 (closeLessonForms)
//
// profiles から移行した授業は formId に公開 URL 側の ID が入っており、
// FormApp.openById が通らない = フェーズごとの Form 開閉が効いていなかった。
// 回答スプレッドシートが編集用 URL を知っているので、そこから復旧する。
// =====================================================================

// 公開 ID では開けず、編集用 URL 経由でのみ開ける Form を再現する harness。
function loadFormRepairContext() {
  const h = loadContext();
  const opened = [];
  const acceptingCalls = [];
  const form = {
    getId: () => 'EDIT_ID_1',
    setAcceptingResponses: (b) => { acceptingCalls.push(b); }
  };
  h.context.FormApp = {
    openById: (id) => {
      opened.push(['byId', id]);
      if (id !== 'EDIT_ID_1') throw new Error('指定した ID のアイテムは見つかりませんでした');
      return form;
    },
    openByUrl: (url) => {
      opened.push(['byUrl', url]);
      if (url !== 'https://docs.google.com/forms/d/EDIT_ID_1/edit') throw new Error('not found');
      return form;
    }
  };
  // 回答 SS だけを差し替え、それ以外 (native の回答シート等) は元の挙動に委ねる。
  const originalOpenById = h.context.SpreadsheetApp.openById;
  h.context.SpreadsheetApp = Object.assign({}, h.context.SpreadsheetApp, {
    openById: (id) => {
      if (id === 'resp_ss_1') {
        return { getFormUrl: () => 'https://docs.google.com/forms/d/EDIT_ID_1/edit' };
      }
      if (id === 'missing_ss') return { getFormUrl: () => null };
      return originalOpenById(id);
    }
  });
  return { h, opened, acceptingCalls };
}

test('closeLessonForms: 公開 ID しか無くても回答シート経由で Form を締め切る', () => {
  const { h, acceptingCalls } = loadFormRepairContext();
  const draft = h.context.createLessonDraft('u1', '移行授業', 'doutoku-3phase');
  const lessonId = draft.data.lesson.lessonId;
  // 移行由来を再現: formId は公開 URL 側の ID、回答 SS はある
  const row = h.lessonsSheet._data[1];
  const j = JSON.parse(row[10]);
  j.phases = [{ name: '導入', formTemplate: 'numberline', formId: '1FAIpQLS_public', spreadsheetId: 'resp_ss_1', sheetName: 'フォームの回答 1' }];
  row[10] = JSON.stringify(j);

  const res = h.context.closeLessonForms('u1', lessonId);
  assert.equal(res.success, true, res.message);
  assert.equal(res.data.closed, 1);
  assert.equal(res.data.failed, 0);
  assert.equal(res.data.repaired, 1, '編集用 ID への復旧が記録されていない');
  assert.deepEqual(Array.from(acceptingCalls), [false]);
});

test('closeLessonForms: 復旧した編集用 ID を保存する (次回は 1 発で開く)', () => {
  const { h } = loadFormRepairContext();
  const draft = h.context.createLessonDraft('u1', '移行授業', 'doutoku-3phase');
  const lessonId = draft.data.lesson.lessonId;
  const row = h.lessonsSheet._data[1];
  const j = JSON.parse(row[10]);
  j.phases = [{ name: '導入', formTemplate: 'numberline', formId: '1FAIpQLS_public', spreadsheetId: 'resp_ss_1', sheetName: 'x' }];
  row[10] = JSON.stringify(j);

  h.context.closeLessonForms('u1', lessonId);
  const after = JSON.parse(h.lessonsSheet._data[1][10]);
  assert.equal(after.phases[0].formId, 'EDIT_ID_1');
});

test('closeLessonForms: 開けないフェーズは失敗として名前付きで報告する', () => {
  const { h } = loadFormRepairContext();
  const draft = h.context.createLessonDraft('u1', '移行授業', 'doutoku-3phase');
  const lessonId = draft.data.lesson.lessonId;
  const row = h.lessonsSheet._data[1];
  const j = JSON.parse(row[10]);
  j.phases = [{ name: '壊れたフェーズ', formTemplate: 'numberline', formId: '1FAIpQLS_public', spreadsheetId: 'missing_ss' }];
  row[10] = JSON.stringify(j);

  const res = h.context.closeLessonForms('u1', lessonId);
  assert.equal(res.data.failed, 1);
  assert.match(res.message, /壊れたフェーズ/);
});

test('closeLessonForms: 授業モード (native) は Form を使わないので no-op', () => {
  const { h } = loadFormRepairContext();
  const lessonId = startNativeLesson(h);
  const res = h.context.closeLessonForms('u1', lessonId);
  assert.equal(res.success, true);
  assert.equal(res.data.total, 0);
  assert.match(res.message, /使っていません/);
});

test('closeLessonForms: 所有者以外は実行できない', () => {
  const { h } = loadFormRepairContext();
  const draft = h.context.createLessonDraft('u1', '移行授業', 'doutoku-3phase');
  h.setEmail('someone@example.com');
  const res = h.context.closeLessonForms('u1', draft.data.lesson.lessonId);
  assert.equal(res.success, false);
});

// =====================================================================
// 教師の見取りグリッド (getLessonReviewGrid)
// =====================================================================

// 3 人の児童に、動かない子 / 大きく動く子 / 少し動く子 の航跡を作る。
function seedThreeStudents(h, lessonId) {
  h.context.getConfigOrDefault = withActiveLesson(lessonId);
  // phase 0 (考える)
  submitAs(h, 'stay@example.com', { lessonId, phaseIndex: 0, numericX: 4, numericY: 2, reason: '友達だから助けたい', name: 'はるき' });
  submitAs(h, 'far@example.com', { lessonId, phaseIndex: 0, numericX: 5, numericY: 1, reason: '逃がす', name: 'みお' });
  submitAs(h, 'near@example.com', { lessonId, phaseIndex: 0, numericX: 2, numericY: 2, reason: '自首を勧める', name: 'そら' });
  // phase 3 (もう一度考える)
  advanceAsTeacher(h, lessonId, 3);
  submitAs(h, 'stay@example.com', {
    lessonId, phaseIndex: 3, numericX: 4, numericY: 2,
    reason: 'やっぱり逃がす。でもロレンゾのこれからを考えて決めたい',
    addedInsight: '友情には厳しさも含まれる', name: 'はるき'
  });
  submitAs(h, 'far@example.com', { lessonId, phaseIndex: 3, numericX: 1, numericY: 5, reason: '自首を勧める', name: 'みお' });
  submitAs(h, 'near@example.com', { lessonId, phaseIndex: 3, numericX: 3, numericY: 2, reason: 'すこし迷う', name: 'そら' });
  h.setEmail('teacher@example.com');
}

test('getLessonReviewGrid: 位置が動かなかった児童を先頭に並べる', () => {
  const h = loadContext();
  const lessonId = startNativeLesson(h);
  seedThreeStudents(h, lessonId);

  const res = h.context.getLessonReviewGrid('u1', lessonId);
  assert.equal(res.success, true, res.message);
  const students = Array.from(res.data.students);
  assert.deepEqual(students.map(s => s.name), ['はるき', 'そら', 'みお']);
  // 動かなかった児童は distance 0 かつ moved=false
  assert.equal(students[0].distance, 0);
  assert.equal(students[0].moved, false);
  assert.equal(students[2].moved, true);
});

test('getLessonReviewGrid: 位置不変でも理由の変化と加わったことが読める', () => {
  const h = loadContext();
  const lessonId = startNativeLesson(h);
  seedThreeStudents(h, lessonId);

  const students = Array.from(h.context.getLessonReviewGrid('u1', lessonId).data.students);
  const stay = students.find(s => s.name === 'はるき');
  assert.equal(stay.first.numericX, stay.last.numericX);
  assert.equal(stay.first.numericY, stay.last.numericY);
  assert.notEqual(stay.first.reason, stay.last.reason);
  assert.equal(stay.last.addedInsight, '友情には厳しさも含まれる');
});

test('getLessonReviewGrid: 1 フェーズしか答えていない児童は last が null で最後に並ぶ', () => {
  const h = loadContext();
  const lessonId = startNativeLesson(h);
  h.context.getConfigOrDefault = withActiveLesson(lessonId);
  submitAs(h, 'both@example.com', { lessonId, phaseIndex: 0, numericX: 3, numericY: 3, reason: 'A', name: 'ふたつ' });
  submitAs(h, 'once@example.com', { lessonId, phaseIndex: 0, numericX: 1, numericY: 1, reason: 'B', name: 'ひとつ' });
  advanceAsTeacher(h, lessonId, 3);
  submitAs(h, 'both@example.com', { lessonId, phaseIndex: 3, numericX: 5, numericY: 5, reason: 'C', name: 'ふたつ' });
  h.setEmail('teacher@example.com');

  const students = Array.from(h.context.getLessonReviewGrid('u1', lessonId).data.students);
  assert.equal(students[students.length - 1].name, 'ひとつ');
  assert.equal(students[students.length - 1].last, null);
  assert.equal(students[students.length - 1].answeredPhases, 1);
});

test('getLessonReviewGrid: 所有者以外は取得できない', () => {
  const h = loadContext();
  const lessonId = startNativeLesson(h);
  h.setEmail('someone@example.com');
  const res = h.context.getLessonReviewGrid('u1', lessonId);
  assert.equal(res.success, false);
});

test('getMyLessonTrajectory: 未投稿なら空 (先に他人を見せない)', () => {
  const h = loadContext();
  const lessonId = startNativeLesson(h);
  h.context.getConfigOrDefault = withActiveLesson(lessonId);

  h.setEmail('child9@example.com');
  const res = h.context.getMyLessonTrajectory('u1');
  assert.equal(res.success, true);
  assert.equal(Array.from(res.data.phases).length, 0);
});

// =====================================================================
// config パッチが実際の検証を通るか (本番で「開始」が失敗した回帰)
//
// Why これが要るか: パッチの中身だけを assert しても、それが
//   validateConfig を通るかは分からない。実際 answer と reason を同じ列 (6) に
//   割り当てていたため validateMapping の「列インデックス重複」で弾かれ、
//   教師が「開始」を押した瞬間に「設定検証エラー」で落ちていた。
//   ここでは本物の validators.js に通して、その経路ごと pin する。
// =====================================================================

const VALIDATORS_SOURCE = fs.readFileSync(path.resolve(__dirname, '../src/validators.js'), 'utf8');
const HELPERS_SOURCE = fs.readFileSync(path.resolve(__dirname, '../src/helpers.js'), 'utf8');

function loadValidator() {
  const ctx = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    SYSTEM_LIMITS: { DEFAULT_PAGE_SIZE: 20, MAX_PAGE_SIZE: 100 }
  };
  vm.createContext(ctx);
  vm.runInContext(HELPERS_SOURCE, ctx, { filename: 'helpers.js' });
  vm.runInContext(VALIDATORS_SOURCE, ctx, { filename: 'validators.js' });
  return ctx.validateConfig;
}

test('native phase の config パッチが validateConfig を通る (全 5 フェーズ)', () => {
  const validateConfig = loadValidator();
  const h = loadContext();
  const draft = h.context.createLessonDraft('u1', 'ロレンゾ', 'dialogue-reconsider-5phase');
  const lessonJson = draft.data.lesson.lessonJson;
  const phases = lessonJson.phases;
  for (let i = 0; i < phases.length; i++) {
    phases[i].spreadsheetId = '1g7cGEiskD7w7DO2a6lSjSsgYHwa_vYRyLnk_BI5fYmY';
    phases[i].sheetName = 'phase' + (i + 1);
    phases[i].columnMapping = h.context.__nativeColumnMapping_(phases[i].formTemplate);
  }
  phases[0].templateOptions = { xLow: '自首を勧める', xHigh: '逃がす', yLow: '迷いあり', yHigh: '迷いなし' };

  for (let i = 0; i < phases.length; i++) {
    const patch = h.context.__buildPhaseConfigPatch_(phases[i], lessonJson, 'lesson_x');
    const res = validateConfig(JSON.parse(JSON.stringify(patch)));
    assert.equal(res.isValid, true,
      `phase ${i} の patch が検証を通らない: ${JSON.stringify(res.errors)}`);
  }
});

test('__nativeColumnMapping_: answer と reason を同じ列にしない', () => {
  const h = loadContext();
  // validateMapping は numericX / numericY だけ重複を免除する
  //   (「立場」列を answer としても numericX としても見る設計)。
  //   それ以外が重複すると config 保存が落ちる。
  for (const tpl of ['matrix', 'numberline', 'board', 'pie']) {
    const m = h.context.__nativeColumnMapping_(tpl);
    const checked = Object.keys(m).filter((k) => k !== 'numericX' && k !== 'numericY');
    const idx = checked.map((k) => m[k]);
    assert.equal(new Set(idx).size, idx.length, `${tpl} で列が重複している: ${JSON.stringify(m)}`);
  }
});

test('__nativeColumnMapping_: answer は必須。matrix は座標、board は本文に割り当てる', () => {
  const h = loadContext();
  // answer は validateMapping の必須列なので、どのテンプレでも必ず存在する。
  for (const tpl of ['matrix', 'numberline', 'board', 'pie']) {
    assert.equal(typeof h.context.__nativeColumnMapping_(tpl).answer, 'number', tpl);
  }
  // matrix は「回答 = 座標」なので answer は軸列 (4)、本文は reason (6)。
  //   象限キーワード抽出が reason を見るので reason も必須。
  const mx = h.context.__nativeColumnMapping_('matrix');
  assert.equal(mx.answer, 4);
  assert.equal(mx.numericX, 4);
  assert.equal(mx.numericY, 5);
  assert.equal(mx.reason, 6);
  // board は自由記述そのものが回答。
  const bd = h.context.__nativeColumnMapping_('board');
  assert.equal(bd.answer, 6);
  assert.equal(bd.reason, undefined);
});

// =====================================================================
// 匿名性 (デジタルシティズンシップ)
//
// 授業モードは道徳的な立場を表明する場なので、児童同士に名前を見せない。
// 「誰の意見か」が見えると同調圧力が働く (少数派だと分かると言い直す等)。
// 教師は isOwnBoard なのでこの設定に関わらず名前を見られる。
// =====================================================================

test('授業モードは児童同士に名前を見せない (ボードの既存設定を引き継がない)', () => {
  const h = loadContext();
  const draft = h.context.createLessonDraft('u1', 'ロレンゾ', 'dialogue-reconsider-5phase');
  const lessonJson = draft.data.lesson.lessonJson;
  const phases = lessonJson.phases;
  for (let i = 0; i < phases.length; i++) {
    phases[i].spreadsheetId = 'ss1';
    phases[i].sheetName = 'phase' + (i + 1);
  }
  for (let i = 0; i < phases.length; i++) {
    const patch = h.context.__buildPhaseConfigPatch_(phases[i], lessonJson, 'l1');
    assert.equal(patch.displaySettings.showNames, false,
      `phase ${i} で名前が見える設定になっている`);
  }
});

test('phase が明示的に showNames=true を指定した場合だけ従う', () => {
  const h = loadContext();
  const draft = h.context.createLessonDraft('u1', 'ロレンゾ', 'dialogue-reconsider-5phase');
  const lessonJson = draft.data.lesson.lessonJson;
  const phase = lessonJson.phases[0];
  phase.spreadsheetId = 'ss1';
  phase.sheetName = 'phase1';
  phase.displaySettings = { showNames: true };
  const patch = h.context.__buildPhaseConfigPatch_(phase, lessonJson, 'l1');
  assert.equal(patch.displaySettings.showNames, true);
});

test('掲示板モード (Form 経由) の授業は既存の設定を尊重する', () => {
  const h = loadContext();
  const draft = h.context.createLessonDraft('u1', '通常授業', 'doutoku-3phase');
  const lessonJson = draft.data.lesson.lessonJson;
  const phase = lessonJson.phases[0];
  phase.formUrl = 'https://forms.example/1';
  phase.spreadsheetId = 'ss1';
  phase.sheetName = 'フォームの回答 1';
  const patch = h.context.__buildPhaseConfigPatch_(phase, lessonJson, 'l1');
  // native ではないので showNames に手を出さない (config 側の設定が生きる)
  assert.equal(patch.displaySettings.showNames, undefined);
});
