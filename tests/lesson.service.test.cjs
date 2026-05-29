/**
 * LessonService の CRUD + lifecycle テスト。
 *
 * Why: lesson は draft → active → completed の状態機械 + Form 一括生成という
 *      副作用の多い箇所なので、startLesson の冪等性 / partial failure / lock を
 *      テストで pin する。
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { gasResponseStubs } = require('./_helpers.cjs');

const LESSON_SOURCE = fs.readFileSync(path.resolve(__dirname, '../src/LessonService.js'), 'utf8');

// In-memory lessons sheet (Google Sheets API の最低限の subset を再現)。
//   現実の sheet と違って row index は 1-based。
function createFakeLessonsSheet(headers) {
  const data = [headers.slice()];
  let textFinderQuery = null;
  const sheet = {
    _data: data,
    getLastRow: () => data.length,
    getLastColumn: () => headers.length,
    getRange: (row, col, numRows, numCols) => {
      const r = row - 1, c = col - 1;
      const rows = numRows || 1;
      const cols = numCols || 1;
      return {
        getValues: () => {
          const out = [];
          for (let i = 0; i < rows; i++) {
            const sourceRow = data[r + i] || [];
            out.push(sourceRow.slice(c, c + cols));
          }
          return out;
        },
        setValue: (v) => { if (!data[r]) data[r] = []; data[r][c] = v; },
        setValues: (vs) => {
          for (let i = 0; i < vs.length; i++) {
            if (!data[r + i]) data[r + i] = [];
            for (let j = 0; j < vs[i].length; j++) {
              data[r + i][c + j] = vs[i][j];
            }
          }
        }
      };
    },
    appendRow: (row) => { data.push(row.slice()); },
    deleteRow: (rowIndex) => { data.splice(rowIndex - 1, 1); },
    createTextFinder: (query) => {
      textFinderQuery = query;
      // 全列を対象に exact-match を探す (実 Sheets API と同じ挙動)。
      return {
        matchEntireCell: () => ({
          findNext: () => {
            for (let i = 1; i < data.length; i++) {
              for (let j = 0; j < data[i].length; j++) {
                if (data[i][j] === textFinderQuery) {
                  return { getRow: () => i + 1, getColumn: () => j + 1 };
                }
              }
            }
            return null;
          },
          findAll: () => {
            const results = [];
            for (let i = 1; i < data.length; i++) {
              for (let j = 0; j < data[i].length; j++) {
                if (data[i][j] === textFinderQuery) {
                  results.push({ getRow: () => i + 1, getColumn: () => j + 1 });
                }
              }
            }
            return results;
          }
        })
      };
    }
  };
  return sheet;
}

function loadLessonContext(overrides = {}) {
  const LESSONS_HEADERS = ['lessonId', 'userId', 'name', 'state', 'createdAt', 'startedAt', 'endedAt', 'schemaVersion', 'sizeBytes', 'etag', 'lessonJson'];
  const lessonsSheet = overrides.lessonsSheet || createFakeLessonsSheet(LESSONS_HEADERS);
  const formCreations = [];
  const formCloses = [];
  const configPatches = [];

  let uuidCounter = 0;
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    ...gasResponseStubs(),
    openDatabase: () => ({
      getSheetByName: (name) => name === 'lessons' ? lessonsSheet : null,
      // Production の sheet 不存在チェック (SA proxy) に合わせて getSheets() も提供。
      getSheets: () => [{ getName: () => 'lessons' }]
    }),
    // SpreadsheetApp.openById で direct admin write (deleteLesson 用)。
    SpreadsheetApp: {
      openById: () => ({ getSheetByName: (name) => name === 'lessons' ? lessonsSheet : null })
    },
    LESSONS_SHEET_HEADERS: ['lessonId', 'userId', 'name', 'state', 'createdAt', 'startedAt', 'endedAt', 'schemaVersion', 'sizeBytes', 'etag', 'lessonJson'],
    // helpers.js の deepClone を test 環境にも提供 (LessonService が依存)。
    deepClone: (v) => (v === null || v === undefined) ? v : JSON.parse(JSON.stringify(v)),
    getCachedProperty: (k) => k === 'DATABASE_SPREADSHEET_ID' ? 'db-id' : null,
    getCurrentEmail: () => 'teacher@example.com',
    isAdministrator: (email) => email === 'admin@example.com',
    findUserByEmail: (email) => email === 'teacher@example.com' ? { userId: 'u1', userEmail: email } : null,
    createTemplateForm: overrides.createTemplateForm || ((templateType, templateOptions) => {
      formCreations.push({ templateType, templateOptions: templateOptions || null });
      return {
        success: true,
        formId: `form_${formCreations.length}`,
        formUrl: `https://forms.example/${formCreations.length}`,
        spreadsheetId: `ss_${formCreations.length}`,
        sheetName: 'フォームの回答 1'
      };
    }),
    applyConfigPatch_: overrides.applyConfigPatch_ || ((userId, patch) => {
      configPatches.push({ userId, patch });
      return { success: true };
    }),
    // Phase 2: snapshot capture が呼ぶ。デフォルトは empty rows。
    getPublishedSheetData: overrides.getPublishedSheetData || (() => ({
      success: true,
      data: []
    })),
    getAllUsers: overrides.getAllUsers || (() => []),
    getConfigOrDefault: overrides.getConfigOrDefault || (() => ({})),
    ScriptApp: overrides.ScriptApp || {
      getProjectTriggers: () => [],
      newTrigger: () => ({
        timeBased: () => ({ everyDays: () => ({ atHour: () => ({ create: () => {} }) }) })
      })
    },
    Date: Date,
    FormApp: {
      openById: (formId) => ({
        setAcceptingResponses: (b) => { formCloses.push({ formId, accepting: b }); }
      })
    },
    LockService: overrides.LockService === null ? undefined : (overrides.LockService || {
      getScriptLock: () => ({
        tryLock: () => true,
        releaseLock: () => {}
      })
    }),
    Utilities: {
      getUuid: () => `uuid${++uuidCounter}aaaaaaaaaaaa`,
      computeDigest: (_alg, str) => {
        // 8-byte fake digest derived from string length + first/last char codes
        const len = str.length;
        const first = str.charCodeAt(0) || 0;
        const last = str.charCodeAt(str.length - 1) || 0;
        return [len & 0xff, first & 0xff, last & 0xff, 0, 0, 0, 0, 0];
      },
      DigestAlgorithm: { SHA_256: 'SHA_256' },
      Charset: { UTF_8: 'UTF_8' },
      newBlob: (str) => ({ getBytes: () => Buffer.from(String(str), 'utf8') })
    }
  };

  vm.createContext(context);
  vm.runInContext(LESSON_SOURCE, context, { filename: 'LessonService.js' });

  return { context, lessonsSheet, formCreations, formCloses, configPatches };
}

// =====================================================================
// createLessonDraft
// =====================================================================

test('createLessonDraft: draft 行を作成 (template = doutoku-3phase)', () => {
  const { context, lessonsSheet } = loadLessonContext();
  const res = context.createLessonDraft('u1', '道徳 5/15', 'doutoku-3phase');
  assert.equal(res.success, true);
  assert.equal(res.data.lesson.state, 'draft');
  assert.equal(res.data.lesson.lessonJson.phases.length, 3);
  // sheet に 1 行追加されている (header + 1 行 = 2 行)
  assert.equal(lessonsSheet._data.length, 2);
});

test('createLessonDraft: 他ユーザーの userId を渡すと reject', () => {
  const { context } = loadLessonContext();
  const res = context.createLessonDraft('u2', 'foo', 'doutoku-3phase');
  assert.equal(res.success, false);
  assert.match(res.message, /他ユーザー/);
});

test('createLessonDraft: 未知のテンプレートは reject', () => {
  const { context } = loadLessonContext();
  const res = context.createLessonDraft('u1', 'foo', 'unknown-template');
  assert.equal(res.success, false);
  assert.match(res.message, /未知のテンプレート/);
});

// =====================================================================
// updateLessonDraft
// =====================================================================

test('updateLessonDraft: name 列と lessonJson 両方を更新', () => {
  const { context } = loadLessonContext();
  const created = context.createLessonDraft('u1', '元の名前', 'doutoku-3phase');
  const lessonId = created.data.lesson.lessonId;

  const res = context.updateLessonDraft('u1', lessonId, 'name', '新しい名前');
  assert.equal(res.success, true);
  assert.equal(res.data.lesson.name, '新しい名前');
});

test('updateLessonDraft: bracket 形式 fieldPath で phases 配列を更新', () => {
  const { context } = loadLessonContext();
  const created = context.createLessonDraft('u1', '5/15', 'doutoku-3phase');
  const lessonId = created.data.lesson.lessonId;

  const res = context.updateLessonDraft('u1', lessonId, 'phases[1].question', '新しい質問');
  assert.equal(res.success, true);
  assert.equal(res.data.lesson.lessonJson.phases[1].question, '新しい質問');
});

test('updateLessonDraft: phases 配列全体を置換できる (add / delete / 編集ができる)', () => {
  const { context } = loadLessonContext();
  const created = context.createLessonDraft('u1', '5/15', 'doutoku-3phase');
  const lessonId = created.data.lesson.lessonId;

  // 5 フェーズに増やして名前 / 形式 / 質問を上書き
  const newPhases = [
    { name: '事前学習', formTemplate: 'board', question: '今日のテーマで知っていることは？' },
    { name: '導入', formTemplate: 'numberline', question: '立場は？' },
    { name: '議論', formTemplate: 'matrix', question: '理由と確信度は？' },
    { name: '深掘り', formTemplate: 'matrix', question: '反対意見をどう受け止める？' },
    { name: '振り返り', formTemplate: 'numberline', question: '今の立場は？' }
  ];
  const res = context.updateLessonDraft('u1', lessonId, 'phases', newPhases);
  assert.equal(res.success, true, JSON.stringify(res));
  assert.equal(res.data.lesson.lessonJson.phases.length, 5);
  assert.equal(res.data.lesson.lessonJson.phases[0].name, '事前学習');
  assert.equal(res.data.lesson.lessonJson.phases[0].formTemplate, 'board');
  assert.equal(res.data.lesson.lessonJson.phases[4].name, '振り返り');
});

test('startLesson: templateOptions が createTemplateForm に渡る (numberline axis label)', () => {
  const { context, formCreations } = loadLessonContext();
  const created = context.createLessonDraft('u1', 'axis lesson', 'doutoku-3phase');
  const lessonId = created.data.lesson.lessonId;
  context.updateLessonDraft('u1', lessonId, 'classes', ['5-1']);
  context.updateLessonDraft('u1', lessonId, 'phases', [
    {
      name: '導入', formTemplate: 'numberline', question: '',
      templateOptions: { lowLabel: '反対', highLabel: '賛成' }
    }
  ]);
  const res = context.startLesson('u1', lessonId);
  assert.equal(res.success, true, JSON.stringify(res));
  assert.equal(formCreations[0].templateType, 'numberline');
  assert.equal(formCreations[0].templateOptions.lowLabel, '反対');
  assert.equal(formCreations[0].templateOptions.highLabel, '賛成');
});

test('startLesson: lessonJson.allowResubmit が全 phase に伝播 + displaySettings にも反映', () => {
  const { context, formCreations, configPatches } = loadLessonContext();
  const created = context.createLessonDraft('u1', '揺らぎ追跡レッスン', 'doutoku-3phase');
  const lessonId = created.data.lesson.lessonId;
  context.updateLessonDraft('u1', lessonId, 'classes', ['5-1']);
  context.updateLessonDraft('u1', lessonId, 'allowResubmit', true);
  context.updateLessonDraft('u1', lessonId, 'phases', [
    { name: '導入', formTemplate: 'numberline', question: 'いまの考えは？',
      templateOptions: { lowLabel: '反対', highLabel: '賛成' } }
  ]);
  const res = context.startLesson('u1', lessonId);
  assert.equal(res.success, true, JSON.stringify(res));
  // createTemplateForm に allowResubmit=true が渡る
  assert.equal(formCreations[0].templateOptions.allowResubmit, true);
  // config.allowResubmit にも反映 (board の ghost dot レンダリング条件)
  assert.equal(configPatches[0].patch.allowResubmit, true);
});

test('startLesson: lessonName / phaseName / question / classChoices が createTemplateForm に流れる (UI ⇄ Form 100% 一致)', () => {
  const { context, formCreations } = loadLessonContext();
  const created = context.createLessonDraft('u1', '5/16 道徳 友だち', 'doutoku-3phase');
  const lessonId = created.data.lesson.lessonId;
  context.updateLessonDraft('u1', lessonId, 'classes', ['5-1', '5-2', '5-3']);
  context.updateLessonDraft('u1', lessonId, 'phases', [
    {
      name: '導入', formTemplate: 'numberline',
      question: 'あなたの立場は？',
      templateOptions: { lowLabel: '反対', highLabel: '賛成' }
    },
    {
      name: '本時', formTemplate: 'matrix',
      question: 'なぜそう考えますか？',
      templateOptions: { xLow: '自分', xHigh: 'みんな', yLow: '今', yHigh: '将来' }
    },
    {
      name: '振り返り', formTemplate: 'pie',
      question: 'いまどう感じる？',
      templateOptions: { choices: ['納得', '迷い', '反対'] }
    }
  ]);
  const res = context.startLesson('u1', lessonId);
  assert.equal(res.success, true, JSON.stringify(res));

  // Phase 0: numberline
  assert.equal(formCreations[0].templateOptions.lessonName, '5/16 道徳 友だち');
  assert.equal(formCreations[0].templateOptions.phaseName, '導入');
  assert.equal(formCreations[0].templateOptions.question, 'あなたの立場は？');
  assert.deepEqual(Array.from(formCreations[0].templateOptions.classChoices), ['5-1', '5-2', '5-3']);
  assert.equal(formCreations[0].templateOptions.lowLabel, '反対');

  // Phase 1: matrix (同じ lesson / classes が全 phase で共有される)
  assert.equal(formCreations[1].templateOptions.lessonName, '5/16 道徳 友だち');
  assert.equal(formCreations[1].templateOptions.phaseName, '本時');
  assert.equal(formCreations[1].templateOptions.question, 'なぜそう考えますか？');
  assert.deepEqual(Array.from(formCreations[1].templateOptions.classChoices), ['5-1', '5-2', '5-3']);
  assert.equal(formCreations[1].templateOptions.xLow, '自分');
  assert.equal(formCreations[1].templateOptions.yHigh, '将来');

  // Phase 2: pie (choices も伝わる)
  assert.equal(formCreations[2].templateType, 'pie');
  assert.equal(formCreations[2].templateOptions.question, 'いまどう感じる？');
  assert.deepEqual(Array.from(formCreations[2].templateOptions.choices), ['納得', '迷い', '反対']);
});

test('startLesson: numberline の templateOptions が config トップレベル xAxisLabels に反映される', () => {
  const { context, configPatches } = loadLessonContext();
  const created = context.createLessonDraft('u1', 'L', 'doutoku-3phase');
  const lessonId = created.data.lesson.lessonId;
  context.updateLessonDraft('u1', lessonId, 'classes', ['5-1']);
  context.updateLessonDraft('u1', lessonId, 'phases', [
    {
      name: '導入', formTemplate: 'numberline', question: '',
      templateOptions: { lowLabel: '反対', highLabel: '賛成' }
    }
  ]);
  const res = context.startLesson('u1', lessonId);
  assert.equal(res.success, true);
  // 軸ラベルは canonical なトップレベルに格納される (旧 nested displaySettings.xAxisLabels は
  //   sanitizeDisplaySettings の allowlist で保存往復ごとに落ちていたため修正済)。
  const patch = configPatches[0].patch;
  assert.equal(patch.displaySettings.xAxisLabels, undefined, 'nested には残さない');
  assert.equal(patch.xAxisLabels.min, '反対');
  assert.equal(patch.xAxisLabels.max, '賛成');
});

test('startLesson: pie テンプレートのフェーズ → displaySettings.boardMode=pie で applyConfigPatch_', () => {
  const { context, configPatches, formCreations } = loadLessonContext();
  const created = context.createLessonDraft('u1', 'pie授業', 'doutoku-3phase');
  const lessonId = created.data.lesson.lessonId;
  // 全 phase を pie に置換
  context.updateLessonDraft('u1', lessonId, 'classes', ['5-1']);
  context.updateLessonDraft('u1', lessonId, 'phases', [
    { name: '導入', formTemplate: 'pie', question: 'A か B か？' }
  ]);

  const res = context.startLesson('u1', lessonId);
  assert.equal(res.success, true, JSON.stringify(res));
  assert.equal(formCreations[0].templateType, 'pie', 'createTemplateForm が pie で呼ばれていない');
  // applyConfigPatch_ の patch.displaySettings.boardMode が 'pie' であること
  assert.equal(configPatches[0].patch.displaySettings.boardMode, 'pie',
    'pie テンプレ phase は displaySettings.boardMode=pie で公開されるべき');
});

test('updateLessonDraft: phases を 1 フェーズに減らせる', () => {
  const { context } = loadLessonContext();
  const created = context.createLessonDraft('u1', '5/15', 'doutoku-3phase');
  const lessonId = created.data.lesson.lessonId;

  const res = context.updateLessonDraft('u1', lessonId, 'phases', [
    { name: '一問だけ', formTemplate: 'numberline', question: 'あなたの立場は？' }
  ]);
  assert.equal(res.success, true);
  assert.equal(res.data.lesson.lessonJson.phases.length, 1);
});

test('updateLessonDraft: state が active なら FORBIDDEN_STATE_MUTATION', () => {
  const ctxBundle = loadLessonContext();
  const created = ctxBundle.context.createLessonDraft('u1', '5/15', 'doutoku-3phase');
  const lessonId = created.data.lesson.lessonId;
  // classes を入れて startLesson
  ctxBundle.context.updateLessonDraft('u1', lessonId, 'classes', ['5-1']);
  const startRes = ctxBundle.context.startLesson('u1', lessonId);
  assert.equal(startRes.success, true, JSON.stringify(startRes));

  const res = ctxBundle.context.updateLessonDraft('u1', lessonId, 'name', '変えたい');
  assert.equal(res.success, false);
  assert.match(res.message, /draft/);
});

// =====================================================================
// listLessons
// =====================================================================

test('listLessons: owner は自分の lesson のみ取得', () => {
  const { context, lessonsSheet } = loadLessonContext();
  context.createLessonDraft('u1', 'mine 1', 'doutoku-3phase');
  context.createLessonDraft('u1', 'mine 2', 'doutoku-3phase');
  // 他ユーザー分を sheet 直接挿入
  lessonsSheet._data.push(['lesson_x', 'u2', 'other', 'draft', '2026-01-01', '', '', 1, 100, 'etag', '{}']);

  const res = context.listLessons('u1');
  assert.equal(res.success, true);
  assert.equal(res.data.lessons.length, 2);
  assert.equal(res.data.lessons.every(l => l.name.startsWith('mine')), true);
});

test('listLessonTemplates: 全テンプレート (4 種) が返る', () => {
  const { context } = loadLessonContext();
  const res = context.listLessonTemplates();
  assert.equal(res.success, true);
  const keys = Array.from(res.data.templates).map(t => t.key).sort();
  assert.deepEqual(keys, ['before-after-2phase', 'doutoku-3phase', 'inquiry-3phase', 'kid-3phase']);
  const kid = Array.from(res.data.templates).find(t => t.key === 'kid-3phase');
  // 2026-05-16: ラベルを「低学年向け」に変更 (旧「児童向け」)。教科ニュートラル化と整合。
  assert.ok(kid && kid.label.includes('低学年'));
  assert.equal(kid.phaseCount, 3);
});

test('createLessonDraft: kid-3phase テンプレで作成すると低学年向けフェーズ名になる', () => {
  const { context } = loadLessonContext();
  const res = context.createLessonDraft('u1', 'k1', 'kid-3phase');
  assert.equal(res.success, true);
  const names = Array.from(res.data.lesson.lessonJson.phases).map(p => p.name);
  // 2026-05-16: 児童向けphase名を「いまの考え→みんなで話す→これからの考え」に変更
  //   (「考えがどう変わったか」が伝わる表現)。
  assert.deepEqual(names, ['いまの考え', 'みんなで話す', 'これからの考え']);
});

test('duplicateLesson: 完了レッスンを複製すると draft 状態 + Form 情報は空 + (コピー) suffix', () => {
  const { context } = loadLessonContext();
  const src = context.createLessonDraft('u1', 'もとの授業', 'doutoku-3phase');
  const sourceId = src.data.lesson.lessonId;
  // 軸ラベルなど詳細設定済を再現
  context.updateLessonDraft('u1', sourceId, 'classes', ['5-1']);
  context.updateLessonDraft('u1', sourceId, 'phases', [
    { name: '導入', formTemplate: 'numberline', question: '立場は？',
      templateOptions: { lowLabel: '反対', highLabel: '賛成' } }
  ]);
  // 完了状態に持っていく
  context.startLesson('u1', sourceId);
  context.endLesson('u1', sourceId);

  const dup = context.duplicateLesson('u1', sourceId, {});
  assert.equal(dup.success, true, JSON.stringify(dup));
  const newL = dup.data.lesson;
  assert.equal(newL.state, 'draft');
  assert.match(newL.name, /もとの授業 \(コピー\)/);
  assert.notEqual(newL.lessonId, sourceId);
  // フェーズの内容は引き継ぐ
  assert.equal(newL.lessonJson.phases[0].name, '導入');
  assert.equal(newL.lessonJson.phases[0].templateOptions.lowLabel, '反対');
  // Form 情報は空 (新規 Form を作るため)
  assert.equal(newL.lessonJson.phases[0].formId, '');
  assert.equal(newL.lessonJson.phases[0].formUrl, '');
  // classes は既定で引き継がない
  assert.equal(Array.from(newL.lessonJson.classes).length, 0);
});

test('duplicateLesson: options.copyClasses=true で classes を引き継ぐ', () => {
  const { context } = loadLessonContext();
  const src = context.createLessonDraft('u1', 'もとの授業', 'doutoku-3phase');
  const sourceId = src.data.lesson.lessonId;
  context.updateLessonDraft('u1', sourceId, 'classes', ['5-1', '5-2']);

  const dup = context.duplicateLesson('u1', sourceId, { copyClasses: true });
  assert.equal(dup.success, true);
  assert.deepEqual(Array.from(dup.data.lesson.lessonJson.classes), ['5-1', '5-2']);
});

test('duplicateLesson: 他ユーザーのレッスンは複製不可', () => {
  const { context, lessonsSheet } = loadLessonContext();
  lessonsSheet._data.push(['lesson_other', 'u2', 'other', 'completed', '2026-01-01', '2026-01-01', '2026-01-01', 1, 100, 'etag', '{}']);
  const dup = context.duplicateLesson('u1', 'lesson_other', {});
  assert.equal(dup.success, false);
});

test('getKnownClassesForUser: 過去レッスンの classes を unique で集約', () => {
  const { context } = loadLessonContext();
  const a = context.createLessonDraft('u1', 'L1', 'doutoku-3phase');
  context.updateLessonDraft('u1', a.data.lesson.lessonId, 'classes', ['5-1', '5-2']);
  const b = context.createLessonDraft('u1', 'L2', 'doutoku-3phase');
  context.updateLessonDraft('u1', b.data.lesson.lessonId, 'classes', ['5-2', '6-3']);

  const res = context.getKnownClassesForUser('u1');
  assert.equal(res.success, true, JSON.stringify(res));
  // Why sort(): test 内で createLessonDraft が同一 ms に呼ばれると createdAt が同値となり、
  //   「新しい順」の決定はできない (実装の sort は stable で挿入順を維持するが、それは
  //   "新しい順" を保証しない)。dedup と set-membership のみ検証。
  //   順序の本番動作は __listLessonsForUser_ の sort に任せる (異なる ms なら正しく動く)。
  const got = Array.from(res.data.classes).slice().sort();
  assert.deepEqual(got, ['5-1', '5-2', '6-3']);
});

test('getKnownClassesForUser: 他ユーザーは reject', () => {
  const { context } = loadLessonContext();
  const res = context.getKnownClassesForUser('other-user');
  assert.equal(res.success, false);
  assert.match(res.message || '', /アクセスできません/);
});

// =====================================================================
// startLesson
// =====================================================================

test('startLesson: 3 phase の Form を生成し state=active', () => {
  const { context, formCreations, configPatches } = loadLessonContext();
  const created = context.createLessonDraft('u1', '5/15', 'doutoku-3phase');
  const lessonId = created.data.lesson.lessonId;
  context.updateLessonDraft('u1', lessonId, 'classes', ['5-1', '5-2']);

  const res = context.startLesson('u1', lessonId);
  assert.equal(res.success, true, JSON.stringify(res));
  assert.equal(res.data.lesson.state, 'active');
  assert.equal(formCreations.length, 3);
  assert.deepEqual(formCreations.map(f => f.templateType), ['numberline', 'matrix', 'numberline']);
  // phase 0 だけ activate されている
  assert.equal(configPatches.length, 1);
  assert.equal(configPatches[0].patch.formUrl, 'https://forms.example/1');
});

test('startLesson: classes 未指定なら reject', () => {
  const { context, formCreations } = loadLessonContext();
  const created = context.createLessonDraft('u1', '5/15', 'doutoku-3phase');
  const lessonId = created.data.lesson.lessonId;

  const res = context.startLesson('u1', lessonId);
  assert.equal(res.success, false);
  assert.match(res.message, /クラス/);
  assert.equal(formCreations.length, 0);
});

test('startLesson: idempotent - 既に formId があるフェーズは skip', () => {
  const ctxBundle = loadLessonContext({
    createTemplateForm: (() => {
      let count = 0;
      return (type) => {
        count++;
        if (count === 2) return { success: false, error: 'simulated phase 2 failure' };
        return { success: true, formId: `form_${count}`, formUrl: `u${count}`, spreadsheetId: `ss${count}`, sheetName: 's' };
      };
    })()
  });
  const created = ctxBundle.context.createLessonDraft('u1', '5/15', 'doutoku-3phase');
  const lessonId = created.data.lesson.lessonId;
  ctxBundle.context.updateLessonDraft('u1', lessonId, 'classes', ['5-1']);

  // 1 回目: phase 2 で失敗 → state は draft のまま、phases[0] のみ formId が入る
  const firstRes = ctxBundle.context.startLesson('u1', lessonId);
  assert.equal(firstRes.success, false);
  assert.match(firstRes.message, /phase 2/);

  // 2 回目: 正常に Form を返すフォームを設定して再試行 → phase 0 は skip される
  ctxBundle.context.createTemplateForm = (type) => ({
    success: true,
    formId: `retry_${type}`, formUrl: 'u', spreadsheetId: 'ss', sheetName: 's'
  });
  const secondRes = ctxBundle.context.startLesson('u1', lessonId);
  assert.equal(secondRes.success, true, JSON.stringify(secondRes));
  // phase 0 (numberline) は 1 回目の form_1 が残っている
  assert.equal(secondRes.data.lesson.lessonJson.phases[0].formId, 'form_1');
});

test('startLesson: LockService が busy なら LESSON_BUSY', () => {
  // Setup phase: default (lock-succeeds) context so createLessonDraft / updateLessonDraft succeed.
  // Then swap LockService to busy stub before startLesson — both __updateLessonRow_ (lesson sheet)
  // and startLesson の lock の両方が busy 扱いになるが、startLesson の lock check が先に動くので
  // LESSON_BUSY が返る。
  const ctxBundle = loadLessonContext();
  const created = ctxBundle.context.createLessonDraft('u1', '5/15', 'doutoku-3phase');
  const lessonId = created.data.lesson.lessonId;
  ctxBundle.context.updateLessonDraft('u1', lessonId, 'classes', ['5-1']);

  ctxBundle.context.LockService = {
    getScriptLock: () => ({
      tryLock: () => false,
      releaseLock: () => {}
    })
  };

  const res = ctxBundle.context.startLesson('u1', lessonId);
  assert.equal(res.success, false);
  assert.match(res.message, /LESSON_BUSY/);
});

// =====================================================================
// advanceLessonPhase
// =====================================================================

test('advanceLessonPhase: next で profileTransitions に entry が追加され Form が切替えられる', () => {
  const { context, formCloses, configPatches } = loadLessonContext();
  const created = context.createLessonDraft('u1', '5/15', 'doutoku-3phase');
  const lessonId = created.data.lesson.lessonId;
  context.updateLessonDraft('u1', lessonId, 'classes', ['5-1']);
  context.startLesson('u1', lessonId);

  const formCloseBefore = formCloses.length;
  const patchesBefore = configPatches.length;
  const res = context.advanceLessonPhase('u1', lessonId, 'next');
  assert.equal(res.success, true);
  assert.equal(res.data.activePhaseIndex, 1);
  // phase 0 close, phase 1 open の 2 つの form 操作が追加
  assert.equal(formCloses.length - formCloseBefore, 2);
  // phase 1 の config が user config に書かれた
  assert.equal(configPatches.length - patchesBefore, 1);
});

test('advanceLessonPhase: 最後のフェーズで next すると reject', () => {
  const { context } = loadLessonContext();
  const created = context.createLessonDraft('u1', '5/15', 'doutoku-3phase');
  const lessonId = created.data.lesson.lessonId;
  context.updateLessonDraft('u1', lessonId, 'classes', ['5-1']);
  context.startLesson('u1', lessonId);
  context.advanceLessonPhase('u1', lessonId, 'next'); // → phase 1
  context.advanceLessonPhase('u1', lessonId, 'next'); // → phase 2 (最後)

  const res = context.advanceLessonPhase('u1', lessonId, 'next');
  assert.equal(res.success, false);
  assert.match(res.message, /最後のフェーズ/);
});

// =====================================================================
// endLesson
// =====================================================================

test('endLesson: state を completed に変更し全 Form を close', () => {
  const { context, formCloses } = loadLessonContext();
  const created = context.createLessonDraft('u1', '5/15', 'doutoku-3phase');
  const lessonId = created.data.lesson.lessonId;
  context.updateLessonDraft('u1', lessonId, 'classes', ['5-1']);
  context.startLesson('u1', lessonId);

  const closeCountBefore = formCloses.length;
  const res = context.endLesson('u1', lessonId);
  assert.equal(res.success, true);
  assert.equal(res.data.lesson.state, 'completed');
  // 3 つの Form が close される
  assert.equal(formCloses.length - closeCountBefore, 3);
  assert.match(res.data.reviewUrl, /lessonId=/);
});

// =====================================================================
// deleteLesson
// =====================================================================

test('deleteLesson: draft は削除可、active は reject', () => {
  const { context } = loadLessonContext();
  const draft = context.createLessonDraft('u1', 'draft1', 'doutoku-3phase');
  const draftId = draft.data.lesson.lessonId;

  const delRes = context.deleteLesson('u1', draftId);
  assert.equal(delRes.success, true);

  const live = context.createLessonDraft('u1', 'live', 'doutoku-3phase');
  const liveId = live.data.lesson.lessonId;
  context.updateLessonDraft('u1', liveId, 'classes', ['5-1']);
  context.startLesson('u1', liveId);
  const denyRes = context.deleteLesson('u1', liveId);
  assert.equal(denyRes.success, false);
  assert.match(denyRes.message, /実行中/);
});

// =====================================================================
// Phase 2: snapshot capture + auto-archive
// =====================================================================

test('endLesson: snapshots に rows を freeze + reactions deep-copy', () => {
  const liveRow = { rowIndex: 2, emailHash: 'h1', answer: '回答1', reactions: { UNDERSTAND: ['h2'], LIKE: [] } };
  const { context } = loadLessonContext({
    getPublishedSheetData: () => ({ success: true, data: [liveRow] })
  });
  const created = context.createLessonDraft('u1', '5/15', 'doutoku-3phase');
  const lessonId = created.data.lesson.lessonId;
  context.updateLessonDraft('u1', lessonId, 'classes', ['5-1']);
  context.startLesson('u1', lessonId);

  const endRes = context.endLesson('u1', lessonId);
  assert.equal(endRes.success, true);

  const snaps = endRes.data.lesson.lessonJson.snapshots;
  assert.equal(snaps.length, 1);
  assert.equal(snaps[0].phaseIndex, 0);
  assert.equal(snaps[0].rowCount, 1);
  assert.equal(snaps[0].rows[0].answer, '回答1');

  // freeze: live row.reactions を後から mutate しても snapshot は影響を受けない
  liveRow.reactions.UNDERSTAND.push('h99');
  // cross-realm Array なので JSON 経由で比較
  assert.equal(JSON.stringify(snaps[0].rows[0].reactions.UNDERSTAND), '["h2"]');
});

test('endLesson: capture 失敗時も lesson は completed に遷移 (reason 付き空 snapshot)', () => {
  const { context } = loadLessonContext({
    getPublishedSheetData: () => ({ success: false, error: 'sheet read failed' })
  });
  const created = context.createLessonDraft('u1', '5/15', 'doutoku-3phase');
  const lessonId = created.data.lesson.lessonId;
  context.updateLessonDraft('u1', lessonId, 'classes', ['5-1']);
  context.startLesson('u1', lessonId);

  const endRes = context.endLesson('u1', lessonId);
  assert.equal(endRes.success, true);
  assert.equal(endRes.data.lesson.state, 'completed');

  const snap = endRes.data.lesson.lessonJson.snapshots[0];
  assert.equal(snap.rows.length, 0);
  assert.match(snap.reason || '', /CAPTURE_FAILED/);
});

test('advanceLessonPhase: 移行前 (fromIdx) の rows を snapshot に保存する', () => {
  let callCount = 0;
  const { context } = loadLessonContext({
    getPublishedSheetData: () => {
      callCount++;
      return { success: true, data: [{ rowIndex: callCount + 1, answer: `phase${callCount} ans` }] };
    }
  });
  const created = context.createLessonDraft('u1', '5/15', 'doutoku-3phase');
  const lessonId = created.data.lesson.lessonId;
  context.updateLessonDraft('u1', lessonId, 'classes', ['5-1']);
  context.startLesson('u1', lessonId);

  const advRes = context.advanceLessonPhase('u1', lessonId, 'next');
  assert.equal(advRes.success, true);

  const snaps = advRes.data.lesson.lessonJson.snapshots;
  assert.equal(snaps.length, 1);
  assert.equal(snaps[0].phaseIndex, 0);
  // phase 0 終了時の row が記録されている
  assert.equal(snaps[0].rows[0].answer, 'phase1 ans');
});

test('advance → end: 同 phaseIndex の double capture は replace (append しない)', () => {
  const { context } = loadLessonContext({
    getPublishedSheetData: () => ({ success: true, data: [{ rowIndex: 2, answer: 'r' }] })
  });
  const created = context.createLessonDraft('u1', '5/15', 'doutoku-3phase');
  const lessonId = created.data.lesson.lessonId;
  context.updateLessonDraft('u1', lessonId, 'classes', ['5-1']);
  context.startLesson('u1', lessonId);
  context.advanceLessonPhase('u1', lessonId, 'next');     // phase 0 snapshot
  context.advanceLessonPhase('u1', lessonId, 'previous'); // phase 1 snapshot (戻る)
  context.advanceLessonPhase('u1', lessonId, 'next');     // phase 0 snapshot (replace されるはず)
  const endRes = context.endLesson('u1', lessonId);

  const snaps = endRes.data.lesson.lessonJson.snapshots;
  // phaseIndex でユニーク
  const indices = snaps.map(s => s.phaseIndex);
  assert.equal(new Set(indices).size, indices.length, 'phaseIndex に重複 = ' + JSON.stringify(indices));
});

test('__maybeAutoArchiveLesson_: 5min 未満は archive しない', () => {
  const { context } = loadLessonContext({
    getPublishedSheetData: () => ({ success: true, data: [{ rowIndex: 2 }] })
  });
  const created = context.createLessonDraft('u1', '5/15', 'doutoku-3phase');
  const lessonId = created.data.lesson.lessonId;
  context.updateLessonDraft('u1', lessonId, 'classes', ['5-1']);
  context.startLesson('u1', lessonId);

  // 1 分前から開始ということにする
  const oneMinAgo = new Date(Date.now() - 60 * 1000).toISOString();
  const res = context.__maybeAutoArchiveLesson_(
    { userId: 'u1', userEmail: 'teacher@example.com' },
    { activeLessonId: lessonId, currentLessonStartedAt: oneMinAgo }
  );
  assert.equal(res.archived, false);
  assert.equal(res.reason, 'too_short');
});

test('__maybeAutoArchiveLesson_: 5min 以上 + 1 回答以上で archive', () => {
  const { context } = loadLessonContext({
    getPublishedSheetData: () => ({ success: true, data: [{ rowIndex: 2, answer: 'r' }] })
  });
  const created = context.createLessonDraft('u1', '5/15', 'doutoku-3phase');
  const lessonId = created.data.lesson.lessonId;
  context.updateLessonDraft('u1', lessonId, 'classes', ['5-1']);
  context.startLesson('u1', lessonId);

  const tenMinAgo = new Date(Date.now() - 10 * 60 * 1000).toISOString();
  const res = context.__maybeAutoArchiveLesson_(
    { userId: 'u1', userEmail: 'teacher@example.com' },
    { activeLessonId: lessonId, currentLessonStartedAt: tenMinAgo }
  );
  assert.equal(res.archived, true);
  assert.equal(res.lessonId, lessonId);
});

test('__maybeAutoArchiveLesson_: 0 回答なら archive しない', () => {
  const { context } = loadLessonContext({
    getPublishedSheetData: () => ({ success: true, data: [] })
  });
  const created = context.createLessonDraft('u1', '5/15', 'doutoku-3phase');
  const lessonId = created.data.lesson.lessonId;
  context.updateLessonDraft('u1', lessonId, 'classes', ['5-1']);
  context.startLesson('u1', lessonId);

  const tenMinAgo = new Date(Date.now() - 10 * 60 * 1000).toISOString();
  const res = context.__maybeAutoArchiveLesson_(
    { userId: 'u1' },
    { activeLessonId: lessonId, currentLessonStartedAt: tenMinAgo }
  );
  assert.equal(res.archived, false);
  assert.equal(res.reason, 'no_responses');
});

test('installLessonTriggers: 既存トリガーがあれば skip (idempotent)', () => {
  let createCount = 0;
  const { context } = loadLessonContext({
    ScriptApp: {
      getProjectTriggers: () => [{ getHandlerFunction: () => 'dailyLessonArchiveSweep' }],
      newTrigger: () => {
        createCount++;
        return { timeBased: () => ({ everyDays: () => ({ atHour: () => ({ create: () => {} }) }) }) };
      }
    }
  });
  context.installLessonTriggers();
  assert.equal(createCount, 0, '既存トリガーがあるのに新規作成された');
});

test('installLessonTriggers: トリガーが無ければ 1 つ作成', () => {
  let createCount = 0;
  let scheduledHour = null;
  const { context } = loadLessonContext({
    ScriptApp: {
      getProjectTriggers: () => [],
      newTrigger: (fn) => ({
        timeBased: () => ({
          everyDays: () => ({
            atHour: (h) => {
              scheduledHour = h;
              return { create: () => { createCount++; } };
            }
          })
        })
      })
    }
  });
  context.installLessonTriggers();
  assert.equal(createCount, 1);
  assert.equal(scheduledHour, 23);
});
