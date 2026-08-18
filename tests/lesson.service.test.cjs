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
  const RESPONSES_HEADERS = ['lessonId', 'phaseIndex', 'rowIndex', 'timestamp', 'class', 'answer', 'reason', 'numericX', 'numericY'];
  const lessonsSheet = overrides.lessonsSheet || createFakeLessonsSheet(LESSONS_HEADERS);
  // アーカイブ側。appendRows (SA proxy 専用) は持たせず getLastRow + setValues の
  //   fallback 経路を通す = native Sheet 相当の挙動でテストする。
  const responsesSheet = overrides.responsesSheet || createFakeLessonsSheet(RESPONSES_HEADERS);
  const formCreations = [];
  const formCloses = [];
  const configPatches = [];

  let uuidCounter = 0;
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    ...gasResponseStubs(),
    openDatabase: () => ({
      getSheetByName: (name) => name === 'lessons' ? lessonsSheet
        : name === 'lesson_responses' ? responsesSheet : null,
      // Production の sheet 不存在チェック (SA proxy) に合わせて getSheets() も提供。
      getSheets: () => [{ getName: () => 'lessons' }, { getName: () => 'lesson_responses' }]
    }),
    // SpreadsheetApp.openById で direct admin write (deleteLesson 用)。
    SpreadsheetApp: {
      openById: () => ({ getSheetByName: (name) => name === 'lessons' ? lessonsSheet : null })
    },
    LESSONS_SHEET_HEADERS: ['lessonId', 'userId', 'name', 'state', 'createdAt', 'startedAt', 'endedAt', 'schemaVersion', 'sizeBytes', 'etag', 'lessonJson'],
    LESSON_RESPONSES_SHEET_HEADERS: RESPONSES_HEADERS,
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

  return { context, lessonsSheet, responsesSheet, formCreations, formCloses, configPatches };
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

test('endLesson: rows は lesson_responses へ書き、snapshot はポインタになる', () => {
  const liveRow = {
    rowIndex: 2, name: '児童A', email: 'kid@example.com', emailHash: 'h1',
    answer: '回答1', reason: '理由1', class: '5-1', numericX: 4, numericY: 5,
    reactions: { UNDERSTAND: ['h2'], LIKE: [] }
  };
  const { context, responsesSheet } = loadLessonContext({
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
  // snapshot 本体は rows を持たない (ポインタのみ)
  assert.equal(snaps[0].rows.length, 0);
  assert.equal(snaps[0].sheet, 'lesson_responses');
  assert.equal(snaps[0].startRow, 2);
  assert.equal(snaps[0].rowCount, 1);

  // アーカイブ行は 9 列に射影され、PII (name/email/reactions) は書かれない
  const archived = responsesSheet._data[1];
  assert.equal(archived[0], lessonId);
  assert.equal(archived[5], '回答1');
  assert.equal(archived[6], '理由1');
  assert.equal(JSON.stringify(archived).includes('児童A'), false);
  assert.equal(JSON.stringify(archived).includes('kid@example.com'), false);

  // getLessonForReview はポインタから rows を読み戻す (クライアント互換の形)
  const review = context.getLessonForReview('u1', lessonId);
  assert.equal(review.success, true);
  const hydrated = review.data.lesson.lessonJson.snapshots[0];
  assert.equal(hydrated.rows.length, 1);
  assert.equal(hydrated.rows[0].answer, '回答1');
  assert.equal(hydrated.rows[0].numericX, 4);
  assert.equal(hydrated.rows[0].numericY, 5);
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
  // phase 0 終了時の row がアーカイブに記録され、hydrate で読み戻せる
  const review = context.getLessonForReview('u1', lessonId);
  assert.equal(review.data.lesson.lessonJson.snapshots[0].rows[0].answer, 'phase1 ans');
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

// ── targetIndex による直接ジャンプ (回答ボード上の phase pill 用) ──
// Why: pill は全 phase を並べるので隣以外も押せる。±1 しか無いと死んだ UI になる。

test('advanceLessonPhase: targetIndex で任意フェーズへ直接ジャンプできる', () => {
  const { context } = loadLessonContext();
  const created = context.createLessonDraft('u1', '5/15', 'doutoku-3phase');
  const lessonId = created.data.lesson.lessonId;
  context.updateLessonDraft('u1', lessonId, 'classes', ['5-1']);
  context.startLesson('u1', lessonId);

  // phase 0 から 2 へ、間を経由せず 1 回で飛ぶ
  const res = context.advanceLessonPhase('u1', lessonId, null, 2);
  assert.equal(res.success, true);
  assert.equal(res.data.activePhaseIndex, 2);
});

test('advanceLessonPhase: targetIndex が現在と同じなら reject', () => {
  const { context } = loadLessonContext();
  const created = context.createLessonDraft('u1', '5/15', 'doutoku-3phase');
  const lessonId = created.data.lesson.lessonId;
  context.updateLessonDraft('u1', lessonId, 'classes', ['5-1']);
  context.startLesson('u1', lessonId);

  const res = context.advanceLessonPhase('u1', lessonId, null, 0);
  assert.equal(res.success, false);
  assert.match(res.message, /既にそのフェーズ/);
});

test('advanceLessonPhase: 範囲外の targetIndex は reject', () => {
  const { context } = loadLessonContext();
  const created = context.createLessonDraft('u1', '5/15', 'doutoku-3phase');
  const lessonId = created.data.lesson.lessonId;
  context.updateLessonDraft('u1', lessonId, 'classes', ['5-1']);
  context.startLesson('u1', lessonId);

  assert.equal(context.advanceLessonPhase('u1', lessonId, null, 99).success, false);
  assert.equal(context.advanceLessonPhase('u1', lessonId, null, -1).success, false);
  // 整数でない値も弾く (client から文字列が来ても壊れない)
  assert.equal(context.advanceLessonPhase('u1', lessonId, null, 'abc').success, false);
});

test('advanceLessonPhase: targetIndex 指定でも移行前フェーズの snapshot を残す', () => {
  const { context } = loadLessonContext();
  const created = context.createLessonDraft('u1', '5/15', 'doutoku-3phase');
  const lessonId = created.data.lesson.lessonId;
  context.updateLessonDraft('u1', lessonId, 'classes', ['5-1']);
  context.startLesson('u1', lessonId);

  const res = context.advanceLessonPhase('u1', lessonId, null, 2);
  const snaps = res.data.lesson.lessonJson.snapshots || [];
  // 出発点 (phase 0) が焼かれていること。飛び先の 2 ではない。
  assert.ok(snaps.some(s => s.phaseIndex === 0));
});

test('getActiveLessonNav: 実行中の授業が無ければ data を付けない (正常系)', () => {
  const { context } = loadLessonContext();
  const res = context.getActiveLessonNav('u1');
  // createSuccessResponse の規約: data が null なら data フィールド自体を出さない。
  //   client 側は falsy 判定なので、これで「授業なし = pill を出さない」になる。
  assert.equal(res.success, true);
  assert.ok(!res.data);
});

test('getActiveLessonNav: 実行中の授業の phase 一覧と現在位置を返す', () => {
  const { context } = loadLessonContext();
  const created = context.createLessonDraft('u1', '5/15', 'doutoku-3phase');
  const lessonId = created.data.lesson.lessonId;
  context.updateLessonDraft('u1', lessonId, 'classes', ['5-1']);
  context.startLesson('u1', lessonId);
  context.advanceLessonPhase('u1', lessonId, 'next');

  const res = context.getActiveLessonNav('u1');
  assert.equal(res.success, true);
  assert.equal(res.data.lessonId, lessonId);
  assert.equal(res.data.activePhaseIndex, 1);
  assert.equal(res.data.phases.length, 3);
  assert.equal(res.data.phases[0].index, 0);
});

// ── アーカイブ分離 (lesson_responses) の周辺仕様 ──

test('migrateLessonArchive: 旧形式 (rows 同居) をポインタ + アーカイブ行へ移す', () => {
  const { context, lessonsSheet, responsesSheet } = loadLessonContext();
  // 旧形式の lesson を直接シートに置く (v2894 移行期のデータを再現)
  const legacyJson = {
    phases: [{ name: 'p0', formTemplate: 'matrix' }],
    snapshots: [{
      phaseIndex: 0, phaseName: 'p0', boardMode: 'matrix',
      columnMapping: {}, displaySettings: {},
      rows: [
        { rowIndex: 2, timestamp: 't1', class: '1組', answer: 'こたえ', reason: 'りゆう', numericX: 3, numericY: 4 },
        { rowIndex: 3, timestamp: 't2', class: '2組', answer: 5, reason: '数値回答', numericX: 5, numericY: 5 }
      ],
      rowCount: 2, truncated: true
    }]
  };
  lessonsSheet.appendRow(['lesson_legacy1', 'u1', '旧授業', 'completed', 't', 't', 't', 1,
    JSON.stringify(legacyJson).length, 'etag0', JSON.stringify(legacyJson)]);

  const res = context.migrateLessonArchive('u1', 'lesson_legacy1');
  assert.equal(res.success, true);
  assert.equal(res.data.migrated.length, 1);
  assert.equal(res.data.migrated[0].rowCount, 2);

  // アーカイブ行が書かれ、lessonJson からは rows が消えた
  assert.equal(responsesSheet._data.length, 3); // header + 2
  const saved = JSON.parse(lessonsSheet._data[1][10]);
  assert.equal(saved.snapshots[0].rows.length, 0);
  assert.equal(saved.snapshots[0].sheet, 'lesson_responses');
  assert.equal(saved.snapshots[0].startRow, 2);

  // hydrate で元どおり読める (answer の数値も保持)
  const review = context.getLessonForReview('u1', 'lesson_legacy1');
  const rows = review.data.lesson.lessonJson.snapshots[0].rows;
  assert.equal(rows.length, 2);
  assert.equal(rows[0].answer, 'こたえ');
  assert.equal(rows[1].answer, 5);
  assert.equal(rows[1].numericY, 5);
});

test('migrateLessonArchive: 二重実行は no-op (rows が空なので移行対象なし)', () => {
  const { context, lessonsSheet, responsesSheet } = loadLessonContext();
  const legacyJson = {
    phases: [{ name: 'p0', formTemplate: 'board' }],
    snapshots: [{ phaseIndex: 0, phaseName: 'p0', rows: [
      { rowIndex: 2, timestamp: 't', class: '', answer: 'a', reason: '' }
    ], rowCount: 1 }]
  };
  lessonsSheet.appendRow(['lesson_legacy2', 'u1', '旧', 'completed', 't', 't', 't', 1, 1, 'e', JSON.stringify(legacyJson)]);
  assert.equal(context.migrateLessonArchive('u1', 'lesson_legacy2').success, true);
  const rowsAfterFirst = responsesSheet._data.length;
  const second = context.migrateLessonArchive('u1', 'lesson_legacy2');
  assert.equal(second.success, true);
  assert.match(second.message, /移行対象なし/);
  assert.equal(responsesSheet._data.length, rowsAfterFirst); // 重複追記しない
});

test('hydrate: ポインタがずれて他授業の行を指しても照合ガードで混入しない', () => {
  const { context, lessonsSheet, responsesSheet } = loadLessonContext();
  // 別授業の行がアーカイブに先に存在する
  responsesSheet.appendRow(['lesson_other', 0, 2, 't', '1組', '他授業の回答', '', '', '']);
  // ポインタが誤ってその行 (row 2) を指している lesson
  const json = {
    phases: [{ name: 'p0', formTemplate: 'board' }],
    snapshots: [{ phaseIndex: 0, phaseName: 'p0', rows: [],
      sheet: 'lesson_responses', startRow: 2, rowCount: 1 }]
  };
  lessonsSheet.appendRow(['lesson_drift', 'u1', 'ずれ', 'completed', 't', 't', 't', 1, 1, 'e', JSON.stringify(json)]);

  const review = context.getLessonForReview('u1', 'lesson_drift');
  const sn = review.data.lesson.lessonJson.snapshots[0];
  assert.equal(sn.rows.length, 0);                    // 他授業の回答は返さない
  assert.match(sn.reason || '', /ARCHIVE_POINTER_DRIFT/);
});

test('recaptureLessonArchive: 現在の公開ボードから全文で焼き直す', () => {
  const { context, responsesSheet } = loadLessonContext({
    getPublishedSheetData: () => ({ success: true, data: [
      { rowIndex: 2, class: '1組', answer: '全文の回答', reason: '全文の理由', numericX: 2, numericY: 3 }
    ] })
  });
  const created = context.createLessonDraft('u1', '5/15', 'doutoku-3phase');
  const lessonId = created.data.lesson.lessonId;
  context.updateLessonDraft('u1', lessonId, 'classes', ['5-1']);
  context.startLesson('u1', lessonId);
  context.endLesson('u1', lessonId); // phase 0 が 1 度 capture される

  const before = responsesSheet._data.length;
  const res = context.recaptureLessonArchive('u1', lessonId, 0);
  assert.equal(res.success, true);
  // 旧範囲は孤児として残り (無害)、ポインタは新範囲を指す
  assert.equal(responsesSheet._data.length, before + 1);
  const review = context.getLessonForReview('u1', lessonId);
  const sn = review.data.lesson.lessonJson.snapshots[0];
  assert.equal(sn.startRow, before + 1);
  assert.equal(sn.rows[0].answer, '全文の回答');

  assert.equal(context.recaptureLessonArchive('u1', lessonId, 99).success, false); // 範囲外
});

test('capture: lesson_responses が無い環境では ARCHIVE_WRITE_FAILED で劣化し授業は止めない', () => {
  const { context } = loadLessonContext({
    getPublishedSheetData: () => ({ success: true, data: [{ rowIndex: 2, answer: 'a' }] }),
    // responses シートを見えなくする
    responsesSheet: null
  });
  // openDatabase を lessons のみに差し替え
  context.openDatabase = () => ({
    getSheetByName: (name) => null,
    getSheets: () => [{ getName: () => 'lessons' }]
  });
  // ここでは __writeArchiveRows_ を直接検証 (シート無し + createIfMissing 失敗)
  const ptr = context.__writeArchiveRows_('lesson_x', 0, [{ rowIndex: 2, answer: 'a' }]);
  assert.equal(ptr, null);
});

// ── 授業の再開 (completed → active) と snapshot 上書きガード ──

test('reopenLesson: completed → active、終了時のフェーズで Form 受付と config が復帰する', () => {
  const { context, formCloses, configPatches } = loadLessonContext({
    getPublishedSheetData: () => ({ success: true, data: [{ rowIndex: 2, answer: 'a' }] })
  });
  const created = context.createLessonDraft('u1', '5/15', 'doutoku-3phase');
  const lessonId = created.data.lesson.lessonId;
  context.updateLessonDraft('u1', lessonId, 'classes', ['5-1']);
  context.startLesson('u1', lessonId);
  context.advanceLessonPhase('u1', lessonId, 'next');   // phase 1 が active
  context.endLesson('u1', lessonId);

  const res = context.reopenLesson('u1', lessonId);
  assert.equal(res.success, true);
  assert.equal(res.data.lesson.state, 'active');
  assert.equal(res.data.lesson.endedAt, '');
  assert.equal(res.data.activePhaseIndex, 1);           // 終了時の phase に戻る

  // 再開 phase の Form だけ受付再開 (最後の open が phase 1 の form)
  const lastOpen = formCloses.filter(f => f.accepting === true).pop();
  assert.equal(lastOpen.formId, 'form_2');

  // config は再開 phase + auto-archive marker
  const lastPatch = configPatches[configPatches.length - 1].patch;
  assert.equal(lastPatch.activeLessonId, lessonId);
  assert.equal(typeof lastPatch.currentLessonStartedAt, 'string');

  // 再開後は getActiveLessonNav にも現れる (board pill の供給元)
  const nav = context.getActiveLessonNav('u1');
  assert.equal(nav.data.lessonId, lessonId);
  assert.equal(nav.data.activePhaseIndex, 1);
});

test('reopenLesson: draft は再開できない / active は冪等', () => {
  const { context } = loadLessonContext();
  const created = context.createLessonDraft('u1', '5/15', 'doutoku-3phase');
  const lessonId = created.data.lesson.lessonId;
  assert.equal(context.reopenLesson('u1', lessonId).success, false); // draft
  context.updateLessonDraft('u1', lessonId, 'classes', ['5-1']);
  context.startLesson('u1', lessonId);
  const res = context.reopenLesson('u1', lessonId);                  // active
  assert.equal(res.success, true);
  assert.match(res.message, /already active/);
});

test('__upsertSnapshot_: 失敗 capture はデータを持つ既存 snapshot を上書きしない', () => {
  const { context } = loadLessonContext();
  const lessonJson = { snapshots: [
    { phaseIndex: 0, phaseName: 'p0', rows: [], rowCount: 96,
      sheet: 'lesson_responses', startRow: 2 }
  ] };
  // 失敗 capture → 保持される
  context.__upsertSnapshot_(lessonJson, {
    phaseIndex: 0, rows: [], rowCount: 0, reason: 'CAPTURE_FAILED:read error'
  });
  assert.equal(lessonJson.snapshots[0].rowCount, 96);
  assert.equal(lessonJson.snapshots[0].startRow, 2);
  // 正常 capture → 置き換わる
  context.__upsertSnapshot_(lessonJson, {
    phaseIndex: 0, rows: [], rowCount: 97, sheet: 'lesson_responses', startRow: 300
  });
  assert.equal(lessonJson.snapshots[0].startRow, 300);
});

test('reopen → advance: 再開後の切替で snapshot が通常どおり焼き直される', () => {
  const { context } = loadLessonContext({
    getPublishedSheetData: () => ({ success: true, data: [{ rowIndex: 2, answer: '再開後の回答' }] })
  });
  const created = context.createLessonDraft('u1', '5/15', 'doutoku-3phase');
  const lessonId = created.data.lesson.lessonId;
  context.updateLessonDraft('u1', lessonId, 'classes', ['5-1']);
  context.startLesson('u1', lessonId);
  context.endLesson('u1', lessonId);      // phase 0 capture (startRow 2)
  context.reopenLesson('u1', lessonId);
  const adv = context.advanceLessonPhase('u1', lessonId, 'next');  // phase 0 を焼き直し
  assert.equal(adv.success, true);
  const sn = adv.data.lesson.lessonJson.snapshots.find(s => s.phaseIndex === 0);
  assert.equal(sn.startRow, 3);           // 新しい範囲を指す (孤児行 1 行は無害)
});

// ── phase 並べ替え (reorderLessonPhases) ──

test('reorderLessonPhases: 配列・遷移・snapshot・アーカイブ行の phaseIndex が揃って入れ替わる', () => {
  const { context, lessonsSheet, responsesSheet } = loadLessonContext();
  // profiles 取り込み由来の「本時が先頭」lesson を再現 (実授業順は 導入(1)→本時(0))
  responsesSheet.appendRow(['lesson_ro1', 0, 2, 't', '1組', '本時の回答', '', 4, 5]);
  responsesSheet.appendRow(['lesson_ro1', 1, 2, 't', '1組', '導入の回答', '', '', '']);
  const json = {
    phases: [
      { name: '本時の議論', formTemplate: 'matrix' },
      { name: '導入アンケート', formTemplate: 'pie' }
    ],
    profileTransitions: [
      { ts: 't1', from: null, to: 1 },
      { ts: 't2', from: 1, to: 0 }
    ],
    snapshots: [
      { phaseIndex: 0, phaseName: '本時の議論', rows: [], sheet: 'lesson_responses', startRow: 2, rowCount: 1 },
      { phaseIndex: 1, phaseName: '導入アンケート', rows: [], sheet: 'lesson_responses', startRow: 3, rowCount: 1 }
    ],
    meta: { profileNames: ['本時の議論', '導入アンケート'] }
  };
  lessonsSheet.appendRow(['lesson_ro1', 'u1', '道徳', 'active', 't', 't', '', 1, 1, 'e', JSON.stringify(json)]);

  const res = context.reorderLessonPhases('u1', 'lesson_ro1', [1, 0]);
  assert.equal(res.success, true);
  assert.equal(res.data.phases[0], '0:導入アンケート');
  assert.equal(res.data.phases[1], '1:本時の議論');
  // 遷移も remap: 授業は「導入(新0) から始まり 本時(新1) へ」
  const saved = JSON.parse(lessonsSheet._data[1][10]);
  assert.equal(saved.profileTransitions[0].to, 0);
  assert.equal(saved.profileTransitions[1].to, 1);
  assert.equal(res.data.activePhaseIndex, 1); // 最終遷移 = 本時 (新 index 1)

  // hydrate で正しい回答が正しい phase に付く (アーカイブ行は新 phaseIndex で焼き直し)
  const review = context.getLessonForReview('u1', 'lesson_ro1');
  const snaps = review.data.lesson.lessonJson.snapshots;
  assert.equal(snaps[0].phaseName, '導入アンケート');
  assert.equal(snaps[0].rows[0].answer, '導入の回答');
  assert.equal(snaps[1].phaseName, '本時の議論');
  assert.equal(snaps[1].rows[0].answer, '本時の回答');
});

test('reorderLessonPhases: 順列でない order は reject / 恒等順は no-op', () => {
  const { context, lessonsSheet } = loadLessonContext();
  const json = { phases: [{ name: 'a', formTemplate: 'board' }, { name: 'b', formTemplate: 'board' }], snapshots: [] };
  lessonsSheet.appendRow(['lesson_ro2', 'u1', 'x', 'completed', 't', 't', 't', 1, 1, 'e', JSON.stringify(json)]);
  assert.equal(context.reorderLessonPhases('u1', 'lesson_ro2', [0, 0]).success, false);
  assert.equal(context.reorderLessonPhases('u1', 'lesson_ro2', [0]).success, false);
  const noop = context.reorderLessonPhases('u1', 'lesson_ro2', [0, 1]);
  assert.equal(noop.success, true);
  assert.equal(noop.data.changed, false);
});
