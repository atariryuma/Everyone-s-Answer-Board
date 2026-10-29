/**
 * lesson.backfillReactions: 終わった授業の記録 (lesson_responses) に、元の回答シートの
 * リアクションの数とハイライトの有無を足す。振り返りはそれを読み戻して見せる。
 *
 * Why: 記録は答え・理由・位置・クラス・時刻だけを残してきた (v2738〜)。リアクションのセルは
 *   押した人のメールなので残さないが、数とハイライトの有無は個人を特定しない。
 *   行がずれていたら別の回答に付けないよう、行番号と時刻 (分まで) の両方で照合する。
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { gasResponseStubs } = require('./_helpers.cjs');

const HEADERS = ['lessonId', 'phaseIndex', 'rowIndex', 'timestamp', 'class', 'answer', 'reason', 'numericX', 'numericY',
  'understand', 'like', 'curious', 'highlight'];

function grid(rows) {
  const data = rows.map((r) => r.slice());
  const writes = [];
  const g = {
    data, writes,
    getMaxColumns: () => 26,
    insertColumnsAfter: () => {},
    getDataRange: () => ({ getValues: () => data.map((r) => r.slice()) }),
    shown: {},  // 'row:col' → セルの見えている文字 (無ければ値の文字)
    getRange: (row, col, numRows = 1, numCols = 1) => ({
      getValues: () => Array.from({ length: numRows }, (_, i) =>
        Array.from({ length: numCols }, (_, j) => ((data[row - 1 + i] || [])[col - 1 + j] ?? ''))),
      getDisplayValues() {
        return Array.from({ length: numRows }, (_, i) => Array.from({ length: numCols }, (_, j) => {
          const k = `${row + i}:${col + j}`;
          return k in this.__shown ? this.__shown[k] : String((data[row - 1 + i] || [])[col - 1 + j] ?? '');
        }));
      },
      __shown: null,
      setValues: (vs) => {
        writes.push({ row, col, rows: vs.length });
        vs.forEach((v, i) => { data[row - 1 + i] = data[row - 1 + i] || []; v.forEach((x, j) => { data[row - 1 + i][col - 1 + j] = x; }); });
      }
    })
  };
  const getRange = g.getRange;
  g.getRange = (...a) => { const r = getRange(...a); r.__shown = g.shown; return r; };
  return g;
}

const t = (s) => new Date(s + ':00Z');
function setup({ state = 'completed' } = {}) {
  // 元の回答シート: 2 行目 = A (なるほど 1・いいね 2・ハイライト)、3 行目 = B (なし)、4 行目 = C (いいね 1)
  const source = grid([
    ['タイムスタンプ', 'メールアドレス', '回答', 'UNDERSTAND', 'LIKE', 'CURIOUS', 'HIGHLIGHT'],
    [t('2026-05-14T10:01'), 'a@x', 'A', 'p@x', 'p@x|q@x', '', true],
    [t('2026-05-14T10:02'), 'b@x', 'B', '', '', '', false],
    [t('2026-05-14T10:03'), 'c@x', 'C', '', 'p@x', '', 'FALSE']
  ]);
  // 記録: 2〜4 行目が L1 のフェーズ 0。C は時刻がずれている (行が入れ替わった想定)。5 行目は別の授業。
  const archive = grid([
    HEADERS.slice(0, 9),
    ['L1', 0, 2, '2026-05-14T10:01', '6-1', 'A', 'ra', '', ''],
    ['L1', 0, 3, '2026-05-14T10:02', '6-1', 'B', 'rb', '', ''],
    ['L1', 0, 4, '2026-05-14T10:09', '6-1', 'C', 'rc', '', ''],
    ['L2', 0, 2, '2026-05-14T10:01', '6-1', 'X', 'rx', '', '']
  ]);
  const lesson = {
    lessonId: 'L1', state,
    lessonJson: {
      phases: [{ name: '本時の議論', spreadsheetId: 'src-ss', sheetName: 'フォームの回答 1' }],
      snapshots: [{ phaseIndex: 0, sheet: 'lesson_responses', startRow: 2, rowCount: 4 }]
    }
  };
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    ...gasResponseStubs(),
    logError_: () => {},
    LESSON_RESPONSES_SHEET_HEADERS: HEADERS,
    openDatabase: () => ({ getSheetByName: (n) => (n === 'lesson_responses' ? archive : null) }),
    openSpreadsheet: (id) => (id === 'src-ss' ? { getSheet: (n) => (n === 'フォームの回答 1' ? source : null) } : null),
    // 本物 (formatters.js) と同じく「分まで」の文字列にそろえる
    formatTimestamp: (v) => { const d = v instanceof Date ? v : new Date(String(v) + (String(v).length === 16 ? ':00Z' : '')); return isNaN(d) ? '-' : d.toISOString().slice(0, 16); }
  };
  vm.createContext(context);
  vm.runInContext(fs.readFileSync(path.resolve(__dirname, '../src/LessonService.js'), 'utf8'), context, { filename: 'LessonService.js' });
  context.__requireLessonOwner_ = () => ({ found: { lesson } });
  return { ctx: context, archive, lesson };
}

test('dryRun: 数えるだけで書かない', () => {
  const { ctx, archive } = setup();
  const res = ctx.__backfillArchiveReactions_('u1', 'L1', { dryRun: true });
  assert.equal(res.success, true, JSON.stringify(res));
  const p = res.data.phases[0];
  assert.deepEqual({ rows: p.rows, matched: p.matched, unmatched: p.unmatched, withReactions: p.withReactions, highlighted: p.highlighted },
    { rows: 3, matched: 2, unmatched: 1, withReactions: 1, highlighted: 1 });
  assert.deepEqual(JSON.parse(JSON.stringify(p.totals)), { UNDERSTAND: 1, LIKE: 2, CURIOUS: 0 });
  assert.equal(archive.writes.length, 0);
});

test('行番号と時刻が一致する行にだけ足し、ずれた行・別の授業の行は触らない', () => {
  const { ctx, archive } = setup();
  const res = ctx.__backfillArchiveReactions_('u1', 'L1', {});
  assert.equal(res.success, true, JSON.stringify(res));
  assert.deepEqual(archive.data[0].slice(9), ['understand', 'like', 'curious', 'highlight'], '見出しを足す');
  assert.deepEqual(archive.data[1].slice(9, 13), [1, 2, '', true], 'A: なるほど 1・いいね 2・ハイライト');
  assert.deepEqual(archive.data[2].slice(9, 13), ['', '', '', ''], 'B: なし');
  assert.deepEqual(archive.data[3].slice(9, 13), ['', '', '', ''], 'C: 時刻がずれているので足さない (元のシートでは いいね 1)');
  assert.deepEqual(archive.data[4].slice(9, 13), ['', '', '', ''], '別の授業の行は触らない');
});

test('振り返りは足した数とハイライトを読み戻す (無い行には項目を付けない)', () => {
  const { ctx, lesson } = setup();
  ctx.__backfillArchiveReactions_('u1', 'L1', {});
  ctx.__hydrateLessonSnapshots_(lesson);
  const rows = lesson.lessonJson.snapshots[0].rows;
  assert.equal(rows.length, 3);
  assert.deepEqual(JSON.parse(JSON.stringify(rows[0].reactions)), { UNDERSTAND: 1, LIKE: 2, CURIOUS: 0 });
  assert.equal(rows[0].highlight, true);
  assert.equal(rows[1].reactions, undefined);
  assert.equal(rows[1].highlight, undefined);
});

test('授業中の授業には足さない', () => {
  const { ctx, archive } = setup({ state: 'active' });
  const res = ctx.__backfillArchiveReactions_('u1', 'L1', {});
  assert.equal(res.success, false);
  assert.equal(archive.writes.length, 0);
});

test('時刻はセルの見えている文字で読む (シートが日付に読み替えてタイムゾーンでずれていても照合でき、振り返りもずれない)', () => {
  const { ctx, archive, lesson } = setup();
  // 本番で起きていた形: 文字で書いた時刻をシートが日付に読み替え、値で読むと 9 時間ずれる
  archive.data[1][3] = new Date('2026-05-14T19:01:00Z');
  archive.shown['2:4'] = '2026-05-14T10:01';
  const res = ctx.__backfillArchiveReactions_('u1', 'L1', { dryRun: true });
  assert.equal(res.data.phases[0].matched, 2, JSON.stringify(res.data.phases[0]));
  ctx.__hydrateLessonSnapshots_(lesson);
  assert.equal(lesson.lessonJson.snapshots[0].rows[0].timestamp, '2026-05-14T10:01', '振り返りの時刻は見えている文字');
});
