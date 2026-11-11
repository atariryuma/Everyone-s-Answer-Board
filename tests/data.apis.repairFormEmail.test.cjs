/**
 * repairFormEmailCollection_ のテスト (2026-11-11)
 *
 * Why: 2026-10-29 〜 2026-11-11 にアプリが作ったフォームは、メールが「回答者の入力 (自己申告)」に
 *      落ちていた。すでに作ったフォームを確認済みに直す一度きりの手当てが、
 *      ・アプリにつながっているフォームだけを触る
 *      ・自己申告のものだけを変える (集めない設定・確認済みは変えない)
 *      ・apply を渡さなければ何も変えない
 *      ・題名や ID の全体を返さない
 *      ことを守る。
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { gasResponseStubs } = require('./_helpers.cjs');

const DATA_APIS_SCRIPT = new vm.Script(
  fs.readFileSync(path.resolve(__dirname, '../src/DataApis.js'), 'utf8'), { filename: 'DataApis.js' });

const TYPES = { DO_NOT_COLLECT: 'DO_NOT_COLLECT', VERIFIED: 'VERIFIED', RESPONDER_INPUT: 'RESPONDER_INPUT' };

function fakeForm(id, type, opts = {}) {
  const f = {
    _type: type,
    _sets: 0,
    getId: () => id,
    getTitle: () => '【ダミー】1年1組 やまだ たろう',
    getEmailCollectionType: () => { if (opts.unreadable) throw new Error('読めない'); return f._type; },
    setEmailCollectionType: (t) => {
      f._sets++;
      if (opts.throwsOnSet) throw new Error('書けない');
      if (!opts.ignoresSet) f._type = t;
      return f;
    }
  };
  return f;
}

/**
 * @param {object} world
 *   forms: { [編集用のID]: fakeForm }  / bySheet: { [スプレッドシートID]: 編集用のID }
 *   users: [{configJson}] / lessons: [{lessonJson}]
 */
function loadCtx(world, overrides = {}) {
  const logged = [];
  const ctx = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    URL,
    ...gasResponseStubs(),
    logError_: (where) => { logged.push(where); },
    safeJsonParse_: (text, fallback) => { try { return JSON.parse(text); } catch (_) { return fallback; } },
    getAllUsers_: () => world.users || [],
    __listLessonsForUser_: (_userId, options) => {
      assert.equal(options && options.allUsers, true, '全員の授業を見る');
      return world.lessons || [];
    },
    // LessonService.js の __openLessonForm_ と同じ開き方 (編集用の ID → だめなら回答のスプレッドシートから)
    __openLessonForm_: (phase) => {
      if (phase.formId && world.forms[phase.formId]) return world.forms[phase.formId];
      const viaSheet = phase.spreadsheetId && (world.bySheet || {})[phase.spreadsheetId];
      return viaSheet ? world.forms[viaSheet] : null;
    },
    FormApp: {
      EmailCollectionType: TYPES,
      openByUrl: (url) => {
        const m = /\/d\/([^/]+)\/edit/.exec(url);
        if (!m || !world.forms[m[1]]) throw new Error('開けない');
        return world.forms[m[1]];
      }
    },
    // world.sheets: { [スプレッドシートID]: つながっているフォームの編集用 URL ('' = フォームなし) }。無い ID は開けない
    SpreadsheetApp: {
      openById: (id) => {
        const sheets = world.sheets || {};
        if (!(id in sheets)) throw new Error('開けない');
        return { getFormUrl: () => sheets[id] || null };
      }
    },
    __logged: logged
  };
  Object.assign(ctx, overrides);
  vm.createContext(ctx);
  DATA_APIS_SCRIPT.runInContext(ctx);
  return ctx;
}

const user = (config) => ({ userId: 'u1', configJson: JSON.stringify(config) });
const lesson = (phases) => ({ lessonId: 'l1', lessonJson: { phases } });
const plain = (x) => JSON.parse(JSON.stringify(x));

function sampleWorld() {
  const forms = {
    selfA1234567: fakeForm('selfA1234567', TYPES.RESPONDER_INPUT),   // いまのボード (自己申告)
    verB12345678: fakeForm('verB12345678', TYPES.VERIFIED),          // profile (確認済み)
    selfC1234567: fakeForm('selfC1234567', TYPES.RESPONDER_INPUT),   // 授業のフェーズ (自己申告)
    noneD1234567: fakeForm('noneD1234567', TYPES.DO_NOT_COLLECT),    // 授業のフェーズ (集めない)
    strayE123456: fakeForm('strayE123456', TYPES.RESPONDER_INPUT)    // アプリが指していない、先生のほかのフォーム
  };
  return {
    forms,
    bySheet: { 'ss-board': 'selfA1234567' },
    users: [user({
      formUrl: 'https://docs.google.com/forms/d/e/1FAIpQLS-public/viewform',   // 公開側の URL (開けない) → スプレッドシートから開く
      spreadsheetId: 'ss-board',
      profiles: [
        { name: 'p1', formUrl: 'https://docs.google.com/forms/d/verB12345678/edit', spreadsheetId: 'ss-none' },
        { name: 'p2', formUrl: 'https://forms.gle/xxxx', spreadsheetId: 'ss-gone' }   // どこからも開けない
      ]
    })],
    lessons: [lesson([
      { name: 'かんがえる', formId: 'selfC1234567', spreadsheetId: 'ss-c' },
      { name: 'ふりかえる', formId: 'noneD1234567', spreadsheetId: 'ss-d' },
      { name: 'おなじフォーム', formId: 'selfA1234567', spreadsheetId: 'ss-board' },   // いまのボードと同じ → 1 回だけ数える
      { name: '授業モード', formId: '', formUrl: '', spreadsheetId: '' }                // フォームなし → 数えない
    ])]
  };
}

test('repairFormEmailCollection_: apply なしは数えるだけで、何も変えない', () => {
  const world = sampleWorld();
  const ctx = loadCtx(world);
  const res = ctx.repairFormEmailCollection_({});
  assert.equal(res.success, true);
  assert.equal(res.data.applied, false);
  assert.deepEqual(plain(res.data.counts), { verified: 1, selfReported: 2, notCollecting: 1, fixed: 0, failed: 0, unreadable: 0 });
  assert.equal(res.data.forms, 4, '同じフォームは 1 回だけ');
  assert.equal(res.data.unopened, 1, '開けなかった指し先は数える');
  Object.values(world.forms).forEach((f) => assert.equal(f._sets, 0, '何も書かない'));
  // apply: true 以外 (文字の 'true' など) も数えるだけ
  assert.equal(ctx.repairFormEmailCollection_({ apply: 'true' }).data.applied, false);
  assert.equal(ctx.repairFormEmailCollection_().data.applied, false);
  Object.values(world.forms).forEach((f) => assert.equal(f._sets, 0));
});

test('repairFormEmailCollection_: 開けなかった指し先を理由で分ける (フォームが無い・シートを開けない・フォームを開けない)', () => {
  const world = {
    forms: { okA123456789: fakeForm('okA123456789', TYPES.VERIFIED) },
    sheets: { 'ss-native': '', 'ss-native2': '', 'ss-linked-lost': 'https://docs.google.com/forms/d/goneZ/edit' },
    users: [user({
      spreadsheetId: 'ss-deleted',                                              // シートを開けない
      profiles: [{ name: 'p', formUrl: 'https://forms.gle/xxxx' }]              // 公開側の URL だけ (たどれない)
    })],
    lessons: [lesson([
      { formId: 'okA123456789', spreadsheetId: 'ss-x' },                        // 開ける
      { formId: '', spreadsheetId: 'ss-native', sheetName: 'phase1' },          // 授業モード (フォームが無い)
      { formId: '', spreadsheetId: 'ss-native2', sheetName: 'phase2' },
      { formId: 'goneZ', spreadsheetId: 'ss-linked-lost' }                      // シートはフォームを指すが、開けない
    ])]
  };
  const res = loadCtx(world).repairFormEmailCollection_({});
  assert.equal(res.success, true);
  assert.equal(res.data.forms, 1);
  assert.equal(res.data.unopened, 5);
  assert.deepEqual(plain(res.data.unopenedBy), {
    'config/sheetUnopenable': 1,
    'profile/formUnopenable': 1,
    'lesson/noForm': 2,
    'lesson/formUnopenable': 1
  });
});

test('repairFormEmailCollection_: apply で、自己申告のフォームだけを確認済みにする', () => {
  const world = sampleWorld();
  const ctx = loadCtx(world);
  const res = ctx.repairFormEmailCollection_({ apply: true });
  assert.equal(res.success, true);
  assert.equal(res.data.applied, true);
  assert.deepEqual(plain(res.data.counts), { verified: 1, selfReported: 2, notCollecting: 1, fixed: 2, failed: 0, unreadable: 0 });
  assert.equal(world.forms.selfA1234567._type, TYPES.VERIFIED);
  assert.equal(world.forms.selfC1234567._type, TYPES.VERIFIED);
  assert.equal(world.forms.noneD1234567._type, TYPES.DO_NOT_COLLECT, '集めない設定は先生が決めたもの。変えない');
  assert.equal(world.forms.noneD1234567._sets, 0);
  assert.equal(world.forms.verB12345678._sets, 0, '確認済みには書かない');
  assert.equal(world.forms.strayE123456._type, TYPES.RESPONDER_INPUT, 'アプリが指していないフォームには触らない');
  assert.equal(world.forms.strayE123456._sets, 0);
  // もう一度流しても、直すものは無い
  const again = ctx.repairFormEmailCollection_({ apply: true });
  assert.deepEqual(plain(again.data.counts), { verified: 3, selfReported: 0, notCollecting: 1, fixed: 0, failed: 0, unreadable: 0 });
});

test('repairFormEmailCollection_: 返事に題名や ID の全体を入れない', () => {
  const ctx = loadCtx(sampleWorld());
  const res = ctx.repairFormEmailCollection_({ apply: true });
  const text = JSON.stringify(res);
  assert.equal(/やまだ|1年1組|ダミー/.test(text), false, '題名を返さない');
  assert.equal(/selfA1234567|selfC1234567|verB12345678|noneD1234567/.test(text), false, 'ID の全体を返さない');
  assert.deepEqual(plain(res.data.list).map((x) => `${x.where}:${x.id}:${x.state}`).sort(), [
    'config:selfA1:fixed', 'lesson:noneD1:notCollecting', 'lesson:selfC1:fixed', 'profile:verB12:verified'
  ]);
});

test('repairFormEmailCollection_: 設定が効かない・投げる・読めないフォームは、直したと言わない', () => {
  const world = {
    forms: {
      ignoreA12345: fakeForm('ignoreA12345', TYPES.RESPONDER_INPUT, { ignoresSet: true }),
      throwB123456: fakeForm('throwB123456', TYPES.RESPONDER_INPUT, { throwsOnSet: true }),
      unreadC12345: fakeForm('unreadC12345', TYPES.RESPONDER_INPUT, { unreadable: true }),
      okD123456789: fakeForm('okD123456789', TYPES.RESPONDER_INPUT)
    },
    lessons: [lesson([
      { formId: 'ignoreA12345' }, { formId: 'throwB123456' }, { formId: 'unreadC12345' }, { formId: 'okD123456789' }
    ])]
  };
  const ctx = loadCtx(world);
  const res = ctx.repairFormEmailCollection_({ apply: true });
  assert.equal(res.success, true, '1 つ失敗しても、残りは続ける');
  assert.deepEqual(plain(res.data.counts), { verified: 0, selfReported: 3, notCollecting: 0, fixed: 1, failed: 2, unreadable: 1 });
  assert.equal(world.forms.okD123456789._type, TYPES.VERIFIED);
  assert.deepEqual(Array.from(ctx.__logged), ['repairFormEmailCollection'], '投げた失敗は記録する');
});

test('repairFormEmailCollection_: 設定が壊れた利用者・フェーズの無い授業があっても止まらない', () => {
  const world = {
    forms: { okA123456789: fakeForm('okA123456789', TYPES.RESPONDER_INPUT) },
    users: [{ userId: 'u0', configJson: '{こわれた' }, { userId: 'u1' }, null, user({ profiles: 'ちがう形' })],
    lessons: [{ lessonId: 'x' }, { lessonJson: {} }, null, lesson([null, 'ちがう形', { formId: 'okA123456789' }])]
  };
  const res = loadCtx(world).repairFormEmailCollection_({ apply: true });
  assert.equal(res.success, true);
  assert.equal(res.data.counts.fixed, 1);
});

test('repairFormEmailCollection_: メールの集め方を読めない実行環境では、何もせずに断る', () => {
  const world = sampleWorld();
  const res = loadCtx(world, { FormApp: {} }).repairFormEmailCollection_({ apply: true });
  assert.equal(res.success, false);
  Object.values(world.forms).forEach((f) => assert.equal(f._sets, 0));
});

test('admin API: repairFormEmailCollection は apply: true のときだけ変える', () => {
  const src = fs.readFileSync(path.resolve(__dirname, '../src/AdminApis.js'), 'utf8');
  assert.match(src, /case 'repairFormEmailCollection': \{[\s\S]{0,400}repairFormEmailCollection_\(\{ apply: params\.apply === true \}\)/);
});
