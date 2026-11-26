/**
 * 授業の自動保存は「編集したときの授業」に届く。
 *
 * Why: 送り先を flush 時の currentLessonId で決めていたため、ある授業を編集した直後
 *   (debounce 600ms 以内) に別の授業を開くと、その変更が開いた先の授業に保存された。
 *   2026-10-21、実行中の「会話のゆくえ」の学級が「6年1組」だけになり、学級の選択欄が消えた。
 *   etag_mismatch の retry も、開いた先の授業を読み直して保存を通してしまっていた。
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const SRC = fs.readFileSync(path.resolve(__dirname, '../src/AdminPanel.js.html'), 'utf8');
const BLOCK = SRC.slice(SRC.indexOf('// ---- lesson autosave: begin ----'), SRC.indexOf('// ---- lesson autosave: end ----'));

function harness(serverEtags) {
  const calls = [];
  let debounced = null;
  const ctx = {
    console: { warn: () => {} },
    Promise,
    userId: 'u1',
    FIELD_DEBOUNCE_MS: 600,
    calls,
    document: { getElementById: () => null },
    window: { sharedUtilities: { debounce: { debounce: (fn) => { debounced = fn; } } } },
    renderConfirmStep: () => {},
    parseFieldValue: (el) => el.value,
    gasRun: (op, params) => {
      calls.push({ op, ...params });
      const etag = serverEtags[params.lessonId];
      if (op === 'lesson.review') return Promise.resolve({ success: true, data: { lesson: { etag } } });
      if (params.expectedEtag !== etag) return Promise.resolve({ success: false, error: 'etag_mismatch' });
      serverEtags[params.lessonId] = etag + '+';
      return Promise.resolve({ success: true, data: { lesson: { etag: serverEtags[params.lessonId], lessonJson: {} } } });
    }
  };
  vm.createContext(ctx);
  vm.runInContext('let currentLessonId = null; let currentLessonEtag = null;\n' + BLOCK +
    '\nthis.api = { schedule: scheduleLessonFieldSave, flush: flushLessonSaves, ' +
    'open: (id, etag) => { flushLessonSaves(); currentLessonId = id; currentLessonEtag = etag; }, ' +
    'queue: () => saveQueue };', ctx);
  return { api: ctx.api, calls, fireDebounce: () => debounced && debounced() };
}

const field = (name, value) => ({ dataset: { lessonField: name }, value });

test('編集した直後に別の授業を開いても、変更は編集した授業に保存される', async () => {
  const h = harness({ trial: 'e1', kaiwa: 'k1' });
  h.api.open('trial', 'e1');
  h.api.schedule(field('classes', '6年1組'));
  h.api.open('kaiwa', 'k1');           // debounce が発火する前に別の授業を開く
  h.fireDebounce();
  await h.api.queue();
  const saves = h.calls.filter((c) => c.op === 'lesson.updateDraft');
  assert.equal(saves.length, 1);
  assert.equal(saves[0].lessonId, 'trial');
  assert.equal(saves[0].expectedEtag, 'e1');
  assert.ok(!h.calls.some((c) => c.lessonId === 'kaiwa'), '開いた先の授業には何も送らない');
});

test('etag がずれて retry するときも、読み直すのは編集した授業', async () => {
  const server = { trial: 'e2', kaiwa: 'k1' };      // trial の etag がサーバで進んでいる
  const h = harness(server);
  h.api.open('trial', 'e1');
  h.api.schedule(field('name', '試し'));
  h.api.flush();
  h.api.open('kaiwa', 'k1');
  await h.api.queue();
  assert.deepEqual(h.calls.map((c) => [c.op, c.lessonId]),
    [['lesson.updateDraft', 'trial'], ['lesson.review', 'trial'], ['lesson.updateDraft', 'trial']]);
});

test('別の授業を編集し始めたら、前の授業の書きかけを先に送る (混ぜない)', async () => {
  const h = harness({ a: 'a1', b: 'b1' });
  h.api.open('a', 'a1');
  h.api.schedule(field('name', 'A'));
  // 授業を開く経路を通らずに currentLessonId が変わっても、混ぜずに分けて送る
  h.api.open('b', 'b1');
  h.api.schedule(field('name', 'B'));
  h.fireDebounce();
  await h.api.queue();
  const saves = h.calls.filter((c) => c.op === 'lesson.updateDraft').map((c) => [c.lessonId, c.value]);
  assert.deepEqual(saves, [['a', 'A'], ['b', 'B']]);
});

test('授業を開く・編集をやめる・戻るボタンの前に、書きかけを送る', () => {
  for (const fn of ['function loadLessonThen', 'function exitLessonCreateMode']) {
    const body = SRC.slice(SRC.indexOf(fn), SRC.indexOf(fn) + 400);
    assert.match(body, /flushLessonSaves\(\);[\s\S]*currentLessonId = /, fn);
  }
  const pop = SRC.slice(SRC.indexOf("window.addEventListener('popstate'"), SRC.indexOf("window.addEventListener('popstate'") + 600);
  assert.match(pop, /flushLessonSaves\(\);\s*currentLessonId = null;/);
});

test('確認の段を描く処理は、一覧の欄が無くても設定漏れの警告まで落ちずに描く', () => {
  // Why: native / firstInput を一覧のブロックの中だけで宣言していて、警告のブロックが
  //   「native is not defined」で落ち、保存は通ったのに表示が「保存失敗」になった (2026-11-26)。
  const start = SRC.indexOf('function renderConfirmStep(lessonJson)');
  const fn = SRC.slice(start, SRC.indexOf('// ---------- Per-field autosave', start));
  const els = { 'lesson-summary-warnings': { innerHTML: '' }, 'lesson-summary-phase-list': { innerHTML: '' },
    'lesson-summary-classes': { textContent: '' } };
  const ctx = {
    document: { getElementById: (id) => els[id] || null, querySelector: () => ({ hidden: false }) },
    window: { sharedUtilities: { security: { escapeHtml: (s) => String(s) } } },
    applyLessonModeCopy: () => {}, roleLabel: (r) => r, PHASE_TEMPLATE_OPTIONS: []
  };
  vm.createContext(ctx);
  vm.runInContext(fn + '\nthis.render = renderConfirmStep;', ctx);
  const lessonJson = { inputMode: 'native', classes: ['6年1組'],
    phases: [{ name: '考える', question: 'q', screenRole: 'input', formTemplate: 'matrix', templateOptions: {} }] };
  ctx.render(lessonJson);
  assert.equal(els['lesson-summary-classes'].textContent, '6年1組');
  assert.match(els['lesson-summary-warnings'].innerHTML, /軸の両端が空です/);
  delete els['lesson-summary-phase-list'];   // 一覧の欄が無い画面でも落ちない
  ctx.render(lessonJson);
});

test('作ったフォームは、先生がこの画面に戻ってきたときに 1 回だけ読み取って接続する', () => {
  const block = SRC.slice(SRC.indexOf('// ---- template form return: begin ----'), SRC.indexOf('// ---- template form return: end ----'));
  const listeners = {};
  const target = (name) => ({
    addEventListener: (ev, fn) => { (listeners[name + ':' + ev] ||= new Set()).add(fn); },
    removeEventListener: (ev, fn) => { (listeners[name + ':' + ev] || new Set()).delete(fn); }
  });
  const fire = (key) => [...(listeners[key] || [])].forEach((fn) => fn());
  let validated = 0;
  const ctx = { document: { hidden: false, ...target('doc') }, window: target('win'), validateCompleteUrl: () => { validated++; } };
  vm.createContext(ctx);
  vm.runInContext(block + '\nthis.wait = connectTemplateFormOnReturn_;', ctx);
  ctx.wait();
  fire('win:focus');                        // 一度も離れていない → 何もしない
  assert.equal(validated, 0);
  ctx.document.hidden = true; fire('doc:visibilitychange');
  assert.equal(validated, 0);
  ctx.document.hidden = false; fire('doc:visibilitychange'); fire('win:focus');
  assert.equal(validated, 1, '戻ったら 1 回だけ');
  ctx.document.hidden = true; fire('doc:visibilitychange');
  ctx.document.hidden = false; fire('doc:visibilitychange');
  assert.equal(validated, 1, '2 回めの戻りでは読み直さない');
  assert.match(SRC, /connectTemplateFormOnReturn_\(\);\s*showSuccess\(result\.message/);
});
