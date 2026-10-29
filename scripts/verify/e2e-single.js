#!/usr/bin/env node
/**
 * 先生専用の配置で、画面の機能を端から端まで一通り動かす (検証用の配置でだけ流す)。
 *
 * 前提: single-tenants.json の対象に extraFiles: ["scripts/verify/DevInvoke.js"] が付いていて
 *   (dev.invoke が使える)、先生が 1 度開いて許可済み。本番の先生には流さない (フォームや授業を作る)。
 *
 * 先生の操作は API キー経路 (= 先生)、児童の操作は dev.invoke の as で児童のふりをして呼ぶ。
 * 出すのは成否と件数だけ。作ったフォーム・授業は名前に【検証】を付ける。
 *
 *   node scripts/verify/e2e-single.js --name t260781p-verify
 */
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
const { getAccessToken, postJSONSync } = require(path.join(__dirname, '..', 'lib', 'gas-auth'));

const CONFIG = path.join(__dirname, '..', 'single-tenants.json');
const args = process.argv.slice(2);
const name = args[args.indexOf('--name') + 1];
const cfgAll = JSON.parse(fs.readFileSync(CONFIG, 'utf8'));
const tn = cfgAll.tenants.find((t) => t.name === name);
if (!tn) throw new Error('--name の配置がありません');
if (!(tn.extraFiles || []).some((f) => /DevInvoke\.js$/.test(f))) throw new Error('検証用の配置 (extraFiles に DevInvoke.js) でだけ流せます');

const domain = tn.adminEmail.split('@')[1];
const url = `https://script.google.com/a/macros/${domain}/s/${tn.deploymentId}/exec`;
const tok = getAccessToken(tn.user || 'organization');
if (!tn.apiKey) {
  const key = crypto.randomBytes(24).toString('base64url');
  const r = postJSONSync(url, { action: 'setupApiKey', apiKey: key }, tok);
  if (!r || !r.success) throw new Error('API キーを設定できませんでした (先生が 1 度開いて許可しましたか?): ' + JSON.stringify(r).slice(0, 200));
  tn.apiKey = key;
  fs.writeFileSync(CONFIG, JSON.stringify(cfgAll, null, 2) + '\n');
}
const op = (operation, params) => postJSONSync(url, { action: 'adminApi', apiKey: tn.apiKey, operation, params: params || {} }, tok);
const inv = (fn, fnArgs, as) => {
  const r = op('dev.invoke', { fn, args: fnArgs || [], as });
  if (!r || r.success === false) return r || { success: false, message: 'no response' };
  return r.devResult;
};
const STUDENT = [1, 2].map((i) => `verify-student-${i}@${domain}`);
const CLS = '6年9組';

const results = [];
const ok = (r) => Boolean(r && r.success !== false && !(r.error && r.success !== true));
const text = (r) => String((r && (r.message || r.error)) || '').replace(/[\w.+-]+@[\w.-]+/g, '<email>').slice(0, 140);
function check(area, label, fn) {
  let pass = false;
  let note = '';
  try {
    const out = fn();
    pass = out === true || (out && out.pass === true);
    note = (out && out.note) || '';
  } catch (e) {
    note = '例外: ' + text(e);
  }
  results.push({ area, label, pass, note });
  console.log(`${pass ? '✓' : '✗'}  [${area}] ${label}${note ? '  — ' + note : ''}`);
  return pass;
}
const rowsOf = (r) => (r && (r.data || r.rows)) || [];

const userId = (op('getUsers').users || []).find((u) => String(u.userEmail).toLowerCase() === tn.adminEmail)?.userId;
if (!userId) throw new Error('先生が登録されていません (管理パネルを 1 度開いてください)');

// ---------------- 先生の画面: 掲示板モード ----------------
let form = null;
let sheet = null;
check('掲示板', '管理パネルを開く (getBatchedAdminData)', () => ({ pass: ok(inv('getBatchedAdminData', [userId])) }));
check('掲示板', '自分のフォーム一覧 (getForms)', () => { const r = inv('getForms'); return { pass: ok(r), note: ok(r) ? `${(r.forms || r.data || []).length} 件` : text(r) }; });
check('掲示板', 'テンプレートからフォームを作る (createTemplateForm: matrix)', () => {
  form = inv('createTemplateForm', ['matrix', { title: '【検証】マトリクス（消してよい）' }]);
  return { pass: ok(form) && Boolean(form.spreadsheetId), note: ok(form) ? '' : text(form) };
});
// 管理パネルと同じ順番: 作ったフォームの回答先を設定に下書き保存してから (自動保存)、URL の読み込み → 接続 → 列の推定。
//   回答シートは「どれかのボードの設定に入っている」ときだけ開ける (validateServiceAccountUsage)。
check('掲示板', '作ったフォームを設定に下書き保存する (saveConfig 自動保存)', () => {
  sheet = { id: form.spreadsheetId, name: form.sheetName };
  const r = inv('saveConfig', [{ spreadsheetId: sheet.id, sheetName: sheet.name, formUrl: form.formUrl, formTitle: form.formTitle }, { isDraft: true }]);
  return { pass: ok(r), note: ok(r) ? '' : text(r) };
});
check('掲示板', 'フォームの URL を読み込む (processFormUrlInput)', () => {
  const r = inv('processFormUrlInput', [form.editUrl]);
  return { pass: ok(r) && r.spreadsheetId === sheet.id, note: ok(r) ? '' : text(r) };
});
check('掲示板', 'データソースをつなぐ (connectDataSource)', () => {
  const ops = ['validateAccess', 'getFormInfo', 'connectDataSource'].map((type) => ({ type, spreadsheetId: sheet.id, sheetName: sheet.name }));
  const r = inv('connectDataSource', [sheet.id, sheet.name, ops]);
  return { pass: ok(r), note: ok(r) ? '' : text(r) };
});
let analysis = null;
check('掲示板', '列の自動推定 (回答 0 件 = 見出しだけ)', () => {
  analysis = inv('getColumnAnalysis', [sheet.id, sheet.name]);
  return { pass: ok(analysis) && (analysis.headers || []).length > 0, note: ok(analysis) ? `見出し ${(analysis.headers || []).length} 列` : text(analysis) };
});
check('掲示板', 'フォームの情報 (getFormInfo)', () => { const r = inv('getFormInfo', [sheet.id, sheet.name]); return { pass: ok(r), note: ok(r) ? '' : text(r) }; });
check('掲示板', '再回答の許可を切り替える (setFormAllowResubmit)', () => { const r = inv('setFormAllowResubmit', [form.formId, true]); return { pass: ok(r), note: ok(r) ? '' : text(r) }; });

// 回答 0 件では軸の列 (X軸 / Y軸) を数値の列と推定できない (値の形も手がかりにするため)。実際の授業と同じく、
//   児童の回答が入ってから推定し直す。フォームの回答のかわりに、見出しに合わせた行を足す。
const headers = (analysis && analysis.headers) || [];
check('掲示板', '児童の回答を入れる (フォームの回答のかわりに行を足す)', () => {
  const rows = STUDENT.map((email, i) => headers.map((h) => {
    if (/タイムスタンプ/.test(h)) return '';
    if (/メール/.test(h)) return email;
    if (/クラス/.test(h)) return CLS;
    if (/名前/.test(h)) return `けんしょう${i + 1}`;
    if (/理由|意見|考え/.test(h)) return `検証の理由 ${i + 1}`;
    if (/^(UNDERSTAND|LIKE|CURIOUS|HIGHLIGHT)$/.test(h)) return '';
    return 2 + i;
  }));
  const r = op('appendRows', { spreadsheetId: sheet.id, sheetName: sheet.name, rows });
  return { pass: ok(r), note: ok(r) ? `${rows.length} 行` : text(r) };
});
check('掲示板', '列の自動推定 (回答が入ったあと)', () => {
  analysis = inv('getColumnAnalysis', [sheet.id, sheet.name]);
  const m = (analysis && analysis.mapping) || {};
  return { pass: ok(analysis) && m.numericX !== undefined && m.numericY !== undefined, note: ok(analysis) ? `推定した列: ${Object.keys(m).join(',')}` : text(analysis) };
});
const boardConfig = {
  spreadsheetId: sheet.id, sheetName: sheet.name, formUrl: form.formUrl, formTitle: form.formTitle,
  columnMapping: (analysis && analysis.mapping) || {}, displaySettings: { boardMode: 'matrix', showNames: false, reactionMode: 'count' }
};
check('掲示板', '設定を保存する (saveConfig 下書き)', () => { const r = inv('saveConfig', [boardConfig, { isDraft: true }]); return { pass: ok(r), note: ok(r) ? '' : text(r) }; });
check('掲示板', '公開する (publishApp)', () => {
  const cur = op('getUserConfig', { userId }).data.config;
  const r = inv('publishApp', [Object.assign({}, boardConfig, { etag: cur.etag })]);
  return { pass: ok(r), note: ok(r) ? '' : text(r) };
});

// ---------------- 児童の画面: 掲示板モード ----------------
check('児童(掲示板)', 'ボードを読む (getPublishedSheetData)', () => { const r = inv('getPublishedSheetData', [null, 'newest', false, userId], STUDENT[0]); return { pass: ok(r) && rowsOf(r).length === 2, note: ok(r) ? `${rowsOf(r).length} 行` : text(r) }; });
check('児童(掲示板)', '更新を確かめる (getNotificationUpdate)', () => { const r = inv('getNotificationUpdate', [userId, {}], STUDENT[0]); return { pass: ok(r), note: ok(r) ? '' : text(r) }; });
check('児童(掲示板)', 'ログイン確認 (checkUserAuthentication)', () => { const r = inv('checkUserAuthentication', [], STUDENT[0]); return { pass: ok(r), note: ok(r) ? '' : text(r) }; });
check('児童(掲示板)', '先生ではない (isAdmin = false)', () => { const r = inv('isAdmin', [], STUDENT[0]); return { pass: r === false || (r && r.isAdmin === false), note: JSON.stringify(r).slice(0, 60) }; });
check('児童(掲示板)', 'フォームの情報 (getActiveFormInfo)', () => { const r = inv('getActiveFormInfo', [userId], STUDENT[0]); return { pass: ok(r), note: ok(r) ? '' : text(r) }; });
let firstRow = null;
check('児童(掲示板)', 'リアクションを付ける (addReaction)', () => {
  const b = inv('getPublishedSheetData', [null, 'newest', false, userId], STUDENT[0]);
  firstRow = rowsOf(b)[0];
  const r = inv('addReaction', [userId, firstRow.rowIndex, 'LIKE'], STUDENT[1]);
  return { pass: ok(r), note: ok(r) ? '' : text(r) };
});
check('先生(掲示板)', 'ハイライト (toggleHighlight)', () => { const r = inv('toggleHighlight', [userId, firstRow.rowIndex]); return { pass: ok(r), note: ok(r) ? '' : text(r) }; });
check('先生(掲示板)', '見出しの確認 (validateHeaderIntegrity)', () => { const r = inv('validateHeaderIntegrity', [userId]); return { pass: ok(r), note: ok(r) ? '' : text(r) }; });
check('先生(掲示板)', '回答を消す (deleteAnswerRow)', () => {
  const b = inv('getPublishedSheetData', [null, 'newest', true, userId]);
  const last = rowsOf(b)[rowsOf(b).length - 1];
  const r = inv('deleteAnswerRow', [userId, last.rowIndex, last.timestamp]);
  const after = rowsOf(inv('getPublishedSheetData', [null, 'newest', false, userId], STUDENT[0])).length;
  return { pass: ok(r) && after === 1, note: ok(r) ? `残り ${after} 行` : text(r) };
});
check('先生(掲示板)', '公開を終了する → 児童は読めない (unpublishBoard)', () => {
  const r = inv('unpublishBoard', [userId]);
  const s = inv('getPublishedSheetData', [null, 'newest', false, userId], STUDENT[0]);
  return { pass: ok(r) && !ok(s), note: ok(r) ? `児童: ${text(s) || '読めた (!!)'}` : text(r) };
});
check('先生(掲示板)', '再び公開する → 児童が読める (republishMyBoard)', () => {
  const r = inv('republishMyBoard', [{}]);
  const s = inv('getPublishedSheetData', [null, 'newest', false, userId], STUDENT[0]);
  return { pass: ok(r) && ok(s), note: ok(r) ? '' : text(r) };
});

// ---------------- 児童が管理の処理を直接呼んでも断られる (先生として実行なので大事) ----------------
for (const [fn, fnArgs] of [
  ['getForms', []], ['createTemplateForm', ['board', {}]], ['processFormUrlInput', [form.editUrl]],
  ['setFormAllowResubmit', [form.formId, false]], ['connectDataSource', [sheet.id, sheet.name]],
  ['getColumnAnalysis', [sheet.id, sheet.name]], ['getFormInfo', [sheet.id, sheet.name]],
  ['uploadLessonImage', ['data:image/png;base64,iVBORw0KGgo=', 'x.png']], ['processLoginAction', []]
]) {
  check('児童(権限)', `${fn} を児童が呼ぶと断られる`, () => {
    const r = inv(fn, fnArgs, STUDENT[0]);
    return { pass: Boolean(r) && r.success === false, note: text(r) };
  });
}

// ---------------- 授業モード (native) ----------------
let lessonId = null;
const nativeIds = [];
check('授業', '授業を作る (lesson.create)', () => {
  const r = op('lesson.create', { userId, name: '【検証】授業モード（消してよい）', template: 'dialogue-reconsider-5phase' });
  lessonId = r.data && (r.data.lessonId || (r.data.lesson && r.data.lesson.lessonId));
  if (lessonId) nativeIds.push(lessonId);
  return { pass: Boolean(lessonId), note: lessonId ? '' : text(r) };
});
check('授業', '学級と問いを書く (lesson.updateDraft)', () => {
  const etag = op('lesson.review', { userId, lessonId }).data.lesson.etag;
  const r1 = op('lesson.updateDraft', { userId, lessonId, fieldPath: 'classes', value: [CLS], expectedEtag: etag });
  const etag2 = op('lesson.review', { userId, lessonId }).data.lesson.etag;
  const r2 = op('lesson.updateDraft', { userId, lessonId, fieldPath: 'phases[0].question', value: '検証の問い', expectedEtag: etag2 });
  return { pass: ok(r1) && ok(r2), note: ok(r1) && ok(r2) ? '' : text(ok(r1) ? r2 : r1) };
});
check('授業', '画像をのせる (uploadLessonImage)', () => {
  const png = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
  const r = inv('uploadLessonImage', [png, '【検証】1px.png']);
  return { pass: ok(r), note: ok(r) ? '' : text(r) };
});
check('授業', '授業を始める → ボードが授業に (lesson.start)', () => {
  const r = op('lesson.start', { userId, lessonId });
  const cfg = op('getUserConfig', { userId }).data.config;
  return { pass: ok(r) && cfg.activeLessonId === lessonId, note: ok(r) ? '' : text(r) };
});
check('授業', 'ボードを公開した状態にする (republishMyBoard)', () => { const cfg = op('getUserConfig', { userId }).data.config; const r = cfg.isPublished ? { success: true } : inv('republishMyBoard', [{}]); return { pass: ok(r), note: cfg.isPublished ? 'もともと公開中' : (ok(r) ? '' : text(r)) }; });
check('授業', '授業のナビ (getActiveLessonNav)', () => { const r = inv('getActiveLessonNav', [userId]); return { pass: ok(r), note: ok(r) ? '' : text(r) }; });
const submit = (who, phaseIndex, x, y) => inv('submitLessonAnswer', [userId, { lessonId, phaseIndex, numericX: x, numericY: y, reason: `検証 ${phaseIndex}`, class: CLS, name: 'けんしょう' }], who);
check('児童(授業)', '「考える」で送る (submitLessonAnswer × 2 人)', () => { const a = submit(STUDENT[0], 0, 2, 4); const b = submit(STUDENT[1], 0, 4, 2); return { pass: ok(a) && ok(b), note: ok(a) && ok(b) ? '' : text(ok(a) ? b : a) }; });
check('児童(授業)', '「考える」の間は他の子の回答が見えない', () => { const r = inv('getPublishedSheetData', [null, 'newest', false, userId], STUDENT[0]); return { pass: ok(r) && rowsOf(r).length === 1, note: ok(r) ? `見えた行 ${rowsOf(r).length}` : text(r) }; });
check('児童(授業)', '同じフェーズで送り直すと行が増えない', () => { submit(STUDENT[0], 0, 3, 3); const r = op('lesson.liveAnswers', { userId, lessonId }); return { pass: ok(r) && (r.data.rows || []).length === 2, note: ok(r) ? `先生の一覧 ${(r.data.rows || []).length} 行` : text(r) }; });
check('先生(授業)', 'ハイライト (lesson.toggleHighlight)', () => {
  const live = op('lesson.liveAnswers', { userId, lessonId }).data;
  const row = (live.rows || [])[0];
  const r = op('lesson.toggleHighlight', { userId, lessonId, rowIndex: row.rowIndex, sheetName: (live.phase && live.phase.sourceSheetName) || '', expectedEmail: row.email || '' });
  return { pass: ok(r), note: ok(r) ? '' : text(r) };
});
check('先生(授業)', 'フェーズを進める: 出会う (advanceLessonPhase)', () => { const r = inv('advanceLessonPhase', [userId, lessonId, 'next', null]); return { pass: ok(r), note: ok(r) ? '' : text(r) }; });
check('児童(授業)', '「出会う」で全員の回答が見え、名前は出ない', () => {
  const r = inv('getPublishedSheetData', [null, 'newest', false, userId], STUDENT[0]);
  const named = rowsOf(r).filter((x) => x.name).length;
  return { pass: ok(r) && rowsOf(r).length === 2 && named === 0, note: ok(r) ? `見えた行 ${rowsOf(r).length}・名前つき ${named}` : text(r) };
});
check('先生(授業)', 'フェーズを進める: 議論する', () => { const r = inv('advanceLessonPhase', [userId, lessonId, 'next', null]); return { pass: ok(r), note: ok(r) ? '' : text(r) }; });
check('児童(授業)', '「議論する」では送れない (サーバが断る)', () => { const r = submit(STUDENT[0], 2, 1, 1); return { pass: !ok(r), note: text(r) }; });
check('先生(授業)', 'フェーズを進める: もう一度考える', () => { const r = inv('advanceLessonPhase', [userId, lessonId, 'next', null]); return { pass: ok(r), note: ok(r) ? '' : text(r) }; });
check('児童(授業)', '「もう一度考える」で送る', () => { const a = submit(STUDENT[0], 3, 5, 5); return { pass: ok(a), note: ok(a) ? '' : text(a) }; });
check('児童(授業)', '自分の考えの移り変わり (getMyLessonTrajectory)', () => { const r = inv('getMyLessonTrajectory', [userId], STUDENT[0]); return { pass: ok(r), note: ok(r) ? '' : text(r) }; });
check('児童(授業)', '更新を確かめる (getNotificationUpdate, 授業中)', () => { const r = inv('getNotificationUpdate', [userId, {}], STUDENT[1]); return { pass: ok(r) && r.lessonPhase !== undefined, note: ok(r) ? `フェーズ ${r.lessonPhase && r.lessonPhase.phaseIndex}` : text(r) }; });
check('授業', 'アーカイブを取り直す (lesson.recaptureArchive, いま映しているフェーズ)', () => { const r = op('lesson.recaptureArchive', { userId, lessonId, phaseIndex: 3 }); return { pass: ok(r), note: ok(r) ? '' : text(r) }; });
check('先生(授業)', 'フェーズを進める: ふりかえる', () => { const r = inv('advanceLessonPhase', [userId, lessonId, 'next', null]); return { pass: ok(r), note: ok(r) ? '' : text(r) }; });
check('授業', '授業を終える → アーカイブ (lesson.end)', () => { const r = op('lesson.end', { userId, lessonId }); return { pass: ok(r), note: ok(r) ? '' : text(r) }; });
check('授業', 'ふりかえり (lesson.review)', () => {
  const r = op('lesson.review', { userId, lessonId });
  const snaps = (r.data && r.data.lesson && r.data.lesson.lessonJson.snapshots) || [];
  return { pass: ok(r) && snaps.length > 0, note: ok(r) ? `アーカイブ ${snaps.length} フェーズ・${snaps.reduce((n, s) => n + (Number(s.rowCount) || (s.rows || []).length), 0)} 行` : text(r) };
});
check('授業', '見取りグリッド (lesson.reviewGrid)', () => { const r = op('lesson.reviewGrid', { userId, lessonId }); return { pass: ok(r) && (r.data.students || []).length === 2, note: ok(r) ? `児童 ${(r.data.students || []).length} 人` : text(r) }; });
check('授業', '再開する → もう一度終える (lesson.reopen / end)', () => { const a = op('lesson.reopen', { userId, lessonId }); const b = op('lesson.end', { userId, lessonId }); return { pass: ok(a) && ok(b), note: ok(a) && ok(b) ? '' : text(ok(a) ? b : a) }; });
check('授業', '複製する (lesson.duplicate)', () => { const r = op('lesson.duplicate', { userId, lessonId, options: {} }); const id = r.data && (r.data.lessonId || (r.data.lesson && r.data.lesson.lessonId)); if (id) nativeIds.push(id); return { pass: Boolean(id), note: id ? '' : text(r) }; });

// ---------------- フォームを使う授業 ----------------
check('授業(フォーム)', 'フォームを使うひな形で作って始める (lesson.create / start)', () => {
  const tpls = op('lesson.templates', { userId }).data;
  const list = Array.isArray(tpls) ? tpls : (tpls.templates || []);
  const formTpl = list.find((t) => t && t.inputMode === 'form');
  if (!formTpl) return { pass: false, note: 'フォームを使うひな形が見つかりません' };
  const c = op('lesson.create', { userId, name: '【検証】フォームの授業（消してよい）', template: formTpl.key });
  const id = c.data && (c.data.lessonId || (c.data.lesson && c.data.lesson.lessonId));
  if (!id) return { pass: false, note: text(c) };
  nativeIds.push(id);
  const etag = op('lesson.review', { userId, lessonId: id }).data.lesson.etag;
  op('lesson.updateDraft', { userId, lessonId: id, fieldPath: 'classes', value: [CLS], expectedEtag: etag });
  const s = op('lesson.start', { userId, lessonId: id });
  return { pass: ok(s), note: ok(s) ? `ひな形 ${formTpl.key}` : text(s) };
});
check('授業(フォーム)', 'フォームを閉じる → 終える (lesson.closeForms / end)', () => {
  const id = nativeIds[nativeIds.length - 1];
  const a = op('lesson.closeForms', { userId, lessonId: id });
  const b = op('lesson.end', { userId, lessonId: id });
  return { pass: ok(a) && ok(b), note: ok(a) && ok(b) ? '' : text(ok(a) ? b : a) };
});

// ---------------- 後片付け (授業は消す。フォームとシートは Drive に残る) ----------------
nativeIds.forEach((id, i) => check('後片付け', `授業を消す ${i + 1} (lesson.delete)`, () => { const r = op('lesson.delete', { userId, lessonId: id }); return { pass: ok(r), note: ok(r) ? '' : text(r) }; }));

const failed = results.filter((r) => !r.pass);
console.log(`\n${results.length - failed.length} / ${results.length} 項目 OK`);
if (failed.length) process.exit(1);
