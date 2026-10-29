/**
 * 検証用の配置にだけ入るファイル。本番の配置には入れない (src/ に置かない)。
 *
 * scripts/deploy-single.js が single-tenants.json の extraFiles に書かれたこのファイルを足す。
 * admin API の 'dev.invoke' で、画面が google.script.run で呼ぶ処理を、そのまま名前で呼ぶ。
 *   Why: 画面からしか呼べない処理 (児童の送信・ボードの更新確認など) は外から試せない。
 *   as を渡すと呼び出し元のメールを差し替える (児童のふり)。getCurrentEmail は API キー経路では
 *   _apiKeyAdminEmail を返すので、呼んでいる間だけそれを置き換える。
 *
 *   node scripts/admin-api.js 相当: { action: 'adminApi', apiKey, operation: 'dev.invoke',
 *     params: { fn: 'submitLessonAnswer', args: [...], as: 'student1@example.jp' } }
 */

/* global _apiKeyAdminEmail:writable, createErrorResponse */

// 画面 (src/*.html) が呼ぶ処理だけ。ここに無い名前は呼ばない。
const DEV_INVOKE_ALLOWED_ = [
  'addReaction', 'advanceLessonPhase', 'checkUserAuthentication', 'connectDataSource', 'createTemplateForm',
  'deleteAnswerRow', 'getActiveFormInfo', 'getActiveLessonNav', 'getBatchedAdminData', 'getColumnAnalysis',
  'getFormInfo', 'getForms', 'getMyLessonTrajectory', 'getNotificationUpdate', 'getPublishedSheetData',
  'getWebAppUrl', 'isAdmin', 'processFormUrlInput', 'publishApp', 'republishMyBoard', 'saveConfig',
  'setFormAllowResubmit', 'submitLessonAnswer', 'toggleHighlight', 'unpublishBoard', 'validateHeaderIntegrity',
  'uploadLessonImage',
  // 失敗の中身を見るための内側の処理 (応答に理由が載らないものを、検証用でだけ直接呼ぶ)
  'validateConfig', '__buildPhaseConfigPatch_', 'getUserConfig'
];

function __devAdminOperation_(op, params) {
  if (op !== 'dev.invoke') return null;
  const p = params || {};
  const fn = String(p.fn || '');
  if (DEV_INVOKE_ALLOWED_.indexOf(fn) < 0 || typeof globalThis[fn] !== 'function') {
    return createErrorResponse('dev.invoke: 呼べない関数です: ' + fn, null, { error: 'DEV_FN_NOT_ALLOWED' });
  }
  const prev = _apiKeyAdminEmail;
  if (p.as) _apiKeyAdminEmail = String(p.as).trim().toLowerCase();
  try {
    const result = globalThis[fn].apply(null, Array.isArray(p.args) ? p.args : []);
    return { success: true, devResult: result === undefined ? null : result };
  } catch (e) {
    return createErrorResponse('dev.invoke: ' + fn + ' が例外を投げました: ' + (e && e.message), null, { error: 'DEV_FN_THREW' });
  } finally {
    _apiKeyAdminEmail = prev;
  }
}
