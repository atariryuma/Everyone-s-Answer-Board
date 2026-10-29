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
 *
 * 'dev.trashTestFiles' は検証で作ったフォーム・シート・画像を先生の Drive のゴミ箱へ移す
 *   (e2e-single.js の最後の項目)。先生として実行するので、検証のたびに先生の Drive に溜まるため。
 */

/* global _apiKeyAdminEmail:writable, createErrorResponse, getCachedProperty_, DriveApp, Utilities */

// 画面 (src/*.html) が呼ぶ処理だけ。ここに無い名前は呼ばない。
const DEV_INVOKE_ALLOWED_ = [
  'addReaction', 'advanceLessonPhase', 'checkUserAuthentication', 'connectDataSource', 'createTemplateForm',
  'deleteAnswerRow', 'getActiveFormInfo', 'getActiveLessonNav', 'getBatchedAdminData', 'getColumnAnalysis',
  'getFormInfo', 'getForms', 'getMyLessonTrajectory', 'getNotificationUpdate', 'getPublishedSheetData',
  'getWebAppUrl', 'isAdmin', 'processFormUrlInput', 'publishApp', 'republishMyBoard', 'saveConfig',
  'setFormAllowResubmit', 'submitLessonAnswer', 'toggleHighlight', 'unpublishBoard', 'validateHeaderIntegrity',
  'uploadLessonImage', 'processLoginAction',
  // 失敗の中身を見るための内側の処理 (応答に理由が載らないものを、検証用でだけ直接呼ぶ)
  'validateConfig', '__buildPhaseConfigPatch_', 'getUserConfig'
];

// 検証で作った Drive のファイル (フォーム・回答のスプレッドシート・画像) の名前。e2e-single.js は
//   作るものの名前を【検証】で始める。画像は uploadLessonImage が【】を _ に置き換えるので _検証_ になる。
const DEV_TEST_FILE_NAME_ = /^(【検証】|_検証_)/;
const DEV_TEST_FILE_MAX_AGE_DAYS_ = 30;

// 先生の Drive から、この 30 日に作った検証のファイルをゴミ箱へ移す (30 日はゴミ箱から戻せる)。
//   dryRun なら名前と作った時刻を返すだけ。
function __devTrashTestFiles_(p) {
  try {
    const since = new Date(Date.now() - DEV_TEST_FILE_MAX_AGE_DAYS_ * 86400000);
    // DriveApp の検索は Drive API v2 の書き方で、作った日では絞れない。この 30 日に作ったファイルは
    //   この 30 日に更新されてもいるので、更新日で絞ってから作った日を確かめる。
    const sinceText = Utilities.formatDate(since, 'UTC', "yyyy-MM-dd'T'HH:mm:ss");
    const it = DriveApp.searchFiles(`modifiedDate > '${sinceText}' and trashed = false and 'me' in owners`);
    // 検証用の DB も名前が【検証】で始まる。消すと検証用の配置が動かなくなる
    const dbId = getCachedProperty_('DATABASE_SPREADSHEET_ID');
    const files = [];
    while (it.hasNext()) {
      const f = it.next();
      const name = f.getName();
      if (f.getId() === dbId || f.getDateCreated() < since) continue;
      if (DEV_TEST_FILE_NAME_.test(name)) files.push(f);
    }
    const list = files.map((f) => ({
      name: f.getName(),
      created: Utilities.formatDate(f.getDateCreated(), 'Asia/Tokyo', 'M/d HH:mm:ss'),
      mimeType: f.getMimeType()
    }));
    if (!p.dryRun) files.forEach((f) => f.setTrashed(true));
    return { success: true, dryRun: Boolean(p.dryRun), count: files.length, files: list };
  } catch (e) {
    return createErrorResponse('dev.trashTestFiles: ' + (e && e.message), null, { error: 'DEV_TRASH_FAILED' });
  }
}

function __devAdminOperation_(op, params) {
  if (op === 'dev.trashTestFiles') return __devTrashTestFiles_(params || {});
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
