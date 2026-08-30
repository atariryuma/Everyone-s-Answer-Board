/**
 * @fileoverview LessonService - 授業 (lesson) の作成・進行・archive。
 *   状態機械: draft (wizard 編集中) → active (Form 生成済み・授業中) → completed (archive 済み)
 *   owner-only auth (管理者は listLessons のみ全件取得可)。
 */

/* global openDatabase, openSpreadsheet, getCurrentEmail, isAdministrator, findUserByEmail, findUserById, findPublishedBoardOwner, createTemplateForm, applyConfigPatch_, applySpreadsheetSharingDefaults, getPublishedSheetData, getPublishedSheetDataForProfile, getAllUsers, getConfigOrDefault, getCachedProperty, bumpBoardDataVersion_, emailToShortHash, LESSONS_SHEET_HEADERS, LESSON_RESPONSES_SHEET_HEADERS, deepClone, createSuccessResponse, createErrorResponse, createExceptionResponse, createUserNotFoundError, createAuthError, isBoardCollaborator, logError_ */

// schemaVersion を bump するときは migration 計画を必ず書く。Phase 1 = 1。
const LESSON_SCHEMA_VERSION = 1;
// Sheets 1 cell 文字数上限 50000 に対する defensive cap。lessonJson が超えるなら save reject。
//   回答アーカイブは lesson_responses シートに分離済みなので、ここに残るのは
//   定義 + 遷移履歴 + 範囲ポインタ (~4KB) のみ。この guard は通常発火しない。
const LESSON_JSON_MAX_BYTES = 45000;
// 回答アーカイブの置き場所。1 回答 = 1 行 (schema は LESSON_RESPONSES_SHEET_HEADERS)。
//   snapshot ポインタは sheet 名を持つので、将来 lesson_responses_2027 のような
//   年次シートへ移行しても過去ポインタはそのまま読める。
const LESSON_RESPONSES_SHEET = 'lesson_responses';
// startLesson 内部の double-click 防止 lock。
const LESSON_START_LOCK_TIMEOUT_MS = 5000;

// ----- 授業モード (native 入力) -----
//
// Google Form を経由せず、回答ボードの画面から直接投稿する phase の入力経路。
// Why Form を使わないか: 「最初の自分を見ながら、もう一度置く」体験は Form では作れない
//   (Form 画面に自分の過去は存在しない)。児童の同定も Form の氏名欄ではなく
//   Session の email で行うため、表記ゆれによる突合ミスが原理的に起きない。
//
// 回答シートの列は Form 回答シートと同じ形に揃える。こうすることで
// getPublishedSheetData / renderMatrix / snapshot capture が一切の分岐なしに動く。
const LESSON_NATIVE_SHEET_HEADERS = [
  'タイムスタンプ', 'メールアドレス', 'クラス', '名前', '横軸', '縦軸', '理由', '加わったこと'
];
// 列推定 (inferColumnRoles) を通さず、この固定 index を columnMapping に直接入れる。
//   native シートは我々が作るので、列の意味は最初から確定している。
// native phase の columnMapping。
//   validateMapping の規約に合わせる (実際に稼働している matrix ボードと同じ形):
//     - answer は必須列
//     - numericX / numericY だけ重複が免除される (「立場」列を answer としても
//       numericX としても見る、という設計が意図されている)
//     - answer と reason が同じ列だと「列インデックスに重複があります」で
//       config 保存が落ちる = 教師が「開始」を押した瞬間に失敗する
//   matrix / numberline は「回答 = 座標」なので answer を軸列に、本文は reason。
//   board / pie は自由記述そのものが回答なので answer を本文列に置く。
function __nativeColumnMapping_(formTemplate) {
  const base = { email: 1, class: 2, name: 3 };
  if (formTemplate === 'board' || formTemplate === 'pie') {
    return Object.assign({}, base, { answer: 6 });
  }
  const m = Object.assign({}, base, { numericX: 4, answer: 4, reason: 6 });
  if (formTemplate === 'matrix') m.numericY = 5;
  return m;
}

const LESSON_NATIVE_COL_TIMESTAMP = 1;   // 1-based (getRange 用)
const LESSON_NATIVE_COL_EMAIL = 2;
const LESSON_NATIVE_COL_X = 5;
const LESSON_NATIVE_COL_INSIGHT = 8;

// 画面の権能。教師のフェーズ送りが、全児童の画面で「何ができるか」を決める。
//   Why 権能を phase に持たせるか: 「アプリが議論を代替しない」「先入観より先に自分の考えを持つ」
//   は機能の有無では守れない。見えるもの・できることをフェーズが制御して初めて構造で保証される。
const LESSON_SCREEN_ROLES = Object.freeze({
  INPUT: 'input',       // 自分の考えを入力。他者は見えない (先入観の遮断)
  BROWSE: 'browse',     // 学級の分布と理由を読む。入力はロック
  DISCUSS: 'discuss',   // 端末は停止。対話の時間 (アプリが黙る)
  REINPUT: 'reinput',   // 自分の ● を見ながら ★ を置き直す + 加わったことを書く
  REFLECT: 'reflect'    // 自分の航跡と自分の言葉だけ。学級の分布は出さない
});
const LESSON_INPUT_ROLES = [LESSON_SCREEN_ROLES.INPUT, LESSON_SCREEN_ROLES.REINPUT];
// 理由 / 加わったこと の文字数上限 (Sheets セル上限ではなく、児童が書く現実的な長さ)。
const LESSON_ANSWER_TEXT_MAX = 500;
// Auto-archive thresholds (unpublishBoard hook + daily sweep)。
const LESSON_AUTO_ARCHIVE_MIN_MS = 5 * 60 * 1000;        // < 5 分 = テスト操作とみなして skip
const LESSON_AUTO_ARCHIVE_MIN_RESPONSES = 1;             // 0 回答 = 授業として成立してない
const LESSON_DAILY_STALE_HOURS = 4;                      // 最終回答から 4h 経過 = もう授業終了とみなす
const LESSON_DAILY_TRIGGER_HOUR = 23;                    // 23:00 JST に sweep

// ----- 内部 CRUD: lessons シートに対する row-level 操作 -----

// Why: SA-backed spreadsheet proxy の getSheetByName('lessons') は sheet 不存在でも
//   proxy オブジェクトを返してしまい、実際に getValues() するまで失敗を検出できない。
//   getSheets() で実存をチェックしてから handle を返す。
function __dbSheetExists_(spreadsheet, name) {
  try {
    const sheets = (spreadsheet.getSheets ? spreadsheet.getSheets() : []) || [];
    return sheets.some(s => {
      try { return s.getName && s.getName() === name; }
      catch (_) { return false; }
    });
  } catch (_) { return false; }
}

function __lessonsSheetExists_(spreadsheet) {
  return __dbSheetExists_(spreadsheet, 'lessons');
}

// 回答アーカイブシートの handle。__getLessonsSheet_ と同じ lazy bootstrap 方式。
//   createIfMissing の SpreadsheetApp 経路は呼び出し元に DB の編集権があるときだけ通る
//   (setup 済みテナントでは setupApp が ensure 済みなので、実運用でここが走ることは稀)。
function __getResponsesSheet_(opts) {
  const spreadsheet = openDatabase();
  if (!spreadsheet) return null;
  if (!__dbSheetExists_(spreadsheet, LESSON_RESPONSES_SHEET)) {
    if (!opts || !opts.createIfMissing) return null;
    try {
      // SA proxy の insertSheet (batchUpdate) を優先: 呼び出し元の権限に依らず動く。
      //   proxy でない native Spreadsheet (setup 経路) でも同名 API があるので共通で呼べる。
      const newSheet = spreadsheet.insertSheet(LESSON_RESPONSES_SHEET);
      if (newSheet && newSheet.appendRow) newSheet.appendRow(LESSON_RESPONSES_SHEET_HEADERS);
    } catch (createErr) {
      logError_('__getResponsesSheet_:create', createErr);
      return null;
    }
  }
  return spreadsheet.getSheetByName(LESSON_RESPONSES_SHEET) || null;
}

/**
 * 回答アーカイブを lesson_responses へ「連続した行範囲」として追記し、ポインタを返す。
 *
 * @param {string} lessonId
 * @param {number} phaseIndex
 * @param {Array} projectedRows - __projectBoardRowForExport_ 済みの行 (PII なし)
 * @returns {Object|null} { sheet, startRow, rowCount } / 書けなければ null
 *
 * 整合性の要:
 *   - 追記は 1 回の呼び出しで行う (SA proxy の appendRows = values:append は連続範囲を保証)。
 *   - proxy 経由でない sheet (test / native) には getLastRow → setValues で fallback する。
 *   - 再 capture 時に旧範囲の行は残る (孤児行)。ポインタが差し替わるだけなので無害で、
 *     読み出しは lessonId + phaseIndex を照合するため誤読もしない。
 */
function __writeArchiveRows_(lessonId, phaseIndex, projectedRows) {
  if (!projectedRows || projectedRows.length === 0) {
    return { sheet: null, startRow: -1, rowCount: 0 };
  }
  const sheet = __getResponsesSheet_({ createIfMissing: true });
  if (!sheet) return null;
  const values = projectedRows.map(r => [
    lessonId,
    phaseIndex,
    (typeof r.rowIndex === 'number') ? r.rowIndex : '',
    r.timestamp || '',
    r.class || '',
    (r.answer === null || r.answer === undefined) ? '' : r.answer,
    (r.reason === null || r.reason === undefined) ? '' : r.reason,
    (typeof r.numericX === 'number') ? r.numericX : '',
    (typeof r.numericY === 'number') ? r.numericY : ''
  ]);
  try {
    if (typeof sheet.appendRows === 'function') {
      const res = sheet.appendRows(values);
      if (!res || !(res.startRow > 0)) return null;
      return { sheet: LESSON_RESPONSES_SHEET, startRow: res.startRow, rowCount: values.length };
    }
    // fallback (test harness / native Sheet): 単一実行内なので getLastRow → setValues で足りる
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, values.length, values[0].length).setValues(values);
    return { sheet: LESSON_RESPONSES_SHEET, startRow, rowCount: values.length };
  } catch (err) {
    logError_('__writeArchiveRows_', err);
    return null;
  }
}

function __getLessonsSheet_(opts) {
  const spreadsheet = openDatabase();
  if (!spreadsheet) return null;
  const exists = __lessonsSheetExists_(spreadsheet);
  if (!exists) {
    if (!opts || !opts.createIfMissing) return null;
    // Lazy bootstrap: 既存 DB に lessons sheet が無いケース (Phase 1+2 デプロイ前に
    //   セットアップされた tenant) で初回 write 時に作成する。SpreadsheetApp 経由は
    //   呼び出し元 (admin) の権限で動くため SA token 不要。
    try {
      const dbId = typeof getCachedProperty === 'function' ? getCachedProperty('DATABASE_SPREADSHEET_ID') : null;
      if (!dbId) return null;
      const ss = SpreadsheetApp.openById(dbId);
      const newSheet = ss.insertSheet('lessons');
      newSheet.appendRow(LESSONS_SHEET_HEADERS);
    } catch (createErr) {
      logError_('__getLessonsSheet_:create', createErr);
      return null;
    }
  }
  return spreadsheet.getSheetByName('lessons') || null;
}

// シートの列インデックス map を生成 (header 行を読んで {col → index})。
function __lessonColumns_(sheet) {
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headerRow.forEach((h, i) => { map[String(h)] = i; });
  return map;
}

function __rowToLesson_(row, cols) {
  const lessonJsonRaw = row[cols.lessonJson];
  let lessonJson = {};
  let parseError = null;
  try { lessonJson = lessonJsonRaw ? JSON.parse(lessonJsonRaw) : {}; }
  catch (parseErr) {
    // Why parseError exposure: 旧コードは parse 失敗を {} で隠していたため、callers が
    //   その lesson を save-back すると corrupted-but-recoverable な lessonJson を {} で
    //   全上書きする data-loss 経路があった。parseError を露出させ、上位の __updateLessonRow_
    //   等が「parse 失敗 lesson は save しない」判定を入れられるようにする。
    logError_('__rowToLesson_:parseLessonJson', parseErr);
    parseError = parseErr.message || String(parseErr);
    lessonJson = {};
  }
  return {
    lessonId: row[cols.lessonId],
    userId: row[cols.userId],
    name: row[cols.name],
    state: row[cols.state],
    createdAt: row[cols.createdAt],
    startedAt: row[cols.startedAt] || null,
    endedAt: row[cols.endedAt] || null,
    schemaVersion: Number(row[cols.schemaVersion]) || LESSON_SCHEMA_VERSION,
    sizeBytes: Number(row[cols.sizeBytes]) || 0,
    etag: row[cols.etag] || null,
    lessonJson,
    parseError
  };
}

function __findLessonRowIndex_(sheet, lessonId) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return -1;
  // TextFinder で lessonId 列を exact match (TextFinder 未対応の test sandbox では fallback)。
  try {
    const finder = sheet.createTextFinder
      ? sheet.createTextFinder(lessonId).matchEntireCell(true)
      : null;
    if (finder) {
      const range = finder.findNext();
      if (range && range.getColumn() === 1) return range.getRow();
    }
  } catch (_) { /* fall through to linear scan */ }
  // Fallback linear scan (test 環境 or TextFinder API 異常時)。
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === lessonId) return i + 2;
  }
  return -1;
}

function __findLessonById_(lessonId) {
  try {
    const sheet = __getLessonsSheet_();
    if (!sheet) return null;
    const rowIndex = __findLessonRowIndex_(sheet, lessonId);
    if (rowIndex < 0) return null;
    const cols = __lessonColumns_(sheet);
    const row = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    return { rowIndex, lesson: __rowToLesson_(row, cols), cols, sheet };
  } catch (error) {
    logError_('__findLessonById_', error);
    return null;
  }
}

// lessonJson を 1 回 stringify して json / sizeBytes / etag を返す。
//   sizeBytes は char-based (Sheets cell 上限 50000 char に対する defensive cap)。
//   etag は ConfigService と同じ ISO+uuid 形式 (時刻ベース optimistic lock)。
function __serializeLessonJson_(lessonJson) {
  const json = JSON.stringify(lessonJson || {});
  const sizeBytes = json.length;
  if (sizeBytes > LESSON_JSON_MAX_BYTES) {
    throw new Error(`LESSON_TOO_LARGE: ${sizeBytes} chars > ${LESSON_JSON_MAX_BYTES} (Sheets cell 50000 char 上限対策)`);
  }
  const etag = new Date().toISOString() + '_' + Utilities.getUuid().replace(/-/g, '').slice(0, 12);
  return { json, sizeBytes, etag };
}

function __buildLessonRow_(record, cols, serialized) {
  const row = new Array(Object.keys(cols).length).fill('');
  row[cols.lessonId] = record.lessonId;
  row[cols.userId] = record.userId;
  row[cols.name] = record.name;
  row[cols.state] = record.state;
  row[cols.createdAt] = record.createdAt;
  row[cols.startedAt] = record.startedAt || '';
  row[cols.endedAt] = record.endedAt || '';
  row[cols.schemaVersion] = LESSON_SCHEMA_VERSION;
  row[cols.sizeBytes] = serialized.sizeBytes;
  row[cols.etag] = serialized.etag;
  row[cols.lessonJson] = serialized.json;
  return row;
}

function __createLessonRow_(record) {
  // write path: 既存 DB に lessons sheet が無ければ lazy bootstrap で作成。
  const sheet = __getLessonsSheet_({ createIfMissing: true });
  if (!sheet) return createErrorResponse('lessons sheet not initialized');
  const cols = __lessonColumns_(sheet);
  const serialized = __serializeLessonJson_(record.lessonJson);
  sheet.appendRow(__buildLessonRow_(record, cols, serialized));
  return { ...record, sizeBytes: serialized.sizeBytes, etag: serialized.etag, schemaVersion: LESSON_SCHEMA_VERSION };
}

// 全 lesson sheet write を serialize する ScriptLock wrapper。
//   GAS LockService は per-key lock 非対応なので script-wide が唯一の選択。
//   acquire / 5s timeout / finally-release の boilerplate を 1 箇所に集約し、
//   __updateLessonRow_ / __deleteLessonRow_ は body だけ lambda で渡す。
function __withLessonLock_(fn) {
  const lock = LockService.getScriptLock();
  let locked = false;
  try {
    locked = lock.tryLock(5000);
    if (!locked) {
      return createErrorResponse('別の処理が実行中です。少し待ってから再試行してください。', null, { error: 'LOCK_TIMEOUT' });
    }
    return fn();
  } finally {
    if (locked) {
      try { lock.releaseLock(); } catch (e) { console.warn('__withLessonLock_ releaseLock failed:', e.message); }
    }
  }
}

// 既存 row を patch。patch は {state?, name?, startedAt?, endedAt?, lessonJson?}。
//   batch write: 1 setValues 呼び出しで完結 (CLAUDE.md batch ルール)。
//   lessonJson 含む patch なら etag / sizeBytes も更新、そうでないなら lessonJson 列は触らず etag のみ。
//
//   Why LockService: 旧コードは read-modify-write を unlocked で行っていたため、
//   advanceLessonPhase / endLesson / updateLessonDraft の同時クリックで write が lost
//   する data-loss race があった。__withLessonLock_ で lessons sheet 全体を serialize する。
function __updateLessonRow_(lessonId, patch, expectedEtag) {
  return __withLessonLock_(() => {
    const found = __findLessonById_(lessonId);
    if (!found) return createErrorResponse('lesson not found', null, { error: 'LESSON_NOT_FOUND' });
    if (expectedEtag && found.lesson.etag && found.lesson.etag !== expectedEtag) {
      // Why lowercase: ConfigService.saveUserConfig / AdminApis.__applyPublishStateChange と
      //   同じ 'etag_mismatch' code に統一。frontend (AdminPanel.js.html) も 1 種類の
      //   コードのみ判定すれば良い (Error envelope audit recommendation #2)。
      return {
        success: false,
        error: 'etag_mismatch',
        message: '別タブで lesson が更新されています。再読込してください。',
        currentEtag: found.lesson.etag
      };
    }
    const { sheet, cols, rowIndex } = found;
    const lessonJson = patch.lessonJson !== undefined ? patch.lessonJson : found.lesson.lessonJson;
    const serialized = __serializeLessonJson_(lessonJson);

    const merged = {
      lessonId: found.lesson.lessonId,
      userId: found.lesson.userId,
      name: patch.name !== undefined ? patch.name : found.lesson.name,
      state: patch.state !== undefined ? patch.state : found.lesson.state,
      createdAt: found.lesson.createdAt,
      startedAt: patch.startedAt !== undefined ? patch.startedAt : found.lesson.startedAt,
      endedAt: patch.endedAt !== undefined ? patch.endedAt : found.lesson.endedAt
    };
    const row = __buildLessonRow_(merged, cols, serialized);
    sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);

    return {
      success: true,
      lesson: { ...found.lesson, ...patch, lessonJson, sizeBytes: serialized.sizeBytes, etag: serialized.etag }
    };
  });
}

function __listLessonsForUser_(userId, options = {}) {
  // Read path は sheet 不存在 / API エラーで silently empty を返す (= まだレッスン無し扱い)。
  try {
    const sheet = __getLessonsSheet_();
    if (!sheet) return [];
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];
    const cols = __lessonColumns_(sheet);
    const results = [];

    // owner-only list: TextFinder で userId 列だけスキャンして該当行 index を得る。
    //   全件 getValues() に比べて O(my-lessons) で済む。allUsers 経路は全件読みを維持。
    if (!options.allUsers && sheet.createTextFinder) {
      try {
        const finder = sheet.createTextFinder(userId).matchEntireCell(true);
        const matches = (finder.findAll ? finder.findAll() : []) || [];
        matches.forEach((range) => {
          if (range.getColumn() !== cols.userId + 1) return;
          const rowIdx = range.getRow();
          if (rowIdx <= 1) return;
          const row = sheet.getRange(rowIdx, 1, 1, sheet.getLastColumn()).getValues()[0];
          results.push(__rowToLesson_(row, cols));
        });
        results.sort((a, b) => String(b.createdAt).localeCompare(String(a.createdAt)));
        return results;
      } catch (_) { /* fall through to bulk scan */ }
    }

    // Fallback: 全件読み (test sandbox or TextFinder 異常時)。
    const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    for (let i = 0; i < data.length; i++) {
      if (!options.allUsers && data[i][cols.userId] !== userId) continue;
      results.push(__rowToLesson_(data[i], cols));
    }
    results.sort((a, b) => String(b.createdAt).localeCompare(String(a.createdAt)));
    return results;
  } catch (error) {
    logError_('__listLessonsForUser_', error);
    return [];
  }
}

function __deleteLessonRow_(lessonId) {
  // Why ScriptLock: deleteRow は行 index を shift する。同時に動いている __updateLessonRow_
  //   は事前に findLessonById_ で取得した rowIndex を保持しているので、その index に対する
  //   setValues が「別の lesson 行」を書き換える data-loss を起こす。__withLessonLock_ で
  //   __updateLessonRow_ と同じ ScriptLock を共有し serialize する。
  return __withLessonLock_(() => {
    try {
      const found = __findLessonById_(lessonId);
      if (!found) return createErrorResponse('LESSON_NOT_FOUND');
      // Why: SA proxy は deleteRow を持たない。admin は DB シートの editor 共有を受けているので
      //   SpreadsheetApp 経由で直接削除する。
      const dbId = typeof getCachedProperty === 'function' ? getCachedProperty('DATABASE_SPREADSHEET_ID') : null;
      if (!dbId) return createErrorResponse('DATABASE_NOT_CONFIGURED');
      const ss = SpreadsheetApp.openById(dbId);
      const sheet = ss.getSheetByName('lessons');
      if (!sheet) return createErrorResponse('LESSONS_SHEET_NOT_FOUND');
      sheet.deleteRow(found.rowIndex);
      return { success: true };
    } catch (error) {
      logError_('__deleteLessonRow_', error);
      return createErrorResponse(error && error.message, null, { error: 'DELETE_FAILED' });
    }
  });
}

// ----- Authorization: owner OR admin (write) / owner OR admin OR collaborator (read) -----
// Why admin allowed: listLessons / getKnownClassesForUser など他の lesson read ops は
//   admin に許可しているのに、advance / end / delete / updateDraft だけ admin 拒否すると
//   admin API 経由のサポート操作 (生徒の質問対応や授業データ修復) ができない。SSOT で揃える。
// Why collaborator (v2855+): ボード SS の editor として共有された共同教師は read 用途
//   (getLessonForReview 等) なら lesson を閲覧してよい。 write 用途 (advance/end/delete/
//   updateDraft) は引き続き owner / admin のみ — allowCollaborator=true で明示的に opt-in。

function __requireLessonOwner_(userId, lessonId, options) {
  const allowCollaborator = !!(options && options.allowCollaborator);
  const email = getCurrentEmail();
  if (!email) return { error: createAuthError() };

  const callerUser = findUserByEmail(email, { requestingUser: email });
  if (!callerUser) return { error: createUserNotFoundError() };

  const isAdmin = isAdministrator(email);
  let isCollaborator = false;
  if (!isAdmin && callerUser.userId !== userId) {
    if (allowCollaborator) {
      const targetUser = (typeof findUserById === 'function') ? findUserById(userId) : null;
      if (targetUser && typeof isBoardCollaborator === 'function'
          && isBoardCollaborator(targetUser, email)) {
        isCollaborator = true;
      }
    }
    if (!isCollaborator) {
      const resourceLabel = (options && options.resourceLabel) || 'lesson';
      return { error: createErrorResponse(`他ユーザーの ${resourceLabel} にはアクセスできません`) };
    }
  }

  if (!lessonId) return { callerUser, isAdmin, isCollaborator };

  const found = __findLessonById_(lessonId);
  if (!found) return { error: createErrorResponse('lesson が見つかりません') };
  if (found.lesson.userId !== userId) {
    return { error: createErrorResponse('lesson の所有者が一致しません') };
  }
  return { callerUser, found, isAdmin, isCollaborator };
}

// ----- Lesson テンプレート (Phase 1 は 1 種類固定) -----

// 児童・教師どちらが画面を見ても自然な命名にする。
//
// 設計指針 (2026-05-16 更新):
//   - phase 名は「教師の指導案語」ではなく「児童が黒板で見る語」に揃える。
//     例: 「本時」「導入」「終末」(指導案語) → 「めあて」「みんなで考える」「ふりかえり」(児童語)。
//   - 道徳科学習指導要領も「めあて」「振り返り」を板書に出すことを推奨 (光村図書 / 沖縄県教委)。
//   - このアプリは道徳に限らず全教科 (国語/算数/社会/外国語/総合等) で使われる前提なので、
//     "定番" テンプレは教科ニュートラルにする (旧「道徳・定番」→「授業の定番」)。
//   - 内部 key (doutoku-3phase) は backward-compat のため残置 (UI には出ない slug)。
//
// 出典: 田村学「カリキュラム・マネジメント」(探究3段階) / 光村図書 道徳 Q&A「めあてと振り返り」/
//       沖縄県教委 R4「めあて・振り返り」資料 / 文科省 特別の教科 道徳編。
const LESSON_TEMPLATES = {
  'doutoku-3phase': {  // 旧名残置 (= "standard-3phase" 相当)。tests / 既存 lessonJson との互換のため。
    label: '授業の定番（3段階）',
    description: 'めあて → みんなで考える → ふりかえり',
    phases: [
      { name: 'めあて', formTemplate: 'numberline', question: 'いまの自分の考えは？' },
      { name: 'みんなで考える', formTemplate: 'matrix', question: 'なぜそう思った？' },
      { name: 'ふりかえり', formTemplate: 'numberline', question: '話し合って、いまの考えは？' }
    ]
  },
  'kid-3phase': {
    label: '低学年向け',
    description: 'いまの考え → みんなで話す → これからの考え',
    phases: [
      { name: 'いまの考え', formTemplate: 'numberline', question: 'いまの考えは？' },
      { name: 'みんなで話す', formTemplate: 'matrix', question: 'どうしてそう思った？' },
      { name: 'これからの考え', formTemplate: 'numberline', question: 'はなしあって、いまの考えは？' }
    ]
  },
  'inquiry-3phase': {
    label: '探究（田村モデル）',
    description: '出会う → ふかめる → つなげる',
    phases: [
      { name: '出会う', formTemplate: 'pie', question: 'まずどっちだと思う？' },
      { name: 'ふかめる', formTemplate: 'matrix', question: '理由と確信度を教えて' },
      { name: 'つなげる', formTemplate: 'numberline', question: '自分の答えはどこに着地した？' }
    ]
  },
  'before-after-2phase': {
    label: '議論前後（2段階）',
    description: '議論のまえ → 議論のあと',
    phases: [
      { name: '議論のまえ', formTemplate: 'numberline', question: 'いまのあなたの立場は？' },
      { name: '議論のあと', formTemplate: 'numberline', question: '議論をしたあと、いまの立場は？' }
    ]
  },
  // 「考え、議論する道徳」向け。他のテンプレと違い Form を作らず、画面から直接投稿する
  //   (inputMode: 'native')。フェーズが児童画面の権能を切り替えるのがこのテンプレの本体。
  //
  // Why 5 段階か: 「自分の考えをもつ → 他者の考えに出会う → 議論する → 問い直す → 言語化する」
  //   が道徳科の学習過程 (文科省 特別の教科 道徳編) だから。可視化はこの過程を支える手段であり、
  //   意見を集めること自体は目的にしない。
  // Why 縦軸が「迷い」か: 立場の正誤を軸にすると多数派が正解に見える。確信度を縦に取ると
  //   「立場は同じだが迷いが増えた」という深まりも位置として現れ、かつ優劣がつかない。
  'dialogue-reconsider-5phase': {
    label: '考え、議論する道徳（5段階）',
    description: '考える → 出会う → 議論する → もう一度考える → ふりかえる',
    inputMode: 'native',
    phases: [
      {
        name: '考える', formTemplate: 'matrix', screenRole: LESSON_SCREEN_ROLES.INPUT,
        question: 'いまのあなたの考えは、どこにありますか？'
      },
      {
        name: '出会う', formTemplate: 'matrix', screenRole: LESSON_SCREEN_ROLES.BROWSE,
        question: '友達はどう考えた？ 理由を読んでみよう'
      },
      {
        name: '議論する', formTemplate: 'matrix', screenRole: LESSON_SCREEN_ROLES.DISCUSS,
        question: '画面をとじて、話し合おう'
      },
      {
        name: 'もう一度考える', formTemplate: 'matrix', screenRole: LESSON_SCREEN_ROLES.REINPUT,
        question: '話し合ったいま、あなたはどこに立ちますか？'
      },
      {
        name: 'ふりかえる', formTemplate: 'matrix', screenRole: LESSON_SCREEN_ROLES.REFLECT,
        question: '自分の考えは、どう変わった／変わらなかった？'
      }
    ]
  }
};

// phase が native 入力かを判定する。phase 個別指定 > lesson 全体の順で解決する。
function __isNativePhase_(phase, lessonJson) {
  if (phase && phase.inputMode) return phase.inputMode === 'native';
  return Boolean(lessonJson && lessonJson.inputMode === 'native');
}

// phase の画面権能。未指定なら 'browse' (= 見るだけ) に倒す。
//   Why 既定を browse にするか: 権能の指定漏れが「誰でも書ける」に倒れると、
//   議論中に投稿できてしまう等、授業の構造が静かに壊れる。安全側は「書けない」。
function __phaseScreenRole_(phase) {
  const role = phase && phase.screenRole;
  const known = Object.keys(LESSON_SCREEN_ROLES).map(k => LESSON_SCREEN_ROLES[k]);
  return known.indexOf(role) >= 0 ? role : LESSON_SCREEN_ROLES.BROWSE;
}

// public: フロントから利用可能なテンプレ一覧を返す (creation dropdown 用)
function listLessonTemplates() {
  return createSuccessResponse('listed', {
    templates: Object.keys(LESSON_TEMPLATES).map((key) => ({
      key,
      label: LESSON_TEMPLATES[key].label,
      description: LESSON_TEMPLATES[key].description,
      phaseCount: LESSON_TEMPLATES[key].phases.length
    }))
  });
}

// ----- 公開 API (dispatchAdminOperation 経由) -----

function createLessonDraft(userId, name, template) {
  try {
    const auth = __requireLessonOwner_(userId, null);
    if (auth.error) return auth.error;

    const templateKey = template || 'doutoku-3phase';
    const tpl = LESSON_TEMPLATES[templateKey];
    if (!tpl) return createErrorResponse(`未知のテンプレート: ${templateKey}`);

    const lessonId = 'lesson_' + Utilities.getUuid().slice(0, 12);
    const now = new Date().toISOString();
    const lessonJson = {
      template: templateKey,
      classes: [],
      // 'native' なら Form を作らず画面から直接投稿する (授業モード)。
      inputMode: tpl.inputMode || 'form',
      phases: tpl.phases.map(p => ({
        name: p.name,
        formTemplate: p.formTemplate,
        question: p.question,
        // 画面の権能 (input/browse/discuss/reinput/reflect)。native テンプレのみ持つ。
        screenRole: p.screenRole || '',
        // Form 生成は startLesson で行うので、draft 時点では空。
        formId: '',
        formUrl: '',
        spreadsheetId: '',
        sheetName: '',
        columnMapping: {},
        displaySettings: {}
      })),
      profileTransitions: [],
      snapshots: [],
      meta: { schemaVersion: LESSON_SCHEMA_VERSION }
    };

    const record = {
      lessonId, userId,
      name: String(name || '').slice(0, 100) || '新しい授業',
      state: 'draft',
      createdAt: now, startedAt: null, endedAt: null,
      lessonJson
    };
    const written = __createLessonRow_(record);
    return createSuccessResponse('lesson draft を作成しました', { lesson: written });
  } catch (error) {
    logError_('createLessonDraft', error);
    return createExceptionResponse(error);
  }
}

// fieldPath ('name', 'classes', 'phases[1].question' 等) を lessonJson にセット。
//   bracket / dot 表記の両方を受け付ける。
function __setByPath_(obj, fieldPath, value) {
  const parts = String(fieldPath).match(/[^.[\]]+/g);
  if (!parts || parts.length === 0) return;
  let cur = obj;
  for (let i = 0; i < parts.length - 1; i++) {
    const key = parts[i];
    const nextKey = parts[i + 1];
    const nextIsIndex = /^\d+$/.test(nextKey);
    if (cur[key] == null || typeof cur[key] !== 'object') {
      cur[key] = nextIsIndex ? [] : {};
    }
    cur = cur[key];
  }
  cur[parts[parts.length - 1]] = value;
}

function updateLessonDraft(userId, lessonId, fieldPath, value, expectedEtag) {
  try {
    const auth = __requireLessonOwner_(userId, lessonId);
    if (auth.error) return auth.error;
    const { found } = auth;

    if (found.lesson.state !== 'draft') {
      return createErrorResponse('FORBIDDEN_STATE: draft 状態でのみ編集できます');
    }

    // top-level の name 列はシート列としても保持しているので別ルートで反映。
    const isNameField = fieldPath === 'name';
    const newName = isNameField
      ? (String(value || '').slice(0, 100) || '新しい授業')
      : found.lesson.name;
    const lessonJson = isNameField
      ? { ...found.lesson.lessonJson }
      : deepClone(found.lesson.lessonJson || {});
    if (!isNameField) __setByPath_(lessonJson, fieldPath, value);

    // Why: IME 入力中など、同じ値が複数回送られてくるケースがある。
    //   既存値と一致するなら sheet 書き込みを skip して I/O と etag 進行を抑制。
    const beforeJson = JSON.stringify(found.lesson.lessonJson || {});
    const afterJson = JSON.stringify(lessonJson);
    if (beforeJson === afterJson && (!isNameField || newName === found.lesson.name)) {
      return createSuccessResponse('unchanged', { lesson: found.lesson });
    }

    const patch = isNameField ? { lessonJson, name: newName } : { lessonJson };
    const result = __updateLessonRow_(lessonId, patch, expectedEtag);
    if (!result.success) return createErrorResponse(result.message || result.error, null, { error: result.error });
    return createSuccessResponse('updated', { lesson: result.lesson });
  } catch (error) {
    logError_('updateLessonDraft', error);
    return createExceptionResponse(error);
  }
}

/**
 * 回答ボード (mode=view) の phase pill 用に、実行中の授業の最小情報だけを返す。
 *
 * Why 専用 API: 生徒は 5 秒ごとに getPublishedSheetData を叩く。そこへ授業情報を
 *   相乗りさせると、生徒に不要なデータを配りながら lessons シートを 5 秒ごとに読む
 *   ことになる。教師が明示的に呼ぶ経路として切り出す。
 * Why owner 限定: 切替は授業の進行そのもの。閲覧者に見せる情報ではない。
 *
 * @returns {Object} 実行中の授業が無ければ data フィールドを付けない (これは正常系)
 */
function getActiveLessonNav(targetUserId) {
  try {
    const userId = targetUserId || null;
    if (!userId) return createErrorResponse('userId is required');
    // lessonId を渡さない = 「この user の lesson を触る権限があるか」だけを見る。
    const auth = __requireLessonOwner_(userId, null);
    if (auth.error) return auth.error;

    const listed = listLessons(userId);
    if (!listed || !listed.success) return listed;
    const lessons = (listed.data && listed.data.lessons) || [];
    const active = lessons.find(l => l && l.state === 'active');
    if (!active) return createSuccessResponse('no active lesson', null);

    const found = __findLessonById_(active.lessonId);
    if (!found) return createSuccessResponse('no active lesson', null);
    const lessonJson = found.lesson.lessonJson || {};
    const phases = Array.isArray(lessonJson.phases) ? lessonJson.phases : [];

    return createSuccessResponse('loaded', {
      lessonId: active.lessonId,
      name: found.lesson.name || '',
      activePhaseIndex: __activePhaseIndex_(lessonJson),
      phases: phases.map((p, i) => ({
        index: i,
        name: (p && p.name) || ('フェーズ ' + (i + 1)),
        formTemplate: (p && p.formTemplate) || ''
      }))
    });
  } catch (error) {
    logError_('getActiveLessonNav', error);
    return createExceptionResponse(error);
  }
}

function listLessons(userId) {
  try {
    const access = __requireLessonOwner_(userId, null, { resourceLabel: 'lesson 一覧' });
    if (access.error) return access.error;

    const lessons = __listLessonsForUser_(userId);
    // 軽量化: 一覧では heavy な snapshots / profileTransitions は除外。
    const summaries = lessons.map(l => ({
      lessonId: l.lessonId,
      name: l.name,
      state: l.state,
      createdAt: l.createdAt,
      startedAt: l.startedAt,
      endedAt: l.endedAt,
      classes: (l.lessonJson && l.lessonJson.classes) || [],
      phaseCount: (l.lessonJson && Array.isArray(l.lessonJson.phases)) ? l.lessonJson.phases.length : 0,
      etag: l.etag
    }));
    return createSuccessResponse('listed', { lessons: summaries });
  } catch (error) {
    logError_('listLessons', error);
    return createExceptionResponse(error);
  }
}

/**
 * 過去レッスンの classes フィールドを集約して unique class 名を返す。
 * Wizard UI が「以前使ったクラス」候補を出すために使う (= テキスト入力を最小化)。
 * createdAt 降順で見て、新しい順で重複排除 → 直近で使ったクラスが先頭に並ぶ。
 */
function getKnownClassesForUser(userId) {
  try {
    const access = __requireLessonOwner_(userId, null, { resourceLabel: 'class 一覧' });
    if (access.error) return access.error;

    const lessons = __listLessonsForUser_(userId); // createdAt 降順済
    const seen = new Set();
    const classes = [];
    lessons.forEach((lesson) => {
      const list = (lesson.lessonJson && lesson.lessonJson.classes) || [];
      list.forEach((c) => {
        const key = String(c || '').trim();
        if (!key || seen.has(key)) return;
        seen.add(key);
        classes.push(key);
      });
    });
    return createSuccessResponse('listed', { classes });
  } catch (error) {
    logError_('getKnownClassesForUser', error);
    return createExceptionResponse(error);
  }
}

function getLessonForReview(userId, lessonId) {
  try {
    // read 用途なので collaborator (ボード SS editor) にも許可 (v2855+)。
    const auth = __requireLessonOwner_(userId, lessonId, { allowCollaborator: true });
    if (auth.error) return auth.error;
    // active / completed どちらも review 可。draft は wizard で開く方が自然。
    // アーカイブ行はポインタ経由で読み戻す (クライアントは rows が埋まった snapshot を期待する)。
    __hydrateLessonSnapshots_(auth.found.lesson);
    return createSuccessResponse('loaded', { lesson: auth.found.lesson });
  } catch (error) {
    logError_('getLessonForReview', error);
    return createExceptionResponse(error);
  }
}

/**
 * snapshot のポインタ {sheet, startRow, rowCount} から lesson_responses の行を読み戻し、
 * snapshots[].rows を埋める (in-place)。旧形式 (rows 同居) の snapshot は触らない。
 *
 * 読み取り量はポインタの範囲だけ (1 フェーズ ≈ 数百セル)。シート全件は読まない。
 * Why 照合ガード: 行の物理削除やシートの手編集でポインタがずれた場合、範囲読みは
 *   「別の授業の行」を返しうる。lessonId + phaseIndex を各行で照合し、他授業の回答を
 *   振り返りに混ぜる事故を構造的に塞ぐ (ずれた行は静かに落ち、rows が減るだけ)。
 */
function __hydrateLessonSnapshots_(lesson) {
  const lessonJson = lesson && lesson.lessonJson;
  const snaps = (lessonJson && Array.isArray(lessonJson.snapshots)) ? lessonJson.snapshots : [];
  const targets = snaps.filter(sn => sn && sn.sheet && sn.startRow > 0 && sn.rowCount > 0
    && (!Array.isArray(sn.rows) || sn.rows.length === 0));
  if (targets.length === 0) return;

  const spreadsheet = openDatabase();
  if (!spreadsheet) return;
  const width = LESSON_RESPONSES_SHEET_HEADERS.length;
  const toNum = (v) => (v === '' || v === null || v === undefined) ? null : Number(v);

  for (const sn of targets) {
    try {
      const sheet = spreadsheet.getSheetByName(sn.sheet);
      if (!sheet) { sn.reason = 'ARCHIVE_SHEET_MISSING:' + sn.sheet; continue; }
      const values = sheet.getRange(sn.startRow, 1, sn.rowCount, width).getValues() || [];
      sn.rows = values
        .filter(v => String(v[0]) === String(lesson.lessonId) && Number(v[1]) === Number(sn.phaseIndex))
        .map(v => ({
          rowIndex: toNum(v[2]),
          timestamp: String(v[3] || ''),
          class: String(v[4] || ''),
          answer: v[5],
          reason: v[6],
          numericX: toNum(v[7]),
          numericY: toNum(v[8])
        }));
      if (sn.rows.length !== sn.rowCount) {
        sn.reason = 'ARCHIVE_POINTER_DRIFT:' + sn.rows.length + '/' + sn.rowCount;
      }
    } catch (err) {
      logError_('__hydrateLessonSnapshots_', err);
      sn.reason = 'ARCHIVE_READ_FAILED';
      sn.rows = [];
    }
  }
}

/**
 * 過去レッスンを「テンプレ的に複製」して新規 draft を作る。
 *   - phases の構造 (name / formTemplate / question / templateOptions) は引き継ぐ
 *   - Form / SS / classes / snapshots / profileTransitions は strip (新規授業 = 別 Form)
 *   - 旧版のクラス構成は明示的に引き継ぎたいケースもあるので options.copyClasses=true で復元
 *   - name は元 name + " (コピー)"。ユーザーは Step 1 で書き換えられる。
 */
function duplicateLesson(userId, sourceLessonId, options) {
  try {
    const auth = __requireLessonOwner_(userId, sourceLessonId);
    if (auth.error) return auth.error;
    const src = auth.found.lesson;
    const opts = options || {};

    const srcPhases = (src.lessonJson && src.lessonJson.phases) || [];
    const newPhases = srcPhases.map((p) => ({
      name: p.name,
      formTemplate: p.formTemplate,
      question: p.question,
      // templateOptions (軸ラベル / 選択肢) は教師の意図そのものなので必ず引き継ぐ
      templateOptions: p.templateOptions ? deepClone(p.templateOptions) : {},
      // 以下は「新しい Form を作る」ために必ず空にする
      formId: '', formUrl: '', spreadsheetId: '', sheetName: '',
      columnMapping: {}, displaySettings: {}
    }));

    const newLessonId = 'lesson_' + Utilities.getUuid().slice(0, 12);
    const now = new Date().toISOString();
    const lessonJson = {
      template: src.lessonJson && src.lessonJson.template || 'doutoku-3phase',
      classes: opts.copyClasses ? ((src.lessonJson && src.lessonJson.classes) || []).slice() : [],
      phases: newPhases,
      profileTransitions: [],
      snapshots: [],
      meta: { schemaVersion: LESSON_SCHEMA_VERSION, duplicatedFrom: sourceLessonId }
    };
    const baseName = String(src.name || '新しい授業').slice(0, 80);
    const record = {
      lessonId: newLessonId, userId,
      name: baseName + ' (コピー)',
      state: 'draft',
      createdAt: now, startedAt: null, endedAt: null,
      lessonJson
    };
    // Why: __createLessonRow_ は record そのもの (etag/sizeBytes 付加) を返し、success フラグは無い。
    //   失敗時は {success:false, message:...} を返すケースのみ。createLessonDraft と同じ慣用。
    const written = __createLessonRow_(record);
    if (written && written.success === false) {
      return createErrorResponse(written.message || 'lesson 作成失敗');
    }
    return createSuccessResponse('duplicated', { lesson: written });
  } catch (error) {
    logError_('duplicateLesson', error);
    return createExceptionResponse(error);
  }
}


// formUrl ('https://docs.google.com/forms/d/e/<published>/viewform') から formId 相当を抽出。
//   published URL の id は canonical formId と一致しないが、UI で開くリンクとしては十分。
//   完全な canonical formId が必要なら FormApp.openByUrl が必要だが、import 時には不要。
function __extractFormPublishedId_(formUrl) {
  if (!formUrl || typeof formUrl !== 'string') return '';
  const m = formUrl.match(/\/forms\/d\/e?\/?([^/]+)\//);
  return m ? m[1] : '';
}

// ボード行を「実践報告書 / 過去授業 archive 用」の slim row に整形する。
//   個人特定可能フィールド (name / email / emailHash / reactions / highlight / opinion / id) は除外。
//   `includeName: true` で名前列も残せる (校内資料用)。
//   exportBoardData (AdminApis.js) と snapshot capture (__captureSnapshot_ / import) で共用。
function __projectBoardRowForExport_(row, options) {
  const includeName = options && options.includeName === true;
  const out = {
    rowIndex: row.rowIndex,
    timestamp: row.formattedTimestamp || row.timestamp || '',
    class: row.class || '',
    answer: row.answer,
    reason: row.reason,
    numericX: row.numericX,
    numericY: row.numericY
  };
  if (includeName) out.name = row.name || '';
  return out;
}


function __extractClassesFromSnapshots_(snapshots) {
  const seen = new Set();
  (snapshots || []).forEach((s) => {
    (s.rows || []).forEach((r) => {
      if (r && r.class) seen.add(String(r.class));
    });
  });
  return Array.from(seen).sort();
}




function deleteLesson(userId, lessonId) {
  try {
    const auth = __requireLessonOwner_(userId, lessonId);
    if (auth.error) return auth.error;

    if (auth.found.lesson.state === 'active') {
      return createErrorResponse('実行中の lesson は削除できません。先に「終了」してください。');
    }
    const result = __deleteLessonRow_(lessonId);
    if (!result.success) return createErrorResponse(result.error || 'delete failed');
    return createSuccessResponse('deleted', { lessonId });
  } catch (error) {
    logError_('deleteLesson', error);
    return createExceptionResponse(error);
  }
}

// ----- Lifecycle: startLesson / advanceLessonPhase / endLesson -----

/**
 * phase の Form を開く。開けなければ null。
 *
 * Why 2 経路あるか: profiles から移行した授業 (meta.importedFromProfiles) は
 *   formId に**公開 URL 側の ID** (`1FAIpQLS...`) が入っている。これは
 *   FormApp.openById が受け付けない (編集用 ID とは別物) ため、移行由来の授業では
 *   フェーズごとの Form 開閉が最初から効いていなかった。
 *   回答スプレッドシートは編集用 URL を知っている (getFormUrl) ので、そこから復旧する。
 *
 * 復旧できたら phase.formId を編集用 ID に**書き直す** (呼び出し元が lessonJson を
 * 保存すれば次回から 1 発で開く)。
 */
function __openLessonForm_(phase) {
  if (typeof FormApp === 'undefined') return null;
  const formId = phase && phase.formId;
  if (formId && FormApp.openById) {
    try {
      return FormApp.openById(formId);
    } catch (_) {
      // 公開 URL 側の ID だった可能性がある。下の SS 経由へ。
    }
  }
  const ssId = phase && phase.spreadsheetId;
  if (!ssId || typeof SpreadsheetApp === 'undefined' || !FormApp.openByUrl) return null;
  try {
    const editUrl = SpreadsheetApp.openById(ssId).getFormUrl();
    if (!editUrl) return null;
    const form = FormApp.openByUrl(editUrl);
    // 復旧した編集用 ID を phase に書き戻す (次回以降は 1 発で開く)。
    if (form && form.getId) phase.formId = form.getId();
    return form;
  } catch (error) {
    logError_('__openLessonForm_', error);
    return null;
  }
}

// Form の回答受付を on/off。テスト sandbox では FormApp が無いので silently fall through。
//   phase オブジェクトを渡すと、移行由来の授業でも SS 経由で Form を解決する。
function __setFormAcceptingResponses_(phaseOrFormId, accepting) {
  const phase = (phaseOrFormId && typeof phaseOrFormId === 'object')
    ? phaseOrFormId
    : { formId: phaseOrFormId };
  if (!phase.formId && !phase.spreadsheetId) return false;
  try {
    const form = __openLessonForm_(phase);
    if (!form) return false;
    form.setAcceptingResponses(Boolean(accepting));
    return true;
  } catch (error) {
    logError_('__setFormAcceptingResponses_', error);
    return false;
  }
}

// startLesson / advanceLessonPhase の両方で使う「user config に phase を適用する」 patch shape。
//   activeLessonId は後段 (autosave / publishApp) が lesson 駆動を判別するためのマーカー。
// formTemplate ('pie' / 'board' 等) → boardMode マップ。
//   numberline / matrix は linearScale 列を含むので auto 検出で動くが、
//   pie / board は同じ多肢選択データ形なので明示指定しないと判別不可。
function __templateToBoardMode_(formTemplate) {
  if (formTemplate === 'pie') return 'pie';
  if (formTemplate === 'numberline') return 'numberline';
  if (formTemplate === 'matrix') return 'matrix';
  if (formTemplate === 'board') return 'board';
  return 'auto';
}

/**
 * native phase 用の回答シートを用意する。
 *
 * 1 授業 = 1 スプレッドシート、1 フェーズ = 1 シート。
 * Why フェーズごとにシートを分けるか: 「最初の考え」と「いまの考え」は別の記録であって
 *   同じ列の上書きではない。分けておけば片方が消えることがなく、フェーズ内では
 *   1 児童 1 行が保証されるので突合も単純になる。
 *
 * @returns {{spreadsheetId:string, sheetName:string}|null}
 */
function __ensureNativeAnswerSheet_(lessonJson, phaseIdx, lessonName, ownerEmail) {
  // スプレッドシートは授業に 1 つ。既に作ってあれば使い回す (resume 時の二重作成防止)。
  let ssId = lessonJson.nativeSpreadsheetId || '';
  let ss = null;
  if (ssId) {
    try {
      ss = SpreadsheetApp.openById(ssId);
    } catch (openErr) {
      logError_('__ensureNativeAnswerSheet_:open', openErr);
      ss = null;
      ssId = '';
    }
  }
  if (!ss) {
    ss = SpreadsheetApp.create(`「${lessonName || '授業'}」の回答`);
    ssId = ss.getId();
    lessonJson.nativeSpreadsheetId = ssId;
    // 児童の投稿は SA pool 経由で書かれる (児童は SS への直接権限を持たない)。
    //   Form 回答シートと同じ共有既定を当てないと、投稿が権限エラーで落ちる。
    try {
      applySpreadsheetSharingDefaults(ssId);
    } catch (shareErr) {
      logError_('__ensureNativeAnswerSheet_:share', shareErr);
    }
    // ボードの所有者 (教師) を editor に加える。
    //   Why 必須か: SpreadsheetApp.create は「実行した人」の Drive に作る。管理者が
    //   教師の代わりに授業を開始すると、教師が所有しない SS がボードの参照先になる。
    //   owner は自分のボードを openById で直接開く経路を通る (SA を使わない) ので、
    //   編集権が無いと教師自身のボードが開けなくなる。作成者本人なら no-op。
    if (ownerEmail) {
      try {
        ss.addEditor(ownerEmail);
      } catch (ownerErr) {
        logError_('__ensureNativeAnswerSheet_:owner', ownerErr);
      }
    }
  }

  const sheetName = `phase${phaseIdx + 1}`;
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    // 1 枚目は create 時の既定シートを流用して名前を変える (空シートを残さない)。
    const sheets = ss.getSheets();
    if (sheets.length === 1 && sheets[0].getLastRow() === 0) {
      sheet = sheets[0].setName(sheetName);
    } else {
      sheet = ss.insertSheet(sheetName);
    }
    sheet.appendRow(LESSON_NATIVE_SHEET_HEADERS);
  }
  return { spreadsheetId: ssId, sheetName };
}

/**
 * native phase が「どのシートのデータを見せるか」を決める。
 *
 * 入力フェーズ (考える / もう一度考える) は自分のシート。
 * 投稿を受け付けないフェーズ (出会う / 議論する / ふりかえる) は、直前の入力フェーズの
 * シートを見る。「出会う」で見るのは「考える」で集まった考えだから。
 *
 * @returns {Object|null} 参照先の phase / 自分のシートでよければ null
 */
function __nativeDataSourcePhase_(phase, lessonJson) {
  const phases = (lessonJson && Array.isArray(lessonJson.phases)) ? lessonJson.phases : [];
  if (LESSON_INPUT_ROLES.indexOf(__phaseScreenRole_(phase)) >= 0) return null;
  const idx = phases.indexOf(phase);
  if (idx < 0) return null;
  for (let i = idx - 1; i >= 0; i--) {
    const p = phases[i];
    if (LESSON_INPUT_ROLES.indexOf(__phaseScreenRole_(p)) >= 0 && p.sheetName) return p;
  }
  return null;  // まだ入力フェーズを通っていない (授業の冒頭) = 空でよい
}

/**
 * 授業全体の座標系を決める phase (= 最初の入力フェーズ)。
 *
 * Why 授業で 1 つに揃えるか: ● 最初 → ★ いま の比較は、両方が同じ軸の上にあって
 *   初めて意味を持つ。「考える」と「もう一度考える」で軸ラベルが違うと、
 *   位置の変化が何を表すのか誰にも説明できなくなる。教師が軸を設定するのは 1 回でよい。
 */
function __nativeAxisPhase_(lessonJson) {
  const phases = (lessonJson && Array.isArray(lessonJson.phases)) ? lessonJson.phases : [];
  for (let i = 0; i < phases.length; i++) {
    if (LESSON_INPUT_ROLES.indexOf(__phaseScreenRole_(phases[i])) >= 0) return phases[i];
  }
  return null;
}

function __buildPhaseConfigPatch_(phase, lessonJson, lessonId) {
  // displaySettings が phase に明示指定されていればそれを優先、
  //   無ければ formTemplate から決定。templateOptions から axis ラベルも反映。
  const baseDisplay = phase.displaySettings || {};
  const isNative = __isNativePhase_(phase, lessonJson);
  // native の非入力フェーズは直前の入力フェーズのデータを表示する。
  const nativeSource = isNative ? __nativeDataSourcePhase_(phase, lessonJson) : null;
  // 軸ラベルは授業を通して 1 つの座標系に揃える (最初の入力フェーズのもの)。
  //   phase 個別に設定されていればそれを優先する (教師が意図的に変えた場合)。
  const axisPhase = isNative ? __nativeAxisPhase_(lessonJson) : null;
  const ownOpts = phase.templateOptions || {};
  const hasOwnAxis = Boolean(ownOpts.xLow || ownOpts.xHigh || ownOpts.yLow || ownOpts.yHigh);
  const opts = (!hasOwnAxis && axisPhase && axisPhase.templateOptions) || ownOpts;
  const displaySettings = Object.assign({}, baseDisplay);
  if (!displaySettings.boardMode) {
    displaySettings.boardMode = __templateToBoardMode_(phase.formTemplate);
  }
  // 軸ラベルは config の **トップレベル** (xAxisLabels / yAxisLabels) に格納する。
  //   これが canonical location: DataApis.getPublishedSheetData の axisConfig、
  //   ConfigService.validateAndSanitizeConfig、 sanitizeProfiles はいずれもトップレベルを参照する。
  //   旧実装は displaySettings.xAxisLabels (nested) に入れていたが、 sanitizeDisplaySettings の
  //   allowlist (showNames/showReactions/theme/pageSize/boardMode) で保存往復ごとに silently
  //   落ち、 board frontend に届かなかった (M1/M2 軸ラベル消失バグ)。
  //   nested で明示指定された legacy phase 互換のため baseDisplay からも拾う。
  let xAxisLabels = baseDisplay.xAxisLabels || null;
  let yAxisLabels = baseDisplay.yAxisLabels || null;
  if (phase.formTemplate === 'numberline' && (opts.lowLabel || opts.highLabel)) {
    xAxisLabels = { min: opts.lowLabel || '', max: opts.highLabel || '' };
  }
  if (phase.formTemplate === 'matrix') {
    if (opts.xLow || opts.xHigh) xAxisLabels = { min: opts.xLow || '', max: opts.xHigh || '' };
    if (opts.yLow || opts.yHigh) yAxisLabels = { min: opts.yLow || '', max: opts.yHigh || '' };
  }
  // nested に残すと sanitizeDisplaySettings が落とすだけで mismatch の温床になるため除去。
  delete displaySettings.xAxisLabels;
  delete displaySettings.yAxisLabels;
  // 揺らぎ追跡: phase or lesson レベルの allowResubmit を config.allowResubmit に反映する。
  //   これで board frontend (page.viz.js) が ghost dot / swing trace を描画する。
  const phaseAllowResubmit = opts.allowResubmit;
  const lessonAllowResubmit = Boolean(lessonJson && lessonJson.allowResubmit);
  const allowResubmit = phaseAllowResubmit != null ? Boolean(phaseAllowResubmit) : lessonAllowResubmit;
  const patch = {
    formUrl: phase.formUrl,
    spreadsheetId: phase.spreadsheetId,
    sheetName: phase.sheetName,
    formTitle: (lessonJson && lessonJson.name) || phase.name,
    columnMapping: phase.columnMapping || {},
    displaySettings,
    allowResubmit,
    activeLessonId: lessonId
  };
  // native phase はシートを我々が作っているので列の意味が確定している。
  //   推定 (inferColumnRoles) を通さず固定 index を渡す = 推定ミスが起きない。
  if (__isNativePhase_(phase, lessonJson)) {
    patch.columnMapping = __nativeColumnMapping_(phase.formTemplate);
    // 児童は画面から投稿するので、Form へ誘導する導線は出さない。
    patch.formUrl = '';
    // 投稿を受け付けないフェーズ (出会う/議論する/ふりかえる) は自分のシートが空なので、
    //   直前の入力フェーズのシートを指す。これがないと「出会う」で分布が空になる
    //   = 他者の考えに出会えず、授業の中心が成立しない。
    if (nativeSource) {
      patch.spreadsheetId = nativeSource.spreadsheetId;
      patch.sheetName = nativeSource.sheetName;
    }
  }
  if (xAxisLabels) patch.xAxisLabels = xAxisLabels;
  if (yAxisLabels) patch.yAxisLabels = yAxisLabels;
  return patch;
}

/**
 * 現在の board データを凍結して snapshot を作る。
 *
 * 回答本文は lesson_responses シートへ 1 件 1 行で追記し、snapshot には範囲ポインタ
 * {sheet, startRow, rowCount} だけを残す。lessonJson は定義 + ポインタのみ (~4KB) になり、
 * Sheets 1 セル 50,000 字上限に触れない = 本文の切り詰めが不要になった。
 *
 * 行は __projectBoardRowForExport_ で PII (name / email / reactions) を落としてから書く。
 * Why: 旧実装は live capture だけ生 row を凍結しており、import 経路 (PII 除去済み) と
 *   非対称だった。アーカイブに児童の氏名・メールを持つ理由はない。
 *
 * 劣化系: fetch 失敗 / アーカイブ書込失敗時は rows 無しの snapshot に reason を残して返す
 * (授業の進行を止めない)。元データは先生の spreadsheet に残っているので、後から
 * lesson.recaptureArchive で焼き直せる。
 */
function __captureSnapshot_(userId, lessonJson, phaseIdx, lessonId) {
  const phases = (lessonJson && lessonJson.phases) || [];
  const phase = phases[phaseIdx] || {};
  const baseSnapshot = {
    phaseIndex: phaseIdx,
    phaseName: phase.name || '',
    capturedAt: new Date().toISOString(),
    boardMode: (phase.displaySettings && phase.displaySettings.boardMode) || 'auto',
    columnMapping: phase.columnMapping || {},
    displaySettings: phase.displaySettings || {},
    rows: [],
    rowCount: 0,
    truncated: false
  };

  let rawRows;
  try {
    const fetched = getPublishedSheetData(null, 'newest', true, userId);
    if (!fetched || fetched.success === false || !Array.isArray(fetched.data)) {
      baseSnapshot.reason = 'CAPTURE_FAILED:' + ((fetched && fetched.error) || 'unknown');
      return baseSnapshot;
    }
    rawRows = fetched.data;
  } catch (error) {
    baseSnapshot.reason = 'CAPTURE_FAILED:' + (error && error.message ? error.message : 'exception');
    return baseSnapshot;
  }

  const projected = rawRows.map(r => __projectBoardRowForExport_(r));
  const pointer = __writeArchiveRows_(lessonId, phaseIdx, projected);
  if (!pointer) {
    baseSnapshot.reason = 'ARCHIVE_WRITE_FAILED';
    return baseSnapshot;
  }
  baseSnapshot.sheet = pointer.sheet;
  baseSnapshot.startRow = pointer.startRow;
  baseSnapshot.rowCount = pointer.rowCount;
  return baseSnapshot;
}

// snapshots[] への upsert: 同 phaseIndex があれば replace、なければ append。
//   phaseIndex 昇順を維持 (replay の slider が時系列で歩けるように)。
function __upsertSnapshot_(lessonJson, snapshot) {
  // 失敗 capture (reason 付き・中身なし) で、データを持つ既存 snapshot を上書きしない。
  //   Why: 再開した授業のフェーズ切替中に一時的な read 失敗があっても、過去の正常な
  //   記録 (ポインタ or 旧形式 rows) が空の失敗記録に置き換わる事故を防ぐ。
  const existing = (lessonJson.snapshots || []).find(s => s && s.phaseIndex === snapshot.phaseIndex);
  const hasData = (sn) => !!sn && ((sn.rowCount > 0) || (Array.isArray(sn.rows) && sn.rows.length > 0));
  if (snapshot.reason && !hasData(snapshot) && hasData(existing)) {
    console.warn('__upsertSnapshot_: capture 失敗のため既存 snapshot を保持: phase '
      + snapshot.phaseIndex + ' (' + snapshot.reason + ')');
    return;
  }
  lessonJson.snapshots = (lessonJson.snapshots || [])
    .filter(s => s && s.phaseIndex !== snapshot.phaseIndex)
    .concat(snapshot)
    .sort((a, b) => Number(a.phaseIndex) - Number(b.phaseIndex));
}

// 現在 active なフェーズ index を profileTransitions の最新エントリから推定。
//   無ければ phase 0 を active とみなす。
function __activePhaseIndex_(lessonJson) {
  const trans = (lessonJson && Array.isArray(lessonJson.profileTransitions))
    ? lessonJson.profileTransitions
    : [];
  if (trans.length === 0) return 0;
  const last = trans[trans.length - 1];
  return Number(last.to) || 0;
}

/**
 * Lesson を draft → active に遷移させ、全 phase の Form を生成する。
 *
 * 設計:
 *   - LockService で double-click 6-form 防止
 *   - phases[i].formId が既にあれば skip (= partial failure 後の resume 可能)
 *   - Form 生成 1 回ごとに lessonJson に書き戻し (= 途中失敗でも進捗を失わない)
 *   - 全 phase 成功後に state='active' + phase 0 を user config に activate
 */
function startLesson(userId, lessonId) {
  try {
    const auth = __requireLessonOwner_(userId, lessonId);
    if (auth.error) return auth.error;
    const { found } = auth;

    if (found.lesson.state === 'active') {
      return createSuccessResponse('already active', { lesson: found.lesson });
    }
    if (found.lesson.state !== 'draft') {
      return createErrorResponse('FORBIDDEN_STATE: lesson の状態が draft でないため開始できません');
    }

    const lessonJson = deepClone(found.lesson.lessonJson || {});
    const phases = Array.isArray(lessonJson.phases) ? lessonJson.phases : [];

    // Pre-flight 検証
    if (!Array.isArray(lessonJson.classes) || lessonJson.classes.length === 0) {
      return createErrorResponse('対象クラスを少なくとも 1 つ指定してください');
    }
    if (phases.length === 0) {
      return createErrorResponse('phase が定義されていません');
    }
    for (let i = 0; i < phases.length; i++) {
      const p = phases[i];
      if (!p || !p.name || !p.formTemplate) {
        return createErrorResponse(`phase ${i + 1} の必須項目 (name / formTemplate) が不足しています`);
      }
    }

    // double-click 6-form 防止 lock。失敗時は LESSON_BUSY を返す。
    let lock = null;
    try {
      if (typeof LockService !== 'undefined' && LockService.getScriptLock) {
        lock = LockService.getScriptLock();
        if (!lock.tryLock(LESSON_START_LOCK_TIMEOUT_MS)) {
          return createErrorResponse('LESSON_BUSY: 別の startLesson 処理が実行中です。少し待ってから再試行してください。');
        }
      }
    } catch (lockErr) {
      logError_('startLesson:lock', lockErr);
      // ロック失敗時も続行 (test sandbox 等で LockService が無いケース)
    }

    try {
      // Form 生成ループ (i=0..N-1)
      //   templateOptions には UI の wizard で教師が入力した値を 100% 流す:
      //     - lessonName + phaseName: Form タイトル "<lesson> / <phase>" に
      //     - question: Form 全体の description に + scaleTitle / choiceTitle の fallback
      //     - classChoices: クラス選択肢 (lesson.classes そのまま)
      //     - lowLabel/highLabel/xLow/xHigh/yLow/yHigh: 線形尺度の両端ラベル
      //     - choices: pie/board の選択肢
      const lessonName = (found && found.lesson && found.lesson.name) || '';
      // 授業モードの回答シートを、ボード所有者 (教師) が必ず開けるようにするため。
      //   管理者が代理で開始したときも教師のボードが壊れない。
      const boardOwner = (typeof findUserById === 'function') ? findUserById(userId) : null;
      const boardOwnerEmail = (boardOwner && boardOwner.userEmail) || '';
      const sharedClasses = (lessonJson && Array.isArray(lessonJson.classes)) ? lessonJson.classes : [];
      // 揺らぎ追跡: lesson レベルの allowResubmit を全 phase に適用 (議論前後の意見変化を取りたい)。
      //   phase 個別に templateOptions.allowResubmit があればそれを優先。
      const lessonAllowResubmit = Boolean(lessonJson && lessonJson.allowResubmit);
      for (let i = 0; i < phases.length; i++) {
        const phase = phases[i];

        // 授業モード (native): Form を作らず、回答シートだけ用意する。
        //   Why 全 phase に作るか: browse/discuss/reflect は投稿を受け付けないが、
        //   config が指す先が無いと board が「データソース未接続」になってしまう。
        //   フェーズの見え方 (何が描かれるか) は screenRole が決める。
        if (__isNativePhase_(phase, lessonJson)) {
          if (phase.sheetName) continue; // resume: 既に用意済みは skip
          try {
            const native = __ensureNativeAnswerSheet_(lessonJson, i, lessonName, boardOwnerEmail);
            phase.spreadsheetId = native.spreadsheetId;
            phase.sheetName = native.sheetName;
            phase.columnMapping = __nativeColumnMapping_(phase.formTemplate);
          } catch (nativeErr) {
            logError_('startLesson:native', nativeErr);
            __updateLessonRow_(lessonId, { lessonJson });
            return createErrorResponse(
              `phase ${i + 1} (${phase.name}) の回答シート作成に失敗しました: ${nativeErr && nativeErr.message ? nativeErr.message : 'unknown'}`,
              null,
              { error: 'NATIVE_SHEET_CREATE_FAILED', completedPhases: i, lessonId }
            );
          }
          if (i < phases.length - 1) __updateLessonRow_(lessonId, { lessonJson });
          continue;
        }

        if (phase.formId) continue; // resume: 既に作成済みは skip

        const phaseAllowResubmit = phase.templateOptions && phase.templateOptions.allowResubmit;
        const formOpts = Object.assign({}, phase.templateOptions || {}, {
          lessonName,
          phaseName: phase.name || '',
          question: phase.question || '',
          classChoices: sharedClasses,
          allowResubmit: phaseAllowResubmit != null ? Boolean(phaseAllowResubmit) : lessonAllowResubmit
        });
        const formResult = createTemplateForm(phase.formTemplate, formOpts);
        if (!formResult || !formResult.success) {
          // Partial failure: ここまでの進捗を draft のまま保存して resumable に。
          __updateLessonRow_(lessonId, { lessonJson });
          return createErrorResponse(
            `phase ${i + 1} (${phase.name}) の Form 作成に失敗しました: ${(formResult && (formResult.error || formResult.message)) || 'unknown'}`,
            null,
            { error: 'FORM_CREATE_FAILED', completedPhases: i, lessonId }
          );
        }

        phase.formId = formResult.formId || formResult.formData?.formId || '';
        phase.formUrl = formResult.formUrl || formResult.formData?.formUrl || '';
        phase.spreadsheetId = formResult.spreadsheetId || '';
        phase.sheetName = formResult.sheetName || 'フォームの回答 1';

        // Form を「受付中」にするのは phase 0 のみ。i>0 は close。
        __setFormAcceptingResponses_(phase, i === 0);

        // 進捗を逐次永続化 (次クリックで resume 可能)。
        //   最後の phase は次の applyConfigPatch_ + state=active の write でまとめて書くので skip。
        if (i < phases.length - 1) {
          __updateLessonRow_(lessonId, { lessonJson });
        }
      }

      // Phase 0 を user config に反映 → 既存 view 経路で即座に board 表示可能。
      //   publishApp を呼ばず applyConfigPatch_ で直接 merge (高速 + lifecycle 干渉なし)。
      const patchResult = applyConfigPatch_(userId, __buildPhaseConfigPatch_(phases[0], lessonJson, lessonId), { publish: false });
      if (!patchResult.success) {
        // Patch 失敗時も Form は既に作成済み。state は draft のまま、lessonJson は最新で書く。
        __updateLessonRow_(lessonId, { lessonJson });
        return createErrorResponse(
          `phase 0 の active 化に失敗しました: ${patchResult.message || 'unknown'}`,
          null,
          { error: 'PHASE_ACTIVATE_FAILED', lessonId }
        );
      }

      // 初回 transition を記録 (phase 0 開始)
      lessonJson.profileTransitions = lessonJson.profileTransitions || [];
      lessonJson.profileTransitions.push({ ts: new Date().toISOString(), from: null, to: 0 });

      // state を active に遷移 + startedAt 記録
      const startedAt = new Date().toISOString();
      const finalResult = __updateLessonRow_(lessonId, {
        state: 'active',
        startedAt,
        lessonJson
      });
      if (!finalResult.success) {
        return createErrorResponse(finalResult.message || finalResult.error);
      }
      return createSuccessResponse('lesson 開始しました', { lesson: finalResult.lesson });
    } finally {
      if (lock && lock.releaseLock) {
        try { lock.releaseLock(); } catch (_) { /* lock は TTL で自動解放される */ }
      }
    }
  } catch (error) {
    logError_('startLesson', error);
    return createExceptionResponse(error);
  }
}

/**
 * フェーズを 1 つ進める/戻す。
 *   - 累積回答は削除しない (= 戻して再進行で同じ Form が再開する)
 *   - publishApp は呼ばず applyConfigPatch_ で user config を直接書き換える
 */
function advanceLessonPhase(userId, lessonId, direction, targetIndex) {
  try {
    const auth = __requireLessonOwner_(userId, lessonId);
    if (auth.error) return auth.error;
    const { found } = auth;

    if (found.lesson.state !== 'active') {
      return createErrorResponse('FORBIDDEN_STATE: lesson が active でないため進められません');
    }

    const lessonJson = deepClone(found.lesson.lessonJson || {});
    const phases = lessonJson.phases || [];
    const fromIdx = __activePhaseIndex_(lessonJson);

    // targetIndex を渡せば任意 phase へ直接ジャンプできる。
    //   Why: 回答ボード側の phase pill は「今どこにいるか」を全 phase 並べて見せるので、
    //   隣以外を押せてしまう。±1 しか無いと押しても何も起きない死んだ UI になる。
    //   遷移の中身 (snapshot → row write → Form 開閉 → config patch) は from/to が
    //   何であっても同じなので、経路は 1 本のまま添字の決め方だけを分ける。
    let toIdx;
    if (targetIndex !== undefined && targetIndex !== null && targetIndex !== '') {
      toIdx = Number(targetIndex);
      if (!Number.isInteger(toIdx)) return createErrorResponse('targetIndex が不正です');
      if (toIdx === fromIdx) return createErrorResponse('既にそのフェーズです');
    } else {
      toIdx = direction === 'previous' ? fromIdx - 1 : fromIdx + 1;
    }

    if (toIdx < 0) return createErrorResponse('既に最初のフェーズです');
    if (toIdx >= phases.length) return createErrorResponse('既に最後のフェーズです (終了するには「⏹ 終了」を押してください)');

    // Why: 移行 *前* に outgoing phase の rows を freeze する。順序を逆にすると
    //   user config が次 phase の columnMapping を指した状態で capture することになり、
    //   replay が破綻する。capture は config 切替より前 (= 現状 fromIdx) で行う。
    __upsertSnapshot_(lessonJson, __captureSnapshot_(userId, lessonJson, fromIdx, lessonId));

    lessonJson.profileTransitions = lessonJson.profileTransitions || [];
    lessonJson.profileTransitions.push({ ts: new Date().toISOString(), from: fromIdx, to: toIdx });

    // Why この順序: lesson row (= 「今どのフェーズか」 の真実) を etag 検証付きで *先に* 確定する。
    //   旧実装は Form/config を先に切替えてから row write していたため、 最後の write が
    //   etag_mismatch で失敗すると「config は次 phase・lessonJson は前 phase」 の不整合が残り
    //   __activePhaseIndex_ がズレた。 row write を concurrency gate にし、 成功後に
    //   *冪等な* 副作用 (Form open/close, config patch — 二重適用しても無害) を適用する。
    const result = __updateLessonRow_(lessonId, { lessonJson });
    if (!result.success) {
      // Why error preservation: __updateLessonRow_ は 'etag_mismatch' を error フィールドで返す。
      //   旧来は createErrorResponse(message || error) のみで wrap し error code を捨てて
      //   いたため、frontend の auto-retry-on-mismatch ロジックが triggers されなかった
      //   (Error envelope audit F4)。
      return createErrorResponse(result.message || result.error, null,
        result.error ? { error: result.error, currentEtag: result.currentEtag } : null);
    }

    // 現フェーズ Form を close、次フェーズ Form を open (冪等)。
    __setFormAcceptingResponses_(phases[fromIdx], false);
    __setFormAcceptingResponses_(phases[toIdx], true);

    // user config を次フェーズに切替 (board が即座に新フェーズに対応; 冪等)
    const target = phases[toIdx];
    const patchResult = applyConfigPatch_(userId, __buildPhaseConfigPatch_(target, lessonJson, lessonId), { publish: false });
    if (!patchResult.success) {
      return createErrorResponse(`フェーズ切替に失敗しました: ${patchResult.message || 'unknown'}`);
    }

    return createSuccessResponse(`フェーズ ${toIdx + 1}: ${target.name} に切替えました`, {
      lesson: result.lesson,
      activePhaseIndex: toIdx
    });
  } catch (error) {
    logError_('advanceLessonPhase', error);
    return createExceptionResponse(error);
  }
}

/**
 * 完了した授業を再開する (completed → active)。
 *
 * Why: 「翌日に続きを受け付けたい」「終了操作が早すぎた」を救う逆遷移。
 *   再開位置は profileTransitions の最終 to (= 終了時に active だった phase)。
 *
 * 順序は advanceLessonPhase と同じ思想: lesson row (真実) を先に確定し、
 * 冪等な副作用 (Form 受付再開 / config patch) を後に適用する。
 * snapshot は触らない — 次にフェーズを離れるとき通常経路で焼き直され、
 * 読めなかった場合も __upsertSnapshot_ のガードが既存記録を守る。
 */
function reopenLesson(userId, lessonId) {
  try {
    const auth = __requireLessonOwner_(userId, lessonId);
    if (auth.error) return auth.error;
    const { found } = auth;

    if (found.lesson.state === 'active') {
      return createSuccessResponse('already active', { lesson: found.lesson });
    }
    if (found.lesson.state !== 'completed') {
      return createErrorResponse('FORBIDDEN_STATE: 完了した授業のみ再開できます');
    }

    const lessonJson = deepClone(found.lesson.lessonJson || {});
    const phases = Array.isArray(lessonJson.phases) ? lessonJson.phases : [];
    if (phases.length === 0) return createErrorResponse('phase が定義されていません');
    const idx = Math.min(__activePhaseIndex_(lessonJson), phases.length - 1);

    const result = __updateLessonRow_(lessonId, { state: 'active', endedAt: '', lessonJson });
    if (!result.success) {
      return createErrorResponse(result.message || result.error, null,
        result.error ? { error: result.error, currentEtag: result.currentEtag } : null);
    }

    // 再開 phase の Form だけ受付再開 (他 phase は advance が通過時に開閉する)。
    __setFormAcceptingResponses_(phases[idx], true);

    // config を再開 phase に合わせる + auto-archive の safety net marker を立てる
    //   (通常は publishApp が立てるが、再開はボード公開済みのまま行われうる)。
    const patch = Object.assign(
      __buildPhaseConfigPatch_(phases[idx], lessonJson, lessonId),
      { currentLessonStartedAt: new Date().toISOString() }
    );
    const patchResult = applyConfigPatch_(userId, patch, { publish: false });
    if (!patchResult.success) {
      return createErrorResponse(`再開しましたが board 設定の切替に失敗しました: ${patchResult.message || 'unknown'}`);
    }

    return createSuccessResponse(`授業を再開しました (フェーズ ${idx + 1}: ${phases[idx].name})`, {
      lesson: result.lesson,
      activePhaseIndex: idx
    });
  } catch (error) {
    logError_('reopenLesson', error);
    return createExceptionResponse(error);
  }
}

/**
 * Lesson を completed に遷移し、最終 snapshot を archive。
 *   - 全 phase Form を close
 *   - 現フェーズの rows を snapshots[currentPhase] に保存
 *   - user config の formUrl 等は触らない (viewer URL は引き続き読める)
 */
function endLesson(userId, lessonId) {
  try {
    const auth = __requireLessonOwner_(userId, lessonId);
    if (auth.error) return auth.error;
    const { found } = auth;

    if (found.lesson.state !== 'active') {
      return createErrorResponse('FORBIDDEN_STATE: lesson が active でないため終了できません');
    }

    const lessonJson = deepClone(found.lesson.lessonJson || {});
    const phases = lessonJson.phases || [];

    // 全 phase Form を close (締切)
    phases.forEach(p => __setFormAcceptingResponses_(p, false));

    // 現フェーズの rows を freeze して snapshots に upsert。capture 失敗時は reason 付き空 snapshot
    //   が積まれ、endLesson 自体は成功する (lesson は必ず completed に遷移する原則)。
    const currentIdx = __activePhaseIndex_(lessonJson);
    __upsertSnapshot_(lessonJson, __captureSnapshot_(userId, lessonJson, currentIdx, lessonId));

    const endedAt = new Date().toISOString();
    const result = __updateLessonRow_(lessonId, {
      state: 'completed',
      endedAt,
      lessonJson
    });
    if (!result.success) {
      // Why: etag_mismatch などの error code を捨てず propagate (Error envelope audit F4)
      return createErrorResponse(result.message || result.error, null,
        result.error ? { error: result.error, currentEtag: result.currentEtag } : null);
    }
    return createSuccessResponse('lesson を終了しました。振り返り画面でいつでも再生できます。', {
      lesson: result.lesson,
      reviewUrl: '?mode=review&lessonId=' + encodeURIComponent(lessonId)
    });
  } catch (error) {
    logError_('endLesson', error);
    return createExceptionResponse(error);
  }
}

/**
 * 授業の全 Form の回答受付を締め切る (状態を問わず実行できる)。
 *
 * Why endLesson と別に要るか: endLesson は active な授業しか対象にできない。
 *   すでに completed の授業で Form が開いたままなら、締め切る手段が
 *   「Google フォームを開いて 1 つずつ」しかなくなる。教師をアプリの外に出さない。
 *
 * Why 教師本人が実行する必要があるか: FormApp は Form の所有者の権限で動く。
 *   API キー経由 (管理者) では他人の Form を開けないので、管理パネルから
 *   本人が実行する導線として用意する。
 *
 * @returns {Object} { closed, failed, repaired, total }
 */
function closeLessonForms(userId, lessonId) {
  try {
    const auth = __requireLessonOwner_(userId, lessonId);
    if (auth.error) return auth.error;
    const { found } = auth;

    const lessonJson = deepClone(found.lesson.lessonJson || {});
    const phases = Array.isArray(lessonJson.phases) ? lessonJson.phases : [];
    if (phases.length === 0) {
      return createErrorResponse('フェーズがありません');
    }
    if (__isNativePhase_({}, lessonJson)) {
      return createSuccessResponse('この授業は Google フォームを使っていません', {
        closed: 0, failed: 0, repaired: 0, total: 0
      });
    }

    let closed = 0, failed = 0, repaired = 0;
    const failedPhases = [];
    for (let i = 0; i < phases.length; i++) {
      const before = phases[i].formId;
      if (__setFormAcceptingResponses_(phases[i], false)) {
        closed++;
        // __openLessonForm_ が編集用 ID に直していたら記録する。
        if (phases[i].formId !== before) repaired++;
      } else {
        failed++;
        failedPhases.push(phases[i].name || ('フェーズ ' + (i + 1)));
      }
    }

    // 復旧した formId を永続化する (次回は 1 発で開く)。
    if (repaired > 0) __updateLessonRow_(lessonId, { lessonJson });

    const msg = failed === 0
      ? `${closed} 件のフォームを締め切りました`
      : `${closed} 件を締め切り、${failed} 件は開けませんでした (${failedPhases.join(' / ')})`;
    return createSuccessResponse(msg, { closed, failed, repaired, total: phases.length });
  } catch (error) {
    logError_('closeLessonForms', error);
    return createExceptionResponse(error);
  }
}

// ----- Auto-archive (unpublishBoard hook + 23:00 cron sweep) -----

/**
 * unpublishBoard 経由で呼ばれる auto-archive 判定 + 実行。
 *
 * Why: 教師が「公開を終了」した瞬間 = 授業終了の意思表示。
 *   下記条件を満たすときだけ endLesson を実行し、archive する。
 *   - currentLessonStartedAt から 5 分以上経過 (< 5 分はデモ操作とみなして skip)
 *   - 回答が 1 件以上 (0 件 = 授業として成立してない)
 *
 *   __applyPublishStateChange の同 save 内で呼ばれ、archived=true なら
 *   呼び出し側で activeLessonId / currentLessonStartedAt をクリアする。
 *
 * @returns {Object} { archived: boolean, lessonId?: string, reason?: string }
 */
function __maybeAutoArchiveLesson_(targetUser, currentConfig) {
  try {
    const activeLessonId = currentConfig.activeLessonId;
    const startedAt = currentConfig.currentLessonStartedAt;
    if (!activeLessonId || !startedAt) return { archived: false, reason: 'no_active_lesson' };

    const elapsedMs = Date.now() - new Date(startedAt).getTime();
    if (!(elapsedMs >= LESSON_AUTO_ARCHIVE_MIN_MS)) {
      return { archived: false, reason: 'too_short' };
    }

    let responseCount = 0;
    try {
      const data = getPublishedSheetData(null, 'newest', true, targetUser.userId);
      responseCount = (data && Array.isArray(data.data)) ? data.data.length : 0;
    } catch (fetchErr) {
      logError_('__maybeAutoArchiveLesson_:fetchCount', fetchErr);
      return { archived: false, reason: 'fetch_failed' };
    }
    if (responseCount < LESSON_AUTO_ARCHIVE_MIN_RESPONSES) {
      return { archived: false, reason: 'no_responses' };
    }

    const endResult = endLesson(targetUser.userId, activeLessonId);
    if (!endResult || !endResult.success) {
      logError_('__maybeAutoArchiveLesson_:endLesson', new Error(endResult && endResult.message));
      return { archived: false, reason: 'end_failed' };
    }
    return { archived: true, lessonId: activeLessonId };
  } catch (error) {
    logError_('__maybeAutoArchiveLesson_', error);
    return { archived: false, reason: 'exception' };
  }
}

/**
 * 1 日 1 回の cron entry。公開し忘れ lesson を回収する。
 *
 * Why: unpublish し忘れて翌日になったケースの safety net。
 *   currentLessonStartedAt が立っているユーザーを enumerate し、
 *   最終回答から LESSON_DAILY_STALE_HOURS 以上経過していれば auto-archive。
 *   isPublished は touch しない (公開ライフサイクル 4 関数のみが扱う規約)。
 *   marker のクリアは applyConfigPatch_ 経由で行う。
 */
function dailyLessonArchiveSweep() {
  const summary = { scanned: 0, archived: 0, skipped: 0, errors: 0 };
  try {
    if (typeof getAllUsers !== 'function') return summary;
    const users = getAllUsers({ activeOnly: false }, { forceServiceAccount: true });
    const staleThresholdMs = LESSON_DAILY_STALE_HOURS * 60 * 60 * 1000;
    const now = Date.now();

    for (const user of (users || [])) {
      summary.scanned++;
      let cfg;
      try {
        cfg = (typeof getConfigOrDefault === 'function') ? getConfigOrDefault(user.userId, user) : null;
      } catch (cfgErr) {
        summary.errors++;
        logError_('dailyLessonArchiveSweep:getConfig', cfgErr);
        continue;
      }
      if (!cfg || !cfg.currentLessonStartedAt || !cfg.activeLessonId) {
        summary.skipped++;
        continue;
      }

      let mostRecent = new Date(cfg.currentLessonStartedAt).getTime();
      try {
        const data = getPublishedSheetData(null, 'newest', true, user.userId);
        const rows = (data && Array.isArray(data.data)) ? data.data : [];
        for (const r of rows) {
          const ts = r && r.timestamp ? new Date(r.timestamp).getTime() : 0;
          if (ts > mostRecent) mostRecent = ts;
        }
      } catch (fetchErr) {
        logError_('dailyLessonArchiveSweep:fetchRows', fetchErr);
      }

      if (now - mostRecent < staleThresholdMs) {
        summary.skipped++;
        continue;
      }

      const archiveResult = __maybeAutoArchiveLesson_(user, cfg);
      if (archiveResult.archived) {
        summary.archived++;
        try {
          applyConfigPatch_(user.userId, {
            activeLessonId: null,
            currentLessonStartedAt: null
          }, { publish: false });
        } catch (patchErr) {
          summary.errors++;
          logError_('dailyLessonArchiveSweep:clearMarker', patchErr);
        }
      } else {
        summary.skipped++;
      }
    }
  } catch (error) {
    logError_('dailyLessonArchiveSweep', error);
    summary.errors++;
  }
  return summary;
}

/**
 * lesson 関連の time-based trigger を冪等にインストール。
 *   setupApp から呼ばれる。既に同名トリガーが居れば何もしない。
 */
function installLessonTriggers() {
  try {
    if (typeof ScriptApp === 'undefined' || !ScriptApp.getProjectTriggers) return;
    const existing = ScriptApp.getProjectTriggers();
    const already = existing.some(t => t.getHandlerFunction && t.getHandlerFunction() === 'dailyLessonArchiveSweep');
    if (already) return;
    ScriptApp.newTrigger('dailyLessonArchiveSweep')
      .timeBased()
      .everyDays(1)
      .atHour(LESSON_DAILY_TRIGGER_HOUR)
      .create();
  } catch (error) {
    logError_('installLessonTriggers', error);
  }
}

/**
 * phase の並び順を入れ替える (保守オペレーション)。
 *
 * Why: v2894 の profiles 取り込みは「保存順」を phase 順として引き継いだため、
 *   実際の授業順 (遷移ログ) と配列順が食い違う lesson が存在する。pill は配列順で
 *   並ぶので、教師には「導入より本時が先」に見える。
 *
 * 整合性: phaseIndex は 3 箇所に現れる — phases 配列 / snapshots / lesson_responses の
 *   行 (hydrate の照合キー)。行は in-place 編集せず、hydrate で読み戻してから新しい
 *   phaseIndex で追記し直し、ポインタを差し替える (旧範囲は孤児 = 無害)。
 *
 * @param {number[]} order - 新しい並び。旧 index の列挙 (例 [1,0,2] = 旧1 が先頭へ)
 */
function reorderLessonPhases(userId, lessonId, order) {
  try {
    const auth = __requireLessonOwner_(userId, lessonId);
    if (auth.error) return auth.error;
    const { found } = auth;
    const lessonJson = deepClone(found.lesson.lessonJson || {});
    const phases = Array.isArray(lessonJson.phases) ? lessonJson.phases : [];

    // order は 0..n-1 の順列であること
    if (!Array.isArray(order) || order.length !== phases.length
        || [...order].sort((a, b) => a - b).some((v, i) => v !== i)) {
      return createErrorResponse('order は phase 数と同じ長さの順列で指定してください');
    }
    if (order.every((v, i) => v === i)) {
      return createSuccessResponse('並び順は既にその通りです', { changed: false });
    }
    const newIndexOf = [];            // 旧 index → 新 index
    order.forEach((oldIdx, newIdx) => { newIndexOf[oldIdx] = newIdx; });

    // アーカイブ行の phaseIndex を新順序で書き直すため、先に読み戻す
    __hydrateLessonSnapshots_({ lessonId, lessonJson });

    lessonJson.phases = order.map(i => phases[i]);
    lessonJson.profileTransitions = (lessonJson.profileTransitions || []).map(t => ({
      ...t,
      from: (t.from === null || t.from === undefined) ? t.from : newIndexOf[t.from],
      to: newIndexOf[t.to]
    }));
    if (lessonJson.meta && Array.isArray(lessonJson.meta.profileNames)) {
      lessonJson.meta.profileNames = lessonJson.phases.map(p => p.name);
    }

    lessonJson.snapshots = (lessonJson.snapshots || []).map(sn => {
      const newIdx = newIndexOf[sn.phaseIndex];
      const rows = Array.isArray(sn.rows) ? sn.rows : [];
      if (rows.length > 0) {
        const pointer = __writeArchiveRows_(lessonId, newIdx, rows);
        if (pointer) {
          return { ...sn, phaseIndex: newIdx, rows: [],
            sheet: pointer.sheet, startRow: pointer.startRow, rowCount: pointer.rowCount };
        }
        // 書き直せなければ旧ポインタを捨てて inline のまま保持 (データ喪失より整合性劣化を選ぶ)
        return { ...sn, phaseIndex: newIdx, sheet: null, startRow: -1, rowCount: rows.length, rows };
      }
      return { ...sn, phaseIndex: newIdx };
    }).sort((a, b) => Number(a.phaseIndex) - Number(b.phaseIndex));

    const result = __updateLessonRow_(lessonId, { lessonJson });
    if (!result.success) return createErrorResponse(result.message || result.error);
    return createSuccessResponse('phase の並び順を変更しました', {
      changed: true,
      phases: lessonJson.phases.map((p, i) => i + ':' + p.name),
      activePhaseIndex: __activePhaseIndex_(lessonJson)
    });
  } catch (error) {
    logError_('reorderLessonPhases', error);
    return createExceptionResponse(error);
  }
}

/**
 * 指定 phase のアーカイブを「現在の公開ボード」から焼き直す。
 * 用途: 移行時に本文が切り詰められていた phase を、元 SS が読める状態で全文化する。
 *   config.spreadsheetId が対象 phase の SS を指していることは呼び出し側の責任
 *   (previewBoard で件数を確認してから呼ぶ)。
 */
function recaptureLessonArchive(userId, lessonId, phaseIndex) {
  try {
    const auth = __requireLessonOwner_(userId, lessonId);
    if (auth.error) return auth.error;
    const { found } = auth;
    const idx = Number(phaseIndex);
    const lessonJson = deepClone(found.lesson.lessonJson || {});
    const phases = Array.isArray(lessonJson.phases) ? lessonJson.phases : [];
    if (!Number.isInteger(idx) || idx < 0 || idx >= phases.length) {
      return createErrorResponse('phaseIndex が不正です');
    }
    const snapshot = __captureSnapshot_(userId, lessonJson, idx, lessonId);
    if (snapshot.reason) {
      return createErrorResponse('再取得失敗: ' + snapshot.reason);
    }
    __upsertSnapshot_(lessonJson, snapshot);
    const result = __updateLessonRow_(lessonId, { lessonJson });
    if (!result.success) return createErrorResponse(result.message || result.error);
    return createSuccessResponse('phase ' + idx + ' を全文で焼き直しました', {
      phaseIndex: idx, rowCount: snapshot.rowCount, startRow: snapshot.startRow
    });
  } catch (error) {
    logError_('recaptureLessonArchive', error);
    return createExceptionResponse(error);
  }
}

// =====================================================================
// 授業モード (native 入力) の公開 API
//
// 児童の端末から直接呼ばれる。owner 限定の管理 API とは別経路で、認可は
// 「公開中のボードを見ている本人」であること (リアクションと同じ基準)。
// =====================================================================

/**
 * 児童が見ている授業の「今のフェーズ」を返す。
 *
 * getActiveLessonNav (owner 限定・全フェーズ一覧) とは別物。こちらは閲覧者に見せてよい
 * 最小限 = 今どのフェーズで、画面で何ができて、問いは何か、だけを返す。
 * Why 分けるか: 児童に未来のフェーズの問いや構成を先に見せない。
 *
 * @param {string} targetUserId - ボード所有者 (教師) の userId
 * @returns {Object|null} 授業中でなければ null
 */
function __getViewerLessonPhase_(targetUserId) {
  try {
    if (!targetUserId) return null;
    const config = getConfigOrDefault(targetUserId);
    const lessonId = config && config.activeLessonId;
    if (!lessonId) return null;

    const found = __findLessonById_(lessonId);
    if (!found || !found.lesson || found.lesson.state !== 'active') return null;

    const lessonJson = found.lesson.lessonJson || {};
    if (!__isNativePhase_({}, lessonJson)) return null;  // Form 経由の授業では使わない

    const phases = Array.isArray(lessonJson.phases) ? lessonJson.phases : [];
    const idx = __activePhaseIndex_(lessonJson);
    const phase = phases[idx];
    if (!phase) return null;

    return {
      lessonId,
      phaseIndex: idx,
      phaseName: phase.name || '',
      screenRole: __phaseScreenRole_(phase),
      question: phase.question || '',
      phaseCount: phases.length
    };
  } catch (error) {
    logError_('__getViewerLessonPhase_', error);
    return null;
  }
}

// 線形尺度 (1-5) の検証。範囲外・非数値は null。
function __validateLessonScale_(value) {
  const n = Number(value);
  if (!Number.isFinite(n)) return null;
  const rounded = Math.round(n);
  if (rounded < 1 || rounded > 5) return null;
  return rounded;
}

// 児童が書くテキストの正規化。制御文字を落とし、長さで切る。
function __sanitizeLessonText_(value, maxLen) {
  if (value === null || value === undefined) return '';
  const limit = maxLen || LESSON_ANSWER_TEXT_MAX;
  return String(value)
    .replace(/[\u0000-\u001f\u007f]/g, ' ')
    .trim()
    .slice(0, limit);
}

// 同一フェーズ内の自分の行番号 (1-based) を返す。無ければ -1。
//   email 列だけを読む (全列読みは 30 人 × 投稿で無駄が大きい)。
function __findOwnLessonRow_(sheet, actorEmail) {
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return -1;
    const emails = sheet.getRange(2, LESSON_NATIVE_COL_EMAIL, lastRow - 1, 1).getValues();
    const target = String(actorEmail || '').trim().toLowerCase();
    for (let i = 0; i < emails.length; i++) {
      if (String(emails[i][0] || '').trim().toLowerCase() === target) return i + 2;
    }
    return -1;
  } catch (error) {
    logError_('__findOwnLessonRow_', error);
    return -1;
  }
}

/**
 * 児童の投稿を受け付ける (doPost: submitLessonAnswer)。
 *
 * 権能の最終防衛線: 画面側で入力 UI を隠していても、フェーズが input/reinput でなければ
 * サーバが拒否する。「議論中は投稿できない」は UI ではなくここで保証される。
 *
 * 1 フェーズ 1 児童 1 行。同じフェーズ内の置き直しは既存行を更新する
 * (「送ったけれど、やっぱり違う」を許す)。フェーズをまたぐ差分だけが航跡になる。
 *
 * @param {string} targetUserId
 * @param {Object} payload - { lessonId, phaseIndex, numericX, numericY, reason, addedInsight, class, name }
 */
function submitLessonAnswer(targetUserId, payload) {
  try {
    const actorEmail = getCurrentEmail();
    if (!actorEmail) return createAuthError();

    const p = payload || {};
    const phase = __getViewerLessonPhase_(targetUserId);
    if (!phase) return createErrorResponse('この授業はいま投稿を受け付けていません');

    // client の lessonId / phaseIndex は「ズレの検出」にのみ使う。真実はサーバの active phase。
    //   フェーズ切替の直後に届いた投稿を、前のフェーズの行として書かないための照合。
    if (p.lessonId && String(p.lessonId) !== String(phase.lessonId)) {
      return createErrorResponse('PHASE_CHANGED: 授業が切り替わりました。画面を読み込み直してください');
    }
    if (p.phaseIndex !== undefined && p.phaseIndex !== null && Number(p.phaseIndex) !== phase.phaseIndex) {
      return createErrorResponse('PHASE_CHANGED: フェーズが切り替わりました。画面を読み込み直してください');
    }
    if (LESSON_INPUT_ROLES.indexOf(phase.screenRole) < 0) {
      return createErrorResponse('いまは考えを送る時間ではありません');
    }

    const found = __findLessonById_(phase.lessonId);
    if (!found || !found.lesson) return createErrorResponse('授業が見つかりません');
    const lessonJson = found.lesson.lessonJson || {};
    const phaseDef = (lessonJson.phases || [])[phase.phaseIndex];
    if (!phaseDef || !phaseDef.spreadsheetId || !phaseDef.sheetName) {
      return createErrorResponse('回答シートが未設定です');
    }

    const isMatrix = phaseDef.formTemplate === 'matrix';
    const x = __validateLessonScale_(p.numericX);
    if (x === null) return createErrorResponse('横軸の値が不正です');
    const y = isMatrix ? __validateLessonScale_(p.numericY) : null;
    if (isMatrix && y === null) return createErrorResponse('縦軸の値が不正です');

    const reason = __sanitizeLessonText_(p.reason);
    if (!reason) return createErrorResponse('理由を書いてください');
    const addedInsight = __sanitizeLessonText_(p.addedInsight);

    const ss = openSpreadsheet(phaseDef.spreadsheetId, { context: 'lesson_submit' });
    if (!ss) return createErrorResponse('回答シートを開けませんでした');
    const sheet = ss.getSheetByName(phaseDef.sheetName);
    if (!sheet) return createErrorResponse('回答シートが見つかりません');

    const row = [
      new Date().toISOString(),
      actorEmail,
      __sanitizeLessonText_(p.class, 50),
      __sanitizeLessonText_(p.name, 50),
      x,
      (y === null) ? '' : y,
      reason,
      addedInsight
    ];

    const existingRow = __findOwnLessonRow_(sheet, actorEmail);
    if (existingRow > 0) {
      sheet.getRange(existingRow, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }

    // browse フェーズを見ている他の児童の画面に反映されるよう board cache を落とす。
    if (typeof bumpBoardDataVersion_ === 'function') {
      try {
        bumpBoardDataVersion_(targetUserId);
      } catch (cacheErr) {
        // 黙って落とさない: cache version を上げ損ねると、他の児童のボードに
        //   古い分布が最大 12 秒残る。症状 (反映されない) だけが出て原因が
        //   追えなくなるので、必ず痕跡を残す。処理自体は続行してよい。
        console.warn('bumpBoardDataVersion_ failed (board may be stale up to 12s):',
          cacheErr && cacheErr.message);
      }
    }

    return createSuccessResponse('考えを送りました', {
      phaseIndex: phase.phaseIndex,
      updated: existingRow > 0
    });
  } catch (error) {
    logError_('submitLessonAnswer', error);
    return createExceptionResponse(error);
  }
}

// 1 フェーズ分の自分の回答を読む。無ければ null。
function __readOwnLessonAnswer_(phaseDef, actorEmail) {
  try {
    if (!phaseDef || !phaseDef.spreadsheetId || !phaseDef.sheetName) return null;
    const ss = openSpreadsheet(phaseDef.spreadsheetId, { context: 'lesson_trajectory' });
    if (!ss) return null;
    const sheet = ss.getSheetByName(phaseDef.sheetName);
    if (!sheet) return null;
    const rowNum = __findOwnLessonRow_(sheet, actorEmail);
    if (rowNum < 2) return null;
    const v = sheet.getRange(rowNum, 1, 1, LESSON_NATIVE_SHEET_HEADERS.length).getValues()[0];
    const x = Number(v[LESSON_NATIVE_COL_X - 1]);
    const y = Number(v[LESSON_NATIVE_COL_X]);
    return {
      numericX: Number.isFinite(x) ? x : null,
      numericY: Number.isFinite(y) ? y : null,
      reason: String(v[LESSON_NATIVE_COL_INSIGHT - 2] || ''),
      addedInsight: String(v[LESSON_NATIVE_COL_INSIGHT - 1] || '')
    };
  } catch (error) {
    logError_('__readOwnLessonAnswer_', error);
    return null;
  }
}

/**
 * 教師の見取りグリッド: 児童ごとの ● 最初 → ★ いま を一覧で返す。
 *
 * 授業「後」に読むための API。授業中の教師はフェーズ送りと投影に集中する想定で、
 * この画面は児童の顔を見ている時間には出さない。
 *
 * 並び順は移動距離の昇順 = 位置が動かなかった児童が先頭に来る。
 * Why: 「位置は変わらないが理由が深まった」学びは、移動量で並べると最後尾に沈む。
 *   教師が最初に読むべきものを最初に置く。深さの判定そのものはしない (教師がする)。
 *
 * @param {string} userId - 授業の所有者
 * @param {string} lessonId
 * @returns {Object} { students: [{ name, email, class, first, last, distance, moved }] }
 */
function getLessonReviewGrid(userId, lessonId) {
  try {
    const auth = __requireLessonOwner_(userId, lessonId);
    if (auth.error) return auth.error;
    const lessonJson = (auth.found.lesson.lessonJson) || {};
    const phases = Array.isArray(lessonJson.phases) ? lessonJson.phases : [];

    // 入力フェーズだけを時系列で読む (browse/discuss には投稿が存在しない)。
    const inputPhases = [];
    for (let i = 0; i < phases.length; i++) {
      if (LESSON_INPUT_ROLES.indexOf(__phaseScreenRole_(phases[i])) >= 0) {
        inputPhases.push({ index: i, def: phases[i] });
      }
    }
    if (inputPhases.length === 0) {
      return createSuccessResponse('入力フェーズなし', { students: [], phaseCount: 0 });
    }

    // email をキーに、フェーズごとの回答を集める。
    const byEmail = new Map();
    for (let p = 0; p < inputPhases.length; p++) {
      const rows = __readAllLessonRows_(inputPhases[p].def);
      for (let r = 0; r < rows.length; r++) {
        const row = rows[r];
        if (!row.email) continue;
        const entry = byEmail.get(row.email) || { email: row.email, name: '', class: '', answers: [] };
        // 名前は後のフェーズで入力されることもあるので、空でなければ最新で更新する。
        if (row.name) entry.name = row.name;
        if (row.class) entry.class = row.class;
        entry.answers.push(Object.assign({ phaseIndex: inputPhases[p].index }, row));
        byEmail.set(row.email, entry);
      }
    }

    const students = [];
    byEmail.forEach((entry) => {
      const answers = entry.answers.sort((a, b) => a.phaseIndex - b.phaseIndex);
      const first = answers[0] || null;
      const last = answers.length > 1 ? answers[answers.length - 1] : null;
      let distance = 0;
      if (first && last) {
        const dx = (Number(last.numericX) || 0) - (Number(first.numericX) || 0);
        const dy = (Number(last.numericY) || 0) - (Number(first.numericY) || 0);
        distance = Math.sqrt(dx * dx + dy * dy);
      }
      students.push({
        name: entry.name,
        email: entry.email,
        class: entry.class,
        first,
        last,
        distance,
        // 位置が動いたかどうかは事実として返すだけ。評価はしない。
        moved: distance > 0,
        answeredPhases: answers.length
      });
    });

    students.sort((a, b) => {
      // 未提出 (last なし) は最後に。それ以外は移動距離の小さい順。
      if (!a.last && b.last) return 1;
      if (a.last && !b.last) return -1;
      return a.distance - b.distance;
    });

    return createSuccessResponse('見取りグリッド', {
      students,
      phaseCount: inputPhases.length
    });
  } catch (error) {
    logError_('getLessonReviewGrid', error);
    return createExceptionResponse(error);
  }
}

// 1 フェーズ分の全回答を読む (教師の見取り用)。
function __readAllLessonRows_(phaseDef) {
  try {
    if (!phaseDef || !phaseDef.spreadsheetId || !phaseDef.sheetName) return [];
    const ss = openSpreadsheet(phaseDef.spreadsheetId, { context: 'lesson_review_grid' });
    if (!ss) return [];
    const sheet = ss.getSheetByName(phaseDef.sheetName);
    if (!sheet) return [];
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    // 全行を 1 回で読む (行ごとの getValue は 70x 遅い)。
    const values = sheet.getRange(2, 1, lastRow - 1, LESSON_NATIVE_SHEET_HEADERS.length).getValues();
    const out = [];
    for (let i = 0; i < values.length; i++) {
      const v = values[i];
      const email = String(v[1] || '').trim().toLowerCase();
      if (!email) continue;
      const x = Number(v[4]);
      const y = Number(v[5]);
      out.push({
        email,
        class: String(v[2] || ''),
        name: String(v[3] || ''),
        numericX: Number.isFinite(x) ? x : null,
        numericY: Number.isFinite(y) ? y : null,
        reason: String(v[6] || ''),
        addedInsight: String(v[7] || '')
      });
    }
    return out;
  } catch (error) {
    logError_('__readAllLessonRows_', error);
    return [];
  }
}

/**
 * 自分の航跡 (● 最初の考え → ★ いまの考え) を返す。
 *
 * Why 本人にしか返さないか: 学級全体の分布に個人の変化を重ねると、誰がどう動いたかが
 * 教室で可視になり、移動そのものが評価に見える。変化は本人と (授業後に) 教師だけが見る。
 */
function getMyLessonTrajectory(targetUserId) {
  try {
    const actorEmail = getCurrentEmail();
    if (!actorEmail) return createAuthError();

    const config = getConfigOrDefault(targetUserId);
    const lessonId = config && config.activeLessonId;
    if (!lessonId) return createSuccessResponse('授業なし', { phases: [] });

    const found = __findLessonById_(lessonId);
    if (!found || !found.lesson) return createSuccessResponse('授業なし', { phases: [] });
    const lessonJson = found.lesson.lessonJson || {};
    const phases = Array.isArray(lessonJson.phases) ? lessonJson.phases : [];
    const activeIdx = __activePhaseIndex_(lessonJson);

    const out = [];
    for (let i = 0; i <= activeIdx && i < phases.length; i++) {
      const ph = phases[i];
      // 入力フェーズだけが航跡の点になる (browse/discuss には投稿が存在しない)。
      if (LESSON_INPUT_ROLES.indexOf(__phaseScreenRole_(ph)) < 0) continue;
      const entry = __readOwnLessonAnswer_(ph, actorEmail);
      if (entry) out.push(Object.assign({ phaseIndex: i, phaseName: ph.name || '' }, entry));
    }
    return createSuccessResponse('航跡', { phases: out });
  } catch (error) {
    logError_('getMyLessonTrajectory', error);
    return createExceptionResponse(error);
  }
}
