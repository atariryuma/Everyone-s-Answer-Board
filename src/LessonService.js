/**
 * @fileoverview LessonService - 授業 (lesson) の作成・進行・archive。
 *   状態機械: draft (wizard 編集中) → active (Form 生成済み・授業中) → completed (archive 済み)
 *   owner-only auth (管理者は listLessons のみ全件取得可)。
 */

/* global openDatabase, getCurrentEmail, isAdministrator, findUserByEmail, findUserById, createTemplateForm, applyConfigPatch_, getPublishedSheetData, getPublishedSheetDataForProfile, getAllUsers, getConfigOrDefault, getCachedProperty, LESSONS_SHEET_HEADERS, deepClone, createSuccessResponse, createErrorResponse, createExceptionResponse, createUserNotFoundError, createAuthError, isBoardCollaborator, logError_ */

// schemaVersion を bump するときは migration 計画を必ず書く。Phase 1 = 1。
const LESSON_SCHEMA_VERSION = 1;
// Sheets 1 cell 文字数上限 50000 に対する defensive cap。lessonJson が超えるなら save reject。
const LESSON_JSON_MAX_BYTES = 45000;
// 1 phase snapshot あたりの char 予算 (Phase 1 設計): 全 phase 合計 + meta が 45000 を超えないよう逆算。
const LESSON_PHASE_SNAPSHOT_BUDGET_BYTES = 12000;
// startLesson 内部の double-click 防止 lock。
const LESSON_START_LOCK_TIMEOUT_MS = 5000;
// Auto-archive thresholds (unpublishBoard hook + daily sweep)。
const LESSON_AUTO_ARCHIVE_MIN_MS = 5 * 60 * 1000;        // < 5 分 = テスト操作とみなして skip
const LESSON_AUTO_ARCHIVE_MIN_RESPONSES = 1;             // 0 回答 = 授業として成立してない
const LESSON_DAILY_STALE_HOURS = 4;                      // 最終回答から 4h 経過 = もう授業終了とみなす
const LESSON_DAILY_TRIGGER_HOUR = 23;                    // 23:00 JST に sweep

// ----- 内部 CRUD: lessons シートに対する row-level 操作 -----

// Why: SA-backed spreadsheet proxy の getSheetByName('lessons') は sheet 不存在でも
//   proxy オブジェクトを返してしまい、実際に getValues() するまで失敗を検出できない。
//   getSheets() で実存をチェックしてから handle を返す。
function __lessonsSheetExists_(spreadsheet) {
  try {
    const sheets = (spreadsheet.getSheets ? spreadsheet.getSheets() : []) || [];
    return sheets.some(s => {
      try { return s.getName && s.getName() === 'lessons'; }
      catch (_) { return false; }
    });
  } catch (_) { return false; }
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
  if (!sheet) return { success: false, message: 'lessons sheet not initialized' };
  const cols = __lessonColumns_(sheet);
  const serialized = __serializeLessonJson_(record.lessonJson);
  sheet.appendRow(__buildLessonRow_(record, cols, serialized));
  return { ...record, sizeBytes: serialized.sizeBytes, etag: serialized.etag, schemaVersion: LESSON_SCHEMA_VERSION };
}

// 既存 row を patch。patch は {state?, name?, startedAt?, endedAt?, lessonJson?}。
//   batch write: 1 setValues 呼び出しで完結 (CLAUDE.md batch ルール)。
//   lessonJson 含む patch なら etag / sizeBytes も更新、そうでないなら lessonJson 列は触らず etag のみ。
//
//   Why LockService: 旧コードは read-modify-write を unlocked で行っていたため、
//   advanceLessonPhase / endLesson / updateLessonDraft の同時クリックで write が lost
//   する data-loss race があった。LockService.getDocumentLock() で lessons sheet 全体を
//   serialize する (細粒度 lock は LessonId 単位の lock service が無いため避ける)。
function __updateLessonRow_(lessonId, patch, expectedEtag) {
  // ScriptLock (script-wide) for serializing all lesson sheet writes.
  // GAS LockService doesn't support per-key locking, so script-wide is the only option.
  const lock = LockService.getScriptLock();
  let locked = false;
  try {
    locked = lock.tryLock(5000);
    if (!locked) {
      return { success: false, error: 'LOCK_TIMEOUT', message: '別の処理が実行中です。少し待ってから再試行してください。' };
    }

    const found = __findLessonById_(lessonId);
    if (!found) return { success: false, error: 'LESSON_NOT_FOUND', message: 'lesson not found' };
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
  } finally {
    if (locked) {
      try { lock.releaseLock(); } catch (e) { console.warn('__updateLessonRow_ releaseLock failed:', e.message); }
    }
  }
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
  //   setValues が「別の lesson 行」を書き換える data-loss を起こす。__updateLessonRow_ と
  //   同じ getScriptLock() で serialize する。
  const lock = LockService.getScriptLock();
  let locked = false;
  try {
    locked = lock.tryLock(5000);
    if (!locked) {
      return { success: false, error: 'LOCK_TIMEOUT', message: '別の処理が実行中です。少し待ってから再試行してください。' };
    }

    const found = __findLessonById_(lessonId);
    if (!found) return { success: false, error: 'LESSON_NOT_FOUND' };
    // Why: SA proxy は deleteRow を持たない。admin は DB シートの editor 共有を受けているので
    //   SpreadsheetApp 経由で直接削除する。
    const dbId = typeof getCachedProperty === 'function' ? getCachedProperty('DATABASE_SPREADSHEET_ID') : null;
    if (!dbId) return { success: false, error: 'DATABASE_NOT_CONFIGURED' };
    const ss = SpreadsheetApp.openById(dbId);
    const sheet = ss.getSheetByName('lessons');
    if (!sheet) return { success: false, error: 'LESSONS_SHEET_NOT_FOUND' };
    sheet.deleteRow(found.rowIndex);
    return { success: true };
  } catch (error) {
    logError_('__deleteLessonRow_', error);
    return { success: false, error: 'DELETE_FAILED', message: error && error.message };
  } finally {
    if (locked) {
      try { lock.releaseLock(); } catch (e) { console.warn('__deleteLessonRow_ releaseLock failed:', e.message); }
    }
  }
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
      return { error: createErrorResponse('他ユーザーの lesson にはアクセスできません') };
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
  }
};

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
      phases: tpl.phases.map(p => ({
        name: p.name,
        formTemplate: p.formTemplate,
        question: p.question,
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

function listLessons(userId) {
  try {
    const email = getCurrentEmail();
    if (!email) return createAuthError();
    const callerUser = findUserByEmail(email, { requestingUser: email });
    if (!callerUser) return createUserNotFoundError();

    const isAdmin = isAdministrator(email);
    if (!isAdmin && callerUser.userId !== userId) {
      return createErrorResponse('他ユーザーの lesson 一覧にはアクセスできません');
    }

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
    const email = getCurrentEmail();
    if (!email) return createAuthError();
    const callerUser = findUserByEmail(email, { requestingUser: email });
    if (!callerUser) return createUserNotFoundError();

    const isAdmin = isAdministrator(email);
    if (!isAdmin && callerUser.userId !== userId) {
      return createErrorResponse('他ユーザーの class 一覧にはアクセスできません');
    }

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
    return createSuccessResponse('loaded', { lesson: auth.found.lesson });
  } catch (error) {
    logError_('getLessonForReview', error);
    return createExceptionResponse(error);
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

// ----- Migration: 既存 profiles[] を「過去の lesson 記録」として lessons シートに取り込む -----

// boardMode → formTemplate の逆引き (importLessonFromProfiles 用)。
//   __templateToBoardMode_ の inverse だが、boardMode の方が表現域が広い (wordcloud 等)
//   ので「最も近い formTemplate」を返す。
function __boardModeToFormTemplate_(boardMode) {
  if (boardMode === 'pie') return 'pie';
  if (boardMode === 'numberline') return 'numberline';
  if (boardMode === 'matrix') return 'matrix';
  if (boardMode === 'board') return 'board';
  // wordcloud / auto などはテンプレ既存形式に無いので 'board' (自由記述ベース) にマップ。
  return 'board';
}

// profile の軸ラベル / 象限ラベル / pieOptions を templateOptions 形に詰め替える。
//   __buildPhaseConfigPatch_ の inverse 相当。lesson の review/replay 表示と整合させる。
function __profileToTemplateOptions_(profile) {
  if (!profile) return {};
  const opts = {};
  const ds = profile.displaySettings || {};
  const xAxis = profile.xAxisLabels || ds.xAxisLabels || null;
  const yAxis = profile.yAxisLabels || ds.yAxisLabels || null;
  const quadrant = profile.matrixQuadrantLabels || ds.matrixQuadrantLabels || null;
  const mode = ds.boardMode || 'auto';
  if (mode === 'numberline' && xAxis) {
    opts.lowLabel = xAxis.min || '';
    opts.highLabel = xAxis.max || '';
  }
  if (mode === 'matrix') {
    if (xAxis) { opts.xLow = xAxis.min || ''; opts.xHigh = xAxis.max || ''; }
    if (yAxis) { opts.yLow = yAxis.min || ''; opts.yHigh = yAxis.max || ''; }
    if (quadrant) opts.matrixQuadrantLabels = quadrant;
  }
  return opts;
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
//   exportBoardData (AdminApis.js) と __buildImportSnapshotRows_ で共用。
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

// answer / reason を cap 文字に切り詰める。row を mutate し、切り詰めたかを返す。
function __truncateRowAnswerReason_(r, cap) {
  let truncated = false;
  if (typeof r.answer === 'string' && r.answer.length > cap) {
    r.answer = r.answer.slice(0, cap) + '…';
    truncated = true;
  }
  if (typeof r.reason === 'string' && r.reason.length > cap) {
    r.reason = r.reason.slice(0, cap) + '…';
    truncated = true;
  }
  return truncated;
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

// 1 phase 分の rows を privacy-stripped + budget-truncated な snapshot rows に整形。
//   __captureSnapshot_ と budget は揃える (LESSON_PHASE_SNAPSHOT_BUDGET_BYTES)。
function __buildImportSnapshotRows_(rawRows) {
  const slim = (rawRows || []).map(r => __projectBoardRowForExport_(r));
  let total = 0;
  let truncated = false;
  for (let i = 0; i < slim.length; i++) {
    const r = slim[i];
    let sz = JSON.stringify(r).length;
    if (total + sz > LESSON_PHASE_SNAPSHOT_BUDGET_BYTES) {
      if (__truncateRowAnswerReason_(r, 60)) {
        truncated = true;
        sz = JSON.stringify(r).length;
      }
    }
    total += sz;
  }
  return { rows: slim, truncated };
}

const LESSON_SHRINK_STAGES = Object.freeze({
  CAP80: 'cap80',
  CAP40: 'cap40',
  ROW_DROP: 'rowDrop',
  META_ONLY: 'metaOnly'
});

// lessonJson 全体が LESSON_JSON_MAX_BYTES を超えていたら snapshots を段階的に縮めて収める。
//   段階: ① answer/reason 80 char cap → ② 40 char cap → ③ 末尾から行を間引き → ④ snapshots 空。
//   Why: 振り返り wordcloud のように 1 行 200-700 char で cell 上限 (50000) 超過するケースを救う。
// 結果は lessonJson を mutate し、shrink した場合は `meta.shrinkStage` を直接書き込む
//   (caller は単に呼ぶだけでよく、後から meta を書き戻す必要なし)。
// パフォーマンス: row drop は per-snapshot サイズキャッシュで `JSON.stringify(snapshot)` を 1 iter / 1 回に。
//   全件 stringify は cap 段階と初回 measure のみ。
function __shrinkLessonJsonToFit_(lessonJson) {
  const measure = () => JSON.stringify(lessonJson || {}).length;
  if (measure() <= LESSON_JSON_MAX_BYTES) return { shrunk: false };

  const snapshots = (lessonJson && Array.isArray(lessonJson.snapshots)) ? lessonJson.snapshots : [];
  if (snapshots.length === 0) return { shrunk: false };
  if (!lessonJson.meta) lessonJson.meta = {};

  const truncateAllRowsToCap = (cap) => {
    snapshots.forEach((s) => {
      (s.rows || []).forEach((r) => {
        if (__truncateRowAnswerReason_(r, cap)) s.truncated = true;
      });
    });
  };

  truncateAllRowsToCap(80);
  if (measure() <= LESSON_JSON_MAX_BYTES) {
    lessonJson.meta.shrinkStage = LESSON_SHRINK_STAGES.CAP80;
    return { shrunk: true };
  }

  truncateAllRowsToCap(40);
  if (measure() <= LESSON_JSON_MAX_BYTES) {
    lessonJson.meta.shrinkStage = LESSON_SHRINK_STAGES.CAP40;
    return { shrunk: true };
  }

  const snapshotSizes = snapshots.map(s => JSON.stringify(s).length);
  let totalSize = measure();
  let guard = 5000;
  while (totalSize > LESSON_JSON_MAX_BYTES && guard-- > 0) {
    let biggestIdx = -1;
    let biggestSize = -1;
    for (let i = 0; i < snapshots.length; i++) {
      if ((snapshots[i].rows || []).length > 0 && snapshotSizes[i] > biggestSize) {
        biggestSize = snapshotSizes[i];
        biggestIdx = i;
      }
    }
    if (biggestIdx < 0) break;
    const target = snapshots[biggestIdx];
    target.rows.pop();
    target.droppedRows = (target.droppedRows || 0) + 1;
    target.truncated = true;
    const newSize = JSON.stringify(target).length;
    totalSize -= (snapshotSizes[biggestIdx] - newSize);
    snapshotSizes[biggestIdx] = newSize;
  }
  if (totalSize <= LESSON_JSON_MAX_BYTES) {
    lessonJson.meta.shrinkStage = LESSON_SHRINK_STAGES.ROW_DROP;
    return { shrunk: true };
  }

  snapshots.forEach((s) => {
    s.droppedRows = (s.rowCount || 0);
    s.rows = [];
    s.truncated = true;
    s.reason = (s.reason ? s.reason + ';' : '') + 'OVERSIZE_SHRINK';
  });
  lessonJson.meta.shrinkStage = LESSON_SHRINK_STAGES.META_ONLY;
  return { shrunk: true };
}

/**
 * 既存 user config の profiles[] を lessons シートに「完了済み授業」として取り込む。
 *
 * 用途: lesson 機能 (Phase 1+2) 導入以前に plain profiles として運用していた授業を、
 *       事後的に授業履歴として残したいときに使う。
 *
 * 動作:
 *   - profiles[] 各エントリを phase に変換 (name / formUrl / spreadsheetId / columnMapping / displaySettings 引き継ぎ)
 *   - boardMode から formTemplate を逆引き
 *   - profileHistory から profileTransitions を再構成 (= 授業中のフェーズ遷移ログ)
 *   - 各 phase のデータを getPublishedSheetDataForProfile で取得し、name フィールドを除外して snapshot 化
 *   - state='completed' で createdAt / startedAt = profileHistory 最古、endedAt = 最新
 *
 * 制約:
 *   - 既に lessons シートに同 profile 構成の record があっても dedup しない (呼び出し側責任)
 *   - lessonJson 全体が 45000 bytes を超えるとエラー (Sheets 1cell 50000 char 上限対策)
 *   - snapshot rows は LESSON_PHASE_SNAPSHOT_BUDGET_BYTES の予算内で answer/reason truncate
 *
 * @param {string} userId - 対象ユーザー (= 呼び出し元 = profiles 所有者)
 * @param {Object} [options]
 * @param {string} [options.name] - lesson 名 (省略時は 1 つ目の profile の formTitle から自動生成)
 * @param {boolean} [options.includeSnapshots=true] - false にすると metadata のみ取り込み (size 圧縮用)
 * @returns {Object} { success, message, data: { lesson } }
 */
function importLessonFromProfiles(userId, options) {
  try {
    const auth = __requireLessonOwner_(userId, null);
    if (auth.error) return auth.error;

    const config = getConfigOrDefault(userId);
    const profiles = Array.isArray(config.profiles) ? config.profiles : [];
    if (profiles.length === 0) {
      return createErrorResponse('プロファイルが設定されていません。取り込み対象がありません。');
    }

    const opts = options || {};
    const includeSnapshots = opts.includeSnapshots !== false;
    const history = Array.isArray(config.profileHistory) ? config.profileHistory : [];

    // phases: profiles[] と同じ順序 (= 教師が AdminPanel で並べた順)
    const phases = profiles.map((p) => ({
      name: p.name || '',
      formTemplate: __boardModeToFormTemplate_(p.displaySettings && p.displaySettings.boardMode),
      question: p.formTitle || p.name || '',
      formId: __extractFormPublishedId_(p.formUrl),
      formUrl: p.formUrl || '',
      spreadsheetId: p.spreadsheetId || '',
      sheetName: p.sheetName || '',
      columnMapping: p.columnMapping || {},
      displaySettings: p.displaySettings || {},
      templateOptions: __profileToTemplateOptions_(p)
    }));

    // profileTransitions: profileHistory から { ts, from, to } に再構成
    const nameToIdx = {};
    phases.forEach((ph, i) => { nameToIdx[ph.name] = i; });
    const transitions = [];
    let prev = null;
    for (const h of history) {
      if (!h || typeof h.name !== 'string') continue;
      const idx = nameToIdx[h.name];
      if (idx === undefined) continue;
      transitions.push({ ts: h.activatedAt || '', from: prev, to: idx });
      prev = idx;
    }

    const snapshots = phases.map((ph, idx) => {
      const base = {
        phaseIndex: idx,
        phaseName: ph.name,
        capturedAt: new Date().toISOString(),
        boardMode: (ph.displaySettings && ph.displaySettings.boardMode) || 'auto',
        columnMapping: ph.columnMapping || {},
        displaySettings: ph.displaySettings || {},
        rows: [],
        rowCount: 0,
        truncated: false
      };
      if (!includeSnapshots) {
        base.reason = 'SKIPPED:includeSnapshots=false';
        return base;
      }
      let fetched;
      try {
        fetched = getPublishedSheetDataForProfile(userId, ph.name, null, 'newest');
      } catch (e) {
        base.reason = 'CAPTURE_FAILED:' + ((e && e.message) || 'exception');
        return base;
      }
      if (!fetched || fetched.success === false || !Array.isArray(fetched.data)) {
        base.reason = 'CAPTURE_FAILED:' + ((fetched && fetched.error) || 'no_data');
        return base;
      }
      const built = __buildImportSnapshotRows_(fetched.data);
      base.rows = built.rows;
      base.rowCount = built.rows.length;
      base.truncated = built.truncated;
      return base;
    });

    const lessonId = 'lesson_' + Utilities.getUuid().slice(0, 12);
    const now = new Date().toISOString();
    const earliest = transitions.length > 0 ? (transitions[0].ts || now) : now;
    const latest = transitions.length > 0 ? (transitions[transitions.length - 1].ts || now) : now;
    const autoName = (profiles[0] && profiles[0].formTitle) ? `[移行] ${profiles[0].formTitle}` : '[移行] 過去の授業';
    const lessonName = String(opts.name || autoName).slice(0, 100);

    const lessonJson = {
      template: 'imported',
      classes: __extractClassesFromSnapshots_(snapshots),
      phases,
      profileTransitions: transitions,
      snapshots,
      meta: {
        schemaVersion: LESSON_SCHEMA_VERSION,
        importedFromProfiles: true,
        importedAt: now,
        originalActiveProfile: config.activeProfile || null,
        profileNames: profiles.map(p => p.name).filter(Boolean)
      }
    };

    // セルサイズ上限 (50000 char) に収まるまで snapshots を段階的に縮める。
    //   原本データは元 spreadsheet に残るので、ここでは「過去授業の概形」を保てれば十分。
    //   shrink 関数自身が lessonJson.meta.shrinkStage を書き込むので caller は意識不要。
    const shrinkInfo = __shrinkLessonJsonToFit_(lessonJson);

    const record = {
      lessonId,
      userId,
      name: lessonName,
      state: 'completed',
      createdAt: earliest,
      startedAt: earliest,
      endedAt: latest,
      lessonJson
    };

    const written = __createLessonRow_(record);
    if (written && written.success === false) {
      return createErrorResponse(written.message || 'lesson 作成失敗');
    }
    return createSuccessResponse('lesson を過去授業として記録しました', {
      lesson: written,
      shrunk: shrinkInfo.shrunk,
      shrinkStage: lessonJson.meta.shrinkStage || null,
      sizeBytes: written.sizeBytes
    });
  } catch (error) {
    logError_('importLessonFromProfiles', error);
    return createExceptionResponse(error);
  }
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

// Form の回答受付を on/off。テスト sandbox では FormApp が無いので silently fall through。
function __setFormAcceptingResponses_(formId, accepting) {
  if (!formId) return false;
  try {
    if (typeof FormApp === 'undefined' || !FormApp.openById) return false;
    FormApp.openById(formId).setAcceptingResponses(Boolean(accepting));
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

function __buildPhaseConfigPatch_(phase, lessonJson, lessonId) {
  // displaySettings が phase に明示指定されていればそれを優先、
  //   無ければ formTemplate から決定。templateOptions から axis ラベルも反映。
  const baseDisplay = phase.displaySettings || {};
  const opts = phase.templateOptions || {};
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
  if (xAxisLabels) patch.xAxisLabels = xAxisLabels;
  if (yAxisLabels) patch.yAxisLabels = yAxisLabels;
  return patch;
}

// Phase boundary で「その時点の rows + reactions + columnMapping」を freeze。
//   - rows は getPublishedSheetData 経由で取得 (live board と同じ projection を再利用)。
//   - reactions は deep-copy で freeze (live row.reactions の以後の mutate を遮断)。
//   - サイズ超過時は answer/reason を切り詰めて truncated=true フラグ (lesson 全体を失わない)。
//   - 失敗時は reason 付きの空 snapshot を残す (endLesson が常に完了するためのフォールバック)。
//   - 同 phaseIndex の snapshot が既にあれば replace (advance→end の double capture 対策)。
function __captureSnapshot_(userId, lessonJson, phaseIdx) {
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

  const frozen = deepClone(rawRows);
  let total = 0;
  let truncated = false;
  for (let i = 0; i < frozen.length; i++) {
    const r = frozen[i];
    let rowSize = JSON.stringify(r).length;
    if (total + rowSize > LESSON_PHASE_SNAPSHOT_BUDGET_BYTES) {
      const remaining = Math.max(0, LESSON_PHASE_SNAPSHOT_BUDGET_BYTES - total);
      if (typeof r.answer === 'string' && r.answer.length > 40) {
        r.answer = r.answer.slice(0, Math.max(20, Math.floor(remaining / 2))) + '…';
      }
      if (typeof r.reason === 'string' && r.reason.length > 40) {
        r.reason = r.reason.slice(0, Math.max(20, Math.floor(remaining / 2))) + '…';
      }
      truncated = true;
      rowSize = JSON.stringify(r).length;
    }
    total += rowSize;
  }

  baseSnapshot.rows = frozen;
  baseSnapshot.rowCount = frozen.length;
  baseSnapshot.truncated = truncated;
  return baseSnapshot;
}

// snapshots[] への upsert: 同 phaseIndex があれば replace、なければ append。
//   phaseIndex 昇順を維持 (replay の slider が時系列で歩けるように)。
function __upsertSnapshot_(lessonJson, snapshot) {
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
      const sharedClasses = (lessonJson && Array.isArray(lessonJson.classes)) ? lessonJson.classes : [];
      // 揺らぎ追跡: lesson レベルの allowResubmit を全 phase に適用 (議論前後の意見変化を取りたい)。
      //   phase 個別に templateOptions.allowResubmit があればそれを優先。
      const lessonAllowResubmit = Boolean(lessonJson && lessonJson.allowResubmit);
      for (let i = 0; i < phases.length; i++) {
        const phase = phases[i];
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
        __setFormAcceptingResponses_(phase.formId, i === 0);

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
        try { lock.releaseLock(); } catch (_) {}
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
function advanceLessonPhase(userId, lessonId, direction) {
  try {
    const auth = __requireLessonOwner_(userId, lessonId);
    if (auth.error) return auth.error;
    const { found } = auth;

    if (found.lesson.state !== 'active') {
      return createErrorResponse('FORBIDDEN_STATE: lesson が active でないため進められません');
    }

    const dir = direction === 'previous' ? 'previous' : 'next';
    const lessonJson = deepClone(found.lesson.lessonJson || {});
    const phases = lessonJson.phases || [];
    const fromIdx = __activePhaseIndex_(lessonJson);
    const toIdx = dir === 'next' ? fromIdx + 1 : fromIdx - 1;

    if (toIdx < 0) return createErrorResponse('既に最初のフェーズです');
    if (toIdx >= phases.length) return createErrorResponse('既に最後のフェーズです (終了するには「⏹ 終了」を押してください)');

    // Why: 移行 *前* に outgoing phase の rows を freeze する。順序を逆にすると
    //   user config が次 phase の columnMapping を指した状態で capture することになり、
    //   replay が破綻する。capture は config 切替より前 (= 現状 fromIdx) で行う。
    __upsertSnapshot_(lessonJson, __captureSnapshot_(userId, lessonJson, fromIdx));

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
    __setFormAcceptingResponses_(phases[fromIdx].formId, false);
    __setFormAcceptingResponses_(phases[toIdx].formId, true);

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
    phases.forEach(p => __setFormAcceptingResponses_(p.formId, false));

    // 現フェーズの rows を freeze して snapshots に upsert。capture 失敗時は reason 付き空 snapshot
    //   が積まれ、endLesson 自体は成功する (lesson は必ず completed に遷移する原則)。
    const currentIdx = __activePhaseIndex_(lessonJson);
    __upsertSnapshot_(lessonJson, __captureSnapshot_(userId, lessonJson, currentIdx));

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
