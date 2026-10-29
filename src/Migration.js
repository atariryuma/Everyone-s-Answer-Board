/**
 * @fileoverview 共有の本体 (サイロ型マルチテナント) から、先生専用の配置へ引っ越す。
 *
 * Why: 2026-10-29、児童は「開いた人として実行」の Apps Script を開けなくなり、先生ごとに
 *   「先生として実行」で配置し直した (isSingleTenant_)。これまでの授業・アーカイブ・ボードの
 *   設定を新しい配置へ移す。元の DB は先生から見えない (SA にだけ共有) ので、元の本体の
 *   admin API で読み出したもの (lesson.review / getUserConfig) を、この admin API で受け取る。
 *   relay は scripts/migrate-to-single.js。
 *
 * 守っていること:
 *   - 何度流しても二重にならない (同じ lessonId は飛ばす)。途中で失敗したら、その回に写した
 *     シートは消して半端を残さない
 *   - 授業モードの回答シートは先生のスプレッドシートへ「授業名｜番号 フェーズ名」で写す
 *     (新しく作る授業と同じ名前の付け方 = __singleNativeSheetName_)。元の場所は meta.migratedFrom に残す
 *   - フォームで集めた回答シートは、フォームが回答先として掴んでいるので動かさない (参照だけ引き継ぐ)
 *   - アーカイブは lesson_responses に 1 件 1 行で書き直し、snapshot にはポインタだけ残す
 *   - 公開状態は publish lifecycle の関数 (republishMyBoard) でだけ変える
 *   - 元の本体のデータは消さない (写すだけ)
 */

/* global isSingleTenant_, getCachedProperty, findUserByEmail, __findLessonById_, __createLessonRow_, __writeArchiveRows_, __singleNativeSheetName_, __withLessonLock_, applyConfigPatch_, republishMyBoard, advanceLessonPhase, bumpBoardDataVersion_, LESSON_NATIVE_SHEET_HEADERS, createErrorResponse, createSuccessResponse, logError_ */

// 引っ越し先の持ち主 (この配置の先生)。登録がまだなら先に管理パネルを開いてもらう。
function __migrationOwner_() {
  if (!(typeof isSingleTenant_ === 'function' && isSingleTenant_())) {
    return { error: createErrorResponse('先生専用の配置でだけ使えます', null, { error: 'NOT_SINGLE_TENANT' }) };
  }
  const adminEmail = getCachedProperty('ADMIN_EMAIL');
  const user = adminEmail ? findUserByEmail(adminEmail, { requestingUser: adminEmail }) : null;
  if (!user || !user.userId) {
    return { error: createErrorResponse('先に先生のアカウントで管理パネルを 1 度開いてください (登録がまだです)', null, { error: 'OWNER_NOT_REGISTERED' }) };
  }
  return { user };
}

/**
 * 元の本体の lesson.review が返した授業 1 つを取り込む。
 * @param {Object} lesson - { lessonId, name, state, createdAt, startedAt, endedAt, lessonJson (snapshots の rows 入り) }
 * @returns {Object} { success, data: { lessonId, skipped?, sheets, archiveRows } }
 */
function migrateImportLesson(lesson) {
  const copied = [];   // この回に写したシート (失敗したら消す)
  try {
    const owner = __migrationOwner_();
    if (owner.error) return owner.error;
    if (!lesson || typeof lesson !== 'object' || !lesson.lessonId || !lesson.lessonJson || typeof lesson.lessonJson !== 'object') {
      return createErrorResponse('lesson が不正です', null, { error: 'INVALID_LESSON' });
    }
    const existing = __findLessonById_(lesson.lessonId);
    if (!existing) return createErrorResponse('lessons シートを読めませんでした', null, { error: 'LESSONS_READ_FAILED' });
    if (existing.lesson) {
      return createSuccessResponse('既に引っ越し済みです', { lessonId: lesson.lessonId, skipped: true, sheets: 0, archiveRows: 0 });
    }

    const lessonJson = JSON.parse(JSON.stringify(lesson.lessonJson));
    const ss = SpreadsheetApp.openById(getCachedProperty('DATABASE_SPREADSHEET_ID'));
    const moved = { sheets: 0, archiveRows: 0 };

    // 1. 授業モードの回答シートを先生のスプレッドシートへ写す
    if (lessonJson.inputMode === 'native' && lessonJson.nativeSpreadsheetId) {
      const src = SpreadsheetApp.openById(lessonJson.nativeSpreadsheetId);
      const from = { nativeSpreadsheetId: lessonJson.nativeSpreadsheetId, sheets: {} };
      delete lessonJson.nativeSheetPrefix;
      (lessonJson.phases || []).forEach((phase, i) => {
        if (!phase || !phase.sheetName) return;   // まだ始めていないフェーズ
        const srcSheet = src.getSheetByName(phase.sheetName);
        if (!srcSheet) throw new Error('元の回答シートが見つかりません: ' + phase.sheetName);
        const name = __singleNativeSheetName_(ss, lessonJson, i, lesson.name);
        const dest = srcSheet.copyTo(ss);
        dest.setName(name);
        copied.push(dest);
        from.sheets[i] = phase.sheetName;
        phase.spreadsheetId = ss.getId();
        phase.sheetName = name;
      });
      moved.sheets = copied.length;
      lessonJson.nativeSpreadsheetId = ss.getId();
      lessonJson.meta = Object.assign({}, lessonJson.meta || {}, { migratedFrom: from });
    }

    // 2. アーカイブ (snapshot の rows) を lesson_responses へ書き直し、ポインタだけ残す
    (lessonJson.snapshots || []).forEach((sn) => {
      const rows = Array.isArray(sn.rows) ? sn.rows : [];
      if (!rows.length && Number(sn.rowCount) > 0) {
        // 元で読み戻せなかった snapshot。空で取り込むと記録が消えるので止める。
        throw new Error('アーカイブを読み戻せていない snapshot があります (phase ' + sn.phaseIndex + ')');
      }
      if (rows.length) {
        const ptr = __writeArchiveRows_(lesson.lessonId, sn.phaseIndex, rows);
        if (!ptr) throw new Error('アーカイブを書けませんでした (phase ' + sn.phaseIndex + ')');
        sn.sheet = ptr.sheet;
        sn.startRow = ptr.startRow;
        sn.rowCount = ptr.rowCount;
        moved.archiveRows += rows.length;
      }
      sn.rows = [];
    });

    // 3. 授業の行 (持ち主はこの配置の先生)
    const record = {
      lessonId: lesson.lessonId,
      userId: owner.user.userId,
      name: String(lesson.name || '').slice(0, 100) || '授業',
      state: lesson.state || 'completed',
      createdAt: lesson.createdAt || new Date().toISOString(),
      startedAt: lesson.startedAt || null,
      endedAt: lesson.endedAt || null,
      lessonJson
    };
    const written = __withLessonLock_(() => __createLessonRow_(record));
    if (!written || written.success === false) {
      throw new Error((written && written.message) || '授業の行を書けませんでした');
    }
    return createSuccessResponse('引っ越しました', { lessonId: lesson.lessonId, skipped: false, sheets: moved.sheets, archiveRows: moved.archiveRows });
  } catch (error) {
    // 写したシートを消して半端を残さない (lesson_responses に書いた行は孤児になるが、
    //   読み出しは snapshot のポインタ経由なので誤読しない = __writeArchiveRows_ の説明と同じ)。
    copied.forEach((sheet) => {
      try { sheet.getParent().deleteSheet(sheet); } catch (_) { /* 消せなくても次の実行は「（2）」で別名になるだけ */ }
    });
    logError_('migrateImportLesson', error, { lessonId: lesson && lesson.lessonId });
    return createErrorResponse('引っ越しに失敗しました: ' + (error && error.message ? error.message : 'unknown'), null, { error: 'MIGRATION_FAILED' });
  }
}

/**
 * 別のスプレッドシートにある授業モードの回答を、この配置の授業のフェーズのシートへ足す。
 *
 * Why: 2026-10-29 の 6年2組は、先生専用の配置の前に作った試作 (授業 1 つ専用の紐づけアプリ) で
 *   「考える」「もう一度考える」を送った。その回答をいまの授業に合流させる。試作は学級を
 *   切り替えないまま使ったので学級が「6年1組」、時刻は 7 時間早く記録されていた
 *   (classMap / shiftMinutes で直す)。
 *
 * 回答は Google の中で写すだけで、応答には件数しか載せない (児童の名前・本文を返さない)。
 * 同じ児童・同じ学級の行がもうあれば足さない (何度流しても二重にならない)。
 *
 * @param {Object} params
 * @param {string} params.lessonId - 足し先の授業 (授業モード)
 * @param {string} params.sourceSpreadsheetId - 元の回答が入ったスプレッドシート
 * @param {Array<{sourceSheet: string, phaseIndex: number}>} params.sheets - 元のシート → 足し先のフェーズ
 * @param {Object} [params.classMap] - 学級の付け替え ({ '6年1組': '6年2組' })。無い学級はそのまま
 * @param {number} [params.shiftMinutes] - 時刻のずらし (分)
 * @param {boolean} [params.dryRun] - true なら数えるだけで書かない
 */
function migrateImportLessonAnswers(params) {
  try {
    const owner = __migrationOwner_();
    if (owner.error) return owner.error;
    const p = params || {};
    if (!p.lessonId || typeof p.lessonId !== 'string') return createErrorResponse('lessonId が必要です', null, { error: 'INVALID_PARAMS' });
    if (!p.sourceSpreadsheetId || typeof p.sourceSpreadsheetId !== 'string') return createErrorResponse('sourceSpreadsheetId が必要です', null, { error: 'INVALID_PARAMS' });
    if (!Array.isArray(p.sheets) || !p.sheets.length) return createErrorResponse('sheets が必要です', null, { error: 'INVALID_PARAMS' });
    const classMap = (p.classMap && typeof p.classMap === 'object') ? p.classMap : {};
    const shiftMs = Number(p.shiftMinutes || 0) * 60 * 1000;
    if (!isFinite(shiftMs)) return createErrorResponse('shiftMinutes が不正です', null, { error: 'INVALID_PARAMS' });

    const found = __findLessonById_(p.lessonId);
    const lesson = found && found.lesson;
    if (!lesson) return createErrorResponse('授業が見つかりません', null, { error: 'LESSON_NOT_FOUND' });
    const lessonJson = lesson.lessonJson || {};
    if (lessonJson.inputMode !== 'native') return createErrorResponse('授業モードの授業ではありません', null, { error: 'NOT_NATIVE' });

    const headers = LESSON_NATIVE_SHEET_HEADERS;
    const COL = { ts: 0, email: 1, cls: 2 };
    const src = SpreadsheetApp.openById(p.sourceSpreadsheetId);
    const result = [];
    let imported = 0;

    p.sheets.forEach((m) => {
      const phase = (lessonJson.phases || [])[m && m.phaseIndex];
      if (!phase || !phase.spreadsheetId || !phase.sheetName) throw new Error('足し先のフェーズにシートがありません (phase ' + (m && m.phaseIndex) + ')');
      const srcSheet = src.getSheetByName(String(m.sourceSheet || ''));
      if (!srcSheet) throw new Error('元のシートが見つかりません: ' + m.sourceSheet);
      const dest = SpreadsheetApp.openById(phase.spreadsheetId).getSheetByName(phase.sheetName);
      if (!dest) throw new Error('足し先のシートが見つかりません: ' + phase.sheetName);

      // 列は見出しの名前で合わせる (元は HIGHLIGHT の位置などが違う)。足し先は授業モードの列の並びであること。
      const destData = dest.getDataRange().getValues();
      const destHead = (destData[0] || []).slice(0, headers.length).map(String);
      if (destHead.join('|') !== headers.join('|')) throw new Error('足し先の列の並びが授業モードと違います: ' + phase.sheetName);
      const srcData = srcSheet.getDataRange().getValues();
      const srcHead = (srcData[0] || []).map(String);
      const idx = headers.map(h => srcHead.indexOf(h));
      if (idx[COL.email] < 0 || idx[COL.cls] < 0) throw new Error('元のシートにメールアドレス / クラスの列がありません: ' + m.sourceSheet);

      const keyOf = (email, cls) => String(email || '').trim().toLowerCase() + '|' + String(cls || '').trim();
      const have = new Set(destData.slice(1).map(r => keyOf(r[COL.email], r[COL.cls])));
      const latest = new Map();   // 元に同じ児童の行が複数あれば最後の行 (最新の送信) を採る = __findOwnLessonRow_ と同じ
      const counts = { phaseIndex: m.phaseIndex, sourceRows: srcData.length - 1, imported: 0, alreadyThere: 0, noEmail: 0, classes: {} };
      srcData.slice(1).forEach((r) => {
        const row = idx.map(i => (i >= 0 ? r[i] : ''));
        if (!String(row[COL.email] || '').trim()) { counts.noEmail++; return; }
        const cls = String(row[COL.cls] || '').trim();
        row[COL.cls] = Object.prototype.hasOwnProperty.call(classMap, cls) ? String(classMap[cls]) : cls;
        const t = row[COL.ts] instanceof Date ? row[COL.ts] : new Date(String(row[COL.ts] || ''));
        if (!isNaN(t.getTime())) row[COL.ts] = new Date(t.getTime() + shiftMs).toISOString();
        latest.set(keyOf(row[COL.email], row[COL.cls]), row);
      });
      const rows = [];
      latest.forEach((row, key) => {
        if (have.has(key)) { counts.alreadyThere++; return; }
        rows.push(row);
        counts.classes[row[COL.cls]] = (counts.classes[row[COL.cls]] || 0) + 1;
      });
      counts.imported = rows.length;
      if (rows.length && !p.dryRun) {
        __withLessonLock_(() => {
          dest.getRange(dest.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
        });
        imported += rows.length;
      }
      result.push(counts);
    });

    if (imported > 0) {
      try { bumpBoardDataVersion_(owner.user.userId); } catch (e) {
        console.warn('migrateImportLessonAnswers: bumpBoardDataVersion_ failed (board may be stale up to 12s):', e && e.message);
      }
    }
    return createSuccessResponse(p.dryRun ? '数えました (書いていません)' : '回答を足しました', { dryRun: Boolean(p.dryRun), sheets: result });
  } catch (error) {
    logError_('migrateImportLessonAnswers', error, { lessonId: params && params.lessonId });
    return createErrorResponse('回答を足せませんでした: ' + (error && error.message ? error.message : 'unknown'), null, { error: 'MIGRATION_FAILED' });
  }
}

/**
 * 元の本体の getUserConfig が返したボードの設定を取り込む。授業より後に呼ぶ
 * (ボードにつながっている授業を、新しいシートを指すようにかけ直すため)。
 * @param {Object} config
 */
function migrateImportConfig(config) {
  try {
    const owner = __migrationOwner_();
    if (owner.error) return owner.error;
    if (!config || typeof config !== 'object') return createErrorResponse('config が不正です', null, { error: 'INVALID_CONFIG' });
    const userId = owner.user.userId;

    // isPublished / publishedAt / userId などは applyConfigPatch_ が外す (公開は下で lifecycle 経由)。
    const saved = applyConfigPatch_(userId, config);
    if (!saved || saved.success === false) {
      return createErrorResponse('ボードの設定を保存できませんでした: ' + ((saved && saved.message) || ''), null, { error: 'CONFIG_SAVE_FAILED' });
    }

    const result = { published: false, reappliedLessonId: null };
    // ボードにつながっていた授業は、いまのフェーズをかけ直してボードの参照先を新しいシートにする。
    const boardLessonId = config.activeLessonId;
    if (boardLessonId) {
      const found = __findLessonById_(boardLessonId);
      const lesson = found && found.lesson;
      if (lesson && lesson.state === 'active') {
        const trans = (lesson.lessonJson && lesson.lessonJson.profileTransitions) || [];
        const idx = trans.length ? Number(trans[trans.length - 1].to) || 0 : 0;
        const re = advanceLessonPhase(userId, boardLessonId, null, idx);
        if (!re || re.success === false) {
          return createErrorResponse('授業をボードにかけ直せませんでした: ' + ((re && re.message) || ''), null, { error: 'REAPPLY_FAILED' });
        }
        result.reappliedLessonId = boardLessonId;
      }
    }
    if (config.isPublished) {
      const pub = republishMyBoard();
      result.published = Boolean(pub && pub.success);
    }
    return createSuccessResponse('ボードの設定を引っ越しました', result);
  } catch (error) {
    logError_('migrateImportConfig', error);
    return createErrorResponse('引っ越しに失敗しました: ' + (error && error.message ? error.message : 'unknown'), null, { error: 'MIGRATION_FAILED' });
  }
}
