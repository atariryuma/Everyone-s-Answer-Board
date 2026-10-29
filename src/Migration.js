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

/* global isSingleTenant_, getCachedProperty, findUserByEmail, __findLessonById_, __createLessonRow_, __writeArchiveRows_, __singleNativeSheetName_, __withLessonLock_, applyConfigPatch_, republishMyBoard, advanceLessonPhase, createErrorResponse, createSuccessResponse, logError_ */

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
