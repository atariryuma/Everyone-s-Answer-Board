/**
 * @fileoverview AdminApis - 管理者 API（ユーザー一覧、フォーム作成、ボード
 *   公開ライフサイクル、列マッピング診断など）。クロスファイル依存は下の
 *   global 宣言を参照。
 */

/* global getCurrentEmail, isAdministrator, findUserById, findUserByEmail, getAllUsers, updateUser, getUserConfig, saveUserConfig, getColumnAnalysis, getPublishedSheetData, getPublishedSheetDataForProfile, createTemplateForm, customizeForm, setFormAllowResubmit, uploadLessonImage, processFormUrlInput, getForms, isValidFormUrl, applySpreadsheetSharingDefaults, listServiceAccountPool, getServiceAccountUsage, addServiceAccountToPool, addServiceAccountsToPoolBatch, reverifyServiceAccountInPool, removeServiceAccountFromPool, bumpBoardDataVersion_, createAdminRequiredError, createAuthError, createUserNotFoundError, createErrorResponse, createSuccessResponse, createExceptionResponse, requireAdmin, getConfigOrDefault, isPlainObject, createLessonDraft, updateLessonDraft, startLesson, advanceLessonPhase, endLesson, listLessons, getLessonForReview, deleteLesson, getKnownClassesForUser, duplicateLesson, listLessonTemplates, importLessonFromProfiles, __projectBoardRowForExport_, __maybeAutoArchiveLesson_, isBoardCollaborator, logError_, safeJsonParse_, sameEmail_ */


// Admin API経由での読み書きから保護する Script Properties キー。
// APIキー漏洩時のブラスト半径を最小化する目的で、
// 認証情報系（substring match）と、攻撃者に差し替えられるとサービス乗っ取り/
// 偽DBへの誘導が可能になる基幹キーを明示的にブロックする。
// 'APP_DISABLED' substring は緊急停止フラグとその監査メタ (APP_DISABLED_REASON/BY/AT) を
//   一括ブロックする。 これらは enableApp/disableApp の専用 op (監査証跡付き) でのみ変更させ、
//   generic setProperty による証跡バイパスを防ぐ。
const PROTECTED_PROPERTY_SUBSTRINGS = ['KEY', 'CREDS', 'SECRET', 'TOKEN', 'PASSWORD', 'SALT', 'APP_DISABLED'];
const PROTECTED_PROPERTY_EXACT_KEYS = ['ADMIN_EMAIL', 'DATABASE_SPREADSHEET_ID', 'DEPLOYED_WEB_APP_URL'];

// Magic-number elimination (Clean Code: replace literal numbers with named constants).
const PROFILE_NAME_MAX_LEN = 50;              // saveProfile: profile name 最大長
const SHEET_NAME_MAX_LEN = 200;               // Forms / Sheets で実用的に許容される sheet 名長
const DRIVE_ERROR_BODY_PREVIEW_LEN = 300;     // Drive API エラーレスポンスのログ表示文字数
const MAX_BULK_INSERT_ROWS = 500;             // bulkInsertRows API 1 回あたりの最大行数
const PROTECTED_PROPERTY_MIN_LENGTH = 4;      // Property name の最低長 (短すぎる input を reject)

function isProtectedPropertyKey(key) {
  if (typeof key !== 'string' || !key) return false;
  const upper = key.toUpperCase();
  if (PROTECTED_PROPERTY_EXACT_KEYS.includes(upper)) return true;
  return PROTECTED_PROPERTY_SUBSTRINGS.some((s) => upper.includes(s));
}

/**
 * 破壊的 admin op (setSheetHeader/clearDataRows/appendRows) が任意の spreadsheetId を
 * 受け取り、 deploy 主の OAuth で読み書きできてしまう blast radius を絞る。
 * app 管理下のボード SS、 または DB SS のみ許可する。
 * @returns {Object|null} 不許可なら error response、 許可なら null
 */
function assertManagedSpreadsheet_(spreadsheetId) {
  try {
    const dbId = getCachedProperty('DATABASE_SPREADSHEET_ID');
    if (dbId && spreadsheetId === dbId) return null;
    const boardIds = (typeof getAllBoardSpreadsheetIds === 'function') ? getAllBoardSpreadsheetIds() : [];
    if (Array.isArray(boardIds) && boardIds.indexOf(spreadsheetId) !== -1) return null;
    return createErrorResponse('指定された spreadsheetId は app 管理下のボードではありません (操作を拒否)');
  } catch (err) {
    logError_('assertManagedSpreadsheet_', err);
    // 検証自体が失敗したら安全側 (拒否) に倒す
    return createErrorResponse('spreadsheetId の検証に失敗しました');
  }
}


function getAdminUsers() {
  try {
    const auth = requireAdmin();
    if (!auth) return createAdminRequiredError();

    const users = getAllUsers({ activeOnly: false }, { forceServiceAccount: true });
    return {
      success: true,
      users: users || []
    };
  } catch (error) {
    logError_('getAdminUsers', error);
    return createExceptionResponse(error);
  }
}

/**
 * Toggle user active status for admin - system admin only
 * @param {string} targetUserId - Target user ID
 * @returns {Object} Toggle result
 */
function toggleUserActiveStatus(targetUserId) {
  try {
    const auth = requireAdmin();
    if (!auth) return createAdminRequiredError();
    const email = auth.email;

    const targetUser = findUserById(targetUserId, { requestingUser: email });
    if (!targetUser) {
      return createUserNotFoundError();
    }

    // Diff-only update so concurrent configJson writes aren't clobbered.
    const newIsActive = !targetUser.isActive;
    const result = updateUser(targetUserId, { isActive: newIsActive });
    if (result.success) {
      return {
        success: true,
        message: `ユーザー状態を${newIsActive ? 'アクティブ' : '非アクティブ'}に変更しました`,
        userId: targetUserId,
        newStatus: newIsActive,
        timestamp: new Date().toISOString()
      };
    } else {
      return createErrorResponse('ユーザー状態の更新に失敗しました');
    }
  } catch (error) {
    logError_('toggleUserActiveStatus', error);
    return createExceptionResponse(error);
  }
}

// =====================================================================
// 公開ライフサイクル: __applyPublishStateChange を通じて isPublished を書く 4 操作
// （publishApp は SystemController.js 側）。詳細仕様: CLAUDE.md の同見出し節。
// =====================================================================

/**
 * 公開状態変更の統一ヘルパー。
 * publishedAt のセマンティクスを「常に現在の状態を反映」に統一する。
 *
 * @param {string|null} targetUserId - 対象ユーザーID（null = 現在のユーザー）
 * @param {boolean|'toggle'} newState - true=公開 / false=非公開 / 'toggle'=反転
 * @param {Object} [options]
 * @param {string} [options.sourceEtag] - 楽観ロック用 etag。指定時は不一致で reject
 * @param {boolean} [options.requireAdmin] - true なら管理者以外は reject（toggleUserBoard 用）
 * @returns {Object} { success, message, userId, isPublished, publishedAt, etag, redirectUrl }
 */
function __applyPublishStateChange(targetUserId, newState, options = {}) {
  const email = getCurrentEmail();
  if (!email) return createAuthError();

  // 対象ユーザーの特定（指定があればそれ、なければ自分）
  let targetUser = targetUserId ? findUserById(targetUserId, { requestingUser: email }) : null;
  if (!targetUser) {
    targetUser = findUserByEmail(email, { requestingUser: email });
  }
  if (!targetUser) return createUserNotFoundError();

  const isAdmin = isAdministrator(email);
  const isOwnBoard = sameEmail_(targetUser.userEmail, email);

  if (options.requireAdmin && !isAdmin) {
    return createErrorResponse('この操作には管理者権限が必要です');
  }
  // v2855+: collaborator (ボード SS の editor) は unpublish (緊急停止) のみ許可。
  //   publish 再開や toggle は owner / admin のみ — 「閉じる方向」 だけ collaborator OK。
  let isCollaborator = false;
  if (!isAdmin && !isOwnBoard) {
    const isUnpublishingOnly = (newState === false)
      && typeof isBoardCollaborator === 'function'
      && isBoardCollaborator(targetUser, email);
    if (isUnpublishingOnly) {
      isCollaborator = true;
    } else {
      return createErrorResponse('ボードの公開状態を変更する権限がありません');
    }
  }

  // Why getUserConfig (not getConfigOrDefault): saveUserConfig 前の read で「failure を empty
  //   {} で隠す」は危険。findUserById が 429/cache miss を返すと {} になり、それを spread して
  //   saveUserConfig するとユーザーの profile/columnMapping/displaySettings を空で全上書きする
  //   (data-loss バグ)。getConfigOrDefault は read-only 用途のみで使う。
  const cfgResult = getUserConfig(targetUser.userId, targetUser);
  if (!cfgResult || !cfgResult.success) {
    return createErrorResponse(`設定の取得に失敗しました: ${(cfgResult && cfgResult.message) || '詳細不明'}`);
  }
  // 破損 config への publish 状態 write は {...currentConfig} で default を保存し原本を
  //   失うため中止する。
  if (cfgResult.corrupted) {
    return createErrorResponse('設定データが破損しているため公開状態を変更できません。設定の復旧が必要です。');
  }
  const currentConfig = cfgResult.config;

  // 楽観ロック: 呼び出し側が sourceEtag を渡したときだけ厳格チェック
  if (options.sourceEtag && currentConfig.etag && currentConfig.etag !== options.sourceEtag) {
    return {
      success: false,
      error: 'etag_mismatch',
      message: '設定が他の場所で更新されています。画面を再読み込みしてください。',
      currentEtag: currentConfig.etag,
      currentConfig
    };
  }

  // 目標状態の決定
  const targetIsPublished = (newState === 'toggle')
    ? !currentConfig.isPublished
    : Boolean(newState);
  const wasPublished = currentConfig.isPublished === true;
  const now = new Date().toISOString();

  const updatedConfig = { ...currentConfig };
  updatedConfig.isPublished = targetIsPublished;
  // publishedAt は「現在の状態の発生時刻」を反映。
  // 公開中は now、非公開化時は null。これにより 4 関数間でセマンティクスが揃う。
  updatedConfig.publishedAt = targetIsPublished ? now : null;
  updatedConfig.lastAccessedAt = now;

  // unpublishBoard 経由で呼ばれて、かつ lesson の auto-archive 条件を満たすなら
  //   この同じ save で endLesson + marker クリアを inline 実行 (一貫性のため同 write 内)。
  if (!targetIsPublished && options.triggerLessonAutoArchive && currentConfig.currentLessonStartedAt) {
    const archiveResult = __maybeAutoArchiveLesson_(targetUser, currentConfig);
    if (archiveResult.archived) {
      updatedConfig.currentLessonStartedAt = null;
      updatedConfig.activeLessonId = null;
    }
  }

  // saveUserConfig は updatedConfig が継承した currentConfig.etag で楽観ロックを行う
  //   (read=L158 と save の間に別 writer が config を変えていれば etag_mismatch)。
  //   この conflict を汎用エラーに埋もれさせず、 publishApp と同じく構造化して surface し、
  //   フロントが再読み込みを促せるようにする (caller に sourceEtag を強制せず非破壊で対称化)。
  const saveResult = saveUserConfig(targetUser.userId, updatedConfig);
  if (!saveResult.success) {
    if (saveResult.error === 'etag_mismatch') {
      return {
        success: false,
        error: 'etag_mismatch',
        message: '設定が他の場所で更新されています。画面を再読み込みしてください。',
        currentEtag: (saveResult.currentConfig && saveResult.currentConfig.etag) || null,
        currentConfig: saveResult.currentConfig
      };
    }
    return createErrorResponse(`ボード状態の更新に失敗しました: ${saveResult.message || '詳細不明'}`);
  }

  // 2 階層の cache を即時 stale 化:
  //   1. sa_validation cache (60s TTL): 未公開ボードへの SA pool access を即 block
  //   2. board data cache (12s TTL): viewer の polling が見ている stale data を即更新
  // 状態が実際に変わったときのみ bump (toggle no-op で 100+ viewer の cache を捨てないため)。
  if (wasPublished !== targetIsPublished) {
    if (currentConfig.spreadsheetId) {
      invalidateSaValidationCache_(currentConfig.spreadsheetId);
    }
    if (Array.isArray(currentConfig.profiles)) {
      for (const p of currentConfig.profiles) {
        if (p && p.spreadsheetId) invalidateSaValidationCache_(p.spreadsheetId);
      }
    }
    // board data cache は userId 単位 (= ボード単位)。 全 profile / 全 filter / 全 sort が
    // 1 度の bump で stale 化される。 typeof check は GAS multi-file env で同名関数が別ファイル
    // (DataApis.js) にあり、 test 単独 load 時に未定義の場合を吸収する。
    if (typeof bumpBoardDataVersion_ === 'function') {
      try { bumpBoardDataVersion_(targetUser.userId); } catch (_) { /* ignore */ }
    }
  }

  const redirectUrl = getWebAppUrl() + '?mode=view&userId=' + targetUser.userId;
  return {
    success: true,
    message: targetIsPublished
      ? (wasPublished ? 'ボードを再び公開しました' : 'ボードを公開しました')
      : (wasPublished ? 'ボードの公開を終了しました' : 'ボードは既に公開されていません'),
    userId: targetUser.userId,
    isPublished: targetIsPublished,
    boardPublished: targetIsPublished,  // back-compat alias
    publishedAt: updatedConfig.publishedAt,
    etag: saveResult.etag || null,
    redirectUrl
  };
}

/**
 * 管理者が他ユーザーのボード公開状態を toggle する。
 * @param {string} targetUserId - 対象ユーザーID（必須）
 * @param {Object} [options] - { sourceEtag }
 * @returns {Object} __applyPublishStateChange の戻り値
 */
function toggleUserBoardStatus(targetUserId, options = {}) {
  try {
    if (!targetUserId) return createErrorResponse('targetUserId is required');
    return __applyPublishStateChange(targetUserId, 'toggle', {
      sourceEtag: options.sourceEtag,
      requireAdmin: true
    });
  } catch (error) {
    logError_('toggleUserBoardStatus', error);
    return createExceptionResponse(error);
  }
}

/**
 * 自分のボードを再公開（所有者用）。
 * @param {Object} [options] - { sourceEtag }
 * @returns {Object} __applyPublishStateChange の戻り値
 */
function republishMyBoard(options = {}) {
  try {
    return __applyPublishStateChange(null, true, { sourceEtag: options.sourceEtag });
  } catch (error) {
    logError_('republishMyBoard', error);
    return createExceptionResponse(error);
  }
}

/**
 * ボードを非公開にする（所有者 or 管理者）。
 * @param {string} [targetUserId] - 対象ユーザーID（省略時は自分）
 * @param {Object} [options] - { sourceEtag }
 * @returns {Object} __applyPublishStateChange の戻り値
 */
function unpublishBoard(targetUserId, options = {}) {
  try {
    return __applyPublishStateChange(targetUserId || null, false, {
      sourceEtag: options.sourceEtag,
      // unpublish 時 = 教師の「授業終了」意思表示。lesson auto-archive を同 save で実行する。
      triggerLessonAutoArchive: options.triggerLessonAutoArchive !== false
    });
  } catch (error) {
    logError_('unpublishBoard', error);
    return createExceptionResponse(error);
  }
}

// =====================================================================
// Profile (multi-board) shared core. AdminApis dispatcher cases と
// DataApis のオーナー向けラッパーが共通で使う実装。
// 直接呼ばず、必ず認可済みの userId を渡すこと。
// =====================================================================

/**
 * profileHistory に新しい active 切替を追記して返す純粋関数。
 *
 * Why (Option B): 生徒が過去フェーズを閲覧できるようにするため、教師が active を切替えるたびに
 *   履歴を append する。3 つの切替経路 (loadProfile / loadProfileForView / __saveProfileCore
 *   autoActivate) 全てで共通利用するため core helper として切り出し。
 *
 * 正規化ルール:
 *   - 直前が同名なら no-op（重複防止: 同じ profile を 2 度連続 click した場合を吸収）
 *   - 末尾が最新。古い entry は ConfigService の sanitizeProfileHistory が cap する。
 *
 * @param {Array} currentHistory - 既存 profileHistory（不在なら [] と等価）
 * @param {string} profileName - 新たに active になった profile 名
 * @returns {Array<{name:string, activatedAt:string}>} 追記後の配列（新配列。元は mutate しない）
 */
function __appendProfileHistory_(currentHistory, profileName) {
  const hist = Array.isArray(currentHistory) ? currentHistory.slice() : [];
  const cleanName = String(profileName == null ? '' : profileName).trim();
  if (!cleanName) return hist;
  const last = hist.length > 0 ? hist[hist.length - 1] : null;
  if (last && last.name === cleanName) return hist;
  hist.push({ name: cleanName, activatedAt: new Date().toISOString() });
  return hist;
}

function __extractProfileFields_(src) {
  return {
    formUrl: src.formUrl || '',
    formTitle: src.formTitle || '',
    spreadsheetId: src.spreadsheetId || '',
    sheetName: src.sheetName || '',
    columnMapping: src.columnMapping || {},
    displaySettings: src.displaySettings || {},
    xAxisLabels: src.xAxisLabels || null,
    yAxisLabels: src.yAxisLabels || null,
    matrixQuadrantLabels: src.matrixQuadrantLabels || null,
    allowResubmit: !!src.allowResubmit
  };
}

// Why: displaySettings が profile に無いときは DEFAULT で埋める。空 {} を渡すと
//   matrix→wordcloud 切替時に旧 columnMapping.numericX/Y が残って空ボードになる。
function __applyProfileToConfig_(currentConfig, profile) {
  const defaultDisplay = typeof DEFAULT_DISPLAY_SETTINGS !== 'undefined' ? DEFAULT_DISPLAY_SETTINGS : {};
  return {
    ...currentConfig,
    formUrl: profile.formUrl || '',
    formTitle: profile.formTitle || '',
    spreadsheetId: profile.spreadsheetId || '',
    sheetName: profile.sheetName || '',
    columnMapping: profile.columnMapping || {},
    displaySettings: profile.displaySettings || defaultDisplay,
    xAxisLabels: profile.xAxisLabels || null,
    yAxisLabels: profile.yAxisLabels || null,
    matrixQuadrantLabels: profile.matrixQuadrantLabels || null,
    allowResubmit: !!profile.allowResubmit
  };
}

function __listProfilesCore(userId) {
  const cfg = getUserConfig(userId);
  if (!cfg.success) return createErrorResponse(cfg.message || 'getUserConfig failed');
  const cur = cfg.config || {};
  const profiles = Array.isArray(cur.profiles) ? cur.profiles : [];
  return createSuccessResponse('Profiles listed', {
    userId,
    activeProfile: cur.activeProfile || null,
    count: profiles.length,
    profiles: profiles.map(p => ({
      name: p.name,
      formTitle: p.formTitle || '',
      formUrl: p.formUrl || '',
      spreadsheetId: p.spreadsheetId || '',
      sheetName: p.sheetName || '',
      boardMode: (p.displaySettings && p.displaySettings.boardMode) || 'auto'
    }))
  });
}

/**
 * @param {string} userId
 * @param {string} name
 * @param {Object} [opts]
 * @param {Object} [opts.snapshot] - 明示指定の保存元 config。省略時は現在の active config。
 * @param {boolean} [opts.autoActivate=false] - 新規追加かつ active 未設定なら activate するか。
 */
function __saveProfileCore(userId, name, opts = {}) {
  if (!name || typeof name !== 'string' || !name.trim()) {
    return createErrorResponse('name (profile name) is required');
  }
  const cleanName = String(name).trim().substring(0, PROFILE_NAME_MAX_LEN);
  const cfg = getUserConfig(userId);
  if (!cfg.success) return createErrorResponse(cfg.message || 'getUserConfig failed');
  const cur = cfg.config || {};
  const src = (opts.snapshot && typeof opts.snapshot === 'object') ? opts.snapshot : cur;

  const newProfile = { name: cleanName, ...__extractProfileFields_(src) };

  const existing = Array.isArray(cur.profiles) ? cur.profiles.slice() : [];
  const idx = existing.findIndex(p => p && p.name === cleanName);
  const isUpdate = idx >= 0;
  if (isUpdate) existing[idx] = newProfile;
  else existing.push(newProfile);

  const patch = { profiles: existing };
  if (opts.autoActivate && !isUpdate && !cur.activeProfile) {
    patch.activeProfile = cleanName;
    // Why: 初回 active 化なので history が空でも 1 件目として記録する。
    //   生徒がこの profile を後で振り返れるよう、明示的に履歴 anchor を作る。
    patch.profileHistory = __appendProfileHistory_(cur.profileHistory, cleanName);
  }

  const saveRes = applyConfigPatch_(userId, patch, { publish: false });
  if (!saveRes.success) return saveRes;

  return createSuccessResponse(
    isUpdate ? `プロファイル「${cleanName}」を更新しました` : `プロファイル「${cleanName}」を保存しました`,
    { profileName: cleanName, count: existing.length, isUpdate, etag: saveRes.etag || null }
  );
}

function __deleteProfileCore(userId, name) {
  if (!name || typeof name !== 'string') {
    return createErrorResponse('name (profile name) is required');
  }
  const cfg = getUserConfig(userId);
  if (!cfg.success) return createErrorResponse(cfg.message || 'getUserConfig failed');
  const cur = cfg.config || {};
  const profiles = Array.isArray(cur.profiles) ? cur.profiles : [];
  const remaining = profiles.filter(p => p && p.name !== name);
  if (remaining.length === profiles.length) {
    return createErrorResponse(`プロファイル「${name}」は存在しません`);
  }
  const patch = { profiles: remaining };
  const wasActive = cur.activeProfile === name;

  // Why: active profile を削除すると top-level (formUrl/spreadsheetId/...) が
  //   削除済 profile の値を指したまま残り、polling / getPublishedSheetData が
  //   消えた spreadsheet を見にいって 403 を出し続ける。
  //   ログ調査 (2026-05-14) で過去 spreadsheet `1tFgE7g…` / `1ikZ63cn…` への
  //   403 が頻発していた根本原因はこれ。残った profile の先頭を新 active として
  //   再 anchor、それも無ければ全クリア + activeProfile=null。
  //   isPublished は触らない（公開ライフサイクルは 4 関数のみが扱う規約）が、
  //   spreadsheetId が空になるので validatePublishConfig は次回 publish 試行で
  //   reject される。
  if (wasActive) {
    if (remaining.length > 0) {
      const next = remaining[0];
      Object.assign(patch, __applyProfileToConfig_({}, next), { activeProfile: next.name });
    } else {
      patch.activeProfile = null;
      patch.formUrl = '';
      patch.formTitle = '';
      patch.spreadsheetId = '';
      patch.sheetName = '';
      patch.columnMapping = {};
    }
  }

  const saveRes = applyConfigPatch_(userId, patch, { publish: false });
  if (!saveRes.success) return saveRes;

  return createSuccessResponse(`プロファイル「${name}」を削除しました`, {
    count: remaining.length,
    remainingActive: patch.activeProfile !== undefined ? patch.activeProfile : cur.activeProfile,
    reAnchored: wasActive && remaining.length > 0
  });
}

function getLogs(options = {}) {
  try {
    const auth = requireAdmin();
    if (!auth) return createAdminRequiredError();

    const limit = Math.max(1, Math.min(Number(options.limit || 50), 200));
    const props = PropertiesService.getScriptProperties().getProperties();
    const logKeys = Object.keys(props)
      .filter(key => key.startsWith('security_log_'))
      .sort((a, b) => {
        const tsA = Number(a.split('_')[2] || 0);
        const tsB = Number(b.split('_')[2] || 0);
        return tsB - tsA;
      })
      .slice(0, limit);

    const logs = logKeys.map(key => {
      try {
        const parsed = JSON.parse(props[key]);
        return {
          key,
          ...parsed
        };
      } catch (parseError) {
        return {
          key,
          timestamp: null,
          severity: 'unknown',
          description: 'ログ解析に失敗しました',
          parseError: parseError.message
        };
      }
    });

    return {
      success: true,
      logs,
      count: logs.length,
      message: logs.length > 0 ? 'セキュリティログを取得しました' : 'セキュリティログはありません'
    };
  } catch (error) {
    logError_('getLogs', error);
    return createExceptionResponse(error);
  }
}


/**
 * アプリ全体のアクセス制限状態をチェック
 * @returns {boolean} true: アプリ停止中, false: 正常運用中
 */
function checkAppAccessRestriction() {
  try {
    const props = PropertiesService.getScriptProperties();
    const appDisabled = props.getProperty('APP_DISABLED');
    return appDisabled === 'true';
  } catch (error) {
    logError_('checkAppAccessRestriction', error);
    return false;
  }
}

/**
 * アプリ全体のアクセスを停止する（管理者専用）
 * @param {string} reason - 停止理由
 * @returns {Object} 実行結果
 */
function disableAppAccess(reason = 'システムメンテナンス') {
  try {
    const auth = requireAdmin();
    if (!auth) return createAdminRequiredError();
    const currentEmail = auth.email;

    PropertiesService.getScriptProperties().setProperties({
      APP_DISABLED: 'true',
      APP_DISABLED_REASON: reason,
      APP_DISABLED_BY: currentEmail,
      APP_DISABLED_AT: new Date().toISOString()
    });

    return createSuccessResponse('アプリケーションを停止しました', {
      reason,
      disabledBy: currentEmail,
      timestamp: new Date().toISOString()
    });
  } catch (error) {
    logError_('disableAppAccess', error);
    return createExceptionResponse(error, 'アプリケーション停止処理');
  }
}

/**
 * アプリ全体のアクセス制限を解除する（管理者専用）
 * @returns {Object} 実行結果
 */
function enableAppAccess() {
  try {
    const auth = requireAdmin();
    if (!auth) return createAdminRequiredError();

    const props = PropertiesService.getScriptProperties();

    // Read all in one call, then mutate.
    const all = props.getProperties();
    const disabledReason = all.APP_DISABLED_REASON || '';
    const disabledBy = all.APP_DISABLED_BY || '';
    const disabledAt = all.APP_DISABLED_AT || '';

    // PropertiesService has no deleteProperties() batch method, so individual
    // calls are unavoidable, but skip the ones whose values are already absent.
    ['APP_DISABLED', 'APP_DISABLED_REASON', 'APP_DISABLED_BY', 'APP_DISABLED_AT'].forEach((key) => {
      if (all[key] !== undefined) props.deleteProperty(key);
    });

    props.setProperty('APP_ENABLED_AT', new Date().toISOString());

    return createSuccessResponse('アプリケーションを再開しました', {
      previousRestriction: {
        reason: disabledReason,
        disabledBy,
        disabledAt
      },
      timestamp: new Date().toISOString()
    });
  } catch (error) {
    logError_('enableAppAccess', error);
    return createExceptionResponse(error, 'アプリケーション再開処理');
  }
}

/**
 * アプリアクセス制限の状態情報を取得（管理者専用）
 * @returns {Object} アクセス制限状態の詳細情報
 */
function getAppAccessStatus() {
  try {
    const auth = requireAdmin();
    if (!auth) return createAdminRequiredError();

    // Why: 6回の個別 getProperty() でなく 1 回の getProperties() で全取得。
    // GAS の PropertiesService 呼出は 1 回あたり 50-100ms のため 500ms級の削減。
    const all = PropertiesService.getScriptProperties().getProperties();
    const isDisabled = all.APP_DISABLED === 'true';

    const status = {
      isDisabled,
      status: isDisabled ? 'disabled' : 'enabled',
      adminEmail: all.ADMIN_EMAIL || '',
      timestamp: new Date().toISOString()
    };

    if (isDisabled) {
      status.restriction = {
        reason: all.APP_DISABLED_REASON || '',
        disabledBy: all.APP_DISABLED_BY || '',
        disabledAt: all.APP_DISABLED_AT || ''
      };
    } else {
      status.lastEnabled = {
        enabledAt: all.APP_ENABLED_AT || ''
      };
    }

    return createSuccessResponse('アクセス制限状態を取得しました', status);
  } catch (error) {
    logError_('getAppAccessStatus', error);
    return createExceptionResponse(error, 'アクセス制限状態取得処理');
  }
}

/**
 * ウェルカムモーダルを既読にする
 * @returns {Object} 実行結果
 */
function markWelcomeSeen() {
  try {
    const email = getCurrentEmail();
    if (!email) {
      return createAuthError();
    }

    const currentUser = findUserByEmail(email, { requestingUser: email });
    if (!currentUser) {
      return createUserNotFoundError('ユーザーが見つかりません');
    }

    // Why getUserConfig (not getConfigOrDefault): markWelcomeSeen は save パス。空 {} で
    //   全上書きすると profile/columnMapping を吹き飛ばす。read 失敗時は中断する。
    const cfgResult = getUserConfig(currentUser.userId);
    if (!cfgResult || !cfgResult.success) {
      return createErrorResponse(`設定の取得に失敗しました: ${(cfgResult && cfgResult.message) || '詳細不明'}`);
    }
    const config = cfgResult.config;
    config.hasSeenWelcome = true;

    const saveResult = saveUserConfig(currentUser.userId, config);
    if (!saveResult.success) {
      return createErrorResponse('設定の保存に失敗しました');
    }

    return createSuccessResponse('ウェルカムモーダルを既読にしました', {
      userId: currentUser.userId,
      hasSeenWelcome: true
    });
  } catch (error) {
    logError_('markWelcomeSeen', error);
    return createExceptionResponse(error);
  }
}

/**
 * Admin API ディスパッチャー
 * doPost の adminApi アクションおよび google.script.run から呼び出される
 * @param {string} operation - 操作名
 * @param {Object} params - パラメータ
 * @returns {Object} 操作結果
 */
// frontend (google.script.run) から call できる「ユーザースコープ」ops の allowlist。
// これらの handler は内部で userId と caller の一致を verify するので admin でなくても呼べる。
// 上記以外のすべての op は admin 認証を要求する (API キー経路 = getCurrentEmail が
// ADMIN_EMAIL を返す経由でパス、フロントエンドからの直叩きは reject)。
const __FRONTEND_USER_DISPATCH_OPS = Object.freeze(new Set([
  'lesson.list', 'lesson.create', 'lesson.delete', 'lesson.duplicate',
  'lesson.templates', 'lesson.review', 'lesson.updateDraft',
  'lesson.start', 'lesson.advance', 'lesson.end', 'lesson.knownClasses',
  'uploadLessonImage'
]));

function dispatchAdminOperation(operation, params) {
  if (!operation || typeof operation !== 'string') {
    return createErrorResponse('operation is required', null, { error: 'MISSING_OPERATION' });
  }

  const op = operation.trim();

  // Why entry-level admin gate: dispatchAdminOperation は google.script.run 経由でも
  //   呼ばれる (AdminPanel lesson UI の gasRun ヘルパー)。frontend からの直叩きで
  //   setProperty / appendRows / previewBoard 等の admin-only op が呼ばれると
  //   privilege escalation になる。API キー経路では _apiKeyAdminEmail が立っていて
  //   getCurrentEmail() = ADMIN_EMAIL となり isAdministrator が true を返すので
  //   この gate を通過できる。
  if (!__FRONTEND_USER_DISPATCH_OPS.has(op)) {
    if (!isAdministrator(getCurrentEmail())) {
      return createAdminRequiredError();
    }
  }

  // case ブランチ冒頭の param 検証を一行に圧縮するためのヘルパー。
  //   `const e = reqStr('userId'); if (e) return e;` の形で使う。
  const reqStr = (key, label) => {
    if (!params[key] || typeof params[key] !== 'string') {
      return createErrorResponse((label || key) + ' is required');
    }
    return null;
  };
  const reqObj = (key, label) => {
    if (!isPlainObject(params[key])) {
      return createErrorResponse((label || (key + ' (object)')) + ' is required');
    }
    return null;
  };

  switch (op) {
    // --- User Management ---
    case 'getUsers':
      return getAdminUsers();

    case 'toggleUserActive':
      { const e = reqStr('targetUserId'); if (e) return e; }
      return toggleUserActiveStatus(params.targetUserId);

    case 'toggleUserBoard':
      { const e = reqStr('targetUserId'); if (e) return e; }
      return toggleUserBoardStatus(params.targetUserId, { sourceEtag: params.etag });

    // --- Board Publish Lifecycle (unified ops) ---
    // Why: CLI/Admin から個別操作を直接呼べるよう dispatcher 経由でも提供。
    //   publishApp は config 差し替えを伴うため main.js の専用 case で別ルート。
    case 'unpublishBoard':
      return unpublishBoard(
        typeof params.targetUserId === 'string' ? params.targetUserId : null,
        { sourceEtag: params.etag }
      );

    case 'republishMyBoard':
      return republishMyBoard({ sourceEtag: params.etag });

    // --- App Control ---
    case 'getLogs':
      return getLogs({ limit: Number(params.limit) || 50 });

    case 'disableApp':
      return disableAppAccess(typeof params.reason === 'string' ? params.reason : 'システムメンテナンス');

    case 'enableApp':
      return enableAppAccess();

    case 'getAppStatus':
      return getAppAccessStatus();

    // --- System Diagnostics ---
    case 'systemDiagnosis':
      return testSystemDiagnosis();

    case 'autoRepair':
      return performAutoRepair();

    case 'cacheReset':
      return forceUrlSystemReset();

    case 'perfMetrics': {
      const category = typeof params.category === 'string' ? params.category : 'all';
      return getPerformanceMetrics(category, params.options || {});
    }

    case 'perfDiagnosis':
      return diagnosePerformance(params.options || {});

    // --- Property Management ---
    case 'getProperty': {
      { const e = reqStr('key'); if (e) return e; }
      if (isProtectedPropertyKey(params.key)) {
        return createErrorResponse(`Property "${params.key}" is protected and cannot be read via admin API`, null, { error: 'PROTECTED_PROPERTY' });
      }
      // Why direct read (lint suppressed): admin "view raw property" 用途。getCachedProperty
      //   は 30s cache が挟まるため、admin が値を更新した直後に古い値が見える事故になる。
      // lint-disable-next-line no-direct-property-fetch
      const value = PropertiesService.getScriptProperties().getProperty(params.key);
      return createSuccessResponse('Property retrieved', { key: params.key, value });
    }

    case 'setProperty': {
      { const e = reqStr('key'); if (e) return e; }
      if (typeof params.value !== 'string') {
        return createErrorResponse('value must be a string');
      }
      if (isProtectedPropertyKey(params.key)) {
        return createErrorResponse(`Property "${params.key}" is protected and cannot be modified via admin API. Use the GAS editor to rotate secrets.`, null, { error: 'PROTECTED_PROPERTY' });
      }
      PropertiesService.getScriptProperties().setProperty(params.key, params.value);
      return createSuccessResponse('Property set', { key: params.key });
    }

    case 'listProperties': {
      const allProps = PropertiesService.getScriptProperties().getProperties();
      const masked = {};
      for (const [k, v] of Object.entries(allProps)) {
        // 文字数も出さない (secret の長さは brute-force/形式推測のヒントになる)。
        masked[k] = isProtectedPropertyKey(k)
          ? '***redacted***'
          : v;
      }
      return createSuccessResponse('Properties listed', { properties: masked });
    }

    // --- User Config Operations ---
    // Why: 既存の getUsers は users 行を返すだけで、config JSON が文字列のまま。
    //      CLI から「ユーザー A の boardMode を numberline に」など個別操作するには
    //      parsed config を返す/受け取る API が必要。
    case 'findUser': {
      { const e = reqStr('email'); if (e) return e; }
      const user = findUserByEmail(String(params.email).trim(), { requestingUser: getCurrentEmail() });
      if (!user) return createErrorResponse('User not found');
      return createSuccessResponse('User found', { user: redactUserForApi_(user) });
    }

    case 'getUserConfig': {
      { const e = reqStr('userId'); if (e) return e; }
      const result = getUserConfig(params.userId);
      if (!result.success) return createErrorResponse(result.message || 'getUserConfig failed');
      return createSuccessResponse('Config retrieved', {
        userId: params.userId,
        config: result.config
      });
    }

    case 'exportConfigs': {
      // Why: 全ユーザーの config を JSON でダンプ。バックアップや別環境への移行、
      //      差分管理、教師ごとの設定比較に使う。configJson 文字列をパース済みで返す。
      const usersResult = getAdminUsers();
      if (!usersResult.success) return usersResult;
      const exported = (usersResult.users || []).map(u => {
        let parsed = {};
        try {
          parsed = u.configJson ? JSON.parse(u.configJson) : {};
        } catch (parseErr) {
          // Why: 設定破損ユーザーを admin が後追いできるよう、どの user の config parse が失敗したかを
          //   ログに残す。空 {} で続行（exportConfigs 全体は止めない）。
          console.warn('exportConfigs: configJson parse failed', {
            userId: u.userId ? `${u.userId.substring(0, 8)}***` : 'N/A',
            error: parseErr.message
          });
          parsed = {};
        }
        return {
          userId: u.userId,
          userEmail: u.userEmail,
          isActive: u.isActive,
          createdAt: u.createdAt,
          lastModified: u.lastModified,
          config: parsed
        };
      });
      return createSuccessResponse('Configs exported', {
        count: exported.length,
        exportedAt: new Date().toISOString(),
        configs: exported
      });
    }

    case 'setUserConfig': {
      { const e = reqStr('userId'); if (e) return e; }
      { const e = reqObj('patch'); if (e) return e; }
      return applyConfigPatch_(params.userId, params.patch, {
        publish: Boolean(params.publish)
      });
    }

    case 'bulkSetUserConfig': {
      // 全ユーザーに同じ patch を適用。dryRun で影響範囲を見られる。
      // Why: 「6 年全クラスに同じ軸ラベル一括展開」など、運用上の bulk 操作用。
      { const e = reqObj('patch'); if (e) return e; }
      const dryRun = Boolean(params.dryRun);
      const targetFilter = params.filter || {};
      const usersResult = getAdminUsers();
      if (!usersResult.success) return usersResult;
      const users = usersResult.users || [];
      const matched = users.filter(u => matchesUserFilter_(u, targetFilter));
      if (dryRun) {
        return createSuccessResponse('Dry run', {
          matchedCount: matched.length,
          totalUsers: users.length,
          sample: matched.slice(0, 5).map(u => ({ userId: u.userId, userEmail: u.userEmail }))
        });
      }
      const results = [];
      for (const u of matched) {
        const r = applyConfigPatch_(u.userId, params.patch, { publish: false });
        results.push({ userId: u.userId, userEmail: u.userEmail, success: r.success, message: r.message });
      }
      const okCount = results.filter(r => r.success).length;
      return createSuccessResponse(`bulkSetUserConfig: ${okCount}/${results.length} succeeded`, {
        results,
        appliedTo: okCount,
        failedCount: results.length - okCount
      });
    }

    case 'runColumnAnalysis': {
      // Why: フロントのみで実行できた列分析を CLI からも実行可能に。
      //      列マッピング自動検出 + numericScaleCandidates を CLI でレビューできる。
      if (!params.spreadsheetId || typeof params.spreadsheetId !== 'string') {
        // userId 指定なら、そのユーザーの config から取得
        if (params.userId && typeof params.userId === 'string') {
          const cfgRes = getUserConfig(params.userId);
          if (!cfgRes.success || !cfgRes.config?.spreadsheetId) {
            return createErrorResponse('User has no spreadsheet configured');
          }
          return getColumnAnalysis(cfgRes.config.spreadsheetId, cfgRes.config.sheetName);
        }
        return createErrorResponse('spreadsheetId or userId is required');
      }
      const sheetName = typeof params.sheetName === 'string' ? params.sheetName : 'フォームの回答 1';
      return getColumnAnalysis(params.spreadsheetId, sheetName);
    }

    // --- Form Operations ---
    // Why: 教師が CLI から「テンプレート Form 作成 → ユーザー設定に紐付け」「既存 Form
    //      URL を渡して config を一括セットアップ」「Drive 内の Form 一覧確認」を完結
    //      できるようにする。frontend ボタン経由と同じヘルパー (createTemplateForm /
    //      processFormUrlInput / getForms / isValidFormUrl) を再利用するので挙動は等価。
    //
    //      重要な制約: FormApp / DriveApp は GAS の executeAs: USER_ACCESSING 設定により
    //      「webapp にアクセスしたユーザー」のコンテキストで動く。adminApi 経由でも
    //      gcloud OAuth トークンの所有者が requester として認識されるので、結果として
    //      その所有者の Drive に Form / Spreadsheet が作成される。
    //      → 中先生が自分のアカウントで CLI を回す典型ケースでは正常に機能する。
    case 'listMyForms': {
      // 自分の Drive 内の Google Forms を最大 30 件返す（getForms の thin wrapper）
      return getForms();
    }

    case 'validateFormUrl': {
      { const e = reqStr('formUrl'); if (e) return e; }
      const trimmed = String(params.formUrl).trim();
      // Why: isValidFormUrl は new URL() ベースで実装されており、GAS V8 で URL constructor が
      //      期待通り動かないコンテキストがある（catch で false を返してしまう known issue）。
      //      CLI 経路では canonical truth として「FormApp.openByUrl で開けるか？」を直接試す。
      //      URL 形式は substring check で補助確認するのみ。
      const looksLikeFormUrl =
        trimmed.startsWith('https://') &&
        (trimmed.includes('docs.google.com/forms/') || trimmed.includes('forms.gle/'));

      if (!looksLikeFormUrl) {
        return createSuccessResponse('URL format check', {
          formUrl: trimmed,
          isValidFormUrl: false,
          reachable: false,
          reason: 'Not a Google Forms URL pattern'
        });
      }
      let reachable = false;
      let formTitle = null;
      let destinationSpreadsheetId = null;
      try {
        const form = FormApp.openByUrl(trimmed);
        reachable = true;
        formTitle = form.getTitle();
        // Why: form.getDestinationId() は destination 未設定時に "フォームに応答先がありません。"
        //   を throw する仕様なので、null と同じ扱いに変換する (誤って reachable=false 判定するのを防ぐ)。
        try {
          destinationSpreadsheetId = form.getDestinationId() || null;
        } catch (_) {
          destinationSpreadsheetId = null;
        }
      } catch (e) {
        return createSuccessResponse('Form URL reachability check', {
          formUrl: trimmed,
          isValidFormUrl: true,
          reachable: false,
          reason: e.message
        });
      }
      return createSuccessResponse('Form URL validated', {
        formUrl: trimmed,
        isValidFormUrl: true,
        reachable,
        formTitle,
        destinationSpreadsheetId
      });
    }

    case 'connectForm': {
      // Form URL を解決してユーザー config に紐付け。frontend の processFormUrlInput
      // と同じパイプライン（spreadsheet 自動作成・回答シート検出も含む）。
      // userId 省略時は requester (gcloud user) の userId に自動解決。
      { const e = reqStr('formUrl'); if (e) return e; }
      const targetUserId = (typeof params.userId === 'string' && params.userId) ? params.userId : null;
      let user = null;
      if (targetUserId) {
        user = findUserById(targetUserId, { requestingUser: getCurrentEmail() });
      } else {
        const email = getCurrentEmail();
        user = email ? findUserByEmail(email, { requestingUser: email }) : null;
      }
      if (!user) return createErrorResponse('User not found');

      const procResult = processFormUrlInput(params.formUrl);
      if (!procResult.success) return procResult;

      // config に formUrl / formTitle / spreadsheetId / sheetName を反映
      const patch = {
        formUrl: procResult.formUrl || procResult.publishedUrl || params.formUrl,
        formTitle: procResult.formTitle || '',
        spreadsheetId: procResult.spreadsheetId || '',
        sheetName: procResult.sheetName || 'フォームの回答 1'
      };
      const saveResult = applyConfigPatch_(user.userId, patch, { publish: false });
      // config 保存が失敗したら全体を失敗として返す。 旧実装は top-level success:true で
      //   失敗を nested configSave に埋もれさせており、 .success を見る CLI が「接続済」と
      //   誤認していた。
      if (!saveResult.success) {
        return createErrorResponse(`Form の config 保存に失敗しました: ${saveResult.message || '詳細不明'}`, {
          userId: user.userId,
          applied: patch,
          formProcessing: procResult
        });
      }
      return createSuccessResponse('Form connected', {
        userId: user.userId,
        userEmail: user.userEmail,
        applied: patch,
        formProcessing: procResult,
        configSave: { success: saveResult.success, message: saveResult.message }
      });
    }

    case 'customizeForm': {
      // 既存 Form の質問項目をスキーマで完全置換。
      // CLI から授業用にカスタマイズしたいときに使う。回答 0 件の新規 Form 対象。
      if (!params.formId && !params.formUrl) {
        return createErrorResponse('formId or formUrl is required');
      }
      { const e = reqObj('schema'); if (e) return e; }
      return customizeForm(params.formId || params.formUrl, params.schema);
    }

    case 'setFormAllowResubmit': {
      if (!params.formId && !params.formUrl) {
        return createErrorResponse('formId or formUrl is required');
      }
      return setFormAllowResubmit(params.formId || params.formUrl, Boolean(params.allowResubmit));
    }

    case 'uploadLessonImage': {
      { const e = reqStr('imageData'); if (e) return e; }
      return uploadLessonImage(params.imageData, params.filename || '');
    }

    case 'createForm': {
      // templateType: 'board' | 'numberline' | 'matrix' | 'pie' (default: 'board')
      const templateType = ['board', 'numberline', 'matrix', 'pie'].includes(params.templateType)
        ? params.templateType : 'board';
      const targetUserId = (typeof params.userId === 'string' && params.userId) ? params.userId : null;
      let user = null;
      if (targetUserId) {
        user = findUserById(targetUserId, { requestingUser: getCurrentEmail() });
      } else {
        const email = getCurrentEmail();
        user = email ? findUserByEmail(email, { requestingUser: email }) : null;
      }
      if (!user) return createErrorResponse('User not found');

      const createResult = createTemplateForm(templateType, params.templateOptions || {});
      if (!createResult.success) return createResult;

      // config に紐付け。numberline/matrix の場合は boardMode も auto / 明示モードに揃える。
      const patch = {
        formUrl: createResult.formUrl || '',
        formTitle: createResult.formTitle || '',
        spreadsheetId: createResult.spreadsheetId || '',
        sheetName: createResult.sheetName || 'フォームの回答 1'
      };
      if (templateType === 'numberline' || templateType === 'matrix' || templateType === 'pie') {
        patch.displaySettings = { boardMode: templateType };
      }
      const saveResult = applyConfigPatch_(user.userId, patch, { publish: false });
      // Form は作成済だが config 保存が失敗した場合は全体を失敗として返す (created は同梱して
      //   作成済 Form URL を呼び出し側に残す)。 top-level success に保存結果を反映する。
      if (!saveResult.success) {
        return createErrorResponse(`Form は作成しましたが config 保存に失敗しました: ${saveResult.message || '詳細不明'}`, {
          userId: user.userId,
          templateType,
          created: createResult,
          applied: patch
        });
      }
      return createSuccessResponse('Form created and linked', {
        userId: user.userId,
        userEmail: user.userEmail,
        templateType,
        created: createResult,
        applied: patch,
        configSave: { success: saveResult.success, message: saveResult.message }
      });
    }

    // --- Multi-board Profile Operations ---
    // Why: 1 ユーザーが複数 Forms を切替えて使えるよう、設定スナップショットの
    //      CRUD を CLI から実行可能にする。教師が「導入アンケート」「本時の議論」
    //      「振り返り」と複数 board を事前に登録して、授業中に切替えるユースケース。
    case 'listProfiles': {
      { const e = reqStr('userId'); if (e) return e; }
      return __listProfilesCore(params.userId);
    }

    case 'saveProfile': {
      { const e = reqStr('userId'); if (e) return e; }
      return __saveProfileCore(params.userId, params.name, { snapshot: params.snapshot });
    }

    case 'loadProfile': {
      // 指定 profile を active 設定に適用して、activeProfile も更新。
      { const e = reqStr('userId'); if (e) return e; }
      { const e = reqStr('name', 'name (profile name)'); if (e) return e; }
      const cfg = getUserConfig(params.userId);
      if (!cfg.success) return createErrorResponse(cfg.message || 'getUserConfig failed');
      const cur = cfg.config || {};
      const profiles = Array.isArray(cur.profiles) ? cur.profiles : [];
      const p = profiles.find(x => x && x.name === params.name);
      if (!p) return createErrorResponse(`Profile "${params.name}" not found`);

      // Why (完全置換セマンティクス): applyConfigPatch_ + deepMerge_ は plain object
      //   を再帰マージするため、新 profile の columnMapping={} が「前 profile の
      //   columnMapping を保持」と解釈されてしまう。loadProfile は active 切替なので
      //   完全置換が正しい。saveUserConfig を直接呼んで merged を渡す。
      const merged = {
        ...__applyProfileToConfig_(cur, p),
        activeProfile: p.name,
        userId: params.userId,
        profileHistory: __appendProfileHistory_(cur.profileHistory, p.name)
      };

      const saveRes = saveUserConfig(params.userId, merged, { isMainConfig: true });
      // profile 切替で board data が変わるので viewer cache を即時 stale 化
      // (cache key に activeProfile が含まれないため bump しないと旧データが残る)。
      if (saveRes.success && typeof bumpBoardDataVersion_ === 'function') {
        try { bumpBoardDataVersion_(params.userId); } catch (_) { /* ignore */ }
      }
      return saveRes;
    }

    case 'deleteProfile': {
      { const e = reqStr('userId'); if (e) return e; }
      return __deleteProfileCore(params.userId, params.name);
    }

    case 'repairSpreadsheetSharing': {
      // SA pool 全員を editor として追加する遡及修復 (v2782 以前作成の SS や、 後から SA を
      // pool 追加した場合に必要)。
      //   npm run api -- repairSpreadsheetSharing --spreadsheetId <ID>
      { const e = reqStr('spreadsheetId'); if (e) return e; }
      try {
        const result = applySpreadsheetSharingDefaults(params.spreadsheetId);
        return createSuccessResponse('Spreadsheet sharing repaired', {
          spreadsheetId: params.spreadsheetId,
          saAdded: result.saAdded,
          saEmails: result.saEmails || [],
          errors: result.errors
        });
      } catch (e) {
        return createExceptionResponse(e);
      }
    }

    // ===== SA pool 管理 (v2782+) =====
    case 'listServiceAccountPool': {
      try {
        return (typeof listServiceAccountPool === 'function')
          ? listServiceAccountPool()
          : createErrorResponse('listServiceAccountPool not available');
      } catch (e) {
        return createExceptionResponse(e);
      }
    }

    case 'getServiceAccountUsage': {
      try {
        return (typeof getServiceAccountUsage === 'function')
          ? getServiceAccountUsage()
          : createErrorResponse('getServiceAccountUsage not available');
      } catch (e) {
        return createExceptionResponse(e);
      }
    }

    case 'addServiceAccountToPool': {
      // 1 個の SA JSON を pool に登録。 戻り値の verified を見て UI に反映する。
      //   npm run api -- addServiceAccountToPool --json '<SA JSON>'
      { const e = reqStr('json'); if (e) return e; }
      try {
        if (typeof addServiceAccountToPool !== 'function') {
          return createErrorResponse('addServiceAccountToPool not available');
        }
        return addServiceAccountToPool(params.json);
      } catch (e) {
        return createExceptionResponse(e);
      }
    }

    case 'addServiceAccountsToPoolBatch': {
      // 複数 SA JSON を一括登録 (改行区切り or `}\n{` 連結 or array)。
      const inputs = params.inputs || params.json;
      if (!inputs) return createErrorResponse('inputs (string or array) is required');
      try {
        if (typeof addServiceAccountsToPoolBatch !== 'function') {
          return createErrorResponse('addServiceAccountsToPoolBatch not available');
        }
        return addServiceAccountsToPoolBatch(inputs);
      } catch (e) {
        return createExceptionResponse(e);
      }
    }

    case 'reverifyServiceAccountInPool': {
      // 共有反映待ちで verified=false だった SA を再確認。
      const slot = params.slot || params.slotKey;
      if (!slot) return createErrorResponse('slot (e.g. "SERVICE_ACCOUNT_CREDS_2") is required');
      try {
        if (typeof reverifyServiceAccountInPool !== 'function') {
          return createErrorResponse('reverifyServiceAccountInPool not available');
        }
        return reverifyServiceAccountInPool(slot);
      } catch (e) {
        return createExceptionResponse(e);
      }
    }

    case 'removeServiceAccountFromPool': {
      // Secondary slot の SA を削除 (primary は別経路で差替)。
      const slot = params.slot || params.slotKey;
      if (!slot) return createErrorResponse('slot (e.g. "SERVICE_ACCOUNT_CREDS_2") is required');
      try {
        if (typeof removeServiceAccountFromPool !== 'function') {
          return createErrorResponse('removeServiceAccountFromPool not available');
        }
        return removeServiceAccountFromPool(slot);
      } catch (e) {
        return createExceptionResponse(e);
      }
    }

    case 'migrateBoardSharing': {
      // v2782 以前作成のボード SS を一括クリーンアップする。
      //   1) SA pool 全員を editor として追加 (board が SA pool 経由で動くようにする)
      //   2) DOMAIN_WITH_LINK 共有を revoke (Drive bleed を解消)
      // dryRun=true で実際の変更なしに対象 SS のリストだけ確認可能。
      //   npm run api -- migrateBoardSharing --dryRun true
      //   npm run api -- migrateBoardSharing
      const dryRun = params.dryRun === true || params.dryRun === 'true';
      try {
        if (typeof migrateBoardSharing !== 'function') {
          return createErrorResponse('migrateBoardSharing not available');
        }
        return migrateBoardSharing({ dryRun });
      } catch (e) {
        return createExceptionResponse(e);
      }
    }

    case 'setSheetHeader': {
      // header 行 (row 1) の任意セルを書き換え。
      //   Why: createTemplateForm のデフォルト設問タイトル (「回答」「理由」等) は汎用的すぎて
      //        ボード上の headingLabel ("eab-question") が意味不明になる。授業ごとに
      //        「あなたが手品師なら、どちらを選ぶ？」のような具体的な問いに差し替えたい。
      //   Form item 側 (setTitle) は触らないので、既存 100+ 件の回答に影響しない。
      //   columnIndex は 0-based (admin api / columnMapping と一致)。
      { const e = reqStr('spreadsheetId'); if (e) return e; }
      { const e = reqStr('sheetName'); if (e) return e; }
      if (typeof params.columnIndex !== 'number' || !Number.isInteger(params.columnIndex) || params.columnIndex < 0) {
        return createErrorResponse('columnIndex (non-negative integer, 0-based) is required');
      }
      if (typeof params.newHeader !== 'string' || !params.newHeader.trim()) {
        return createErrorResponse('newHeader (non-empty string) is required');
      }
      { const denied = assertManagedSpreadsheet_(params.spreadsheetId); if (denied) return denied; }
      try {
        const ss = SpreadsheetApp.openById(params.spreadsheetId);
        const sheet = ss.getSheetByName(params.sheetName);
        if (!sheet) return createErrorResponse(`Sheet "${params.sheetName}" not found`);
        const colNum = params.columnIndex + 1;
        const lastCol = sheet.getLastColumn();
        if (colNum > lastCol) {
          return createErrorResponse(`columnIndex ${params.columnIndex} is out of range (lastColumn=${lastCol - 1})`);
        }
        const cell = sheet.getRange(1, colNum);
        const oldHeader = String(cell.getValue() || '');
        const cleanHeader = String(params.newHeader).trim().substring(0, 200);
        cell.setValue(cleanHeader);
        // Why: SHEET_HEADERS は 10 分 TTL でキャッシュされる (SystemController.DATABASE_LONG)。
        //   invalidate しないと getQuestionText が古い「回答」を返し続け、headingLabel が
        //   更新されない (CLAUDE.md: setupColumns 時は明示的に invalidate せよ規約と同じ)。
        try {
          if (typeof invalidateSheetHeadersCache === 'function') {
            invalidateSheetHeadersCache(params.spreadsheetId, params.sheetName);
          }
        } catch (cacheErr) {
          console.warn('setSheetHeader: cache invalidate failed', cacheErr.message);
        }
        return createSuccessResponse('Header updated', {
          spreadsheetId: params.spreadsheetId,
          sheetName: params.sheetName,
          columnIndex: params.columnIndex,
          oldHeader,
          newHeader: cleanHeader
        });
      } catch (e) {
        return createExceptionResponse(e);
      }
    }

    case 'clearDataRows': {
      // データ行（2行目以降）を全削除して、ヘッダーだけ残す。
      // appendRows と組み合わせて「テストデータをきれいに作り直す」用途。
      // 本番運用では絶対に使わないこと（明示の確認フラグ params.confirm = "yes-i-mean-it" 必須）。
      { const e = reqStr('spreadsheetId'); if (e) return e; }
      { const e = reqStr('sheetName'); if (e) return e; }
      if (params.confirm !== 'yes-i-mean-it') {
        return createErrorResponse('confirm: "yes-i-mean-it" is required (safeguard against accidental data loss)');
      }
      { const denied = assertManagedSpreadsheet_(params.spreadsheetId); if (denied) return denied; }
      try {
        const ss = SpreadsheetApp.openById(params.spreadsheetId);
        const sheet = ss.getSheetByName(params.sheetName);
        if (!sheet) return createErrorResponse(`Sheet "${params.sheetName}" not found`);
        const lastRow = sheet.getLastRow();
        if (lastRow <= 1) {
          return createSuccessResponse('No data rows to clear', { deletedCount: 0 });
        }
        // Why clearContent (not deleteRows): Sheets は「sheet に行が 1 つも残らない」
        //   deleteRows を拒否し「固定されていない行をすべて削除することはできません」
        //   エラーになる (ヘッダー 1 行 + データ 100 行を deleteRows(2, 100) しても起こる)。
        //   clearContent で内容だけ消せば getLastRow() は 1 に戻り、appendRows が row 2
        //   から書ける。Reaction 列等の lazy provisioning ヘッダーも保持される。
        const lastCol = Math.max(1, sheet.getLastColumn());
        sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
        return createSuccessResponse('Data rows cleared', {
          spreadsheetId: params.spreadsheetId,
          sheetName: params.sheetName,
          deletedCount: lastRow - 1
        });
      } catch (e) {
        return createExceptionResponse(e);
      }
    }

    case 'renameDriveFile': {
      // Drive 上のファイル名を変更する（Form / Spreadsheet どちらも file.setName で OK）。
      // 主用途: customizeForm でタイトルを差し替えた既存 Form / SS の名前同期、
      //        または手動命名済の SS を後から命名し直す等。
      //
      //   注意: form.setTitle() は Forms 内部タイトル、SpreadsheetApp.create(name) や
      //         DriveApp.setName() は Drive 上の file name を別々に管理する。AdminPanel
      //         「一覧から選択」dropdown は file name を表示するので、両者を揃える必要がある。
      { const e = reqStr('fileId'); if (e) return e; }
      if (!params.newName || typeof params.newName !== 'string' || !params.newName.trim()) {
        return createErrorResponse('newName (non-empty string) is required');
      }
      try {
        const file = DriveApp.getFileById(params.fileId);
        const oldName = file.getName();
        const cleanName = String(params.newName).trim().substring(0, SHEET_NAME_MAX_LEN);
        file.setName(cleanName);
        return createSuccessResponse('Drive file renamed', {
          fileId: params.fileId,
          oldName,
          newName: cleanName,
          mimeType: file.getMimeType()
        });
      } catch (e) {
        return createExceptionResponse(e);
      }
    }

    case 'appendRows': {
      // Why: テストデータ投入用。Spreadsheet に複数 row を一括 append。
      //      Forms 経由ではなく直接書き込むので、本番運用では使わないこと。
      //      授業前の動作確認・デモ・空状態回避でのみ使う。
      //
      //   セーフガード:
      //   - 最大 500 row まで（巨大投入を防ぐ）
      //   - first col が null/'' なら自動で now() タイムスタンプ
      //   - 列数は既存ヘッダ幅に padding して整列
      { const e = reqStr('spreadsheetId'); if (e) return e; }
      { const e = reqStr('sheetName'); if (e) return e; }
      if (!Array.isArray(params.rows) || params.rows.length === 0) {
        return createErrorResponse('rows (non-empty array) is required');
      }
      if (params.rows.length > MAX_BULK_INSERT_ROWS) {
        return createErrorResponse(`rows array too large (max ${MAX_BULK_INSERT_ROWS})`);
      }
      { const denied = assertManagedSpreadsheet_(params.spreadsheetId); if (denied) return denied; }
      try {
        const ss = SpreadsheetApp.openById(params.spreadsheetId);
        const sheet = ss.getSheetByName(params.sheetName);
        if (!sheet) return createErrorResponse(`Sheet "${params.sheetName}" not found`);

        const now = new Date();
        const lastCol = Math.max(1, sheet.getLastColumn());
        const normalized = params.rows.map((row, idx) => {
          if (!Array.isArray(row)) return null;
          const r = row.slice();
          // first column を timestamp として扱う。
          //   - null / '' / undefined : 1 秒ずらしの auto-timestamp で並び順を安定化
          //   - string (ISO 8601 等) : Date オブジェクトに変換して Sheets に日付セルとして書く。
          //     Why: Sheets はテキスト ISO 文字列を「日付セル」として認識しないため、
          //          aggregator が new Date(r.timestamp) で NaN を出し buildHistoricalSnapshots が
          //          空配列を返す。Date 化することでフォーム自動入力と同じ型にそろう。
          //   - Date / number 等は素通し
          if (r[0] === null || r[0] === undefined || r[0] === '') {
            r[0] = new Date(now.getTime() + idx * 1000);
          } else if (typeof r[0] === 'string') {
            const parsed = new Date(r[0]);
            if (Number.isFinite(parsed.getTime())) r[0] = parsed;
          }
          return r;
        }).filter(Boolean);

        if (normalized.length === 0) {
          return createErrorResponse('no valid rows after normalization');
        }
        const maxLen = Math.max(lastCol, ...normalized.map(r => r.length));
        const padded = normalized.map(r => {
          const out = r.slice();
          while (out.length < maxLen) out.push('');
          return out;
        });

        const startRow = sheet.getLastRow() + 1;
        sheet.getRange(startRow, 1, padded.length, maxLen).setValues(padded);

        return createSuccessResponse('Rows appended', {
          spreadsheetId: params.spreadsheetId,
          sheetName: params.sheetName,
          appendedCount: padded.length,
          startRow,
          endRow: startRow + padded.length - 1
        });
      } catch (e) {
        return createExceptionResponse(e);
      }
    }

    case 'previewBoard': {
      // Why: 教師がボードを公開する前に「viewer から何が見えるか」を CLI で
      //      確認できる（プライバシー設定通りに匿名化されているか等）。
      { const e = reqStr('userId'); if (e) return e; }
      // getPublishedSheetData は admin/owner 判定のため getBatchedAdminAuth を内部で使う。
      // 本関数は API キー経路（既に admin と同等）なので、targetUserId を渡せば直接呼べる。
      const result = getPublishedSheetData(
        params.classFilter || null,
        params.sortOrder || 'newest',
        false,
        params.userId
      );
      // 大きい data 配列は CLI で扱いにくいので、件数だけ返してサンプル 3 件を見せる
      const trimmed = {
        success: result.success,
        header: result.header,
        sheetName: result.sheetName,
        displaySettings: result.displaySettings,
        axisConfig: result.axisConfig,
        // Why: CLI から profile / formMeta の状態確認をしたいケースが多い。
        //      ペイロードに含めても軽量なので preview にも乗せる。
        formMeta: result.formMeta || null,
        profiles: result.profiles || null,
        dataCount: Array.isArray(result.data) ? result.data.length : 0,
        sampleData: Array.isArray(result.data) ? result.data.slice(0, 3) : [],
        error: result.error
      };
      return createSuccessResponse('Board preview', trimmed);
    }

    case 'exportBoardData': {
      // 実践報告書 / 授業記録の作成用に、ボード公開済みの全件データを CLI から取得する。
      //   - profileName を渡せば「過去 phase」も読める（getPublishedSheetDataForProfile 経由）
      //   - 個人情報保護のため、デフォルトで `name` フィールドを除外する（stripName=false で残す）
      //   - reactions / emailHash / highlight / id / rowIndex は report 用には不要なので削ぎ落とす
      { const e = reqStr('userId'); if (e) return e; }
      const profileName = typeof params.profileName === 'string' && params.profileName ? params.profileName : null;
      const stripName = params.stripName !== false; // default true
      const result = profileName
        ? getPublishedSheetDataForProfile(params.userId, profileName, params.classFilter || null, params.sortOrder || 'newest')
        : getPublishedSheetData(params.classFilter || null, params.sortOrder || 'newest', false, params.userId);
      const rows = Array.isArray(result.data) ? result.data : [];
      const slim = rows.map(r => __projectBoardRowForExport_(r, { includeName: !stripName }));
      return createSuccessResponse('Board data export', {
        success: result.success,
        header: result.header || '',
        sheetName: result.sheetName || '',
        boardMode: result?.displaySettings?.boardMode || null,
        axisConfig: result.axisConfig || null,
        profile: profileName,
        rowCount: slim.length,
        stripName,
        data: slim,
        error: result.error
      });
    }

    // --- Lessons (lesson archive + workspace) ---
    case 'lesson.create': {
      { const e = reqStr('userId'); if (e) return e; }
      return createLessonDraft(params.userId, params.name, params.template);
    }
    case 'lesson.updateDraft': {
      { const e = reqStr('userId'); if (e) return e; }
      { const e = reqStr('lessonId'); if (e) return e; }
      { const e = reqStr('fieldPath'); if (e) return e; }
      return updateLessonDraft(params.userId, params.lessonId, params.fieldPath, params.value, params.expectedEtag);
    }
    case 'lesson.start': {
      { const e = reqStr('userId'); if (e) return e; }
      { const e = reqStr('lessonId'); if (e) return e; }
      return startLesson(params.userId, params.lessonId);
    }
    case 'lesson.advance': {
      { const e = reqStr('userId'); if (e) return e; }
      { const e = reqStr('lessonId'); if (e) return e; }
      { const e = reqStr('direction'); if (e) return e; }
      return advanceLessonPhase(params.userId, params.lessonId, params.direction);
    }
    case 'lesson.end': {
      { const e = reqStr('userId'); if (e) return e; }
      { const e = reqStr('lessonId'); if (e) return e; }
      return endLesson(params.userId, params.lessonId);
    }
    case 'lesson.list': {
      { const e = reqStr('userId'); if (e) return e; }
      return listLessons(params.userId);
    }
    case 'lesson.knownClasses': {
      { const e = reqStr('userId'); if (e) return e; }
      return getKnownClassesForUser(params.userId);
    }
    case 'lesson.review': {
      { const e = reqStr('userId'); if (e) return e; }
      { const e = reqStr('lessonId'); if (e) return e; }
      return getLessonForReview(params.userId, params.lessonId);
    }
    case 'lesson.delete': {
      { const e = reqStr('userId'); if (e) return e; }
      { const e = reqStr('lessonId'); if (e) return e; }
      return deleteLesson(params.userId, params.lessonId);
    }
    case 'lesson.duplicate': {
      { const e = reqStr('userId'); if (e) return e; }
      { const e = reqStr('lessonId'); if (e) return e; }
      return duplicateLesson(params.userId, params.lessonId, params.options || {});
    }
    case 'lesson.templates': {
      return listLessonTemplates();
    }
    case 'lesson.importFromProfiles': {
      // 既存 profiles[] (= plain profile 構成) を「過去授業の lesson 記録」として
      //   lessons シートに取り込む。lesson 機能 (Phase 1+2) 導入以前に運用していた授業を
      //   履歴に残すための一方向 import。詳細は LessonService.importLessonFromProfiles の jsdoc。
      { const e = reqStr('userId'); if (e) return e; }
      return importLessonFromProfiles(params.userId, {
        name: params.name,
        includeSnapshots: params.includeSnapshots !== false && params.includeSnapshots !== 'false'
      });
    }

    default:
      return createErrorResponse(`Unknown operation: ${op}`, null, { error: 'UNKNOWN_OPERATION' });
  }
}

/**
 * users 行から API レスポンス向けの安全な表現に整形（configJson 文字列をパース）。
 * Why: getUsers は raw 行を返すが、CLI で扱いやすくパース済みオブジェクトに整形。
 */
function redactUserForApi_(user) {
  if (!user || typeof user !== 'object') return user;
  let config = null;
  try {
    config = user.configJson ? JSON.parse(user.configJson) : null;
  } catch (parseErr) {
    // Why: redactUserForApi_ は admin API レスポンスの一部。config parse 失敗は珍しい
    //   が起きると config-derived fields が null になるので、トラブルシュートのために残す。
    console.warn('redactUserForApi_: configJson parse failed', {
      userId: user.userId ? `${String(user.userId).substring(0, 8)}***` : 'N/A',
      error: parseErr.message
    });
  }
  return {
    userId: user.userId,
    userEmail: user.userEmail,
    googleId: user.googleId || null,
    isActive: !!user.isActive,
    createdAt: user.createdAt,
    lastModified: user.lastModified,
    config
  };
}

/**
 * ユーザーフィルタ判定。`bulkSetUserConfig` 用。
 * Why: target を全件か一部に絞るシンプルな機構。複雑な条件は将来追加。
 *   filter = { isActive: true, isPublished: false, emailContains: "naha" }
 */
function matchesUserFilter_(user, filter) {
  if (!filter || typeof filter !== 'object') return true;
  if (filter.isActive !== undefined && Boolean(user.isActive) !== Boolean(filter.isActive)) return false;
  if (filter.emailContains && typeof filter.emailContains === 'string') {
    if (!String(user.userEmail || '').toLowerCase().includes(filter.emailContains.toLowerCase())) return false;
  }
  if (filter.isPublished !== undefined) {
    const cfg = safeJsonParse_(user.configJson, null);
    if (!cfg) return !filter.isPublished;  // parse 不能 = unpublished 扱い
    if (Boolean(cfg.isPublished) !== Boolean(filter.isPublished)) return false;
  }
  return true;
}

/**
 * 既存 config に patch を deep-merge して保存する。
 *
 * Why: CLI から `setUserConfig --patch '{"displaySettings":{"boardMode":"numberline"}}'`
 *      で部分更新できるようにするための内部実装。深いマージで他フィールドを保持。
 *
 *      sanitize は saveUserConfig が validateAndSanitizeConfig 経由で行うので、
 *      ここでは patch を merge するだけ。許可フィールドはバックエンドの sanitize
 *      関数の allowlist が最終ゲート。
 *
 *      protected fields (userId/userEmail/etag) は明示的に削除。
 */
function applyConfigPatch_(userId, patch, options) {
  if (!userId || typeof userId !== 'string') {
    return createErrorResponse('userId is required');
  }
  const PROTECTED = ['userId', 'userEmail', 'googleId', 'createdAt', 'etag'];
  // Clone & strip protected keys top-level
  const safePatch = {};
  for (const k of Object.keys(patch)) {
    if (PROTECTED.includes(k)) continue;
    safePatch[k] = patch[k];
  }

  // Load existing config
  const cur = getUserConfig(userId);
  if (!cur.success) return createErrorResponse(cur.message || 'getUserConfig failed');
  // 破損 config (parse 失敗で default fallback) への patch は復旧可能な原本を default で
  //   上書きしてしまうため中止する。
  if (cur.corrupted) {
    return createErrorResponse('設定データが破損しているため上書きを中止しました。設定の復旧が必要です。');
  }
  const merged = deepMerge_(cur.config || {}, safePatch);
  merged.userId = userId;

  // Save through existing sanitizing pipeline (validateAndSanitizeConfig handles everything)
  const saveOptions = options && options.publish ? { isPublish: true } : { isMainConfig: true };
  const result = saveUserConfig(userId, merged, saveOptions);
  return result;
}

/**
 * Deep merge: target に source を mutate せずマージ。
 * - plain object 同士は再帰
 * - 配列は置換（merge しない）
 * - source の null は target のキーを削除する明示シグナル
 */
function deepMerge_(target, source) {
  if (target === null || typeof target !== 'object' || Array.isArray(target)) {
    return source === undefined ? target : source;
  }
  const out = { ...target };
  for (const k of Object.keys(source)) {
    const v = source[k];
    if (v === null) {
      delete out[k];
    } else if (Array.isArray(v)) {
      out[k] = v.slice();
    } else if (typeof v === 'object') {
      out[k] = deepMerge_(target[k] || {}, v);
    } else {
      out[k] = v;
    }
  }
  return out;
}

// =========================================================================
// migrateBoardSharing - v2782 以前作成ボードの共有設定一括クリーンアップ
// =========================================================================

/**
 * 全 users.config の spreadsheetId を走査し、 以下を一括適用する:
 *   1. SA pool 全員を editor として追加 (board が SA pool 経由で動くようにする)
 *   2. DOMAIN_WITH_LINK / DOMAIN 共有を revoke (Drive bleed と直接編集を解消)
 *
 * 冪等。 何度呼んでも安全。 各 SS の処理は fail-soft (1 つの SS で失敗しても他を続行)。
 *
 * @param {Object} [options]
 * @param {boolean} [options.dryRun=false] - true なら共有変更なし、 対象 SS リストのみ返す
 * @returns {Object} 結果サマリ
 */
function migrateBoardSharing(options = {}) {
  const dryRun = options.dryRun === true;

  const ssIds = getAllBoardSpreadsheetIds();
  if (ssIds.length === 0) {
    return createSuccessResponse('No board spreadsheets found', { total: 0, processed: [], dryRun });
  }

  const processed = [];
  let saAddedCount = 0;
  let domainRevokedCount = 0;
  let errorCount = 0;

  for (const ssId of ssIds) {
    const entry = { spreadsheetId: ssId, saEmails: [], domainRevoked: false, errors: [] };

    if (dryRun) {
      processed.push(entry);
      continue;
    }

    // 1) SA pool 全員を editor 追加
    try {
      const shareResult = applySpreadsheetSharingDefaults(ssId);
      entry.saEmails = (shareResult && shareResult.saEmails) || [];
      if (shareResult && shareResult.errors && shareResult.errors.length) {
        entry.errors.push(...shareResult.errors);
      }
      if (entry.saEmails.length > 0) saAddedCount++;
    } catch (shareError) {
      entry.errors.push('SA share: ' + (shareError.message || shareError));
    }

    // 2) DOMAIN_WITH_LINK / DOMAIN 共有を revoke
    try {
      const revoked = revokeDomainSharing_(ssId);
      if (revoked.success) {
        entry.domainRevoked = true;
        if (revoked.changed) domainRevokedCount++;
      } else if (revoked.error) {
        entry.errors.push('Revoke: ' + revoked.error);
      }
    } catch (revokeError) {
      entry.errors.push('Revoke: ' + (revokeError.message || revokeError));
    }

    if (entry.errors.length > 0) errorCount++;
    processed.push(entry);
  }

  return createSuccessResponse(
    dryRun ? `Dry run: ${ssIds.length} board SS would be processed` : `${ssIds.length} board SS processed`,
    { total: ssIds.length, saAddedCount, domainRevokedCount, errorCount, dryRun, processed }
  );
}

/**
 * SS の DOMAIN_WITH_LINK / DOMAIN 共有を revoke する。
 *   - 既に PRIVATE なら no-op (changed=false)
 *   - 通常 setSharing で PRIVATE に戻す。 失敗時は Drive REST API で domain permission を削除
 *
 * @returns {{success:boolean, changed:boolean, error?:string}}
 */
function revokeDomainSharing_(spreadsheetId) {
  try {
    const file = DriveApp.getFileById(spreadsheetId);
    const access = file.getSharingAccess();

    if (access === DriveApp.Access.PRIVATE) {
      return { success: true, changed: false };
    }

    try {
      file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
      return { success: true, changed: true };
    } catch (_setSharingError) {
      // Workspace ポリシーで setSharing が落ちる場合は Drive REST API で domain permission を
      // 直接削除する fallback。
      const token = ScriptApp.getOAuthToken();
      const listUrl = 'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(spreadsheetId) +
        '/permissions?fields=permissions(id,type,role,domain)&supportsAllDrives=true';
      const listResp = UrlFetchApp.fetch(listUrl, {
        method: 'get',
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true
      });
      if (listResp.getResponseCode() !== 200) {
        return { success: false, changed: false,
          error: `Drive list permissions ${listResp.getResponseCode()}: ${listResp.getContentText().substring(0, DRIVE_ERROR_BODY_PREVIEW_LEN)}` };
      }
      const data = JSON.parse(listResp.getContentText());
      const perms = Array.isArray(data.permissions) ? data.permissions : [];
      const domainPerms = perms.filter((p) => p.type === 'domain');
      let changed = false;
      for (const p of domainPerms) {
        const delUrl = 'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(spreadsheetId) +
          '/permissions/' + encodeURIComponent(p.id) + '?supportsAllDrives=true';
        const delResp = UrlFetchApp.fetch(delUrl, {
          method: 'delete',
          headers: { Authorization: 'Bearer ' + token },
          muteHttpExceptions: true
        });
        const code = delResp.getResponseCode();
        if (code >= 200 && code < 300) { changed = true; }
        else if (code !== 404) {
          return { success: false, changed,
            error: `Drive delete permission ${code}: ${delResp.getContentText().substring(0, DRIVE_ERROR_BODY_PREVIEW_LEN)}` };
        }
      }
      return { success: true, changed };
    }
  } catch (err) {
    return { success: false, changed: false, error: err.message || String(err) };
  }
}

