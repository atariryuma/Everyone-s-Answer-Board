/**
 * @fileoverview AdminApis - 管理者専用API
 *
 * ✅ V8ランタイム対応（2022年6月アップデート準拠）
 * - 関数定義は順序に関係なく呼び出し可能
 * - グローバルスコープでのコード実行を完全排除
 *
 * 依存関係（呼び出す関数）:
 * - getCurrentEmail() - main.jsで定義
 * - isAdministrator() - main.jsで定義
 * - findUserById() - DatabaseCore.jsで定義
 * - findUserByEmail() - DatabaseCore.jsで定義
 * - getAllUsers() - DatabaseCore.jsで定義
 * - updateUser() - DatabaseCore.jsで定義
 * - getUserConfig() - ConfigService.jsで定義
 * - saveUserConfig() - ConfigService.jsで定義
 * - createAdminRequiredError() - helpers.jsで定義
 * - createAuthError() - helpers.jsで定義
 * - createUserNotFoundError() - helpers.jsで定義
 * - createErrorResponse() - helpers.jsで定義
 * - createSuccessResponse() - helpers.jsで定義
 * - createExceptionResponse() - helpers.jsで定義
 *
 * 移動元: main.js
 * 移動日: 2025-12-13
 */

/* global getCurrentEmail, isAdministrator, findUserById, findUserByEmail, getAllUsers, updateUser, getUserConfig, saveUserConfig, getColumnAnalysis, getPublishedSheetData, createTemplateForm, customizeForm, processFormUrlInput, getForms, isValidFormUrl, applySpreadsheetSharingDefaults, FormApp, SpreadsheetApp, ScriptApp, UrlFetchApp, createAdminRequiredError, createAuthError, createUserNotFoundError, createErrorResponse, createSuccessResponse, createExceptionResponse, requireAdmin, getConfigOrDefault */


// Admin API経由での読み書きから保護する Script Properties キー。
// APIキー漏洩時のブラスト半径を最小化する目的で、
// 認証情報系（substring match）と、攻撃者に差し替えられるとサービス乗っ取り/
// 偽DBへの誘導が可能になる基幹キーを明示的にブロックする。
const PROTECTED_PROPERTY_SUBSTRINGS = ['KEY', 'CREDS', 'SECRET', 'TOKEN', 'PASSWORD'];
const PROTECTED_PROPERTY_EXACT_KEYS = ['ADMIN_EMAIL', 'DATABASE_SPREADSHEET_ID'];

function isProtectedPropertyKey(key) {
  if (typeof key !== 'string' || !key) return false;
  const upper = key.toUpperCase();
  if (PROTECTED_PROPERTY_EXACT_KEYS.includes(upper)) return true;
  return PROTECTED_PROPERTY_SUBSTRINGS.some((s) => upper.includes(s));
}


/**
 * Get users - simplified name for admin panel
 * @param {Object} _options - Options (reserved for future use)
 * @returns {Object} User list result
 */
function getAdminUsers(_options = {}) {
  try {
    const auth = requireAdmin();
    if (!auth) return createAdminRequiredError();

    const users = getAllUsers({ activeOnly: false }, { forceServiceAccount: true });
    return {
      success: true,
      users: users || []
    };
  } catch (error) {
    console.error('getAdminUsers error:', error.message);
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
    console.error('toggleUserActiveStatus error:', error.message);
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
  const isOwnBoard = targetUser.userEmail === email;

  if (options.requireAdmin && !isAdmin) {
    return createErrorResponse('この操作には管理者権限が必要です');
  }
  if (!isAdmin && !isOwnBoard) {
    return createErrorResponse('ボードの公開状態を変更する権限がありません');
  }

  // 第2引数で preloaded user を渡し、getUserConfig 内の findUserById 再実行を避ける。
  const currentConfig = getConfigOrDefault(targetUser.userId, targetUser);

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

  const saveResult = saveUserConfig(targetUser.userId, updatedConfig);
  if (!saveResult.success) {
    return createErrorResponse(`ボード状態の更新に失敗しました: ${saveResult.message || '詳細不明'}`);
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
    console.error('toggleUserBoardStatus error:', error.message);
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
    console.error('republishMyBoard error:', error.message);
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
    return __applyPublishStateChange(targetUserId || null, false, { sourceEtag: options.sourceEtag });
  } catch (error) {
    console.error('unpublishBoard error:', error.message);
    return createExceptionResponse(error);
  }
}

/**
 * @deprecated Use unpublishBoard() instead. Kept for back-compat with existing
 *   client code (page.js endPublication / AdminPanel unpublishBoardBtn).
 *   このエイリアスは v2700 以降で削除予定。新規コードは unpublishBoard を使うこと。
 * @param {string} [targetUserId]
 * @returns {Object} 実行結果
 */
function clearActiveSheet(targetUserId) {
  return unpublishBoard(targetUserId);
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
  const cleanName = String(name).trim().substring(0, 50);
  const cfg = getUserConfig(userId);
  if (!cfg.success) return createErrorResponse(cfg.message || 'getUserConfig failed');
  const cur = cfg.config || {};
  const src = (opts.snapshot && typeof opts.snapshot === 'object') ? opts.snapshot : cur;

  const newProfile = {
    name: cleanName,
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
      patch.activeProfile = next.name;
      patch.formUrl = next.formUrl || '';
      patch.formTitle = next.formTitle || '';
      patch.spreadsheetId = next.spreadsheetId || '';
      patch.sheetName = next.sheetName || '';
      patch.columnMapping = next.columnMapping || {};
      patch.displaySettings = next.displaySettings || (typeof DEFAULT_DISPLAY_SETTINGS !== 'undefined' ? DEFAULT_DISPLAY_SETTINGS : {});
      patch.xAxisLabels = next.xAxisLabels || null;
      patch.yAxisLabels = next.yAxisLabels || null;
      patch.matrixQuadrantLabels = next.matrixQuadrantLabels || null;
      patch.allowResubmit = !!next.allowResubmit;
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

/**
 * Get logs - simplified name
 * @param {Object} _options - Options (reserved for future use)
 * @returns {Object} Logs result
 */
function getLogs(_options = {}) {
  try {
    const auth = requireAdmin();
    if (!auth) return createAdminRequiredError();

    const limit = Math.max(1, Math.min(Number(_options.limit || 50), 200));
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
    console.error('getLogs error:', error.message);
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
    console.error('checkAppAccessRestriction error:', error.message);
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
    console.error('disableAppAccess error:', error.message);
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
    console.error('enableAppAccess error:', error.message);
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
    console.error('getAppAccessStatus error:', error.message);
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
      return createAuthError('ユーザー認証が必要です');
    }

    const currentUser = findUserByEmail(email, { requestingUser: email });
    if (!currentUser) {
      return createUserNotFoundError('ユーザーが見つかりません');
    }

    const config = getConfigOrDefault(currentUser.userId);

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
    console.error('markWelcomeSeen error:', error.message);
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
function dispatchAdminOperation(operation, params) {
  if (!operation || typeof operation !== 'string') {
    return createErrorResponse('operation is required', null, { error: 'MISSING_OPERATION' });
  }

  const op = operation.trim();

  switch (op) {
    // --- User Management ---
    case 'getUsers':
      return getAdminUsers();

    case 'toggleUserActive':
      if (!params.targetUserId || typeof params.targetUserId !== 'string') {
        return createErrorResponse('targetUserId is required');
      }
      return toggleUserActiveStatus(params.targetUserId);

    case 'toggleUserBoard':
      if (!params.targetUserId || typeof params.targetUserId !== 'string') {
        return createErrorResponse('targetUserId is required');
      }
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
      if (!params.key || typeof params.key !== 'string') {
        return createErrorResponse('key is required');
      }
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
      if (!params.key || typeof params.key !== 'string') {
        return createErrorResponse('key is required');
      }
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
        masked[k] = isProtectedPropertyKey(k)
          ? `***${v.length}chars***`
          : v;
      }
      return createSuccessResponse('Properties listed', { properties: masked });
    }

    // --- User Config Operations ---
    // Why: 既存の getUsers は users 行を返すだけで、config JSON が文字列のまま。
    //      CLI から「ユーザー A の boardMode を numberline に」など個別操作するには
    //      parsed config を返す/受け取る API が必要。
    case 'findUser': {
      if (!params.email || typeof params.email !== 'string') {
        return createErrorResponse('email is required');
      }
      const user = findUserByEmail(String(params.email).trim(), { requestingUser: getCurrentEmail() });
      if (!user) return createErrorResponse('User not found');
      return createSuccessResponse('User found', { user: redactUserForApi_(user) });
    }

    case 'getUserConfig': {
      if (!params.userId || typeof params.userId !== 'string') {
        return createErrorResponse('userId is required');
      }
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
      const exported = (usersResult.data?.users || usersResult.users || []).map(u => {
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
      if (!params.userId || typeof params.userId !== 'string') {
        return createErrorResponse('userId is required');
      }
      if (!params.patch || typeof params.patch !== 'object' || Array.isArray(params.patch)) {
        return createErrorResponse('patch (object) is required');
      }
      return applyConfigPatch_(params.userId, params.patch, {
        publish: Boolean(params.publish)
      });
    }

    case 'bulkSetUserConfig': {
      // 全ユーザーに同じ patch を適用。dryRun で影響範囲を見られる。
      // Why: 「6 年全クラスに同じ軸ラベル一括展開」など、運用上の bulk 操作用。
      if (!params.patch || typeof params.patch !== 'object' || Array.isArray(params.patch)) {
        return createErrorResponse('patch (object) is required');
      }
      const dryRun = Boolean(params.dryRun);
      const targetFilter = params.filter || {};
      const usersResult = getAdminUsers();
      if (!usersResult.success) return usersResult;
      const users = usersResult.data?.users || usersResult.users || [];
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
      if (!params.formUrl || typeof params.formUrl !== 'string') {
        return createErrorResponse('formUrl is required');
      }
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
        // form.getDestinationId() は destination 未設定時に "フォームに応答先がありません。"
        //   を throw する仕様なので、null と同じ扱いに変換する。
        //   2026-05-14 v2657: 教師が destination 無しの Form を validateFormUrl で
        //   調べたとき "reachable: false" と誤判定されるバグを修正。
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
      if (!params.formUrl || typeof params.formUrl !== 'string') {
        return createErrorResponse('formUrl is required');
      }
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
      if (!params.schema || typeof params.schema !== 'object') {
        return createErrorResponse('schema (object) is required');
      }
      return customizeForm(params.formId || params.formUrl, params.schema);
    }

    case 'createForm': {
      // テンプレート Form を作成して、指定ユーザー (または requester) の config に紐付け。
      // templateType: 'board' | 'numberline' | 'matrix' (default: 'board')
      const templateType = ['board', 'numberline', 'matrix'].includes(params.templateType)
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

      const createResult = createTemplateForm(templateType);
      if (!createResult.success) return createResult;

      // config に紐付け。numberline/matrix の場合は boardMode も auto / 明示モードに揃える。
      const patch = {
        formUrl: createResult.formUrl || '',
        formTitle: createResult.formTitle || '',
        spreadsheetId: createResult.spreadsheetId || '',
        sheetName: createResult.sheetName || 'フォームの回答 1'
      };
      if (templateType === 'numberline' || templateType === 'matrix') {
        patch.displaySettings = { boardMode: templateType };
      }
      const saveResult = applyConfigPatch_(user.userId, patch, { publish: false });

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
      if (!params.userId || typeof params.userId !== 'string') {
        return createErrorResponse('userId is required');
      }
      return __listProfilesCore(params.userId);
    }

    case 'saveProfile': {
      if (!params.userId || typeof params.userId !== 'string') {
        return createErrorResponse('userId is required');
      }
      return __saveProfileCore(params.userId, params.name, { snapshot: params.snapshot });
    }

    case 'loadProfile': {
      // 指定 profile を active 設定に適用して、activeProfile も更新。
      if (!params.userId || typeof params.userId !== 'string') {
        return createErrorResponse('userId is required');
      }
      if (!params.name || typeof params.name !== 'string') {
        return createErrorResponse('name (profile name) is required');
      }
      const cfg = getUserConfig(params.userId);
      if (!cfg.success) return createErrorResponse(cfg.message || 'getUserConfig failed');
      const cur = cfg.config || {};
      const profiles = Array.isArray(cur.profiles) ? cur.profiles : [];
      const p = profiles.find(x => x && x.name === params.name);
      if (!p) return createErrorResponse(`Profile "${params.name}" not found`);

      // profile の中身を active config の対応フィールドに反映
      //
      // Why (完全置換セマンティクス): applyConfigPatch_ + deepMerge_ は plain object
      //   を再帰マージするため、新 profile の columnMapping={} が「前 profile の
      //   columnMapping を保持」と解釈されてしまう (matrix → wordcloud で numericX/Y
      //   が居残る原因)。loadProfile は active 切替なので「完全置換」が正しい
      //   セマンティクス。saveUserConfig を直接呼んで merged を渡す。
      const merged = { ...cur };
      merged.formUrl = p.formUrl || '';
      merged.formTitle = p.formTitle || '';
      merged.spreadsheetId = p.spreadsheetId || '';
      merged.sheetName = p.sheetName || '';
      merged.activeProfile = p.name;
      merged.columnMapping = p.columnMapping || {};
      merged.displaySettings = p.displaySettings || (typeof DEFAULT_DISPLAY_SETTINGS !== 'undefined' ? DEFAULT_DISPLAY_SETTINGS : {});
      merged.xAxisLabels = p.xAxisLabels || null;
      merged.yAxisLabels = p.yAxisLabels || null;
      merged.matrixQuadrantLabels = p.matrixQuadrantLabels || null;
      merged.allowResubmit = !!p.allowResubmit;
      merged.userId = params.userId;
      // Option B: 生徒が後で振り返れるよう、active 切替を時系列で記録。
      merged.profileHistory = __appendProfileHistory_(cur.profileHistory, p.name);

      return saveUserConfig(params.userId, merged, { isMainConfig: true });
    }

    case 'deleteProfile': {
      if (!params.userId || typeof params.userId !== 'string') {
        return createErrorResponse('userId is required');
      }
      return __deleteProfileCore(params.userId, params.name);
    }

    case 'shareWithDomain': {
      // Drive REST API でファイル（Spreadsheet / Form）を指定ドメインに共有する。
      //
      // Why: CLI からテンプレート Form を作ると、ファイル所有者は CLI ユーザー
      //      （gas-deploy-bot サービスアカウント等）になり、別ドメインの生徒・教師は
      //      アクセスできない。これでドメイン全体（naha-okinawa.ed.jp 等）に
      //      reader として共有することで access denied を解消する。
      //
      // 注意: Drive Advanced Service は使わず UrlFetchApp + REST API でアクセス。
      //      auth/drive スコープが appsscript.json に必要（既に granted）。
      if (!params.fileId || typeof params.fileId !== 'string') {
        return createErrorResponse('fileId is required');
      }
      if (!params.domain || typeof params.domain !== 'string') {
        return createErrorResponse('domain is required (e.g. "naha-okinawa.ed.jp")');
      }
      const role = ['reader', 'writer', 'commenter'].includes(params.role) ? params.role : 'reader';
      try {
        const token = ScriptApp.getOAuthToken();
        const url = 'https://www.googleapis.com/drive/v3/files/' +
                    encodeURIComponent(params.fileId) + '/permissions?supportsAllDrives=true';
        const resp = UrlFetchApp.fetch(url, {
          method: 'post',
          contentType: 'application/json',
          muteHttpExceptions: true,
          headers: { Authorization: 'Bearer ' + token },
          payload: JSON.stringify({
            type: 'domain',
            role: role,
            domain: params.domain,
            allowFileDiscovery: false  // Drive 検索結果には出さない（リンクを持つ人のみ）
          })
        });
        const code = resp.getResponseCode();
        const body = resp.getContentText();
        if (code >= 400) {
          return createErrorResponse('Drive API error ' + code + ': ' + body.substring(0, 300));
        }
        return createSuccessResponse('Shared with domain', {
          fileId: params.fileId,
          domain: params.domain,
          role,
          permission: JSON.parse(body)
        });
      } catch (e) {
        return createExceptionResponse(e);
      }
    }

    case 'repairSpreadsheetSharing': {
      // Why: createForm / customizeForm で過去に作成された SS のうち、
      //   サービスアカウントが editor に入っていないものは DatabaseCore の Sheets API JWT
      //   経路で 403 を返し、view モードの getPublishedSheetData が落ちる。
      //   現在の createTemplateForm / customizeForm は applySpreadsheetSharingDefaults を
      //   作成直後に呼ぶよう修正済みだが、修正前に作られた SS の遡及修復が必要。
      //
      //   呼び出し例:
      //     npm run api -- repairSpreadsheetSharing --spreadsheetId <ID>
      //     npm run api -- repairSpreadsheetSharing --spreadsheetId <ID> --ownerEmail teacher@example
      //
      //   ownerEmail を省略すると getCurrentEmail() を使用。ドメイン共有はスキップしたい
      //   ローカルテスト用に、明示的に ownerEmail を渡せる。
      if (!params.spreadsheetId || typeof params.spreadsheetId !== 'string') {
        return createErrorResponse('spreadsheetId is required');
      }
      const ownerEmail = (typeof params.ownerEmail === 'string' && params.ownerEmail)
        ? params.ownerEmail
        : getCurrentEmail();
      try {
        const result = applySpreadsheetSharingDefaults(params.spreadsheetId, ownerEmail);
        return createSuccessResponse('Spreadsheet sharing repaired', {
          spreadsheetId: params.spreadsheetId,
          ownerEmail,
          saAdded: result.saAdded,
          domainShared: result.domainShared,
          errors: result.errors
        });
      } catch (e) {
        return createExceptionResponse(e);
      }
    }

    case 'clearDataRows': {
      // データ行（2行目以降）を全削除して、ヘッダーだけ残す。
      // appendRows と組み合わせて「テストデータをきれいに作り直す」用途。
      // 本番運用では絶対に使わないこと（明示の確認フラグ params.confirm = "yes-i-mean-it" 必須）。
      if (!params.spreadsheetId || typeof params.spreadsheetId !== 'string') {
        return createErrorResponse('spreadsheetId is required');
      }
      if (!params.sheetName || typeof params.sheetName !== 'string') {
        return createErrorResponse('sheetName is required');
      }
      if (params.confirm !== 'yes-i-mean-it') {
        return createErrorResponse('confirm: "yes-i-mean-it" is required (safeguard against accidental data loss)');
      }
      try {
        const ss = SpreadsheetApp.openById(params.spreadsheetId);
        const sheet = ss.getSheetByName(params.sheetName);
        if (!sheet) return createErrorResponse(`Sheet "${params.sheetName}" not found`);
        const lastRow = sheet.getLastRow();
        if (lastRow <= 1) {
          return createSuccessResponse('No data rows to clear', { deletedCount: 0 });
        }
        // header (row 1) は残し、row 2 〜 lastRow を削除
        sheet.deleteRows(2, lastRow - 1);
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
      if (!params.fileId || typeof params.fileId !== 'string') {
        return createErrorResponse('fileId is required');
      }
      if (!params.newName || typeof params.newName !== 'string' || !params.newName.trim()) {
        return createErrorResponse('newName (non-empty string) is required');
      }
      try {
        const file = DriveApp.getFileById(params.fileId);
        const oldName = file.getName();
        const cleanName = String(params.newName).trim().substring(0, 200);
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
      if (!params.spreadsheetId || typeof params.spreadsheetId !== 'string') {
        return createErrorResponse('spreadsheetId is required');
      }
      if (!params.sheetName || typeof params.sheetName !== 'string') {
        return createErrorResponse('sheetName is required');
      }
      if (!Array.isArray(params.rows) || params.rows.length === 0) {
        return createErrorResponse('rows (non-empty array) is required');
      }
      if (params.rows.length > 500) {
        return createErrorResponse('rows array too large (max 500)');
      }
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
      if (!params.userId || typeof params.userId !== 'string') {
        return createErrorResponse('userId is required');
      }
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
    try {
      const cfg = user.configJson ? JSON.parse(user.configJson) : {};
      if (Boolean(cfg.isPublished) !== Boolean(filter.isPublished)) return false;
    } catch (_) {
      if (filter.isPublished) return false;
    }
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

