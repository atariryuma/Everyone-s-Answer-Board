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

/* global getCurrentEmail, isAdministrator, findUserById, findUserByEmail, getAllUsers, updateUser, getUserConfig, saveUserConfig, getColumnAnalysis, getPublishedSheetData, createTemplateForm, processFormUrlInput, getForms, isValidFormUrl, FormApp, createAdminRequiredError, createAuthError, createUserNotFoundError, createErrorResponse, createSuccessResponse, createExceptionResponse, requireAdmin, getConfigOrDefault */


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

/**
 * Toggle user board publication status - admin only (for managing other users)
 * @param {string} targetUserId - Target user ID
 * @returns {Object} Toggle result
 */
function toggleUserBoardStatus(targetUserId) {
  try {
    const auth = requireAdmin();
    if (!auth) return createAdminRequiredError();
    const email = auth.email;

    const targetUser = findUserById(targetUserId, { requestingUser: email });
    if (!targetUser) {
      return createUserNotFoundError();
    }

    let currentConfig = {};
    try {
      const configJsonStr = targetUser.configJson || '{}';
      currentConfig = JSON.parse(configJsonStr);
    } catch (error) {
      console.error('toggleUserBoardStatus: Invalid configJson for user', targetUserId, ':', error.message);
      currentConfig = {};
    }

    const newIsPublished = !currentConfig.isPublished;
    const updates = {};

    const updatedConfig = { ...currentConfig };
    updatedConfig.isPublished = newIsPublished;
    if (newIsPublished && !updatedConfig.publishedAt) {
      updatedConfig.publishedAt = new Date().toISOString();
    }
    updatedConfig.lastAccessedAt = new Date().toISOString();

    updates.configJson = JSON.stringify(updatedConfig);
    updates.lastModified = new Date().toISOString();

    const updateResult = updateUser(targetUserId, updates);
    if (!updateResult.success) {
      return createErrorResponse(`Failed to toggle board status: ${updateResult.message || '詳細不明'}`);
    }

    return {
      success: true,
      message: `ボードを${newIsPublished ? '公開' : '非公開'}に変更しました`,
      userId: targetUserId,
      boardPublished: newIsPublished,
      timestamp: new Date().toISOString()
    };
  } catch (error) {
    console.error('toggleUserBoardStatus error:', error.message);
    return createExceptionResponse(error);
  }
}

/**
 * 自分のボードを再公開（所有者用のシンプル関数）
 * @returns {Object} Republish result
 */
function republishMyBoard() {
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

    config.isPublished = true;
    config.publishedAt = new Date().toISOString();

    const saveResult = saveUserConfig(currentUser.userId, config);
    if (!saveResult.success) {
      return {
        success: false,
        message: '設定の保存に失敗しました'
      };
    }

    // リダイレクトURL（GAS iFrame対応）
    const redirectUrl = getWebAppUrl() + '?mode=view&userId=' + currentUser.userId;

    return {
      success: true,
      message: '回答ボードを再公開しました',
      userId: currentUser.userId,
      redirectUrl: redirectUrl
    };

  } catch (error) {
    console.error('republishMyBoard error:', error.message);
    return createExceptionResponse(error);
  }
}

/**
 * Clear active sheet publication (set board to unpublished)
 * @param {string} targetUserId - 対象ユーザーID（省略時は現在のユーザー）
 * @returns {Object} 実行結果
 */
function clearActiveSheet(targetUserId) {
  try {
    const email = getCurrentEmail();
    if (!email) {
      return createAuthError();
    }

    let targetUser = targetUserId ? findUserById(targetUserId, { requestingUser: email }) : null;
    if (!targetUser) {
      targetUser = findUserByEmail(email, { requestingUser: email });
    }

    if (!targetUser) {
      return createUserNotFoundError();
    }

    const isAdmin = isAdministrator(email);
    const isOwnBoard = targetUser.userEmail === email;

    if (!isAdmin && !isOwnBoard) {
      return createErrorResponse('ボードの非公開権限がありません');
    }

    const config = getConfigOrDefault(targetUser.userId);

    const wasPublished = config.isPublished === true;
    config.isPublished = false;
    config.publishedAt = null;
    config.lastAccessedAt = new Date().toISOString();

    const saveResult = saveUserConfig(targetUser.userId, config, { forceUpdate: false });
    if (!saveResult.success) {
      return createErrorResponse(`ボード状態の更新に失敗しました: ${saveResult.message || '詳細不明'}`);
    }

    // リダイレクトURL（GAS iFrame対応 - viewモードでサーバー側が公開状態を判定）
    const redirectUrl = getWebAppUrl() + '?mode=view&userId=' + targetUser.userId;

    return {
      success: true,
      message: wasPublished ? 'ボードを非公開にしました' : 'ボードは既に非公開です',
      boardPublished: false,
      userId: targetUser.userId,
      redirectUrl: redirectUrl
    };
  } catch (error) {
    console.error('clearActiveSheet error:', error.message);
    return createExceptionResponse(error);
  }
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
      return toggleUserBoardStatus(params.targetUserId);

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
        try { parsed = u.configJson ? JSON.parse(u.configJson) : {}; } catch (_) { parsed = {}; }
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
        destinationSpreadsheetId = form.getDestinationId() || null;
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
  try { config = user.configJson ? JSON.parse(user.configJson) : null; } catch (_) {}
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
