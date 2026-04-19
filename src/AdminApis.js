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

/* global getCurrentEmail, isAdministrator, findUserById, findUserByEmail, getAllUsers, updateUser, getUserConfig, saveUserConfig, createAdminRequiredError, createAuthError, createUserNotFoundError, createErrorResponse, createSuccessResponse, createExceptionResponse, requireAdmin, getConfigOrDefault, PropertiesService */


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

    const updatedUser = {
      ...targetUser,
      isActive: !targetUser.isActive,
      lastModified: new Date().toISOString()
    };

    const result = updateUser(targetUserId, updatedUser);
    if (result.success) {
      return {
        success: true,
        message: `ユーザー状態を${updatedUser.isActive ? 'アクティブ' : '非アクティブ'}に変更しました`,
        userId: targetUserId,
        newStatus: updatedUser.isActive,
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

    const props = PropertiesService.getScriptProperties();
    props.setProperty('APP_DISABLED', 'true');
    props.setProperty('APP_DISABLED_REASON', reason);
    props.setProperty('APP_DISABLED_BY', currentEmail);
    props.setProperty('APP_DISABLED_AT', new Date().toISOString());

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

    default:
      return createErrorResponse(`Unknown operation: ${op}`, null, { error: 'UNKNOWN_OPERATION' });
  }
}
