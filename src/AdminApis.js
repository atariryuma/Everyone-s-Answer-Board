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

/* global getCurrentEmail, isAdministrator, findUserById, findUserByEmail, getAllUsers, updateUser, getUserConfig, saveUserConfig, createAdminRequiredError, createAuthError, createUserNotFoundError, createErrorResponse, createSuccessResponse, createExceptionResponse, PropertiesService */


/**
 * Get users - simplified name for admin panel
 * @param {Object} _options - Options (reserved for future use)
 * @returns {Object} User list result
 */
function getAdminUsers(_options = {}) {
  try {
    const email = getCurrentEmail();
    if (!email || !isAdministrator(email)) {
      return createAdminRequiredError();
    }

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
    const email = getCurrentEmail();
    if (!email || !isAdministrator(email)) {
      return createAdminRequiredError();
    }

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
    const email = getCurrentEmail();
    if (!email) {
      return createAuthError('ユーザー認証が必要です');
    }

    if (!isAdministrator(email)) {
      return createAuthError('管理者権限が必要です');
    }

    const targetUser = findUserById(targetUserId, { requestingUser: email });
    if (!targetUser) {
      return createUserNotFoundError();
    }

    let currentConfig = {};
    try {
      const configJsonStr = targetUser.configJson || '{}';
      currentConfig = JSON.parse(configJsonStr);
    } catch (error) {
      console.warn('toggleUserBoardStatus: Invalid configJson, using empty config:', error.message);
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

    const configResult = getUserConfig(currentUser.userId);
    const config = configResult.success ? configResult.config : {};

    config.isPublished = true;
    config.publishedAt = new Date().toISOString();

    const saveResult = saveUserConfig(currentUser.userId, config);
    if (!saveResult.success) {
      return {
        success: false,
        message: '設定の保存に失敗しました'
      };
    }

    return {
      success: true,
      message: '回答ボードを再公開しました',
      userId: currentUser.userId
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

    const configResult = getUserConfig(targetUser.userId);
    const config = configResult.success ? configResult.config : {};

    const wasPublished = config.isPublished === true;
    config.isPublished = false;
    config.publishedAt = null;
    config.lastAccessedAt = new Date().toISOString();

    const saveResult = saveUserConfig(targetUser.userId, config, { forceUpdate: false });
    if (!saveResult.success) {
      return createErrorResponse(`ボード状態の更新に失敗しました: ${saveResult.message || '詳細不明'}`);
    }

    return {
      success: true,
      message: wasPublished ? 'ボードを非公開にしました' : 'ボードは既に非公開です',
      boardPublished: false,
      userId: targetUser.userId
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
    const email = getCurrentEmail();
    if (!email || !isAdministrator(email)) {
      return createAdminRequiredError();
    }

    return {
      success: true,
      logs: [],
      message: 'Logs functionality available'
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
    const currentEmail = getCurrentEmail();
    if (!currentEmail || !isAdministrator(currentEmail)) {
      return createAdminRequiredError();
    }

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
    const currentEmail = getCurrentEmail();
    if (!currentEmail || !isAdministrator(currentEmail)) {
      return createAdminRequiredError();
    }

    const props = PropertiesService.getScriptProperties();

    const disabledReason = props.getProperty('APP_DISABLED_REASON') || '';
    const disabledBy = props.getProperty('APP_DISABLED_BY') || '';
    const disabledAt = props.getProperty('APP_DISABLED_AT') || '';

    props.deleteProperty('APP_DISABLED');
    props.deleteProperty('APP_DISABLED_REASON');
    props.deleteProperty('APP_DISABLED_BY');
    props.deleteProperty('APP_DISABLED_AT');

    props.setProperty('APP_ENABLED_BY', currentEmail);
    props.setProperty('APP_ENABLED_AT', new Date().toISOString());

    return createSuccessResponse('アプリケーションを再開しました', {
      enabledBy: currentEmail,
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
    const currentEmail = getCurrentEmail();
    if (!currentEmail || !isAdministrator(currentEmail)) {
      return createAdminRequiredError();
    }

    const props = PropertiesService.getScriptProperties();
    const isDisabled = props.getProperty('APP_DISABLED') === 'true';

    const status = {
      isDisabled,
      status: isDisabled ? 'disabled' : 'enabled',
      timestamp: new Date().toISOString()
    };

    if (isDisabled) {
      status.restriction = {
        reason: props.getProperty('APP_DISABLED_REASON') || '',
        disabledBy: props.getProperty('APP_DISABLED_BY') || '',
        disabledAt: props.getProperty('APP_DISABLED_AT') || ''
      };
    } else {
      status.lastEnabled = {
        enabledBy: props.getProperty('APP_ENABLED_BY') || '',
        enabledAt: props.getProperty('APP_ENABLED_AT') || ''
      };
    }

    return createSuccessResponse('アクセス制限状態を取得しました', status);
  } catch (error) {
    console.error('getAppAccessStatus error:', error.message);
    return createExceptionResponse(error, 'アクセス制限状態取得処理');
  }
}
