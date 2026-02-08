/**
 * @fileoverview UserApis - ユーザー関連API
 *
 * ✅ V8ランタイム対応（2022年6月アップデート準拠）
 * - 関数定義は順序に関係なく呼び出し可能
 * - グローバルスコープでのコード実行を完全排除
 *
 * 依存関係（呼び出す関数）:
 * - getCurrentEmail() - main.jsで定義
 * - findUserByEmail() - DatabaseCore.jsで定義
 * - isAdministrator() - main.jsで定義
 * - getUserConfig() - ConfigService.jsで定義
 * - createUser() - UserService.jsで定義
 * - createAuthError() - helpers.jsで定義
 * - createUserNotFoundError() - helpers.jsで定義
 * - createExceptionResponse() - helpers.jsで定義
 *
 * 移動元: main.js
 * 移動日: 2025-12-13
 */

/* global getCurrentEmail, findUserByEmail, isAdministrator, getUserConfig, createUser, createAuthError, createUserNotFoundError, createExceptionResponse, ScriptApp, Utilities, shouldEnforceDomainRestrictions, validateDomainAccess */

/**
 * ドメインアクセス検証（必要時のみ）
 * @param {string} email - 検証対象メール
 * @returns {Object} 検証結果
 */
function ensureDomainAccess(email) {
  const enforce = (typeof shouldEnforceDomainRestrictions === 'function')
    ? shouldEnforceDomainRestrictions()
    : true;

  if (!enforce || typeof validateDomainAccess !== 'function') {
    return { allowed: true };
  }

  return validateDomainAccess(email, {
    allowIfAdminUnconfigured: true,
    allowIfEmailMissing: false
  });
}


/**
 * Get user configuration - unified function for current user
 * @returns {Object} User configuration result
 */
function getConfig() {
  try {
    const email = getCurrentEmail();
    if (!email) {
      return createAuthError();
    }

    const domainAccess = ensureDomainAccess(email);
    if (!domainAccess.allowed) {
      return {
        success: false,
        message: '管理者と同一ドメインのアカウントでアクセスしてください'
      };
    }

    const user = findUserByEmail(email, { requestingUser: email });
    if (!user) {
      return createUserNotFoundError();
    }

    const configResult = getUserConfig(user.userId);
    const config = configResult.success ? configResult.config : {};

    return { success: true, data: { config, userId: user.userId } };
  } catch (error) {
    console.error('getConfig error:', error.message);
    return createExceptionResponse(error);
  }
}

/**
 * Batched user configuration retrieval
 * Batch operation: Get email, user, and config in single coordinated call
 * @returns {Object} Batched result with email, user, and config
 */
function getBatchedUserConfig() {
  try {
    const email = getCurrentEmail();
    if (!email) {
      return {
        success: false,
        authenticated: false,
        error: 'ユーザー認証が必要です',
        user: null,
        config: null,
        authError: createAuthError()
      };
    }

    const domainAccess = ensureDomainAccess(email);
    if (!domainAccess.allowed) {
      return {
        success: false,
        authenticated: true,
        error: '管理者と同一ドメインのアカウントでアクセスしてください',
        email,
        user: null,
        config: null
      };
    }

    const user = findUserByEmail(email, { requestingUser: email });
    if (!user) {
      return {
        success: false,
        authenticated: true,
        error: 'ユーザー情報が見つかりません',
        email,
        user: null,
        config: null,
        userError: createUserNotFoundError()
      };
    }

    const configResult = getUserConfig(user.userId);
    const config = configResult.success ? configResult.config : {};

    return {
      success: true,
      authenticated: true,
      email,
      user,
      config,
      configSuccess: configResult.success,
      configMessage: configResult.message || 'Configuration retrieved successfully'
    };

  } catch (error) {
    console.error('getBatchedUserConfig error:', error.message);
    return {
      success: false,
      authenticated: false,
      error: `ユーザー設定取得エラー: ${error.message}`,
      user: null,
      config: null,
      exception: createExceptionResponse(error)
    };
  }
}


/**
 * Process login action - handles user login flow
 * @returns {Object} Login result with redirect URL
 */
function processLoginAction() {
  try {
    const email = getCurrentEmail();
    if (!email) {
      return {
        success: false,
        message: 'Authentication failed - no user email available'
      };
    }

    const domainAccess = ensureDomainAccess(email);
    if (!domainAccess.allowed) {
      return {
        success: false,
        message: '管理者と同一ドメインのアカウントでログインしてください。'
      };
    }

    let user = findUserByEmail(email, { requestingUser: email });
    if (!user) {
      user = createUser(email);
      if (!user) {
        console.warn('createUser failed, creating fallback user object');
        user = {
          userId: Utilities.getUuid(),
          userEmail: email,
          isActive: true,
          configJson: JSON.stringify({
            setupStatus: 'pending',
            isPublished: false
          }),
          lastModified: new Date().toISOString()
        };
      }
    }

    const baseUrl = getWebAppUrl(); // eslint-disable-line no-undef
    if (!baseUrl) {
      return {
        success: false,
        message: 'Web app URL not available (deployment may be initializing)'
      };
    }
    if (!user || !user.userId) {
      return {
        success: false,
        message: 'Invalid user data'
      };
    }
    const redirectUrl = `${baseUrl}?mode=admin&userId=${user.userId}`;

    return {
      success: true,
      message: 'Login successful',
      redirectUrl,
      adminUrl: redirectUrl,
      url: redirectUrl,
      data: {
        userId: user.userId,
        email,
        redirectUrl
      }
    };
  } catch (error) {
    console.error('processLoginAction error:', error.message);
    return {
      success: false,
      message: `Login failed: ${error.message || '詳細不明'}`
    };
  }
}

/**
 * Check user authentication status
 * @returns {Object} Authentication status with user details
 */
function checkUserAuthentication() {
  try {
    const email = getCurrentEmail();
    if (!email) {
      return {
        success: false,
        authenticated: false,
        message: 'No user session found',
        timestamp: new Date().toISOString()
      };
    }

    const domainAccess = ensureDomainAccess(email);
    if (!domainAccess.allowed) {
      return {
        success: false,
        authenticated: true,
        message: '管理者と同一ドメインのアカウントでアクセスしてください',
        timestamp: new Date().toISOString()
      };
    }

    const user = findUserByEmail(email, { requestingUser: email });
    const userExists = !!user;

    const isAdminUser = isAdministrator(email);

    let hasValidConfig = false;
    if (user) {
      try {
        const configResult = getUserConfig(user.userId);
        hasValidConfig = configResult.success;
      } catch (configError) {
        console.warn('Config check failed:', configError.message);
      }
    }

    const authLevel = isAdminUser ? 'administrator' : userExists ? 'user' : 'guest';

    return {
      success: true,
      authenticated: true,
      email,
      userExists,
      isAdministrator: isAdminUser,
      hasValidConfig,
      authLevel,
      userId: user ? user.userId : null,
      sessionValid: true,
      message: `User authenticated as ${authLevel}`,
      timestamp: new Date().toISOString()
    };

  } catch (error) {
    console.error('checkUserAuthentication error:', error.message);
    return {
      success: false,
      authenticated: false,
      message: `Authentication check failed: ${error.message}`,
      error: error.message,
      timestamp: new Date().toISOString()
    };
  }
}
