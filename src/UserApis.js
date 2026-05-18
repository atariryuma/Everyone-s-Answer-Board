/**
 * @fileoverview UserApis - ユーザー関連 API（プロフィール取得、Welcome 表示状態
 *   管理など）。クロスファイル依存は下の global 宣言を参照。
 */

/* global getCurrentEmail, findUserByEmail, findUserByGoogleId, isAdministrator, getUserConfig, getConfigOrDefault, createUser, updateUser, getGoogleIdFromToken, getUserInfoCacheKey, getCachedProperty, setCachedProperty, createAuthError, createUserNotFoundError, createExceptionResponse, shouldEnforceDomainRestrictions, validateDomainAccess, logError_ */

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
        message: '管理者と同一ドメインのアカウントでアクセスしてください。'
      };
    }

    const user = findUserByEmail(email, { requestingUser: email });
    if (!user) {
      return createUserNotFoundError();
    }

    const config = getConfigOrDefault(user.userId);

    return { success: true, data: { config, userId: user.userId } };
  } catch (error) {
    logError_('getConfig', error);
    return createExceptionResponse(error);
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

    const tokenInfo = getGoogleIdFromToken();
    const googleId = tokenInfo?.googleId || '';

    let user = findUserByEmail(email, { requestingUser: email });

    // Google Workspace管理者がメール変更した場合、sub(googleId)は不変なので旧ユーザーを特定できる
    if (!user && googleId) {
      user = findUserByGoogleId(googleId, { requestingUser: email, skipAccessCheck: true });
      if (user) {
        const oldEmail = user.userEmail;
        const result = updateUser(user.userId, { userEmail: email }, { skipAccessCheck: true });
        if (result?.success) {
          user.userEmail = email;
          try {
            CacheService.getScriptCache().remove(getUserInfoCacheKey(oldEmail));
          } catch (e) {
            console.warn('processLoginAction: Failed to clear old email cache:', e.message);
          }
          // 管理者メールも自動更新
          const adminEmail = getCachedProperty('ADMIN_EMAIL');
          if (adminEmail && adminEmail.toLowerCase() === oldEmail.toLowerCase()) {
            // Store normalized: isAdministrator compares case-insensitively but
            // other write sites (setupApp) persist the lowercased form.
            setCachedProperty('ADMIN_EMAIL', email.toLowerCase());
            console.log('processLoginAction: ADMIN_EMAIL updated', { oldEmail, newEmail: email });
          }
        }
        console.log('processLoginAction: Email change detected', {
          userId: user.userId,
          oldEmail,
          newEmail: email,
          updated: !!result?.success
        });
      }
    }

    if (user && googleId && !user.googleId) {
      const result = updateUser(user.userId, { googleId });
      if (result?.success) {
        user.googleId = googleId;
      }
    }

    if (!user) {
      user = createUser(email, {}, { googleId });
      if (!user) {
        logError_('processLoginAction.createUser', new Error('createUser returned null'), { email });
        return {
          success: false,
          message: 'ユーザー登録に失敗しました。しばらくしてから再度お試しください。'
        };
      }
    }

    const baseUrl = getWebAppUrl();
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
    logError_('processLoginAction', error);
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
        message: '管理者と同一ドメインのアカウントでアクセスしてください。',
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
    logError_('checkUserAuthentication', error);
    return {
      success: false,
      authenticated: false,
      message: `Authentication check failed: ${error.message}`,
      error: error.message,
      timestamp: new Date().toISOString()
    };
  }
}
