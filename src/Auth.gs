/**
 * @fileoverview Auth - 統一認証システム
 *
 * 責任範囲:
 * - 真のサービスアカウント認証
 * - ユーザーセッション管理
 * - JWT-based認証フロー
 * - Zero-Dependency Architecture準拠
 */

/* global ServiceFactory, Data, validateEmail, getCurrentEmail, getUserConfig, SLEEP_MS, CACHE_DURATION */

/**
 * 統一認証クラス
 * CLAUDE.md準拠のZero-Dependency実装
 */
class Auth {
  /**
   * 真のサービスアカウント認証
   * JWT-based OAuth 2.0 flow implementation
   * @returns {Object} Service account authentication object
   */
  static serviceAccount() {
    try {
      const config = this.getServiceAccountConfig();
      if (!config) {
        throw new Error('Service account configuration not found');
      }

      // Create JWT assertion for OAuth 2.0
      const jwt = this.createJWTAssertion(config);

      // Exchange JWT for access token
      const token = this.exchangeJWTForToken(jwt);

      return {
        token,
        email: config.client_email,
        isValid: !!token,
        type: 'service_account'
      };
    } catch (error) {
      console.error('Auth.serviceAccount error:', error.message);
      return {
        token: null,
        email: null,
        isValid: false,
        type: 'service_account',
        error: error.message
      };
    }
  }

  /**
   * ユーザーセッション取得
   * Session API統一アクセス
   * @returns {Object} User session object
   */
  static session() {
    try {
      const session = ServiceFactory.getSession();
      const {email} = session;
      return {
        email,
        isValid: !!email,
        type: 'user_session'
      };
    } catch (error) {
      console.warn('Auth.session error:', error.message);
      return {
        email: null,
        isValid: false,
        type: 'user_session',
        error: error.message
      };
    }
  }

  /**
   * サービスアカウント設定取得
   * @returns {Object|null} Service account configuration
   */
  static getServiceAccountConfig() {
    try {
      // 🔧 CLAUDE.md準拠: 循環参照解決 - 直接PropertiesService使用
      const props = PropertiesService.getScriptProperties();
      const credsJson = props.getProperty('SERVICE_ACCOUNT_CREDS');

      if (!credsJson) {
        return null;
      }

      const config = JSON.parse(credsJson);

      // Validate required fields
      if (!config.client_email || !config.private_key || !config.project_id) {
        throw new Error('Invalid service account configuration');
      }

      return config;
    } catch (error) {
      console.error('Auth.getServiceAccountConfig error:', error.message);
      return null;
    }
  }

  /**
   * JWT assertion作成
   * RSA-SHA256 signature generation
   * @param {Object} config - Service account configuration
   * @returns {string} JWT assertion
   */
  static createJWTAssertion(config) {
    const now = Math.floor(Date.now() / 1000);

    const header = {
      alg: 'RS256',
      typ: 'JWT'
    };

    const payload = {
      iss: config.client_email,
      scope: 'https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive',
      aud: 'https://oauth2.googleapis.com/token',
      exp: now + 3600,
      iat: now
    };

    const headerB64 = Utilities.base64EncodeWebSafe(JSON.stringify(header)).replace(/=+$/, '');
    const payloadB64 = Utilities.base64EncodeWebSafe(JSON.stringify(payload)).replace(/=+$/, '');

    // Validate base64 encoding results before template literal usage
    if (!headerB64 || !payloadB64) {
      throw new Error('Failed to encode JWT components - invalid base64 result');
    }

    const signatureInput = `${headerB64}.${payloadB64}`;

    // Use Utilities.computeRsaSha256Signature for real RSA-SHA256 signing
    const signature = Utilities.computeRsaSha256Signature(signatureInput, config.private_key);
    const signatureB64 = Utilities.base64EncodeWebSafe(signature).replace(/=+$/, '');

    // Validate signature encoding before final JWT assembly
    if (!signatureB64) {
      throw new Error('Failed to encode JWT signature - invalid signature result');
    }

    return `${signatureInput}.${signatureB64}`;
  }

  /**
   * JWT assertion をアクセストークンに交換
   * @param {string} jwt - JWT assertion
   * @returns {string|null} Access token
   */
  static exchangeJWTForToken(jwt) {
    try {
      const response = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded'
        },
        payload: `grant_type=urn:ietf:params:oauth:grant-type:jwt-bearer&assertion=${jwt}`,
        muteHttpExceptions: true
      });

      const responseData = JSON.parse(response.getContentText());

      if (response.getResponseCode() === 200 && responseData.access_token) {
        return responseData.access_token;
      } else {
        throw new Error(`Token exchange failed: ${JSON.stringify(responseData)}`);
      }
    } catch (error) {
      console.error('Auth.exchangeJWTForToken error:', error.message);
      return null;
    }
  }

  /**
   * 認証診断
   * @returns {Object} Diagnostic information
   */
  static diagnose() {
    const results = {
      service: 'Auth',
      timestamp: new Date().toISOString(),
      checks: []
    };

    // Service account check
    try {
      const sa = this.serviceAccount();
      results.checks.push({
        name: 'Service Account Authentication',
        status: sa.isValid ? '✅' : '❌',
        details: sa.isValid ? 'Service account authentication working' : sa.error || 'Authentication failed'
      });
    } catch (error) {
      results.checks.push({
        name: 'Service Account Authentication',
        status: '❌',
        details: error.message
      });
    }

    // User session check
    try {
      const session = this.session();
      results.checks.push({
        name: 'User Session',
        status: session.isValid ? '✅' : '⚠️',
        details: session.isValid ? `User: ${session.email}` : 'No active session'
      });
    } catch (error) {
      results.checks.push({
        name: 'User Session',
        status: '❌',
        details: error.message
      });
    }

    results.overall = results.checks.every(check => check.status === '✅') ? '✅' : '⚠️';
    return results;
  }

  // ===========================================
  // 🔐 統一アクセス制御ミドルウェア（Administrator/Editor/Viewer）
  // ===========================================

  /**
   * 統一アクセス制御チェック
   * @param {string} mode - アクセスモード
   * @param {Object} params - パラメータ
   * @returns {Object} 認可結果
   */
  static checkAccess(mode, params = {}) {
    const session = ServiceFactory.getSession();

    const accessCheckers = {
      login: () => ({ allowed: true }),
      admin: () => this.checkEditorAccess(session),
      appSetup: () => this.checkAdministratorAccess(session),
      view: () => this.checkViewerAccess(params),
      main: () => ({ allowed: false, redirect: 'AccessRestricted' })
    };

    const checker = accessCheckers[mode];
    if (!checker) {
      console.warn(`Auth.checkAccess: Unknown mode '${mode}'`);
      return { allowed: false, redirect: 'AccessRestricted' };
    }

    const result = checker();

    // 監査ログ
    this.logAuthEvent(mode, session.email, result.allowed);

    return result;
  }

  /**
   * Administrator（システム管理者）アクセス確認
   */
  static checkAdministratorAccess(session) {
    if (!session.email) {
      return { allowed: false, redirect: 'LoginPage' };
    }

    // 🔧 重複解決: UserService.isAdministratorを直接使用（循環依存を避けるためstaticアクセス）
    const isAdmin = (() => {
      try {
        const adminEmail = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL');
        return adminEmail && session.email && session.email.toLowerCase() === adminEmail.toLowerCase();
      } catch (error) {
        console.error('Administrator check error:', error.message);
        return false;
      }
    })();

    if (!isAdmin) {
      console.warn('Administrator access denied:', `${session.email.split('@')[0]}@***`);
      return { allowed: false, redirect: 'AccessRestricted' };
    }

    return {
      allowed: true,
      accessLevel: 'administrator',
      email: session.email
    };
  }

  /**
   * Editor（編集者）アクセス確認（安全なユーザー作成込み）
   */
  static checkEditorAccess(session) {
    if (!session.email) {
      return { allowed: false, redirect: 'LoginPage' };
    }

    let user = Data.findUserByEmail(session.email);
    if (!user) {
      user = this.createEditorSafely(session.email);
      if (!user) {
        return { allowed: false, redirect: 'ErrorBoundary' };
      }
    }

    return {
      allowed: true,
      accessLevel: 'editor',
      user,
      email: session.email
    };
  }

  /**
   * Viewer（閲覧者）アクセス確認
   */
  static checkViewerAccess(params) {
    const { userId } = params;
    if (!userId) {
      return { allowed: false, redirect: 'AccessRestricted' };
    }

    const user = Data.findUserById(userId);
    if (!user) {
      return {
        allowed: false,
        redirect: 'ErrorBoundary',
        error: 'User not found'
      };
    }

    // 🔧 循環依存解決: getUserConfigの直接使用を避けてData経由でconfig取得
    let config = {};
    try {
      if (user && user.configJson) {
        config = JSON.parse(user.configJson);
      }
    } catch (error) {
      console.warn('Auth.checkViewerAccess: config parse error:', error.message);
      config = {};
    }

    if (!config.isPublished) {
      return {
        allowed: false,
        redirect: 'ErrorBoundary',
        error: 'Board not published'
      };
    }

    return {
      allowed: true,
      accessLevel: 'viewer',
      user,
      config
    };
  }

  /**
   * 安全なEditor作成（RequestGate代替）
   */
  static createEditorSafely(email) {
    const lockKey = `create_editor_${email.replace(/[^a-zA-Z0-9]/g, '_')}`;
    const cache = ServiceFactory.getCache();

    // 排他制御（Cache-based mutex）
    if (cache.get(lockKey)) {
      console.info('Editor creation in progress, waiting...');
      Utilities.sleep(SLEEP_MS.LONG);
      return Data.findUserByEmail(email);
    }

    try {
      cache.put(lockKey, true, CACHE_DURATION.MEDIUM); // 30秒ロック

      // 再度確認（double-check）
      let user = Data.findUserByEmail(email);
      if (user) return user;

      // 新規作成
      user = Data.createUser(email);
      if (user) {
        console.log('✅ Editor user created successfully:', `${email.split('@')[0]}@***`);
        return user;
      }

      // フォールバック作成
      return this.createFallbackUser(email);

    } catch (error) {
      console.error('Editor creation failed:', error.message);
      return this.createFallbackUser(email);
    } finally {
      cache.remove(lockKey);
    }
  }

  /**
   * フォールバック用ユーザー作成
   */
  static createFallbackUser(email) {
    return {
      userId: Utilities.getUuid(),
      userEmail: email,
      isActive: true,
      configJson: JSON.stringify({
        setupStatus: 'pending',
        isPublished: false,
        createdAt: new Date().toISOString()
      }),
      lastModified: new Date().toISOString()
    };
  }


  /**
   * 監査ログ記録
   */
  static logAuthEvent(mode, email, allowed) {
    console.info('Auth Event:', {
      mode,
      email: email ? `${email.split('@')[0]}@***` : 'anonymous',
      allowed,
      timestamp: new Date().toISOString()
    });
  }
}

// Export for global access
if (typeof globalThis !== 'undefined') {
  globalThis.Auth = Auth;
} else if (typeof global !== 'undefined') {
  global.Auth = Auth;
} else {
  this.Auth = Auth;
}