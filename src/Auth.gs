/**
 * @fileoverview Auth - 統一認証システム
 *
 * 責任範囲:
 * - 真のサービスアカウント認証
 * - ユーザーセッション管理
 * - JWT-based認証フロー
 * - Zero-Dependency Architecture準拠
 */

/* global ServiceFactory */

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
      const email = Session.getActiveUser().getEmail();
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
      const props = ServiceFactory.getProperties();
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

    const signatureInput = `${headerB64  }.${  payloadB64}`;

    // Use Utilities.computeRsaSha256Signature for real RSA-SHA256 signing
    const signature = Utilities.computeRsaSha256Signature(signatureInput, config.private_key);
    const signatureB64 = Utilities.base64EncodeWebSafe(signature).replace(/=+$/, '');

    return `${signatureInput  }.${  signatureB64}`;
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
        payload: `grant_type=urn:ietf:params:oauth:grant-type:jwt-bearer&assertion=${  jwt}`,
        muteHttpExceptions: true
      });

      const responseData = JSON.parse(response.getContentText());

      if (response.getResponseCode() === 200 && responseData.access_token) {
        return responseData.access_token;
      } else {
        throw new Error(`Token exchange failed: ${  JSON.stringify(responseData)}`);
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
}

// Export for global access
const __globalRoot = (typeof globalThis !== 'undefined') ? globalThis : (typeof global !== 'undefined' ? global : this);
__globalRoot.Auth = Auth;