/**
 * @fileoverview 認証管理クラス - JWTトークンキャッシュと最適化
 */

class AuthManager {
  static CACHE_KEY = 'SA_TOKEN_CACHE';
  static TOKEN_EXPIRY_BUFFER = 300; // 5分のバッファ
  
  /**
   * キャッシュされたサービスアカウントトークンを取得
   * @returns {string} アクセストークン
   */
  static getServiceAccountToken() {
    const cache = CacheService.getScriptCache();
    let tokenData = cache.get(this.CACHE_KEY);
    
    if (tokenData) {
      tokenData = JSON.parse(tokenData);
      const now = Math.floor(Date.now() / 1000);
      
      // トークンがまだ有効な場合は再利用
      if (tokenData.expiresAt > now + this.TOKEN_EXPIRY_BUFFER) {
        debugLog('キャッシュされたトークンを使用');
        return tokenData.token;
      }
    }
    
    // 新しいトークンを生成
    const newToken = this._generateNewToken();
    return newToken;
  }
  
  /**
   * 新しいJWTトークンを生成
   * @private
   * @returns {string} アクセストークン
   */
  static _generateNewToken() {
    const props = PropertiesService.getScriptProperties();
    const serviceAccountCreds = JSON.parse(props.getProperty(SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS));
    
    const privateKey = serviceAccountCreds.private_key;
    const clientEmail = serviceAccountCreds.client_email;
    const tokenUrl = "https://www.googleapis.com/oauth2/v4/token";
    
    const now = Math.floor(Date.now() / 1000);
    const expiresAt = now + 3600; // 1時間後
    
    // JWT生成
    const jwtHeader = { alg: "RS256", typ: "JWT" };
    const jwtClaimSet = {
      iss: clientEmail,
      scope: "https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive",
      aud: tokenUrl,
      exp: expiresAt,
      iat: now
    };
    
    const encodedHeader = Utilities.base64EncodeWebSafe(JSON.stringify(jwtHeader));
    const encodedClaimSet = Utilities.base64EncodeWebSafe(JSON.stringify(jwtClaimSet));
    const signatureInput = encodedHeader + '.' + encodedClaimSet;
    const signature = Utilities.computeRsaSha256Signature(signatureInput, privateKey);
    const encodedSignature = Utilities.base64EncodeWebSafe(signature);
    const jwt = signatureInput + '.' + encodedSignature;
    
    // トークンリクエスト
    const response = UrlFetchApp.fetch(tokenUrl, {
      method: "post",
      contentType: "application/x-www-form-urlencoded",
      payload: {
        grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
        assertion: jwt
      }
    });
    
    const responseData = JSON.parse(response.getContentText());
    const accessToken = responseData.access_token;
    
    // キャッシュに保存
    const tokenData = {
      token: accessToken,
      expiresAt: expiresAt - this.TOKEN_EXPIRY_BUFFER
    };
    
    const cache = CacheService.getScriptCache();
    cache.put(this.CACHE_KEY, JSON.stringify(tokenData), 3600); // 1時間キャッシュ
    
    debugLog('新しいトークンを生成・キャッシュしました');
    return accessToken;
  }
  
  /**
   * トークンキャッシュをクリア
   */
  static clearTokenCache() {
    const cache = CacheService.getScriptCache();
    cache.remove(this.CACHE_KEY);
    debugLog('トークンキャッシュをクリアしました');
  }
}