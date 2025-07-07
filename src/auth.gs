/**
 * @fileoverview 認証管理 - JWTトークンキャッシュと最適化
 * GAS互換の関数ベースの実装
 */

// 認証管理のための定数
var AUTH_CACHE_KEY = 'SA_TOKEN_CACHE';
var TOKEN_EXPIRY_BUFFER = 300; // 5分のバッファ

/**
 * キャッシュされたサービスアカウントトークンを取得
 * @returns {string} アクセストークン
 */
function getServiceAccountTokenCached() {
  return cacheManager.get(AUTH_CACHE_KEY, generateNewServiceAccountToken, { ttl: 3500 });
}
  
/**
 * 新しいJWTトークンを生成
 * @returns {string} アクセストークン
 */
function generateNewServiceAccountToken() {
  var props = PropertiesService.getScriptProperties();
  var serviceAccountCreds = JSON.parse(props.getProperty(SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS));
  
  var privateKey = serviceAccountCreds.private_key.replace(/\\n/g, '\n'); // 改行文字を正規化
  var clientEmail = serviceAccountCreds.client_email;
  var tokenUrl = "https://www.googleapis.com/oauth2/v4/token";
  
  var now = Math.floor(Date.now() / 1000);
  var expiresAt = now + 3600; // 1時間後
  
  // JWT生成
  var jwtHeader = { alg: "RS256", typ: "JWT" };
  var jwtClaimSet = {
    iss: clientEmail,
    scope: "https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive",
    aud: tokenUrl,
    exp: expiresAt,
    iat: now
  };
  
  var encodedHeader = Utilities.base64EncodeWebSafe(JSON.stringify(jwtHeader));
  var encodedClaimSet = Utilities.base64EncodeWebSafe(JSON.stringify(jwtClaimSet));
  var signatureInput = encodedHeader + '.' + encodedClaimSet;
  var signature = Utilities.computeRsaSha256Signature(signatureInput, privateKey);
  var encodedSignature = Utilities.base64EncodeWebSafe(signature);
  var jwt = signatureInput + '.' + encodedSignature;
  
  // トークンリクエスト
  var response = UrlFetchApp.fetch(tokenUrl, {
    method: "post",
    contentType: "application/x-www-form-urlencoded",
    payload: {
      grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
      assertion: jwt
    }
  });
  
  var responseData = JSON.parse(response.getContentText());
  return responseData.access_token;
}

/**
 * トークンキャッシュをクリア
 */
function clearServiceAccountTokenCache() {
  cacheManager.remove(AUTH_CACHE_KEY);
  debugLog('トークンキャッシュをクリアしました');
}