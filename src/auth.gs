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
  if (typeof cacheManager === 'undefined') {
    return generateNewServiceAccountToken();
  }
  return cacheManager.get(AUTH_CACHE_KEY, generateNewServiceAccountToken, { 
    ttl: 3500, 
    enableMemoization: true 
  }); // メモ化対応でトークン取得を高速化
}
  
/**
 * 新しいJWTトークンを生成
 * @returns {string} アクセストークン
 */
function generateNewServiceAccountToken() {
  var props = PropertiesService.getScriptProperties();
  var serviceAccountCredsString = props.getProperty(SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS);
  
  if (!serviceAccountCredsString) {
    throw new Error('サービスアカウント認証情報が設定されていません。スクリプトプロパティで SERVICE_ACCOUNT_CREDS を設定してください。');
  }
  
  var serviceAccountCreds;
  try {
    serviceAccountCreds = JSON.parse(serviceAccountCredsString);
  } catch (e) {
    throw new Error('サービスアカウント認証情報の形式が正しくありません: ' + e.message);
  }
  
  if (!serviceAccountCreds.client_email || !serviceAccountCreds.private_key) {
    throw new Error('サービスアカウント認証情報にclient_emailまたはprivate_keyが含まれていません。');
  }
  
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
    },
    muteHttpExceptions: true
  });
  
  if (response.getResponseCode() !== 200) {
    console.error('Token request failed. Status:', response.getResponseCode());
    console.error('Response:', response.getContentText());
    throw new Error('サービスアカウントトークンの取得に失敗しました。認証情報を確認してください。Status: ' + response.getResponseCode());
  }
  
  var responseData = JSON.parse(response.getContentText());
  if (!responseData.access_token) {
    console.error('No access token in response:', responseData);
    throw new Error('アクセストークンが返されませんでした。サービスアカウント設定を確認してください。');
  }
  
  console.log('Service account token generated successfully for:', clientEmail);
  return responseData.access_token;
}

/**
 * トークンキャッシュをクリア
 */
function clearServiceAccountTokenCache() {
  cacheManager.remove(AUTH_CACHE_KEY);
  debugLog('トークンキャッシュをクリアしました');
}

/**
 * 設定されているサービスアカウントのメールアドレスを取得
 * @returns {string} サービスアカウントのメールアドレス
 */
function getServiceAccountEmail() {
  try {
    var props = PropertiesService.getScriptProperties();
    var serviceAccountCredsString = props.getProperty(SCRIPT_PROPS_KEYS.SERVICE_ACCOUNT_CREDS);
    
    if (!serviceAccountCredsString) {
      return 'サービスアカウント未設定';
    }
    
    var serviceAccountCreds = JSON.parse(serviceAccountCredsString);
    return serviceAccountCreds.client_email || 'メールアドレス不明';
  } catch (e) {
    return 'サービスアカウント設定エラー';
  }
}