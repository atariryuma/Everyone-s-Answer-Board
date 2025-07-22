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
  
  var privateKey = serviceAccountCreds.private_key.replace(/\n/g, '\n'); // 改行文字を正規化
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

/**
 * 指定されたuserIdと現在ログイン中のユーザーのメールアドレスが一致し、
 * かつアクティブな管理者権限を持っているかを確認します。
 * @param {string} userId - URLパラメータから受け取ったユーザーID
 * @returns {boolean} 検証に成功した場合は true、それ以外は false
 */
function verifyAdminAccess(userId) {
  try {
    // 引数チェック
    if (!userId || typeof userId !== 'string' || userId.trim() === '') {
      console.warn('verifyAdminAccess: 無効なuserIdが渡されました:', userId);
      return false;
    }

    // 現在操作しているGoogleアカウントのメールアドレスを取得
    var activeUserEmail = Session.getActiveUser().getEmail();
    if (!activeUserEmail) {
      console.warn('verifyAdminAccess: アクティブユーザーのメールアドレスが取得できませんでした');
      return false;
    }

    // データベースから指定されたIDのユーザー情報を取得（キャッシュ活用）
    // セキュリティ検証のため、最新データを確実に取得
    console.log('verifyAdminAccess: ユーザー検索開始 - userId:', userId);
    
    // キャッシュを完全にクリアして最新データを取得
    var cacheKey = 'user_' + userId;
    var emailCacheKey = 'email_' + activeUserEmail;
    
    console.log('verifyAdminAccess: キャッシュクリア実行');
    cacheManager.remove(cacheKey);
    cacheManager.remove(emailCacheKey);
    
    // 直接データベースから取得（キャッシュを経由しない）
    console.log('verifyAdminAccess: データベースから直接ユーザー検索');
    var userFromDb = fetchUserFromDatabase('userId', userId);
    
    console.log('verifyAdminAccess: データベース検索結果:', {
      found: !!userFromDb,
      userId: userFromDb ? userFromDb.userId : 'なし',
      adminEmail: userFromDb ? userFromDb.adminEmail : 'なし',
      isActive: userFromDb ? userFromDb.isActive : 'なし',
      activeUserEmail: activeUserEmail
    });

    if (!userFromDb) {
      console.warn('verifyAdminAccess: 指定されたIDのユーザーが存在しません。ID:', userId);
      return false;
    }

    // データベースのメールアドレスと、現在ログイン中のメールアドレスを比較
    var dbEmail = userFromDb.adminEmail ? String(userFromDb.adminEmail).trim() : '';
    var currentEmail = activeUserEmail ? String(activeUserEmail).trim() : '';
    var isEmailMatched = dbEmail && currentEmail && 
                        dbEmail.toLowerCase() === currentEmail.toLowerCase();
    
    console.log('verifyAdminAccess: メールアドレス照合:', {
      dbEmail: dbEmail,
      currentEmail: currentEmail,
      isEmailMatched: isEmailMatched
    });
    
    // ユーザーがアクティブであるかを確認（型安全な判定）
    console.log('verifyAdminAccess: isActive検証 - raw:', userFromDb.isActive, 'type:', typeof userFromDb.isActive);
    var isActive = (userFromDb.isActive === true || 
                    userFromDb.isActive === 'true' || 
                    String(userFromDb.isActive).toLowerCase() === 'true');
    console.log('verifyAdminAccess: isActive結果:', isActive);

    if (isEmailMatched && isActive) {
      console.log('✅ 管理者本人によるアクセスを確認しました:', activeUserEmail, 'UserID:', userId);
      return true; // メールが一致し、かつアクティブであれば成功
    } else {
      console.warn('⚠️ 不正なアクセス試行をブロックしました。' +
                  'DB Email: ' + userFromDb.adminEmail + 
                  ', Active Email: ' + activeUserEmail + 
                  ', Is Active: ' + isActive);
      return false; // 一致しない、またはアクティブでない場合は失敗
    }
  } catch (e) {
    console.error('verifyAdminAccess: 管理者検証中にエラーが発生しました:', e.message);
    return false;
  }
}

/**
 * ログインフローを処理し、適切なページに直接リダイレクトする
 * 自動リダイレクト強化版：即座に管理パネルへ遷移
 * @param {string} userEmail ログインユーザーのメールアドレス
 * @returns {HtmlOutput} 表示するHTMLコンテンツ
 */
function processLoginFlow(userEmail) {
  try {
    if (!userEmail) {
      throw new Error('ユーザーメールアドレスが指定されていません');
    }

    console.log('🔐 processLoginFlow開始:', userEmail);

    // 1. ユーザー情報をデータベースから取得
    var userInfo = findUserByEmail(userEmail);

    // 2. 既存ユーザーの処理
    if (userInfo) {
      // 2a. アクティブユーザーの場合
      if (isTrue(userInfo.isActive)) {
        console.log('✅ 既存アクティブユーザー確認:', userEmail);
        
        // 最終アクセス時刻を更新（設定は保護）
        updateUserLastAccess(userInfo.userId);
        
        // 統一されたURL生成関数を使用
        const adminUrl = buildAdminPanelUrl(userInfo.userId);
        console.log('🚀 管理パネルへリダイレクト:', adminUrl);
        
        // 高速自動リダイレクトを実行
        return createRedirectResponse(adminUrl);
      } 
      // 2b. 非アクティブユーザーの場合
      else {
        console.warn('⚠️ 非アクティブユーザーのアクセス試行:', userEmail);
        return showErrorPage(
          'アカウントが無効です', 
          'あなたのアカウントは現在無効化されています。管理者にお問い合わせください。'
        );
      }
    } 
    // 3. 新規ユーザーの処理
    else {
      console.log('👤 新規ユーザー登録開始:', userEmail);
      
      // 3a. 新規ユーザーデータを準備（初期設定でpending状態）
      const initialConfig = {
        setupStatus: 'pending',
        createdAt: new Date().toISOString(),
        formCreated: false,
        appPublished: false
      };
      
      const newUser = {
        userId: Utilities.getUuid(),
        adminEmail: userEmail,
        createdAt: new Date().toISOString(),
        isActive: true, // 即時有効化
        configJson: JSON.stringify(initialConfig),
        spreadsheetId: '',
        spreadsheetUrl: '',
        lastAccessedAt: new Date().toISOString()
      };
      
      // 3b. データベースに作成
      createUser(newUser);
      if (!waitForUserRecord(newUser.userId, 3000, 500)) {
        console.warn('⚠️ ユーザー作成後の確認に失敗:', newUser.userId);
      }
      console.log('✅ 新規ユーザー作成完了:', newUser.userId);
      
      // 3c. 新規ユーザーの管理パネルへ高速リダイレクト
      const adminUrl = buildAdminPanelUrl(newUser.userId);
      console.log('🚀 新規ユーザー管理パネルへリダイレクト:', adminUrl);
      
      return createRedirectResponse(adminUrl);
    }
  } catch (error) {
    console.error('❌ processLoginFlowでエラー:', error.stack);
    return showErrorPage('ログインエラー', 'ログイン処理中にエラーが発生しました。', error);
  }
}

/**
 * ユーザーの最終アクセス時刻のみを更新（設定は保護）
 * @param {string} userId - 更新対象のユーザーID
 */
function updateUserLastAccess(userId) {
  try {
    if (!userId) {
      console.warn('updateUserLastAccess: userIdが指定されていません');
      return;
    }
    
    const now = new Date().toISOString();
    console.log('最終アクセス時刻を更新:', userId, now);
    
    // lastAccessedAtフィールドのみを更新（他の設定は保護）
    updateUserField(userId, 'lastAccessedAt', now);
    
  } catch (error) {
    console.error('updateUserLastAccess エラー:', error.message);
  }
}

/**
 * configJsonからsetupStatusを安全に取得
 * @param {string} configJsonString - JSONエンコードされた設定文字列
 * @returns {string} setupStatus ('pending', 'completed', 'error')
 */
function getSetupStatusFromConfig(configJsonString) {
  try {
    if (!configJsonString || configJsonString.trim() === '' || configJsonString === '{}') {
      return 'pending'; // 空の場合はセットアップ未完了とみなす
    }
    
    const config = JSON.parse(configJsonString);
    
    // setupStatusが明示的に設定されている場合はそれを使用
    if (config.setupStatus) {
      return config.setupStatus;
    }
    
    // setupStatusがない場合、他のフィールドから推測
    if (config.appPublished === true && config.formCreated === true) {
      return 'completed';
    }
    
    return 'pending';
    
  } catch (error) {
    console.warn('getSetupStatusFromConfig JSON解析エラー:', error.message);
    return 'pending'; // エラー時はセットアップ未完了とみなす
  }
}
