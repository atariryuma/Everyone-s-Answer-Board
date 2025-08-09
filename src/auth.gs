/**
 * @fileoverview 認証管理 - JWTトークンキャッシュと最適化
 * GAS互換の関数ベースの実装
 */

// 認証管理のための定数
const AUTH_CACHE_KEY = 'SA_TOKEN_CACHE';
const TOKEN_EXPIRY_BUFFER = 300; // 5分のバッファ

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
  // 統一秘密情報管理システムで安全に取得
  const serviceAccountCreds = getSecureServiceAccountCreds();
  
  const privateKey = serviceAccountCreds.private_key.replace(/\\n/g, '\n'); // 改行文字を正規化
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
  const response = resilientUrlFetch(tokenUrl, {
    method: "post",
    contentType: "application/x-www-form-urlencoded",
    payload: {
      grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
      assertion: jwt
    },
    muteHttpExceptions: true
  });

  // レスポンスオブジェクト検証（resilientUrlFetchで既に検証済みだが、念のため）
  if (!response || typeof response.getResponseCode !== 'function') {
    throw new Error('サービスアカウント認証: 無効なレスポンスオブジェクトが返されました');
  }

  const responseCode = response.getResponseCode();
  if (responseCode !== 200) {
    const responseText = response.getContentText();
    errorLog('Token request failed. Status:', responseCode);
    errorLog('Response:', responseText);
    
    // より詳細なエラーメッセージ
    let errorMessage = 'サービスアカウントトークンの取得に失敗しました。';
    if (responseCode === 400) {
      errorMessage += ' 認証情報またはJWTの形式に問題があります。';
    } else if (responseCode === 401) {
      errorMessage += ' 認証情報が無効です。サービスアカウントキーを確認してください。';
    } else if (responseCode === 403) {
      errorMessage += ' 権限が不足しています。サービスアカウントの権限を確認してください。';
    } else {
      errorMessage += ` Status: ${responseCode}`;
    }
    
    throw new Error(errorMessage);
  }

  var responseData = JSON.parse(response.getContentText());
  if (!responseData.access_token) {
    errorLog('No access token in response:', responseData);
    throw new Error('アクセストークンが返されませんでした。サービスアカウント設定を確認してください。');
  }

  infoLog('Service account token generated successfully for:', clientEmail);
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
    const serviceAccountCreds = getSecureServiceAccountCreds();
    return serviceAccountCreds.client_email || 'メールアドレス不明';
  } catch (error) {
    warnLog('サービスアカウントメール取得エラー:', error.message);
    return 'メールアドレス取得エラー';
  }
}

// handleNewUserAuthentication関数は簡素化により削除されました
// 新規ユーザー認証は、confirmUserRegistrationとverifyAdminAccessの統合フローで処理されます

// 不要な複雑な認証システムを削除（簡素化のため）

/**
 * シンプル・確実な管理者権限検証（3重チェック）
 * メールアドレス + ユーザーID + アクティブ状態の照合
 * @param {string} userId - 検証するユーザーのID
 * @returns {boolean} 管理者権限がある場合は true、そうでなければ false
 */
function verifyAdminAccess(userId) {
  try {
    // 基本的な引数チェック
    if (!userId || typeof userId !== 'string' || userId.trim() === '') {
      warnLog('verifyAdminAccess: 無効なuserIdが渡されました:', userId);
      return false;
    }

    // 現在のログインユーザーのメールアドレスを取得
    const activeUserEmail = getCurrentUserEmail();
    if (!activeUserEmail) {
      warnLog('verifyAdminAccess: アクティブユーザーのメールアドレスが取得できませんでした');
      return false;
    }

    debugLog('verifyAdminAccess: 認証開始', { userId, activeUserEmail });

    // データベースからユーザー情報を取得（統一検索関数使用）
    let userFromDb = null;
    
    // まずは通常のキャッシュ付き検索を試行
    userFromDb = findUserById(userId);
    
    // 見つからない場合は強制フレッシュで再試行
    if (!userFromDb) {
      debugLog('verifyAdminAccess: 強制フレッシュで再検索中...');
      userFromDb = fetchUserFromDatabase('userId', userId, { forceFresh: true });
    }

    // ユーザーが見つからない場合は認証失敗
    if (!userFromDb) {
      warnLog('verifyAdminAccess: ユーザーが見つかりません:', { 
        userId, 
        activeUserEmail 
      });
      return false;
    }

    // 3重チェック実行
    // 1. メールアドレス照合
    const dbEmail = String(userFromDb.adminEmail || '').toLowerCase().trim();
    const currentEmail = String(activeUserEmail).toLowerCase().trim();
    const isEmailMatched = dbEmail && currentEmail && dbEmail === currentEmail;

    // 2. ユーザーID照合（念のため）
    const isUserIdMatched = String(userFromDb.userId) === String(userId);

    // 3. アクティブ状態確認
    const isActive = (userFromDb.isActive === true || 
                     userFromDb.isActive === 'true' || 
                     String(userFromDb.isActive).toLowerCase() === 'true');

    debugLog('verifyAdminAccess: 3重チェック結果:', {
      isEmailMatched,
      isUserIdMatched,
      isActive,
      dbEmail,
      currentEmail
    });

    // 3つの条件すべてが満たされた場合のみ認証成功
    if (isEmailMatched && isUserIdMatched && isActive) {
      infoLog('✅ verifyAdminAccess: 認証成功', { userId, email: activeUserEmail });
      return true;
    } else {
      warnLog('❌ verifyAdminAccess: 認証失敗', {
        userId,
        activeUserEmail,
        failures: {
          email: !isEmailMatched,
          userId: !isUserIdMatched,
          active: !isActive
        }
      });
      return false;
    }

  } catch (error) {
    errorLog('❌ verifyAdminAccess: 認証処理エラー:', error.message);
    return false;
  }
}


/**
 * ユーザーの最終アクセス時刻のみを更新（設定は保護）
 * @param {string} userId - 更新対象のユーザーID
 */
function updateUserLastAccess(userId) {
  try {
    if (!userId) {
      warnLog('updateUserLastAccess: userIdが指定されていません');
      return;
    }

    const now = new Date().toISOString();
    debugLog('最終アクセス時刻を更新:', userId, now);

    // lastAccessedAtフィールドのみを更新（他の設定は保護）
    updateUserField(userId, 'lastAccessedAt', now);

  } catch (error) {
    errorLog('updateUserLastAccess エラー:', error.message);
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

    // setupStatusがない場合、他のフィールドから推測（循環参照回避）
    // Note: この推測ロジックは循環参照を避けるため、formUrlベースに変更
    if (config.formCreated === true && config.formUrl && config.formUrl.trim()) {
      return 'completed';
    }

    return 'pending';

  } catch (error) {
    warnLog('getSetupStatusFromConfig JSON解析エラー:', error.message);
    return 'pending'; // エラー時はセットアップ未完了とみなす
  }
}
