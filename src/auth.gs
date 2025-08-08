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

  if (response.getResponseCode() !== 200) {
    errorLog('Token request failed. Status:', response.getResponseCode());
    errorLog('Response:', response.getContentText());
    throw new Error('サービスアカウントトークンの取得に失敗しました。認証情報を確認してください。Status: ' + response.getResponseCode());
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
      warnLog('verifyAdminAccess: 無効なuserIdが渡されました:', userId);
      return false;
    }

    // 現在操作しているGoogleアカウントのメールアドレスを取得
    var activeUserEmail = getCurrentUserEmail();
    if (!activeUserEmail) {
      warnLog('verifyAdminAccess: アクティブユーザーのメールアドレスが取得できませんでした');
      return false;
    }

    // データベースから指定されたIDのユーザー情報を取得
    // セキュリティ認証では最新データが必要だが、過度なキャッシュクリアを避ける
    debugLog('verifyAdminAccess: ユーザー検索開始 - userId:', userId);

    // まずキャッシュからユーザー情報を取得を試行
    var userFromDb = getOrFetchUserInfo(userId, 'userId', {
      useExecutionCache: false, // セキュリティ認証のため実行時キャッシュは使用しない
      ttl: 60 // 短いTTLで最新性を確保
    });

    // セキュリティクリティカルな場合のみ、追加検証として直接取得
    if (!userFromDb || !userFromDb.adminEmail) {
      debugLog('verifyAdminAccess: 直接データベース検索にフォールバック');
      userFromDb = fetchUserFromDatabase('userId', userId);
    }

    debugLog('verifyAdminAccess: ユーザー検索結果:', {
      found: !!userFromDb,
      userId: userFromDb ? userFromDb.userId : 'なし',
      adminEmail: userFromDb ? userFromDb.adminEmail : 'なし',
      isActive: userFromDb ? userFromDb.isActive : 'なし',
      activeUserEmail: activeUserEmail,
      cacheStrategy: 'optimized' // 最適化されたキャッシュ戦略を使用
    });

    if (!userFromDb) {
      warnLog('verifyAdminAccess: 指定されたIDのユーザーが存在しません。ID:', userId);
      return false;
    }

    // データベースのメールアドレスと、現在ログイン中のメールアドレスを比較
    var dbEmail = userFromDb.adminEmail ? String(userFromDb.adminEmail).trim() : '';
    var currentEmail = activeUserEmail ? String(activeUserEmail).trim() : '';
    var isEmailMatched = dbEmail && currentEmail &&
                        dbEmail.toLowerCase() === currentEmail.toLowerCase();

    debugLog('verifyAdminAccess: メールアドレス照合:', {
      dbEmail: dbEmail,
      currentEmail: currentEmail,
      isEmailMatched: isEmailMatched
    });

    // ユーザーがアクティブであるかを確認（型安全な判定）
    debugLog('verifyAdminAccess: isActive検証 - raw:', userFromDb.isActive, 'type:', typeof userFromDb.isActive);
    var isActive = (userFromDb.isActive === true ||
                    userFromDb.isActive === 'true' ||
                    String(userFromDb.isActive).toLowerCase() === 'true');
    debugLog('verifyAdminAccess: isActive結果:', isActive);

    if (isEmailMatched && isActive) {
      infoLog('✅ 管理者本人によるアクセスを確認しました:', activeUserEmail, 'UserID:', userId);
      return true; // メールが一致し、かつアクティブであれば成功
    } else {
      // セキュリティログの構造化
      const securityAlert = {
        timestamp: new Date().toISOString(),
        event: 'unauthorized_access_attempt',
        severity: 'high',
        details: {
          attemptedUserId: userId,
          dbEmail: userFromDb.adminEmail,
          activeUserEmail: activeUserEmail,
          isUserActive: isActive,
          sourceFunction: 'verifyAdminAccess'
        }
      };
      warnLog('🚨 セキュリティアラート:', JSON.stringify(securityAlert, null, 2));
      return false; // 一致しない、またはアクティブでない場合は失敗
    }
  } catch (e) {
    errorLog('verifyAdminAccess: 管理者検証中にエラーが発生しました:', e.message);
    return false;
  }
}

/**
 * ログインフローを処理し、適切なページにリダイレクトする
 * 既存ユーザーの設定を保護しつつ、セットアップ状況に応じたメッセージを表示
 * @param {string} userEmail ログインユーザーのメールアドレス
 * @returns {HtmlOutput} 表示するHTMLコンテンツ
 */
function processLoginFlow(userEmail) {
  try {
    if (!userEmail) {
      throw new Error('ユーザーメールアドレスが指定されていません');
    }

    // 1. ユーザー情報をデータベースから取得
    var userInfo = findUserByEmail(userEmail);

    // 2. 既存ユーザーの処理
    if (userInfo) {
      // 2a. アクティブユーザーの場合
      if (isTrue(userInfo.isActive)) {
        debugLog('processLoginFlow: 既存アクティブユーザー:', userEmail);

        // 最終アクセス時刻を更新（設定は保護）
        updateUserLastAccess(userInfo.userId);

        // セットアップ状況を確認してメッセージを調整
        const setupStatus = getSetupStatusFromConfig(userInfo.configJson);
        let welcomeMessage = '管理パネルへようこそ';

        if (setupStatus === 'pending') {
          welcomeMessage = 'セットアップを続行してください';
        } else if (setupStatus === 'completed') {
          welcomeMessage = 'おかえりなさい！';
        }

        const adminUrl = buildUserAdminUrl(userInfo.userId);
        return createSecureRedirect(adminUrl, welcomeMessage);
      }
      // 2b. 非アクティブユーザーの場合
      else {
        warnLog('processLoginFlow: 既存だが非アクティブなユーザー:', userEmail);
        return showErrorPage(
          'アカウントが無効です',
          'あなたのアカウントは現在無効化されています。管理者にお問い合わせください。'
        );
      }
    }
    // 3. 新規ユーザーの処理
    else {
      debugLog('processLoginFlow: 新規ユーザー登録開始:', userEmail);

      // 3a. 新規ユーザーデータを準備（統一された初期設定）
      const initialConfig = {
        // セットアップ管理
        setupStatus: 'pending',
        createdAt: new Date().toISOString(),

        // フォーム設定
        formCreated: false,
        formUrl: '',
        editFormUrl: '',

        // 公開設定
        appPublished: false,
        publishedSheetName: '',
        publishedSpreadsheetId: '',

        // 表示設定
        displayMode: 'anonymous',
        showCounts: false,
        sortOrder: 'newest',

        // メタデータ
        version: '1.0.0',
        lastModified: new Date().toISOString()
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
        warnLog('processLoginFlow: user not found after create:', newUser.userId);
      }
      debugLog('processLoginFlow: 新規ユーザー作成完了:', newUser.userId);

      // 3c. 新規ユーザーの管理パネルへリダイレクト
      const adminUrl = buildUserAdminUrl(newUser.userId);
      return createSecureRedirect(adminUrl, 'ようこそ！セットアップを開始してください');
    }
  } catch (error) {
    // 構造化エラーログの出力
    const errorInfo = {
      timestamp: new Date().toISOString(),
      function: 'processLoginFlow',
      userEmail: userEmail || 'unknown',
      errorType: error.name || 'UnknownError',
      message: error.message,
      stack: error.stack,
      severity: 'high' // ログインエラーは高重要度
    };
    errorLog('🚨 processLoginFlow 重大エラー:', JSON.stringify(errorInfo, null, 2));

    // ユーザーフレンドリーなエラーメッセージ
    const userMessage = error.message.includes('ユーザー')
      ? error.message
      : 'ログイン処理中に予期しないエラーが発生しました。しばらく待ってから再度お試しください。';

    return showErrorPage('ログインエラー', userMessage, error);
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
