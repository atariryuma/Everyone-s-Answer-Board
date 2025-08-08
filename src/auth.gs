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

/**
 * 指定されたユーザーの管理者権限を検証する - 統合検索システム対応版
 * @param {string} userId - 検証するユーザーのID
 * @returns {boolean} 管理者権限がある場合は true、そうでなければ false
 */
function verifyAdminAccess(userId) {
  try {
    const startTime = Date.now();
    
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

    debugLog('🔍 verifyAdminAccess: 統合ユーザー検索開始', {
      userId: userId,
      activeUserEmail: activeUserEmail,
      timestamp: new Date().toISOString()
    });

    // 統合ユーザー検索システムを使用
    var userFromDb = unifiedUserSearch(userId);
    const searchDuration = Date.now() - startTime;

    if (!userFromDb) {
      // 新規ユーザー作成直後の特別処理
      let isRecentlyCreated = false;
      try {
        const userProperties = PropertiesService.getUserProperties();
        const lastCreatedUserId = userProperties.getProperty('lastCreatedUserId');
        const lastCreatedTime = userProperties.getProperty('lastCreatedUserTime');
        
        if (lastCreatedUserId === userId && lastCreatedTime) {
          const timeDiff = Date.now() - parseInt(lastCreatedTime);
          isRecentlyCreated = timeDiff < 60000; // 60秒以内に作成された場合（時間を延長）
          
          debugLog('verifyAdminAccess: 新規ユーザー作成チェック:', {
            userId: userId,
            lastCreatedUserId: lastCreatedUserId,
            timeDiff: timeDiff,
            isRecentlyCreated: isRecentlyCreated,
            threshold: '60秒'
          });
        }
      } catch (propError) {
        warnLog('verifyAdminAccess: ユーザープロパティ取得エラー:', propError.message);
      }
      
      if (isRecentlyCreated) {
        warnLog('verifyAdminAccess: ⏰ 新規作成直後のユーザーです。段階的リトライを実行します:', userId);
        
        // 段階的リトライ（データベース同期を待つ）
        for (let retryCount = 1; retryCount <= 3; retryCount++) {
          const waitTime = retryCount * 1000; // 1秒、2秒、3秒
          warnLog(`verifyAdminAccess: リトライ ${retryCount}/3 - ${waitTime}ms待機後に再検索`);
          
          Utilities.sleep(waitTime);
          userFromDb = unifiedUserSearch(userId);
          
          if (userFromDb) {
            infoLog(`✅ verifyAdminAccess: リトライ${retryCount}回目で成功!`, userId);
            break;
          }
        }
        
        // まだ見つからない場合は仮承認
        if (!userFromDb) {
          warnLog('verifyAdminAccess: 🕒 リトライ後もデータなし - メールベースで仮承認を試行');
          const currentEmailLower = activeUserEmail ? activeUserEmail.toLowerCase().trim() : '';
          if (currentEmailLower) {
            infoLog('verifyAdminAccess: 🎫 新規ユーザー仮承認 - データベース同期完了を待つ間の暫定認証');
            return true; // 仮承認
          }
        }
      }
      
      if (!userFromDb) {
        const errorDetail = {
          requestedUserId: userId,
          activeUserEmail: activeUserEmail,
          isRecentlyCreated: isRecentlyCreated,
          searchDuration: searchDuration + 'ms',
          timestamp: new Date().toISOString()
        };
        errorLog('🚨 verifyAdminAccess: 統合検索システムでもユーザーが見つかりませんでした:', errorDetail);
        return false;
      }
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
  debugLog('processLoginFlow: Starting login flow for user:', userEmail); // 追加
  try {
    if (!userEmail) {
      debugLog('processLoginFlow: User email is not defined.'); // 追加
      throw new Error('ユーザーメールアドレスが指定されていません');
    }

    // 1. ユーザー情報をデータベースから取得
    debugLog('processLoginFlow: Attempting to find user by email:', userEmail); // 追加
    var userInfo = findUserByEmail(userEmail);
    debugLog('processLoginFlow: User info found:', userInfo ? 'Yes' : 'No'); // 追加

    // 2. 既存ユーザーの処理
    if (userInfo) {
      // 2a. アクティブユーザーの場合
      if (isTrue(userInfo.isActive)) {
        debugLog('processLoginFlow: Existing active user:', userEmail);

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
        debugLog('processLoginFlow: Redirecting to admin panel:', adminUrl); // 追加
        return createSecureRedirect(adminUrl, welcomeMessage);
      }
      // 2b. 非アクティブユーザーの場合
      else {
        warnLog('processLoginFlow: 既存だが非アクティブなユーザー:', userEmail);
        debugLog('processLoginFlow: User is inactive, showing error page.'); // 追加
        return showErrorPage(
          'アカウントが無効です',
          'あなたのアカウントは現在無効化されています。管理者にお問い合わせください。'
        );
      }
    }
    // 3. 新規ユーザーの処理
    else {
      debugLog('processLoginFlow: New user registration started:', userEmail);

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
        configJson: JSON.stringify(initialConfig),
        spreadsheetId: '',
        spreadsheetUrl: '',
        lastAccessedAt: new Date().toISOString(),
        isActive: true // 即時有効化
      };

      // 3b. データベースに作成
      createUser(newUser);
      
      // 3c. 新規ユーザー作成後のキャッシュクリア（権限確認問題の解決）
      debugLog('processLoginFlow: 新規ユーザー作成後、全キャッシュをクリアします', newUser.userId);
      try {
        // 実行キャッシュとScriptキャッシュを完全にクリア
        clearAllExecutionCache();
        CacheService.getScriptCache().removeAll();
        debugLog('✅ processLoginFlow: キャッシュクリア完了');
      } catch (cacheError) {
        warnLog('⚠️ processLoginFlow: キャッシュクリアでエラー:', cacheError.message);
      }
      
      // 3d. ユーザー作成後の検証を強化（待機時間とリトライ回数を増加）
      debugLog('processLoginFlow: ユーザー作成完了、データベース検証を実行中...', newUser.userId);
      if (!waitForUserRecord(newUser.userId, 5000, 300)) { // 5秒間待機、300ms間隔でリトライ
        warnLog('processLoginFlow: ユーザーレコード検証でタイムアウト:', newUser.userId);
        
        // 追加検証: 複数の方法でユーザーの存在を確認
        let verifyUser = null;
        const verificationMethods = [];
        
        // 方法1: findUserById
        try {
          verifyUser = findUserById(newUser.userId, { useExecutionCache: false, forceRefresh: true });
          verificationMethods.push({ method: 'findUserById', success: !!verifyUser });
        } catch (error) {
          verificationMethods.push({ method: 'findUserById', error: error.message });
        }
        
        // 方法2: fetchUserFromDatabase (直接アクセス)
        if (!verifyUser) {
          try {
            verifyUser = fetchUserFromDatabase('userId', newUser.userId, {
              enableDiagnostics: false,
              autoRepair: false,
              retryCount: 1
            });
            verificationMethods.push({ method: 'fetchUserFromDatabase', success: !!verifyUser });
          } catch (error) {
            verificationMethods.push({ method: 'fetchUserFromDatabase', error: error.message });
          }
        }
        
        // 方法3: adminEmailでの検索
        if (!verifyUser) {
          try {
            verifyUser = fetchUserFromDatabase('adminEmail', newUser.adminEmail, {
              enableDiagnostics: false,
              autoRepair: false,
              retryCount: 1
            });
            verificationMethods.push({ method: 'fetchByEmail', success: !!verifyUser });
          } catch (error) {
            verificationMethods.push({ method: 'fetchByEmail', error: error.message });
          }
        }
        
        debugLog('processLoginFlow: ユーザー検証結果:', {
          userId: newUser.userId,
          email: newUser.adminEmail,
          found: !!verifyUser,
          verificationMethods: verificationMethods
        });
        
        if (!verifyUser) {
          errorLog('processLoginFlow: 🚨 全ての検証方法でユーザーが見つかりません:', {
            userId: newUser.userId,
            email: newUser.adminEmail,
            verificationMethods: verificationMethods
          });
          throw new Error('ユーザー登録の処理中にエラーが発生しました。データベースの同期に時間がかかっている可能性があります。数分後に再度お試しください。');
        } else {
          infoLog('processLoginFlow: ✅ ユーザー検証完了:', {
            userId: newUser.userId,
            verifiedBy: verificationMethods.find(m => m.success)?.method || 'unknown'
          });
        }
      } else {
        infoLog('processLoginFlow: ✅ ユーザーレコード検証が完了しました:', newUser.userId);
      }
      
      debugLog('processLoginFlow: New user creation completed:', newUser.userId);

      // 3e. 新規ユーザー作成の完了を記録（管理パネルアクセス時の参考用）
      try {
        const userProperties = PropertiesService.getUserProperties();
        userProperties.setProperty('lastCreatedUserId', newUser.userId);
        userProperties.setProperty('lastCreatedUserTime', Date.now().toString());
        debugLog('✅ processLoginFlow: 新規ユーザー作成情報を記録しました');
      } catch (propError) {
        warnLog('⚠️ processLoginFlow: ユーザープロパティ記録でエラー:', propError.message);
      }

      // 3f. 新規ユーザーの管理パネルへリダイレクト
      const adminUrl = buildUserAdminUrl(newUser.userId);
      debugLog('processLoginFlow: Redirecting new user to admin panel:', adminUrl);
      
      // 新規ユーザー登録完了メッセージを明確に表示
      return createSecureRedirect(adminUrl, '✨ 新規ユーザー登録が完了しました！セットアップを開始してください');
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
    debugLog('processLoginFlow: Error in login flow. Error:', error.message); // 追加

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
