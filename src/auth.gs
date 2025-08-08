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
 * 緊急メール認証システム - データベースで見つからない新規ユーザーの認証
 * @param {string} userId - 検証するユーザーID
 * @param {string} activeUserEmail - 現在のログインユーザーのメールアドレス
 * @param {Array} detectionResults - 新規ユーザー検出結果
 * @returns {Object} 認証結果 {approved: boolean, method: string, reason: string}
 */
function performEmergencyEmailAuth(userId, activeUserEmail, detectionResults) {
  const authResult = {
    approved: false,
    method: 'none',
    reason: 'no_match',
    details: {}
  };
  
  try {
    const currentEmail = activeUserEmail ? activeUserEmail.toLowerCase().trim() : '';
    if (!currentEmail) {
      authResult.reason = 'no_active_email';
      return authResult;
    }
    
    // 方法1: ScriptPropertiesで記録されたメールアドレスとの一致確認
    for (const result of detectionResults) {
      if (result.method === 'ScriptProperties' && result.data && result.data.email) {
        const recordedEmail = result.data.email.toLowerCase().trim();
        if (recordedEmail === currentEmail) {
          authResult.approved = true;
          authResult.method = 'ScriptProperties_email_match';
          authResult.reason = 'email_match_in_script_props';
          authResult.details = {
            recordedEmail: recordedEmail,
            timeDiff: result.timeDiff,
            key: result.key
          };
          return authResult;
        }
      }
    }
    
    // 方法2: EmailBasedでの検出結果を使用
    for (const result of detectionResults) {
      if (result.method === 'EmailBased' && result.data) {
        const recordedEmail = result.data.email.toLowerCase().trim();
        if (recordedEmail === currentEmail) {
          authResult.approved = true;
          authResult.method = 'EmailBased_detection';
          authResult.reason = 'email_based_detection_match';
          authResult.details = {
            recordedEmail: recordedEmail,
            timeDiff: result.timeDiff,
            key: result.key
          };
          return authResult;
        }
      }
    }
    
    // 方法3: ドメイン一致確認（教育機関ドメインなどの信頼できるドメイン）
    const trustedDomains = ['naha-okinawa.ed.jp', 'gmail.com']; // 必要に応じて拡張
    const emailDomain = currentEmail.split('@')[1];
    
    if (trustedDomains.includes(emailDomain)) {
      // さらに最近の作成記録があるかチェック
      try {
        const scriptProps = PropertiesService.getScriptProperties();
        const allProps = scriptProps.getProperties();
        const recentThreshold = Date.now() - (5 * 60 * 1000); // 5分以内
        
        for (const [key, value] of Object.entries(allProps)) {
          if (key.startsWith('newUser_') && key.toLowerCase().includes(emailDomain)) {
            try {
              const data = JSON.parse(value);
              const createdTime = parseInt(data.createdTime);
              
              if (createdTime > recentThreshold) {
                authResult.approved = true;
                authResult.method = 'trusted_domain_recent';
                authResult.reason = 'trusted_domain_with_recent_activity';
                authResult.details = {
                  domain: emailDomain,
                  recentActivity: true,
                  timeDiff: Date.now() - createdTime,
                  key: key
                };
                return authResult;
              }
            } catch (parseError) {
              // パース失敗は無視
            }
          }
        }
      } catch (domainCheckError) {
        warnLog('performEmergencyEmailAuth: ドメインチェックでエラー:', domainCheckError.message);
      }
    }
    
    authResult.reason = 'no_emergency_criteria_met';
    authResult.details = {
      currentEmail: currentEmail,
      detectionResultsCount: detectionResults.length,
      domain: emailDomain
    };
    
  } catch (error) {
    authResult.reason = 'emergency_auth_error';
    authResult.details = { error: error.message };
    errorLog('performEmergencyEmailAuth: エラー:', error.message);
  }
  
  return authResult;
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
      // 新規ユーザー作成直後の特別処理 - 改善された検出ロジック
      let isRecentlyCreated = false;
      let detectionSource = 'none';
      const detectionResults = [];
      
      try {
        // 方法1: UserPropertiesでの検出
        try {
          const userProperties = PropertiesService.getUserProperties();
          const lastCreatedUserId = userProperties.getProperty('lastCreatedUserId');
          const lastCreatedTime = userProperties.getProperty('lastCreatedUserTime');
          const lastCreatedEmail = userProperties.getProperty('lastCreatedUserEmail');
          
          const userPropsResult = {
            method: 'UserProperties',
            lastCreatedUserId: lastCreatedUserId,
            lastCreatedTime: lastCreatedTime,
            lastCreatedEmail: lastCreatedEmail,
            found: false,
            timeDiff: null
          };
          
          if (lastCreatedUserId === userId && lastCreatedTime) {
            const timeDiff = Date.now() - parseInt(lastCreatedTime);
            userPropsResult.timeDiff = timeDiff;
            userPropsResult.found = timeDiff < 120000; // 120秒以内に作成された場合
            
            if (userPropsResult.found) {
              isRecentlyCreated = true;
              detectionSource = 'UserProperties';
            }
          }
          detectionResults.push(userPropsResult);
        } catch (userPropError) {
          detectionResults.push({ method: 'UserProperties', error: userPropError.message });
        }
        
        // 方法2: ScriptPropertiesでの検出（フォールバック）
        if (!isRecentlyCreated) {
          try {
            const scriptProps = PropertiesService.getScriptProperties();
            const allScriptProps = scriptProps.getProperties();
            
            Object.keys(allScriptProps).forEach(key => {
              if (key.startsWith('newUser_') && key.includes(userId)) {
                try {
                  const data = JSON.parse(allScriptProps[key]);
                  const timeDiff = Date.now() - parseInt(data.createdTime);
                  
                  const scriptPropsResult = {
                    method: 'ScriptProperties',
                    key: key,
                    data: data,
                    timeDiff: timeDiff,
                    found: timeDiff < 120000
                  };
                  
                  if (scriptPropsResult.found) {
                    isRecentlyCreated = true;
                    detectionSource = 'ScriptProperties';
                  }
                  detectionResults.push(scriptPropsResult);
                } catch (parseError) {
                  detectionResults.push({ method: 'ScriptProperties', key: key, parseError: parseError.message });
                }
              }
            });
          } catch (scriptPropError) {
            detectionResults.push({ method: 'ScriptProperties', error: scriptPropError.message });
          }
        }
        
        // 方法3: メールアドレスベースの検出
        if (!isRecentlyCreated) {
          try {
            const scriptProps = PropertiesService.getScriptProperties();
            const allScriptProps = scriptProps.getProperties();
            const currentEmail = activeUserEmail.toLowerCase();
            
            Object.keys(allScriptProps).forEach(key => {
              if (key.startsWith('newUser_') && key.toLowerCase().includes(currentEmail)) {
                try {
                  const data = JSON.parse(allScriptProps[key]);
                  const timeDiff = Date.now() - parseInt(data.createdTime);
                  
                  const emailBasedResult = {
                    method: 'EmailBased',
                    key: key,
                    data: data,
                    timeDiff: timeDiff,
                    found: timeDiff < 120000
                  };
                  
                  if (emailBasedResult.found) {
                    isRecentlyCreated = true;
                    detectionSource = 'EmailBased';
                  }
                  detectionResults.push(emailBasedResult);
                } catch (parseError) {
                  detectionResults.push({ method: 'EmailBased', key: key, parseError: parseError.message });
                }
              }
            });
          } catch (emailDetectionError) {
            detectionResults.push({ method: 'EmailBased', error: emailDetectionError.message });
          }
        }
        
        infoLog('verifyAdminAccess: 新規ユーザー検出結果:', {
          userId: userId,
          activeUserEmail: activeUserEmail,
          isRecentlyCreated: isRecentlyCreated,
          detectionSource: detectionSource,
          threshold: '120秒',
          detectionResults: detectionResults
        });
        
      } catch (generalError) {
        warnLog('verifyAdminAccess: 新規ユーザー検出で一般エラー:', generalError.message);
        detectionResults.push({ method: 'general', error: generalError.message });
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
        
        // まだ見つからない場合は緊急メール認証システムを実行
        if (!userFromDb) {
          warnLog('verifyAdminAccess: 🕒 リトライ後もデータなし - 緊急メール認証システムを実行');
          
          // 緊急メール認証: 複数の検証方法
          const emergencyAuth = performEmergencyEmailAuth(userId, activeUserEmail, detectionResults);
          
          if (emergencyAuth.approved) {
            infoLog('verifyAdminAccess: 🚨 緊急メール認証で承認', {
              userId: userId,
              method: emergencyAuth.method,
              reason: emergencyAuth.reason
            });
            return true; // 緊急承認
          } else {
            warnLog('verifyAdminAccess: 🚫 緊急メール認証でも承認できません', emergencyAuth);
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
      
      // 書き込み完了の検証
      debugLog('processLoginFlow: データベース書き込み完了検証を開始:', newUser.userId);
      let writeVerificationSuccess = false;
      let verificationAttempts = 0;
      const maxVerificationAttempts = 3;
      
      while (!writeVerificationSuccess && verificationAttempts < maxVerificationAttempts) {
        verificationAttempts++;
        try {
          Utilities.sleep(500 * verificationAttempts); // 段階的待機: 500ms, 1000ms, 1500ms
          
          const verificationUser = fetchUserFromDatabase('userId', newUser.userId, {
            enableDiagnostics: false,
            autoRepair: false,
            retryCount: 0
          });
          
          if (verificationUser && verificationUser.userId === newUser.userId) {
            writeVerificationSuccess = true;
            infoLog(`✅ processLoginFlow: 書き込み完了検証成功 (試行${verificationAttempts}回目)`, newUser.userId);
            break;
          } else {
            warnLog(`⚠️ processLoginFlow: 書き込み完了検証失敗 (試行${verificationAttempts}/${maxVerificationAttempts})`, {
              userId: newUser.userId,
              found: !!verificationUser,
              foundUserId: verificationUser ? verificationUser.userId : null
            });
          }
        } catch (verifyError) {
          warnLog(`❌ processLoginFlow: 書き込み完了検証でエラー (試行${verificationAttempts}/${maxVerificationAttempts}):`, verifyError.message);
        }
      }
      
      if (!writeVerificationSuccess) {
        warnLog('🚨 processLoginFlow: 書き込み完了検証が最終的に失敗しました:', {
          userId: newUser.userId,
          attempts: verificationAttempts,
          warning: 'データベース同期に遅延がある可能性があります'
        });
      }

      debugLog('processLoginFlow: New user creation completed:', newUser.userId);

      // 3e. 新規ユーザー作成の完了を記録（管理パネルアクセス時の参考用）
      try {
        const createdTime = Date.now().toString();
        const userEmail = newUser.adminEmail;
        
        // UserPropertiesとScriptPropertiesの両方に記録（フォールバック対応）
        const userProps = PropertiesService.getUserProperties();
        const scriptProps = PropertiesService.getScriptProperties();
        
        // UserPropertiesに記録
        userProps.setProperties({
          'lastCreatedUserId': newUser.userId,
          'lastCreatedUserTime': createdTime,
          'lastCreatedUserEmail': userEmail
        });
        
        // ScriptPropertiesにも記録（フォールバック用）
        const scriptKey = `newUser_${userEmail}_${newUser.userId}`;
        scriptProps.setProperty(scriptKey, JSON.stringify({
          userId: newUser.userId,
          email: userEmail,
          createdTime: createdTime,
          timestamp: new Date().toISOString()
        }));
        
        // 古い記録をクリーンアップ（1時間以上前の記録）
        try {
          const allScriptProps = scriptProps.getProperties();
          const oneHourAgo = Date.now() - (60 * 60 * 1000);
          
          Object.keys(allScriptProps).forEach(key => {
            if (key.startsWith('newUser_')) {
              try {
                const data = JSON.parse(allScriptProps[key]);
                if (parseInt(data.createdTime) < oneHourAgo) {
                  scriptProps.deleteProperty(key);
                }
              } catch (e) {
                // 無効なデータは削除
                scriptProps.deleteProperty(key);
              }
            }
          });
        } catch (cleanupError) {
          warnLog('processLoginFlow: クリーンアップでエラー:', cleanupError.message);
        }
        
        infoLog('✅ processLoginFlow: 新規ユーザー作成情報を記録しました', {
          userId: newUser.userId,
          email: userEmail,
          userProps: 'success',
          scriptProps: 'success'
        });
      } catch (propError) {
        errorLog('🚨 processLoginFlow: ユーザープロパティ記録でエラー:', propError.message);
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
