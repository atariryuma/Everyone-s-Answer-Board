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
 * 合理的な新規ユーザー認証処理 - 適切なバランスでセキュリティと利便性を両立
 * @param {string} userId - 検証するユーザーID
 * @param {string} activeUserEmail - 現在のログインユーザーのメールアドレス
 * @param {Object} searchSummary - 基本検索の結果サマリー
 * @returns {Object} 認証結果 {approved: boolean, method: string, reason: string}
 */
function handleNewUserAuthentication(userId, activeUserEmail, searchSummary) {
  const authResult = {
    approved: false,
    method: 'none',
    reason: 'no_new_user_criteria_met',
    details: {}
  };

  try {
    const currentEmail = activeUserEmail ? activeUserEmail.toLowerCase().trim() : '';
    const emailDomain = currentEmail.split('@')[1];

    // 新規ユーザーかどうかをScriptPropertiesで確認（UserPropertiesは使用しない）
    let isRecentlyCreated = false;
    let createdTimeInfo = null;

    try {
      const scriptProps = PropertiesService.getScriptProperties();
      const allProps = scriptProps.getProperties();
      
      // デバッグログ: ScriptPropertiesの内容と検索パラメーターを記録
      const newUserKeys = Object.keys(allProps).filter(k => k.startsWith('newUser_'));
      debugLog('handleNewUserAuthentication: 検索パラメーター:', {
        userId: userId,
        currentEmail: currentEmail,
        totalProps: Object.keys(allProps).length,
        newUserKeys: newUserKeys.length,
        availableKeys: newUserKeys.slice(0, 5) // 最初の5件のみ表示
      });
      
      // 新規ユーザー記録を探す - より柔軟な検索ロジック
      for (const [key, value] of Object.entries(allProps)) {
        if (key.startsWith('newUser_')) {
          try {
            const data = JSON.parse(value);
            const timeDiff = Date.now() - parseInt(data.createdTime);
            
            // 5分以内に作成されたユーザーのみ対象
            if (timeDiff < 300000) { // 5分 = 300秒
              // より柔軟なマッチングロジック
              const keyLower = key.toLowerCase();
              const emailLower = currentEmail.toLowerCase();
              const dataEmailLower = data.email ? data.email.toLowerCase() : '';
              
              const matchConditions = {
                userIdExact: data.userId === userId,
                userIdInKey: key.includes(userId),
                emailInKey: keyLower.includes(emailLower),
                emailExact: dataEmailLower === emailLower,
                emailLocal: emailLower.split('@')[0] && keyLower.includes(emailLower.split('@')[0])
              };
              
              debugLog('handleNewUserAuthentication: キーマッチング検証:', {
                key: key,
                timeDiff: timeDiff,
                matchConditions: matchConditions,
                data: data
              });
              
              // いずれかの条件に一致すれば有効とみなす
              if (matchConditions.userIdExact || 
                  matchConditions.userIdInKey || 
                  matchConditions.emailInKey || 
                  matchConditions.emailExact || 
                  matchConditions.emailLocal) {
                
                isRecentlyCreated = true;
                createdTimeInfo = {
                  timeDiff: timeDiff,
                  key: key,
                  data: data,
                  matchedBy: Object.keys(matchConditions).filter(k => matchConditions[k])
                };
                
                infoLog('handleNewUserAuthentication: ✅ 新規ユーザーマッチング成功:', {
                  matchedBy: createdTimeInfo.matchedBy,
                  timeDiff: timeDiff + 'ms'
                });
                break;
              }
            } else {
              debugLog('handleNewUserAuthentication: タイムアウト:', {
                key: key,
                timeDiff: timeDiff,
                timeoutThreshold: 300000
              });
            }
          } catch (parseError) {
            warnLog('handleNewUserAuthentication: JSONパースエラー:', {
              key: key,
              error: parseError.message
            });
            continue;
          }
        }
      }
      
      // 検索結果のサマリーログ
      if (!isRecentlyCreated && newUserKeys.length > 0) {
        warnLog('handleNewUserAuthentication: 新規ユーザー検索で一致なし:', {
          searchedUserId: userId,
          searchedEmail: currentEmail,
          availableNewUserKeys: newUserKeys,
          totalNewUserRecords: newUserKeys.length
        });
      }
    } catch (propsError) {
      errorLog('handleNewUserAuthentication: ScriptProperties取得でエラー:', propsError.message);
    }

    if (isRecentlyCreated && createdTimeInfo) {
      warnLog('handleNewUserAuthentication: 新規作成ユーザーを検出、追加検証を実行');
      
      // 2-3秒待機してから追加検証
      Utilities.sleep(2500); // 2.5秒待機
      
      // 追加検証: データベース同期を待ってもう一度検索
      let retryUserFromDb = null;
      try {
        retryUserFromDb = fetchUserFromDatabase('userId', userId, {
          enableDiagnostics: false,
          autoRepair: false,
          retryCount: 1
        });
      } catch (retryError) {
        warnLog('handleNewUserAuthentication: 追加検索でエラー:', retryError.message);
      }

      if (retryUserFromDb) {
        authResult.approved = true;
        authResult.method = 'delayed_database_sync';
        authResult.reason = 'new_user_found_after_sync_wait';
        authResult.details = {
          waitTime: '2.5秒',
          createdTimeDiff: createdTimeInfo.timeDiff + 'ms',
          domain: emailDomain
        };
        return authResult;
      }

      // 教育機関ドメインの場合のみ、合理的な緊急措置
      if (emailDomain === 'naha-okinawa.ed.jp') {
        warnLog('handleNewUserAuthentication: 教育機関ドメインによる合理的緊急措置を実行');
        authResult.approved = true;
        authResult.method = 'educational_emergency_measure';
        authResult.reason = 'educational_domain_with_recent_creation';
        authResult.details = {
          domain: emailDomain,
          createdTimeDiff: createdTimeInfo.timeDiff + 'ms',
          securityNote: '5分以内の新規作成 + 教育機関ドメインによる限定的承認'
        };
        return authResult;
      }
    }

    authResult.reason = 'no_recent_creation_or_invalid_domain';
    authResult.details = {
      domain: emailDomain,
      isRecentlyCreated: isRecentlyCreated,
      createdTimeInfo: createdTimeInfo
    };

  } catch (error) {
    authResult.reason = 'new_user_auth_error';
    authResult.details = { error: error.message };
    errorLog('handleNewUserAuthentication: エラー:', error.message);
  }

  return authResult;
}

// 不要な複雑な認証システムを削除（簡素化のため）

/**
 * 指定されたユーザーの管理者権限を検証する - 安定版ベース + 合理的な新規ユーザー対応
 * @param {string} userId - 検証するユーザーのID
 * @returns {boolean} 管理者権限がある場合は true、そうでなければ false
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

    // 3段階でのユーザー情報取得（安定版と同じロジック）
    var userFromDb = null;
    var searchAttempts = [];

    // 第1段階: キャッシュからユーザー情報を取得を試行
    try {
      userFromDb = getOrFetchUserInfo(userId, 'userId', {
        useExecutionCache: false, // セキュリティ認証のため実行時キャッシュは使用しない
        ttl: 30 // より短いTTLで最新性を確保
      });
      searchAttempts.push({ method: 'getOrFetchUserInfo', success: !!userFromDb });
    } catch (error) {
      warnLog('verifyAdminAccess: getOrFetchUserInfo でエラー:', error.message);
      searchAttempts.push({ method: 'getOrFetchUserInfo', error: error.message });
    }

    // 第2段階: 直接データベース検索にフォールバック
    if (!userFromDb || !userFromDb.adminEmail) {
      debugLog('verifyAdminAccess: 直接データベース検索にフォールバック');
      try {
        userFromDb = fetchUserFromDatabase('userId', userId);
        searchAttempts.push({ method: 'fetchUserFromDatabase', success: !!userFromDb });
      } catch (error) {
        errorLog('verifyAdminAccess: fetchUserFromDatabase でエラー:', error.message);
        searchAttempts.push({ method: 'fetchUserFromDatabase', error: error.message });
      }
    }

    // 第3段階: findUserById による追加検証
    if (!userFromDb) {
      debugLog('verifyAdminAccess: findUserById による最終検証を実行');
      try {
        userFromDb = findUserById(userId, { useExecutionCache: false, forceRefresh: true });
        searchAttempts.push({ method: 'findUserById', success: !!userFromDb });
      } catch (error) {
        errorLog('verifyAdminAccess: findUserById でエラー:', error.message);
        searchAttempts.push({ method: 'findUserById', error: error.message });
      }
    }

    // 検索結果の詳細ログ
    const searchSummary = {
      found: !!userFromDb,
      userId: userFromDb ? userFromDb.userId : 'なし',
      adminEmail: userFromDb ? userFromDb.adminEmail : 'なし',
      isActive: userFromDb ? userFromDb.isActive : 'なし',
      activeUserEmail: activeUserEmail,
      searchAttempts: searchAttempts,
      timestamp: new Date().toISOString()
    };
    debugLog('verifyAdminAccess: ユーザー検索結果:', searchSummary);

    // 新規ユーザー対応: 合理的な追加検証
    if (!userFromDb) {
      const newUserAuth = handleNewUserAuthentication(userId, activeUserEmail, searchSummary);
      if (newUserAuth.approved) {
        warnLog('verifyAdminAccess: 新規ユーザー認証で承認:', newUserAuth);
        return true;
      } else {
        errorLog('verifyAdminAccess: 🚨 全ての検索方法でユーザーが見つかりませんでした:', {
          requestedUserId: userId,
          activeUserEmail: activeUserEmail,
          searchSummary: searchSummary,
          newUserAuthResult: newUserAuth
        });
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
