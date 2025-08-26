/**
 * サービスアカウント管理モジュール
 * Google Apps Scriptでサービスアカウント認証を行うための関数群
 */

/**
 * サービスアカウントのメールアドレスを取得
 * @returns {string} サービスアカウントのメールアドレス
 */
function getServiceAccountEmail() {
  try {
    /** @type {Object} プロパティサービス */
    const props = PropertiesService.getScriptProperties();
    
    /** @type {string|null} サービスアカウント認証情報JSON */
    const serviceAccountCredsJson = props.getProperty('SERVICE_ACCOUNT_CREDS');
    
    if (!serviceAccountCredsJson) {
      logWarn('サービスアカウント認証情報が設定されていません');
      return 'サービスアカウント未設定';
    }

    /** @type {Object} サービスアカウント認証情報 */
    const serviceAccountCreds = JSON.parse(serviceAccountCredsJson);
    
    if (!serviceAccountCreds.client_email) {
      logWarn('サービスアカウントのclient_emailが見つかりません');
      return 'サービスアカウント設定エラー';
    }

    return serviceAccountCreds.client_email;
    
  } catch (error) {
    logError(error, 'getServiceAccountEmail', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.AUTHENTICATION);
    return 'サービスアカウント設定エラー';
  }
}

/**
 * サービスアカウントのアクセストークンを生成（新規）
 * @returns {string|null} アクセストークン
 */
function generateNewServiceAccountToken() {
  try {
    /** @type {Object} プロパティサービス */
    const props = PropertiesService.getScriptProperties();
    
    /** @type {string|null} サービスアカウント認証情報JSON */
    const serviceAccountCredsJson = props.getProperty('SERVICE_ACCOUNT_CREDS');
    
    if (!serviceAccountCredsJson) {
      logWarn('サービスアカウント認証情報が設定されていません');
      return null;
    }

    /** @type {Object} サービスアカウント認証情報 */
    const serviceAccountCreds = JSON.parse(serviceAccountCredsJson);
    
    // 必要なフィールドの検証
    const requiredFields = ['client_email', 'private_key', 'token_uri'];
    for (const field of requiredFields) {
      if (!serviceAccountCreds[field]) {
        throw new Error('サービスアカウント設定に' + field + 'が不足しています');
      }
    }

    /** @type {number} 現在時刻（UNIX秒） */
    const now = Math.floor(Date.now() / 1000);
    
    /** @type {number} トークン有効期限（1時間後） */
    const exp = now + 3600;

    // JWT ヘッダー
    /** @type {Object} JWTヘッダー */
    const header = {
      alg: 'RS256',
      typ: 'JWT'
    };

    // JWT ペイロード（Google OAuth2の要求）
    /** @type {Object} JWTペイロード */
    const payload = {
      iss: serviceAccountCreds.client_email,
      scope: [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive',
        'https://www.googleapis.com/auth/forms'
      ].join(' '),
      aud: serviceAccountCreds.token_uri || 'https://oauth2.googleapis.com/token',
      iat: now,
      exp: exp
    };

    // JWT を作成
    /** @type {string} Base64エンコードされたヘッダー */
    const encodedHeader = Utilities.base64EncodeWebSafe(JSON.stringify(header)).replace(/=+$/, '');
    
    /** @type {string} Base64エンコードされたペイロード */
    const encodedPayload = Utilities.base64EncodeWebSafe(JSON.stringify(payload)).replace(/=+$/, '');
    
    /** @type {string} 署名対象文字列 */
    const signatureInput = encodedHeader + '.' + encodedPayload;
    
    // 秘密鍵で署名
    /** @type {string} クリーンアップされた秘密鍵 */
    const privateKey = serviceAccountCreds.private_key.replace(/\\n/g, '\n');
    
    /** @type {Uint8Array} 署名 */
    const signature = Utilities.computeRsaSha256Signature(signatureInput, privateKey);
    
    /** @type {string} Base64エンコードされた署名 */
    const encodedSignature = Utilities.base64EncodeWebSafe(signature).replace(/=+$/, '');
    
    /** @type {string} 完成したJWT */
    const jwt = signatureInput + '.' + encodedSignature;

    // トークンエンドポイントにリクエスト
    /** @type {Object} リクエストパラメータ */
    const tokenRequestPayload = {
      grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
      assertion: jwt
    };

    /** @type {Object} HTTPリクエストオプション */
    const options = {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded'
      },
      payload: Object.keys(tokenRequestPayload)
        .map(key => key + '=' + encodeURIComponent(tokenRequestPayload[key]))
        .join('&')
    };

    /** @type {Object} HTTPレスポンス */
    const response = UrlFetchApp.fetch(serviceAccountCreds.token_uri, options);
    
    /** @type {number} レスポンスコード */
    const responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      throw new Error('Token request failed with status ' + responseCode + ': ' + response.getContentText());
    }

    /** @type {Object} レスポンスデータ */
    const responseData = JSON.parse(response.getContentText());
    
    if (!responseData.access_token) {
      throw new Error('アクセストークンが含まれていません: ' + response.getContentText());
    }

    logInfo('サービスアカウントのアクセストークンを生成しました');
    
    return responseData.access_token;
    
  } catch (error) {
    logError(error, 'generateNewServiceAccountToken', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.AUTHENTICATION);
    return null;
  }
}

/**
 * サービスアカウントの設定状況を確認
 * @returns {Object} 設定状況
 */
function checkServiceAccountConfiguration() {
  try {
    /** @type {Object} プロパティサービス */
    const props = PropertiesService.getScriptProperties();
    
    /** @type {string|null} サービスアカウント認証情報JSON */
    const serviceAccountCredsJson = props.getProperty('SERVICE_ACCOUNT_CREDS');
    
    /** @type {Object} 結果 */
    const result = {
      configured: false,
      email: null,
      hasRequiredFields: false,
      errors: []
    };

    if (!serviceAccountCredsJson) {
      result.errors.push('サービスアカウント認証情報が設定されていません');
      return result;
    }

    try {
      /** @type {Object} サービスアカウント認証情報 */
      const serviceAccountCreds = JSON.parse(serviceAccountCredsJson);
      
      result.email = serviceAccountCreds.client_email || null;
      result.configured = true;

      // 必要なフィールドをチェック
      const requiredFields = [
        'type', 'project_id', 'private_key_id', 'private_key', 
        'client_email', 'client_id', 'auth_uri', 'token_uri'
      ];

      const missingFields = [];
      for (const field of requiredFields) {
        if (!serviceAccountCreds[field]) {
          missingFields.push(field);
        }
      }

      if (missingFields.length === 0) {
        result.hasRequiredFields = true;
      } else {
        result.errors.push('必要なフィールドが不足: ' + missingFields.join(', '));
      }

      // タイプチェック
      if (serviceAccountCreds.type !== 'service_account') {
        result.errors.push('認証情報のタイプが不正: ' + serviceAccountCreds.type + ' (期待値: service_account)');
      }

    } catch (parseError) {
      result.errors.push('サービスアカウント認証情報のJSON解析に失敗: ' + parseError.message);
    }

    return result;
    
  } catch (error) {
    logError(error, 'checkServiceAccountConfiguration', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.CONFIG);
    return {
      configured: false,
      email: null,
      hasRequiredFields: false,
      errors: ['設定確認中にエラーが発生: ' + error.message]
    };
  }
}

/**
 * サービスアカウント設定をセットアップ
 * @param {string} serviceAccountJson サービスアカウントのJSONキー
 * @returns {Object} 設定結果
 */
function setupServiceAccount(serviceAccountJson) {
  try {
    if (!serviceAccountJson || typeof serviceAccountJson !== 'string') {
      throw new Error('サービスアカウントJSONが提供されていません');
    }

    // JSON検証
    /** @type {Object} サービスアカウント認証情報 */
    let serviceAccountCreds;
    try {
      serviceAccountCreds = JSON.parse(serviceAccountJson);
    } catch (parseError) {
      throw new Error('サービスアカウントJSONの解析に失敗: ' + parseError.message);
    }

    // タイプ検証
    if (serviceAccountCreds.type !== 'service_account') {
      throw new Error('認証情報のタイプが不正: ' + serviceAccountCreds.type + ' (期待値: service_account)');
    }

    // 必要フィールド検証
    const requiredFields = [
      'type', 'project_id', 'private_key_id', 'private_key', 
      'client_email', 'client_id', 'auth_uri', 'token_uri'
    ];

    const missingFields = [];
    for (const field of requiredFields) {
      if (!serviceAccountCreds[field]) {
        missingFields.push(field);
      }
    }

    if (missingFields.length > 0) {
      throw new Error('必要なフィールドが不足: ' + missingFields.join(', '));
    }

    // プロパティに保存
    /** @type {Object} プロパティサービス */
    const props = PropertiesService.getScriptProperties();
    props.setProperty('SERVICE_ACCOUNT_CREDS', serviceAccountJson);

    // テスト用にトークン生成を試行
    /** @type {string|null} テストトークン */
    const testToken = generateNewServiceAccountToken();
    
    if (!testToken) {
      throw new Error('サービスアカウント設定後のトークン生成に失敗');
    }

    logInfo('サービスアカウント設定が完了しました', { email: serviceAccountCreds.client_email });

    return {
      success: true,
      email: serviceAccountCreds.client_email,
      projectId: serviceAccountCreds.project_id,
      message: 'サービスアカウント設定が完了しました'
    };

  } catch (error) {
    logError(error, 'setupServiceAccount', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.CONFIG);
    
    return {
      success: false,
      error: error.message,
      message: 'サービスアカウント設定に失敗しました'
    };
  }
}

/**
 * サービスアカウント設定をクリア
 * @returns {Object} クリア結果
 */
function clearServiceAccountConfiguration() {
  try {
    /** @type {Object} プロパティサービス */
    const props = PropertiesService.getScriptProperties();
    props.deleteProperty('SERVICE_ACCOUNT_CREDS');

    // キャッシュもクリア
    if (typeof cacheManager !== 'undefined' && cacheManager.remove) {
      cacheManager.remove('service_account_token');
    }

    logInfo('サービスアカウント設定をクリアしました');
    
    return {
      success: true,
      message: 'サービスアカウント設定をクリアしました'
    };
    
  } catch (error) {
    logError(error, 'clearServiceAccountConfiguration', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.CONFIG);
    
    return {
      success: false,
      error: error.message,
      message: 'サービスアカウント設定のクリアに失敗しました'
    };
  }
}

/**
 * サービスアカウントのヘルスチェック
 * @returns {Object} ヘルスチェック結果
 */
function performServiceAccountHealthCheck() {
  try {
    /** @type {Object} 設定確認結果 */
    const configCheck = checkServiceAccountConfiguration();
    
    /** @type {Object} 結果 */
    const result = {
      timestamp: new Date().toISOString(),
      configured: configCheck.configured,
      email: configCheck.email,
      hasRequiredFields: configCheck.hasRequiredFields,
      tokenGeneration: false,
      overallStatus: 'unknown',
      details: {
        configurationErrors: configCheck.errors,
        tokenErrors: []
      }
    };

    // 設定が正常な場合のみトークン生成テスト
    if (configCheck.configured && configCheck.hasRequiredFields) {
      try {
        /** @type {string|null} テストトークン */
        const testToken = generateNewServiceAccountToken();
        result.tokenGeneration = !!testToken;
        
        if (!testToken) {
          result.details.tokenErrors.push('トークン生成に失敗しました');
        }
      } catch (tokenError) {
        result.details.tokenErrors.push('トークン生成エラー: ' + tokenError.message);
      }
    } else {
      result.details.tokenErrors.push('設定が不完全なためトークン生成をスキップ');
    }

    // 総合ステータス判定
    if (result.configured && result.hasRequiredFields && result.tokenGeneration) {
      result.overallStatus = 'healthy';
    } else if (result.configured) {
      result.overallStatus = 'configured_but_not_functional';
    } else {
      result.overallStatus = 'not_configured';
    }

    return result;
    
  } catch (error) {
    logError(error, 'performServiceAccountHealthCheck', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.SYSTEM);
    
    return {
      timestamp: new Date().toISOString(),
      configured: false,
      email: null,
      hasRequiredFields: false,
      tokenGeneration: false,
      overallStatus: 'error',
      details: {
        configurationErrors: [],
        tokenErrors: ['ヘルスチェック実行エラー: ' + error.message]
      }
    };
  }
}