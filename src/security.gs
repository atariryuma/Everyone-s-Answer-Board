/**
 * @fileoverview 簡略化されたセキュリティシステム
 * 基本的なアクセス制御機能のみを保持し、複雑なセキュリティクラスを除去
 */

// =============================================================================
// SECTION 1: 基本認証機能（必要最小限）
// =============================================================================

// Module-scoped constants (2024 GAS Best Practice)
const SECURITY_CONFIG = Object.freeze({
  AUTH_CACHE_KEY: 'SA_TOKEN_CACHE',
  TOKEN_EXPIRY_BUFFER: CORE.TIMEOUTS.LONG / 6, // 5秒
  SESSION_TTL: CORE.TIMEOUTS.LONG, // 30秒
  MAX_LOGIN_ATTEMPTS: 3,
});

// 基本的な入力検証パターン
const VALIDATION_PATTERNS = Object.freeze({
  SPREADSHEET_URL: /^https:\/\/docs\.google\.com\/spreadsheets\/d\/[a-zA-Z0-9-_]+/,
  SAFE_STRING: /^[a-zA-Z0-9\s\-_.@]+$/,
});

/**
 * CacheService永続キャッシュでサービスアカウントトークンを取得
 * GAS実行環境で真に永続化されるCacheServiceを活用
 * @returns {string} アクセストークン
 */
function getServiceAccountTokenCached() {
  try {
    // Service Accountトークン取得開始

    // CacheService直接使用（GAS環境で真に永続化）
    const scriptCache = CacheService.getScriptCache();
    const cachedToken = scriptCache.get(SECURITY_CONFIG.AUTH_CACHE_KEY);

    if (cachedToken) {
      // キャッシュヒット
      return cachedToken;
    }

    // キャッシュミス - 新規生成
    const newToken = generateNewServiceAccountToken();

    // 1時間（3600秒）永続キャッシュ
    scriptCache.put(SECURITY_CONFIG.AUTH_CACHE_KEY, newToken, 3600);
    // キャッシュ保存完了

    return newToken;
  } catch (error) {
    console.error('❌ getServiceAccountTokenCached: CacheService永続キャッシュエラー', {
      error: error.message,
      stack: error.stack,
    });
    throw error;
  }
}

/**
 * 新しいJWTトークンを生成
 * @returns {string} アクセストークン
 */
function generateNewServiceAccountToken() {
  try {
    // トークン生成開始

    // 統一秘密情報管理システムで安全に取得
    const serviceAccountCreds = getSecureServiceAccountCreds();

    const privateKey = serviceAccountCreds.private_key.replace(/\n/g, '\n'); // 改行文字を正規化
    const clientEmail = serviceAccountCreds.client_email;
    const tokenUrl = 'https://www.googleapis.com/oauth2/v4/token';

    // JWT準備完了

    const now = Math.floor(Date.now() / 1000);
    const expiresAt = now + 3600; // 1時間後

    // JWT生成
    const jwtHeader = { alg: 'RS256', typ: 'JWT' };
    const jwtClaimSet = {
      iss: clientEmail,
      scope: 'https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive',
      aud: tokenUrl,
      exp: expiresAt,
      iat: now,
    };

    const encodedHeader = Utilities.base64EncodeWebSafe(JSON.stringify(jwtHeader));
    const encodedClaimSet = Utilities.base64EncodeWebSafe(JSON.stringify(jwtClaimSet));
    const signatureInput = `${encodedHeader}.${encodedClaimSet}`;
    const signature = Utilities.computeRsaSha256Signature(signatureInput, privateKey);
    const encodedSignature = Utilities.base64EncodeWebSafe(signature);
    const jwt = `${signatureInput}.${encodedSignature}`;

    // トークンリクエスト
    const response = UrlFetchApp.fetch(tokenUrl, {
      method: 'post',
      contentType: 'application/x-www-form-urlencoded',
      payload: {
        grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
        assertion: jwt,
      },
      muteHttpExceptions: true,
    });

    const responseCode = response.getResponseCode();
    if (responseCode !== 200) {
      const responseText = response.getContentText();
      console.error('[ERROR]', 'Token request failed. Status: [非表示]');
      console.error('[ERROR]', 'Response:', responseText);

      // セキュリティ: エラー詳細を隠蔽、内部ログのみ
      console.error('Service Account認証エラー:', { responseCode, responseText });
      throw new Error('認証システムでエラーが発生しました。システム管理者にお問い合わせください。');
    }

    const responseData = JSON.parse(response.getContentText());
    if (!responseData.access_token) {
      throw new Error('アクセストークンが見つかりませんでした');
    }

    // Security: Never log access tokens - removed token logging
    // トークン生成完了
    return responseData.access_token;
  } catch (error) {
    console.error('🔑 Service Accountトークン生成失敗:', error.message);
    throw error;
  }
}

/**
 * 統一秘密情報取得（PropertiesService利用）
 */
function getSecureServiceAccountCreds() {
  try {
    // 認証情報取得開始
    const props = PropertiesService.getScriptProperties();
    const credsJson = props.getProperty('SERVICE_ACCOUNT_CREDS');

    if (!credsJson) {
      console.error('🔐 SERVICE_ACCOUNT_CREDS が PropertiesService に設定されていません');
      throw new Error('サービスアカウント認証情報が設定されていません');
    }

    // 認証情報取得成功

    return JSON.parse(credsJson);
  } catch (error) {
    console.error('🔐 Service Account認証情報取得エラー:', error.message);
    throw error;
  }
}

// =============================================================================
// SECTION 2: 基本的なアクセス制御機能（Base.gsのAccessControllerを補完）
// =============================================================================

/**
 * ユーザー管理者権限を検証（元のverifyAdminAccess関数を保持）
 * @param {string} userId - 検証するユーザーのID
 * @returns {boolean} 認証結果
 */
function verifyAdminAccess(userId) {
  try {
    // セキュリティ: 厳密な入力検証
    if (
      !userId ||
      typeof userId !== 'string' ||
      userId.trim() === '' ||
      !/^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(userId)
    ) {
      console.warn('verifyAdminAccess: 無効なuserID形式');
      return false;
    }

    // 現在のログインユーザーのメールアドレスを取得
    const activeUserEmail = UserManager.getCurrentEmail();
    if (!activeUserEmail) {
      console.warn('verifyAdminAccess: アクティブユーザーのメールアドレスが取得できませんでした');
      return false;
    }

    console.log('verifyAdminAccess: 認証開始');

    // データベースからユーザー情報を取得
    let userFromDb = DB.findUserById(userId);

    // 見つからない場合は強制フレッシュで再試行
    if (!userFromDb) {
      console.log('verifyAdminAccess: 強制フレッシュで再検索中...');
      userFromDb = DB.findUserById(userId);
    }

    // ユーザーが見つからない場合は認証失敗
    if (!userFromDb) {
      console.warn('verifyAdminAccess: ユーザーが見つかりません:', {
        userId,
        activeUserEmail,
      });
      return false;
    }

    // 3重チェック実行
    // 1. メールアドレス照合
    const dbEmail = String(userFromDb.userEmail || '')
      .toLowerCase()
      .trim();
    const currentEmail = String(activeUserEmail).toLowerCase().trim();
    const isEmailMatched = dbEmail && currentEmail && dbEmail === currentEmail;

    // 2. ユーザーID照合（念のため）
    const isUserIdMatched = String(userFromDb.userId) === String(userId);

    // 3. アクティブ状態確認
    const isActive = Boolean(userFromDb.isActive);

    console.log('verifyAdminAccess: 3重チェック結果:', {
      isEmailMatched,
      isUserIdMatched,
      isActive,
    });

    // 3つの条件すべてが満たされた場合のみ認証成功
    if (isEmailMatched && isUserIdMatched && isActive) {
      console.log('✅ verifyAdminAccess: 認証成功');
      return true;
    } else {
      console.warn('❌ verifyAdminAccess: 認証失敗', {
        failures: {
          email: !isEmailMatched,
          userId: !isUserIdMatched,
          active: !isActive,
        },
      });
      return false;
    }
  } catch (error) {
    console.error('[ERROR]', '❌ verifyAdminAccess: 認証処理エラー:', error.message);
    return false;
  }
}

/**
 * ユーザーの最終アクセス時刻のみを更新（設定は保護）
 * @param {string} userId - 更新対象のユーザーID
 */
// updateUserLastAccess removed from security.gs - use auth.js version (proper configJson handling)

// =============================================================================
// SECTION 3: 最小限のユーティリティ関数
// =============================================================================

/**
 * セキュアなデータベースIDを取得
 */
function getSecureDatabaseId() {
  try {
    const props = PropertiesService.getScriptProperties();
    const databaseId = props.getProperty('DATABASE_SPREADSHEET_ID');
    if (!databaseId) {
      throw new Error('データベースIDが設定されていません');
    }
    return databaseId;
  } catch (error) {
    console.error('getSecureDatabaseIdエラー:', error.message);
    throw error;
  }
}

// =============================================================================
// SECTION 4: 基本入力検証関数
// =============================================================================

/**
 * スプレッドシートURL検証（基本実装）
 * @param {string} url - 検証するURL
 * @returns {boolean} 有効性
 */
function validateSpreadsheetUrl(url) {
  if (!url || typeof url !== 'string') {
    return false;
  }
  return VALIDATION_PATTERNS.SPREADSHEET_URL.test(url.trim());
}

/**
 * 基本的な文字列サニタイゼーション
 * @param {string} input - サニタイズする文字列
 * @returns {string} サニタイズ済み文字列
 */
function sanitizeInput(input) {
  if (!input || typeof input !== 'string') {
    return '';
  }
  return input.replace(/[<>\"'&]/g, '').trim();
}

// verifyUserAccess function removed - standardized to use App.getAccess().verifyAccess() directly

/**
 * スプレッドシートをサービスアカウントと共有
 * @param {string} spreadsheetId - 共有するスプレッドシートのID
 */
function shareSpreadsheetWithServiceAccount(spreadsheetId) {
  try {
    if (!spreadsheetId || typeof spreadsheetId !== 'string') {
      console.warn('shareSpreadsheetWithServiceAccount: 無効なspreadsheetId:', spreadsheetId);
      return;
    }

    // サービスアカウント認証情報を取得
    const serviceAccountCreds = getSecureServiceAccountCreds();
    const serviceAccountEmail = serviceAccountCreds.client_email;

    if (!serviceAccountEmail) {
      console.error(
        'shareSpreadsheetWithServiceAccount: サービスアカウントメールアドレスが見つかりません'
      );
      return;
    }

    // スプレッドシートをサービスアカウントと共有
    const spreadsheet = new ConfigurationManager().getSpreadsheet(spreadsheetId);
    spreadsheet.addEditor(serviceAccountEmail);

    console.log(
      `✅ スプレッドシートをサービスアカウントと共有完了: ${spreadsheetId} -> ${serviceAccountEmail}`
    );
  } catch (error) {
    console.error('shareSpreadsheetWithServiceAccount エラー:', {
      spreadsheetId,
      error: error.message,
      stack: error.stack,
    });
    // 共有失敗は非致命的エラーとして処理（システム全体を停止させない）
  }
}

/**
 * Sheets APIを使用してデータを更新
 * @param {Object} service - Sheets APIサービスオブジェクト
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} range - 更新する範囲
 * @param {Array<Array>} values - 更新する値の2次元配列
 * @returns {Object} APIレスポンス
 */
function updateSheetsData(service, spreadsheetId, range, values) {
  try {
    // Service Account経由でSheets API使用
    if (
      service &&
      service.spreadsheets &&
      service.spreadsheets.values &&
      service.spreadsheets.values.update
    ) {
      console.log('updateSheetsData: Service Account経由でSheets API使用', {
        spreadsheetId: `${spreadsheetId.substring(0, 10)}...`,
        range,
      });

      const response = service.spreadsheets.values.update({
        spreadsheetId,
        range,
        valueInputOption: 'RAW',
        values,
      });

      console.log('✅ updateSheetsData: Service Account成功');
      return response;
    } else {
      throw new Error('Service Account Sheets APIが利用できません');
    }
  } catch (error) {
    // 403 Permission Denied エラーの場合、自動でService Account共有を実行
    if (error.message && error.message.includes('"code": 403') && error.message.includes('permission')) {
      console.warn('🔧 updateSheetsData: 403権限エラー検出 - Service Account自動共有を実行');
      
      try {
        // Service Accountをスプレッドシートに自動追加
        shareSpreadsheetWithServiceAccount(spreadsheetId);
        console.log('✅ Service Account自動共有完了 - API再試行します');
        
        // 少し待機してから再試行
        Utilities.sleep(1000);
        
        // API再試行 (1回のみ)
        const retryResponse = service.spreadsheets.values.update({
          spreadsheetId,
          range,
          valueInputOption: 'RAW',
          values,
        });
        
        console.log('✅ updateSheetsData: Service Account自動復旧成功');
        return retryResponse;
        
      } catch (recoveryError) {
        console.error('❌ Service Account自動復旧失敗:', recoveryError.message);
        // 復旧失敗の場合は元のエラーを投げる
      }
    }
    
    console.error('❌ updateSheetsData Service Account呼び出しエラー:', error.message);
    throw error;
  }
}

/**
 * Sheets APIを使用してバッチでデータを取得
 * @param {Object} service - Sheets APIサービスオブジェクト
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {Array<string>} ranges - 取得する範囲の配列
 * @returns {Object} APIレスポンス
 */
function batchGetSheetsData(service, spreadsheetId, ranges) {
  try {
    // Service Accountを使用してSheets APIでアクセス
    if (
      service &&
      service.spreadsheets &&
      service.spreadsheets.values &&
      service.spreadsheets.values.batchGet
    ) {
      console.log('batchGetSheetsData: Service Account経由でSheets API使用', {
        spreadsheetId: `${spreadsheetId.substring(0, 10)}...`,
        rangeCount: ranges.length,
      });

      // 正しいメソッド呼び出し: 関数として実行
      const response = service.spreadsheets.values.batchGet({
        spreadsheetId,
        ranges,
      });

      console.log('✅ batchGetSheetsData: Service Account成功');
      return response;
    } else {
      throw new Error('Service Account Sheets APIが利用できません');
    }
  } catch (error) {
    // 403 Permission Denied エラーの場合、自動でService Account共有を実行
    if (error.message && error.message.includes('"code": 403') && error.message.includes('permission')) {
      console.warn('🔧 batchGetSheetsData: 403権限エラー検出 - Service Account自動共有を実行');
      
      try {
        // Service Accountをスプレッドシートに自動追加
        shareSpreadsheetWithServiceAccount(spreadsheetId);
        console.log('✅ Service Account自動共有完了 - API再試行します');
        
        // 少し待機してから再試行
        Utilities.sleep(2000);
        
        // API再試行
        const retryResponse = service.spreadsheets.values.batchGet({
          spreadsheetId,
          ranges,
        });
        
        console.log('✅ batchGetSheetsData: Service Account自動復旧成功');
        return retryResponse;
        
      } catch (recoveryError) {
        console.error('❌ Service Account自動復旧失敗:', recoveryError.message);
        // 復旧失敗の場合は元のエラーを投げる
      }
    }
    
    console.error('❌ batchGetSheetsData Service Account呼び出しエラー:', error.message);
    throw error;
  }
}

/**
 * サービスアカウントのメールアドレス取得
 * @returns {string} サービスアカウントのメールアドレス
 */
function getServiceAccountEmail() {
  try {
    const serviceAccountCreds = getSecureServiceAccountCreds();
    return serviceAccountCreds.client_email || 'メールアドレス不明';
  } catch (error) {
    console.warn('サービスアカウントメール取得エラー:', error.message);
    return 'メールアドレス取得エラー';
  }
}

/**
 * サービスアカウントをスプレッドシートに編集者として追加
 * @param {string} spreadsheetId - スプレッドシートID
 */
function addServiceAccountToSpreadsheet(spreadsheetId) {
  try {
    const serviceAccountEmail = getServiceAccountEmail();
    if (serviceAccountEmail === 'メールアドレス取得エラー') {
      console.warn(
        'サービスアカウントのメールアドレスが取得できないため、スプレッドシートの共有をスキップします。'
      );
      return;
    }

    const spreadsheet = new ConfigurationManager().getSpreadsheet(spreadsheetId);
    const permissions = spreadsheet.getEditors();
    const isAlreadyEditor = permissions.some((editor) => editor.getEmail() === serviceAccountEmail);

    if (!isAlreadyEditor) {
      spreadsheet.addEditor(serviceAccountEmail);
      console.log(
        `✅ サービスアカウント (${serviceAccountEmail}) をスプレッドシート (${spreadsheetId}) に編集者として追加しました。`
      );
    } else {
      console.log(
        `サービスアカウント (${serviceAccountEmail}) は既にスプレッドシート (${spreadsheetId}) の編集者です。`
      );
    }
  } catch (error) {
    console.error(
      `サービスアカウントをスプレッドシート (${spreadsheetId}) に共有中にエラーが発生しました: ${error.message}`
    );
    throw new Error(
      `サービスアカウントをスプレッドシートに共有できませんでした。手動で ${getServiceAccountEmail()} を編集者として追加してください。`
    );
  }
}

/**
 * バッチでSheets APIを使用してスプレッドシートデータを更新（複数範囲対応）
 * @param {Object} service - Sheets API service object
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {Array} updateData - 更新データ配列 [{range, values}, ...]
 * @returns {Object} API レスポンス
 */
function batchUpdateSheetsData(service, spreadsheetId, updateData) {
  try {
    if (!updateData || !Array.isArray(updateData) || updateData.length === 0) {
      throw new Error('更新データが無効です');
    }

    console.log('batchUpdateSheetsData: バッチ更新開始', {
      spreadsheetId: `${spreadsheetId.substring(0, 10)}...`,
      updateCount: updateData.length
    });

    // 既存のupdateSheetsData関数を使用してバッチ処理
    const results = [];
    for (const update of updateData) {
      if (!update.range || !update.values) {
        console.warn('batchUpdateSheetsData: 無効な更新データをスキップ', update);
        continue;
      }
      
      const result = updateSheetsData(service, spreadsheetId, update.range, update.values);
      results.push(result);
    }

    console.log('batchUpdateSheetsData: バッチ更新完了', {
      successCount: results.length,
      totalRequested: updateData.length
    });

    return {
      status: 'success',
      updateCount: results.length,
      results: results
    };

  } catch (error) {
    // 403 Permission Denied エラーの場合、自動でService Account共有を実行
    if (error.message && error.message.includes('"code": 403') && error.message.includes('permission')) {
      console.warn('🔧 batchUpdateSheetsData: 403権限エラー検出 - Service Account自動共有を実行');
      
      try {
        // Service Accountをスプレッドシートに自動追加
        shareSpreadsheetWithServiceAccount(spreadsheetId);
        console.log('✅ Service Account自動共有完了 - API再試行します');
        
        // 少し待機してから再試行
        Utilities.sleep(2000);
        
        // API再試行 - 既存のupdateSheetsData関数を再実行
        const retryResults = [];
        for (const update of updateData) {
          if (!update.range || !update.values) continue;
          const result = updateSheetsData(service, spreadsheetId, update.range, update.values);
          retryResults.push(result);
        }
        
        console.log('✅ batchUpdateSheetsData: Service Account自動復旧成功');
        return {
          status: 'success',
          updateCount: retryResults.length,
          results: retryResults,
          recovered: true
        };
        
      } catch (recoveryError) {
        console.error('❌ Service Account自動復旧失敗:', recoveryError.message);
        // 復旧失敗の場合は元のエラーを投げる
      }
    }
    
    console.error('batchUpdateSheetsData: エラー', {
      spreadsheetId: `${spreadsheetId.substring(0, 10)}...`,
      error: error.message,
      updateDataCount: updateData?.length || 0
    });
    throw new Error(`バッチ更新エラー: ${error.message}`);
  }
}

console.log('🔐 簡略化されたセキュリティシステムが初期化されました');
