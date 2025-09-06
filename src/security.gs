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

/**
 * キャッシュされたサービスアカウントトークンを取得
 * @returns {string} アクセストークン
 */
function getServiceAccountTokenCached() {
  return cacheManager.get(SECURITY_CONFIG.AUTH_CACHE_KEY, generateNewServiceAccountToken, {
    ttl: 3500,
    enableMemoization: true,
  }); // メモ化対応でトークン取得を高速化
}

/**
 * 新しいJWTトークンを生成
 * @returns {string} アクセストークン
 */
function generateNewServiceAccountToken() {
  // 統一秘密情報管理システムで安全に取得
  const serviceAccountCreds = getSecureServiceAccountCreds();

  const privateKey = serviceAccountCreds.private_key.replace(/\n/g, '\n'); // 改行文字を正規化
  const clientEmail = serviceAccountCreds.client_email;
  const tokenUrl = 'https://www.googleapis.com/oauth2/v4/token';

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
  const signatureInput = `${encodedHeader  }.${  encodedClaimSet}`;
  const signature = Utilities.computeRsaSha256Signature(signatureInput, privateKey);
  const encodedSignature = Utilities.base64EncodeWebSafe(signature);
  const jwt = `${signatureInput  }.${  encodedSignature}`;

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
    console.error('[ERROR]', 'Token request failed. Status:', responseCode);
    console.error('[ERROR]', 'Response:', responseText);

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

  const responseData = JSON.parse(response.getContentText());
  if (!responseData.access_token) {
    throw new Error('アクセストークンが見つかりませんでした');
  }

  console.log(
    `🔐 新しいサービスアカウントトークンを生成しました (有効期限: ${new Date(expiresAt * 1000).toISOString()})`
  );
  return responseData.access_token;
}

/**
 * 統一秘密情報取得（PropertiesService利用）
 */
function getSecureServiceAccountCreds() {
  const props = PropertiesService.getScriptProperties();
  const credsJson = props.getProperty('SERVICE_ACCOUNT_CREDS');

  if (!credsJson) {
    throw new Error('サービスアカウント認証情報が設定されていません');
  }

  return JSON.parse(credsJson);
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
    // 基本的な引数チェック
    if (!userId || typeof userId !== 'string' || userId.trim() === '') {
      console.warn('verifyAdminAccess: 無効なuserIdが渡されました:', userId);
      return false;
    }

    // 現在のログインユーザーのメールアドレスを取得
    const activeUserEmail = UserManager.getCurrentEmail();
    if (!activeUserEmail) {
      console.warn('verifyAdminAccess: アクティブユーザーのメールアドレスが取得できませんでした');
      return false;
    }

    console.log('verifyAdminAccess: 認証開始', { userId, activeUserEmail });

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
      dbEmail,
      currentEmail,
    });

    // 3つの条件すべてが満たされた場合のみ認証成功
    if (isEmailMatched && isUserIdMatched && isActive) {
      console.log('✅ verifyAdminAccess: 認証成功', { userId, email: activeUserEmail });
      return true;
    } else {
      console.warn('❌ verifyAdminAccess: 認証失敗', {
        userId,
        activeUserEmail,
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
  const props = PropertiesService.getScriptProperties();
  return props.getProperty('DATABASE_SPREADSHEET_ID');
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
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
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
 * Sheets APIサービスを取得（getSheetsServiceCachedへのエイリアス）
 * @deprecated getSheetsServiceCached()を使用してください
 * @returns {Object|null} Sheetsサービスオブジェクト
 */
function getSheetsService() {
  // キャッシュ版に統一
  return getSheetsServiceCached();
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
    // 🔧 修正：Sheets API未有効化対応 - 直接SpreadsheetAppを使用
    console.log('updateSheetsData: SpreadsheetApp直接使用（API未有効化対応）');
    
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    // シート名と範囲を分離
    const match = range.match(/^'?([^'!]+)'?!(.+)$/);
    if (match) {
      const sheetName = match[1];
      const rangeSpec = match[2];
      const sheet = spreadsheet.getSheetByName(sheetName);
      if (sheet) {
        const targetRange = sheet.getRange(rangeSpec);
        targetRange.setValues(values);
        return {
          updatedCells: values.length * (values[0] ? values[0].length : 0),
          updatedRows: values.length,
          updatedColumns: values[0] ? values[0].length : 0,
          spreadsheetId,
          updatedRange: range
        };
      }
    }
    throw new Error(`範囲の解析に失敗しました: ${  range}`);
  } catch (error) {
    console.error('updateSheetsData SpreadsheetAppエラー:', error.message);
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
    // Service Accountを使用してSheets APIでアクセスを試みる
    if (service && service.spreadsheets) {
      try {
        console.log('batchGetSheetsData: Service Account経由でSheets API使用', {
          hasService: !!service,
          hasSpreadsheets: !!service.spreadsheets,
          hasValues: !!service.spreadsheets.values,
          hasBatchGet: !!service.spreadsheets.values.batchGet,
          batchGetType: typeof service.spreadsheets.values.batchGet
        });
        const response = service.spreadsheets.values.batchGet({
          spreadsheetId: spreadsheetId,
          ranges: ranges
        });
        return response;
      } catch (apiError) {
        console.warn('Sheets API呼び出し失敗、SpreadsheetAppにフォールバック:', apiError.message);
      }
    } else {
      console.log('batchGetSheetsData: Service利用不可', {
        hasService: !!service,
        hasSpreadsheets: service ? !!service.spreadsheets : 'service is null'
      });
    }
    
    // フォールバック：Sheets APIが使えない場合はSpreadsheetAppを使用
    console.log('batchGetSheetsData: SpreadsheetApp直接使用（フォールバック）');
    
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const valueRanges = ranges.map(range => {
      // シート名と範囲を分離
      const match = range.match(/^'?([^'!]+)'?!(.+)$/);
      if (match) {
        const sheetName = match[1];
        const rangeSpec = match[2];
        const sheet = spreadsheet.getSheetByName(sheetName);
        if (sheet) {
          const values = sheet.getRange(rangeSpec).getValues();
          return {
            range,
            values
          };
        }
      }
      return null;
    }).filter(Boolean);
    
    return {
      valueRanges
    };
  } catch (error) {
    console.error('batchGetSheetsData エラー:', error.message);
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

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
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

console.log('🔐 簡略化されたセキュリティシステムが初期化されました');
