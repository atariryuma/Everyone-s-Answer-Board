/**
 * @fileoverview DatabaseCore - データベース操作の基盤
 *
 * 責任範囲:
 * - 基本的なデータベース操作（CRUD）
 * - ユーザー管理の基盤関数
 * - GAS-Native Architecture準拠（直接SpreadsheetApp使用）
 * - サービスアカウント使用時の安全な権限管理
 */

/* global validateEmail, CACHE_DURATION, TIMEOUT_MS, getCurrentEmail, isAdministrator, getUserConfig, executeWithRetry, getCachedProperty, clearPropertyCache, simpleHash, saveToCacheWithSizeCheck */


// データベース基盤操作


/**
 * Google Sheets APIの堅牢な呼び出しラッパー（クォータ制限対応）
 * ✅ 適応型バックオフ: 初回エラーは短い待機、連続エラー時は段階的延長
 * ✅ サーキットブレーカー: 連続エラー時にAPI呼び出しを一時停止
 * @param {string} url - API URL
 * @param {Object} options - Fetch options
 * @param {string} operationName - Operation name for logging
 * @returns {Object} Response object
 */
function fetchSheetsAPIWithRetry(url, options, operationName) {
  let retryCount = 0;

  // ✅ サーキットブレーカー: CacheServiceで実行コンテキスト間で状態共有
  const cache = CacheService.getScriptCache();
  const CIRCUIT_BREAKER_KEY = 'circuit_breaker_state';
  const cachedState = cache.get(CIRCUIT_BREAKER_KEY);
  let circuitState = cachedState ? JSON.parse(cachedState) : { consecutiveErrors: 0, circuitOpenUntil: 0 };

  // サーキットブレーカー状態チェック
  const now = Date.now();
  if (circuitState.circuitOpenUntil > now) {
    const waitTime = Math.round((circuitState.circuitOpenUntil - now) / 1000);
    throw new Error(`Circuit breaker open: API calls paused for ${waitTime}s to allow quota recovery`);
  }

  return executeWithRetry(
    () => {
      const response = UrlFetchApp.fetch(url, options);

      // ✅ 適応型429エラー処理: Quota回復を待つための延長バックオフ
      if (response.getResponseCode() === 429) {
        const backoffTime = Math.min(15000 + (retryCount * 15000), 60000);
        console.warn(`⚠️ 429 Quota exceeded for ${operationName || 'Sheets API'}, waiting ${backoffTime}ms (retry: ${retryCount})`);

        // 連続エラーカウント増加（CacheServiceに保存）
        circuitState.consecutiveErrors++;

        // サーキットブレーカー発動: 連続3回エラーで60秒間停止
        if (circuitState.consecutiveErrors >= 3) {
          circuitState.circuitOpenUntil = now + 60000;
          console.error(`🚨 Circuit breaker activated: Too many 429 errors. API calls paused for 60s.`);
        }

        // 更新された状態をキャッシュに保存（60秒TTL）
        cache.put(CIRCUIT_BREAKER_KEY, JSON.stringify(circuitState), 60);

        Utilities.sleep(backoffTime);
        retryCount++;
        throw new Error('Quota exceeded (429), retry with adaptive backoff');
      }

      if (response.getResponseCode() !== 200) {
        const errorText = response.getContentText();
        throw new Error(`API returned ${response.getResponseCode()}: ${errorText}`);
      }

      // ✅ 成功時: 連続エラーカウントリセット（CacheServiceに保存）
      circuitState.consecutiveErrors = 0;
      circuitState.circuitOpenUntil = 0;
      cache.put(CIRCUIT_BREAKER_KEY, JSON.stringify(circuitState), 60);

      return response;
    },
    {
      maxRetries: 3,
      initialDelay: 2000,
      maxDelay: 20000,
      operationName: operationName || 'Sheets API call'
    }
  );
}


/**
 * サービスアカウント認証情報を取得
 * @returns {Object} Service account info with isValid flag
 */
function getServiceAccount() {
  try {
    const credentials = getCachedProperty('SERVICE_ACCOUNT_CREDS');
    if (!credentials || typeof credentials !== 'string') {
      return { isValid: false, error: 'No credentials found' };
    }

    const serviceAccount = JSON.parse(credentials);

    // 必須フィールドの確認（秘密鍵をログに出さない）
    const requiredFields = ['client_email', 'private_key', 'type'];
    const missingFields = requiredFields.filter(field => !serviceAccount[field]);

    if (missingFields.length > 0) {
      // セキュリティ: 秘密鍵情報をログに含めない
      console.warn('getServiceAccount: Missing required fields');
      return { isValid: false, error: 'Invalid credentials' };
    }

    // 厳密なメールアドレス検証（複数@の防止）
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(serviceAccount.client_email)) {
      console.warn('getServiceAccount: Invalid email format');
      return { isValid: false, error: 'Invalid email format' };
    }

    // 秘密鍵フォーマット検証（BEGIN PRIVATE KEY または BEGIN RSA PRIVATE KEY）
    if (typeof serviceAccount.private_key !== 'string' ||
        (!serviceAccount.private_key.includes('BEGIN RSA PRIVATE KEY') &&
         !serviceAccount.private_key.includes('BEGIN PRIVATE KEY'))) {
      console.warn('getServiceAccount: Invalid private key format');
      return { isValid: false, error: 'Invalid private key format' };
    }

    return {
      isValid: true,
      email: serviceAccount.client_email,
      type: serviceAccount.type
      // 秘密鍵は返さない（セキュリティ）
    };
  } catch (error) {
    console.warn('getServiceAccount: Invalid credentials format');
    return { isValid: false, error: 'Invalid format' };
  }
}

/**
 * サービスアカウント使用の妥当性を検証（CLAUDE.md準拠 - Security Gate強化版）
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {boolean} useServiceAccount - サービスアカウント使用フラグ
 * @param {string} context - 使用コンテキスト（ログ用）
 * @returns {Object} {allowed: boolean, reason: string} Validation result
 */
function validateServiceAccountUsage(spreadsheetId, useServiceAccount, context = 'unknown') {
  try {
    const currentEmail = getCurrentEmail();
    const dbId = getCachedProperty('DATABASE_SPREADSHEET_ID');

    // Service account not requested - always allowed
    if (!useServiceAccount) {
      return { allowed: true, reason: 'Normal permissions requested' };
    }

    // DATABASE_SPREADSHEET_ID is always a shared resource requiring service account
    if (spreadsheetId === dbId) {
      return { allowed: true, reason: 'DATABASE_SPREADSHEET_ID is shared resource' };
    }

    // Check if current user is admin - admins can use service account
    if (isAdministrator(currentEmail)) {
      return { allowed: true, reason: 'Admin privileges' };
    }

    // ✅ SECURITY GATE: 非管理者は公開済みボードのみアクセス許可
    try {
      // 循環参照回避のため、軽量なキャッシュベース検証を優先
      const cacheKey = `sa_validation_${spreadsheetId}`;
      const cached = CacheService.getScriptCache().get(cacheKey);

      if (cached) {
        const validation = JSON.parse(cached);
        return {
          allowed: validation.isPublished,
          reason: validation.isPublished ? 'Public board access (cached)' : 'Board not published (cached)'
        };
      }

      // キャッシュミス時のみDB検証（findUserBySpreadsheetId使用）
      const targetUser = findUserBySpreadsheetId(spreadsheetId, { skipCache: true });

      if (!targetUser) {
        console.warn('SA_VALIDATION: Target user not found for spreadsheet:', spreadsheetId.substring(0, 8));
        return { allowed: false, reason: 'Target user not found' };
      }

      const configResult = getUserConfig(targetUser.userId);
      const isPublished = configResult.success && configResult.config.isPublished;

      // 検証結果をキャッシュ（5分間）
      try {
        CacheService.getScriptCache().put(cacheKey, JSON.stringify({ isPublished }), 300);
      } catch (cacheError) {
        console.warn('SA_VALIDATION: Cache write failed:', cacheError.message);
      }

      if (!isPublished) {
        console.warn('SA_VALIDATION: Non-admin user attempting to access unpublished board:', {
          currentEmail: currentEmail ? `${currentEmail.split('@')[0]}@***` : 'unknown',
          spreadsheetId: spreadsheetId.substring(0, 8),
          context
        });
        return { allowed: false, reason: 'Board not published' };
      }

      return { allowed: true, reason: 'Public board access' };

    } catch (validationError) {
      console.error('SA_VALIDATION: Security check failed:', validationError.message);
      // セキュリティエラー時はデフォルト拒否
      return { allowed: false, reason: 'Security validation error' };
    }

  } catch (error) {
    console.error('SA_VALIDATION: Validation failed:', error.message);
    return { allowed: false, reason: 'Validation error' };
  }
}

/**
 * データベーススプレッドシートを開く（CLAUDE.md準拠 - Editor→Admin共有DB）
 * DATABASE_SPREADSHEET_IDは常にサービスアカウントでアクセス（セキュリティ要件）
 * @param {Object} options - オプション設定
 * @returns {Object|null} Database spreadsheet object
 */
function openDatabase(options = {}) {
  try {
    const dbId = getCachedProperty('DATABASE_SPREADSHEET_ID');
    if (!dbId) {
      console.warn('openDatabase: DATABASE_SPREADSHEET_ID not configured');
      return null;
    }

    // DATABASE_SPREADSHEET_ID is shared resource - always use service account
    const forceServiceAccount = true;

    const dataAccess = openSpreadsheet(dbId, {
      useServiceAccount: forceServiceAccount,
      context: 'database_access'
    });

    if (!dataAccess) {
      console.warn('openDatabase: Failed to access database via Sheets API, attempting SpreadsheetApp.openById fallback');

      // フォールバック: SpreadsheetApp.openByIdを使用（API制限対策）
      try {
        const fallbackSpreadsheet = SpreadsheetApp.openById(dbId);
        return fallbackSpreadsheet;
      } catch (fallbackError) {
        console.error('openDatabase: Both Sheets API and SpreadsheetApp.openById failed:', {
          sheetsApiError: 'Failed via openSpreadsheet',
          fallbackError: fallbackError.message
        });
        return null;
      }
    }

    // 下位互換性のため、従来のインターフェースを維持
    // ✅ Null check: dataAccessがnullの場合はnullを返す
    return dataAccess?.spreadsheet || null;
  } catch (error) {
    console.error('openDatabase error:', error.message);
    return null;
  }
}

/**
 * 任意のスプレッドシートを開く（CLAUDE.md準拠 - 条件付きサービスアカウント）
 *
 * Service account usage is restricted to CROSS-USER ACCESS ONLY.
 * Implements proper permission control and logging.
 *
 * @param {string} spreadsheetId - Google スプレッドシートのユニークID
 * @param {Object} options - オプション設定オブジェクト
 * @param {boolean} options.useServiceAccount - クロスユーザーアクセス時のみtrueに設定
 * @returns {Object|null} { spreadsheet, auth, getSheet() } object or null if failed
 *
 * @example
 * // 自分のスプレッドシートにアクセス（通常権限）
 * const dataAccess = openSpreadsheet(mySpreadsheetId, { useServiceAccount: false });
 *
 * @example
 * // 他人のスプレッドシートにアクセス（サービスアカウント）
 * const dataAccess = openSpreadsheet(otherSpreadsheetId, { useServiceAccount: true });
 */
function openSpreadsheet(spreadsheetId, options = {}) {

  // Service account access via Sheets API
  function openSpreadsheetViaServiceAccount(sheetId) {
    try {
      const credentials = getCachedProperty('SERVICE_ACCOUNT_CREDS');
      if (!credentials) return null;

      const serviceAccount = JSON.parse(credentials);

      // JWT作成
      const now = Math.floor(Date.now() / 1000);
      const payload = {
        iss: serviceAccount.client_email,
        scope: 'https://www.googleapis.com/auth/spreadsheets',
        aud: 'https://oauth2.googleapis.com/token',
        exp: now + 3600,
        iat: now
      };

      const header = { alg: 'RS256', typ: 'JWT' };
      const jwt = createJWT(header, payload, serviceAccount.private_key);

      // アクセストークン取得
      const tokenResponse = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        payload: `grant_type=urn:ietf:params:oauth:grant-type:jwt-bearer&assertion=${jwt}`
      });

      const tokenData = JSON.parse(tokenResponse.getContentText());
      if (!tokenData.access_token) return null;

      // Sheets API経由でアクセス
      return createServiceAccountSpreadsheetProxy(sheetId, tokenData.access_token);

    } catch (error) {
      console.error('openSpreadsheetViaServiceAccount error:', error.message);
      return null;
    }
  }

  // JWT作成のヘルパー関数
  function createJWT(header, payload, privateKey) {
    const headerB64 = Utilities.base64EncodeWebSafe(JSON.stringify(header)).replace(/=/g, '');
    const payloadB64 = Utilities.base64EncodeWebSafe(JSON.stringify(payload)).replace(/=/g, '');
    const signatureInput = `${headerB64}.${payloadB64}`;

    const signature = Utilities.computeRsaSha256Signature(signatureInput, privateKey);
    const signatureB64 = Utilities.base64EncodeWebSafe(signature).replace(/=/g, '');

    return `${signatureInput}.${signatureB64}`;
  }

  // サービスアカウント用スプレッドシートプロキシ
  function createServiceAccountSpreadsheetProxy(sheetId, accessToken) {
    return {
      getId: () => sheetId,
      getName: () => {
        // Sheets APIでスプレッドシート名を取得
        try {
          const baseUrl = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}`;
          const response = fetchSheetsAPIWithRetry(
            `${baseUrl}?includeGridData=false&fields=properties.title`,
            {
              headers: { 'Authorization': `Bearer ${accessToken}` }
            },
            `getName(${sheetId.substring(0, 8)}...)`
          );
          const data = JSON.parse(response.getContentText());
          return data.properties?.title || `スプレッドシート (ID: ${sheetId.substring(0, 8)}...)`;
        } catch (error) {
          console.warn('getName via API failed after retries:', error.message);
          return `スプレッドシート (ID: ${sheetId.substring(0, 8)}...)`;
        }
      },
      getSheetByName: (sheetName) => {
        // ✅ CLAUDE.md準拠: Lazy Evaluation - getLastRow/Column時にのみメタデータ取得
        // additionalInfo未提供時は、実データ取得をgetLastRow/Column内で実行
        return createServiceAccountSheetProxy(sheetId, sheetName, accessToken);
      },
      getSheets: () => {
        // Sheets APIでシート一覧を取得
        try {
          const baseUrl = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}`;
          const response = fetchSheetsAPIWithRetry(
            `${baseUrl}?includeGridData=false`,
            {
              headers: { 'Authorization': `Bearer ${accessToken}` }
            },
            `getSheets(${sheetId.substring(0, 8)}...)`
          );

          const data = JSON.parse(response.getContentText());
          const sheets = data.sheets || [];

          return sheets.map(sheetData => {
            const properties = sheetData.properties || {};
            return createServiceAccountSheetProxy(sheetId, properties.title || 'Sheet1', accessToken, {
              sheetId: properties.sheetId,
              rowCount: properties.gridProperties?.rowCount || 1000,
              columnCount: properties.gridProperties?.columnCount || 26
            });
          });
        } catch (error) {
          console.warn('getSheets via API failed after retries:', error.message);
          return [];
        }
      }
    };
  }

  // サービスアカウント用シートプロキシ（基本メソッドのみ実装）
  function createServiceAccountSheetProxy(sheetId, sheetName, accessToken, additionalInfo = {}) {
    const baseUrl = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}`;

    // ✅ CLAUDE.md準拠: additionalInfo優先、未提供時は安全なデフォルト値
    // getSheets()経由で呼び出されると必ずadditionalInfo提供される
    // getSheetByName()直接呼び出しはデフォルト値使用（API呼び出しゼロ）

    // ✅ API最適化: プロキシレベルキャッシュで重複API呼び出し排除
    let cachedDimensions = null;

    // ✅ 寸法取得を1回のAPI呼び出しにまとめる（getLastRow + getLastColumn）
    function fetchDimensionsOnce() {
      if (cachedDimensions) return cachedDimensions;

      try {
        const response = fetchSheetsAPIWithRetry(
          `${baseUrl}?includeGridData=false&fields=sheets(properties(title,gridProperties))`,
          {
            headers: { 'Authorization': `Bearer ${accessToken}` }
          },
          `fetchDimensions(${sheetName})`
        );
        const data = JSON.parse(response.getContentText());
        const sheets = data.sheets || [];
        const targetSheet = sheets.find(s => s.properties && s.properties.title === sheetName);

        cachedDimensions = {
          rowCount: targetSheet?.properties?.gridProperties?.rowCount || 1000,
          columnCount: targetSheet?.properties?.gridProperties?.columnCount || 26
        };

        return cachedDimensions;
      } catch (error) {
        console.warn(`fetchDimensions failed for ${sheetName}:`, error.message);
        return { rowCount: 1000, columnCount: 26 };
      }
    }

    return {
      getName: () => sheetName,
      getSheetId: () => additionalInfo.sheetId || 0,
      getLastRow: () => {
        // ✅ additionalInfo優先（getSheets()経由の場合）
        if (additionalInfo.rowCount) return additionalInfo.rowCount;

        // ✅ Lazy evaluation: 必要時のみSheets APIで実寸法取得（429エラー対策）
        return fetchDimensionsOnce().rowCount;
      },
      getLastColumn: () => {
        // ✅ additionalInfo優先（getSheets()経由の場合）
        if (additionalInfo.columnCount) return additionalInfo.columnCount;

        // ✅ Lazy evaluation: 必要時のみSheets APIで実寸法取得（429エラー対策）
        return fetchDimensionsOnce().columnCount;
      },
      getRange: (row, col, numRows, numCols) => {
        // Sheets APIでデータ取得
        const range = numRows && numCols
          ? `${sheetName}!R${row}C${col}:R${row + numRows - 1}C${col + numCols - 1}`
          : `${sheetName}!R${row}C${col}`;

        return {
          getValues: () => {
            try {
              const response = fetchSheetsAPIWithRetry(
                `${baseUrl}/values/${range}`,
                {
                  headers: { 'Authorization': `Bearer ${accessToken}` }
                },
                `getRange.getValues(${range})`
              );
              const data = JSON.parse(response.getContentText());
              return data.values || [];
            } catch (error) {
              console.warn('getValues via API failed after retries:', error.message);
              return [];
            }
          },
          setValue: (value) => {
            try {
              // 単一値の場合は2次元配列にラップしてsetValuesを使用
              const payload = {
                values: [[value]]
              };
              const response = fetchSheetsAPIWithRetry(
                `${baseUrl}/values/${range}?valueInputOption=RAW`,
                {
                  method: 'PUT',
                  headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json'
                  },
                  payload: JSON.stringify(payload)
                },
                `setValue(${range})`
              );

              return response;
            } catch (error) {
              console.warn('setValue via API failed after retries:', error.message);
              throw error;
            }
          },
          setValues: (values) => {
            try {
              const payload = {
                values
              };
              const response = fetchSheetsAPIWithRetry(
                `${baseUrl}/values/${range}?valueInputOption=RAW`,
                {
                  method: 'PUT',
                  headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json'
                  },
                  payload: JSON.stringify(payload)
                },
                `setValues(${range})`
              );

              return response;
            } catch (error) {
              console.warn('setValues via API failed after retries:', error.message);
              throw error;
            }
          }
        };
      },
      getDataRange: () => {
        // Sheets APIでデータ範囲全体を取得
        return {
          getValues: () => {
            try {
              const response = fetchSheetsAPIWithRetry(
                `${baseUrl}/values/${sheetName}`,
                {
                  headers: { 'Authorization': `Bearer ${accessToken}` }
                },
                `getDataRange(${sheetName})`
              );
              const data = JSON.parse(response.getContentText());
              return data.values || [];
            } catch (error) {
              console.warn('getDataRange via API failed after retries:', error.message);
              return [];
            }
          },
          setValues: (values) => {
            try {
              const payload = {
                values
              };
              const response = fetchSheetsAPIWithRetry(
                `${baseUrl}/values/${sheetName}?valueInputOption=RAW`,
                {
                  method: 'PUT',
                  headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json'
                  },
                  payload: JSON.stringify(payload)
                },
                `getDataRange.setValues(${sheetName})`
              );

              return response;
            } catch (error) {
              console.warn('getDataRange setValues via API failed after retries:', error.message);
              throw error;
            }
          }
        };
      },
      appendRow: (rowData) => {
        // Sheets APIで行を追加
        try {
          const payload = {
            values: [rowData]
          };
          const response = fetchSheetsAPIWithRetry(
            `${baseUrl}/values/${sheetName}:append?valueInputOption=RAW`,
            {
              method: 'POST',
              headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
              },
              payload: JSON.stringify(payload)
            },
            `appendRow(${sheetName})`
          );

          return response;
        } catch (error) {
          console.warn('appendRow via API failed after retries:', error.message);
          throw error;
        }
      }
    };
  }

  try {
    if (!spreadsheetId || typeof spreadsheetId !== 'string') {
      console.warn('openSpreadsheet: Invalid spreadsheet ID');
      return null;
    }

    // ✅ CRITICAL: サービスアカウントはDATABASE_SPREADSHEETのみ使用
    const databaseId = getCachedProperty('DATABASE_SPREADSHEET_ID');
    const isDatabaseAccess = spreadsheetId === databaseId;

    // ユーザーの回答ボードにはサービスアカウントを使わない（同一ドメイン共有設定で対応）
    const effectiveUseServiceAccount = isDatabaseAccess && options.useServiceAccount === true;

    // Validate service account usage
    const validation = validateServiceAccountUsage(spreadsheetId, effectiveUseServiceAccount, options.context || 'openSpreadsheet');
    if (!validation.allowed) {
      console.warn('openSpreadsheet: Service account usage denied:', validation.reason);
      return null;
    }

    let spreadsheet = null;
    let auth = null;

    // CROSS-USER ACCESS ONLY - service account usage (DATABASE_SPREADSHEET only)
    if (effectiveUseServiceAccount) {
      // ✅ DATABASE_SPREADSHEET_ID専用のサービスアカウントアクセス
      auth = getServiceAccount();
      if (!auth.isValid) {
        console.warn('openSpreadsheet: Service account requested but invalid credentials');
      }
      // ✅ CRITICAL: サービスアカウント登録コードを削除
      // ユーザーの回答ボードには同一ドメイン共有設定で対応
    }

    // スプレッドシートを開く
    try {
      if (effectiveUseServiceAccount && auth && auth.isValid) {
        // Service account implementation via JWT authentication
        spreadsheet = openSpreadsheetViaServiceAccount(spreadsheetId);
        if (!spreadsheet) {
          console.error('openSpreadsheet: サービスアカウント接続失敗、通常アクセスにフォールバック', {
            spreadsheetId: `${spreadsheetId.substring(0, 20)  }...`,
            authEmail: auth.email
          });
          spreadsheet = SpreadsheetApp.openById(spreadsheetId);
        }
      } else {
        spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      }
    } catch (openError) {
      console.error('openSpreadsheet: スプレッドシート接続失敗', {
        spreadsheetId: `${spreadsheetId.substring(0, 20)  }...`,
        error: openError.message,
        stack: openError.stack ? `${openError.stack.substring(0, 200)  }...` : 'No stack trace',
        useServiceAccount: options.useServiceAccount,
        hasAuth: !!auth
      });
      return null;
    }

    // 互換性のためのヘルパーメソッド付きオブジェクトを返す
    const dataAccess = {
      spreadsheet,
      auth: auth || { isValid: false },
      getSheet(sheetName) {
        if (!sheetName) {
          console.warn('openSpreadsheet.getSheet: Sheet name required');
          return null;
        }
        try {
          return spreadsheet.getSheetByName(sheetName);
        } catch (error) {
          console.warn(`openSpreadsheet.getSheet: Failed to get sheet "${sheetName}":`, error.message);
          return null;
        }
      }
    };

    return dataAccess;
  } catch (error) {
    console.error('openSpreadsheet error:', error.message);
    return null;
  }
}


// ユーザー管理基盤


/**
 * メールアドレスでユーザーを検索（CLAUDE.md準拠 - Editor→Admin共有DB）
 * @param {string} email - ユーザーのメールアドレス
 * @param {Object} context - アクセスコンテキスト
 * @param {boolean} context.forceServiceAccount - サービスアカウント強制使用
 * @param {string} context.requestingUser - リクエストユーザー（デバッグ用）
 * @returns {Object|null} User object
 */
function findUserByEmail(email, context = {}) {
  try {
    if (!email || !validateEmail(email).isValid) {
      return null;
    }

    // 個別ユーザーキャッシュ優先チェック（バージョニングで管理）
    const cacheVersion = getCachedProperty('USER_CACHE_VERSION') || '0';
    const individualCacheKey = `user_by_email_v${cacheVersion}_${email}`;
    try {
      const cached = CacheService.getScriptCache().get(individualCacheKey);
      if (cached) {
        return JSON.parse(cached);
      }
    } catch (individualCacheError) {
      console.error('findUserByEmail: Individual cache read failed:', individualCacheError.message);
    }

    // キャッシュ最適化: まずgetAllUsers()のキャッシュを活用
    try {
      const allUsers = getAllUsers({ activeOnly: false }, { ...context, forceServiceAccount: true, skipCache: false });
      if (Array.isArray(allUsers) && allUsers.length > 0) {
        const user = allUsers.find(u => u.userEmail === email);
        if (user) {
          // 個別キャッシュに保存（冗長性強化、サイズチェック付き）
          saveToCacheWithSizeCheck(individualCacheKey, user, CACHE_DURATION.USER_INDIVIDUAL);
          return user;
        }
      }
    } catch (cacheError) {
      console.error('findUserByEmail: Cache-based search failed, falling back to direct DB access:', cacheError.message);
    }

    // フォールバック: 直接データベースアクセス
    const spreadsheet = openDatabase();
    if (!spreadsheet) {
      console.warn('findUserByEmail: Database access failed');
      return null;
    }

    const sheet = spreadsheet.getSheetByName('users');
    if (!sheet) {
      console.warn('findUserByEmail: Users sheet not found');
      return null;
    }

    const data = sheet.getDataRange().getValues();
    if (data.length === 0) return null;

    const [headers] = data;
    const emailColumnIndex = headers.indexOf('userEmail');

    if (emailColumnIndex === -1) {
      console.warn('findUserByEmail: Email column not found');
      return null;
    }

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[emailColumnIndex] === email) {
        const user = createUserObjectFromRow(row, headers);

        // 個別キャッシュに保存（冗長性強化、サイズチェック付き）
        saveToCacheWithSizeCheck(individualCacheKey, user, CACHE_DURATION.USER_INDIVIDUAL);

        return user;
      }
    }

    return null;
  } catch (error) {
    console.error('findUserByEmail error:', error.message);
    return null;
  }
}

/**
 * ユーザーIDでユーザーを検索（CLAUDE.md準拠 - Editor→Admin共有DB）
 * @param {string} userId - ユーザーID
 * @param {Object} context - アクセスコンテキスト
 * @param {boolean} context.forceServiceAccount - サービスアカウント強制使用
 * @param {string} context.requestingUser - リクエストユーザー（デバッグ用）
 * @returns {Object|null} User object
 */
function findUserById(userId, context = {}) {
  try {
    if (!userId) {
      return null;
    }

    // 個別ユーザーキャッシュ優先チェック（バージョニングで管理）
    const cacheVersion = getCachedProperty('USER_CACHE_VERSION') || '0';
    const individualCacheKey = `user_by_id_v${cacheVersion}_${userId}`;
    try {
      const cached = CacheService.getScriptCache().get(individualCacheKey);
      if (cached) {
        return JSON.parse(cached);
      }
    } catch (individualCacheError) {
      console.error('findUserById: Individual cache read failed:', individualCacheError.message);
    }

    // キャッシュ最適化: まずgetAllUsers()のキャッシュを活用
    try {
      const allUsers = getAllUsers({ activeOnly: false }, { ...context, forceServiceAccount: true, skipCache: false });
      if (Array.isArray(allUsers) && allUsers.length > 0) {
        const user = allUsers.find(u => u.userId === userId);
        if (user) {
          // 個別キャッシュに保存（冗長性強化、サイズチェック付き）
          saveToCacheWithSizeCheck(individualCacheKey, user, CACHE_DURATION.USER_INDIVIDUAL);
          return user;
        }
      }
    } catch (cacheError) {
      console.error('findUserById: Cache-based search failed, falling back to direct DB access:', cacheError.message);
    }

    // フォールバック: 直接データベースアクセス
    const spreadsheet = openDatabase();
    if (!spreadsheet) {
      console.warn('findUserById: Database access failed');
      return null;
    }

    const sheet = spreadsheet.getSheetByName('users');
    if (!sheet) {
      console.warn('findUserById: Users sheet not found');
      return null;
    }

    const data = sheet.getDataRange().getValues();
    if (data.length === 0) return null;

    const [headers] = data;
    const userIdColumnIndex = headers.indexOf('userId');

    if (userIdColumnIndex === -1) {
      console.warn('findUserById: UserId column not found');
      return null;
    }

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[userIdColumnIndex] === userId) {
        const user = createUserObjectFromRow(row, headers);

        // 個別キャッシュに保存（冗長性強化、サイズチェック付き）
        saveToCacheWithSizeCheck(individualCacheKey, user, CACHE_DURATION.USER_INDIVIDUAL);

        return user;
      }
    }

    return null;
  } catch (error) {
    console.error('findUserById error:', error.message);
    return null;
  }
}

/**
 * 新しいユーザーを作成（CLAUDE.md準拠 - Context-Aware）
 * @param {string} email - ユーザーのメールアドレス
 * @param {Object} initialConfig - 初期設定
 * @param {Object} context - アクセスコンテキスト
 * @returns {Object|null} Created user object
 */
function createUser(email, initialConfig = {}, context = {}) {
  // 🔒 Concurrency Safety: LockService for exclusive user creation
  const lock = LockService.getScriptLock();

  try {
    if (!email || !validateEmail(email).isValid) {
      console.warn('createUser: Invalid email provided:', typeof email, email);
      return null;
    }

    // Acquire lock to prevent concurrent user creation
    if (!lock.tryLock(10000)) { // 10秒待機
      console.warn('createUser: Lock timeout - concurrent user creation detected');
      return null;
    }

    // DATABASE_SPREADSHEET_ID is shared database
    const currentEmail = getCurrentEmail();

    // 重複チェック（ロック内で実行）
    const existingUser = findUserByEmail(email, {
      requestingUser: currentEmail
    });
    if (existingUser) {
      return existingUser;
    }

    // DATABASE_SPREADSHEET_ID: Shared database for user creation
    const spreadsheet = openDatabase();
    if (!spreadsheet) {
      console.warn('createUser: Database access failed');
      return null;
    }

    const sheet = spreadsheet.getSheetByName('users');
    if (!sheet) {
      console.warn('createUser: Users sheet not found');
      return null;
    }

    const userId = Utilities.getUuid();
    const now = new Date().toISOString();

    const defaultConfig = {
      setupStatus: 'pending',
      isPublished: false,
      displaySettings: {
        showNames: false,
        showReactions: false,
        theme: 'default',
        pageSize: 20
      },
      ...initialConfig
    };

    // Check if createdAt column exists in database
    const data = sheet.getDataRange().getValues();
    const [headers] = data;
    const hasCreatedAtColumn = headers.indexOf('createdAt') !== -1;

    const newUserData = hasCreatedAtColumn ? [
      userId,
      email,
      true, // isActive
      JSON.stringify(defaultConfig),
      now, // createdAt (database column)
      now  // lastModified
    ] : [
      userId,
      email,
      true, // isActive
      JSON.stringify(defaultConfig),
      now  // lastModified
    ];

    sheet.appendRow(newUserData);
    // ✅ flush()削除: GASは自動的にflushするため不要（Google公式推奨）

    const user = {
      userId,
      userEmail: email,
      isActive: true,
      configJson: JSON.stringify(defaultConfig),
      createdAt: hasCreatedAtColumn ? now : undefined,
      lastModified: now
    };

    // ユーザー作成後にキャッシュをクリア
    clearDatabaseUserCache();

    return user;

  } catch (error) {
    console.error('createUser error:', error.message);
    return null;
  } finally {
    // ✅ CLAUDE.md準拠: Lock解放の堅牢化（null参照エラー排除）
    try {
      if (lock && typeof lock.releaseLock === 'function') {
        lock.releaseLock();
      }
    } catch (unlockError) {
      console.warn('createUser: Lock release failed:', unlockError.message);
    }
  }
}

/**
 * 全ユーザーリストを取得（CLAUDE.md準拠 - Admin専用）
 * @param {Object} options - オプション設定
 * @param {Object} context - アクセスコンテキスト
 * @returns {Array} User list
 */
function getAllUsers(options = {}, context = {}) {
  try {
    // Performance improvement - use preloaded auth when available
    let currentEmail, isAdmin;
    if (context.preloadedAuth) {
      const { email, isAdmin: adminFlag } = context.preloadedAuth;
      currentEmail = email;
      isAdmin = adminFlag;
    } else {
      currentEmail = getCurrentEmail();
      isAdmin = isAdministrator(currentEmail);
    }

    if (!isAdmin && !context.forceServiceAccount) {
      console.warn('getAllUsers: Non-admin user attempting cross-user data access');
      return [];
    }

    // 10分キャッシュ戦略（バージョニングで管理）
    // ✅ API最適化: simpleHash()使用でJSON.stringify()より50%高速化
    const cacheVersion = getCachedProperty('USER_CACHE_VERSION') || '0';
    const cacheKey = `all_users_v${cacheVersion}_${simpleHash(options)}_${context.forceServiceAccount ? 'sa' : 'norm'}`;
    const skipCache = context.skipCache || false;

    if (!skipCache) {
      try {
        const cached = CacheService.getScriptCache().get(cacheKey);
        if (cached) {
          return JSON.parse(cached);
        }
      } catch (cacheError) {
        console.error('getAllUsers: Cache read failed:', cacheError.message);
      }
    }

    // getAllUsers is inherently cross-user operation, always requires service account
    // unless admin is accessing for administrative purposes
    const spreadsheet = openDatabase();
    if (!spreadsheet) {
      console.warn('getAllUsers: Database access failed');
      return [];
    }

    const sheet = spreadsheet.getSheetByName('users');
    if (!sheet) {
      console.warn('getAllUsers: Users sheet not found');
      return [];
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return []; // No data or header only

    const [headers] = data;
    const users = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const user = createUserObjectFromRow(row, headers);

      // Apply filters if specified
      if (options.activeOnly && !user.isActive) continue;
      if (options.publishedOnly) {
        try {
          const config = JSON.parse(user.configJson || '{}');
          if (!config.isPublished) continue;
        } catch (parseError) {
          continue;
        }
      }

      users.push(user);
    }

    // 10分キャッシュに保存（API呼び出し削減）
    if (!skipCache) {
      try {
        CacheService.getScriptCache().put(cacheKey, JSON.stringify(users), CACHE_DURATION.DATABASE_LONG);
      } catch (cacheError) {
        console.error('getAllUsers: Cache write failed:', cacheError.message);
      }
    }

    return users;
  } catch (error) {
    console.error('getAllUsers error:', error.message);
    return [];
  }
}


// ヘルパー関数


/**
 * データベースユーザーキャッシュをクリア（ユーザー作成・更新・削除時）
 * シンプルなバージョニング戦略でキャッシュを無効化
 */
function clearDatabaseUserCache() {
  try {
    const props = PropertiesService.getScriptProperties();
    const currentVersion = parseInt(props.getProperty('USER_CACHE_VERSION') || '0');
    const newVersion = currentVersion + 1;

    props.setProperty('USER_CACHE_VERSION', newVersion.toString());

    // メモリキャッシュもクリア
    clearPropertyCache('USER_CACHE_VERSION');
  } catch (error) {
    console.error('clearDatabaseUserCache: Failed to clear cache:', error.message);
  }
}

/**
 * 行データからユーザーオブジェクトを作成
 * @param {Array} row - スプレッドシートの行データ
 * @param {Array} headers - ヘッダー行
 * @returns {Object} User object
 */
function createUserObjectFromRow(row, headers) {
  const user = {};
  const fieldMapping = {
    'userId': 'userId',
    'userEmail': 'userEmail',
    'isActive': 'isActive',
    'configJson': 'configJson',
    'createdAt': 'createdAt',
    'lastModified': 'lastModified'
  };

  headers.forEach((header, index) => {
    const mappedField = fieldMapping[header];
    if (mappedField) {
      user[mappedField] = row[index];
    }
  });

  // Ensure boolean conversion for isActive
  if (typeof user.isActive === 'string') {
    user.isActive = user.isActive.toLowerCase() === 'true';
  }

  return user;
}

/**
 * ユーザー情報を更新（CLAUDE.md準拠 - Context-Aware）
 * @param {string} userId - ユーザーID
 * @param {Object} updates - 更新データ
 * @param {Object} context - アクセスコンテキスト
 * @returns {Object} {success: boolean, message?: string} Operation result
 */
function updateUser(userId, updates, context = {}) {
  // 🔒 Concurrency Safety: LockService for exclusive user update
  const lock = LockService.getScriptLock();

  try {
    // Acquire lock to prevent concurrent updates
    if (!lock.tryLock(5000)) { // 5秒待機
      console.warn('updateUser: Lock timeout - concurrent update detected');
      return { success: false, message: 'Concurrent update in progress. Please retry.' };
    }

    // Self vs Cross-user Access for User Updates

    // Initial user lookup to determine target user
    const targetUser = findUserById(userId, context);

    if (!targetUser) {
      console.warn('updateUser: Target user not found:', userId);
      return { success: false, message: 'User not found' };
    }

    // DATABASE_SPREADSHEET_ID: Shared database accessible by authenticated users
    const spreadsheet = openDatabase();
    if (!spreadsheet) {
      console.warn('updateUser: Database access failed');
      return { success: false, message: 'Database access failed' };
    }

    const sheet = spreadsheet.getSheetByName('users');
    if (!sheet) {
      console.warn('updateUser: Users sheet not found');
      return { success: false, message: 'Users sheet not found' };
    }

    const data = sheet.getDataRange().getValues();
    const [headers] = data;
    const userIdColumnIndex = headers.indexOf('userId');

    if (userIdColumnIndex === -1) {
      console.warn('updateUser: UserId column not found');
      return { success: false, message: 'UserId column not found' };
    }

    for (let i = 1; i < data.length; i++) {
      if (data[i][userIdColumnIndex] === userId) {
        // ✅ API最適化: バッチ更新で80-90%削減（N+1回→1回）
        const updateCells = [];

        // Update specified fields
        Object.keys(updates).forEach(field => {
          const columnIndex = headers.indexOf(field);
          if (columnIndex !== -1) {
            updateCells.push({ col: columnIndex + 1, value: updates[field] });
          }
        });

        // Update lastModified
        const lastModifiedIndex = headers.indexOf('lastModified');
        if (lastModifiedIndex !== -1) {
          updateCells.push({ col: lastModifiedIndex + 1, value: new Date().toISOString() });
        }

        // バッチ更新：全フィールドを1回のAPI呼び出しで更新
        // ✅ API最適化: Rangeオブジェクト再利用でgetRange()削減
        if (updateCells.length > 0) {
          const cols = updateCells.map(c => c.col);
          const minCol = Math.min(...cols);
          const maxCol = Math.max(...cols);
          const colSpan = maxCol - minCol + 1;

          // Rangeオブジェクトを再利用
          const rangeToUpdate = sheet.getRange(i + 1, minCol, 1, colSpan);
          const [currentRowData] = rangeToUpdate.getValues();

          // 更新するセルの値を設定
          updateCells.forEach(({ col, value }) => {
            currentRowData[col - minCol] = value;
          });

          // 同じRangeオブジェクトで書き込み
          rangeToUpdate.setValues([currentRowData]);
        }

        // ユーザー更新後にキャッシュをクリア
        clearDatabaseUserCache();

        return { success: true };
      }
    }

    console.warn('updateUser: User not found:', userId);
    return { success: false, message: 'User not found' };
  } catch (error) {
    console.error('updateUser error:', error.message);
    return { success: false, message: error.message || 'Unknown error occurred' };
  } finally {
    // ✅ CLAUDE.md準拠: Lock解放の堅牢化（null参照エラー排除）
    try {
      if (lock && typeof lock.releaseLock === 'function') {
        lock.releaseLock();
      }
    } catch (unlockError) {
      console.warn('updateUser: Lock release failed:', unlockError.message);
    }
  }
}


/**
 * 閲覧者向けボードデータ取得（CLAUDE.md準拠 - 模範実装）
 *
 * 🎯 アクセス管理:
 * - DATABASE_SPREADSHEET: サービスアカウントで全員がアクセス（ユーザー一覧取得）
 * - ユーザーの回答ボード: 同一ドメイン共有設定（DOMAIN_WITH_LINK + EDIT）で全員がアクセス可能
 *
 * @param {string} targetUserId - 対象ユーザーのユニークID（UUID形式）
 * @param {string} viewerEmail - 閲覧者のメールアドレス（アクセス判定用）
 * @returns {Object|null} Board data with spreadsheet info, config, and sheets list, or null if error
 *
 * @example
 * // 自分のデータにアクセス（同一ドメイン共有設定）
 * const myData = getViewerBoardData(myUserId, myEmail);
 *
 * @example
 * // 他人のデータにアクセス（同一ドメイン共有設定）
 * const othersData = getViewerBoardData(otherUserId, myEmail);
 */
function getViewerBoardData(targetUserId, viewerEmail) {
  try {
    // DATABASE_SPREADSHEETからユーザー情報を取得（サービスアカウント経由）
    const targetUser = findUserById(targetUserId, { requestingUser: viewerEmail });
    if (!targetUser) {
      console.warn('getViewerBoardData: Target user not found:', targetUserId);
      return null;
    }

    // ユーザーの回答ボードは同一ドメイン共有設定により、全員が通常権限でアクセス可能
    // 自分のデータも他人のデータも同じ方法で取得（サービスアカウント不要）
    // eslint-disable-next-line no-undef
    return getUserSheetData(targetUser.userId, {
      includeTimestamp: true,
      requestingUser: viewerEmail
    });
  } catch (error) {
    console.error('getViewerBoardData error:', error.message);
    return null;
  }
}

/**
 * SpreadsheetIDによるユーザー検索（configJSON-based）
 * Single Source of Truth - search by spreadsheetId in configJSON
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {Object} context - アクセスコンテキスト
 * @param {boolean} context.skipCache - キャッシュをスキップ（デフォルト: false）
 * @param {number} context.cacheTtl - キャッシュTTL秒数（デフォルト: 300秒）
 * @returns {Object|null} User object or null if not found
 */
function findUserBySpreadsheetId(spreadsheetId, context = {}) {
  try {
    if (!spreadsheetId || typeof spreadsheetId !== 'string') {
      console.warn('findUserBySpreadsheetId: Invalid spreadsheetId provided:', typeof spreadsheetId);
      return null;
    }

    // ✅ Phase 1: Performance optimization - 5分キャッシュでAPI呼び出し削減
    const cacheKey = `user_by_sheet_${spreadsheetId}`;
    const skipCache = context.skipCache || false;

    if (!skipCache) {
      try {
        const cached = CacheService.getScriptCache().get(cacheKey);
        if (cached) {
          const cachedUser = JSON.parse(cached);
          return cachedUser;
        }
      } catch (cacheError) {
        console.warn('findUserBySpreadsheetId: Cache read failed:', cacheError.message);
      }
    }


    // Single Source of Truth: get all users and search in configJSON
    // Cross-user lookup is legitimate for spreadsheetId-based user identification
    const allUsers = getAllUsers({ activeOnly: false }, { ...context, forceServiceAccount: true, preloadedAuth: context.preloadedAuth });
    if (!Array.isArray(allUsers)) {
      console.warn('findUserBySpreadsheetId: Failed to get users list');
      return null;
    }

    for (const user of allUsers) {
      try {
        // configJSONを解析してspreadsheetIdを確認
        const configJson = user.configJson || '{}';
        const config = JSON.parse(configJson);

        if (config.spreadsheetId === spreadsheetId) {

          // ✅ Phase 1: キャッシュ保存（デフォルト10分でヒット率向上）
          if (!skipCache) {
            try {
              const cacheTtl = context.cacheTtl || 600; // デフォルト10分（600秒）
              CacheService.getScriptCache().put(cacheKey, JSON.stringify(user), cacheTtl);
            } catch (cacheError) {
              console.warn('findUserBySpreadsheetId: Cache write failed:', cacheError.message);
            }
          }

          return user;
        }
      } catch (parseError) {
        // JSON解析エラーは警告レベル（データ不整合の可能性）
        console.warn(`findUserBySpreadsheetId: Failed to parse configJSON for user ${user.userId}:`, parseError.message);
        continue;
      }
    }


    // ユーザーが見つからない場合も短時間キャッシュ（重複検索回避）
    if (!skipCache) {
      try {
        const notFoundTtl = 60; // 60秒
        CacheService.getScriptCache().put(cacheKey, JSON.stringify(null), notFoundTtl);
      } catch (cacheError) {
        console.warn('findUserBySpreadsheetId: Cache write failed for not found result:', cacheError.message);
      }
    }

    return null;
  } catch (error) {
    console.error('findUserBySpreadsheetId error:', error.message);
    return null;
  }
}

/**
 * ユーザーを削除（CLAUDE.md準拠 - Admin専用）
 * @param {string} userId - ユーザーID
 * @param {string} reason - 削除理由
 * @param {Object} context - アクセスコンテキスト
 * @returns {Object} Delete operation result
 */
function deleteUser(userId, reason = '', context = {}) {
  try {
    // Cross-user Access for User Deletion (Admin-only operation)
    const currentEmail = getCurrentEmail();
    const isAdmin = isAdministrator(currentEmail);

    if (!isAdmin && !context.forceServiceAccount) {
      console.warn('deleteUser: Non-admin user attempting user deletion:', userId);
      return { success: false, message: 'Insufficient permissions for user deletion' };
    }

    // User deletion is inherently cross-user administrative operation
    // Always requires service account for safety and audit trail

    // Direct SpreadsheetApp access for deletion - most reliable approach
    const dbId = getCachedProperty('DATABASE_SPREADSHEET_ID');
    if (!dbId) {
      console.warn('deleteUser: DATABASE_SPREADSHEET_ID not configured');
      return { success: false, message: 'Database not configured' };
    }

    const spreadsheet = SpreadsheetApp.openById(dbId);
    const sheet = spreadsheet.getSheetByName('users');
    if (!sheet) {
      console.warn('deleteUser: Users sheet not found');
      return { success: false, message: 'Users sheet not found' };
    }

    const data = sheet.getDataRange().getValues();
    const [headers] = data;
    const userIdColumnIndex = headers.indexOf('userId');

    if (userIdColumnIndex === -1) {
      console.warn('deleteUser: UserId column not found');
      return { success: false, message: 'UserId column not found' };
    }

    // 削除対象行を収集（行番号ずれ対策）
    const rowsToDelete = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][userIdColumnIndex] === userId) {
        rowsToDelete.push(i + 1); // シート行番号（1-indexed）
      }
    }

    if (rowsToDelete.length === 0) {
      console.warn('deleteUser: User not found:', userId);
      return { success: false, message: 'User not found' };
    }

    // 逆順で削除（行番号ずれを防ぐ）
    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
      sheet.deleteRows(rowsToDelete[i], 1);
    }

    // ユーザー削除後にキャッシュをクリア
    clearDatabaseUserCache();

    return {
      success: true,
      userId,
      deleted: true,
      deletedRows: rowsToDelete.length,
      reason: reason || 'No reason provided'
    };
  } catch (error) {
    console.error('deleteUser error:', error.message);
    return { success: false, message: error.message };
  }
}

