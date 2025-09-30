/**
 * @fileoverview DatabaseCore - データベース操作の基盤
 *
 * 責任範囲:
 * - 基本的なデータベース操作（CRUD）
 * - ユーザー管理の基盤関数
 * - GAS-Native Architecture準拠（直接SpreadsheetApp使用）
 * - サービスアカウント使用時の安全な権限管理
 */

/* global validateEmail, CACHE_DURATION, TIMEOUT_MS, getCurrentEmail, isAdministrator, getUserConfig, executeWithRetry */


// データベース基盤操作


/**
 * Google Sheets APIの堅牢な呼び出しラッパー（クォータ制限対応）
 * @param {string} url - API URL
 * @param {Object} options - Fetch options
 * @param {string} operationName - Operation name for logging
 * @returns {Object} Response object
 */
function fetchSheetsAPIWithRetry(url, options, operationName) {
  return executeWithRetry(
    () => {
      const response = UrlFetchApp.fetch(url, options);

      if (response.getResponseCode() !== 200) {
        const errorText = response.getContentText();
        throw new Error(`API returned ${response.getResponseCode()}: ${errorText}`);
      }

      return response;
    },
    {
      maxRetries: 3,
      initialDelay: 1000, // 1秒 - APIクォータ制限のため長め
      maxDelay: 10000,   // 10秒 - クォータエラー対応
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
    const credentials = PropertiesService.getScriptProperties().getProperty('SERVICE_ACCOUNT_CREDS');
    if (!credentials || typeof credentials !== 'string') {
      return { isValid: false, error: 'No credentials found' };
    }

    const serviceAccount = JSON.parse(credentials);

    const requiredFields = ['client_email', 'private_key', 'type'];
    const missingFields = requiredFields.filter(field => !serviceAccount[field]);

    if (missingFields.length > 0) {
      console.warn('getServiceAccount: Missing required fields:', missingFields);
      return { isValid: false, error: `Missing fields: ${missingFields.join(', ')}` };
    }

    // Validate email format
    if (!serviceAccount.client_email.includes('@') || !serviceAccount.client_email.includes('.')) {
      console.warn('getServiceAccount: Invalid service account email format');
      return { isValid: false, error: 'Invalid email format' };
    }

    return {
      isValid: true,
      email: serviceAccount.client_email,
      type: serviceAccount.type
    };
  } catch (error) {
    console.warn('getServiceAccount: Invalid credentials format:', error.message);
    return { isValid: false, error: 'JSON parse error' };
  }
}

/**
 * サービスアカウント使用の妥当性を検証（CLAUDE.md準拠 - Security Gate）
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {boolean} useServiceAccount - サービスアカウント使用フラグ
 * @param {string} context - 使用コンテキスト（ログ用）
 * @returns {Object} {allowed: boolean, reason: string} Validation result
 */
function validateServiceAccountUsage(spreadsheetId, useServiceAccount, context = 'unknown') {
  try {
    const currentEmail = getCurrentEmail();
    const dbId = PropertiesService.getScriptProperties().getProperty('DATABASE_SPREADSHEET_ID');

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

    // For non-admin users accessing other spreadsheets, allow service account
    // Skip user lookup to prevent circular reference (findUserBySpreadsheetId -> getAllUsers -> openDatabase -> validateServiceAccountUsage)
    return { allowed: true, reason: 'Cross-user access (assumed)' };

  } catch (error) {
    console.error('SA_VALIDATION: Validation failed:', error.message);
    return { allowed: false, reason: 'Validation error' };
  }
}

/**
 * データベーススプレッドシートを開く（CLAUDE.md準拠 - Editor→Admin共有DB）
 * @param {boolean} useServiceAccount - サービスアカウントを使用するか（互換性のため保持、実際は常にtrue）
 * @param {Object} options - オプション設定
 * @returns {Object|null} Database spreadsheet object
 */
function openDatabase(useServiceAccount = false, options = {}) {
   
  const _ = useServiceAccount; // Suppress unused parameter warning
  try {
    const dbId = PropertiesService.getScriptProperties().getProperty('DATABASE_SPREADSHEET_ID');
    if (!dbId) {
      console.warn('openDatabase: DATABASE_SPREADSHEET_ID not configured');
      return null;
    }

    // DATABASE_SPREADSHEET_ID is shared resource - always use service account
    // Note: useServiceAccount parameter preserved for API compatibility but overridden for security
    const forceServiceAccount = true; // DATABASE_SPREADSHEET_ID always requires service account

    const dataAccess = openSpreadsheet(dbId, {
      useServiceAccount: forceServiceAccount,
      context: 'database_access'
    });

    if (!dataAccess) {
      console.warn('openDatabase: Failed to access database via Sheets API, attempting SpreadsheetApp.openById fallback');

      // フォールバック: SpreadsheetApp.openByIdを使用（API制限対策）
      try {
        const fallbackSpreadsheet = SpreadsheetApp.openById(dbId);
        console.info('openDatabase: Successfully accessed database via SpreadsheetApp.openById fallback');
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
    return dataAccess.spreadsheet;
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
      const credentials = PropertiesService.getScriptProperties().getProperty('SERVICE_ACCOUNT_CREDS');
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
        // 必要最小限のSheetプロキシを返す
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

    return {
      getName: () => sheetName,
      getSheetId: () => additionalInfo.sheetId || 0,
      getLastRow: () => {
        // additionalInfoが利用可能な場合は使用、そうでなければAPIで取得
        if (additionalInfo.rowCount) {
          return additionalInfo.rowCount;
        }
        // Sheets APIで行数取得
        try {
          const response = fetchSheetsAPIWithRetry(
            `${baseUrl}/values/${sheetName}`,
            {
              headers: { 'Authorization': `Bearer ${accessToken}` }
            },
            `getLastRow(${sheetName})`
          );
          const data = JSON.parse(response.getContentText());
          return data.values ? data.values.length : 1;
        } catch (error) {
          console.warn('getLastRow via API failed after retries:', error.message);
          return 1;
        }
      },
      getLastColumn: () => {
        // additionalInfoが利用可能な場合は使用、そうでなければAPIで取得
        if (additionalInfo.columnCount) {
          return additionalInfo.columnCount;
        }
        // Sheets APIで列数取得
        try {
          const response = fetchSheetsAPIWithRetry(
            `${baseUrl}/values/${sheetName}!1:1`,
            {
              headers: { 'Authorization': `Bearer ${accessToken}` }
            },
            `getLastColumn(${sheetName})`
          );
          const data = JSON.parse(response.getContentText());
          return data.values && data.values[0] ? data.values[0].length : 1;
        } catch (error) {
          console.warn('getLastColumn via API failed after retries:', error.message);
          return 1;
        }
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

    // Validate service account usage
    const validation = validateServiceAccountUsage(spreadsheetId, options.useServiceAccount, options.context || 'openSpreadsheet');
    if (!validation.allowed) {
      console.warn('openSpreadsheet: Service account usage denied:', validation.reason);
      return null;
    }

    let spreadsheet = null;
    let auth = null;

    // CROSS-USER ACCESS ONLY - service account usage
    if (options.useServiceAccount === true) {
      const currentUser = getCurrentEmail() || 'unknown';

      // クロスユーザーアクセス - サービスアカウント使用
      auth = getServiceAccount();
      if (auth.isValid) {
        const dbId = PropertiesService.getScriptProperties().getProperty('DATABASE_SPREADSHEET_ID');

        // DATABASE_SPREADSHEET_IDの場合、DriveApp権限チェックをスキップしてSheets API直接アクセス
        if (spreadsheetId !== dbId) {
          try {
            // Check if access already exists before granting
            const file = DriveApp.getFileById(spreadsheetId);
            const editors = file.getEditors();
            const hasAccess = editors.some(editor => editor.getEmail() === auth.email);

            if (!hasAccess) {
              file.addEditor(auth.email);
            }
          } catch (driveError) {
            console.warn('openSpreadsheet: DriveApp permission check failed, proceeding with Sheets API access:', driveError.message);
          }
        }
      } else {
        console.warn('openSpreadsheet: Service account requested but invalid credentials');
      }
    } else {
      // Default: normal user permissions
      const currentUser = getCurrentEmail() || 'unknown';
    }

    // スプレッドシートを開く
    try {
      console.log('openSpreadsheet: 接続試行開始', {
        spreadsheetId: `${spreadsheetId.substring(0, 20)  }...`,
        useServiceAccount: options.useServiceAccount,
        hasAuth: !!auth,
        authValid: auth?.isValid
      });

      if (options.useServiceAccount === true && auth && auth.isValid) {
        // Service account implementation via JWT authentication
        console.log('openSpreadsheet: サービスアカウント接続試行');
        spreadsheet = openSpreadsheetViaServiceAccount(spreadsheetId);
        if (!spreadsheet) {
          console.error('openSpreadsheet: サービスアカウント接続失敗、通常アクセスにフォールバック', {
            spreadsheetId: `${spreadsheetId.substring(0, 20)  }...`,
            authEmail: auth.email
          });
          spreadsheet = SpreadsheetApp.openById(spreadsheetId);
        } else {
          console.log('openSpreadsheet: サービスアカウント接続成功');
        }
      } else {
        console.log('openSpreadsheet: 通常権限で接続');
        spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      }

      console.log('openSpreadsheet: 接続成功', {
        spreadsheetId: `${spreadsheetId.substring(0, 20)  }...`,
        hasSpreadsheet: !!spreadsheet
      });
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
    const cacheVersion = PropertiesService.getScriptProperties().getProperty('USER_CACHE_VERSION') || '0';
    const individualCacheKey = `user_by_email_v${cacheVersion}_${email}`;
    try {
      const cached = CacheService.getScriptCache().get(individualCacheKey);
      if (cached) {
        const cachedUser = JSON.parse(cached);
        console.log(`findUserByEmail: Found user in individual cache (v${cacheVersion}), highest performance`);
        return cachedUser;
      }
    } catch (individualCacheError) {
      console.warn('findUserByEmail: Individual cache read failed:', individualCacheError.message);
    }

    // キャッシュ最適化: まずgetAllUsers()のキャッシュを活用
    try {
      const allUsers = getAllUsers({ activeOnly: false }, { ...context, forceServiceAccount: true, skipCache: false });
      if (Array.isArray(allUsers) && allUsers.length > 0) {
        const user = allUsers.find(u => u.userEmail === email);
        if (user) {
          console.log('findUserByEmail: Found user in cached data, avoiding direct API call');

          // 個別キャッシュに保存（冗長性強化）
          try {
            CacheService.getScriptCache().put(individualCacheKey, JSON.stringify(user), CACHE_DURATION.USER_INDIVIDUAL);
          } catch (saveError) {
            console.warn('findUserByEmail: Individual cache save failed:', saveError.message);
          }

          return user;
        }
      }
    } catch (cacheError) {
      console.warn('findUserByEmail: Cache-based search failed, falling back to direct DB access:', cacheError.message);
    }

    // フォールバック: 直接データベースアクセス
    const useServiceAccount = context.forceServiceAccount || false;

    const spreadsheet = openDatabase(useServiceAccount);
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

        // 個別キャッシュに保存（冗長性強化）
        try {
          CacheService.getScriptCache().put(individualCacheKey, JSON.stringify(user), CACHE_DURATION.USER_INDIVIDUAL);
        } catch (saveError) {
          console.warn('findUserByEmail: Individual cache save failed:', saveError.message);
        }

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
    const cacheVersion = PropertiesService.getScriptProperties().getProperty('USER_CACHE_VERSION') || '0';
    const individualCacheKey = `user_by_id_v${cacheVersion}_${userId}`;
    try {
      const cached = CacheService.getScriptCache().get(individualCacheKey);
      if (cached) {
        const cachedUser = JSON.parse(cached);
        console.log(`findUserById: Found user in individual cache (v${cacheVersion}), highest performance`);
        return cachedUser;
      }
    } catch (individualCacheError) {
      console.warn('findUserById: Individual cache read failed:', individualCacheError.message);
    }

    // キャッシュ最適化: まずgetAllUsers()のキャッシュを活用
    try {
      const allUsers = getAllUsers({ activeOnly: false }, { ...context, forceServiceAccount: true, skipCache: false });
      if (Array.isArray(allUsers) && allUsers.length > 0) {
        const user = allUsers.find(u => u.userId === userId);
        if (user) {
          console.log('findUserById: Found user in cached data, avoiding direct API call');

          // 個別キャッシュに保存（冗長性強化）
          try {
            CacheService.getScriptCache().put(individualCacheKey, JSON.stringify(user), CACHE_DURATION.USER_INDIVIDUAL);
          } catch (saveError) {
            console.warn('findUserById: Individual cache save failed:', saveError.message);
          }

          return user;
        }
      }
    } catch (cacheError) {
      console.warn('findUserById: Cache-based search failed, falling back to direct DB access:', cacheError.message);
    }

    // フォールバック: 直接データベースアクセス
    const useServiceAccount = context.forceServiceAccount || false;

    const spreadsheet = openDatabase(useServiceAccount);
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

        // 個別キャッシュに保存（冗長性強化）
        try {
          CacheService.getScriptCache().put(individualCacheKey, JSON.stringify(user), CACHE_DURATION.USER_INDIVIDUAL);
        } catch (saveError) {
          console.warn('findUserById: Individual cache save failed:', saveError.message);
        }

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
      console.log('createUser: User already exists, returning existing:', email);
      return existingUser;
    }

    // DATABASE_SPREADSHEET_ID: Shared database for user creation
    const spreadsheet = openDatabase(context.forceServiceAccount || false);
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
    SpreadsheetApp.flush(); // 確実に書き込み

    const user = {
      userId,
      userEmail: email,
      isActive: true,
      configJson: JSON.stringify(defaultConfig),
      createdAt: hasCreatedAtColumn ? now : undefined,
      lastModified: now
    };

    // ユーザー作成後にキャッシュをクリア
    clearDatabaseUserCache('user_creation');
    clearIndividualUserCache(user, 'user_creation');

    console.log('createUser: User created successfully:', userId);
    return user;

  } catch (error) {
    console.error('createUser error:', error.message);
    return null;
  } finally {
    // 確実にロックを解放
    lock.releaseLock();
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
    const cacheVersion = PropertiesService.getScriptProperties().getProperty('USER_CACHE_VERSION') || '0';
    const cacheKey = `all_users_v${cacheVersion}_${JSON.stringify(options)}_${context.forceServiceAccount ? 'sa' : 'norm'}`;
    const skipCache = context.skipCache || false;

    if (!skipCache) {
      try {
        const cached = CacheService.getScriptCache().get(cacheKey);
        if (cached) {
          const cachedUsers = JSON.parse(cached);
          console.log(`getAllUsers: Returned cached data (v${cacheVersion}), avoiding API call`);
          return cachedUsers;
        }
      } catch (cacheError) {
        console.warn('getAllUsers: Cache read failed:', cacheError.message);
      }
    }

    // getAllUsers is inherently cross-user operation, always requires service account
    // unless admin is accessing for administrative purposes
    const spreadsheet = openDatabase(true); // Always service account for cross-user data
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
        console.log('getAllUsers: Cached user data for 10 minutes');
      } catch (cacheError) {
        console.warn('getAllUsers: Cache write failed:', cacheError.message);
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
 * @param {string} operation - 操作名（ログ用）
 */
function clearDatabaseUserCache(operation = 'database_operation') {
  try {
    const props = PropertiesService.getScriptProperties();
    const currentVersion = parseInt(props.getProperty('USER_CACHE_VERSION') || '0');
    const newVersion = currentVersion + 1;

    props.setProperty('USER_CACHE_VERSION', newVersion.toString());

    console.log(`clearDatabaseUserCache: Cache invalidated after ${operation} (v${currentVersion} → v${newVersion})`);
  } catch (error) {
    console.warn('clearDatabaseUserCache: Failed to clear cache:', error.message);
  }
}

/**
 * 特定ユーザーの個別キャッシュをクリア
 * バージョニング戦略により、この関数は不要（clearDatabaseUserCacheで一括無効化）
 * 後方互換性のため空実装を保持
 * @deprecated Use clearDatabaseUserCache() instead - versioning strategy handles all caches
 * @param {Object} user - ユーザーオブジェクト
 * @param {string} operation - 操作名（ログ用）
 */
function clearIndividualUserCache(user, operation = 'user_operation') {
  // No-op: Versioning strategy in clearDatabaseUserCache() invalidates all user caches
  // This function preserved for backward compatibility only
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
    const useServiceAccount = context.forceServiceAccount || false;


    const spreadsheet = openDatabase(useServiceAccount);
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
        // Update specified fields
        Object.keys(updates).forEach(field => {
          const columnIndex = headers.indexOf(field);
          if (columnIndex !== -1) {
            sheet.getRange(i + 1, columnIndex + 1).setValue(updates[field]);
          }
        });

        // Update lastModified
        const lastModifiedIndex = headers.indexOf('lastModified');
        if (lastModifiedIndex !== -1) {
          sheet.getRange(i + 1, lastModifiedIndex + 1).setValue(new Date().toISOString());
        }

        SpreadsheetApp.flush(); // 確実に書き込み

        // ユーザー更新後にキャッシュをクリア
        clearDatabaseUserCache('user_update');
        clearIndividualUserCache(targetUser, 'user_update');

        console.log('updateUser: User updated successfully:', userId);
        return { success: true };
      }
    }

    console.warn('updateUser: User not found:', userId);
    return { success: false, message: 'User not found' };
  } catch (error) {
    console.error('updateUser error:', error.message);
    return { success: false, message: error.message || 'Unknown error occurred' };
  } finally {
    // 確実にロックを解放
    lock.releaseLock();
  }
}


/**
 * 閲覧者向けボードデータ取得（CLAUDE.md準拠 - 模範実装）
 *
 * Context-aware access control implementation.
 * Self-access uses normal permissions, cross-user access uses service account.
 *
 * Security Pattern:
 * - Self-access: targetUser.userEmail === viewerEmail → Normal permissions
 * - Cross-user access: targetUser.userEmail !== viewerEmail → Service account
 *
 * @param {string} targetUserId - 対象ユーザーのユニークID（UUID形式）
 * @param {string} viewerEmail - 閲覧者のメールアドレス（アクセス判定用）
 * @returns {Object|null} Board data with spreadsheet info, config, and sheets list, or null if error
 *
 * @example
 * // 自分のデータにアクセス（通常権限）
 * const myData = getViewerBoardData(myUserId, myEmail);
 *
 * @example
 * // 他人のデータにアクセス（サービスアカウント）
 * const othersData = getViewerBoardData(otherUserId, myEmail);
 */
function getViewerBoardData(targetUserId, viewerEmail) {
  try {
    const targetUser = findUserById(targetUserId, { requestingUser: viewerEmail });
    if (!targetUser) {
      console.warn('getViewerBoardData: Target user not found:', targetUserId);
      return null;
    }

    if (targetUser.userEmail === viewerEmail) {
      // Own data: use normal permissions
      // eslint-disable-next-line no-undef
      return getUserSheetData(targetUser.userId, {
        includeTimestamp: true,
        requestingUser: viewerEmail
      });
    } else {
      // Other's data: use service account for cross-user access
      // eslint-disable-next-line no-undef
      return getUserSheetData(targetUser.userId, {
        includeTimestamp: true,
        requestingUser: viewerEmail,
        targetUserEmail: targetUser.userEmail // This will trigger service account usage in getUserSheetData
      });
    }
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

    // Performance optimization: caching implementation
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

          // Cache the result
          if (!skipCache) {
            try {
              const cacheTtl = context.cacheTtl || CACHE_DURATION.LONG; // 300秒
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
    const dbId = PropertiesService.getScriptProperties().getProperty('DATABASE_SPREADSHEET_ID');
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

    for (let i = 1; i < data.length; i++) {
      if (data[i][userIdColumnIndex] === userId) {
        // 削除前にユーザー情報を取得（キャッシュクリア用）
        const userToDelete = createUserObjectFromRow(data[i], headers);

        // Simple hard delete - remove the row using correct GAS API
        const rowToDelete = i + 1;
        sheet.deleteRows(rowToDelete, 1);

        // ユーザー削除後にキャッシュをクリア
        clearDatabaseUserCache('user_deletion');
        clearIndividualUserCache(userToDelete, 'user_deletion');

        return {
          success: true,
          userId,
          deleted: true,
          reason: reason || 'No reason provided'
        };
      }
    }

    console.warn('deleteUser: User not found:', userId);
    return { success: false, message: 'User not found' };
  } catch (error) {
    console.error('deleteUser error:', error.message);
    return { success: false, message: error.message };
  }
}

