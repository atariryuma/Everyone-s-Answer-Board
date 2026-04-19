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

  const cache = CacheService.getScriptCache();
  const CIRCUIT_BREAKER_KEY = 'circuit_breaker_state';
  const cachedState = cache.get(CIRCUIT_BREAKER_KEY);
  let circuitState = cachedState ? JSON.parse(cachedState) : { consecutiveErrors: 0, circuitOpenUntil: 0 };

  const now = Date.now();
  if (circuitState.circuitOpenUntil > now) {
    const waitTime = Math.round((circuitState.circuitOpenUntil - now) / 1000);
    throw new Error(`Circuit breaker open: API calls paused for ${waitTime}s to allow quota recovery`);
  }

  return executeWithRetry(
    () => {
      const response = UrlFetchApp.fetch(url, options);

      if (response.getResponseCode() === 429) {
        const backoffTime = Math.min(15000 + (retryCount * 15000), 60000);
        console.warn(`⚠️ 429 Quota exceeded for ${operationName || 'Sheets API'}, waiting ${backoffTime}ms (retry: ${retryCount})`);

        circuitState.consecutiveErrors++;

        if (circuitState.consecutiveErrors >= 7) {
          // 段階的バックオフ: 7回=30s, 10回=60s, 15回+=120s
          const lockoutMs = circuitState.consecutiveErrors >= 15 ? 120000
            : circuitState.consecutiveErrors >= 10 ? 60000 : 30000;
          circuitState.circuitOpenUntil = now + lockoutMs;
          console.error(`Circuit breaker activated: ${circuitState.consecutiveErrors} errors. Paused for ${lockoutMs / 1000}s.`);
        }

        cache.put(CIRCUIT_BREAKER_KEY, JSON.stringify(circuitState), 120);

        Utilities.sleep(backoffTime);
        retryCount++;
        throw new Error('Quota exceeded (429), retry with adaptive backoff');
      }

      if (response.getResponseCode() !== 200) {
        const errorText = response.getContentText();
        throw new Error(`API returned ${response.getResponseCode()}: ${errorText}`);
      }

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

    const requiredFields = ['client_email', 'private_key', 'type'];
    const missingFields = requiredFields.filter(field => !serviceAccount[field]);

    if (missingFields.length > 0) {
      console.warn('getServiceAccount: Missing required fields');
      return { isValid: false, error: 'Invalid credentials' };
    }

    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(serviceAccount.client_email)) {
      console.warn('getServiceAccount: Invalid email format');
      return { isValid: false, error: 'Invalid email format' };
    }

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

    if (!useServiceAccount) {
      return { allowed: true, reason: 'Normal permissions requested' };
    }

    if (spreadsheetId === dbId) {
      return { allowed: true, reason: 'DATABASE_SPREADSHEET_ID is shared resource' };
    }

    if (isAdministrator(currentEmail)) {
      return { allowed: true, reason: 'Admin privileges' };
    }

    // ✅ SECURITY GATE: 非管理者は公開済みボードのみアクセス許可
    try {
      const cacheKey = `sa_validation_${spreadsheetId}`;
      const cached = CacheService.getScriptCache().get(cacheKey);

      if (cached) {
        const validation = JSON.parse(cached);
        return {
          allowed: validation.isPublished,
          reason: validation.isPublished ? 'Public board access (cached)' : 'Board not published (cached)'
        };
      }

      const targetUser = findUserBySpreadsheetId(spreadsheetId, { skipCache: true });

      if (!targetUser) {
        console.warn('SA_VALIDATION: Target user not found for spreadsheet:', spreadsheetId.substring(0, 8));
        return { allowed: false, reason: 'Target user not found' };
      }

      const configResult = getUserConfig(targetUser.userId);
      const isPublished = configResult.success && configResult.config.isPublished;

      try {
        CacheService.getScriptCache().put(cacheKey, JSON.stringify({ isPublished }), 60);
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

    const forceServiceAccount = true;

    const dataAccess = openSpreadsheet(dbId, {
      useServiceAccount: forceServiceAccount,
      context: 'database_access'
    });

    if (!dataAccess) {
      console.warn('openDatabase: Failed to access database via Sheets API, attempting SpreadsheetApp.openById fallback');

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
  try {
    if (!spreadsheetId || typeof spreadsheetId !== 'string') {
      console.warn('openSpreadsheet: Invalid spreadsheet ID');
      return null;
    }

    // Why: サービスアカウントはDATABASE_SPREADSHEETのみに限定。
    // ユーザーのシートに対するSA使用はセキュリティ上許可しない。
    const databaseId = getCachedProperty('DATABASE_SPREADSHEET_ID');
    const isDatabaseAccess = spreadsheetId === databaseId;
    const effectiveUseServiceAccount = isDatabaseAccess && options.useServiceAccount === true;

    const validation = validateServiceAccountUsage(spreadsheetId, effectiveUseServiceAccount, options.context || 'openSpreadsheet');
    if (!validation.allowed) {
      console.warn('openSpreadsheet: Service account usage denied:', validation.reason);
      return null;
    }

    let spreadsheet = null;
    let auth = null;

    if (effectiveUseServiceAccount) {
      auth = getServiceAccount();
      if (!auth.isValid) {
        console.warn('openSpreadsheet: Service account requested but invalid credentials');
      }
    }

    try {
      if (effectiveUseServiceAccount && auth && auth.isValid) {
        spreadsheet = openSpreadsheetViaServiceAccount(spreadsheetId);
        if (!spreadsheet) {
          console.error('openSpreadsheet: サービスアカウント接続失敗、通常アクセスにフォールバック', {
            spreadsheetId: `${spreadsheetId.substring(0, 20)}...`,
            authEmail: auth.email
          });
          spreadsheet = SpreadsheetApp.openById(spreadsheetId);
        }
      } else {
        spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      }
    } catch (openError) {
      console.error('openSpreadsheet: スプレッドシート接続失敗', {
        spreadsheetId: `${spreadsheetId.substring(0, 20)}...`,
        error: openError.message,
        stack: openError.stack ? `${openError.stack.substring(0, 200)}...` : 'No stack trace',
        useServiceAccount: options.useServiceAccount,
        hasAuth: !!auth
      });
      return null;
    }

    return {
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
  } catch (error) {
    console.error('openSpreadsheet error:', error.message);
    return null;
  }
}

// ---------------------------------------------------------------------
// サービスアカウント経路のヘルパー群
// openSpreadsheet内部からmodule-levelへ抽出（クロージャ未使用のため純粋な字句移動）。
// デバッグ時のスタックトレース可読性とテスト可能性の向上が目的。
// ---------------------------------------------------------------------

function openSpreadsheetViaServiceAccount(sheetId) {
  try {
    const credentials = getCachedProperty('SERVICE_ACCOUNT_CREDS');
    if (!credentials) return null;

    const serviceAccount = JSON.parse(credentials);

    const now = Math.floor(Date.now() / 1000);
    const payload = {
      iss: serviceAccount.client_email,
      scope: 'https://www.googleapis.com/auth/spreadsheets',
      aud: 'https://oauth2.googleapis.com/token',
      exp: now + 3600,
      iat: now
    };

    const header = { alg: 'RS256', typ: 'JWT' };
    const jwt = createServiceAccountJWT(header, payload, serviceAccount.private_key);

    const tokenResponse = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      payload: `grant_type=urn:ietf:params:oauth:grant-type:jwt-bearer&assertion=${jwt}`
    });

    const tokenData = JSON.parse(tokenResponse.getContentText());
    if (!tokenData.access_token) return null;

    return createServiceAccountSpreadsheetProxy(sheetId, tokenData.access_token);

  } catch (error) {
    console.error('openSpreadsheetViaServiceAccount error:', error.message);
    return null;
  }
}

function createServiceAccountJWT(header, payload, privateKey) {
  const headerB64 = Utilities.base64EncodeWebSafe(JSON.stringify(header)).replace(/=/g, '');
  const payloadB64 = Utilities.base64EncodeWebSafe(JSON.stringify(payload)).replace(/=/g, '');
  const signatureInput = `${headerB64}.${payloadB64}`;

  const signature = Utilities.computeRsaSha256Signature(signatureInput, privateKey);
  const signatureB64 = Utilities.base64EncodeWebSafe(signature).replace(/=/g, '');

  return `${signatureInput}.${signatureB64}`;
}

function createServiceAccountSpreadsheetProxy(sheetId, accessToken) {
  return {
    getId: () => sheetId,
    getName: () => {
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
      return createServiceAccountSheetProxy(sheetId, sheetName, accessToken);
    },
    getSheets: () => {
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

function createServiceAccountSheetProxy(sheetId, sheetName, accessToken, additionalInfo = {}) {
  const baseUrl = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}`;
  let cachedDimensions = null;

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
      if (additionalInfo.rowCount) return additionalInfo.rowCount;
      return fetchDimensionsOnce().rowCount;
    },
    getLastColumn: () => {
      if (additionalInfo.columnCount) return additionalInfo.columnCount;
      return fetchDimensionsOnce().columnCount;
    },
    getRange: (row, col, numRows, numCols) => {
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
            const payload = { values: [[value]] };
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
            const payload = { values };
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
            const payload = { values };
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
      try {
        const payload = { values: [rowData] };
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




/**
 * ユーザー検索の呼び出し元メールアドレスを解決
 * @param {Object} context - アクセスコンテキスト
 * @returns {string|null} リクエストユーザーのメール
 */
function resolveRequestingUser(context = {}) {
  if (context.requestingUser && typeof context.requestingUser === 'string' && context.requestingUser.trim()) {
    return context.requestingUser.trim();
  }
  try {
    return getCurrentEmail();
  } catch (error) {
    console.error('resolveRequestingUser: getCurrentEmail failed:', error.message);
    return null;
  }
}

/**
 * 対象ユーザーのボード公開状態を判定
 * @param {Object} targetUser - 対象ユーザー
 * @returns {boolean} 公開中かどうか
 */
function isUserBoardPublished(targetUser) {
  if (!targetUser) {
    return false;
  }

  try {
    if (targetUser.configJson && typeof targetUser.configJson === 'string') {
      const parsed = JSON.parse(targetUser.configJson);
      return Boolean(parsed && parsed.isPublished);
    }
  } catch (parseError) {
    console.warn('isUserBoardPublished: Failed to parse configJson:', parseError.message);
  }

  if (!targetUser.userId) {
    return false;
  }

  try {
    const configResult = getUserConfig(targetUser.userId, targetUser);
    return Boolean(configResult && configResult.success && configResult.config && configResult.config.isPublished);
  } catch (configError) {
    console.warn('isUserBoardPublished: Failed to resolve config:', configError.message);
    return false;
  }
}

/**
 * ユーザー検索のアクセス可否を判定
 * @param {Object} targetUser - 対象ユーザー
 * @param {Object} context - アクセスコンテキスト
 * @returns {boolean} アクセス許可されるか
 */
function canAccessTargetUser(targetUser, context = {}) {
  if (!targetUser) {
    return false;
  }

  if (context.skipAccessCheck === true) {
    return true;
  }

  const requestingUser = resolveRequestingUser(context);
  if (!requestingUser) {
    return false;
  }

  let isAdmin = false;
  if (context.preloadedAuth && typeof context.preloadedAuth.isAdmin === 'boolean') {
    isAdmin = context.preloadedAuth.isAdmin;
  } else {
    isAdmin = isAdministrator(requestingUser);
  }

  if (isAdmin) {
    return true;
  }

  if (targetUser.userEmail === requestingUser) {
    return true;
  }

  if (context.allowPublishedRead === true) {
    return isUserBoardPublished(targetUser);
  }

  return false;
}

/**
 * ユーザー検索結果にアクセス制御を適用
 * @param {Object|null} user - 検索結果ユーザー
 * @param {Object} context - アクセスコンテキスト
 * @param {string} operation - 操作名
 * @returns {Object|null} 許可済みユーザーまたはnull
 */
function applyUserAccessControl(user, context = {}, operation = 'user_lookup') {
  if (!user) {
    return null;
  }

  if (canAccessTargetUser(user, context)) {
    return user;
  }

  const requestingUser = resolveRequestingUser(context);
  const maskedRequester = requestingUser ? `${requestingUser.split('@')[0]}@***` : 'N/A';
  const maskedTarget = user.userEmail ? `${String(user.userEmail).split('@')[0]}@***` : 'unknown';
  console.warn(`${operation}: Access denied`, {
    requestingUser: maskedRequester,
    target: maskedTarget
  });
  return null;
}

/**
 * メールアドレスでユーザーを検索（CLAUDE.md準拠 - Editor→Admin共有DB）
 * @param {string} email - ユーザーのメールアドレス
 * @param {Object} context - アクセスコンテキスト
 * @param {boolean} context.forceServiceAccount - サービスアカウント強制使用
 * @param {string} context.requestingUser - リクエストユーザー（デバッグ用）
 * @returns {Object|null} User object
 */
function findUserByEmail(email, context = {}) {
  if (!email || !validateEmail(email).isValid) return null;
  return findUserByField('userEmail', email, { ...context, cacheKeyPrefix: 'user_by_email', label: 'findUserByEmail' });
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
  if (!userId) return null;
  return findUserByField('userId', userId, { ...context, cacheKeyPrefix: 'user_by_id', label: 'findUserById' });
}

/**
 * GoogleアカウントID（不変ID）でユーザーを検索
 * メールアドレス変更時のフォールバック検索に使用
 * @param {string} googleId - Googleアカウントのsub値
 * @param {Object} context - アクセスコンテキスト
 * @returns {Object|null} User object
 */
function findUserByGoogleId(googleId, context = {}) {
  if (!googleId) return null;
  // Note: historically this path did not populate an individual cache entry, so
  // we explicitly opt out via cacheKeyPrefix=null to preserve behavior.
  return findUserByField('googleId', googleId, { ...context, cacheKeyPrefix: null, label: 'findUserByGoogleId' });
}

/**
 * findUserBy*系の共通実装。
 * 探索は (1) individual cache → (2) getAllUsers キャッシュ → (3) 直接DB の順。
 * @param {string} fieldName - 検索するユーザーフィールド ('userEmail' | 'userId' | 'googleId')
 * @param {*} fieldValue - 一致させる値
 * @param {Object} context - requestingUser / preloadedAuth / forceServiceAccount 等に加え、
 *                          内部用の cacheKeyPrefix (null で個別キャッシュ無効) と label を受け取る
 * @returns {Object|null} アクセス制御適用済みのユーザー、または null
 */
function findUserByField(fieldName, fieldValue, context = {}) {
  const { cacheKeyPrefix = null, label = `findUserBy_${fieldName}` } = context;

  try {
    let individualCacheKey = null;
    if (cacheKeyPrefix) {
      const cacheVersion = getCachedProperty('USER_CACHE_VERSION') || '0';
      individualCacheKey = `${cacheKeyPrefix}_v${cacheVersion}_${fieldValue}`;
      try {
        const cached = CacheService.getScriptCache().get(individualCacheKey);
        if (cached) {
          const cachedUser = JSON.parse(cached);
          return applyUserAccessControl(cachedUser, context, label);
        }
      } catch (individualCacheError) {
        console.error(`${label}: Individual cache read failed:`, individualCacheError.message);
      }
    }

    try {
      const allUsers = getAllUsers(
        { activeOnly: false },
        { ...context, forceServiceAccount: true, skipCache: false }
      );
      if (Array.isArray(allUsers) && allUsers.length > 0) {
        const user = allUsers.find((u) => u[fieldName] === fieldValue);
        if (user) {
          const allowedUser = applyUserAccessControl(user, context, label);
          if (!allowedUser) return null;
          if (individualCacheKey) {
            saveToCacheWithSizeCheck(individualCacheKey, allowedUser, CACHE_DURATION.USER_INDIVIDUAL);
          }
          return allowedUser;
        }
      }
    } catch (cacheError) {
      console.error(`${label}: Cache-based search failed, falling back to direct DB access:`, cacheError.message);
    }

    const spreadsheet = openDatabase();
    if (!spreadsheet) {
      console.warn(`${label}: Database access failed`);
      return null;
    }
    const sheet = spreadsheet.getSheetByName('users');
    if (!sheet) {
      console.warn(`${label}: Users sheet not found`);
      return null;
    }
    const data = sheet.getDataRange().getValues();
    if (data.length === 0) return null;

    const [headers] = data;
    const columnIndex = headers.indexOf(fieldName);
    if (columnIndex === -1) {
      console.warn(`${label}: ${fieldName} column not found`);
      return null;
    }

    for (let i = 1; i < data.length; i++) {
      if (data[i][columnIndex] === fieldValue) {
        const user = createUserObjectFromRow(data[i], headers);
        const allowedUser = applyUserAccessControl(user, context, label);
        if (!allowedUser) return null;
        if (individualCacheKey) {
          saveToCacheWithSizeCheck(individualCacheKey, allowedUser, CACHE_DURATION.USER_INDIVIDUAL);
        }
        return allowedUser;
      }
    }

    return null;
  } catch (error) {
    console.error(`${label} error:`, error.message);
    return null;
  }
}

/**
 * 新しいユーザーを作成（CLAUDE.md準拠 - Context-Aware）
 * @param {string} email - ユーザーのメールアドレス
 * @param {Object} initialConfig - 初期設定
 * @param {Object} context - アクセスコンテキスト
 * @param {string} [context.googleId] - Googleアカウント不変ID
 * @returns {Object|null} Created user object
 */
function createUser(email, initialConfig = {}, context = {}) {
  const lock = LockService.getScriptLock();

  try {
    if (!email || !validateEmail(email).isValid) {
      console.warn('createUser: Invalid email provided:', typeof email, email);
      return null;
    }

    if (!lock.tryLock(10000)) { // 10秒待機
      console.warn('createUser: Lock timeout - concurrent user creation detected');
      return null;
    }

    const currentEmail = getCurrentEmail();

    const existingUser = findUserByEmail(email, {
      requestingUser: currentEmail
    });
    if (existingUser) {
      return existingUser;
    }

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

    const data = sheet.getDataRange().getValues();
    const [headers] = data;
    const hasCreatedAtColumn = headers.indexOf('createdAt') !== -1;
    const hasGoogleIdColumn = headers.indexOf('googleId') !== -1;
    const googleId = context.googleId || '';

    const newUserRow = new Array(headers.length).fill('');
    const setCol = (name, value) => {
      const idx = headers.indexOf(name);
      if (idx !== -1) newUserRow[idx] = value;
    };
    setCol('userId', userId);
    setCol('userEmail', email);
    if (hasGoogleIdColumn) setCol('googleId', googleId);
    setCol('isActive', true);
    setCol('configJson', JSON.stringify(defaultConfig));
    if (hasCreatedAtColumn) setCol('createdAt', now);
    setCol('lastModified', now);

    if (sheet.getLastRow() >= 10000) {
      console.error('createUser: Users sheet at capacity');
      return null;
    }

    sheet.appendRow(newUserRow);

    const user = {
      userId,
      userEmail: email,
      googleId: hasGoogleIdColumn ? googleId : undefined,
      isActive: true,
      configJson: JSON.stringify(defaultConfig),
      createdAt: hasCreatedAtColumn ? now : undefined,
      lastModified: now
    };

    clearDatabaseUserCache();

    return user;

  } catch (error) {
    console.error('createUser error:', error.message);
    return null;
  } finally {
    try {
      lock.releaseLock();
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

    if (!skipCache) {
      // Why: 100+ユーザでJSONが100KB超えるとCacheService.putが黙って失敗する。
      //      saveToCacheWithSizeCheckでサイズ超過を検出しログに残す。
      saveToCacheWithSizeCheck(cacheKey, users, CACHE_DURATION.DATABASE_LONG);
    }

    return users;
  } catch (error) {
    console.error('getAllUsers error:', error.message);
    return [];
  }
}




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
    'googleId': 'googleId',
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
  const lock = LockService.getScriptLock();

  try {
    if (!lock.tryLock(5000)) { // 5秒待機
      console.warn('updateUser: Lock timeout - concurrent update detected');
      return { success: false, message: 'Concurrent update in progress. Please retry.' };
    }


    const requestingUser = context.requestingUser || getCurrentEmail();
    const targetUser = findUserById(userId, {
      ...context,
      requestingUser
    });

    if (!targetUser) {
      console.warn('updateUser: Target user not found:', userId);
      return { success: false, message: 'User not found' };
    }

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
        const updateCells = [];

        Object.keys(updates).forEach(field => {
          const columnIndex = headers.indexOf(field);
          if (columnIndex !== -1) {
            updateCells.push({ col: columnIndex + 1, value: updates[field] });
          }
        });

        const lastModifiedIndex = headers.indexOf('lastModified');
        if (lastModifiedIndex !== -1) {
          updateCells.push({ col: lastModifiedIndex + 1, value: new Date().toISOString() });
        }

        if (updateCells.length > 0) {
          const cols = updateCells.map(c => c.col);
          const minCol = Math.min(...cols);
          const maxCol = Math.max(...cols);
          const colSpan = maxCol - minCol + 1;

          const rangeToUpdate = sheet.getRange(i + 1, minCol, 1, colSpan);
          const [currentRowData] = rangeToUpdate.getValues();

          updateCells.forEach(({ col, value }) => {
            currentRowData[col - minCol] = value;
          });

          rangeToUpdate.setValues([currentRowData]);
        }

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
    try {
      lock.releaseLock();
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
    const targetUser = findUserById(targetUserId, {
      requestingUser: viewerEmail,
      allowPublishedRead: true
    });
    if (!targetUser) {
      console.warn('getViewerBoardData: Target user not found:', targetUserId);
      return null;
    }

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


    const allUsers = getAllUsers({ activeOnly: false }, { ...context, forceServiceAccount: true, preloadedAuth: context.preloadedAuth });
    if (!Array.isArray(allUsers)) {
      console.warn('findUserBySpreadsheetId: Failed to get users list');
      return null;
    }

    for (const user of allUsers) {
      try {
        const configJson = user.configJson || '{}';
        const config = JSON.parse(configJson);

        if (config.spreadsheetId === spreadsheetId) {

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
        console.warn(`findUserBySpreadsheetId: Failed to parse configJSON for user ${user.userId}:`, parseError.message);
        continue;
      }
    }


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
    const currentEmail = getCurrentEmail();
    const isAdmin = isAdministrator(currentEmail);

    if (!isAdmin && !context.forceServiceAccount) {
      console.warn('deleteUser: Non-admin user attempting user deletion:', userId);
      return { success: false, message: 'Insufficient permissions for user deletion' };
    }


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

    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
      sheet.deleteRows(rowsToDelete[i], 1);
    }

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
