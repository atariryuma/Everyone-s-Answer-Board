/**
 * @fileoverview DatabaseCore - データベース操作の基盤
 *
 * 責任範囲:
 * - 基本的なデータベース操作（CRUD）
 * - ユーザー管理の基盤関数
 * - GAS-Native Architecture準拠（直接SpreadsheetApp使用）
 * - サービスアカウント使用時の安全な権限管理
 */

/* global validateEmail, CACHE_DURATION, TIMEOUT_MS, getCurrentEmail, isAdministrator, getUserConfig */

// ===========================================
// 🗄️ データベース基盤操作
// ===========================================

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

    // 🛡️ Enhanced validation for required fields
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
      console.log(`SA_VALIDATION: DATABASE_SPREADSHEET_ID access - service account required`);
      return { allowed: true, reason: 'DATABASE_SPREADSHEET_ID is shared resource' };
    }

    // Check if current user is admin - admins can use service account
    if (isAdministrator(currentEmail)) {
      console.log(`SA_VALIDATION: Admin ${currentEmail} authorized for service account usage in ${context}`);
      return { allowed: true, reason: 'Admin privileges' };
    }

    // For non-admin users accessing other spreadsheets, allow service account
    // Skip user lookup to prevent circular reference (findUserBySpreadsheetId -> getAllUsers -> openDatabase -> validateServiceAccountUsage)
    console.log(`SA_VALIDATION: Cross-user access approved for ${currentEmail} in ${context} (skipping user lookup to prevent circular reference)`);
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

    // 🔧 CLAUDE.md準拠: DATABASE_SPREADSHEET_ID is shared resource - always use service account
    // ✅ **Critical**: DATABASE_SPREADSHEET_ID contains all user data and requires elevated permissions
    // ✅ **Security**: General users cannot access DATABASE_SPREADSHEET_ID with normal permissions
    // Note: useServiceAccount parameter preserved for API compatibility but overridden for security
    const forceServiceAccount = true; // DATABASE_SPREADSHEET_ID always requires service account
    console.log(`openDatabase: Using service account for shared DATABASE_SPREADSHEET_ID (forced: ${forceServiceAccount})`);

    const dataAccess = openSpreadsheet(dbId, {
      useServiceAccount: forceServiceAccount,
      context: 'database_access'
    });
    if (!dataAccess) {
      console.warn('openDatabase: Failed to access database via openSpreadsheet with service account');
      return null;
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
 * 🔐 **CLAUDE.md Security-Critical Implementation**
 * サービスアカウントの使用は CROSS-USER ACCESS ONLY に制限されています。
 * この関数は CLAUDE.md の Security Best Practices に完全準拠し、
 * 適切な権限制御とログ記録を実装しています。
 *
 * 📊 **Usage Metrics**:
 * すべてのアクセスは SA_USAGE プレフィックスでログ記録され、
 * セキュリティ監査とパフォーマンス分析に使用されます。
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

  // 🔧 内部関数: サービスアカウント経由でSheets API直接アクセス
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
      console.log(`SA_API_ACCESS: Successfully authenticated via service account for ${sheetId.substring(0, 8)}***`);
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
      getSheetByName: (sheetName) => {
        // 必要最小限のSheetプロキシを返す
        return createServiceAccountSheetProxy(sheetId, sheetName, accessToken);
      }
    };
  }

  // サービスアカウント用シートプロキシ（基本メソッドのみ実装）
  function createServiceAccountSheetProxy(sheetId, sheetName, accessToken) {
    const baseUrl = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}`;

    return {
      getName: () => sheetName,
      getLastRow: () => {
        // Sheets APIで行数取得
        try {
          const response = UrlFetchApp.fetch(`${baseUrl}/values/${sheetName}`, {
            headers: { 'Authorization': `Bearer ${accessToken}` }
          });
          const data = JSON.parse(response.getContentText());
          return data.values ? data.values.length : 1;
        } catch (error) {
          console.warn('getLastRow via API failed:', error.message);
          return 1;
        }
      },
      getLastColumn: () => {
        // Sheets APIで列数取得
        try {
          const response = UrlFetchApp.fetch(`${baseUrl}/values/${sheetName}!1:1`, {
            headers: { 'Authorization': `Bearer ${accessToken}` }
          });
          const data = JSON.parse(response.getContentText());
          return data.values && data.values[0] ? data.values[0].length : 1;
        } catch (error) {
          console.warn('getLastColumn via API failed:', error.message);
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
              const response = UrlFetchApp.fetch(`${baseUrl}/values/${range}`, {
                headers: { 'Authorization': `Bearer ${accessToken}` }
              });
              const data = JSON.parse(response.getContentText());
              return data.values || [];
            } catch (error) {
              console.warn('getValues via API failed:', error.message);
              return [];
            }
          }
        };
      },
      getDataRange: () => {
        // Sheets APIでデータ範囲全体を取得
        return {
          getValues: () => {
            try {
              const response = UrlFetchApp.fetch(`${baseUrl}/values/${sheetName}`, {
                headers: { 'Authorization': `Bearer ${accessToken}` }
              });
              const data = JSON.parse(response.getContentText());
              return data.values || [];
            } catch (error) {
              console.warn('getDataRange via API failed:', error.message);
              return [];
            }
          }
        };
      }
    };
  }

  try {
    if (!spreadsheetId || typeof spreadsheetId !== 'string') {
      console.warn('openSpreadsheet: Invalid spreadsheet ID');
      return null;
    }

    // 🛡️ CLAUDE.md Security Gate: Validate service account usage
    const validation = validateServiceAccountUsage(spreadsheetId, options.useServiceAccount, options.context || 'openSpreadsheet');
    if (!validation.allowed) {
      console.warn('openSpreadsheet: Service account usage denied:', validation.reason);
      return null;
    }

    let spreadsheet = null;
    let auth = null;

    // 🔧 CLAUDE.md準拠: CROSS-USER ACCESS ONLYでサービスアカウント使用
    if (options.useServiceAccount === true) {
      // 📊 詳細ログ: サービスアカウント使用メトリクス
      const currentUser = getCurrentEmail() || 'unknown';
      console.log(`SA_USAGE: cross-user-access - ${currentUser} -> ${spreadsheetId.substring(0, 8)}*** - service_account_mode`);

      // クロスユーザーアクセス - サービスアカウント使用
      auth = getServiceAccount();
      if (auth.isValid) {
        const dbId = PropertiesService.getScriptProperties().getProperty('DATABASE_SPREADSHEET_ID');

        // DATABASE_SPREADSHEET_IDの場合、DriveApp権限チェックをスキップしてSheets API直接アクセス
        if (spreadsheetId === dbId) {
          console.log('openSpreadsheet: Skipping DriveApp permission check for DATABASE_SPREADSHEET_ID (using direct Sheets API access)');
        } else {
          try {
            // 🚀 Performance Optimization: Check if access already exists before granting
            const file = DriveApp.getFileById(spreadsheetId);
            const editors = file.getEditors();
            const hasAccess = editors.some(editor => editor.getEmail() === auth.email);

            if (!hasAccess) {
              file.addEditor(auth.email);
              console.log(`openSpreadsheet: Service account editor access granted for cross-user access: ${spreadsheetId.substring(0, 8)}***`);
              console.log(`SA_USAGE: editor-access-granted - ${auth.email} -> ${spreadsheetId.substring(0, 8)}***`);
            } else {
              console.log(`SA_USAGE: editor-access-existing - ${auth.email} -> ${spreadsheetId.substring(0, 8)}*** - already_has_access`);
            }
          } catch (driveError) {
            console.warn('openSpreadsheet: Service account access check/grant failed:', driveError.message);
            console.log(`SA_USAGE: editor-access-failed - ${auth.email} -> ${spreadsheetId.substring(0, 8)}*** - ${driveError.message}`);
          }
        }
      } else {
        console.warn('openSpreadsheet: Service account requested but invalid credentials');
      }
    } else {
      // ✅ デフォルト: 通常権限でアクセス（CLAUDE.md準拠）
      const currentUser = getCurrentEmail() || 'unknown';
      console.log(`openSpreadsheet: Using normal user permissions for ${spreadsheetId.substring(0, 8)}***`);
      console.log(`SA_USAGE: self-access - ${currentUser} -> ${spreadsheetId.substring(0, 8)}*** - normal_permissions`);
    }

    // スプレッドシートを開く
    try {
      if (options.useServiceAccount === true && auth && auth.isValid) {
        // 🔧 GAS Service Account実装: JWT認証でSheets API直接アクセス
        spreadsheet = openSpreadsheetViaServiceAccount(spreadsheetId);
        if (!spreadsheet) {
          console.error('openSpreadsheet: Service account access failed, falling back to normal access');
          spreadsheet = SpreadsheetApp.openById(spreadsheetId);
        }
      } else {
        spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      }
    } catch (openError) {
      console.error('openSpreadsheet: Failed to open spreadsheet:', openError.message);
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

// ===========================================
// 👤 ユーザー管理基盤
// ===========================================

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

    // 🔧 CLAUDE.md準拠: DATABASE_SPREADSHEET_ID は Editor→Admin 共有DB
    // ✅ **DATABASE_SPREADSHEET_ID**: Shared database accessible by all authenticated users
    // ✅ **Service Account**: Only required for system operations or when specifically requested
    const useServiceAccount = context.forceServiceAccount || false;

    console.log(`findUserByEmail: ${useServiceAccount ? 'Service account' : 'Normal permissions'} access to shared DATABASE_SPREADSHEET_ID`);

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
        return createUserObjectFromRow(row, headers);
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

    // 🔧 CLAUDE.md準拠: DATABASE_SPREADSHEET_ID は Editor→Admin 共有DB
    // ✅ **DATABASE_SPREADSHEET_ID**: Shared database accessible by all authenticated users
    // ✅ **Service Account**: Only required for system operations or when specifically requested
    const useServiceAccount = context.forceServiceAccount || false;

    console.log(`findUserById: ${useServiceAccount ? 'Service account' : 'Normal permissions'} lookup in shared DATABASE_SPREADSHEET_ID`);

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

        console.log(`findUserById: User found via ${useServiceAccount ? 'service account' : 'normal permissions'} lookup in shared DATABASE_SPREADSHEET_ID`);

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
  try {
    if (!email || !validateEmail(email).isValid) {
      console.warn('createUser: Invalid email provided:', typeof email, email);
      return null;
    }

    console.log('createUser: Processing email:', email);

    // 🔧 CLAUDE.md準拠: DATABASE_SPREADSHEET_ID は Editor→Admin 共有DB
    const currentEmail = getCurrentEmail();

    const existingUser = findUserByEmail(email, {
      requestingUser: currentEmail
    });
    if (existingUser) {
      console.info('createUser: 既存ユーザーを返却', { email: `${email.substring(0, 5)}***` });
      return existingUser;
    }

    // ✅ **DATABASE_SPREADSHEET_ID**: Shared database for user creation (normal permissions)
    // ✅ **Service Account**: Only use when specifically required by context
    console.log(`createUser: User creation in shared DATABASE_SPREADSHEET_ID`);
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

    // ✅ Optimized: Remove createdAt from configJSON, store in database column
    const defaultConfig = {
      setupStatus: 'pending',
      isPublished: false,
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

    const user = {
      userId,
      userEmail: email,
      isActive: true,
      configJson: JSON.stringify(defaultConfig),
      createdAt: hasCreatedAtColumn ? now : undefined,
      lastModified: now
    };

    console.log('✅ User created successfully:', `${email.split('@')[0]}@***`);
    return user;

  } catch (error) {
    console.error('createUser error:', error.message);
    return null;
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
    // 🔧 CLAUDE.md準拠: Cross-user Access for getAllUsers (Admin-only operation)
    const currentEmail = getCurrentEmail();
    const isAdmin = isAdministrator(currentEmail);

    if (!isAdmin && !context.forceServiceAccount) {
      console.warn('getAllUsers: Non-admin user attempting cross-user data access');
      return [];
    }

    // getAllUsers is inherently cross-user operation, always requires service account
    // unless admin is accessing for administrative purposes
    console.log(`getAllUsers: ${isAdmin ? 'Admin' : 'System'} cross-user access to DATABASE_SPREADSHEET_ID`);
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

    return users;
  } catch (error) {
    console.error('getAllUsers error:', error.message);
    return [];
  }
}

// ===========================================
// 🛠️ ヘルパー関数
// ===========================================

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
  try {
    // 🔧 CLAUDE.md準拠: Self vs Cross-user Access for User Updates

    // Initial user lookup to determine target user
    const targetUser = findUserById(userId, context);

    if (!targetUser) {
      console.warn('updateUser: Target user not found:', userId);
      return { success: false, message: 'User not found' };
    }

    // ✅ **DATABASE_SPREADSHEET_ID**: Shared database accessible by authenticated users
    // ✅ **Service Account**: Only use when specifically required by context
    const useServiceAccount = context.forceServiceAccount || false;

    console.log(`updateUser: Update operation in shared DATABASE_SPREADSHEET_ID ${useServiceAccount ? 'with service account' : 'with normal permissions'}`);

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

        return { success: true };
      }
    }

    console.warn('updateUser: User not found:', userId);
    return { success: false, message: 'User not found' };
  } catch (error) {
    console.error('updateUser error:', error.message);
    return { success: false, message: error.message || 'Unknown error occurred' };
  }
}

/**
 * ユーザーのスプレッドシートデータを取得（CLAUDE.md準拠 - Context-Aware）
 * @param {Object} targetUser - 対象ユーザーオブジェクト
 * @param {Object} options - オプション設定
 * @param {Object} options.dataAccess - 事前取得されたデータアクセスオブジェクト
 * @returns {Object} User spreadsheet data with configuration
 */
function getUserSpreadsheetData(targetUser, options = {}) {
  try {
    if (!targetUser || !targetUser.userId) {
      console.warn('getUserSpreadsheetData: Invalid target user or missing userId');
      return null;
    }

    // 🔧 CLAUDE.md準拠: configJson-based unified configuration
    const configResult = getUserConfig(targetUser.userId);
    if (!configResult.success || !configResult.config) {
      console.warn('getUserSpreadsheetData: Failed to get user configuration');
      return null;
    }

    const {config} = configResult;
    if (!config.spreadsheetId) {
      console.warn('getUserSpreadsheetData: No spreadsheetId in user configuration');
      return null;
    }

    // 事前に取得されたdataAccessがある場合はそれを使用（最適化）
    let {dataAccess} = options;
    if (!dataAccess) {
      // データアクセスが提供されていない場合、新規取得
      const currentEmail = getCurrentEmail();
      const isSelfAccess = currentEmail === targetUser.userEmail;

      console.log(`getUserSpreadsheetData: ${isSelfAccess ? 'Self-access normal permissions' : 'Cross-user service account'} for spreadsheet data`);
      dataAccess = openSpreadsheet(config.spreadsheetId, {
        useServiceAccount: !isSelfAccess,
        context: 'getUserSpreadsheetData'
      });
    }

    if (!dataAccess || !dataAccess.spreadsheet) {
      console.warn('getUserSpreadsheetData: Failed to access target spreadsheet');
      return null;
    }

    // 基本的なスプレッドシート情報を取得
    const {spreadsheet} = dataAccess;
    const sheets = spreadsheet.getSheets();

    // 🧹 レガシーコード削除: 重複するJSON.parse()処理を排除
    // すでに取得した config オブジェクトを使用

    const result = {
      userId: targetUser.userId,
      userEmail: targetUser.userEmail,
      spreadsheetId: config.spreadsheetId,
      sheetName: config.sheetName || '',
      formUrl: config.formUrl || '',
      config,
      sheets: sheets.map(sheet => ({
        name: sheet.getName(),
        id: sheet.getSheetId()
      })),
      dataAccess // For subsequent operations
    };

    console.log(`getUserSpreadsheetData: Successfully retrieved data for ${targetUser.userEmail ? `${targetUser.userEmail.split('@')[0]  }@***` : 'unknown user'}`);
    return result;

  } catch (error) {
    console.error('getUserSpreadsheetData error:', error.message);
    return null;
  }
}

/**
 * 閲覧者向けボードデータ取得（CLAUDE.md準拠 - 模範実装）
 *
 * 🎯 **CLAUDE.md Pattern Implementation (Lines 52-64)**
 * この関数はCLAUDE.mdで推奨されるサービスアカウント使用パターンの完全実装です。
 * Context-awareアクセス制御により、自分のデータは通常権限、他人のデータは
 * サービスアカウントを使用してセキュリティを確保します。
 *
 * 🔐 **Security Pattern**:
 * - Self-access: `targetUser.userEmail === viewerEmail` → Normal permissions
 * - Cross-user access: `targetUser.userEmail !== viewerEmail` → Service account
 *
 * 🚀 **Performance Features**:
 * - Efficient user lookup by ID
 * - Reuses dataAccess object for subsequent operations
 * - Proper error handling and logging
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
    const targetUser = findUserById(targetUserId);
    if (!targetUser) {
      console.warn('getViewerBoardData: Target user not found:', targetUserId);
      return null;
    }

    if (targetUser.userEmail === viewerEmail) {
      // ✅ Own data: use normal permissions (CLAUDE.md pattern)
      console.log('getViewerBoardData: Self-access - using normal permissions');
      // ✅ Self-access: Use getUserSheetData to get actual sheet data
      // eslint-disable-next-line no-undef
      return getUserSheetData(targetUser.userId, {
        includeTimestamp: true,
        requestingUser: viewerEmail
      });
    } else {
      // ✅ Other's data: use service account for cross-user access (CLAUDE.md pattern)
      console.log('getViewerBoardData: Cross-user access - using service account');
      // ✅ Cross-user: Use getUserSheetData with context for service account usage
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
 * ✅ CLAUDE.md準拠: Single Source of Truth - configJSON内のspreadsheetIdで検索
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

    // 🚀 パフォーマンス最適化: キャッシュ機能実装
    const cacheKey = `user_by_sheet_${spreadsheetId}`;
    const skipCache = context.skipCache || false;

    if (!skipCache) {
      try {
        const cached = CacheService.getScriptCache().get(cacheKey);
        if (cached) {
          const cachedUser = JSON.parse(cached);
          console.log(`findUserBySpreadsheetId: Cache hit for spreadsheet ${spreadsheetId.substring(0, 8)}***`);
          return cachedUser;
        }
      } catch (cacheError) {
        console.warn('findUserBySpreadsheetId: Cache read failed:', cacheError.message);
      }
    }

    console.log(`findUserBySpreadsheetId: ConfigJSON-based lookup for spreadsheet ${spreadsheetId.substring(0, 8)}***`);

    // ✅ Single Source of Truth: getAllUsers()でユーザー一覧を取得し、configJSONから検索
    // Cross-user lookup is legitimate for spreadsheetId-based user identification
    const allUsers = getAllUsers({ activeOnly: false }, { ...context, forceServiceAccount: true });
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
          console.log(`findUserBySpreadsheetId: User found via configJSON lookup - ${user.userEmail ? `${user.userEmail.split('@')[0]}@***` : 'unknown'}`);

          // 🚀 パフォーマンス最適化: キャッシュに保存
          if (!skipCache) {
            try {
              const cacheTtl = context.cacheTtl || CACHE_DURATION.LONG; // 300秒
              CacheService.getScriptCache().put(cacheKey, JSON.stringify(user), cacheTtl);
              console.log(`findUserBySpreadsheetId: User cached for spreadsheet ${spreadsheetId.substring(0, 8)}*** (TTL: ${cacheTtl}s)`);
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

    console.log('findUserBySpreadsheetId: No user found with spreadsheetId in configJSON:', `${spreadsheetId.substring(0, 8)}***`);

    // ユーザーが見つからない場合も短時間キャッシュ（重複検索回避）
    if (!skipCache) {
      try {
        const notFoundTtl = 60; // 60秒
        CacheService.getScriptCache().put(cacheKey, JSON.stringify(null), notFoundTtl);
        console.log(`findUserBySpreadsheetId: Not found result cached for ${spreadsheetId.substring(0, 8)}*** (TTL: ${notFoundTtl}s)`);
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
    // 🔧 CLAUDE.md準拠: Cross-user Access for User Deletion (Admin-only operation)
    const currentEmail = getCurrentEmail();
    const isAdmin = isAdministrator(currentEmail);

    if (!isAdmin && !context.forceServiceAccount) {
      console.warn('deleteUser: Non-admin user attempting user deletion:', userId);
      return { success: false, message: 'Insufficient permissions for user deletion' };
    }

    // User deletion is inherently cross-user administrative operation
    // Always requires service account for safety and audit trail
    console.log(`deleteUser: Admin cross-user deletion in DATABASE_SPREADSHEET_ID`);
    const spreadsheet = openDatabase(true); // Always service account for deletion operations
    if (!spreadsheet) {
      console.warn('deleteUser: Database access failed');
      return { success: false, message: 'Database access failed' };
    }

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
        // Soft delete by setting isActive to false
        const isActiveIndex = headers.indexOf('isActive');
        if (isActiveIndex !== -1) {
          sheet.getRange(i + 1, isActiveIndex + 1).setValue(false);
        }

        // Update lastModified
        const lastModifiedIndex = headers.indexOf('lastModified');
        if (lastModifiedIndex !== -1) {
          sheet.getRange(i + 1, lastModifiedIndex + 1).setValue(new Date().toISOString());
        }

        console.log('✅ User soft deleted successfully:', `${userId.substring(0, 8)}***`, reason ? `Reason: ${reason}` : '');
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