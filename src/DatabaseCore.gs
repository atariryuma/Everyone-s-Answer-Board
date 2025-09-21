/**
 * @fileoverview DatabaseCore - データベース操作の基盤
 *
 * 責任範囲:
 * - 基本的なデータベース操作（CRUD）
 * - ユーザー管理の基盤関数
 * - GAS-Native Architecture準拠（直接SpreadsheetApp使用）
 * - サービスアカウント使用時の安全な権限管理
 */

/* global validateEmail, CACHE_DURATION, TIMEOUT_MS */

// ===========================================
// 🗄️ データベース基盤操作
// ===========================================

/**
 * サービスアカウント認証情報を取得
 * @returns {Object} Service account info with isValid flag
 */
function getServiceAccount() {
  try {
    const creds = PropertiesService.getScriptProperties().getProperty('SERVICE_ACCOUNT_CREDS');
    if (!creds) {
      return { isValid: false };
    }

    const serviceAccount = JSON.parse(creds);
    return {
      isValid: !!(serviceAccount.client_email && serviceAccount.private_key),
      email: serviceAccount.client_email
    };
  } catch (error) {
    console.warn('getServiceAccount: Invalid credentials format');
    return { isValid: false };
  }
}

/**
 * データベーススプレッドシートを開く
 * @param {boolean} useServiceAccount - サービスアカウントを使用するか
 * @returns {Object|null} Database spreadsheet object
 */
function openDatabase(useServiceAccount = false) {
  try {
    const dbId = PropertiesService.getScriptProperties().getProperty('DATABASE_SPREADSHEET_ID');
    if (!dbId) {
      console.warn('openDatabase: DATABASE_SPREADSHEET_ID not configured');
      return null;
    }

    if (useServiceAccount) {
      const auth = getServiceAccount();
      if (!auth.isValid) {
        console.warn('openDatabase: Service account authentication failed');
        return null;
      }

      // Grant service account access if needed
      try {
        DriveApp.getFileById(dbId).addEditor(auth.email);
      } catch (driveError) {
        console.warn('openDatabase: Service account access already granted or failed:', driveError.message);
      }
    }

    return SpreadsheetApp.openById(dbId);
  } catch (error) {
    console.error('openDatabase error:', error.message);
    return null;
  }
}

// ===========================================
// 👤 ユーザー管理基盤
// ===========================================

/**
 * メールアドレスでユーザーを検索
 * @param {string} email - ユーザーのメールアドレス
 * @returns {Object|null} User object
 */
function findUserByEmail(email) {
  try {
    if (!email || !validateEmail(email).isValid) {
      return null;
    }

    const spreadsheet = openDatabase(true); // Service account required
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
 * ユーザーIDでユーザーを検索
 * @param {string} userId - ユーザーID
 * @returns {Object|null} User object
 */
function findUserById(userId) {
  try {
    if (!userId) {
      return null;
    }

    const spreadsheet = openDatabase(true); // Service account required
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
        return createUserObjectFromRow(row, headers);
      }
    }

    return null;
  } catch (error) {
    console.error('findUserById error:', error.message);
    return null;
  }
}

/**
 * 新しいユーザーを作成
 * @param {string} email - ユーザーのメールアドレス
 * @param {Object} initialConfig - 初期設定
 * @returns {Object|null} Created user object
 */
function createUser(email, initialConfig = {}) {
  try {
    if (!email || !validateEmail(email).isValid) {
      console.warn('createUser: Invalid email provided');
      return null;
    }

    const existingUser = findUserByEmail(email);
    if (existingUser) {
      console.info('createUser: 既存ユーザーを返却', { email: `${email.substring(0, 5)}***` });
      return existingUser;
    }

    const spreadsheet = openDatabase(true); // Service account required
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
      createdAt: now,
      ...initialConfig
    };

    const newUserData = [
      userId,
      email,
      true, // isActive
      JSON.stringify(defaultConfig),
      now, // lastModified
      '', // spreadsheetId (to be set later)
      '', // sheetName (to be set later)
      '', // formUrl (to be set later)
    ];

    sheet.appendRow(newUserData);

    const user = {
      userId,
      userEmail: email,
      isActive: true,
      configJson: JSON.stringify(defaultConfig),
      lastModified: now,
      spreadsheetId: '',
      sheetName: '',
      formUrl: ''
    };

    console.log('✅ User created successfully:', `${email.split('@')[0]}@***`);
    return user;

  } catch (error) {
    console.error('createUser error:', error.message);
    return null;
  }
}

/**
 * 全ユーザーリストを取得
 * @param {Object} options - オプション設定
 * @returns {Array} User list
 */
function getAllUsers(options = {}) {
  try {
    const spreadsheet = openDatabase(true); // Service account required
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
    'lastModified': 'lastModified',
    'spreadsheetId': 'spreadsheetId',
    'sheetName': 'sheetName',
    'formUrl': 'formUrl'
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
 * ユーザー情報を更新
 * @param {string} userId - ユーザーID
 * @param {Object} updates - 更新データ
 * @returns {Object} {success: boolean, message?: string} Operation result
 */
function updateUser(userId, updates) {
  try {
    const spreadsheet = openDatabase(true); // Service account required
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
 * ユーザーを削除
 * @param {string} userId - ユーザーID
 * @param {string} reason - 削除理由
 * @returns {Object} Delete operation result
 */
function deleteUser(userId, reason = '') {
  try {
    const spreadsheet = openDatabase(true); // Service account required
    if (!spreadsheet) {
      console.warn('deleteUser: Database access failed');
      return { success: false, error: 'Database access failed' };
    }

    const sheet = spreadsheet.getSheetByName('users');
    if (!sheet) {
      console.warn('deleteUser: Users sheet not found');
      return { success: false, error: 'Users sheet not found' };
    }

    const data = sheet.getDataRange().getValues();
    const [headers] = data;
    const userIdColumnIndex = headers.indexOf('userId');

    if (userIdColumnIndex === -1) {
      console.warn('deleteUser: UserId column not found');
      return { success: false, error: 'UserId column not found' };
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
    return { success: false, error: 'User not found' };
  } catch (error) {
    console.error('deleteUser error:', error.message);
    return { success: false, error: error.message };
  }
}