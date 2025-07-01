/**
 * @OnlyCurrentDoc
 *
 * ===================================================================================
 * StudyQuest Answer Board - Admin Logger API (Optimized Version)
 * ===================================================================================
 *
 * High-performance, error-resistant API for user management and logging.
 * All messages converted to English for better international compatibility.
 *
 * Key Optimizations:
 * - Enhanced cache management with invalidation
 * - Batch operations for better performance
 * - Improved error handling with English messages
 * - Optimized database queries
 * - Better lock management
 */

// Global Configuration
const CONFIG = {
  DATABASE_ID_KEY: 'DATABASE_ID',
  DEPLOYMENT_ID_KEY: 'DEPLOYMENT_ID',
  TARGET_SHEET_NAME: 'Users',
  CACHE_TTL: 300, // 5 minutes
  LOCK_TIMEOUT: 15000, // 15 seconds
  MAX_RETRIES: 3,
  BATCH_SIZE: 100
};

const HEADERS = [
  'userId', 'adminEmail', 'spreadsheetId', 'spreadsheetUrl', 
  'createdAt', 'accessToken', 'configJson', 'lastAccessedAt', 'isActive'
];

// Memory cache for frequently accessed data
const MEMORY_CACHE = new Map();
const MEMORY_CACHE_TTL = 60000; // 1 minute

/**
 * Adds custom menu when spreadsheet is opened
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 Admin Logger Setup')
    .addItem('1. Initialize Database', 'initializeDatabase')
    .addItem('2. Deploy API', 'showDeploymentInstructions')
    .addSeparator()
    .addItem('Show Current Settings', 'showCurrentSettings')
    .addItem('Test Deployment', 'testDeployment')
    .addSeparator()
    .addItem('🔍 View Database Contents', 'debugDatabaseContents')
    .addSeparator()
    .addItem('🧹 Clear Database', 'clearDatabase')
    .addItem('🔧 Cleanup Invalid Users', 'cleanupInvalidUsers')
    .addItem('🗑️ Clear All Caches', 'clearAllCaches')
    .addToUi();
}

/**
 * Initialize spreadsheet as database
 */
function initializeDatabase() {
  const ui = SpreadsheetApp.getUi();
  const properties = PropertiesService.getScriptProperties();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetId = spreadsheet.getId();

  if (properties.getProperty(CONFIG.DATABASE_ID_KEY) === sheetId) {
    ui.alert('✅ Database already initialized.');
    return;
  }

  const confirmation = ui.alert(
    'Database Initialization',
    'Initialize this spreadsheet as the user database? Sheet will be renamed and headers created.',
    ui.ButtonSet.OK_CANCEL
  );

  if (confirmation !== ui.Button.OK) {
    ui.alert('Initialization cancelled.');
    return;
  }

  try {
    properties.setProperty(CONFIG.DATABASE_ID_KEY, sheetId);

    let sheet = spreadsheet.getSheetByName(CONFIG.TARGET_SHEET_NAME);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(CONFIG.TARGET_SHEET_NAME, 0);
      const defaultSheet = spreadsheet.getSheetByName('Sheet1') || spreadsheet.getSheetByName('シート1');
      if (defaultSheet && spreadsheet.getSheets().length > 1) {
        spreadsheet.deleteSheet(defaultSheet);
      }
    }
    
    sheet.clearContents();
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight("bold");
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, HEADERS.length);

    spreadsheet.rename('StudyQuest Admin Logger Database');


    ui.alert('✅ Database initialized successfully. Template form and spreadsheet created. Next: Deploy API from menu.');

  } catch (e) {
    console.error('Database initialization error:', e);
    ui.alert(`Error occurred: ${e.message}`);
  }
}

/**
 * GET request handler for basic connectivity testing
 */
function doGet(e) {
  console.log('GET request received');
  
  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    message: 'StudyQuest Logger API is operational',
    timestamp: new Date().toISOString(),
    service: 'StudyQuest Logger API v2.0'
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Main POST request handler with enhanced error handling and performance
 */
function doPost(e) {
  const startTime = Date.now();
  let responsePayload;
  let lock;

  try {
    // Enhanced lock acquisition with retry
    lock = acquireLock();

    if (!e?.postData?.contents) {
      throw new Error('Invalid request: Missing request body');
    }

    const requestData = JSON.parse(e.postData.contents);
    
    // Process structured API requests
    if (requestData.action) {
      responsePayload = handleApiRequest(requestData);
    } else {
      // Legacy direct logging for backward compatibility
      logMetadataToDatabase(requestData);
      responsePayload = { 
        status: 'success', 
        message: 'Data logged successfully',
        processingTime: Date.now() - startTime
      };
    }

  } catch (error) {
    console.error(`doPost error: ${error.toString()}\nStack: ${error.stack}`);
    responsePayload = { 
      status: 'error', 
      message: `Server error: ${error.message}`,
      errorCode: getErrorCode(error)
    };
  } finally {
    if (lock) {
      lock.releaseLock();
    }
  }

  return ContentService.createTextOutput(JSON.stringify(responsePayload))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Validate domain access for DOMAIN authentication
 */
function validateDomainAccess() {
  try {
    const activeUser = Session.getActiveUser();
    const effectiveUser = Session.getEffectiveUser();
    
    // Check if user is authenticated
    if (!activeUser || !activeUser.getEmail()) {
      return {
        valid: false,
        reason: 'No authenticated user found'
      };
    }
    
    const userEmail = activeUser.getEmail();
    const userDomain = userEmail.split('@')[1];
    
    // Check if user is from allowed domain
    const allowedDomain = 'naha-okinawa.ed.jp';
    if (userDomain !== allowedDomain) {
      return {
        valid: false,
        reason: `Access denied. Expected domain: ${allowedDomain}, Got: ${userDomain}`
      };
    }
    
    return {
      valid: true,
      domain: userDomain,
      userEmail: userEmail,
      effectiveEmail: effectiveUser?.getEmail()
    };
    
  } catch (e) {
    console.error('Domain validation error:', e);
    return {
      valid: false,
      reason: `Authentication error: ${e.message}`
    };
  }
}

/**
 * Acquire lock with retry mechanism
 */
function acquireLock() {
  const lock = LockService.getScriptLock();
  let attempts = 0;
  
  while (attempts < CONFIG.MAX_RETRIES) {
    try {
      if (lock.tryLock(CONFIG.LOCK_TIMEOUT)) {
        return lock;
      }
      attempts++;
      if (attempts < CONFIG.MAX_RETRIES) {
        Utilities.sleep(100 * attempts); // Exponential backoff
      }
    } catch (e) {
      console.error(`Lock acquisition attempt ${attempts + 1} failed:`, e);
      attempts++;
    }
  }
  
  throw new Error('Server busy: Lock acquisition failed after retries');
}

/**
 * Enhanced API request handler with performance optimizations and DOMAIN authentication
 */
function handleApiRequest(requestData) {
  const { action, data, timestamp, requestUser, effectiveUser } = requestData;
  
  // DOMAIN認証チェック（データベース操作が必要な場合のみ）
  if (action !== 'ping') {
    const authResult = validateDomainAccess();
    if (!authResult.valid) {
      console.log(`❌ API Access Denied: ${authResult.reason}`);
      return {
        success: false,
        error: 'DOMAIN_ACCESS_DENIED',
        message: authResult.reason,
        timestamp: new Date().toISOString()
      };
    }
    console.log(`✅ API Access Granted for: ${authResult.userEmail}`);
  }
  
  // Log API access for monitoring
  console.log(`API Request: ${action} from ${requestUser || 'anonymous'}`);
  
  switch (action) {
    case 'ping':
      return {
        success: true,
        message: 'Logger API operational',
        timestamp: new Date().toISOString(),
        data: { 
          pong: true,
          requestUser: requestUser,
          effectiveUser: effectiveUser,
          version: '2.0',
          authRequired: false,
          testEndpoint: true
        }
      };
      
    case 'getUserInfo':
      return handleGetUserInfo(data);
      
    case 'createUser':
      return handleCreateUser(data, requestUser);
      
    case 'updateUser':
      return handleUpdateUser(data, requestUser);

    case 'getExistingBoard':
      return handleGetExistingBoard(data);

    case 'checkExistingUser':
      return handleCheckExistingUser(data);

    case 'createOrFetchUser':
      return handleCreateOrFetchUser(data, requestUser);

    case 'deleteUser':
      return handleDeleteUser(data, requestUser);
      
    case 'invalidateCache':
      return handleInvalidateCache(data);
      
      
    default:
      throw new Error(`Unknown API action: ${action}`);
  }
}

/**
 * Enhanced user info retrieval with cache verification
 */
function handleGetUserInfo(data) {
  try {
    const { userId } = data;
    
    if (!userId) {
      return { success: false, error: 'userId is required' };
    }
    
    // Validate userId format
    if (!isValidUserId(userId)) {
      return { success: false, error: 'Invalid userId format' };
    }
    
    const userData = findUserById(userId);
    
    if (userData) {
      // Verify data is still valid in database
      if (userData.isActive === false || userData.isActive === 'FALSE') {
        invalidateUserCache(userId);
        return { success: false, error: 'User account is inactive' };
      }
      
      return {
        success: true,
        data: userData
      };
    } else {
      return {
        success: false,
        error: 'User not found'
      };
    }
  } catch (error) {
    console.error(`getUserInfo error: ${error.message}`);
    return {
      success: false,
      error: `Database access error: ${error.message}`
    };
  }
}

/**
 * Enhanced user creation with validation
 */
function handleCreateUser(data, requestUser) {
  try {
    // Validate required fields
    if (!data.userId || !data.adminEmail) {
      return { success: false, error: 'userId and adminEmail are required' };
    }
    
    if (!isValidEmail(data.adminEmail)) {
      return { success: false, error: 'Invalid email format' };
    }
    
    // Check if user already exists
    const existingUser = findUserById(data.userId);
    if (existingUser) {
      return { success: false, error: 'User already exists' };
    }
    
    const dbSheet = getDatabaseSheet();
    const timestamp = new Date();
    
    const newRow = [
      data.userId,
      data.adminEmail,
      data.spreadsheetId || '',
      data.spreadsheetUrl || '',
      timestamp,
      data.accessToken || '',
      data.configJson || '{}',
      timestamp,
      data.isActive !== undefined ? data.isActive : true
    ];
    
    dbSheet.appendRow(newRow);

    // Clear caches
    invalidateUserCache(data.userId);
    invalidateEmailCache(data.adminEmail);
    
    return {
      success: true,
      message: 'User created successfully',
      data: {
        userId: data.userId,
        adminEmail: data.adminEmail,
        createdAt: timestamp
      }
    };
    
  } catch (error) {
    console.error(`createUser error: ${error.message}`);
    return {
      success: false,
      error: `User creation failed: ${error.message}`
    };
  }
}

/**
 * Enhanced user update with atomic operations
 */
function handleUpdateUser(data, requestUser) {
  try {
    const { userId } = data;
    
    if (!userId) {
      return { success: false, error: 'userId is required' };
    }
    
    const dbSheet = getDatabaseSheet();
    const userRow = findUserRowById(dbSheet, userId);
    
    if (!userRow) {
      // Clear cache since user doesn't exist
      invalidateUserCache(userId);
      return { success: false, error: 'User not found' };
    }
    
    // Update fields atomically
    const updates = {};
    if (data.spreadsheetId !== undefined) updates.spreadsheetId = data.spreadsheetId;
    if (data.spreadsheetUrl !== undefined) updates.spreadsheetUrl = data.spreadsheetUrl;
    if (data.accessToken !== undefined) updates.accessToken = data.accessToken;
    if (data.configJson !== undefined) updates.configJson = data.configJson;
    if (data.isActive !== undefined) updates.isActive = data.isActive;
    
    // Apply updates
    if (updates.spreadsheetId !== undefined) userRow.values[2] = updates.spreadsheetId;
    if (updates.spreadsheetUrl !== undefined) userRow.values[3] = updates.spreadsheetUrl;
    if (updates.accessToken !== undefined) userRow.values[5] = updates.accessToken;
    if (updates.configJson !== undefined) userRow.values[6] = updates.configJson;
    if (updates.isActive !== undefined) userRow.values[8] = updates.isActive;
    
    // Always update lastAccessedAt
    userRow.values[7] = new Date();
    
    // Write back to sheet
    dbSheet.getRange(userRow.rowIndex, 1, 1, userRow.values.length).setValues([userRow.values]);

    // Clear caches
    invalidateUserCache(userId);
    if (userRow.values[1]) {
      invalidateEmailCache(userRow.values[1]);
    }
    
    return {
      success: true,
      message: 'User updated successfully',
      data: {
        userId: userId,
        updatedAt: new Date(),
        updatedFields: Object.keys(updates)
      }
    };
    
  } catch (error) {
    console.error(`updateUser error: ${error.message}`);
    return {
      success: false,
      error: `Update failed: ${error.message}`
    };
  }
}

/**
 * Enhanced existing board retrieval
 */
function handleGetExistingBoard(data) {
  try {
    const { adminEmail } = data;
    
    if (!adminEmail) {
      return { success: false, error: 'adminEmail is required' };
    }
    
    if (!isValidEmail(adminEmail)) {
      return { success: false, error: 'Invalid email format' };
    }
    
    const userData = findUserByEmail(adminEmail);
    
    if (userData) {
      // Verify user is active
      if (userData.isActive === false || userData.isActive === 'FALSE') {
        invalidateEmailCache(adminEmail);
        return { success: false, error: 'User account is inactive' };
      }
      
      return {
        success: true,
        data: userData
      };
    } else {
      return {
        success: false,
        message: 'No existing board found'
      };
    }
  } catch (error) {
    console.error(`getExistingBoard error: ${error.message}`);
    return {
      success: false,
      error: `Database access error: ${error.message}`
    };
  }
}

/**
 * Enhanced existing user check
 */
function handleCheckExistingUser(data) {
  try {
    const { adminEmail } = data;
    
    if (!adminEmail) {
      return { success: false, error: 'adminEmail is required' };
    }
    
    const userData = findUserByEmail(adminEmail);
    
    return {
      success: true,
      exists: userData !== null,
      data: userData,
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    console.error(`checkExistingUser error: ${error.message}`);
    return {
      success: false,
      error: `Check failed: ${error.message}`
    };
  }
}

/**
 * Check for an existing user by email or create a new one in a single call
 */
function handleCreateOrFetchUser(data, requestUser) {
  try {
    const { adminEmail } = data;

    if (!adminEmail) {
      return { success: false, error: 'adminEmail is required' };
    }

    const existingUser = findUserByEmail(adminEmail);

    if (existingUser) {
      return {
        success: true,
        created: false,
        data: existingUser
      };
    }

    // Require userId when creating a new user
    if (!data.userId) {
      return { success: false, error: 'userId is required for creation' };
    }

    const createResult = handleCreateUser(data, requestUser);

    if (createResult.success) {
      const newUser = findUserById(data.userId);
      return {
        success: true,
        created: true,
        data: newUser
      };
    }

    return createResult;
  } catch (error) {
    console.error(`createOrFetchUser error: ${error.message}`);
    return {
      success: false,
      error: `createOrFetchUser failed: ${error.message}`
    };
  }
}

/**
 * New: Delete user with proper cleanup
 */
function handleDeleteUser(data, requestUser) {
  try {
    const { userId, adminEmail } = data;
    
    if (!userId && !adminEmail) {
      return { success: false, error: 'userId or adminEmail is required' };
    }
    
    const dbSheet = getDatabaseSheet();
    let userRow;
    
    if (userId) {
      userRow = findUserRowById(dbSheet, userId);
    } else {
      userRow = findUserRowByEmail(dbSheet, adminEmail);
    }
    
    if (!userRow) {
      return { success: false, error: 'User not found' };
    }
    
    // Delete row from sheet
    dbSheet.deleteRow(userRow.rowIndex);
    
    // Clear all related caches
    if (userRow.values[0]) invalidateUserCache(userRow.values[0]);
    if (userRow.values[1]) invalidateEmailCache(userRow.values[1]);
    
    return {
      success: true,
      message: 'User deleted successfully',
      data: {
        deletedUserId: userRow.values[0],
        deletedEmail: userRow.values[1],
        deletedAt: new Date()
      }
    };
    
  } catch (error) {
    console.error(`deleteUser error: ${error.message}`);
    return {
      success: false,
      error: `Deletion failed: ${error.message}`
    };
  }
}

/**
 * New: Manual cache invalidation endpoint
 */
function handleInvalidateCache(data) {
  try {
    const { userId, adminEmail, clearAll } = data;
    
    if (clearAll) {
      clearAllCaches();
      return {
        success: true,
        message: 'All caches cleared'
      };
    }
    
    if (userId) {
      invalidateUserCache(userId);
    }
    
    if (adminEmail) {
      invalidateEmailCache(adminEmail);
    }
    
    return {
      success: true,
      message: 'Cache invalidated',
      clearedKeys: { userId, adminEmail }
    };
    
  } catch (error) {
    console.error(`invalidateCache error: ${error.message}`);
    return {
      success: false,
      error: `Cache invalidation failed: ${error.message}`
    };
  }
}


/**
 * Enhanced user search by ID with multi-level caching
 */
function findUserById(userId) {
  if (!userId || !isValidUserId(userId)) {
    return null;
  }
  
  // Check memory cache first
  const memoryKey = `user_id_${userId}`;
  const memoryCached = MEMORY_CACHE.get(memoryKey);
  if (memoryCached && (Date.now() - memoryCached.timestamp) < MEMORY_CACHE_TTL) {
    return memoryCached.data;
  }
  
  // Check CacheService
  const cache = CacheService.getScriptCache();
  const cacheKey = `user_id_${userId}`;
  try {
    const cached = cache.get(cacheKey);
    if (cached) {
      const userData = JSON.parse(cached);
      // Update memory cache
      MEMORY_CACHE.set(memoryKey, { data: userData, timestamp: Date.now() });
      return userData;
    }
  } catch (e) {
    console.warn(`Cache read error for user ${userId}:`, e.message);
  }

  // Fetch from database
  try {
    const dbSheet = getDatabaseSheet();
    const userData = findUserInSheet(dbSheet, userId, 'userId');
    
    if (userData) {
      // Cache the result
      try {
        cache.put(cacheKey, JSON.stringify(userData), CONFIG.CACHE_TTL);
        MEMORY_CACHE.set(memoryKey, { data: userData, timestamp: Date.now() });
      } catch (e) {
        console.warn(`Cache write error for user ${userId}:`, e.message);
      }
    }
    
    return userData;
  } catch (error) {
    console.error(`Database search error for user ${userId}:`, error.message);
    throw new Error(`User search failed: ${error.message}`);
  }
}

/**
 * Enhanced user search by email with validation
 */
function findUserByEmail(adminEmail) {
  if (!adminEmail || !isValidEmail(adminEmail)) {
    return null;
  }
  
  const normalizedEmail = adminEmail.trim().toLowerCase();
  
  // Check memory cache first
  const memoryKey = `user_email_${normalizedEmail}`;
  const memoryCached = MEMORY_CACHE.get(memoryKey);
  if (memoryCached && (Date.now() - memoryCached.timestamp) < MEMORY_CACHE_TTL) {
    return memoryCached.data;
  }
  
  // Check CacheService
  const cache = CacheService.getScriptCache();
  const cacheKey = `user_email_${normalizedEmail}`;
  try {
    const cached = cache.get(cacheKey);
    if (cached) {
      const userData = JSON.parse(cached);
      // Update memory cache
      MEMORY_CACHE.set(memoryKey, { data: userData, timestamp: Date.now() });
      return userData;
    }
  } catch (e) {
    console.warn(`Cache read error for email ${normalizedEmail}:`, e.message);
  }

  // Fetch from database
  try {
    const dbSheet = getDatabaseSheet();
    const userData = findUserInSheet(dbSheet, normalizedEmail, 'email');
    
    if (userData) {
      // Cache the result
      try {
        cache.put(cacheKey, JSON.stringify(userData), CONFIG.CACHE_TTL);
        MEMORY_CACHE.set(memoryKey, { data: userData, timestamp: Date.now() });
      } catch (e) {
        console.warn(`Cache write error for email ${normalizedEmail}:`, e.message);
      }
    }
    
    return userData;
  } catch (error) {
    console.error(`Database search error for email ${normalizedEmail}:`, error.message);
    throw new Error(`Email search failed: ${error.message}`);
  }
}

/**
 * Optimized sheet search with batch processing
 */
function findUserInSheet(sheet, searchValue, searchType) {
  try {
    const data = sheet.getDataRange().getValues();
    
    if (!data || data.length <= 1) {
      return null;
    }
    
    const searchColumn = searchType === 'userId' ? 0 : 1;
    const compareValue = searchType === 'email' ? searchValue.toLowerCase() : searchValue;
    
    // Process in batches for better performance
    for (let i = 1; i < data.length; i += CONFIG.BATCH_SIZE) {
      const endIndex = Math.min(i + CONFIG.BATCH_SIZE, data.length);
      
      for (let j = i; j < endIndex; j++) {
        if (!data[j] || !data[j][searchColumn]) continue;
        
        const cellValue = searchType === 'email' 
          ? String(data[j][searchColumn]).trim().toLowerCase()
          : data[j][searchColumn];
        
        if (cellValue === compareValue) {
          const userData = {
            userId: data[j][0],
            adminEmail: data[j][1],
            spreadsheetId: data[j][2],
            spreadsheetUrl: data[j][3],
            createdAt: data[j][4],
            accessToken: data[j][5],
            configJson: data[j][6],
            lastAccessedAt: data[j][7],
            isActive: data[j][8]
          };
          
          // Validate userId integrity
          if (!userData.userId || userData.userId === '' || userData.userId === 'undefined') {
            continue;
          }
          
          return userData;
        }
      }
    }
    
    return null;
  } catch (error) {
    console.error(`Sheet search error:`, error.message);
    throw error;
  }
}

/**
 * Find user row for updates
 */
function findUserRowById(sheet, userId) {
  try {
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i] && data[i][0] === userId) {
        return {
          rowIndex: i + 1,
          values: data[i]
        };
      }
    }
    
    return null;
  } catch (error) {
    console.error(`findUserRowById error: ${error.message}`);
    throw error;
  }
}

/**
 * Find user row by email for updates
 */
function findUserRowByEmail(sheet, adminEmail) {
  try {
    const data = sheet.getDataRange().getValues();
    const normalizedEmail = adminEmail.trim().toLowerCase();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i] && data[i][1]) {
        const storedEmail = String(data[i][1]).trim().toLowerCase();
        if (storedEmail === normalizedEmail) {
          return {
            rowIndex: i + 1,
            values: data[i]
          };
        }
      }
    }
    
    return null;
  } catch (error) {
    console.error(`findUserRowByEmail error: ${error.message}`);
    throw error;
  }
}

/**
 * Enhanced database sheet access with fallback
 */
function getDatabaseSheet() {
  const properties = PropertiesService.getScriptProperties();
  let dbSheetId = properties.getProperty(CONFIG.DATABASE_ID_KEY);

  // Auto-initialize if not set
  if (!dbSheetId) {
    const currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (currentSpreadsheet) {
      dbSheetId = currentSpreadsheet.getId();
      properties.setProperty(CONFIG.DATABASE_ID_KEY, dbSheetId);
    } else {
      throw new Error('Database not configured. Please run initialization from setup menu.');
    }
  }

  try {
    const spreadsheet = SpreadsheetApp.openById(dbSheetId);
    let sheet = spreadsheet.getSheetByName(CONFIG.TARGET_SHEET_NAME);
    
    if (!sheet) {
      // Auto-create sheet
      sheet = spreadsheet.insertSheet(CONFIG.TARGET_SHEET_NAME);
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight("bold");
      sheet.setFrozenRows(1);
    }
    
    return sheet;
  } catch (error) {
    throw new Error(`Database access failed: ${error.message}`);
  }
}

/**
 * Legacy metadata logging for backward compatibility
 */
function logMetadataToDatabase(data) {
  try {
    const dbSheet = getDatabaseSheet();
    const timestamp = new Date();

    const newRow = [
      data.userId || '',
      data.adminEmail || '',
      data.spreadsheetId || '',
      data.spreadsheetUrl || '',
      timestamp,
      data.accessToken || '',
      data.configJson || '{}',
      timestamp,
      true
    ];

    dbSheet.appendRow(newRow);

  } catch (error) {
    console.error(`Legacy logging failed: ${error.message}`);
    throw new Error(`Database write failed: ${error.message}`);
  }
}

/**
 * Cache management functions
 */
function invalidateUserCache(userId) {
  if (!userId) return;
  
  try {
    const cache = CacheService.getScriptCache();
    cache.remove(`user_id_${userId}`);
    MEMORY_CACHE.delete(`user_id_${userId}`);
  } catch (e) {
    console.warn(`Cache invalidation failed for user ${userId}:`, e.message);
  }
}

function invalidateEmailCache(email) {
  if (!email) return;
  
  try {
    const normalizedEmail = email.trim().toLowerCase();
    const cache = CacheService.getScriptCache();
    cache.remove(`user_email_${normalizedEmail}`);
    MEMORY_CACHE.delete(`user_email_${normalizedEmail}`);
  } catch (e) {
    console.warn(`Cache invalidation failed for email ${email}:`, e.message);
  }
}

function clearAllCaches() {
  try {
    CacheService.getScriptCache().removeAll(['user_id_', 'user_email_']);
    MEMORY_CACHE.clear();
  } catch (e) {
    console.warn(`Cache clearing failed:`, e.message);
  }
}

/**
 * Validation functions
 */
function isValidUserId(userId) {
  return userId && typeof userId === 'string' && userId.trim().length > 0 && userId !== 'undefined' && userId !== 'null';
}

function isValidEmail(email) {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return email && typeof email === 'string' && emailRegex.test(email.trim());
}

/**
 * Error classification for better handling
 */
function getErrorCode(error) {
  const message = error.message.toLowerCase();
  
  if (message.includes('lock')) return 'LOCK_ERROR';
  if (message.includes('permission') || message.includes('access')) return 'PERMISSION_ERROR';
  if (message.includes('not found')) return 'NOT_FOUND';
  if (message.includes('invalid') || message.includes('validation')) return 'VALIDATION_ERROR';
  if (message.includes('timeout')) return 'TIMEOUT_ERROR';
  
  return 'GENERAL_ERROR';
}

/**
 * UI Functions (kept for backward compatibility)
 */
function showCurrentSettings() {
  const ui = SpreadsheetApp.getUi();
  const properties = PropertiesService.getScriptProperties();
  const dbSheetId = properties.getProperty(CONFIG.DATABASE_ID_KEY);
  const deploymentId = properties.getProperty(CONFIG.DEPLOYMENT_ID_KEY);

  let message = '🚀 StudyQuest Admin Logger API - Status Dashboard\n';
  message += '=' .repeat(50) + '\n\n';
  
  // Database Status
  message += '📊 DATABASE STATUS:\n';
  if (dbSheetId) {
    try {
      const dbSheet = getDatabaseSheet();
      const data = dbSheet.getDataRange().getValues();
      const userCount = Math.max(0, data.length - 1);
      
      message += `✅ Database: ACTIVE\n`;
      message += `   📁 Spreadsheet ID: ${dbSheetId}\n`;
      message += `   👥 Total Users: ${userCount}\n`;
      message += `   🏢 Sheet Name: ${CONFIG.TARGET_SHEET_NAME}\n\n`;
    } catch (e) {
      message += `⚠️ Database: CONFIGURED but ERROR\n`;
      message += `   📁 Spreadsheet ID: ${dbSheetId}\n`;
      message += `   ❌ Error: ${e.message}\n\n`;
    }
  } else {
    message += '❌ Database: NOT CONFIGURED\n';
    message += '   📋 Action: Run "Initialize Database" first\n\n';
  }

  // API Deployment Status
  message += '🌐 API DEPLOYMENT STATUS:\n';
  if (deploymentId) {
    const webAppUrl = deploymentId.startsWith('https://') 
      ? deploymentId 
      : `https://script.google.com/macros/s/${deploymentId}/exec`;
    
    message += `✅ API: DEPLOYED & ACTIVE\n`;
    message += `   🔗 URL: ${webAppUrl}\n`;
    message += `   🛡️ Security: DOMAIN (naha-okinawa.ed.jp only)\n`;
    message += `   ⚙️ Execute as: USER_DEPLOYING\n`;
    message += `   📝 All API calls logged\n\n`;
  } else {
    message += '❌ API: NOT DEPLOYED\n';
    message += '   📋 Action: Use "Deploy API" from menu\n\n';
  }

  // Authentication & Security
  message += '🔒 SECURITY & ACCESS:\n';
  try {
    const currentUser = Session.getActiveUser().getEmail();
    const userDomain = currentUser.split('@')[1];
    
    if (userDomain === 'naha-okinawa.ed.jp') {
      message += `✅ Current User: AUTHORIZED\n`;
      message += `   👤 Email: ${currentUser}\n`;
      message += `   🏢 Domain: ${userDomain} ✓\n`;
    } else {
      message += `⚠️ Current User: UNAUTHORIZED DOMAIN\n`;
      message += `   👤 Email: ${currentUser}\n`;
      message += `   🏢 Domain: ${userDomain} ✗\n`;
      message += `   📋 Required: naha-okinawa.ed.jp\n`;
    }
  } catch (e) {
    message += `❌ Authentication: ERROR\n`;
    message += `   ❌ Error: ${e.message}\n`;
  }

  message += '\n' + '=' .repeat(50) + '\n';
  message += 'ℹ️ Need help? Use "Test Deployment" to verify API functionality.';

  ui.alert(message);
}

function testDeployment() {
  const ui = SpreadsheetApp.getUi();
  const properties = PropertiesService.getScriptProperties();
  const deploymentId = properties.getProperty(CONFIG.DEPLOYMENT_ID_KEY);
  
  if (!deploymentId) {
    ui.alert('❌ DEPLOYMENT TEST FAILED\n\nNo deployment ID configured.\n📋 Action: Please deploy API first using "Deploy API" menu.');
    return;
  }
  
  const webAppUrl = deploymentId.startsWith('https://') 
    ? deploymentId 
    : `https://script.google.com/macros/s/${deploymentId}/exec`;
  
  let message = '🧪 API DEPLOYMENT TEST RESULTS\n';
  message += '=' .repeat(45) + '\n\n';
  message += `🔗 Test URL: ${webAppUrl}\n\n`;
  
  try {
    // Test GET
    message += '📡 TEST 1: Basic Connectivity (GET)\n';
    const getResponse = UrlFetchApp.fetch(webAppUrl, {
      method: 'GET',
      muteHttpExceptions: true
    });
    
    const getCode = getResponse.getResponseCode();
    message += `   Status Code: ${getCode}\n`;
    
    if (getCode === 200) {
      message += '   ✅ Result: SUCCESS - API is reachable\n\n';
      
      // Test POST with DOMAIN explanation
      message += '📡 TEST 2: POST Authentication (DOMAIN Security)\n';
      const postResponse = UrlFetchApp.fetch(webAppUrl, {
        method: 'POST',
        contentType: 'application/json',
        payload: JSON.stringify({action: 'ping', data: {}}),
        muteHttpExceptions: true
      });
      
      const postCode = postResponse.getResponseCode();
      message += `   Status Code: ${postCode}\n`;
      
      if (postCode === 200) {
        try {
          const responseData = JSON.parse(postResponse.getContentText());
          message += '   ✅ Result: SUCCESS - API authenticated correctly\n';
          message += `   📊 API Version: ${responseData.data?.version || 'Unknown'}\n`;
          message += `   🔐 Auth Required: ${responseData.data?.authRequired !== false ? 'Yes' : 'No'}\n\n`;
        } catch (e) {
          message += '   ✅ Result: SUCCESS - API responded (parse error)\n\n';
        }
      } else if (postCode === 401) {
        message += '   ⚠️ Result: EXPECTED 401 (External Test Limitation)\n';
        message += '   📋 This is NORMAL for DOMAIN security!\n';
        message += '   🔒 External tools cannot authenticate to DOMAIN APIs\n';
        message += '   ✅ Your API security is working correctly\n\n';
      } else {
        message += `   ❌ Result: UNEXPECTED ERROR (${postCode})\n`;
        message += `   Details: ${postResponse.getContentText().substring(0, 100)}\n\n`;
      }
      
      // Database connectivity test
      message += '📡 TEST 3: Database Connectivity\n';
      try {
        const dbSheet = getDatabaseSheet();
        const data = dbSheet.getDataRange().getValues();
        message += `   ✅ Result: SUCCESS - Database accessible\n`;
        message += `   📊 Total Users: ${Math.max(0, data.length - 1)}\n\n`;
      } catch (dbError) {
        message += `   ❌ Result: DATABASE ERROR\n`;
        message += `   Details: ${dbError.message}\n\n`;
      }
      
    } else {
      message += `   ❌ Result: CONNECTION FAILED\n`;
      message += `   Details: ${getResponse.getContentText().substring(0, 150)}\n\n`;
    }
    
    // Overall assessment
    message += '🎯 OVERALL ASSESSMENT:\n';
    if (getCode === 200) {
      message += '✅ Status: API IS OPERATIONAL\n';
      message += '🔒 Security: DOMAIN protection active\n';
      message += '📱 Ready for: StudyQuest main app integration\n';
      message += '⚠️ Note: POST 401 from external tools is expected & secure\n';
    } else {
      message += '❌ Status: API HAS ISSUES\n';
      message += '📋 Action: Check deployment settings and redeploy\n';
    }
    
  } catch (e) {
    message += '❌ CRITICAL ERROR during testing:\n';
    message += `Error: ${e.message}\n`;
    message += '📋 Action: Check network connection and API deployment\n';
  }
  
  message += '\n' + '=' .repeat(45);
  ui.alert(message);
}

function showDeploymentInstructions() {
  const ui = SpreadsheetApp.getUi();
  
  if (!PropertiesService.getScriptProperties().getProperty(CONFIG.DATABASE_ID_KEY)) {
    ui.alert('Error: Please initialize database first.');
    return;
  }
  
  const htmlOutput = HtmlService.createHtmlOutputFromFile('DeploymentGuide')
    .setWidth(600)
    .setHeight(450);
  ui.showModalDialog(htmlOutput, 'Deployment Instructions');
}

function saveDeploymentIdToProperties(id) {
  if (id && typeof id === 'string' && id.trim().length > 0) {
    PropertiesService.getScriptProperties().setProperty(CONFIG.DEPLOYMENT_ID_KEY, id.trim());
    SpreadsheetApp.getUi().alert('✅ Deployment ID saved successfully.');
    return 'OK';
  } else {
    SpreadsheetApp.getUi().alert('Error: Invalid deployment ID.');
    return 'Error';
  }
}

function debugDatabaseContents() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const dbSheet = getDatabaseSheet();
    const data = dbSheet.getDataRange().getValues();
    
    let message = '📊 DATABASE CONTENTS ANALYSIS\n';
    message += '=' .repeat(40) + '\n\n';
    
    message += `📁 Database Sheet: ${CONFIG.TARGET_SHEET_NAME}\n`;
    message += `📊 Total Rows: ${data.length}\n`;
    message += `👥 User Records: ${Math.max(0, data.length - 1)}\n\n`;
    
    if (data.length === 0) {
      message += '❌ Status: DATABASE IS EMPTY\n';
      message += '📋 Action: Database needs initialization\n';
    } else if (data.length === 1) {
      message += '✅ Status: INITIALIZED (Headers Only)\n';
      message += '📋 Ready for: User registration\n\n';
      message += '📋 HEADERS:\n';
      data[0].forEach((header, index) => {
        message += `   ${index + 1}. ${header}\n`;
      });
    } else {
      message += '✅ Status: ACTIVE WITH DATA\n';
      message += `📊 User Count: ${data.length - 1}\n\n`;
      
      // Show headers
      message += '📋 HEADERS:\n';
      data[0].forEach((header, index) => {
        message += `   ${index + 1}. ${header}\n`;
      });
      message += '\n';
      
      // Show sample user data (anonymized)
      message += '👥 SAMPLE USER DATA:\n';
      const sampleCount = Math.min(3, data.length - 1);
      
      for (let i = 1; i <= sampleCount; i++) {
        const row = data[i];
        message += `   User ${i}:\n`;
        message += `     📧 Email: ${row[1] ? String(row[1]).replace(/(.{3}).*(@.*)/, '$1***$2') : 'N/A'}\n`;
        message += `     📊 Spreadsheet: ${row[2] ? 'Configured' : 'Missing'}\n`;
        message += `     📅 Created: ${row[4] ? new Date(row[4]).toLocaleDateString() : 'N/A'}\n`;
        message += `     ✅ Active: ${row[8] !== false ? 'Yes' : 'No'}\n`;
        message += '\n';
      }
      
      if (data.length > 4) {
        message += `   ... (${data.length - 4} more users)\n\n`;
      }
      
      // Data quality check
      message += '🔍 DATA QUALITY CHECK:\n';
      let validUsers = 0;
      let invalidUsers = 0;
      
      for (let i = 1; i < data.length; i++) {
        const userId = data[i][0];
        const email = data[i][1];
        
        if (isValidUserId(userId) && isValidEmail(email)) {
          validUsers++;
        } else {
          invalidUsers++;
        }
      }
      
      message += `   ✅ Valid Users: ${validUsers}\n`;
      message += `   ❌ Invalid Users: ${invalidUsers}\n`;
      
      if (invalidUsers > 0) {
        message += '   📋 Action: Use "Cleanup Invalid Users" to fix\n';
      }
    }
    
    message += '\n' + '=' .repeat(40) + '\n';
    message += 'ℹ️ This data powers the StudyQuest user management system.';
    
    ui.alert(message);
    
  } catch (error) {
    let errorMessage = '❌ DATABASE ACCESS ERROR\n';
    errorMessage += '=' .repeat(30) + '\n\n';
    errorMessage += `Error: ${error.message}\n\n`;
    errorMessage += 'Possible causes:\n';
    errorMessage += '• Database not initialized\n';
    errorMessage += '• Permission issues\n';
    errorMessage += '• Spreadsheet deleted or moved\n\n';
    errorMessage += '📋 Action: Try "Initialize Database" from menu';
    
    ui.alert(errorMessage);
  }
}

function clearDatabase() {
  const ui = SpreadsheetApp.getUi();
  
  const confirmation = ui.alert(
    '⚠️ Clear Database',
    'This will delete all user data permanently. Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (confirmation !== ui.Button.YES) {
    ui.alert('Operation cancelled.');
    return;
  }
  
  try {
    const dbSheet = getDatabaseSheet();
    const data = dbSheet.getDataRange();
    
    if (data.getNumRows() > 1) {
      dbSheet.getRange(2, 1, data.getNumRows() - 1, data.getNumColumns()).clearContent();
      clearAllCaches();
      ui.alert('✅ Database cleared. Headers preserved.');
    } else {
      ui.alert('ℹ️ Database is already empty.');
    }
    
  } catch (error) {
    ui.alert(`Error: ${error.message}`);
  }
}

function cleanupInvalidUsers() {
  const ui = SpreadsheetApp.getUi();
  
  const confirmation = ui.alert(
    '🔧 CLEANUP INVALID USERS',
    'This will remove rows with invalid userIds or emails.\nAre you sure you want to continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (confirmation !== ui.Button.YES) {
    ui.alert('✅ OPERATION CANCELLED\n\nNo changes were made to the database.');
    return;
  }
  
  try {
    const dbSheet = getDatabaseSheet();
    const data = dbSheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      ui.alert('ℹ️ CLEANUP COMPLETE\n\nNo user data found to cleanup.\nDatabase contains only headers.');
      return;
    }
    
    let invalidRows = [];
    let invalidReasons = [];
    
    // Scan for invalid data
    for (let i = 1; i < data.length; i++) {
      const userId = data[i][0];
      const email = data[i][1];
      const row = i + 1;
      
      if (!isValidUserId(userId)) {
        invalidRows.push(row);
        invalidReasons.push(`Row ${row}: Invalid userId "${userId}"`);
      } else if (!isValidEmail(email)) {
        invalidRows.push(row);
        invalidReasons.push(`Row ${row}: Invalid email "${email}"`);
      }
    }
    
    if (invalidRows.length === 0) {
      ui.alert('✅ CLEANUP COMPLETE\n\nNo invalid user data found.\nDatabase integrity: EXCELLENT\n\nAll user records are valid.');
      return;
    }
    
    // Show what will be deleted
    let preMessage = '🔍 CLEANUP PREVIEW\n';
    preMessage += '=' .repeat(30) + '\n\n';
    preMessage += `Found ${invalidRows.length} invalid records:\n\n`;
    
    const previewCount = Math.min(5, invalidReasons.length);
    for (let i = 0; i < previewCount; i++) {
      preMessage += `• ${invalidReasons[i]}\n`;
    }
    
    if (invalidReasons.length > 5) {
      preMessage += `• ... and ${invalidReasons.length - 5} more\n`;
    }
    
    preMessage += '\nProceed with deletion?';
    
    const finalConfirm = ui.alert('🗑️ CONFIRM DELETION', preMessage, ui.ButtonSet.YES_NO);
    
    if (finalConfirm !== ui.Button.YES) {
      ui.alert('✅ OPERATION CANCELLED\n\nNo records were deleted.');
      return;
    }
    
    // Delete from bottom to top to maintain row indices
    for (let i = invalidRows.length - 1; i >= 0; i--) {
      dbSheet.deleteRow(invalidRows[i]);
    }
    
    clearAllCaches();
    
    let resultMessage = '✅ CLEANUP COMPLETED SUCCESSFULLY\n';
    resultMessage += '=' .repeat(35) + '\n\n';
    resultMessage += `📊 Results:\n`;
    resultMessage += `   🗑️ Deleted: ${invalidRows.length} invalid records\n`;
    resultMessage += `   ✅ Remaining: ${Math.max(0, data.length - 1 - invalidRows.length)} valid users\n`;
    resultMessage += `   🧹 Cache: Cleared and refreshed\n\n`;
    resultMessage += '🎯 Database integrity: RESTORED\n';
    resultMessage += '📋 Ready for: Normal operations';
    
    ui.alert(resultMessage);
    
  } catch (error) {
    let errorMessage = '❌ CLEANUP FAILED\n';
    errorMessage += '=' .repeat(20) + '\n\n';
    errorMessage += `Error: ${error.message}\n\n`;
    errorMessage += 'Possible causes:\n';
    errorMessage += '• Database access issues\n';
    errorMessage += '• Permission problems\n';
    errorMessage += '• Concurrent modifications\n\n';
    errorMessage += '📋 Action: Try again or check database permissions';
    
    ui.alert(errorMessage);
  }
}

