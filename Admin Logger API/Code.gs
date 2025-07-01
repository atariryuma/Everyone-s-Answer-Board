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
  
  // DOMAIN認証チェック
  const authResult = validateDomainAccess();
  if (!authResult.valid) {
    console.log(`❌ API Access Denied: ${authResult.reason}`);
    return {
      success: false,
      error: 'DOMAIN_ACCESS_DENIED',
      message: authResult.reason
    };
  }
  
  // Log API access for monitoring
  console.log(`API Request: ${action} from ${requestUser || 'anonymous'} (Domain: ${authResult.domain})`);
  
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
          version: '2.0'
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

  let message = 'Current Configuration:\n\n';
  
  if (dbSheetId) {
    message += `✅ Database: Configured\n   (Spreadsheet ID: ${dbSheetId})\n\n`;
  } else {
    message += '❌ Database: Not configured\n\n';
  }

  // Templates are no longer used - removed for optimization

  if (deploymentId) {
    const webAppUrl = deploymentId.startsWith('https://') 
      ? deploymentId 
      : `https://script.google.com/macros/s/${deploymentId}/exec`;
    message += `✅ API Deployment: Active\n   URL: ${webAppUrl}\n\n`;
    message += '🔒 Security Info:\n• Admin privileges required\n• URL is confidential\n• All API calls are logged\n';
  } else {
    message += '❌ API Deployment: Not configured\n';
  }

  ui.alert(message);
}

function testDeployment() {
  const ui = SpreadsheetApp.getUi();
  const properties = PropertiesService.getScriptProperties();
  const deploymentId = properties.getProperty(CONFIG.DEPLOYMENT_ID_KEY);
  
  if (!deploymentId) {
    ui.alert('❌ No deployment ID configured. Please deploy API first.');
    return;
  }
  
  const webAppUrl = deploymentId.startsWith('https://') 
    ? deploymentId 
    : `https://script.google.com/macros/s/${deploymentId}/exec`;
  
  try {
    // Test GET
    const getResponse = UrlFetchApp.fetch(webAppUrl, {
      method: 'GET',
      muteHttpExceptions: true
    });
    
    let message = `Deployment Test Results:\n\nURL: ${webAppUrl}\nGET Test: ${getResponse.getResponseCode()}\n\n`;
    
    if (getResponse.getResponseCode() === 200) {
      message += '✅ GET connection successful\n\n';
      
      // Test POST
      const postResponse = UrlFetchApp.fetch(webAppUrl, {
        method: 'POST',
        contentType: 'application/json',
        payload: JSON.stringify({action: 'ping', data: {}}),
        muteHttpExceptions: true
      });
      
      message += `POST Test: ${postResponse.getResponseCode()}\n`;
      
      if (postResponse.getResponseCode() === 200) {
        message += '✅ API operational';
      } else {
        message += `❌ POST failed: ${postResponse.getContentText().substring(0, 100)}`;
      }
      
    } else {
      message += `❌ Connection failed\nDetails: ${getResponse.getContentText().substring(0, 200)}`;
    }
    
    ui.alert(message);
    
  } catch (e) {
    ui.alert(`Test error: ${e.message}`);
  }
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
    
    let message = `Database Contents:\n\nTotal rows: ${data.length}\n\n`;
    
    if (data.length === 0) {
      message += '❌ Database is empty';
    } else if (data.length === 1) {
      message += '✅ Headers only (no user data)\n';
      message += `Headers: ${JSON.stringify(data[0])}`;
    } else {
      message += `✅ Headers + ${data.length - 1} user records\n\n`;
      message += `Headers: ${JSON.stringify(data[0])}\n\n`;
      
      for (let i = 1; i < Math.min(data.length, 4); i++) {
        message += `Row ${i}: ${JSON.stringify(data[i])}\n`;
      }
      
      if (data.length > 4) {
        message += `... (${data.length - 4} more rows)`;
      }
    }
    
    ui.alert(message);
    
  } catch (error) {
    ui.alert(`Error: ${error.message}`);
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
    '🔧 Cleanup Invalid Users',
    'Remove rows with empty, null, or undefined userIds?',
    ui.ButtonSet.YES_NO
  );
  
  if (confirmation !== ui.Button.YES) {
    ui.alert('Operation cancelled.');
    return;
  }
  
  try {
    const dbSheet = getDatabaseSheet();
    const data = dbSheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      ui.alert('ℹ️ No data to cleanup.');
      return;
    }
    
    let invalidRows = [];
    
    for (let i = 1; i < data.length; i++) {
      const userId = data[i][0];
      if (!isValidUserId(userId)) {
        invalidRows.push(i + 1);
      }
    }
    
    if (invalidRows.length === 0) {
      ui.alert('✅ No invalid user data found.');
      return;
    }
    
    // Delete from bottom to top to maintain row indices
    for (let i = invalidRows.length - 1; i >= 0; i--) {
      dbSheet.deleteRow(invalidRows[i]);
    }
    
    clearAllCaches();
    ui.alert(`✅ Cleaned up ${invalidRows.length} invalid user records.`);
    
  } catch (error) {
    ui.alert(`Error: ${error.message}`);
  }
}

