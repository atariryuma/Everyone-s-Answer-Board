/**
 * Jest Setup File
 * Global configuration and mocks for Google Apps Script environment
 */

// Extend Jest matchers
require('jest-extended');

// Global test timeout
jest.setTimeout(30000);

// Suppress console.log during tests unless explicitly needed
const originalConsoleLog = console.log;
const originalConsoleWarn = console.warn;
const originalConsoleError = console.error;

// Allow console.error for actual errors but suppress info logging
console.log = jest.fn();
console.warn = jest.fn();
console.error = originalConsoleError;

// Restore original console for debugging when needed
global.enableConsoleLogging = () => {
  console.log = originalConsoleLog;
  console.warn = originalConsoleWarn;
};

// Mock Google Apps Script globals
global.SpreadsheetApp = {
  openById: jest.fn(),
  getActiveSpreadsheet: jest.fn(),
  getActiveSheet: jest.fn(),
  create: jest.fn()
};

global.DriveApp = {
  getFileById: jest.fn(),
  createFile: jest.fn(),
  getFiles: jest.fn()
};

global.FormApp = {
  openByUrl: jest.fn(),
  create: jest.fn()
};

global.Session = {
  getActiveUser: jest.fn(() => ({
    getEmail: jest.fn(() => 'test@example.com')
  })),
  getScriptTimeZone: jest.fn(() => 'Asia/Tokyo')
};

global.Utilities = {
  getUuid: jest.fn(() => 'test-uuid-123'),
  base64Encode: jest.fn(),
  base64Decode: jest.fn(),
  computeDigest: jest.fn(),
  formatDate: jest.fn()
};

global.PropertiesService = {
  getScriptProperties: jest.fn(() => ({
    getProperty: jest.fn(),
    setProperty: jest.fn(),
    deleteProperty: jest.fn(),
    getProperties: jest.fn(() => ({}))
  })),
  getUserProperties: jest.fn(() => ({
    getProperty: jest.fn(),
    setProperty: jest.fn(),
    deleteProperty: jest.fn()
  }))
};

global.CacheService = {
  getScriptCache: jest.fn(() => ({
    get: jest.fn(),
    put: jest.fn(),
    remove: jest.fn(),
    removeAll: jest.fn()
  })),
  getUserCache: jest.fn(() => ({
    get: jest.fn(),
    put: jest.fn(),
    remove: jest.fn()
  }))
};

global.HtmlService = {
  createTemplateFromFile: jest.fn(),
  createHtmlOutput: jest.fn(),
  createHtmlOutputFromFile: jest.fn(),
  XFrameOptionsMode: {
    ALLOWALL: 'ALLOWALL',
    SAMEORIGIN: 'SAMEORIGIN',
    DENY: 'DENY'
  }
};

global.ContentService = {
  createTextOutput: jest.fn(() => ({
    setMimeType: jest.fn(),
    getContent: jest.fn()
  })),
  MimeType: {
    JSON: 'application/json',
    XML: 'application/xml',
    TEXT: 'text/plain'
  }
};

global.UrlFetchApp = {
  fetch: jest.fn(),
  getRequest: jest.fn()
};

// Custom matchers for Google Apps Script testing
expect.extend({
  toBeValidUserId(received) {
    const isValid = typeof received === 'string' && received.length > 0;
    return {
      message: () => `expected ${received} to be a valid user ID`,
      pass: isValid
    };
  },
  
  toBeValidEmail(received) {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    const isValid = emailRegex.test(received);
    return {
      message: () => `expected ${received} to be a valid email address`,
      pass: isValid
    };
  },
  
  toBeValidConfigJSON(received) {
    try {
      const parsed = typeof received === 'string' ? JSON.parse(received) : received;
      const hasRequiredFields = parsed && 
        typeof parsed === 'object' &&
        'setupStatus' in parsed;
      
      return {
        message: () => `expected ${received} to be valid configJSON`,
        pass: hasRequiredFields
      };
    } catch (error) {
      return {
        message: () => `expected ${received} to be valid JSON`,
        pass: false
      };
    }
  }
});

// Mock performance APIs for performance testing
global.performance = {
  now: jest.fn(() => Date.now()),
  mark: jest.fn(),
  measure: jest.fn()
};

// Reset all mocks utility function
global.resetAllMocks = () => {
  jest.clearAllMocks();
  
  // Reset all Google Apps Script API mocks to default state
  global.Session.getActiveUser.mockReturnValue({
    getEmail: jest.fn(() => 'test@example.com')
  });
  
  global.PropertiesService.getScriptProperties.mockReturnValue({
    getProperty: jest.fn(),
    setProperty: jest.fn(),
    deleteProperty: jest.fn(),
    getProperties: jest.fn(() => ({}))
  });
  
  global.CacheService.getScriptCache.mockReturnValue({
    get: jest.fn(),
    put: jest.fn(),
    remove: jest.fn(),
    removeAll: jest.fn()
  });
  
  global.SpreadsheetApp.openById.mockReturnValue({
    getSheetByName: jest.fn(() => ({
      getDataRange: jest.fn(() => ({
        getValues: jest.fn(() => [])
      })),
      getRange: jest.fn(() => ({
        getValues: jest.fn(() => []),
        setValues: jest.fn()
      }))
    }))
  });
};

// Global test utilities
global.testUtils = {
  createMockUser: (overrides = {}) => ({
    userId: 'test-user-123',
    userEmail: 'test@example.com',
    isActive: true,
    configJson: JSON.stringify({
      setupStatus: 'pending',
      appPublished: false,
      ...overrides.config
    }),
    lastModified: new Date().toISOString(),
    ...overrides
  }),
  
  createMockConfig: (overrides = {}) => ({
    setupStatus: 'pending',
    appPublished: false,
    spreadsheetId: 'test-sheet-id',
    sheetName: 'テストシート',
    displaySettings: { showNames: true, showReactions: true },
    ...overrides
  }),
  
  createMockSheetData: (count = 3) => {
    return Array(count).fill().map((_, i) => ({
      id: i + 1,
      timestamp: new Date().toISOString(),
      email: `student${i + 1}@example.com`,
      class: '3年A組',
      name: `テストユーザー${i + 1}`,
      answer: `テスト回答${i + 1}`,
      reason: `テスト理由${i + 1}`,
      reactions: { understand: 0, like: 0, curious: 0 },
      highlight: false
    }));
  }
};

// Cleanup after each test
afterEach(() => {
  jest.clearAllMocks();
});