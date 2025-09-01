/**
 * Jest Global Setup Configuration
 * TypeScript対応の高度なテスト環境セットアップ
 */

import 'jest-extended';

// Global test configuration
declare global {
  namespace jest {
    interface Matchers<R> {
      toBeValidEmail(): R;
      toBeValidSpreadsheetId(): R;
      toBeValidUserId(): R;
      toHaveValidColumnMapping(): R;
      toThrowGASError(expectedError?: string | RegExp): R;
    }
  }
}

// Custom matchers for Google Apps Script specific validations
expect.extend({
  /**
   * カスタムマッチャー: 有効なメールアドレスかチェック
   */
  toBeValidEmail(received: string) {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    const pass = typeof received === 'string' && emailRegex.test(received);
    
    if (pass) {
      return {
        message: () => `expected ${received} not to be a valid email address`,
        pass: true
      };
    } else {
      return {
        message: () => `expected ${received} to be a valid email address`,
        pass: false
      };
    }
  },

  /**
   * カスタムマッチャー: 有効なスプレッドシートIDかチェック
   */
  toBeValidSpreadsheetId(received: string) {
    const spreadsheetIdRegex = /^[a-zA-Z0-9-_]{44}$/;
    const pass = typeof received === 'string' && spreadsheetIdRegex.test(received);
    
    if (pass) {
      return {
        message: () => `expected ${received} not to be a valid spreadsheet ID`,
        pass: true
      };
    } else {
      return {
        message: () => `expected ${received} to be a valid spreadsheet ID`,
        pass: false
      };
    }
  },

  /**
   * カスタムマッチャー: 有効なユーザーIDかチェック（UUID形式）
   */
  toBeValidUserId(received: string) {
    const uuidRegex = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;
    const pass = typeof received === 'string' && uuidRegex.test(received);
    
    if (pass) {
      return {
        message: () => `expected ${received} not to be a valid UUID user ID`,
        pass: true
      };
    } else {
      return {
        message: () => `expected ${received} to be a valid UUID user ID`,
        pass: false
      };
    }
  },

  /**
   * カスタムマッチャー: 有効な列マッピングオブジェクトかチェック
   */
  toHaveValidColumnMapping(received: any) {
    const isObject = typeof received === 'object' && received !== null;
    const hasRequiredFields = isObject && 
      'answer' in received && 
      typeof received.answer === 'number';
    
    const pass = hasRequiredFields;
    
    if (pass) {
      return {
        message: () => `expected ${JSON.stringify(received)} not to have valid column mapping`,
        pass: true
      };
    } else {
      return {
        message: () => `expected ${JSON.stringify(received)} to have valid column mapping with at least answer field`,
        pass: false
      };
    }
  },

  /**
   * カスタムマッチャー: GAS特有のエラーをチェック
   */
  toThrowGASError(received: () => void, expectedError?: string | RegExp) {
    let thrown = false;
    let error: Error | null = null;
    
    try {
      received();
    } catch (e) {
      thrown = true;
      error = e as Error;
    }
    
    if (!thrown) {
      return {
        message: () => `expected function to throw a GAS error`,
        pass: false
      };
    }
    
    if (expectedError) {
      const messageMatches = typeof expectedError === 'string' 
        ? error!.message.includes(expectedError)
        : expectedError.test(error!.message);
        
      if (messageMatches) {
        return {
          message: () => `expected function not to throw error "${error!.message}"`,
          pass: true
        };
      } else {
        return {
          message: () => `expected function to throw error matching "${expectedError}", but got "${error!.message}"`,
          pass: false
        };
      }
    }
    
    return {
      message: () => `expected function not to throw any error, but it threw "${error!.message}"`,
      pass: true
    };
  }
});

// Global test environment configuration
beforeEach(() => {
  // Reset console mocks
  jest.clearAllMocks();
  
  // Reset global Date to avoid test instability
  jest.useFakeTimers().setSystemTime(new Date('2024-01-01T12:00:00Z'));
});

afterEach(() => {
  // Clean up timers
  jest.useRealTimers();
  
  // Reset all mocks
  jest.resetAllMocks();
});

// Enhanced console output for debugging
const originalConsoleError = console.error;
console.error = (...args: any[]) => {
  // Filter out expected test errors to reduce noise
  if (args.some(arg => 
    typeof arg === 'string' && 
    (arg.includes('Warning: ') || arg.includes('[EXPECTED TEST ERROR]'))
  )) {
    return;
  }
  originalConsoleError.apply(console, args);
};

// Performance monitoring for tests
let testStartTime: number;
beforeEach(() => {
  testStartTime = Date.now();
});

afterEach(() => {
  const testEndTime = Date.now();
  const duration = testEndTime - testStartTime;
  
  // Warn about slow tests (> 5 seconds)
  if (duration > 5000) {
    console.warn(`⚠️  Slow test detected: ${expect.getState().currentTestName} took ${duration}ms`);
  }
});

// Global error handler for unhandled promise rejections
process.on('unhandledRejection', (reason, promise) => {
  console.error('Unhandled promise rejection in test:', reason);
  throw reason;
});

export {};