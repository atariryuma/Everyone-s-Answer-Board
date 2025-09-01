/**
 * 高精度 Google Apps Script API モック (TypeScript版)
 * 2024年ベストプラクティス準拠 - リアルなエラー条件・動作を再現
 */

import type { 
  GoogleAppsScript 
} from 'google-apps-script';

// Mock data storage for realistic testing
class MockDataStore {
  private data = new Map<string, any>();
  private errorConditions = new Map<string, Error>();
  
  setValue(key: string, value: any): void {
    this.data.set(key, value);
  }
  
  getValue(key: string, defaultValue: any = null): any {
    return this.data.get(key) ?? defaultValue;
  }
  
  deleteValue(key: string): void {
    this.data.delete(key);
  }
  
  setErrorCondition(key: string, error: Error): void {
    this.errorConditions.set(key, error);
  }
  
  checkErrorCondition(key: string): void {
    const error = this.errorConditions.get(key);
    if (error) {
      this.errorConditions.delete(key);
      throw error;
    }
  }
  
  clear(): void {
    this.data.clear();
    this.errorConditions.clear();
  }
}

const mockDataStore = new MockDataStore();

// Realistic GAS Error Classes
class GASError extends Error {
  constructor(message: string, public code?: string) {
    super(message);
    this.name = 'GoogleAppsScriptError';
  }
}

// Enhanced SpreadsheetApp Mock with realistic behavior
const createSpreadsheetAppMock = () => {
  const mockSpreadsheet = (id: string) => ({
    getId: jest.fn(() => id),
    getName: jest.fn(() => mockDataStore.getValue(`spreadsheet_${id}_name`, `Test Spreadsheet ${id}`)),
    getUrl: jest.fn(() => `https://docs.google.com/spreadsheets/d/${id}/edit`),
    
    getSheetByName: jest.fn((name: string) => {
      mockDataStore.checkErrorCondition(`sheet_access_${name}`);
      
      const sheetExists = mockDataStore.getValue(`sheet_${name}_exists`, true);
      if (!sheetExists) {
        return null; // Realistic GAS behavior
      }
      
      return createMockSheet(name);
    }),
    
    getSheets: jest.fn(() => {
      const sheetNames = mockDataStore.getValue('sheet_names', ['Sheet1', 'Form Responses 1']);
      return sheetNames.map((name: string) => createMockSheet(name));
    }),
    
    insertSheet: jest.fn((name?: string) => {
      const sheetName = name || `Sheet${Date.now()}`;
      const currentSheets = mockDataStore.getValue('sheet_names', ['Sheet1']);
      mockDataStore.setValue('sheet_names', [...currentSheets, sheetName]);
      return createMockSheet(sheetName);
    }),
    
    deleteSheet: jest.fn((sheet: any) => {
      const sheetName = sheet.getName();
      const currentSheets = mockDataStore.getValue('sheet_names', ['Sheet1']);
      mockDataStore.setValue('sheet_names', currentSheets.filter((name: string) => name !== sheetName));
    })
  });

  const createMockSheet = (name: string) => {
    const getStoredData = (rowStart = 1, colStart = 1, numRows = 10, numCols = 5) => {
      const key = `sheet_${name}_data`;
      const defaultData = Array(numRows).fill(null).map(() => Array(numCols).fill(''));
      
      // Add realistic header data for first row
      if (rowStart === 1 && mockDataStore.getValue(`sheet_${name}_has_headers`, true)) {
        const headers = mockDataStore.getValue(`sheet_${name}_headers`, 
          ['タイムスタンプ', 'メールアドレス', '回答', '理由', 'クラス']);
        defaultData[0] = headers.slice(colStart - 1, colStart - 1 + numCols);
      }
      
      return mockDataStore.getValue(key, defaultData);
    };

    return {
      getName: jest.fn(() => name),
      getIndex: jest.fn(() => mockDataStore.getValue(`sheet_${name}_index`, 0)),
      
      getRange: jest.fn((row: number, col: number, numRows = 1, numCols = 1) => {
        // Validate range parameters (realistic GAS validation)
        if (row < 1 || col < 1) {
          throw new GASError('Range row and column indices must be positive integers', 'INVALID_ARGUMENT');
        }
        if (numRows < 1 || numCols < 1) {
          throw new GASError('Range must contain at least one cell', 'INVALID_ARGUMENT');
        }
        
        return createMockRange(name, row, col, numRows, numCols);
      }),
      
      getLastRow: jest.fn(() => {
        mockDataStore.checkErrorCondition(`sheet_access_${name}`);
        return mockDataStore.getValue(`sheet_${name}_lastRow`, 10);
      }),
      
      getLastColumn: jest.fn(() => {
        mockDataStore.checkErrorCondition(`sheet_access_${name}`);
        return mockDataStore.getValue(`sheet_${name}_lastColumn`, 5);
      }),
      
      appendRow: jest.fn((values: any[]) => {
        mockDataStore.checkErrorCondition(`sheet_write_${name}`);
        
        const lastRow = mockDataStore.getValue(`sheet_${name}_lastRow`, 10);
        mockDataStore.setValue(`sheet_${name}_lastRow`, lastRow + 1);
        
        const currentData = getStoredData();
        currentData.push(values);
        mockDataStore.setValue(`sheet_${name}_data`, currentData);
      }),
      
      deleteRow: jest.fn((rowIndex: number) => {
        mockDataStore.checkErrorCondition(`sheet_write_${name}`);
        
        if (rowIndex < 1) {
          throw new GASError('Row index must be positive', 'INVALID_ARGUMENT');
        }
        
        const lastRow = mockDataStore.getValue(`sheet_${name}_lastRow`, 10);
        if (lastRow > 1) {
          mockDataStore.setValue(`sheet_${name}_lastRow`, lastRow - 1);
        }
      }),
      
      insertRowAfter: jest.fn((afterRow: number) => {
        mockDataStore.checkErrorCondition(`sheet_write_${name}`);
        
        const lastRow = mockDataStore.getValue(`sheet_${name}_lastRow`, 10);
        mockDataStore.setValue(`sheet_${name}_lastRow`, lastRow + 1);
      }),
      
      clear: jest.fn(() => {
        mockDataStore.checkErrorCondition(`sheet_write_${name}`);
        mockDataStore.setValue(`sheet_${name}_data`, []);
        mockDataStore.setValue(`sheet_${name}_lastRow`, 0);
      })
    };
  };

  const createMockRange = (sheetName: string, row: number, col: number, numRows: number, numCols: number) => {
    return {
      getRow: jest.fn(() => row),
      getColumn: jest.fn(() => col),
      getNumRows: jest.fn(() => numRows),
      getNumColumns: jest.fn(() => numCols),
      
      getValues: jest.fn(() => {
        mockDataStore.checkErrorCondition(`range_access_${sheetName}`);
        
        const sheetData = mockDataStore.getValue(`sheet_${sheetName}_data`, []);
        const result: any[][] = [];
        
        for (let r = 0; r < numRows; r++) {
          const rowData: any[] = [];
          const actualRow = sheetData[row - 1 + r] || [];
          
          for (let c = 0; c < numCols; c++) {
            rowData.push(actualRow[col - 1 + c] || '');
          }
          result.push(rowData);
        }
        
        return result;
      }),
      
      setValues: jest.fn((values: any[][]) => {
        mockDataStore.checkErrorCondition(`range_write_${sheetName}`);
        
        // Realistic validation - array dimensions must match range
        if (values.length !== numRows) {
          throw new GASError(
            `The number of rows in the data (${values.length}) does not match the number of rows in the range (${numRows})`,
            'INVALID_ARGUMENT'
          );
        }
        
        values.forEach((rowValues, rowIndex) => {
          if (rowValues.length !== numCols) {
            throw new GASError(
              `Row ${rowIndex + 1}: The number of columns in the data (${rowValues.length}) does not match the number of columns in the range (${numCols})`,
              'INVALID_ARGUMENT'
            );
          }
        });
        
        // Update stored data
        let sheetData = mockDataStore.getValue(`sheet_${sheetName}_data`, []);
        
        // Ensure sheet data is large enough
        while (sheetData.length < row + numRows - 1) {
          sheetData.push([]);
        }
        
        for (let r = 0; r < numRows; r++) {
          const targetRowIndex = row - 1 + r;
          if (!sheetData[targetRowIndex]) {
            sheetData[targetRowIndex] = [];
          }
          
          // Ensure row is wide enough
          while (sheetData[targetRowIndex].length < col + numCols - 1) {
            sheetData[targetRowIndex].push('');
          }
          
          for (let c = 0; c < numCols; c++) {
            sheetData[targetRowIndex][col - 1 + c] = values[r][c];
          }
        }
        
        mockDataStore.setValue(`sheet_${sheetName}_data`, sheetData);
      }),
      
      getValue: jest.fn(() => {
        if (numRows !== 1 || numCols !== 1) {
          throw new GASError('Range must be a single cell for getValue()', 'INVALID_ARGUMENT');
        }
        
        const values = (createMockRange(sheetName, row, col, numRows, numCols).getValues as jest.Mock)();
        return values[0][0];
      }),
      
      setValue: jest.fn((value: any) => {
        if (numRows !== 1 || numCols !== 1) {
          throw new GASError('Range must be a single cell for setValue()', 'INVALID_ARGUMENT');
        }
        
        (createMockRange(sheetName, row, col, numRows, numCols).setValues as jest.Mock)([[value]]);
      }),
      
      clear: jest.fn(() => {
        mockDataStore.checkErrorCondition(`range_write_${sheetName}`);
        
        const emptyValues = Array(numRows).fill(null).map(() => Array(numCols).fill(''));
        (createMockRange(sheetName, row, col, numRows, numCols).setValues as jest.Mock)(emptyValues);
      })
    };
  };

  return {
    getActiveSpreadsheet: jest.fn(() => {
      mockDataStore.checkErrorCondition('activeSpreadsheet');
      const activeId = mockDataStore.getValue('activeSpreadsheetId', 'test-active-spreadsheet-id');
      return mockSpreadsheet(activeId);
    }),
    
    getActiveSheet: jest.fn(() => {
      mockDataStore.checkErrorCondition('activeSheet');
      const activeSheetName = mockDataStore.getValue('activeSheetName', 'Sheet1');
      return createMockSheet(activeSheetName);
    }),
    
    openById: jest.fn((id: string) => {
      mockDataStore.checkErrorCondition(`spreadsheet_${id}`);
      
      // Simulate permission error for invalid IDs
      if (!id || id === 'invalid-id') {
        throw new GASError('You do not have permission to access the requested document', 'PERMISSION_DENIED');
      }
      
      return mockSpreadsheet(id);
    }),
    
    create: jest.fn((name: string) => {
      const newId = `new-spreadsheet-${Date.now()}`;
      mockDataStore.setValue(`spreadsheet_${newId}_name`, name);
      return mockSpreadsheet(newId);
    })
  };
};

// Enhanced PropertiesService Mock
const createPropertiesServiceMock = () => {
  const createPropertiesStore = (storeType: string) => ({
    getProperty: jest.fn((key: string) => {
      mockDataStore.checkErrorCondition(`properties_${storeType}_read`);
      return mockDataStore.getValue(`${storeType}_${key}`, null);
    }),
    
    setProperty: jest.fn((key: string, value: string) => {
      mockDataStore.checkErrorCondition(`properties_${storeType}_write`);
      if (typeof value !== 'string') {
        throw new GASError('Property value must be a string', 'INVALID_ARGUMENT');
      }
      mockDataStore.setValue(`${storeType}_${key}`, value);
    }),
    
    deleteProperty: jest.fn((key: string) => {
      mockDataStore.checkErrorCondition(`properties_${storeType}_write`);
      mockDataStore.deleteValue(`${storeType}_${key}`);
    }),
    
    getProperties: jest.fn(() => {
      mockDataStore.checkErrorCondition(`properties_${storeType}_read`);
      // Return all properties for this store type
      const properties: {[key: string]: string} = {};
      // This would need more complex implementation for full coverage
      return properties;
    }),
    
    setProperties: jest.fn((properties: {[key: string]: string}) => {
      mockDataStore.checkErrorCondition(`properties_${storeType}_write`);
      Object.entries(properties).forEach(([key, value]) => {
        if (typeof value !== 'string') {
          throw new GASError('All property values must be strings', 'INVALID_ARGUMENT');
        }
        mockDataStore.setValue(`${storeType}_${key}`, value);
      });
    }),
    
    deleteAllProperties: jest.fn(() => {
      mockDataStore.checkErrorCondition(`properties_${storeType}_write`);
      // Clear all properties for this store type
      // This would need more complex implementation for full coverage
    })
  });

  return {
    getScriptProperties: jest.fn(() => createPropertiesStore('script')),
    getUserProperties: jest.fn(() => createPropertiesStore('user')),
    getDocumentProperties: jest.fn(() => createPropertiesStore('document'))
  };
};

// Global GAS API Mocks
const SpreadsheetApp = createSpreadsheetAppMock();
const PropertiesService = createPropertiesServiceMock();

// Mock Test Utilities
export const MockTestUtils = {
  /**
   * Set mock data for testing
   */
  setMockData: (key: string, value: any) => mockDataStore.setValue(key, value),
  
  /**
   * Get mock data
   */
  getMockData: (key: string, defaultValue?: any) => mockDataStore.getValue(key, defaultValue),
  
  /**
   * Clear all mock data
   */
  clearAllMockData: () => mockDataStore.clear(),
  
  /**
   * Simulate GAS API error conditions
   */
  simulateError: (condition: string, error: Error) => mockDataStore.setErrorCondition(condition, error),
  
  /**
   * Create realistic spreadsheet test data
   */
  createRealisticSpreadsheetData: (headers: string[], dataRows: any[][]) => {
    mockDataStore.setValue('sheet_Sheet1_headers', headers);
    mockDataStore.setValue('sheet_Sheet1_data', [headers, ...dataRows]);
    mockDataStore.setValue('sheet_Sheet1_lastRow', dataRows.length + 1);
    mockDataStore.setValue('sheet_Sheet1_lastColumn', headers.length);
  },
  
  /**
   * Simulate permission errors
   */
  simulatePermissionError: (resource: string) => {
    mockDataStore.setErrorCondition(resource, new GASError('Permission denied', 'PERMISSION_DENIED'));
  }
};

// Legacy compatibility - existing mock objects
const legacyMockObjects = {
  // 統一レスポンス関数 モック
  createSuccessResponse: jest.fn((data = null, message = null) => ({
    success: true,
    data,
    message,
    error: null,
  })),

  createErrorResponse: jest.fn((error, message = null, data = null) => ({
    success: false,
    data,
    message,
    error,
  })),

  createUnifiedResponse: jest.fn((success, data = null, message = null, error = null) => ({
    success,
    data,
    message,
    error,
  })),

  // 統一キャッシュマネージャー モック
  cacheManager: {
    get: jest.fn(),
    put: jest.fn(),
    remove: jest.fn(),
    clear: jest.fn(),
    getStats: jest.fn(() => ({ hits: 0, misses: 0 })),
  },

  // 名前空間オブジェクト モック
  User: {
    email: jest.fn(() => 'test@example.com'),
    info: jest.fn(() => ({
      userId: 'test-user-id',
      adminEmail: 'test@example.com',
      isActive: true,
    })),
  },

  Access: {
    check: jest.fn(() => ({ hasAccess: true, message: 'アクセス許可' })),
  },

  Deploy: {
    domain: jest.fn(() => ({
      currentDomain: 'example.com',
      deployDomain: 'example.com',
      isDomainMatch: true,
      webAppUrl: 'https://script.google.com/test',
    })),
    isUser: jest.fn(() => true),
  },

  DB: {
    createUser: jest.fn((userData) => userData),
    findUserByEmail: jest.fn((email) => ({
      userId: 'test-user-id',
      adminEmail: email,
      isActive: true,
    })),
    findUserById: jest.fn((userId) => ({
      userId,
      adminEmail: 'test@example.com',
      isActive: true,
    })),
  }
};

// Export everything for global scope
declare global {
  const SpreadsheetApp: typeof SpreadsheetApp;
  const PropertiesService: typeof PropertiesService;
  const MockTestUtils: typeof MockTestUtils;
  
  // Legacy compatibility
  const createSuccessResponse: typeof legacyMockObjects.createSuccessResponse;
  const createErrorResponse: typeof legacyMockObjects.createErrorResponse;
  const createUnifiedResponse: typeof legacyMockObjects.createUnifiedResponse;
  const cacheManager: typeof legacyMockObjects.cacheManager;
  const User: typeof legacyMockObjects.User;
  const Access: typeof legacyMockObjects.Access;
  const Deploy: typeof legacyMockObjects.Deploy;
  const DB: typeof legacyMockObjects.DB;
}

// Set global mocks
(global as any).SpreadsheetApp = SpreadsheetApp;
(global as any).PropertiesService = PropertiesService;
(global as any).MockTestUtils = MockTestUtils;

// Set legacy compatibility mocks
Object.entries(legacyMockObjects).forEach(([key, value]) => {
  (global as any)[key] = value;
});

export {
  SpreadsheetApp,
  PropertiesService,
  MockTestUtils,
  GASError
};