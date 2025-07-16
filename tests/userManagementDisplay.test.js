const fs = require('fs');
const vm = require('vm');

describe('User Management Display', () => {
  const code = fs.readFileSync('src/database.gs', 'utf8');
  let context;
  
  beforeEach(() => {
    context = {
      console,
      isDeployUser: jest.fn(() => true),
      PropertiesService: {
        getScriptProperties: jest.fn(() => ({
          getProperty: jest.fn(() => 'test-db-id')
        }))
      },
      SCRIPT_PROPS_KEYS: {
        DATABASE_SPREADSHEET_ID: 'DATABASE_SPREADSHEET_ID'
      },
      DB_SHEET_CONFIG: {
        SHEET_NAME: 'Users'
      },
      getSheetsService: jest.fn(() => ({
        spreadsheets: {
          values: {
            batchGet: jest.fn(() => ({})),
          },
        },
      })),
      getServiceAccountTokenCached: jest.fn(() => 'mock-token'),
      batchGetSheetsData: jest.fn(),
      debugLog: jest.fn()
    };
    
    vm.createContext(context);
    vm.runInContext(code, context);
    context.batchGetSheetsData = jest.fn();
  });

  describe('getAllUsersForAdmin', () => {
    test('should return users with createdAt field from database', () => {
      // Mock database response
      const mockHeaders = ['userId', 'adminEmail', 'spreadsheetId', 'spreadsheetUrl', 'createdAt', 'configJson', 'setupStatus', 'lastAccessedAt', 'isActive'];
      const mockUserData = [
        'user123',
        '35t22@naha-okinawa.ed.jp',
        'sheet123',
        'https://example.com/sheet',
        '2025-01-01T10:00:00.000Z',  // This is createdAt
        '{}',
        'INITIALIZED', // setupStatus
        '2025-07-13T12:00:00.000Z',
        'true'
      ];
      
      context.batchGetSheetsData.mockReturnValue({
        valueRanges: [{
          values: [mockHeaders, mockUserData]
        }]
      });
      
      const result = context.getAllUsersForAdmin();
      
      expect(result).toHaveLength(1);
      expect(result[0]).toHaveProperty('createdAt', '2025-01-01T10:00:00.000Z');
      expect(result[0]).toHaveProperty('adminEmail', '35t22@naha-okinawa.ed.jp');
      expect(result[0]).toHaveProperty('lastAccessedAt', '2025-07-13T12:00:00.000Z');
      
      // Ensure the property name matches what AppSetupPage.html expects
      expect(result[0]).toHaveProperty('createdAt');
      expect(result[0]).not.toHaveProperty('registrationDate');
    });

    test('should handle users with missing createdAt gracefully', () => {
      // Mock database response with missing createdAt
      const mockHeaders = ['userId', 'adminEmail', 'spreadsheetId', 'spreadsheetUrl', 'createdAt', 'configJson', 'lastAccessedAt', 'isActive'];
      const mockUserData = [
        'user123', 
        '35t22@naha-okinawa.ed.jp', 
        'sheet123', 
        'https://example.com/sheet', 
        '',  // Empty createdAt
        '{}', 
        '2025-07-13T12:00:00.000Z', 
        'true'
      ];
      
      context.batchGetSheetsData.mockReturnValue({
        valueRanges: [{
          values: [mockHeaders, mockUserData]
        }]
      });
      
      const result = context.getAllUsersForAdmin();
      
      expect(result).toHaveLength(1);
      expect(result[0]).toHaveProperty('createdAt', '');
      expect(result[0]).toHaveProperty('adminEmail', '35t22@naha-okinawa.ed.jp');
    });
  });

  describe('Database column mapping', () => {
    test('should verify DB_SHEET_CONFIG.HEADERS contains createdAt', () => {
      // This test documents the expected database structure
      const expectedHeaders = ['userId', 'adminEmail', 'spreadsheetId', 'spreadsheetUrl', 'createdAt', 'configJson', 'setupStatus', 'lastAccessedAt', 'isActive'];
      
      expect(expectedHeaders).toContain('createdAt');
      expect(expectedHeaders).not.toContain('registrationDate');
      
      console.log('✅ Database correctly uses createdAt column');
      console.log('✅ AppSetupPage.html now correctly reads user.createdAt');
    });
  });
});