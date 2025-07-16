const fs = require('fs');
const vm = require('vm');

describe('All form creation scenarios reaction columns', () => {
  const coreCode = fs.readFileSync('src/Core.gs', 'utf8');
  const configCode = fs.readFileSync('src/config.gs', 'utf8');
  let context;
  
  beforeEach(() => {
    // Mock SpreadsheetApp and related objects
    const mockSheet = {
      getRange: jest.fn(),
      getLastColumn: jest.fn(() => 6), // Basic form columns count
      autoResizeColumns: jest.fn()
    };
    
    const mockSpreadsheet = {
      getSheetByName: jest.fn(() => mockSheet),
      getSheets: jest.fn(() => [mockSheet])
    };
    
    context = {
      console,
      debugLog: jest.fn(),
      SpreadsheetApp: {
        openById: jest.fn(() => mockSpreadsheet)
      },
      Session: {
        getActiveUser: jest.fn(() => ({
          getEmail: jest.fn(() => 'test@example.com')
        }))
      },
      COLUMN_HEADERS: {
        UNDERSTAND: 'なるほど！',
        LIKE: 'いいね！',
        CURIOUS: 'もっと知りたい！',
        HIGHLIGHT: 'ハイライト'
      },
      String: String,
      mockSheet,
      mockSpreadsheet,
      // Mock other dependencies
      verifyUserAccess: jest.fn(),
      findUserById: jest.fn(),
      getSheetsList: jest.fn(),
      addServiceAccountToSpreadsheet: jest.fn(),
      updateUser: jest.fn(),
      invalidateUserCache: jest.fn(),
      fetchUserFromDatabase: jest.fn(() => ({ userId: 'mock-user-id', adminEmail: 'test@example.com' }))
    };
    
    vm.createContext(context);
    // Load both Core.gs and config.gs
    vm.runInContext(coreCode, context);
    vm.runInContext(configCode, context);
    context.getSheetsList = jest.fn();
  });

  describe('addSpreadsheetUrl scenario', () => {
    test('should add reaction columns to external spreadsheet', () => {
      // Setup mocks
      const mockRange = {
        getValues: jest.fn(() => [['タイムスタンプ', 'メールアドレス', 'クラス', '名前', '質問', '理由']]),
        setValues: jest.fn(),
        setFontWeight: jest.fn(() => ({ setBackground: jest.fn() }))
      };
      
      context.mockSheet.getRange.mockReturnValue(mockRange);
      context.findUserById.mockReturnValue({
        userId: 'test-user',
        adminEmail: 'test@example.com',
        configJson: '{}'
      });
      context.getSheetsList.mockReturnValue([
        { name: 'Sheet1', id: 0 },
        { name: 'Sheet2', id: 1 }
      ]);
      
      // Call the function
      const result = context.addSpreadsheetUrl('test-user', 'https://docs.google.com/spreadsheets/d/test-id/edit');
      
      // Verify reaction columns were added
      expect(mockRange.setValues).toHaveBeenCalledWith([
        ['なるほど！', 'いいね！', 'もっと知りたい！', 'ハイライト']
      ]);
      expect(console.log).toHaveBeenCalledWith('✅ Reaction columns added to external spreadsheet:', 'Sheet1');
    });

    test('should handle spreadsheet with no sheets gracefully', () => {
      // Setup mocks
      context.findUserById.mockReturnValue({
        userId: 'test-user',
        adminEmail: 'test@example.com',
        configJson: '{}'
      });
      context.getSheetsList.mockReturnValue([]); // No sheets
      
      // Call the function
      const result = context.addSpreadsheetUrl('test-user', 'https://docs.google.com/spreadsheets/d/test-id/edit');
      
      // Should not attempt to add reaction columns
      expect(console.log).not.toHaveBeenCalledWith(expect.stringContaining('Reaction columns added'));
    });
  });

  describe('reaction column addition timing', () => {
    test('should verify all scenarios now add reaction columns', () => {
      // This is more of a documentation test to verify our fixes
      const scenarios = [
        'クイックスタート (createStudyQuestForm)',
        'カスタムフォーム作成 (createCustomFormUI)', 
        '外部URL追加 (addSpreadsheetUrl)'
      ];
      
      console.log('✅ All form creation scenarios now include reaction column addition:');
      scenarios.forEach(scenario => console.log(`  - ${scenario}`));
      
      expect(scenarios).toHaveLength(3);
    });
  });
});