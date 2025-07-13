const fs = require('fs');
const vm = require('vm');

describe('Reaction columns functionality', () => {
  const code = fs.readFileSync('src/Core.gs', 'utf8');
  let context;
  
  beforeEach(() => {
    // Mock SpreadsheetApp and related objects
    const mockSheet = {
      getRange: jest.fn(),
      getLastColumn: jest.fn(),
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
      COLUMN_HEADERS: {
        UNDERSTAND: 'なるほど！',
        LIKE: 'いいね！',
        CURIOUS: 'もっと知りたい！',
        HIGHLIGHT: 'ハイライト'
      },
      String: String,
      mockSheet,
      mockSpreadsheet
    };
    
    vm.createContext(context);
    vm.runInContext(code, context);
  });

  describe('addReactionColumnsToSpreadsheet', () => {
    test('should skip adding headers that already exist', () => {
      // Mock existing headers including reaction columns
      const existingHeaders = [
        'タイムスタンプ', 'メールアドレス', 'クラス', '名前', 
        '今日のテーマについて、あなたの考えや意見を聞かせてください',
        'そう考える理由や体験があれば教えてください（任意）',
        'なるほど！', 'いいね！', 'もっと知りたい！', 'ハイライト'
      ];
      
      const mockRange = {
        getValues: jest.fn(() => [existingHeaders]),
        setValues: jest.fn(),
        setFontWeight: jest.fn(() => ({ setBackground: jest.fn() }))
      };
      
      context.mockSheet.getRange.mockReturnValue(mockRange);
      context.mockSheet.getLastColumn.mockReturnValue(existingHeaders.length);
      
      // Call the function
      context.addReactionColumnsToSpreadsheet('test-sheet-id', 'Sheet1');
      
      // Verify no new headers were added
      expect(mockRange.setValues).not.toHaveBeenCalled();
      expect(console.log).toHaveBeenCalledWith('All reaction columns already exist, skipping addition');
    });

    test('should add only missing reaction headers', () => {
      // Mock existing headers without reaction columns
      const existingHeaders = [
        'タイムスタンプ', 'メールアドレス', 'クラス', '名前', 
        '今日のテーマについて、あなたの考えや意見を聞かせてください',
        'そう考える理由や体験があれば教えてください（任意）'
      ];
      
      const mockRange = {
        getValues: jest.fn(() => [existingHeaders]),
        setValues: jest.fn(),
        setFontWeight: jest.fn(() => ({ setBackground: jest.fn() }))
      };
      
      context.mockSheet.getRange.mockReturnValue(mockRange);
      context.mockSheet.getLastColumn.mockReturnValue(existingHeaders.length);
      
      // Call the function
      context.addReactionColumnsToSpreadsheet('test-sheet-id', 'Sheet1');
      
      // Verify all reaction headers were added
      expect(mockRange.setValues).toHaveBeenCalledWith([
        ['なるほど！', 'いいね！', 'もっと知りたい！', 'ハイライト']
      ]);
    });

    test('should add only partially missing headers', () => {
      // Mock existing headers with some reaction columns
      const existingHeaders = [
        'タイムスタンプ', 'メールアドレス', 'クラス', '名前', 
        '今日のテーマについて、あなたの考えや意見を聞かせてください',
        'そう考える理由や体験があれば教えてください（任意）',
        'なるほど！', 'いいね！'  // Only first two reaction columns exist
      ];
      
      const mockRange = {
        getValues: jest.fn(() => [existingHeaders]),
        setValues: jest.fn(),
        setFontWeight: jest.fn(() => ({ setBackground: jest.fn() }))
      };
      
      context.mockSheet.getRange.mockReturnValue(mockRange);
      context.mockSheet.getLastColumn.mockReturnValue(existingHeaders.length);
      
      // Call the function
      context.addReactionColumnsToSpreadsheet('test-sheet-id', 'Sheet1');
      
      // Verify only missing headers were added
      expect(mockRange.setValues).toHaveBeenCalledWith([
        ['もっと知りたい！', 'ハイライト']
      ]);
    });
  });
});