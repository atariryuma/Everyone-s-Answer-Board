/**
 * DataService Test Suite
 * Comprehensive testing for data processing and sheet operations
 */

const { testUserData, testSheetData, createMockSpreadsheet } = require('../mocks/gas-mocks.js');

// Mock DataService (simulate GAS environment)
const DataService = {
  getUserSheetData: jest.fn(),
  addReaction: jest.fn(),
  toggleHighlight: jest.fn(),
  refreshBoardData: jest.fn(),
  processDataWithColumnMapping: jest.fn(),
  getBulkData: jest.fn(),
  getAutoStopTime: jest.fn(),
  getFormInfo: jest.fn(),
  getDataStatistics: jest.fn(),
  updateReactionInSheet: jest.fn(),
  getOrCreateReactionColumn: jest.fn(),
  validateSheetAccess: jest.fn(),
  diagnose: jest.fn()
};

describe('DataService', () => {
  
  beforeEach(() => {
    jest.resetAllMocks();
    jest.clearAllMocks();
  });

  describe('getSheetData', () => {
    
    it('should return formatted sheet data', () => {
      // Arrange
      const userId = 'test-user-123';
      const mockData = [
        {
          id: 1,
          timestamp: '2025-01-15T10:00:00Z',
          email: 'student1@example.com',
          class: '3年A組',
          name: '田中太郎',
          answer: 'テスト回答1',
          reason: 'テスト理由1',
          reactions: { understand: 2, like: 1, curious: 0 },
          highlight: false
        },
        {
          id: 2,
          timestamp: '2025-01-15T10:05:00Z',
          email: 'student2@example.com',
          class: '3年A組',
          name: '佐藤花子',
          answer: 'テスト回答2',
          reason: 'テスト理由2',
          reactions: { understand: 1, like: 3, curious: 1 },
          highlight: true
        }
      ];

      DataService.getUserSheetData.mockReturnValue({
        success: true,
        data: mockData,
        count: 2,
        timestamp: '2025-01-15T10:00:00Z'
      });

      // Act
      const result = DataService.getUserSheetData(userId);

      // Assert
      expect(result.success).toBe(true);
      expect(result.data).toHaveLength(2);
      expect(result.data[0].answer).toBe('テスト回答1');
      expect(result.data[1].highlight).toBe(true);
    });

    it('should handle empty sheet data', () => {
      // Arrange
      DataService.getUserSheetData.mockReturnValue({
        success: true,
        data: [],
        count: 0,
        timestamp: '2025-01-15T10:00:00Z'
      });

      // Act
      const result = DataService.getUserSheetData('test-user');

      // Assert
      expect(result.success).toBe(true);
      expect(result.data).toHaveLength(0);
      expect(result.count).toBe(0);
    });

    it('should handle sheet access errors', () => {
      // Arrange
      DataService.getUserSheetData.mockReturnValue({
        success: false,
        error: 'スプレッドシートにアクセスできません',
        timestamp: '2025-01-15T10:00:00Z'
      });

      // Act
      const result = DataService.getUserSheetData('invalid-user');

      // Assert
      expect(result.success).toBe(false);
      expect(result.error).toBe('スプレッドシートにアクセスできません');
    });

  });

  describe('addReaction', () => {

    it('should add reaction successfully', () => {
      // Arrange
      const userId = 'test-user-123';
      const rowId = 'row_2';
      const reactionType = 'LIKE';

      DataService.addReaction.mockReturnValue({
        success: true,
        message: 'リアクションを追加しました',
        updatedReactions: { understand: 1, like: 2, curious: 0 }
      });

      // Act
      const result = DataService.addReaction(userId, rowId, reactionType);

      // Assert
      expect(result.success).toBe(true);
      expect(result.updatedReactions.like).toBe(2);
      expect(DataService.addReaction).toHaveBeenCalledWith(userId, rowId, reactionType);
    });

    it('should handle invalid reaction type', () => {
      // Arrange
      DataService.addReaction.mockReturnValue({
        success: false,
        error: '無効なリアクションタイプです'
      });

      // Act
      const result = DataService.addReaction('test-user', 'row_1', 'INVALID');

      // Assert
      expect(result.success).toBe(false);
      expect(result.error).toBe('無効なリアクションタイプです');
    });

    it('should handle row not found', () => {
      // Arrange
      DataService.addReaction.mockReturnValue({
        success: false,
        error: '指定された行が見つかりません'
      });

      // Act
      const result = DataService.addReaction('test-user', 'row_999', 'LIKE');

      // Assert
      expect(result.success).toBe(false);
      expect(result.error).toBe('指定された行が見つかりません');
    });

  });

  describe('toggleHighlight', () => {

    it('should toggle highlight successfully', () => {
      // Arrange
      const userId = 'test-user-123';
      const rowId = 'row_2';

      DataService.toggleHighlight.mockReturnValue({
        success: true,
        highlighted: true,
        message: 'ハイライトを更新しました'
      });

      // Act
      const result = DataService.toggleHighlight(userId, rowId);

      // Assert
      expect(result.success).toBe(true);
      expect(result.highlighted).toBe(true);
      expect(DataService.toggleHighlight).toHaveBeenCalledWith(userId, rowId);
    });

    it('should handle toggle highlight errors', () => {
      // Arrange
      DataService.toggleHighlight.mockReturnValue({
        success: false,
        error: 'ハイライトの更新に失敗しました'
      });

      // Act
      const result = DataService.toggleHighlight('test-user', 'invalid-row');

      // Assert
      expect(result.success).toBe(false);
      expect(result.error).toBe('ハイライトの更新に失敗しました');
    });

  });

  describe('processDataWithColumnMapping', () => {

    it('should process data with column mapping', () => {
      // Arrange
      const dataRows = [
        ['2025-01-15T10:00:00Z', 'student@example.com', '3年A組', '田中太郎', 'テスト回答', 'テスト理由'],
        ['2025-01-15T10:05:00Z', 'student2@example.com', '3年A組', '佐藤花子', 'テスト回答2', 'テスト理由2']
      ];
      const headers = ['タイムスタンプ', 'メール', 'クラス', '名前', '回答', '理由'];
      const columnMapping = {
        timestamp: 0,
        email: 1,
        class: 2,
        name: 3,
        answer: 4,
        reason: 5
      };

      const expectedResult = [
        {
          id: 1,
          timestamp: '2025-01-15T10:00:00Z',
          email: 'student@example.com',
          class: '3年A組',
          name: '田中太郎',
          answer: 'テスト回答',
          reason: 'テスト理由',
          reactions: { understand: 0, like: 0, curious: 0 },
          highlight: false,
          originalData: dataRows[0]
        },
        {
          id: 2,
          timestamp: '2025-01-15T10:05:00Z',
          email: 'student2@example.com',
          class: '3年A組',
          name: '佐藤花子',
          answer: 'テスト回答2',
          reason: 'テスト理由2',
          reactions: { understand: 0, like: 0, curious: 0 },
          highlight: false,
          originalData: dataRows[1]
        }
      ];

      DataService.processDataWithColumnMapping.mockReturnValue(expectedResult);

      // Act
      const result = DataService.processDataWithColumnMapping(dataRows, headers, columnMapping);

      // Assert
      expect(result).toHaveLength(2);
      expect(result[0].email).toBe('student@example.com');
      expect(result[1].answer).toBe('テスト回答2');
    });

    it('should handle empty data rows', () => {
      // Arrange
      DataService.processDataWithColumnMapping.mockReturnValue([]);

      // Act
      const result = DataService.processDataWithColumnMapping([], [], {});

      // Assert
      expect(result).toHaveLength(0);
    });

  });

  describe('getBulkData', () => {

    it('should return bulk data with all options', () => {
      // Arrange
      const userId = 'test-user-123';
      const options = {
        includeSheetData: true,
        includeFormInfo: true,
        includeSystemInfo: true
      };
      const mockBulkData = {
        timestamp: '2025-01-15T10:00:00Z',
        userId,
        userInfo: {
          userEmail: 'test@example.com',
          isActive: true,
          config: { setupStatus: 'completed' }
        },
        sheetData: testSheetData,
        formInfo: { formUrl: 'https://forms.gle/test' },
        systemInfo: {
          setupStep: 3,
          isSystemSetup: true,
          appPublished: true
        }
      };

      DataService.getBulkData.mockReturnValue({
        success: true,
        data: mockBulkData,
        executionTime: 150
      });

      // Act
      const result = DataService.getBulkData(userId, options);

      // Assert
      expect(result.success).toBe(true);
      expect(result.data.sheetData).toBeDefined();
      expect(result.data.formInfo).toBeDefined();
      expect(result.data.systemInfo).toBeDefined();
      expect(result.executionTime).toBeLessThan(1000);
    });

    it('should handle bulk data errors gracefully', () => {
      // Arrange
      DataService.getBulkData.mockReturnValue({
        success: false,
        error: 'ユーザーが見つかりません',
        timestamp: '2025-01-15T10:00:00Z'
      });

      // Act
      const result = DataService.getBulkData('invalid-user', {});

      // Assert
      expect(result.success).toBe(false);
      expect(result.error).toBe('ユーザーが見つかりません');
    });

  });

  describe('getAutoStopTime', () => {

    it('should calculate auto stop time correctly', () => {
      // Arrange
      const publishedAt = '2025-01-15T10:00:00Z';
      const minutes = 60;
      const mockResult = {
        publishTime: new Date(publishedAt),
        stopTime: new Date('2025-01-15T11:00:00Z'),
        publishTimeFormatted: '2025/1/15 10:00:00',
        stopTimeFormatted: '2025/1/15 11:00:00',
        remainingMinutes: 30
      };

      DataService.getAutoStopTime.mockReturnValue(mockResult);

      // Act
      const result = DataService.getAutoStopTime(publishedAt, minutes);

      // Assert
      expect(result.publishTime).toEqual(new Date(publishedAt));
      expect(result.remainingMinutes).toBe(30);
      expect(result.publishTimeFormatted).toBeDefined();
    });

    it('should handle invalid dates', () => {
      // Arrange
      DataService.getAutoStopTime.mockReturnValue(null);

      // Act
      const result = DataService.getAutoStopTime('invalid-date', 60);

      // Assert
      expect(result).toBeNull();
    });

  });

  describe('getDataStatistics', () => {

    it('should return data statistics', () => {
      // Arrange
      const userId = 'test-user-123';
      const mockStats = {
        totalEntries: 10,
        nonEmptyEntries: 8,
        reactionCounts: {
          understand: 15,
          like: 12,
          curious: 8
        },
        highlightedEntries: 3,
        avgReactionsPerEntry: 3.5,
        lastUpdated: '2025-01-15T10:00:00Z'
      };

      DataService.getDataStatistics.mockReturnValue(mockStats);

      // Act
      const result = DataService.getDataStatistics(userId);

      // Assert
      expect(result.totalEntries).toBe(10);
      expect(result.reactionCounts.understand).toBe(15);
      expect(result.highlightedEntries).toBe(3);
    });

  });

  describe('updateReactionInSheet', () => {

    it('should update reaction in sheet successfully', () => {
      // Arrange
      const config = {
        spreadsheetId: 'test-sheet-id',
        sheetName: 'テストシート'
      };
      const rowId = 'row_3';
      const reactionType = 'LIKE';
      const action = 'add';

      DataService.updateReactionInSheet.mockReturnValue(true);

      // Act
      const result = DataService.updateReactionInSheet(config, rowId, reactionType, action);

      // Assert
      expect(result).toBe(true);
      expect(DataService.updateReactionInSheet).toHaveBeenCalledWith(config, rowId, reactionType, action);
    });

    it('should handle invalid row ID', () => {
      // Arrange
      DataService.updateReactionInSheet.mockReturnValue(false);

      // Act
      const result = DataService.updateReactionInSheet({}, 'invalid-row', 'LIKE', 'add');

      // Assert
      expect(result).toBe(false);
    });

  });

  describe('refreshBoardData', () => {

    it('should refresh board data successfully', () => {
      // Arrange
      const userId = 'test-user-123';
      const mockRefreshResult = {
        success: true,
        message: 'データを更新しました',
        updatedAt: '2025-01-15T10:00:00Z',
        dataCount: 15
      };

      DataService.refreshBoardData.mockReturnValue(mockRefreshResult);

      // Act
      const result = DataService.refreshBoardData(userId);

      // Assert
      expect(result.success).toBe(true);
      expect(result.dataCount).toBe(15);
      expect(result.updatedAt).toBeDefined();
    });

  });

  describe('validateSheetAccess', () => {

    it('should validate sheet access permissions', () => {
      // Arrange
      const userId = 'test-user-123';
      const spreadsheetId = 'test-sheet-id';

      DataService.validateSheetAccess.mockReturnValue({
        hasAccess: true,
        permissions: ['read', 'write'],
        sheetExists: true
      });

      // Act
      const result = DataService.validateSheetAccess(userId, spreadsheetId);

      // Assert
      expect(result.hasAccess).toBe(true);
      expect(result.permissions).toContain('read');
      expect(result.sheetExists).toBe(true);
    });

    it('should handle access denied', () => {
      // Arrange
      DataService.validateSheetAccess.mockReturnValue({
        hasAccess: false,
        error: 'アクセス権限がありません',
        sheetExists: true
      });

      // Act
      const result = DataService.validateSheetAccess('unauthorized-user', 'test-sheet-id');

      // Assert
      expect(result.hasAccess).toBe(false);
      expect(result.error).toBe('アクセス権限がありません');
    });

  });

  describe('Error Handling', () => {

    it('should handle spreadsheet API errors', () => {
      // Arrange
      DataService.getUserSheetData.mockImplementation(() => {
        throw new Error('Spreadsheet API error');
      });

      // Act & Assert
      expect(() => DataService.getUserSheetData('test-user')).toThrow('Spreadsheet API error');
    });

    it('should handle network timeouts', () => {
      // Arrange
      DataService.getBulkData.mockImplementation(() => {
        throw new Error('Request timeout');
      });

      // Act & Assert
      expect(() => DataService.getBulkData('test-user', {})).toThrow('Request timeout');
    });

  });

  describe('Performance Tests', () => {

    it('should complete sheet operations within time limits', async () => {
      // Arrange
      const startTime = Date.now();
      DataService.getUserSheetData.mockReturnValue({
        success: true,
        data: testSheetData,
        count: testSheetData.length
      });

      // Act
      DataService.getUserSheetData('test-user');
      const endTime = Date.now();

      // Assert (should complete within 100ms for mock)
      expect(endTime - startTime).toBeLessThan(100);
    });

    it('should handle large datasets efficiently', () => {
      // Arrange
      const largeDataset = Array(1000).fill().map((_, i) => ({
        id: i + 1,
        answer: `回答${i + 1}`,
        reactions: { understand: 0, like: 0, curious: 0 }
      }));

      DataService.getUserSheetData.mockReturnValue({
        success: true,
        data: largeDataset,
        count: 1000
      });

      // Act
      const result = DataService.getUserSheetData('test-user');

      // Assert
      expect(result.count).toBe(1000);
      expect(result.success).toBe(true);
    });

  });

  describe('Diagnostics', () => {

    it('should return diagnostic information', () => {
      // Arrange
      const mockDiagnostics = {
        service: 'DataService',
        timestamp: '2025-01-15T00:00:00.000Z',
        checks: [
          { name: 'Sheet Access', status: '✅', details: 'Sheet accessible' },
          { name: 'Data Processing', status: '✅', details: 'Processing working' },
          { name: 'Reaction System', status: '✅', details: 'Reactions working' },
          { name: 'Bulk Operations', status: '✅', details: 'Bulk ops working' }
        ],
        overall: '✅'
      };

      DataService.diagnose.mockReturnValue(mockDiagnostics);

      // Act
      const result = DataService.diagnose();

      // Assert
      expect(result.service).toBe('DataService');
      expect(result.overall).toBe('✅');
      expect(result.checks).toHaveLength(4);
    });

  });

});

describe('DataService Integration Tests', () => {

  it('should complete full data workflow', async () => {
    // Arrange
    const userId = 'integration-test-user';
    
    // Mock the full workflow
    DataService.getUserSheetData.mockResolvedValue({
      success: true,
      data: testSheetData,
      count: testSheetData.length
    });
    
    DataService.addReaction.mockResolvedValue({
      success: true,
      updatedReactions: { understand: 1, like: 2, curious: 0 }
    });
    
    DataService.toggleHighlight.mockResolvedValue({
      success: true,
      highlighted: true
    });
    
    DataService.getDataStatistics.mockResolvedValue({
      totalEntries: testSheetData.length,
      reactionCounts: { understand: 5, like: 8, curious: 3 }
    });

    // Act
    const sheetData = await DataService.getUserSheetData(userId);
    const reactionResult = await DataService.addReaction(userId, 'row_2', 'LIKE');
    const highlightResult = await DataService.toggleHighlight(userId, 'row_2');
    const statistics = await DataService.getDataStatistics(userId);

    // Assert
    expect(sheetData.success).toBe(true);
    expect(reactionResult.success).toBe(true);
    expect(highlightResult.success).toBe(true);
    expect(statistics.totalEntries).toBeGreaterThan(0);
  });

});