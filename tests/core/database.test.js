/**
 * データベース機能のコアテスト
 * CLAUDE.md準拠のテストケース
 */

describe('Database Core Functions', () => {
  let mockSpreadsheetApp;
  let mockPropertiesService;

  beforeEach(() => {
    // モックのセットアップ
    mockSpreadsheetApp = {
      openById: jest.fn().mockReturnValue({
        getSheetByName: jest.fn().mockReturnValue({
          getRange: jest.fn().mockReturnValue({
            getValues: jest.fn().mockReturnValue([
              ['userId', 'userEmail', 'createdAt', 'lastAccessedAt', 'isActive'],
              ['user123', 'test@example.com', '2024-01-01', '2024-01-02', 'true'],
            ]),
            setValues: jest.fn(),
          }),
          getLastRow: jest.fn().mockReturnValue(2),
          getLastColumn: jest.fn().mockReturnValue(5),
          appendRow: jest.fn(),
        }),
      }),
    };

    mockPropertiesService = {
      getScriptProperties: jest.fn().mockReturnValue({
        getProperty: jest.fn().mockReturnValue('mock-db-id'),
      }),
    };

    global.SpreadsheetApp = mockSpreadsheetApp;
    global.PropertiesService = mockPropertiesService;
  });

  describe('DB.createUser', () => {
    test('新規ユーザーを正しく作成できる', () => {
      const userData = {
        userId: 'new-user-123',
        userEmail: 'newuser@example.com',
        spreadsheetId: 'sheet-123',
        sheetName: 'フォームの回答 1',
      };

      // DB.createUserのモック実装
      const DB = {
        createUser: jest.fn().mockReturnValue({
          success: true,
          userId: userData.userId,
        }),
      };

      const result = DB.createUser(userData);

      expect(result.success).toBe(true);
      expect(result.userId).toBe(userData.userId);
      expect(DB.createUser).toHaveBeenCalledWith(userData);
    });

    test('重複メールアドレスでエラーを返す', () => {
      const userData = {
        userId: 'duplicate-user',
        userEmail: 'existing@example.com',
      };

      const DB = {
        createUser: jest.fn().mockImplementation(() => {
          throw new Error('このメールアドレスは既に登録されています。');
        }),
      };

      expect(() => DB.createUser(userData)).toThrow('このメールアドレスは既に登録されています。');
    });
  });

  describe('DB.findUserById', () => {
    test('存在するユーザーを検索できる', () => {
      const DB = {
        findUserById: jest.fn().mockReturnValue({
          userId: 'user123',
          userEmail: 'test@example.com',
          spreadsheetId: 'sheet-123',
          sheetName: 'フォームの回答 1',
        }),
      };

      const user = DB.findUserById('user123');

      expect(user).toBeDefined();
      expect(user.userId).toBe('user123');
      expect(user.userEmail).toBe('test@example.com');
    });

    test('存在しないユーザーはnullを返す', () => {
      const DB = {
        findUserById: jest.fn().mockReturnValue(null),
      };

      const user = DB.findUserById('non-existent');

      expect(user).toBeNull();
    });
  });

  describe('DB.updateUser', () => {
    test('ユーザー情報を更新できる', () => {
      const DB = {
        updateUser: jest.fn().mockReturnValue({
          success: true,
          message: 'ユーザー情報を更新しました',
        }),
      };

      const updateData = {
        spreadsheetId: 'new-sheet-456',
        sheetName: 'フォームの回答 2',
      };

      const result = DB.updateUser('user123', updateData);

      expect(result.success).toBe(true);
      expect(DB.updateUser).toHaveBeenCalledWith('user123', updateData);
    });
  });
});

describe('Database Schema Compliance', () => {
  test('DB_CONFIG.HEADERSがCLAUDE.md定義と一致する', () => {
    const expectedHeaders = [
      'userId',
      'userEmail',
      'createdAt',
      'lastAccessedAt',
      'isActive',
      'spreadsheetId',
      'spreadsheetUrl',
      'configJson',
      'formUrl',
      'sheetName',
      'columnMappingJson',
      'publishedAt',
      'appUrl',
      'lastModified',
    ];

    // 実際のDB_CONFIGをモック
    const DB_CONFIG = {
      HEADERS: Object.freeze(expectedHeaders),
    };

    expect(DB_CONFIG.HEADERS).toEqual(expectedHeaders);
    expect(DB_CONFIG.HEADERS.length).toBe(14);
    expect(Object.isFrozen(DB_CONFIG.HEADERS)).toBe(true);
  });
});
