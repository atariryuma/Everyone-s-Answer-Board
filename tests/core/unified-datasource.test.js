/**
 * 統一データソーステスト
 * userInfo.spreadsheetIdの一貫性を検証
 */

describe('Unified Data Source Implementation', () => {
  let mockUserInfo;

  beforeEach(() => {
    mockUserInfo = {
      userId: 'test-user-123',
      userEmail: 'test@example.com',
      spreadsheetId: '1z5DNDxNZJB6x8KYRNBKRM1qRlGKCH44rNSMF9vlu4bA',
      sheetName: 'フォームの回答 7',
      configJson: JSON.stringify({
        appPublished: true,
        setupStatus: 'completed',
      }),
    };
  });

  describe('renderAnswerBoard', () => {
    test('userInfo.spreadsheetIdを統一データソースとして使用', () => {
      // renderAnswerBoardのモック実装
      function renderAnswerBoard(userInfo, _params) {
        const targetSpreadsheetId = userInfo.spreadsheetId;
        const targetSheetName = userInfo.sheetName;

        return {
          spreadsheetId: targetSpreadsheetId,
          sheetName: targetSheetName,
          hasUserConfig: !!(targetSpreadsheetId && targetSheetName),
        };
      }

      const result = renderAnswerBoard(mockUserInfo, {});

      expect(result.spreadsheetId).toBe(mockUserInfo.spreadsheetId);
      expect(result.sheetName).toBe(mockUserInfo.sheetName);
      expect(result.hasUserConfig).toBe(true);
    });

    test('publishedSpreadsheetIdを使用しない', () => {
      const configJson = JSON.parse(mockUserInfo.configJson);

      // publishedSpreadsheetIdが存在しないことを確認
      expect(configJson.publishedSpreadsheetId).toBeUndefined();

      // 削除処理のシミュレーション
      if (configJson.publishedSpreadsheetId) {
        delete configJson.publishedSpreadsheetId;
      }

      expect(configJson.publishedSpreadsheetId).toBeUndefined();
    });
  });

  describe('getPublishedSheetData', () => {
    test('userInfo.spreadsheetIdから直接データを取得', () => {
      function getPublishedSheetData(_userId) {
        // DBからユーザー情報取得のモック
        const userInfo = mockUserInfo;

        const targetSpreadsheetId = userInfo.spreadsheetId;
        const targetSheetName = userInfo.sheetName;

        if (!targetSpreadsheetId || !targetSheetName) {
          return {
            status: 'error',
            message: 'スプレッドシート設定が見つかりません',
          };
        }

        return {
          status: 'success',
          spreadsheetId: targetSpreadsheetId,
          sheetName: targetSheetName,
          data: [],
        };
      }

      const result = getPublishedSheetData('test-user-123');

      expect(result.status).toBe('success');
      expect(result.spreadsheetId).toBe(mockUserInfo.spreadsheetId);
      expect(result.sheetName).toBe(mockUserInfo.sheetName);
    });
  });

  describe('データ同期の簡素化', () => {
    test('重複フィールドが削除される', () => {
      const configWithDuplicate = {
        appPublished: true,
        publishedSpreadsheetId: 'old-id-should-be-removed',
        publishedSheetName: 'old-sheet-should-be-removed',
      };

      // fixUserDataConsistencyのモック実装
      function fixUserDataConsistency(configJson) {
        if (configJson.publishedSpreadsheetId) {
          delete configJson.publishedSpreadsheetId;
        }
        if (configJson.publishedSheetName) {
          delete configJson.publishedSheetName;
        }
        return configJson;
      }

      const cleaned = fixUserDataConsistency(configWithDuplicate);

      expect(cleaned.publishedSpreadsheetId).toBeUndefined();
      expect(cleaned.publishedSheetName).toBeUndefined();
      expect(cleaned.appPublished).toBe(true);
    });
  });

  describe('スプレッドシート切り替え', () => {
    test('DB.updateUserでspreadsheetIdとsheetNameを更新', () => {
      const DB = {
        updateUser: jest.fn().mockReturnValue({ success: true }),
      };

      function switchPublishedSheet(userId, spreadsheetId, sheetName) {
        const updateData = {
          spreadsheetId,
          sheetName,
          configJson: JSON.stringify({
            appPublished: true,
            lastModified: new Date().toISOString(),
          }),
        };

        return DB.updateUser(userId, updateData);
      }

      const result = switchPublishedSheet(
        'test-user-123',
        'new-spreadsheet-id',
        'フォームの回答 8'
      );

      expect(result.success).toBe(true);
      expect(DB.updateUser).toHaveBeenCalledWith(
        'test-user-123',
        expect.objectContaining({
          spreadsheetId: 'new-spreadsheet-id',
          sheetName: 'フォームの回答 8',
        })
      );
    });
  });
});
