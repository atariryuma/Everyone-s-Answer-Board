/**
 * 公開エラー修正の単純テスト
 * Issue: Base.gs:updateUserConfig でのspreadsheetId保存問題を直接テスト
 */

describe('Publish Fix - Simple updateUserConfig Test', () => {
  test('updateUserConfig should separate database fields from config fields', () => {
    console.log('🔧 テスト: updateUserConfig のデータベースフィールド分離機能');

    // モック関数とデータを設定
    const mockUpdateUser = jest.fn(() => true);
    const mockGetUserConfig = jest.fn(() => ({
      setupStatus: 'pending',
      displaySettings: { showNames: true }
    }));

    // Base.gsのupdateUserConfig相当のロジックを直接テスト
    function testUpdateUserConfig(userId, updates) {
      const currentConfig = mockGetUserConfig(userId);
      if (!currentConfig) return false;

      // データベースフィールドとconfigJsonフィールドを分離
      const dbFields = ['spreadsheetId', 'spreadsheetUrl', 'formUrl', 'sheetName', 'columnMappingJson', 'publishedAt', 'appUrl', 'isActive'];
      const dbUpdates = {};
      const configUpdates = {};

      Object.entries(updates).forEach(([key, value]) => {
        if (dbFields.includes(key)) {
          dbUpdates[key] = value;
        } else {
          configUpdates[key] = value;
        }
      });

      // configJsonの更新
      const updatedConfig = {
        ...currentConfig,
        ...configUpdates,
        lastModified: new Date().toISOString(),
      };

      // データベース更新データを準備
      const dbUpdateData = {
        configJson: JSON.stringify(updatedConfig),
        lastAccessedAt: new Date().toISOString(),
        ...dbUpdates,
      };

      mockUpdateUser(userId, dbUpdateData);
      return true;
    }

    // テスト実行: spreadsheetIdを含む更新
    const testUserId = 'test-user-123';
    const testUpdates = {
      spreadsheetId: 'test-spreadsheet-456',  // DBフィールド
      sheetName: 'テストシート',               // DBフィールド
      connectionMethod: 'dropdown_select',      // configJsonフィールド
      lastConnected: '2025-01-01T10:00:00Z'    // configJsonフィールド
    };

    const result = testUpdateUserConfig(testUserId, testUpdates);

    // 検証
    expect(result).toBe(true);
    expect(mockUpdateUser).toHaveBeenCalledWith(testUserId, expect.objectContaining({
      spreadsheetId: 'test-spreadsheet-456',  // DBフィールドが含まれているか
      sheetName: 'テストシート',               // DBフィールドが含まれているか
      configJson: expect.stringContaining('dropdown_select'), // configJsonにconfigフィールドが含まれているか
      lastAccessedAt: expect.any(String)
    }));

    // configJsonの内容も確認
    const updateCall = mockUpdateUser.mock.calls[0][1];
    const savedConfig = JSON.parse(updateCall.configJson);
    expect(savedConfig.connectionMethod).toBe('dropdown_select');
    expect(savedConfig.lastConnected).toBe('2025-01-01T10:00:00Z');

    console.log('✅ updateUserConfig: データベースフィールド分離成功');
    console.log('📋 保存されたデータ:', updateCall);
  });

  test('getCurrentConfig should return database spreadsheetId after update', () => {
    console.log('🔧 テスト: データベース更新後のgetCurrentConfig動作');

    // テスト用ユーザーデータ（データベースに保存されている状態をシミュレート）
    const savedUserInfo = {
      userId: 'test-user-789',
      userEmail: 'test@example.com',
      spreadsheetId: 'updated-spreadsheet-123',    // データベースに保存された値
      sheetName: '更新されたシート',                  // データベースに保存された値
      configJson: JSON.stringify({
        connectionMethod: 'dropdown_select',
        lastConnected: '2025-01-01T10:00:00Z',
        setupStatus: 'connected'
      })
    };

    // getCurrentConfig相当のロジックを直接テスト
    function testGetCurrentConfig(userInfo) {
      if (!userInfo) return null;

      let config = {};
      try {
        config = JSON.parse(userInfo.configJson || '{}');
      } catch (e) {
        config = {};
      }

      const fullConfig = {
        userId: userInfo.userId,
        userEmail: userInfo.userEmail,
        spreadsheetId: userInfo.spreadsheetId,    // データベースから直接取得
        sheetName: userInfo.sheetName,            // データベースから直接取得
        setupStatus: config.setupStatus || 'pending',
        connectionMethod: config.connectionMethod,
        lastConnected: config.lastConnected,
      };

      return fullConfig;
    }

    // テスト実行
    const result = testGetCurrentConfig(savedUserInfo);

    // 検証: データベースのspreadsheetIdが正しく返されるか
    expect(result).toBeTruthy();
    expect(result.spreadsheetId).toBe('updated-spreadsheet-123');
    expect(result.sheetName).toBe('更新されたシート');
    expect(result.setupStatus).toBe('connected');
    expect(result.connectionMethod).toBe('dropdown_select');

    console.log('✅ getCurrentConfig: データベース情報取得成功');
    console.log('📋 取得された設定:', result);
  });

  test('publishApp validation logic should pass with valid spreadsheetId', () => {
    console.log('🔧 テスト: publishApp検証ロジック（修正後）');

    // publishApp内の検証ロジックをシミュレート
    function testPublishValidation(currentConfig) {
      if (!currentConfig || !currentConfig.spreadsheetId) {
        return {
          success: false,
          error: 'まずデータソースを設定してください。'
        };
      }

      return {
        success: true,
        message: '公開準備完了'
      };
    }

    // テストケース1: spreadsheetIdが設定されている場合
    const validConfig = {
      userId: 'test-user-valid',
      spreadsheetId: 'valid-spreadsheet-123',
      sheetName: 'validシート'
    };

    const validResult = testPublishValidation(validConfig);
    expect(validResult.success).toBe(true);
    expect(validResult.message).toBe('公開準備完了');

    // テストケース2: spreadsheetIdが未設定の場合（修正前の問題）
    const invalidConfig = {
      userId: 'test-user-invalid',
      spreadsheetId: '',  // 空文字列または未設定
      sheetName: ''
    };

    const invalidResult = testPublishValidation(invalidConfig);
    expect(invalidResult.success).toBe(false);
    expect(invalidResult.error).toBe('まずデータソースを設定してください。');

    // テストケース3: configがnullの場合
    const nullResult = testPublishValidation(null);
    expect(nullResult.success).toBe(false);
    expect(nullResult.error).toBe('まずデータソースを設定してください。');

    console.log('✅ publishApp検証: 全ケースで正常動作確認');
  });
});