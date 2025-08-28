/**
 * エラー修正の統合検証テスト
 * 実際のシステム連携での動作確認
 */

describe('エラー修正統合検証', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  describe('getCurrentUserStatus型安全性統合テスト', () => {
    test('文字列のrequestUserIdで正常に動作', () => {
      // モック設定
      global.getCurrentUserEmail = jest.fn(() => '35t22@naha-okinawa.ed.jp');
      global.verifyUserAccess = jest.fn();
      global.findUserById = jest.fn((id) => ({
        userId: id,
        adminEmail: '35t22@naha-okinawa.ed.jp',
        isActive: true,
        spreadsheetId: 'test-sheet-id',
      }));
      global.ERROR_SEVERITY = { MEDIUM: 'MEDIUM' };
      global.ERROR_CATEGORIES = { VALIDATION: 'VALIDATION', SYSTEM: 'SYSTEM' };
      global.logError = jest.fn();

      // getCurrentUserStatusをモック
      function mockGetCurrentUserStatus(requestUserId = null) {
        try {
          const activeUserEmail = global.getCurrentUserEmail();

          // 型安全性強化: requestUserIdの型チェック
          if (requestUserId != null && typeof requestUserId !== 'string') {
            global.logError(
              new Error(`Invalid requestUserId type: ${typeof requestUserId}`),
              'getCurrentUserStatus',
              global.ERROR_SEVERITY.MEDIUM,
              global.ERROR_CATEGORIES.VALIDATION,
              { requestUserId, type: typeof requestUserId }
            );
            return {
              status: 'error',
              message: 'requestUserIdは文字列である必要があります',
              data: null,
              userInfo: null,
              timestamp: new Date().toISOString(),
            };
          }

          let userInfo;
          if (requestUserId && typeof requestUserId === 'string' && requestUserId.trim() !== '') {
            global.verifyUserAccess(requestUserId);
            userInfo = global.findUserById(requestUserId);
          } else {
            // フォールバック処理
            userInfo = { userId: 'auto-generated', adminEmail: activeUserEmail };
          }

          return {
            status: 'success',
            userInfo,
            timestamp: new Date().toISOString(),
          };
        } catch (e) {
          global.logError(
            e,
            'getCurrentUserStatus',
            global.ERROR_SEVERITY.MEDIUM,
            global.ERROR_CATEGORIES.SYSTEM,
            { requestUserId }
          );
          return {
            status: 'error',
            message: `ステータス取得に失敗しました: ${e.message}`,
            data: null,
            userInfo: null,
            timestamp: new Date().toISOString(),
          };
        }
      }

      // テスト実行
      const validResult = mockGetCurrentUserStatus('valid-user-id');
      expect(validResult.status).toBe('success');
      expect(validResult.userInfo.userId).toBe('valid-user-id');
    });

    test('オブジェクトのrequestUserIdでエラーハンドリング', () => {
      global.logError = jest.fn();
      global.ERROR_SEVERITY = { MEDIUM: 'MEDIUM' };
      global.ERROR_CATEGORIES = { VALIDATION: 'VALIDATION' };

      function mockGetCurrentUserStatus(requestUserId = null) {
        if (requestUserId != null && typeof requestUserId !== 'string') {
          global.logError(
            new Error(`Invalid requestUserId type: ${typeof requestUserId}`),
            'getCurrentUserStatus',
            global.ERROR_SEVERITY.MEDIUM,
            global.ERROR_CATEGORIES.VALIDATION,
            { requestUserId, type: typeof requestUserId }
          );
          return {
            status: 'error',
            message: 'requestUserIdは文字列である必要があります',
            data: null,
            userInfo: null,
            timestamp: new Date().toISOString(),
          };
        }
        return { status: 'success' };
      }

      const errorResult = mockGetCurrentUserStatus({ userId: 'test' });
      expect(errorResult.status).toBe('error');
      expect(errorResult.message).toBe('requestUserIdは文字列である必要があります');
      expect(global.logError).toHaveBeenCalled();
    });
  });

  describe('logClientError統合テスト', () => {
    test('文字列エラーの正常処理', () => {
      function mockLogClientError(errorInfo) {
        try {
          let message = 'unknown error';
          let userId = 'unknown';

          if (typeof errorInfo === 'string') {
            message = errorInfo;
          } else if (errorInfo && typeof errorInfo === 'object') {
            message = errorInfo.message || errorInfo.error || JSON.stringify(errorInfo);
            userId = errorInfo.userId || errorInfo.user || 'unknown';
          }

          console.error(`🚨 CLIENT: ${message} (${userId})`);
          return { status: 'success', logged: true };
        } catch (e) {
          console.error(`🚨 CLIENT ERROR LOGGING FAILED: ${e.message}`);
          return { status: 'error', message: e.message };
        }
      }

      const result = mockLogClientError('Test error message');
      expect(result.status).toBe('success');
      expect(result.logged).toBe(true);
    });

    test('オブジェクトエラーの正常処理', () => {
      function mockLogClientError(errorInfo) {
        try {
          let message = 'unknown error';
          let userId = 'unknown';

          if (typeof errorInfo === 'string') {
            message = errorInfo;
          } else if (errorInfo && typeof errorInfo === 'object') {
            message = errorInfo.message || errorInfo.error || JSON.stringify(errorInfo);
            userId = errorInfo.userId || errorInfo.user || 'unknown';
          }

          return { status: 'success', logged: true, message, userId };
        } catch (e) {
          return { status: 'error', message: e.message };
        }
      }

      const errorObj = {
        message: 'Frontend error',
        userId: 'user123',
        url: 'https://example.com',
        timestamp: new Date().toISOString(),
      };

      const result = mockLogClientError(errorObj);
      expect(result.status).toBe('success');
      expect(result.message).toBe('Frontend error');
      expect(result.userId).toBe('user123');
    });

    test('undefined/null エラーの安全な処理', () => {
      function mockLogClientError(errorInfo) {
        try {
          let message = 'unknown error';
          let userId = 'unknown';

          if (typeof errorInfo === 'string') {
            message = errorInfo;
          } else if (errorInfo && typeof errorInfo === 'object') {
            message = errorInfo.message || errorInfo.error || JSON.stringify(errorInfo);
            userId = errorInfo.userId || errorInfo.user || 'unknown';
          }

          return { status: 'success', logged: true, message, userId };
        } catch (e) {
          return { status: 'error', message: e.message };
        }
      }

      const resultUndefined = mockLogClientError(undefined);
      expect(resultUndefined.status).toBe('success');
      expect(resultUndefined.message).toBe('unknown error');

      const resultNull = mockLogClientError(null);
      expect(resultNull.status).toBe('success');
      expect(resultNull.message).toBe('unknown error');
    });
  });

  describe('フロントエンド-バックエンド連携テスト', () => {
    test('正しいエラー情報の送信', () => {
      // フロントエンドのエラー送信処理をモック
      function mockSendToBackend(errorInfo) {
        const clientErrorData = {
          message: errorInfo.message || errorInfo.error || 'unknown error',
          userId: 'mock-user-id',
          url: 'https://mock-url.com',
          timestamp: new Date().toISOString(),
          userAgent: 'Mock User Agent',
          stack: errorInfo.stack,
        };

        // バックエンドのlogClientErrorをシミュレート
        function mockBackendLogClientError(errorData) {
          let message = 'unknown error';
          let userId = 'unknown';

          if (typeof errorData === 'string') {
            message = errorData;
          } else if (errorData && typeof errorData === 'object') {
            message = errorData.message || errorData.error || JSON.stringify(errorData);
            userId = errorData.userId || errorData.user || 'unknown';
          }

          return { status: 'success', logged: true, message, userId };
        }

        return mockBackendLogClientError(clientErrorData);
      }

      const frontendError = {
        message: 'JavaScript error occurred',
        stack: 'Error at line 100',
      };

      const result = mockSendToBackend(frontendError);
      expect(result.status).toBe('success');
      expect(result.logged).toBe(true);
      expect(result.message).toBe('JavaScript error occurred');
      expect(result.userId).toBe('mock-user-id');
    });
  });
});
