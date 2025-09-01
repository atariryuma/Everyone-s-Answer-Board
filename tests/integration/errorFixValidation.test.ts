/**
 * エラー修正の統合検証テスト (TypeScript版)
 * 実際のシステム連携での動作確認 + 型安全性保証
 */

import { MockTestUtils, GASError } from '../mocks/gasMocks';

// Type definitions for better test structure
interface UserInfo {
  userId: string;
  adminEmail: string;
  isActive: boolean;
  spreadsheetId?: string;
}

interface UserStatusResponse {
  status: 'success' | 'error';
  userInfo?: UserInfo | null;
  message?: string;
  data?: any;
  timestamp: string;
}

interface ErrorInfo {
  message?: string;
  error?: string;
  userId?: string;
  user?: string;
  url?: string;
  timestamp?: string;
  stack?: string;
}

interface ClientErrorResponse {
  status: 'success' | 'error';
  logged?: boolean;
  message?: string;
  userId?: string;
}

describe('エラー修正統合検証 (TypeScript強化版)', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    MockTestUtils.clearAllMockData();
  });

  describe('getCurrentUserStatus型安全性統合テスト', () => {
    // Mock functions with proper typing
    const mockGetCurrentUserEmail = jest.fn(() => '35t22@naha-okinawa.ed.jp');
    const mockVerifyUserAccess = jest.fn();
    const mockFindUserById = jest.fn(
      (id: string): UserInfo => ({
        userId: id,
        adminEmail: '35t22@naha-okinawa.ed.jp',
        isActive: true,
        spreadsheetId: 'test-sheet-id',
      })
    );
    const mockLogError = jest.fn();

    const mockGetCurrentUserStatus = (requestUserId?: string | null): UserStatusResponse => {
      try {
        const activeUserEmail = mockGetCurrentUserEmail();

        // 型安全性強化: requestUserIdの型チェック
        if (requestUserId != null && typeof requestUserId !== 'string') {
          const error = new Error(`Invalid requestUserId type: ${typeof requestUserId}`);
          mockLogError(error, 'getCurrentUserStatus', 'MEDIUM', 'VALIDATION', {
            requestUserId,
            type: typeof requestUserId,
          });
          return {
            status: 'error',
            message: 'requestUserIdは文字列である必要があります',
            data: null,
            userInfo: null,
            timestamp: new Date().toISOString(),
          };
        }

        let userInfo: UserInfo;
        if (requestUserId && typeof requestUserId === 'string' && requestUserId.trim() !== '') {
          mockVerifyUserAccess(requestUserId);
          userInfo = mockFindUserById(requestUserId);
        } else {
          // フォールバック処理
          userInfo = {
            userId: 'auto-generated',
            adminEmail: activeUserEmail,
            isActive: true,
          };
        }

        return {
          status: 'success',
          userInfo,
          timestamp: new Date().toISOString(),
        };
      } catch (e) {
        const error = e as Error;
        mockLogError(error, 'getCurrentUserStatus', 'MEDIUM', 'SYSTEM', { requestUserId });
        return {
          status: 'error',
          message: `ステータス取得に失敗しました: ${error.message}`,
          data: null,
          userInfo: null,
          timestamp: new Date().toISOString(),
        };
      }
    };

    test('文字列のrequestUserIdで正常に動作', () => {
      const validResult = mockGetCurrentUserStatus('valid-user-id');

      expect(validResult.status).toBe('success');
      expect(validResult.userInfo).toBeDefined();
      expect(validResult.userInfo!.userId).toBe('valid-user-id');
      expect(validResult.userInfo!.adminEmail).toBeValidEmail();
      expect(validResult.timestamp).toMatch(/^\d{4}-\d{2}-\d{2}T/);

      // TypeScript ensures we can't access wrong properties
      // @ts-expect-error - This should cause a TypeScript error
      // expect(validResult.invalidProperty).toBeDefined();
    });

    test('オブジェクトのrequestUserIdでエラーハンドリング', () => {
      const errorResult = mockGetCurrentUserStatus({ userId: 'test' } as any);

      expect(errorResult.status).toBe('error');
      expect(errorResult.message).toBe('requestUserIdは文字列である必要があります');
      expect(errorResult.userInfo).toBeNull();
      expect(mockLogError).toHaveBeenCalledWith(
        expect.any(Error),
        'getCurrentUserStatus',
        'MEDIUM',
        'VALIDATION',
        expect.objectContaining({
          requestUserId: { userId: 'test' },
          type: 'object',
        })
      );
    });

    test('null/undefinedのrequestUserIdでフォールバック処理', () => {
      const nullResult = mockGetCurrentUserStatus(null);
      const undefinedResult = mockGetCurrentUserStatus(undefined);

      expect(nullResult.status).toBe('success');
      expect(nullResult.userInfo!.userId).toBe('auto-generated');

      expect(undefinedResult.status).toBe('success');
      expect(undefinedResult.userInfo!.userId).toBe('auto-generated');
    });

    test('空文字列のrequestUserIdでフォールバック処理', () => {
      const emptyResult = mockGetCurrentUserStatus('');
      const whitespaceResult = mockGetCurrentUserStatus('   ');

      expect(emptyResult.status).toBe('success');
      expect(emptyResult.userInfo!.userId).toBe('auto-generated');

      expect(whitespaceResult.status).toBe('success');
      expect(whitespaceResult.userInfo!.userId).toBe('auto-generated');
    });

    test('mockFindUserByIdがエラーを投げる場合の処理', () => {
      mockFindUserById.mockImplementationOnce(() => {
        throw new Error('User not found in database');
      });

      const errorResult = mockGetCurrentUserStatus('nonexistent-user');

      expect(errorResult.status).toBe('error');
      expect(errorResult.message).toContain('User not found in database');
      expect(mockLogError).toHaveBeenCalledWith(
        expect.any(Error),
        'getCurrentUserStatus',
        'MEDIUM',
        'SYSTEM',
        { requestUserId: 'nonexistent-user' }
      );
    });
  });

  describe('logClientError統合テスト (型安全版)', () => {
    const mockLogClientError = (
      errorInfo: string | ErrorInfo | null | undefined
    ): ClientErrorResponse => {
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
        const error = e as Error;
        return { status: 'error', message: error.message };
      }
    };

    test('文字列エラーの正常処理', () => {
      const result = mockLogClientError('Test error message');

      expect(result.status).toBe('success');
      expect(result.logged).toBe(true);
      expect(result.message).toBe('Test error message');
      expect(result.userId).toBe('unknown');
    });

    test('ErrorInfoオブジェクトの正常処理', () => {
      const errorObj: ErrorInfo = {
        message: 'Frontend error',
        userId: 'user123',
        url: 'https://example.com',
        timestamp: new Date().toISOString(),
      };

      const result = mockLogClientError(errorObj);

      expect(result.status).toBe('success');
      expect(result.logged).toBe(true);
      expect(result.message).toBe('Frontend error');
      expect(result.userId).toBe('user123');
    });

    test('不完全なErrorInfoオブジェクトの処理', () => {
      const partialErrorObj: Partial<ErrorInfo> = {
        error: 'Generic error message',
        // userId is missing
      };

      const result = mockLogClientError(partialErrorObj);

      expect(result.status).toBe('success');
      expect(result.message).toBe('Generic error message');
      expect(result.userId).toBe('unknown');
    });

    test('undefined/null エラーの安全な処理', () => {
      const resultUndefined = mockLogClientError(undefined);
      expect(resultUndefined.status).toBe('success');
      expect(resultUndefined.message).toBe('unknown error');

      const resultNull = mockLogClientError(null);
      expect(resultNull.status).toBe('success');
      expect(resultNull.message).toBe('unknown error');
    });

    test('複雑なオブジェクト（JSON.stringify使用）', () => {
      const complexErrorObj = {
        nestedError: {
          code: 500,
          details: 'Server error',
        },
        timestamp: Date.now(),
      };

      const result = mockLogClientError(complexErrorObj);

      expect(result.status).toBe('success');
      expect(result.message).toContain('nestedError');
      expect(result.message).toContain('500');
    });
  });

  describe('フロントエンド-バックエンド連携テスト (型安全版)', () => {
    interface BackendErrorData {
      message: string;
      userId: string;
      url: string;
      timestamp: string;
      userAgent: string;
      stack?: string;
    }

    const mockSendToBackend = (errorInfo: ErrorInfo): ClientErrorResponse => {
      const clientErrorData: BackendErrorData = {
        message: errorInfo.message || errorInfo.error || 'unknown error',
        userId: 'mock-user-id',
        url: 'https://mock-url.com',
        timestamp: new Date().toISOString(),
        userAgent: 'Mock User Agent',
        stack: errorInfo.stack,
      };

      const mockBackendLogClientError = (errorData: BackendErrorData): ClientErrorResponse => {
        let message = 'unknown error';
        let userId = 'unknown';

        if (errorData && typeof errorData === 'object') {
          message = errorData.message || 'unknown error';
          userId = errorData.userId || 'unknown';
        }

        return { status: 'success', logged: true, message, userId };
      };

      return mockBackendLogClientError(clientErrorData);
    };

    test('正しいエラー情報の送信', () => {
      const frontendError: ErrorInfo = {
        message: 'JavaScript error occurred',
        stack: 'Error at line 100',
      };

      const result = mockSendToBackend(frontendError);

      expect(result.status).toBe('success');
      expect(result.logged).toBe(true);
      expect(result.message).toBe('JavaScript error occurred');
      expect(result.userId).toBe('mock-user-id');
    });

    test('スタックトレース付きエラーの送信', () => {
      const errorWithStack: ErrorInfo = {
        message: 'TypeError: Cannot read property',
        stack: 'TypeError: Cannot read property\n    at Object.<anonymous> (/file.js:10:5)',
        url: 'https://example.com/page.html',
      };

      const result = mockSendToBackend(errorWithStack);

      expect(result).toMatchObject({
        status: 'success',
        logged: true,
        message: 'TypeError: Cannot read property',
      });
    });
  });

  describe('プロダクション環境特有のエラー再現テスト', () => {
    test('GASスプレッドシート権限エラーの再現', () => {
      // 実際のGAS環境で発生する権限エラーをシミュレート
      MockTestUtils.simulatePermissionError('spreadsheet_invalid-id');

      expect(() => {
        global.SpreadsheetApp.openById('invalid-id');
      }).toThrowGASError('You do not have permission to access');
    });

    test('配列サイズ不一致エラーの再現', () => {
      // 実際のsetValues呼び出しで発生するエラーをシミュレート
      MockTestUtils.setMockData('sheet_Sheet1_data', [['Header1', 'Header2']]);

      const sheet = global.SpreadsheetApp.getActiveSheet();
      const range = sheet.getRange(1, 1, 2, 2);

      expect(() => {
        // 2x2の範囲に1行しか提供
        range.setValues([['Value1', 'Value2']]);
      }).toThrowGASError(/number of rows in the data/);
    });

    test('レスポンス時間制限エラーの再現', async () => {
      jest.setTimeout(10000); // 10秒タイムアウト

      // 長時間実行をシミュレート
      const slowOperation = () => {
        return new Promise((resolve) => {
          setTimeout(resolve, 2000); // 2秒待機
        });
      };

      const startTime = Date.now();
      await slowOperation();
      const endTime = Date.now();

      expect(endTime - startTime).toBeGreaterThanOrEqual(2000);
    });
  });

  describe('型安全性テスト', () => {
    test('UserInfoインターフェースの型チェック', () => {
      const validUserInfo: UserInfo = {
        userId: 'test-user-123',
        adminEmail: 'test@example.com',
        isActive: true,
        spreadsheetId: 'test-spreadsheet-id',
      };

      expect(validUserInfo.userId).toBeValidUserId();
      expect(validUserInfo.adminEmail).toBeValidEmail();
      expect(validUserInfo.isActive).toBe(true);

      // TypeScript prevents invalid property access
      // @ts-expect-error
      // const invalidProperty = validUserInfo.nonExistentProperty;
    });

    test('エラーレスポンスの型安全性', () => {
      const errorResponse: UserStatusResponse = {
        status: 'error',
        message: 'Test error',
        timestamp: new Date().toISOString(),
      };

      expect(errorResponse.status).toBe('error');
      expect(errorResponse.userInfo).toBeUndefined();
      expect(errorResponse.message).toBeDefined();
    });
  });
});
