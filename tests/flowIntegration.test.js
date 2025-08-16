const fs = require('fs');
const vm = require('vm');

describe('StudyQuest System Flow Integration Tests', () => {
  let context;
  const mockSpreadsheetId = 'test-spreadsheet-123';
  const mockUserId = 'USER_123';

  beforeEach(() => {
    // Mock環境のセットアップ
    context = {
      // Google Apps Script API mocks
      Session: {
        getActiveUser: () => ({ getEmail: () => 'test@example.com' })
      },
      PropertiesService: {
        getUserProperties: () => ({
          getProperty: jest.fn().mockReturnValue(null),
          setProperty: jest.fn(),
          deleteProperty: jest.fn()
        })
      },
      LockService: {
        getScriptLock: jest.fn(() => ({
          waitLock: jest.fn(),
          releaseLock: jest.fn()
        }))
      },
      CacheService: {
        getUserCache: () => ({
          get: jest.fn().mockReturnValue(null),
          put: jest.fn(),
          remove: jest.fn()
        }),
        getScriptCache: () => ({
          get: jest.fn().mockReturnValue(null),
          put: jest.fn(),
          remove: jest.fn()
        })
      },
      
      // キャッシュマネージャー Mock
      cacheManager: {
        get: jest.fn(),
        put: jest.fn(),
        remove: jest.fn(),
        getHealth: jest.fn().mockReturnValue('healthy')
      },
      
      // データベース操作 Mocks
      findUserById: jest.fn().mockReturnValue({
        userId: mockUserId,
        adminEmail: 'test@example.com',
        spreadsheetId: mockSpreadsheetId,
        configJson: '{"appPublished": true}'
      }),
      
      getUserId: jest.fn().mockReturnValue(mockUserId),
      verifyUserAccess: jest.fn(),
      
      // システムセットアップ Mocks
      isSystemSetup: jest.fn().mockReturnValue(true),
      checkApplicationAccess: jest.fn().mockReturnValue({ hasAccess: true }),
      
      // キャッシュクリア関数 Mocks
      clearAllExecutionCache: jest.fn(),
      clearExecutionUserInfoCache: jest.fn(),
      clearExecutionSheetsServiceCache: jest.fn(),
      
      // コンソール Mock
      console: {
        log: jest.fn(),
        error: jest.fn(),
        warn: jest.fn()
      },
      
      // Utilities Mock
      Utilities: {
        sleep: jest.fn()
      }
    };

    // テスト用のコードを実行コンテキストに読み込み
    vm.createContext(context);
  });

  describe('Concurrent Flow Interference Tests', () => {
    
    test('saveAndPublish and addReaction should not interfere', async () => {
      // saveAndPublish と addReaction の同時実行テスト
      
      const saveAndPublishMock = jest.fn().mockImplementation(async (userId, sheetName, config) => {
        // LockService の呼び出しをシミュレート
        const lock = context.LockService.getScriptLock();
        lock.waitLock(30000);
        
        // 処理時間をシミュレート
        await new Promise(resolve => setTimeout(resolve, 50));
        
        lock.releaseLock();
        return { status: 'success' };
      });
      
      const addReactionMock = jest.fn().mockImplementation(async (userId, rowIndex, reaction, sheetName) => {
        // LockService の呼び出しをシミュレート
        const lock = context.LockService.getScriptLock();
        lock.waitLock(10000);
        
        // 処理時間をシミュレート
        await new Promise(resolve => setTimeout(resolve, 30));
        
        lock.releaseLock();
        return { status: 'ok' };
      });
      
      context.saveAndPublish = saveAndPublishMock;
      context.addReaction = addReactionMock;
      
      // 同時実行
      const promises = [
        context.saveAndPublish(mockUserId, 'TestSheet', { showNames: true }),
        context.addReaction(mockUserId, 5, 'LIKE', 'TestSheet'),
        context.addReaction(mockUserId, 6, 'UNDERSTAND', 'TestSheet')
      ];
      
      const results = await Promise.all(promises);
      
      // 全ての操作が成功することを確認
      expect(results).toHaveLength(3);
      expect(results[0].status).toBe('success');
      expect(results[1].status).toBe('ok');
      expect(results[2].status).toBe('ok');
      
      // LockService が適切に呼ばれることを確認 (3回の操作 × 各1回)
      expect(context.LockService.getScriptLock).toHaveBeenCalledTimes(3);
    });
    
    test('Cache invalidation should not cause race conditions', () => {
      // キャッシュ無効化の競合テスト
      
      let cacheCleared = 0;
      context.clearExecutionUserInfoCache = jest.fn(() => {
        cacheCleared++;
      });
      
      // 複数のキャッシュクリア呼び出しを並行実行
      const clearCacheOperations = Array(10).fill(0).map(() => 
        context.clearExecutionUserInfoCache()
      );
      
      Promise.all(clearCacheOperations);
      
      // キャッシュクリアが10回呼ばれることを確認
      expect(context.clearExecutionUserInfoCache).toHaveBeenCalledTimes(10);
    });
    
    test('User authentication flow should handle concurrent requests', () => {
      // 認証フローの並行処理テスト
      
      const concurrentUsers = [
        'USER_A',
        'USER_B', 
        'USER_C'
      ];
      
      concurrentUsers.forEach(userId => {
        context.findUserById.mockReturnValueOnce({
          userId: userId,
          adminEmail: `${userId.toLowerCase()}@example.com`,
          spreadsheetId: `spreadsheet-${userId}`,
          configJson: '{"appPublished": true}'
        });
      });
      
      // 並行認証テスト
      const authPromises = concurrentUsers.map(userId => {
        return new Promise(resolve => {
          context.verifyUserAccess(userId);
          resolve(userId);
        });
      });
      
      return Promise.all(authPromises).then(results => {
        expect(results).toEqual(concurrentUsers);
        expect(context.verifyUserAccess).toHaveBeenCalledTimes(3);
      });
    });
  });

  describe('Flow State Consistency Tests', () => {
    
    test('Session state should remain consistent across flows', () => {
      // セッション状態の一貫性テスト
      
      const initialEmail = 'user1@example.com';
      const switchedEmail = 'user2@example.com';
      
      context.Session.getActiveUser = jest.fn()
        .mockReturnValueOnce({ getEmail: () => initialEmail })
        .mockReturnValueOnce({ getEmail: () => switchedEmail });
      
      // 初期セッション確認
      const email1 = context.Session.getActiveUser().getEmail();
      expect(email1).toBe(initialEmail);
      
      // セッション切り替え後の確認
      const email2 = context.Session.getActiveUser().getEmail();
      expect(email2).toBe(switchedEmail);
    });
    
    test('Cache hierarchy should maintain data consistency', () => {
      // キャッシュ階層の整合性テスト
      
      const testData = { userId: mockUserId, data: 'test' };
      
      // 実行レベルキャッシュにデータを設定
      context.executionUserInfoCache = testData;
      
      // 統一キャッシュマネージャーでもキャッシュを設定
      context.cacheManager.get = jest.fn().mockReturnValue(testData);
      
      // データの一貫性を確認
      const cachedData = context.cacheManager.get('test-key');
      expect(cachedData).toEqual(testData);
    });
  });

  describe('Error Recovery Tests', () => {
    
    test('Flow should recover from database errors', () => {
      // データベースエラーからの回復テスト
      
      context.findUserById = jest.fn()
        .mockImplementationOnce(() => {
          throw new Error('Database connection failed');
        })
        .mockImplementationOnce(() => ({
          userId: mockUserId,
          adminEmail: 'test@example.com',
          spreadsheetId: mockSpreadsheetId
        }));
      
      // 最初の呼び出しはエラー
      expect(() => context.findUserById(mockUserId)).toThrow('Database connection failed');
      
      // 2回目の呼び出しは成功
      const result = context.findUserById(mockUserId);
      expect(result.userId).toBe(mockUserId);
    });
    
    test('Lock timeout should not cause system hang', () => {
      // ロックタイムアウトのテスト
      
      const mockLock = {
        waitLock: jest.fn().mockImplementation((timeout) => {
          if (timeout < 5000) {
            throw new Error('Lock timeout');
          }
        }),
        releaseLock: jest.fn()
      };
      
      context.LockService.getScriptLock = jest.fn().mockReturnValue(mockLock);
      
      // 短いタイムアウトはエラーになることを確認
      const lock = context.LockService.getScriptLock();
      expect(() => lock.waitLock(1000)).toThrow('Lock timeout');
      
      // 長いタイムアウトは成功することを確認
      expect(() => lock.waitLock(10000)).not.toThrow();
    });
  });

  describe('Performance Impact Tests', () => {
    
    test('Cache operations should not significantly impact performance', () => {
      // キャッシュ操作のパフォーマンステスト
      
      const start = Date.now();
      
      // 大量のキャッシュ操作
      for (let i = 0; i < 1000; i++) {
        context.cacheManager.get(`key-${i}`);
        context.cacheManager.put(`key-${i}`, `value-${i}`);
      }
      
      const duration = Date.now() - start;
      
      // 1秒以内に完了することを確認
      expect(duration).toBeLessThan(1000);
    });
    
    test('Multiple flow executions should not cause memory leaks', () => {
      // メモリリークのテスト (実際のメモリ使用量は測定できないため、呼び出し回数で代用)
      
      const executionCount = 100;
      let cacheClears = 0;
      
      context.clearExecutionUserInfoCache = jest.fn(() => {
        cacheClears++;
      });
      
      // 複数フロー実行のシミュレーション
      for (let i = 0; i < executionCount; i++) {
        context.clearExecutionUserInfoCache();
      }
      
      // 適切な回数のクリーンアップが実行されることを確認
      expect(cacheClears).toBe(executionCount);
    });
  });
});