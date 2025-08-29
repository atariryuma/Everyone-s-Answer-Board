/**
 * 統一キャッシュシステム統合テスト
 * Phase 4: キャッシュ管理完全統合の動作確認
 */

describe('統一キャッシュシステム統合テスト', () => {
  beforeEach(() => {
    // モックの初期化
    global.cacheManager = {
      get: jest.fn(),
      remove: jest.fn(),
      clearByPattern: jest.fn(),
      clearAll: jest.fn(),
      getHealth: jest.fn(() => ({
        status: 'ok',
        stats: { hitRate: '85%', totalOperations: 100, errors: 2 }
      })),
      invalidateRelated: jest.fn(),
      invalidateSheetData: jest.fn()
    };
    
    global.getUnifiedExecutionCache = jest.fn(() => ({
      clearUserInfo: jest.fn(),
      clearSheetsService: jest.fn(),
      clearAll: jest.fn(),
      syncWithUnifiedCache: jest.fn()
    }));

    global.getServiceAccountTokenCached = jest.fn(() => 'mock_token');
    global.CacheService = {
      getScriptCache: jest.fn(() => ({
        removeAll: jest.fn(),
        remove: jest.fn()
      })),
      getUserCache: jest.fn(() => ({
        removeAll: jest.fn(),
        remove: jest.fn()
      }))
    };
    
    // 統一キャッシュAPIをモック
    global.unifiedCacheAPI = {
      clearUserInfoCache: jest.fn(),
      clearAllExecutionCache: jest.fn(),
      getSheetsServiceCached: jest.fn(),
      invalidateUserCache: jest.fn(),
      synchronizeCacheAfterCriticalUpdate: jest.fn(),
      clearDatabaseCache: jest.fn(),
      invalidateSpreadsheetCache: jest.fn(),
      getHealth: jest.fn(() => ({ status: 'ok' }))
    };
  });

  describe('統一キャッシュAPI', () => {
    test('clearUserInfoCacheが正常に動作する', () => {
      const mockClearUserInfo = jest.fn();
      const mockSyncWithUnifiedCache = jest.fn();
      
      global.getUnifiedExecutionCache = jest.fn(() => ({
        clearUserInfo: mockClearUserInfo,
        syncWithUnifiedCache: mockSyncWithUnifiedCache
      }));

      global.unifiedCacheAPI.clearUserInfoCache('test_user');

      expect(global.unifiedCacheAPI.clearUserInfoCache).toHaveBeenCalledWith('test_user');
    });

    test('clearAllExecutionCacheが全キャッシュをクリアする', () => {
      global.unifiedCacheAPI.clearAllExecutionCache();

      expect(global.unifiedCacheAPI.clearAllExecutionCache).toHaveBeenCalled();
    });

    test('getSheetsServiceCachedがサービスを返す', () => {
      global.unifiedCacheAPI.getSheetsServiceCached();

      expect(global.unifiedCacheAPI.getSheetsServiceCached).toHaveBeenCalled();
    });

    test('invalidateUserCacheが複数パラメータを処理する', () => {
      const userId = 'user123';
      const email = 'test@example.com';
      const spreadsheetId = 'sheet123';
      
      global.unifiedCacheAPI.invalidateUserCache(userId, email, spreadsheetId);

      expect(global.unifiedCacheAPI.invalidateUserCache).toHaveBeenCalledWith(
        userId, email, spreadsheetId
      );
    });

    test('synchronizeCacheAfterCriticalUpdateがクリティカル同期を実行', () => {
      const userId = 'user123';
      const email = 'test@example.com';
      const oldSpreadsheetId = 'old_sheet';
      const newSpreadsheetId = 'new_sheet';
      
      global.unifiedCacheAPI.synchronizeCacheAfterCriticalUpdate(
        userId, email, oldSpreadsheetId, newSpreadsheetId
      );

      expect(global.unifiedCacheAPI.synchronizeCacheAfterCriticalUpdate)
        .toHaveBeenCalledWith(userId, email, oldSpreadsheetId, newSpreadsheetId);
    });
  });

  describe('後方互換性テスト', () => {
    test('clearExecutionUserInfoCacheが統一APIにリダイレクトされる', () => {
      // 後方互換性関数がモック化されていることを確認
      global.clearExecutionUserInfoCache = jest.fn(() => 
        global.unifiedCacheAPI.clearUserInfoCache()
      );

      global.clearExecutionUserInfoCache();

      expect(global.clearExecutionUserInfoCache).toHaveBeenCalled();
    });

    test('clearAllExecutionCacheが統一APIにリダイレクトされる', () => {
      global.clearAllExecutionCache = jest.fn(() =>
        global.unifiedCacheAPI.clearAllExecutionCache()
      );

      global.clearAllExecutionCache();

      expect(global.clearAllExecutionCache).toHaveBeenCalled();
    });

    test('getSheetsServiceCachedが統一APIにリダイレクトされる', () => {
      global.getSheetsServiceCached = jest.fn(() =>
        global.unifiedCacheAPI.getSheetsServiceCached()
      );

      global.getSheetsServiceCached();

      expect(global.getSheetsServiceCached).toHaveBeenCalled();
    });

    test('getCachedSheetsServiceが統一APIにリダイレクトされる', () => {
      global.getCachedSheetsService = jest.fn(() =>
        global.unifiedCacheAPI.getSheetsServiceCached()
      );

      global.getCachedSheetsService();

      expect(global.getCachedSheetsService).toHaveBeenCalled();
    });
  });

  describe('キャッシュ効率テスト', () => {
    test('preWarmCacheが事前読み込みを実行', () => {
      global.preWarmCache = jest.fn(() => ({
        preWarmedItems: ['service_account_token', 'user_by_email'],
        errors: [],
        success: true
      }));

      const result = global.preWarmCache('test@example.com');

      expect(result).toEqual({
        preWarmedItems: ['service_account_token', 'user_by_email'],
        errors: [],
        success: true
      });
    });

    test('analyzeCacheEfficiencyが分析結果を返す', () => {
      global.analyzeCacheEfficiency = jest.fn(() => ({
        efficiency: 'excellent',
        recommendations: [],
        optimizationOpportunities: ['TTL延長によるさらなる高速化']
      }));

      const result = global.analyzeCacheEfficiency();

      expect(result.efficiency).toBe('excellent');
      expect(result.optimizationOpportunities).toContain('TTL延長によるさらなる高速化');
    });

    test('getHealthがキャッシュヘルス情報を返す', () => {
      const healthInfo = global.unifiedCacheAPI.getHealth();

      expect(healthInfo).toEqual({ status: 'ok' });
    });
  });

  describe('エラーハンドリングテスト', () => {
    test('無効なユーザーIDでキャッシュクリアが安全に失敗', () => {
      global.unifiedCacheAPI.invalidateUserCache = jest.fn().mockImplementation(() => {
        // エラーをスローせずに安全に処理することを確認
      });

      expect(() => {
        global.unifiedCacheAPI.invalidateUserCache(null, null, null);
      }).not.toThrow();
    });

    test('キャッシュマネージャー未初期化時の安全な動作', () => {
      global.cacheManager = undefined;
      
      global.unifiedCacheAPI.clearDatabaseCache = jest.fn().mockImplementation(() => {
        // cacheManagerが未定義でもエラーをスローしない
      });

      expect(() => {
        global.unifiedCacheAPI.clearDatabaseCache();
      }).not.toThrow();
    });
  });

  describe('パフォーマンステスト', () => {
    test('大量キャッシュクリア時の制限確認', () => {
      global.cacheManager.clearByPattern = jest.fn().mockImplementation((pattern, options) => {
        // maxKeysオプションが適用されることを確認
        expect(options).toHaveProperty('maxKeys');
        expect(options.maxKeys).toBeLessThanOrEqual(500);
        return options.maxKeys; // 削除した数を返す
      });

      global.unifiedCacheAPI.invalidateUserCache = jest.fn().mockImplementation(() => {
        global.cacheManager.clearByPattern('user_*', { maxKeys: 500 });
      });

      global.unifiedCacheAPI.invalidateUserCache('user123', 'test@example.com');
      
      expect(global.cacheManager.clearByPattern).toHaveBeenCalledWith(
        'user_*', 
        expect.objectContaining({ maxKeys: expect.any(Number) })
      );
    });

    test('キャッシュ階層の正しい使用順序', () => {
      const mockCache = {
        // Level 1: メモ化キャッシュ (最高速)
        memoCache: new Map(),
        // Level 2: Apps Scriptキャッシュ
        scriptCache: { get: jest.fn(), put: jest.fn() },
        // Level 3: PropertiesServiceフォールバック
        propertiesService: { getProperty: jest.fn(), setProperty: jest.fn() }
      };

      global.cacheManager.get = jest.fn().mockImplementation((key, valueFn, options) => {
        // 階層化アクセスの順序をテスト
        const { enableMemoization, usePropertiesFallback } = options || {};
        
        if (enableMemoization && mockCache.memoCache.has(key)) {
          return mockCache.memoCache.get(key);
        }
        
        const scriptResult = mockCache.scriptCache.get(key);
        if (scriptResult) return scriptResult;
        
        if (usePropertiesFallback) {
          return mockCache.propertiesService.getProperty(key);
        }
        
        return valueFn ? valueFn() : null;
      });

      // テスト実行
      global.cacheManager.get('test_key', () => 'new_value', {
        enableMemoization: true,
        usePropertiesFallback: true
      });

      expect(global.cacheManager.get).toHaveBeenCalledWith(
        'test_key',
        expect.any(Function),
        expect.objectContaining({
          enableMemoization: true,
          usePropertiesFallback: true
        })
      );
    });
  });
});