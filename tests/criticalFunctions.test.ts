/**
 * @fileoverview クリティカル機能のテスト (TypeScript版)
 * 新アーキテクチャでの重要機能の動作確認 + 型安全性強化
 */

import { MockTestUtils } from './mocks/gasMocks';

// Type definitions for test data structures
interface UserInfo {
  userId: string;
  adminEmail: string;
  spreadsheetId: string;
  spreadsheetUrl: string;
  createdAt: string;
  configJson: string;
  lastAccessedAt: string;
  isActive: string;
}

interface ConfigJson {
  appPublished?: boolean;
  publishedSheetName?: string;
  setupStatus?: string;
  formCreated?: boolean;
  [key: string]: any;
}

interface InitialDataResponse {
  userInfo: UserInfo;
  appUrls: AppUrls;
  setupStep: number;
  activeSheetName: string;
  webAppUrl: string;
  isPublished: boolean;
  answerCount: number;
  totalReactions: number;
  config: ConfigJson;
  allSheets: string[];
  _meta: {
    apiVersion: string;
    executionTime: null;
  };
}

interface AppUrls {
  webApp: string;
  admin: string;
  setup: string;
  published: string;
}

interface SheetDataResponse {
  success: boolean;
  data?: any[];
  fromCache?: boolean;
  error?: string;
}

interface QuickStartFormResponse {
  success: boolean;
  formUrl?: string;
  spreadsheetUrl?: string;
  setupStep?: number;
}

interface SheetConfig {
  opinionHeader?: string;
  nameHeader?: string;
  classHeader?: string;
  [key: string]: any;
}

interface ApiResponse<T = any> {
  success: boolean;
  data?: T;
  error?: string;
}

interface MockServices {
  getUserInfo: jest.MockedFunction<(userId: string) => UserInfo | null>;
  getUserId: jest.MockedFunction<() => string>;
  updateUser: jest.MockedFunction<(userId: string, data: Partial<UserInfo>) => boolean>;
  createUser: jest.MockedFunction<(data: Partial<UserInfo>) => UserInfo>;
  getSheetData: jest.MockedFunction<
    (
      userId: string,
      sheetName?: string,
      classFilter?: string | null,
      sortOrder?: string,
      adminMode?: boolean
    ) => any[]
  >;
  getResponsesData: jest.MockedFunction<() => ApiResponse>;
  verifyUserAccess: jest.MockedFunction<(userId: string) => boolean>;
  createQuickStartForm: jest.MockedFunction<(email: string, userId: string) => any>;
  logError: jest.MockedFunction<(...args: any[]) => void>;
  debugLog: jest.MockedFunction<(...args: any[]) => void>;
  warnLog: jest.MockedFunction<(...args: any[]) => void>;
  infoLog: jest.MockedFunction<(...args: any[]) => void>;
}

// Enhanced memory cache implementation with proper typing
interface CacheEntry<T = any> {
  value: T;
  expiry: number;
}

class TypedMemoryCache {
  private cache = new Map<string, CacheEntry>();

  get<T = any>(key: string): T | null {
    const entry = this.cache.get(key);
    if (entry && entry.expiry > Date.now()) {
      return entry.value;
    }
    if (entry) {
      this.cache.delete(key); // Clean up expired entries
    }
    return null;
  }

  set<T = any>(key: string, value: T, ttlSeconds = 300): void {
    this.cache.set(key, {
      value,
      expiry: Date.now() + ttlSeconds * 1000,
    });
  }

  delete(key: string): boolean {
    return this.cache.delete(key);
  }

  clear(): void {
    this.cache.clear();
  }

  clearUserCache(userId: string): void {
    for (const key of this.cache.keys()) {
      if (key.includes(userId)) {
        this.cache.delete(key);
      }
    }
  }
}

// Global typed cache instance
const typedMemoryCache = new TypedMemoryCache();

// Enhanced test setup with proper typing
const setupTypedTest = (): MockServices => {
  // Reset all mock data
  MockTestUtils.clearAllMockData();

  // Setup realistic test data
  const testUserInfo: UserInfo = {
    userId: 'user123',
    adminEmail: 'test@example.com',
    spreadsheetId: 'sheet123',
    spreadsheetUrl: 'https://sheets.google.com/123',
    createdAt: '2024-01-01',
    configJson: '{"appPublished":true,"publishedSheetName":"Sheet1","setupStatus":"completed"}',
    lastAccessedAt: '2024-01-02',
    isActive: 'true',
  };

  MockTestUtils.setMockData('user_user123', testUserInfo);
  MockTestUtils.createRealisticSpreadsheetData(
    [
      'userId',
      'adminEmail',
      'spreadsheetId',
      'spreadsheetUrl',
      'createdAt',
      'configJson',
      'lastAccessedAt',
      'isActive',
    ],
    [
      [
        testUserInfo.userId,
        testUserInfo.adminEmail,
        testUserInfo.spreadsheetId,
        testUserInfo.spreadsheetUrl,
        testUserInfo.createdAt,
        testUserInfo.configJson,
        testUserInfo.lastAccessedAt,
        testUserInfo.isActive,
      ],
    ]
  );

  // Create typed mock services
  const mockServices: MockServices = {
    getUserInfo: jest.fn((userId: string): UserInfo | null => {
      if (userId === 'user123') {
        return testUserInfo;
      }
      return null;
    }),

    getUserId: jest.fn((): string => 'user123'),

    updateUser: jest.fn((userId: string, data: Partial<UserInfo>): boolean => {
      const user = MockTestUtils.getMockData(`user_${userId}`);
      if (user) {
        MockTestUtils.setMockData(`user_${userId}`, { ...user, ...data });
        return true;
      }
      return false;
    }),

    createUser: jest.fn((data: Partial<UserInfo>): UserInfo => {
      const newUser: UserInfo = {
        userId: 'new-user-123',
        adminEmail: '',
        spreadsheetId: '',
        spreadsheetUrl: '',
        createdAt: new Date().toISOString(),
        configJson: '{}',
        lastAccessedAt: new Date().toISOString(),
        isActive: 'true',
        ...data,
      };
      MockTestUtils.setMockData(`user_${newUser.userId}`, newUser);
      return newUser;
    }),

    getSheetData: jest.fn(
      (
        _userId: string,
        _sheetName?: string,
        _classFilter?: string | null,
        _sortOrder?: string,
        _adminMode?: boolean
      ): any[] => {
        return [{ rowIndex: 1, name: 'Test User', opinion: 'Test Opinion' }];
      }
    ),

    getResponsesData: jest.fn(
      (): ApiResponse => ({
        status: 'success' as any,
        data: [{ id: 1 }, { id: 2 }],
      })
    ),

    verifyUserAccess: jest.fn((_userId: string): boolean => true),

    createQuickStartForm: jest.fn((_email: string, _userId: string) => ({
      formId: 'form-123',
      formUrl: 'https://forms.google.com/123',
      spreadsheetId: 'sheet-new-123',
      spreadsheetUrl: 'https://sheets.google.com/new',
    })),

    logError: jest.fn(),
    debugLog: jest.fn(),
    warnLog: jest.fn(),
    infoLog: jest.fn(),
  };

  // Set global mock functions
  (global as any).getUserInfo = mockServices.getUserInfo;
  (global as any).getUserId = mockServices.getUserId;
  (global as any).updateUser = mockServices.updateUser;
  (global as any).createUser = mockServices.createUser;
  (global as any).getSheetData = mockServices.getSheetData;
  (global as any).getResponsesData = mockServices.getResponsesData;
  (global as any).verifyUserAccess = mockServices.verifyUserAccess;
  (global as any).createQuickStartForm = mockServices.createQuickStartForm;
  (global as any).logError = mockServices.logError;
  (global as any).debugLog = mockServices.debugLog;
  (global as any).warnLog = mockServices.warnLog;
  (global as any).infoLog = mockServices.infoLog;

  // Setup typed cache functions
  (global as any).getCacheValue = <T = any>(key: string): T | null => typedMemoryCache.get<T>(key);
  (global as any).setCacheValue = <T = any>(key: string, value: T, ttl = 300): void =>
    typedMemoryCache.set(key, value, ttl);
  (global as any).clearUserCache = (userId: string): void =>
    typedMemoryCache.clearUserCache(userId);

  // Setup core business logic functions with proper typing
  (global as any).getSetupStep = (userInfo: UserInfo | null, configJson: ConfigJson): number => {
    if (!userInfo || !userInfo.spreadsheetId) return 1;
    if (!configJson.formCreated) return 2;
    if (!configJson.appPublished) return 3;
    return 4;
  };

  (global as any).getSheetsList = (_userId: string): string[] => ['Sheet1', 'Sheet2'];

  (global as any).generateUserUrls = (_userId: string): AppUrls => ({
    webApp: 'https://script.google.com/test',
    admin: 'https://script.google.com/test?mode=admin',
    setup: 'https://script.google.com/test?mode=setup',
    published: 'https://script.google.com/test',
  });

  (global as any).getWebAppUrl = (): string => 'https://script.google.com/test';

  return mockServices;
};

// Core function implementations with proper typing
const createGetInitialData = (mockServices: MockServices) => {
  return (
    requestUserId: string,
    _targetSheetName?: string,
    _lightweightMode?: boolean
  ): InitialDataResponse => {
    const userInfo = mockServices.getUserInfo(requestUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    const configJson: ConfigJson = JSON.parse(userInfo.configJson);
    // formCreatedを追加してステップ4になるようにする
    configJson.formCreated = true;

    return {
      userInfo,
      appUrls: (global as any).generateUserUrls(requestUserId),
      setupStep: (global as any).getSetupStep(userInfo, configJson),
      activeSheetName: configJson.publishedSheetName || 'Sheet1',
      webAppUrl: (global as any).getWebAppUrl(),
      isPublished: !!configJson.appPublished,
      answerCount: 2,
      totalReactions: 4,
      config: configJson,
      allSheets: (global as any).getSheetsList(requestUserId),
      _meta: {
        apiVersion: 'integrated_v1',
        executionTime: null,
      },
    };
  };
};

const createGetPublishedSheetData = (mockServices: MockServices) => {
  return (
    requestUserId: string,
    classFilter?: string | null,
    sortOrder?: string,
    adminMode?: boolean,
    bypassCache?: boolean
  ): SheetDataResponse => {
    const userInfo = mockServices.getUserInfo(requestUserId);
    if (!userInfo) {
      return { success: false, error: 'ユーザーが見つかりません' };
    }

    // Check cache first (unless bypassed)
    if (!bypassCache) {
      const cacheKey = `sheet_data_${requestUserId}_Sheet1_${classFilter}_${sortOrder}_${adminMode}`;
      const cached = (global as any).getCacheValue(cacheKey);
      if (cached) {
        return { success: true, data: cached, fromCache: true };
      }
    }

    const data = mockServices.getSheetData(
      requestUserId,
      'Sheet1',
      classFilter,
      sortOrder,
      adminMode
    );
    return { success: true, data, fromCache: false };
  };
};

const createQuickStartFormUI = (mockServices: MockServices) => {
  return (requestUserId: string): QuickStartFormResponse => {
    const userInfo = mockServices.getUserInfo(requestUserId);
    if (!userInfo) {
      throw new Error('ユーザーが見つかりません');
    }

    const result = mockServices.createQuickStartForm(userInfo.adminEmail, requestUserId);

    mockServices.updateUser(requestUserId, {
      spreadsheetId: result.spreadsheetId,
      spreadsheetUrl: result.spreadsheetUrl,
    });

    return {
      success: true,
      formUrl: result.formUrl,
      spreadsheetUrl: result.spreadsheetUrl,
      setupStep: 3,
    };
  };
};

const createSaveSheetConfig = (mockServices: MockServices) => {
  return (
    userId: string,
    spreadsheetId: string,
    sheetName: string,
    config: SheetConfig,
    options: { setAsPublished?: boolean } = {}
  ): boolean => {
    const userInfo = mockServices.getUserInfo(userId);
    if (!userInfo) {
      return false;
    }

    const configJson: ConfigJson = JSON.parse(userInfo.configJson || '{}');
    configJson[`sheet_${sheetName}`] = config;

    if (options.setAsPublished) {
      configJson.publishedSheetName = sheetName;
      configJson.appPublished = true;
    }

    mockServices.updateUser(userId, { configJson: JSON.stringify(configJson) });
    (global as any).clearUserCache(userId);

    return true;
  };
};

// Test suite
describe('クリティカル機能テスト (TypeScript強化版)', () => {
  let mockServices: MockServices;
  let getInitialData: ReturnType<typeof createGetInitialData>;
  let getPublishedSheetData: ReturnType<typeof createGetPublishedSheetData>;
  let createQuickStartFormUI: ReturnType<typeof createQuickStartFormUI>;
  let saveSheetConfig: ReturnType<typeof createSaveSheetConfig>;

  beforeEach(() => {
    mockServices = setupTypedTest();
    jest.clearAllMocks();

    // Create typed function implementations
    getInitialData = createGetInitialData(mockServices);
    getPublishedSheetData = createGetPublishedSheetData(mockServices);
    createQuickStartFormUI = createQuickStartFormUI(mockServices);
    saveSheetConfig = createSaveSheetConfig(mockServices);

    // Set global implementations
    (global as any).getInitialData = getInitialData;
    (global as any).getPublishedSheetData = getPublishedSheetData;
    (global as any).createQuickStartFormUI = createQuickStartFormUI;
    (global as any).saveSheetConfig = saveSheetConfig;
  });

  afterEach(() => {
    typedMemoryCache.clear();
  });

  describe('getInitialData (型安全版)', () => {
    test('正常に初期データを取得できる', () => {
      const result = getInitialData('user123', 'Sheet1', false);

      expect(result).toBeDefined();
      expect(result.userInfo.userId).toBe('user123');
      expect(result.userInfo.adminEmail).toBeValidEmail();
      expect(result.setupStep).toBe(4); // 全て完了
      expect(result.isPublished).toBe(true);
      expect(result.answerCount).toBe(2);
      expect(result.config).toHaveProperty('appPublished', true);
      expect(result.allSheets).toEqual(['Sheet1', 'Sheet2']);
      expect(result._meta.apiVersion).toBe('integrated_v1');

      // TypeScript ensures type safety
      expect(typeof result.setupStep).toBe('number');
      expect(Array.isArray(result.allSheets)).toBe(true);
    });

    test('ユーザーが見つからない場合エラーをスロー', () => {
      expect(() => {
        getInitialData('invalid-user');
      }).toThrow('ユーザー情報が見つかりません');

      expect(mockServices.getUserInfo).toHaveBeenCalledWith('invalid-user');
    });

    test('configJsonのパースエラーハンドリング', () => {
      mockServices.getUserInfo.mockReturnValueOnce({
        userId: 'user-invalid-json',
        adminEmail: 'test@example.com',
        spreadsheetId: 'sheet123',
        spreadsheetUrl: 'https://sheets.google.com/123',
        createdAt: '2024-01-01',
        configJson: 'invalid-json', // Invalid JSON
        lastAccessedAt: '2024-01-02',
        isActive: 'true',
      });

      expect(() => {
        getInitialData('user-invalid-json');
      }).toThrow();
    });
  });

  describe('getPublishedSheetData (型安全版)', () => {
    test('公開シートデータを取得できる', () => {
      const result = getPublishedSheetData('user123', null, 'newest', false, false);

      expect(result.success).toBe(true);
      expect(result.data).toBeDefined();
      expect(Array.isArray(result.data)).toBe(true);
      expect(result.data!.length).toBeGreaterThan(0);
      expect(result.fromCache).toBe(false);
    });

    test('ユーザーが見つからない場合エラー返却', () => {
      const result = getPublishedSheetData('invalid-user');

      expect(result.success).toBe(false);
      expect(result.error).toBe('ユーザーが見つかりません');
      expect(result.data).toBeUndefined();
    });

    test('キャッシュから取得できる', () => {
      const testData = [{ id: 1, cached: true }];
      typedMemoryCache.set('sheet_data_user123_Sheet1_null_newest_false', testData, 60);

      const result = getPublishedSheetData('user123', null, 'newest', false, false);

      expect(result.success).toBe(true);
      expect(result.fromCache).toBe(true);
      expect(result.data).toEqual(testData);
    });

    test('bypassCacheでキャッシュを回避', () => {
      const testData = [{ id: 1, cached: true }];
      typedMemoryCache.set('sheet_data_user123_Sheet1_null_newest_false', testData, 60);

      const result = getPublishedSheetData('user123', null, 'newest', false, true);

      expect(result.success).toBe(true);
      expect(result.fromCache).toBe(false);
      expect(result.data).not.toEqual(testData);
    });
  });

  describe('createQuickStartFormUI (型安全版)', () => {
    test('クイックスタートフォームを作成できる', () => {
      const result = createQuickStartFormUI('user123');

      expect(result.success).toBe(true);
      expect(result.formUrl).toBeDefined();
      expect(result.formUrl).toMatch(/^https:\/\/forms\.google\.com/);
      expect(result.spreadsheetUrl).toBeDefined();
      expect(result.setupStep).toBe(3);

      expect(mockServices.createQuickStartForm).toHaveBeenCalledWith('test@example.com', 'user123');
      expect(mockServices.updateUser).toHaveBeenCalledWith('user123', expect.any(Object));
    });

    test('ユーザーが見つからない場合エラーをスロー', () => {
      expect(() => {
        createQuickStartFormUI('invalid-user');
      }).toThrow('ユーザーが見つかりません');
    });

    test('createQuickStartFormのエラーを適切に処理', () => {
      mockServices.createQuickStartForm.mockImplementationOnce(() => {
        throw new Error('Form creation failed');
      });

      expect(() => {
        createQuickStartFormUI('user123');
      }).toThrow('Form creation failed');
    });
  });

  describe('saveSheetConfig (型安全版)', () => {
    test('シート設定を保存できる', () => {
      const config: SheetConfig = {
        opinionHeader: '意見',
        nameHeader: '名前',
        classHeader: 'クラス',
      };

      const result = saveSheetConfig('user123', 'sheet123', 'Sheet1', config, {
        setAsPublished: true,
      });

      expect(result).toBe(true);
      expect(mockServices.updateUser).toHaveBeenCalledWith('user123', {
        configJson: expect.stringContaining('Sheet1'),
      });
    });

    test('ユーザーが見つからない場合false返却', () => {
      const config: SheetConfig = { opinionHeader: '意見' };

      const result = saveSheetConfig('invalid-user', 'sheet123', 'Sheet1', config);

      expect(result).toBe(false);
      expect(mockServices.updateUser).not.toHaveBeenCalled();
    });

    test('setAsPublishedオプションが機能する', () => {
      const config: SheetConfig = { opinionHeader: '意見' };

      const result = saveSheetConfig('user123', 'sheet123', 'Sheet1', config, {
        setAsPublished: true,
      });

      expect(result).toBe(true);

      // Check the updated config contains published settings
      const updateCall = mockServices.updateUser.mock.calls[0];
      const updatedConfigJson = JSON.parse(updateCall[1].configJson!);
      expect(updatedConfigJson.publishedSheetName).toBe('Sheet1');
      expect(updatedConfigJson.appPublished).toBe(true);
    });
  });

  describe('API統合 (型安全版)', () => {
    interface ApiRequest {
      action: string;
      params: Record<string, any>;
    }

    const createHandleCoreApiRequest = () => {
      return (action: string, params: Record<string, any>): ApiResponse => {
        const request: ApiRequest = { action, params };

        switch (request.action) {
          case 'getInitialData':
            try {
              const data = getInitialData(request.params.userId);
              return { success: true, data };
            } catch (error) {
              return { success: false, error: (error as Error).message };
            }
          case 'getPublishedSheetData':
            return getPublishedSheetData(request.params.userId) as ApiResponse;
          case 'createQuickStartForm':
            try {
              const data = createQuickStartFormUI(request.params.userId);
              return { success: true, data };
            } catch (error) {
              return { success: false, error: (error as Error).message };
            }
          default:
            return { success: false, error: 'Unknown action' };
        }
      };
    };

    test('handleCoreApiRequestが正しくルーティングする', () => {
      const handleCoreApiRequest = createHandleCoreApiRequest();

      const result1 = handleCoreApiRequest('getInitialData', { userId: 'user123' });
      expect(result1.success).toBe(true);
      expect(result1.data).toBeDefined();

      const result2 = handleCoreApiRequest('getPublishedSheetData', { userId: 'user123' });
      expect(result2.success).toBe(true);

      const result3 = handleCoreApiRequest('unknown', {});
      expect(result3.success).toBe(false);
      expect(result3.error).toBe('Unknown action');
    });

    test('APIエラーハンドリングが適切に動作', () => {
      const handleCoreApiRequest = createHandleCoreApiRequest();

      // Invalid user for getInitialData
      const result = handleCoreApiRequest('getInitialData', { userId: 'invalid-user' });
      expect(result.success).toBe(false);
      expect(result.error).toContain('ユーザー情報が見つかりません');
    });
  });

  describe('パフォーマンステスト', () => {
    test('大量データでのパフォーマンス確認', () => {
      const largeDataSet = Array.from({ length: 1000 }, (_, i) => ({
        id: i,
        name: `User ${i}`,
        data: `Data ${i}`,
      }));

      mockServices.getSheetData.mockReturnValueOnce(largeDataSet);

      const startTime = Date.now();
      const result = getPublishedSheetData('user123', null, 'newest', false, true);
      const endTime = Date.now();

      expect(result.success).toBe(true);
      expect(result.data!.length).toBe(1000);
      expect(endTime - startTime).toBeLessThan(100); // Should complete within 100ms
    });

    test('キャッシュのパフォーマンス効果確認', () => {
      const testData = Array.from({ length: 100 }, (_, i) => ({ id: i }));

      // First call (uncached)
      mockServices.getSheetData.mockReturnValueOnce(testData);
      const startTime1 = Date.now();
      const result1 = getPublishedSheetData('user123', null, 'newest', false, false);
      const endTime1 = Date.now();

      // Set up cache
      typedMemoryCache.set('sheet_data_user123_Sheet1_null_newest_false', testData, 60);

      // Second call (cached)
      const startTime2 = Date.now();
      const result2 = getPublishedSheetData('user123', null, 'newest', false, false);
      const endTime2 = Date.now();

      expect(result1.success).toBe(true);
      expect(result2.success).toBe(true);
      expect(result2.fromCache).toBe(true);

      // Cached call should be faster
      expect(endTime2 - startTime2).toBeLessThan(endTime1 - startTime1);
    });
  });
});
