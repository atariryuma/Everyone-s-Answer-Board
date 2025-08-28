/**
 * @fileoverview getServiceAccountTokenCached関数のテスト
 */

const fs = require('fs');
const path = require('path');

// serviceAccount.gsからgetServiceAccountTokenCached関数を抽出
function loadGetServiceAccountTokenCachedFunction() {
  const file = fs.readFileSync(path.join(__dirname, '../src/serviceAccount.gs'), 'utf8');

  // getServiceAccountTokenCached関数を抽出
  const start = file.indexOf('function getServiceAccountTokenCached');
  if (start === -1) {
    throw new Error('getServiceAccountTokenCached function not found');
  }

  let i = file.indexOf('{', start);
  let depth = 1;
  i += 1;
  while (i < file.length && depth > 0) {
    const char = file[i];
    if (char === '{') depth += 1;
    if (char === '}') depth -= 1;
    i += 1;
  }

  const fnStr = file.slice(start, i);

  // 必要なグローバル関数とサービスのモック
  const mockCacheService = {
    getScriptCache: jest.fn(() => ({
      get: jest.fn(),
      put: jest.fn(),
    })),
  };

  const mockGenerateNewServiceAccountToken = jest.fn();
  const mockLogError = jest.fn();
  const mockConsole = {
    warn: jest.fn(),
  };

  const fn = new Function(`
    const CacheService = arguments[0];
    const generateNewServiceAccountToken = arguments[1];
    const logError = arguments[2];
    const console = arguments[3];
    const ERROR_SEVERITY = { HIGH: 'high' };
    const ERROR_CATEGORIES = { AUTHENTICATION: 'authentication' };
    
    ${fnStr}; 
    return { getServiceAccountTokenCached, CacheService, generateNewServiceAccountToken };
  `)(mockCacheService, mockGenerateNewServiceAccountToken, mockLogError, mockConsole);

  return {
    ...fn,
    mockCacheService,
    mockGenerateNewServiceAccountToken,
    mockLogError,
    mockConsole,
  };
}

describe('getServiceAccountTokenCached関数テスト', () => {
  let functions;
  let mockCache;

  beforeEach(() => {
    functions = loadGetServiceAccountTokenCachedFunction();
    mockCache = {
      get: jest.fn(),
      put: jest.fn(),
    };
    functions.mockCacheService.getScriptCache.mockReturnValue(mockCache);
    jest.clearAllMocks();
  });

  test('キャッシュにトークンがある場合はキャッシュから返す', () => {
    const cachedToken = 'cached-access-token-12345';
    mockCache.get.mockReturnValue(cachedToken);

    const result = functions.getServiceAccountTokenCached();

    expect(result).toBe(cachedToken);
    expect(mockCache.get).toHaveBeenCalledWith('service_account_token');
    expect(functions.mockGenerateNewServiceAccountToken).not.toHaveBeenCalled();
    expect(mockCache.put).not.toHaveBeenCalled();
  });

  test('キャッシュにトークンがない場合は新規生成してキャッシュに保存', () => {
    const newToken = 'new-access-token-67890';
    mockCache.get.mockReturnValue(null);
    functions.mockGenerateNewServiceAccountToken.mockReturnValue(newToken);

    const result = functions.getServiceAccountTokenCached();

    expect(result).toBe(newToken);
    expect(mockCache.get).toHaveBeenCalledWith('service_account_token');
    expect(functions.mockGenerateNewServiceAccountToken).toHaveBeenCalled();
    expect(mockCache.put).toHaveBeenCalledWith('service_account_token', newToken, 3300);
  });

  test('forceRefreshがtrueの場合はキャッシュを無視して新規生成', () => {
    const cachedToken = 'cached-token';
    const newToken = 'fresh-token';
    mockCache.get.mockReturnValue(cachedToken);
    functions.mockGenerateNewServiceAccountToken.mockReturnValue(newToken);

    const result = functions.getServiceAccountTokenCached(true);

    expect(result).toBe(newToken);
    expect(mockCache.get).not.toHaveBeenCalled();
    expect(functions.mockGenerateNewServiceAccountToken).toHaveBeenCalled();
    expect(mockCache.put).toHaveBeenCalledWith('service_account_token', newToken, 3300);
  });

  test('新規トークン生成に失敗した場合はnullを返す', () => {
    mockCache.get.mockReturnValue(null);
    functions.mockGenerateNewServiceAccountToken.mockReturnValue(null);

    const result = functions.getServiceAccountTokenCached();

    expect(result).toBeNull();
    expect(functions.mockGenerateNewServiceAccountToken).toHaveBeenCalled();
    expect(mockCache.put).not.toHaveBeenCalled();
  });

  test('キャッシュ保存に失敗してもトークンは返す', () => {
    const newToken = 'valid-token';
    mockCache.get.mockReturnValue(null);
    functions.mockGenerateNewServiceAccountToken.mockReturnValue(newToken);
    mockCache.put.mockImplementation(() => {
      throw new Error('Cache storage failed');
    });

    const result = functions.getServiceAccountTokenCached();

    expect(result).toBe(newToken);
    expect(functions.mockConsole.warn).toHaveBeenCalledWith(
      expect.stringContaining('トークンキャッシュの保存に失敗'),
      'Cache storage failed'
    );
  });

  test('予期しないエラーが発生した場合は適切にハンドリング', () => {
    functions.mockCacheService.getScriptCache.mockImplementation(() => {
      throw new Error('CacheService unavailable');
    });

    const result = functions.getServiceAccountTokenCached();

    expect(result).toBeNull();
    expect(functions.mockLogError).toHaveBeenCalledWith(
      expect.any(Error),
      'getServiceAccountTokenCached',
      'high',
      'authentication'
    );
  });
});
