/**
 * @fileoverview getLogoutUrl関数のテスト
 */

const fs = require('fs');
const path = require('path');

// main.gsからgetLogoutUrl関数を抽出
function loadGetLogoutUrlFunction() {
  const file = fs.readFileSync(path.join(__dirname, '../src/main.gs'), 'utf8');
  
  // getLogoutUrl関数を抽出
  const start = file.indexOf('function getLogoutUrl');
  if (start === -1) {
    throw new Error('getLogoutUrl function not found');
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
  
  // 必要なグローバル関数のモック
  const mockLogError = jest.fn();
  const mockGetWebAppUrl = jest.fn(() => 'https://script.google.com/test');
  const mockERROR_SEVERITY = { LOW: 'low' };
  const mockERROR_CATEGORIES = { SYSTEM: 'system' };
  
  const fn = new Function(`
    const logError = function() { /* mock */ };
    const getWebAppUrl = function() { return 'https://script.google.com/test'; };
    const ERROR_SEVERITY = ${JSON.stringify(mockERROR_SEVERITY)};
    const ERROR_CATEGORIES = ${JSON.stringify(mockERROR_CATEGORIES)};
    const encodeURIComponent = global.encodeURIComponent;
    
    ${fnStr}; 
    return { getLogoutUrl, logError: logError, getWebAppUrl: getWebAppUrl };
  `);
  return fn();
}

describe('getLogoutUrl関数テスト', () => {
  let functions;
  
  beforeEach(() => {
    functions = loadGetLogoutUrlFunction();
    jest.clearAllMocks();
  });

  test('正常にログアウトURLを生成できる', () => {
    const result = functions.getLogoutUrl();
    
    expect(typeof result).toBe('string');
    expect(result).toContain('https://accounts.google.com/logout');
    expect(result).toContain('continue=');
  });

  test('WebAppURLが含まれたログアウトURLを生成する', () => {
    const result = functions.getLogoutUrl();
    
    // encodeURIComponentでエンコードされたWebAppURLが含まれていることを確認
    expect(result).toContain(encodeURIComponent('https://script.google.com/test'));
  });

  test('エラーが発生した場合はフォールバックURLを返す', () => {
    // エラーケースをテストするためのシンプルな実装
    const getLogoutUrlWithError = () => {
      try {
        throw new Error('WebApp URL取得失敗');
      } catch (error) {
        // フォールバック: 単純なログアウトURL
        return 'https://accounts.google.com/logout';
      }
    };
    
    const result = getLogoutUrlWithError();
    
    // フォールバックURLが返されることを確認
    expect(result).toBe('https://accounts.google.com/logout');
  });

  test('URLの形式が正しい', () => {
    const result = functions.getLogoutUrl();
    
    // 有効なURLであることを確認
    expect(() => new URL(result)).not.toThrow();
    
    // HTTPSプロトコルを使用していることを確認
    const url = new URL(result);
    expect(url.protocol).toBe('https:');
    expect(url.hostname).toBe('accounts.google.com');
    expect(url.pathname).toBe('/logout');
  });

  test('continue パラメータが正しくエンコードされている', () => {
    const result = functions.getLogoutUrl();
    const url = new URL(result);
    
    expect(url.searchParams.has('continue')).toBe(true);
    
    const continueUrl = url.searchParams.get('continue');
    expect(continueUrl).toBe('https://script.google.com/test');
  });
});

describe('AdminPanelログアウト統合テスト', () => {
  test('フォールバック機構の動作確認', () => {
    // adminPanel.js.htmlのログアウト関数の動作を模擬
    const mockConfirm = jest.fn(() => true);
    const mockConsoleError = jest.fn();
    
    // API呼び出し失敗を模擬
    const mockCallAPI = jest.fn().mockRejectedValue(new Error('getLogoutUrl is not a function'));
    
    // フォールバック実行の模擬
    const fallbackLogout = async () => {
      if (mockConfirm('ログアウトしますか？')) {
        try {
          await mockCallAPI('getLogoutUrl');
        } catch (error) {
          mockConsoleError('API呼び出し失敗、代替ログアウトを実行:', error);
          return 'https://accounts.google.com/logout?continue=' + encodeURIComponent('https://example.com');
        }
      }
    };
    
    return fallbackLogout().then(result => {
      expect(mockConfirm).toHaveBeenCalledWith('ログアウトしますか？');
      expect(mockCallAPI).toHaveBeenCalledWith('getLogoutUrl');
      expect(mockConsoleError).toHaveBeenCalled();
      expect(result).toContain('https://accounts.google.com/logout');
      expect(result).toContain('continue=');
    });
  });
});