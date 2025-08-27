/**
 * @fileoverview getLoginStatus関数のテスト
 */

const fs = require('fs');
const path = require('path');

// main.gsからgetLoginStatus関数を抽出
function loadGetLoginStatusFunction() {
  const file = fs.readFileSync(path.join(__dirname, '../src/main.gs'), 'utf8');
  
  // getLoginStatus関数を抽出
  const start = file.indexOf('function getLoginStatus');
  if (start === -1) {
    throw new Error('getLoginStatus function not found');
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
  const mockGetCurrentUser = jest.fn();
  const mockLogError = jest.fn();
  const mockGetWebAppUrl = jest.fn(() => 'https://script.google.com/test');
  const mockFetchUserFromDatabase = jest.fn();
  const mockPropertiesService = {
    getUserProperties: jest.fn(() => ({
      getProperty: jest.fn()
    }))
  };
  
  const fn = new Function(`
    const getCurrentUser = arguments[0];
    const logError = arguments[1];
    const getWebAppUrl = arguments[2];
    const fetchUserFromDatabase = arguments[3];
    const PropertiesService = arguments[4];
    const MAIN_ERROR_SEVERITY = { MEDIUM: 'medium' };
    const MAIN_ERROR_CATEGORIES = { AUTH: 'auth', DATA: 'data' };
    const JSON = global.JSON;
    
    ${fnStr}; 
    return { getLoginStatus, getCurrentUser, fetchUserFromDatabase, PropertiesService };
  `)(mockGetCurrentUser, mockLogError, mockGetWebAppUrl, mockFetchUserFromDatabase, mockPropertiesService);
  
  return { ...fn, mockGetCurrentUser, mockLogError, mockGetWebAppUrl, mockFetchUserFromDatabase, mockPropertiesService };
}

describe('getLoginStatus関数テスト', () => {
  let functions;
  
  beforeEach(() => {
    functions = loadGetLoginStatusFunction();
    jest.clearAllMocks();
  });

  test('ログインしていない場合はerrorを返す', () => {
    functions.mockGetCurrentUser.mockReturnValue({
      isLoggedIn: false,
      email: null
    });

    const result = functions.getLoginStatus();
    
    expect(result.status).toBe('error');
    expect(result.isLoggedIn).toBe(false);
    expect(result.message).toContain('ログインしていません');
  });

  test('未登録ユーザーの場合はunregisteredを返す', () => {
    functions.mockGetCurrentUser.mockReturnValue({
      isLoggedIn: true,
      email: 'test@example.com',
      id: 'test-id'
    });
    
    // データベースにユーザーが見つからない
    functions.mockFetchUserFromDatabase.mockReturnValue(null);

    const result = functions.getLoginStatus();
    
    expect(result.status).toBe('unregistered');
    expect(result.isLoggedIn).toBe(true);
    expect(result.email).toBe('test@example.com');
    expect(result.message).toContain('新規ユーザー登録が必要');
    expect(functions.mockFetchUserFromDatabase).toHaveBeenCalledWith('adminEmail', 'test@example.com');
  });

  test('セットアップ未完了の場合はsetup_requiredを返す', () => {
    functions.mockGetCurrentUser.mockReturnValue({
      isLoggedIn: true,
      email: 'test@example.com',
      id: 'test-id'
    });
    
    // データベースにユーザーが存在
    functions.mockFetchUserFromDatabase.mockReturnValue({
      userId: 'test-id',
      adminEmail: 'test@example.com'
    });
    
    const mockUserProperties = {
      getProperty: jest.fn((key) => {
        if (key === 'app_config') {
          return null; // セットアップ未完了
        }
        return null;
      })
    };
    functions.mockPropertiesService.getUserProperties.mockReturnValue(mockUserProperties);

    const result = functions.getLoginStatus();
    
    expect(result.status).toBe('setup_required');
    expect(result.isLoggedIn).toBe(true);
    expect(result.adminUrl).toContain('page=admin');
    expect(result.message).toContain('セットアップを完了');
    expect(functions.mockFetchUserFromDatabase).toHaveBeenCalledWith('adminEmail', 'test@example.com');
  });

  test('既存ユーザーの場合はexisting_userを返す', () => {
    functions.mockGetCurrentUser.mockReturnValue({
      isLoggedIn: true,
      email: 'test@example.com',
      id: 'test-id'
    });
    
    // データベースにユーザーが存在
    functions.mockFetchUserFromDatabase.mockReturnValue({
      userId: 'test-id',
      adminEmail: 'test@example.com'
    });
    
    const mockUserProperties = {
      getProperty: jest.fn((key) => {
        if (key === 'app_config') {
          return JSON.stringify({ setupComplete: true }); // セットアップ完了
        }
        return null;
      })
    };
    functions.mockPropertiesService.getUserProperties.mockReturnValue(mockUserProperties);

    const result = functions.getLoginStatus();
    
    expect(result.status).toBe('existing_user');
    expect(result.isLoggedIn).toBe(true);
    expect(result.adminUrl).toContain('page=admin');
    expect(result.message).toContain('ログイン完了');
    expect(functions.mockFetchUserFromDatabase).toHaveBeenCalledWith('adminEmail', 'test@example.com');
  });

  test('データベースエラーの場合でも動作する', () => {
    functions.mockGetCurrentUser.mockReturnValue({
      isLoggedIn: true,
      email: 'test@example.com',
      id: 'test-id'
    });
    
    // データベース検索でエラーが発生
    functions.mockFetchUserFromDatabase.mockImplementation(() => {
      throw new Error('Database connection failed');
    });

    const result = functions.getLoginStatus();
    
    expect(result.status).toBe('unregistered');
    expect(result.message).toContain('新規ユーザー登録が必要');
    expect(functions.mockFetchUserFromDatabase).toHaveBeenCalledWith('adminEmail', 'test@example.com');
  });
});