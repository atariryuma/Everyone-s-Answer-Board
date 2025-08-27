/**
 * @fileoverview confirmUserRegistration関数のテスト
 */

const fs = require('fs');
const path = require('path');

// main.gsからconfirmUserRegistration関数を抽出
function loadConfirmUserRegistrationFunction() {
  const file = fs.readFileSync(path.join(__dirname, '../src/main.gs'), 'utf8');
  
  // confirmUserRegistration関数を抽出
  const start = file.indexOf('function confirmUserRegistration');
  if (start === -1) {
    throw new Error('confirmUserRegistration function not found');
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
  const mockFetchUserFromDatabase = jest.fn();
  const mockUtilities = {
    getUuid: jest.fn(() => 'test-uuid-12345')
  };
  const mockPropertiesService = {
    getUserProperties: jest.fn(() => ({
      setProperty: jest.fn()
    }))
  };
  
  const fn = new Function(`
    const logError = arguments[0];
    const fetchUserFromDatabase = arguments[1];
    const Utilities = arguments[2];
    const PropertiesService = arguments[3];
    const MAIN_ERROR_SEVERITY = { MEDIUM: 'medium' };
    const MAIN_ERROR_CATEGORIES = { AUTH: 'auth', DATA: 'data' };
    const JSON = global.JSON;
    
    ${fnStr}; 
    return { confirmUserRegistration, fetchUserFromDatabase, Utilities, PropertiesService };
  `)(mockLogError, mockFetchUserFromDatabase, mockUtilities, mockPropertiesService);
  
  return { ...fn, mockLogError, mockFetchUserFromDatabase, mockUtilities, mockPropertiesService };
}

describe('confirmUserRegistration関数テスト', () => {
  let functions;
  
  beforeEach(() => {
    functions = loadConfirmUserRegistrationFunction();
    jest.clearAllMocks();
  });

  test('無効なメールパラメータの場合はエラーを返す', () => {
    const result = functions.confirmUserRegistration(null);
    
    expect(result.status).toBe('error');
    expect(result.message).toContain('Invalid email parameter');
  });

  test('既存データベースユーザーの場合は登録済みを返す', () => {
    // データベースにユーザーが存在
    functions.mockFetchUserFromDatabase.mockReturnValue({
      userId: 'existing-user-id',
      adminEmail: 'test@example.com'
    });

    const result = functions.confirmUserRegistration('test@example.com');
    
    expect(result.status).toBe('success');
    expect(result.isRegistered).toBe(true);
    expect(result.message).toContain('既にデータベースに登録済み');
    expect(result.userId).toBe('existing-user-id');
    expect(functions.mockFetchUserFromDatabase).toHaveBeenCalledWith('adminEmail', 'test@example.com');
  });

  test('新規ユーザーの場合は登録準備完了を返す', () => {
    // データベースにユーザーが見つからない
    functions.mockFetchUserFromDatabase.mockReturnValue(null);
    
    const mockUserProperties = {
      setProperty: jest.fn()
    };
    functions.mockPropertiesService.getUserProperties.mockReturnValue(mockUserProperties);

    const result = functions.confirmUserRegistration('new@example.com');
    
    expect(result.status).toBe('success');
    expect(result.isRegistered).toBe(false);
    expect(result.message).toContain('登録準備が完了');
    expect(result.userId).toBe('test-uuid-12345');
    expect(functions.mockFetchUserFromDatabase).toHaveBeenCalledWith('adminEmail', 'new@example.com');
    expect(mockUserProperties.setProperty).toHaveBeenCalledWith(
      'user_new@example.com',
      expect.stringContaining('new@example.com')
    );
  });

  test('データベースエラーの場合でも新規登録処理を継続する', () => {
    // データベース検索でエラーが発生
    functions.mockFetchUserFromDatabase.mockImplementation(() => {
      throw new Error('Database connection failed');
    });
    
    const mockUserProperties = {
      setProperty: jest.fn()
    };
    functions.mockPropertiesService.getUserProperties.mockReturnValue(mockUserProperties);

    const result = functions.confirmUserRegistration('test@example.com');
    
    expect(result.status).toBe('success');
    expect(result.isRegistered).toBe(false);
    expect(result.message).toContain('登録準備が完了');
    expect(functions.mockLogError).toHaveBeenCalled();
    expect(functions.mockFetchUserFromDatabase).toHaveBeenCalledWith('adminEmail', 'test@example.com');
  });

  test('PropertiesServiceエラーの場合は適切にエラーを返す', () => {
    functions.mockFetchUserFromDatabase.mockReturnValue(null);
    
    functions.mockPropertiesService.getUserProperties.mockImplementation(() => {
      throw new Error('PropertiesService error');
    });

    const result = functions.confirmUserRegistration('test@example.com');
    
    expect(result.status).toBe('error');
    expect(result.message).toContain('PropertiesService error');
    expect(functions.mockLogError).toHaveBeenCalled();
  });
});