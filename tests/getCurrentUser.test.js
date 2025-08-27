/**
 * @fileoverview getCurrentUser関数のテスト
 */

const fs = require('fs');
const path = require('path');

// main.gsからgetCurrentUser関数を抽出
function loadGetCurrentUserFunction() {
  const file = fs.readFileSync(path.join(__dirname, '../src/main.gs'), 'utf8');
  
  // getCurrentUser関数を抽出
  const start = file.indexOf('function getCurrentUser');
  if (start === -1) {
    throw new Error('getCurrentUser function not found');
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
  const mockSession = {
    getActiveUser: jest.fn(() => ({
      getEmail: jest.fn()
    }))
  };
  
  const mockFetchUserFromDatabase = jest.fn();
  const mockLogError = jest.fn();
  const mockConsole = {
    warn: jest.fn()
  };
  
  const fn = new Function(`
    const Session = arguments[0];
    const fetchUserFromDatabase = arguments[1];
    const logError = arguments[2];
    const console = arguments[3];
    const MAIN_ERROR_SEVERITY = { HIGH: 'high' };
    const MAIN_ERROR_CATEGORIES = { AUTH: 'auth' };
    
    ${fnStr}; 
    return { getCurrentUser, Session, fetchUserFromDatabase };
  `)(mockSession, mockFetchUserFromDatabase, mockLogError, mockConsole);
  
  return { 
    ...fn, 
    mockSession, 
    mockFetchUserFromDatabase, 
    mockLogError,
    mockConsole
  };
}

describe('getCurrentUser関数テスト', () => {
  let functions;
  let mockUser;
  
  beforeEach(() => {
    functions = loadGetCurrentUserFunction();
    mockUser = {
      getEmail: jest.fn()
    };
    functions.mockSession.getActiveUser.mockReturnValue(mockUser);
    jest.clearAllMocks();
  });

  test('データベースにユーザーが存在する場合はUUIDを返す', () => {
    const email = 'test@example.com';
    const uuid = '123e4567-e89b-12d3-a456-426614174000';
    
    mockUser.getEmail.mockReturnValue(email);
    functions.mockFetchUserFromDatabase.mockReturnValue({
      userId: uuid,
      adminEmail: email
    });

    const result = functions.getCurrentUser();
    
    expect(result.email).toBe(email);
    expect(result.id).toBe(uuid);
    expect(result.isLoggedIn).toBe(true);
    expect(functions.mockFetchUserFromDatabase).toHaveBeenCalledWith('adminEmail', email);
  });

  test('データベースにユーザーが存在しない場合はメールアドレスをIDとする', () => {
    const email = 'newuser@example.com';
    
    mockUser.getEmail.mockReturnValue(email);
    functions.mockFetchUserFromDatabase.mockReturnValue(null);

    const result = functions.getCurrentUser();
    
    expect(result.email).toBe(email);
    expect(result.id).toBe(email); // メールアドレスがIDとして使用される
    expect(result.isLoggedIn).toBe(true);
  });

  test('データベースエラーの場合もメールアドレスをIDとする', () => {
    const email = 'user@example.com';
    
    mockUser.getEmail.mockReturnValue(email);
    functions.mockFetchUserFromDatabase.mockImplementation(() => {
      throw new Error('Database connection failed');
    });

    const result = functions.getCurrentUser();
    
    expect(result.email).toBe(email);
    expect(result.id).toBe(email);
    expect(result.isLoggedIn).toBe(true);
    expect(functions.mockConsole.warn).toHaveBeenCalledWith(
      expect.stringContaining('データベース検索失敗'),
      'Database connection failed'
    );
  });

  test('メールアドレスが取得できない場合はエラー状態を返す', () => {
    mockUser.getEmail.mockReturnValue(null);

    const result = functions.getCurrentUser();
    
    expect(result.email).toBeNull();
    expect(result.id).toBeNull();
    expect(result.isLoggedIn).toBe(false);
    expect(result.error).toBe('No user email found');
    expect(functions.mockLogError).toHaveBeenCalled();
  });

  test('Session.getActiveUserがエラーをスローした場合の処理', () => {
    functions.mockSession.getActiveUser.mockImplementation(() => {
      throw new Error('Session error');
    });

    const result = functions.getCurrentUser();
    
    expect(result.email).toBeNull();
    expect(result.id).toBeNull();
    expect(result.isLoggedIn).toBe(false);
    expect(result.error).toBe('Session error');
    expect(functions.mockLogError).toHaveBeenCalled();
  });
});