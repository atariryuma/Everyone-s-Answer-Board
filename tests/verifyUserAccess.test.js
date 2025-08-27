/**
 * @fileoverview verifyUserAccess関数のテスト
 */

const fs = require('fs');
const path = require('path');

// main.gsからverifyUserAccess関数を抽出
function loadVerifyUserAccessFunction() {
  const file = fs.readFileSync(path.join(__dirname, '../src/main.gs'), 'utf8');
  
  // verifyUserAccess関数を抽出
  const start = file.indexOf('function verifyUserAccess');
  if (start === -1) {
    throw new Error('verifyUserAccess function not found');
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
  
  const fn = new Function(`
    const getCurrentUser = arguments[0];
    const logError = arguments[1];
    const MAIN_ERROR_SEVERITY = { HIGH: 'high' };
    const MAIN_ERROR_CATEGORIES = { AUTH: 'auth' };
    
    ${fnStr}; 
    return { verifyUserAccess, getCurrentUser, logError };
  `)(mockGetCurrentUser, mockLogError);
  
  return { 
    ...fn, 
    mockGetCurrentUser, 
    mockLogError
  };
}

describe('verifyUserAccess関数テスト', () => {
  let functions;
  
  beforeEach(() => {
    functions = loadVerifyUserAccessFunction();
    jest.clearAllMocks();
  });

  test('UUIDが一致する場合はアクセスを許可', () => {
    const userId = '123e4567-e89b-12d3-a456-426614174000';
    functions.mockGetCurrentUser.mockReturnValue({
      email: 'user@example.com',
      id: userId,
      isLoggedIn: true
    });

    // エラーをスローしないことを確認
    expect(() => {
      functions.verifyUserAccess(userId);
    }).not.toThrow();
    
    expect(functions.mockLogError).not.toHaveBeenCalled();
  });

  test('メールアドレスが一致する場合はアクセスを許可', () => {
    const email = 'user@example.com';
    functions.mockGetCurrentUser.mockReturnValue({
      email: email,
      id: '123e4567-e89b-12d3-a456-426614174000',
      isLoggedIn: true
    });

    // メールアドレスをrequestUserIdとして渡す
    expect(() => {
      functions.verifyUserAccess(email);
    }).not.toThrow();
    
    expect(functions.mockLogError).not.toHaveBeenCalled();
  });

  test('ID・メールアドレスが一致しない場合はアクセス拒否', () => {
    functions.mockGetCurrentUser.mockReturnValue({
      email: 'user1@example.com',
      id: '123e4567-e89b-12d3-a456-426614174000',
      isLoggedIn: true
    });

    expect(() => {
      functions.verifyUserAccess('different-uuid-or-email');
    }).toThrow('Access denied');
    
    expect(functions.mockLogError).toHaveBeenCalledTimes(2); // エラーログと例外の両方
  });

  test('ユーザーがログインしていない場合はエラー', () => {
    functions.mockGetCurrentUser.mockReturnValue({
      email: null,
      id: null,
      isLoggedIn: false
    });

    expect(() => {
      functions.verifyUserAccess('any-id');
    }).toThrow('User not logged in');
    
    expect(functions.mockLogError).toHaveBeenCalled();
  });

  test('requestUserIdが無効な場合はエラー', () => {
    functions.mockGetCurrentUser.mockReturnValue({
      email: 'user@example.com',
      id: '123e4567-e89b-12d3-a456-426614174000',
      isLoggedIn: true
    });

    // null の場合
    expect(() => {
      functions.verifyUserAccess(null);
    }).toThrow('Invalid request user ID');

    // 空文字列の場合
    expect(() => {
      functions.verifyUserAccess('');
    }).toThrow('Invalid request user ID');

    // 数値の場合（型チェック）
    expect(() => {
      functions.verifyUserAccess(123);
    }).toThrow('Invalid request user ID');
  });

  test('getCurrentUserがエラーをスローした場合の処理', () => {
    functions.mockGetCurrentUser.mockImplementation(() => {
      throw new Error('Session error');
    });

    expect(() => {
      functions.verifyUserAccess('any-id');
    }).toThrow('Session error');
    
    expect(functions.mockLogError).toHaveBeenCalled();
  });
});