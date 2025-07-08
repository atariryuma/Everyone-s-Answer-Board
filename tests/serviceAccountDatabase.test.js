const { setupGlobalMocks, ensureCodeEvaluated, resetMocks } = require('./shared-mocks');

describe('Service Account Database Architecture Tests', () => {
  beforeEach(() => {
    resetMocks();
    setupGlobalMocks();
    ensureCodeEvaluated();
  });

  test('1. セットアップ関数のテスト', () => {
    console.log('\n=== Test 1: Setup Functions ===');
    
    const mockCredsJson = JSON.stringify({
      "type": "service_account",
      "project_id": "test-project",
      "client_email": "test@test-project.iam.gserviceaccount.com",
      "private_key": "-----BEGIN PRIVATE KEY-----\ntest-key\n-----END PRIVATE KEY-----"
    });
    
    const mockDbId = '1BcD3fGhIjKlMnOpQrStUvWxYz0123456789ABCDEFGH';
    
    console.log('Testing setupApplication...');
    expect(() => setupApplication(mockCredsJson, mockDbId)).not.toThrow();
    console.log('✅ setupApplication completed successfully');
  });

  test('2. サービスアカウント認証のテスト', () => {
    console.log('\n=== Test 2: Service Account Authentication ===');
    
    console.log('Testing getServiceAccountToken...');
    const token = getServiceAccountToken();
    
    expect(token).toBeDefined();
    expect(token).toMatch(/^mock-access-token-/);
    console.log(`✅ Token retrieved: ${token.substring(0, 20)}...`);
  });

  test('3. Sheets APIサービスのテスト', () => {
    console.log('\n=== Test 3: Sheets API Service ===');
    
    console.log('Testing getSheetsService...');
    const service = getSheetsService();
    
    expect(service).toBeDefined();
    expect(service.baseUrl).toBeDefined();
    
    // Mocking the actual Sheets API call within the test
    global.UrlFetchApp.fetch = jest.fn((url, options) => {
      if (url.includes('/spreadsheets/') && !url.includes('/values/')) {
        return {
          getContentText: () => JSON.stringify({
            properties: { title: `Test Database Spreadsheet` },
            sheets: [
              { properties: { title: 'Users', sheetId: 0 } },
              { properties: { title: 'フォームの回答 1', sheetId: 1 } },
              { properties: { title: '名簿', sheetId: 2 } }
            ]
          })
        };
      }
      return { getContentText: () => JSON.stringify({}) };
    });

    const spreadsheet = getSpreadsheetsData(service, 'test-spreadsheet-id');
    
    expect(spreadsheet).toBeDefined();
    expect(spreadsheet.properties.title).toBe(`Test Database Spreadsheet`);
    console.log(`✅ Spreadsheet retrieved: ${spreadsheet.properties.title}`);
  });

  test('4. データベース操作のテスト', () => {
    console.log('\n=== Test 4: Database Operations ===');
    
    // Mocking UrlFetchApp.fetch for database operations
    global.UrlFetchApp.fetch = jest.fn((url, options) => {
      if (url.includes('oauth2/v4/token')) {
        return {
          getContentText: () => JSON.stringify({
            access_token: 'mock-access-token-' + Date.now(),
            expires_in: 3600
          })
        };
      } else if (url.includes('/values:batchGet')) {
        return {
          getContentText: () => JSON.stringify({
            valueRanges: [{
              range: 'Users!A:H',
              values: mockDatabase
            }]
          })
        };
      } else if (url.includes(':append')) {
        const newData = JSON.parse(options.payload);
        if (newData.values && newData.values[0]) {
          mockDatabase.push(newData.values[0]);
        }
        return {
          getContentText: () => JSON.stringify({
            updates: { updatedRows: 1, updatedCells: newData.values[0].length }
          })
        };
      } else if (url.includes(':batchUpdate')) {
        return {
          getContentText: () => JSON.stringify({
            replies: [{ updatedRows: 1 }]
          })
        };
      } else if (url.includes('/values/') && options && options.method === 'put') {
        return {
          getContentText: () => JSON.stringify({
            updatedRows: 1,
            updatedCells: 1
          })
        };
      }
      throw new Error(`Unexpected API call: ${url}`);
    });

    // ユーザー検索テスト
    console.log('Testing findUserById...');
    const user = findUserById('test-user-id-12345');
    
    expect(user).toBeDefined();
    expect(user.userId).toBe('test-user-id-12345');
    console.log(`✅ User found: ${user.adminEmail}`);
    
    // メールでのユーザー検索テスト
    console.log('Testing findUserByEmail...');
    const userByEmail = findUserByEmail('test@example.com');
    
    expect(userByEmail).toBeDefined();
    expect(userByEmail.adminEmail).toBe('test@example.com');
    console.log(`✅ User found by email: ${userByEmail.userId}`);
    
    // 新規ユーザー作成テスト
    console.log('Testing createUser...');
    const newUserData = {
      userId: 'new-test-user-' + Date.now(),
      adminEmail: 'newuser@example.com',
      spreadsheetId: 'new-spreadsheet-id',
      spreadsheetUrl: 'https://docs.google.com/spreadsheets/d/new',
      createdAt: new Date().toISOString(),
      configJson: '{"formUrl":"https://forms.google.com/new"}',
      lastAccessedAt: new Date().toISOString(),
      isActive: 'true'
    };
    
    const createdUser = createUser(newUserData);
    expect(createdUser).toBeDefined();
    expect(createdUser.userId).toBe(newUserData.userId);
    console.log(`✅ User created: ${createdUser.userId}`);
    
    // ユーザー更新テスト
    console.log('Testing updateUser...');
    const updateResult = updateUser('test-user-id-12345', {
      lastAccessedAt: new Date().toISOString()
    });
    
    expect(updateResult).toBeDefined();
    expect(updateResult.success).toBe(true);
    console.log('✅ User updated successfully');
  });

  test('5. キャッシュ管理のテスト', () => {
    console.log('\n=== Test 5: Cache Management ===');
    
    // ユーザー情報キャッシュテスト
    console.log('Testing user info cache...');
    const testUserId = 'cache-test-user';
    const testUserInfo = { userId: testUserId, adminEmail: 'cache@test.com' };
    
    // キャッシュに保存
    setCachedUserInfo(testUserId, testUserInfo);
    
    // キャッシュから取得
    const cachedInfo = getCachedUserInfo(testUserId);
    
    expect(cachedInfo).toBeDefined();
    expect(cachedInfo.userId).toBe(testUserId);
    console.log(`✅ Cache test passed: ${cachedInfo.adminEmail}`);
    
    // キャッシュクリアテスト
    console.log('Testing cache clear...');
    clearAllCaches();
    
    const clearedCache = getCachedUserInfo(testUserId);
    expect(clearedCache).toBeNull();
    console.log('✅ Cache clear test passed');
  });

  test('6. ヘルパー関数のテスト', () => {
    console.log('\n=== Test 6: Helper Functions ===');
    
    // メール検証テスト
    console.log('Testing email validation...');
    const validEmail = 'test@example.com';
    const invalidEmail = 'invalid-email';
    
    expect(isValidEmail(validEmail)).toBe(true);
    expect(isValidEmail(invalidEmail)).toBe(false);
    console.log('✅ Email validation test passed');
    
    // ドメイン抽出テスト
    console.log('Testing domain extraction...');
    const domain = getEmailDomain('user@example.com');
    
    expect(domain).toBe('example.com');
    console.log('✅ Domain extraction test passed');
    
    // リアクション文字列解析テスト
    console.log('Testing reaction string parsing...');
    const reactionString = 'user1@example.com, user2@example.com, user3@example.com';
    const parsedReactions = parseReactionString(reactionString);
    
    expect(parsedReactions).toHaveLength(3);
    expect(parsedReactions[0]).toBe('user1@example.com');
    console.log('✅ Reaction string parsing test passed');
  });

  test('7. エラーハンドリングのテスト', () => {
    console.log('\n=== Test 7: Error Handling ===');
    
    // 存在しないユーザーの検索
    console.log('Testing non-existent user search...');
    const nonExistentUser = findUserById('non-existent-user-id');
    
    expect(nonExistentUser).toBeNull();
    console.log('✅ Non-existent user handling test passed');
    
    // 無効なメールでの検索
    console.log('Testing invalid email search...');
    const invalidEmailUser = findUserByEmail('invalid-email-format');
    
    expect(invalidEmailUser).toBeNull();
    console.log('✅ Invalid email handling test passed');
    
    // 安全なスプレッドシート操作テスト
    console.log('Testing safe spreadsheet operation...');
    const result = safeSpreadsheetOperation(() => {
      throw new Error('Test error');
    }, 'fallback-value');
    
    expect(result).toBe('fallback-value');
    console.log('✅ Safe operation test passed');
  });
});
