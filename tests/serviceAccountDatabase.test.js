/**
 * Service Account Database Architecture Test
 * 新アーキテクチャのデータベース管理機能をテストします
 */

const { setupGlobalMocks, ensureCodeEvaluated, resetMocks } = require('./shared-mocks');

// Setup the mock environment
setupGlobalMocks();

// テスト関数
describe('Service Account Database Architecture Tests', () => {
  
  test('1. セットアップ関数のテスト', () => {
    console.log('\n=== Test 1: Setup Functions ===');
    
    try {
      const mockCredsJson = JSON.stringify({
        "type": "service_account",
        "project_id": "test-project",
        "client_email": "test@test-project.iam.gserviceaccount.com",
        "private_key": "-----BEGIN PRIVATE KEY-----\ntest-key\n-----END PRIVATE KEY-----"
      });
      
      // 実際のGoogleスプレッドシートIDは44文字
      const mockDbId = '1BcD3fGhIjKlMnOpQrStUvWxYz0123456789ABCDEFGH';
      
      console.log('Testing setupApplication...');
      setupApplication(mockCredsJson, mockDbId);
      console.log('✅ setupApplication completed successfully');
      
    } catch (error) {
      console.error('❌ Setup test failed:', error.message);
      throw error;
    }
  });

  test('2. サービスアカウント認証のテスト', () => {
    console.log('\n=== Test 2: Service Account Authentication ===');
    
    try {
      console.log('Testing getServiceAccountToken...');
      const token = getServiceAccountToken();
      
      if (!token || !token.startsWith('mock-access-token-')) {
        throw new Error('Invalid access token returned');
      }
      
      console.log(`✅ Token retrieved: ${token.substring(0, 20)}...`);
      
    } catch (error) {
      console.error('❌ Authentication test failed:', error.message);
      throw error;
    }
  });

  test('3. Sheets APIサービスのテスト', () => {
    console.log('\n=== Test 3: Sheets API Service ===');
    
    try {
      console.log('Testing getSheetsService...');
      const service = getSheetsService();
      
      if (!service || !service.spreadsheets) {
        throw new Error('Invalid service object returned');
      }
      
      console.log('Testing spreadsheet get operation...');
      const spreadsheet = service.spreadsheets.get('test-spreadsheet-id');
      
      if (!spreadsheet || !spreadsheet.properties) {
        throw new Error('Invalid spreadsheet object returned');
      }
      
      console.log(`✅ Spreadsheet retrieved: ${spreadsheet.properties.title}`);
      
    } catch (error) {
      console.error('❌ Sheets API test failed:', error.message);
      throw error;
    }
  });

  test('4. データベース操作のテスト', () => {
    console.log('\n=== Test 4: Database Operations ===');
    
    try {
      // ユーザー検索テスト
      console.log('Testing findUserById...');
      const user = findUserById('test-user-id-12345');
      
      if (!user || user.userId !== 'test-user-id-12345') {
        throw new Error('User not found or invalid user data');
      }
      
      console.log(`✅ User found: ${user.adminEmail}`);
      
      // メールでのユーザー検索テスト
      console.log('Testing findUserByEmail...');
      const userByEmail = findUserByEmail('test@example.com');
      
      if (!userByEmail || userByEmail.adminEmail !== 'test@example.com') {
        throw new Error('User not found by email or invalid user data');
      }
      
      console.log(`✅ User found by email: ${userByEmail.userId}`);
      
      // 新規ユーザー作成テスト
      console.log('Testing createUserInDb...');
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
      
      const createdUser = createUserInDb(newUserData);
      console.log(`✅ User created: ${createdUser.userId}`);
      
      // ユーザー更新テスト
      console.log('Testing updateUserInDb...');
      const updateResult = updateUserInDb('test-user-id-12345', {
        lastAccessedAt: new Date().toISOString()
      });
      
      if (!updateResult || !updateResult.success) {
        throw new Error('User update failed');
      }
      
      console.log('✅ User updated successfully');
      
    } catch (error) {
      console.error('❌ Database operations test failed:', error.message);
      throw error;
    }
  });

  test('5. キャッシュ管理のテスト', () => {
    console.log('\n=== Test 5: Cache Management ===');
    
    try {
      // ユーザー情報キャッシュテスト
      console.log('Testing user info cache...');
      const testUserId = 'cache-test-user';
      const testUserInfo = { userId: testUserId, adminEmail: 'cache@test.com' };
      
      // キャッシュに保存
      setCachedUserInfo(testUserId, testUserInfo);
      
      // キャッシュから取得
      const cachedInfo = getCachedUserInfo(testUserId);
      
      if (!cachedInfo || cachedInfo.userId !== testUserId) {
        throw new Error('Cache operation failed');
      }
      
      console.log(`✅ Cache test passed: ${cachedInfo.adminEmail}`);
      
      // キャッシュクリアテスト
      console.log('Testing cache clear...');
      clearAllCaches();
      
      const clearedCache = getCachedUserInfo(testUserId);
      if (clearedCache !== null) {
        throw new Error('Cache clear failed');
      }
      
      console.log('✅ Cache clear test passed');
      
    } catch (error) {
      console.error('❌ Cache management test failed:', error.message);
      throw error;
    }
  });

  test('6. ヘルパー関数のテスト', () => {
    console.log('\n=== Test 6: Helper Functions ===');
    
    try {
      // メール検証テスト
      console.log('Testing email validation...');
      const validEmail = 'test@example.com';
      const invalidEmail = 'invalid-email';
      
      if (!isValidEmail(validEmail)) {
        throw new Error('Valid email rejected');
      }
      
      if (isValidEmail(invalidEmail)) {
        throw new Error('Invalid email accepted');
      }
      
      console.log('✅ Email validation test passed');
      
      // ドメイン抽出テスト
      console.log('Testing domain extraction...');
      const domain = getEmailDomain('user@example.com');
      
      if (domain !== 'example.com') {
        throw new Error(`Expected 'example.com', got '${domain}'`);
      }
      
      console.log('✅ Domain extraction test passed');
      
      // リアクション文字列解析テスト
      console.log('Testing reaction string parsing...');
      const reactionString = 'user1@example.com, user2@example.com, user3@example.com';
      const parsedReactions = parseReactionString(reactionString);
      
      if (parsedReactions.length !== 3 || parsedReactions[0] !== 'user1@example.com') {
        throw new Error('Reaction string parsing failed');
      }
      
      console.log('✅ Reaction string parsing test passed');
      
    } catch (error) {
      console.error('❌ Helper functions test failed:', error.message);
      throw error;
    }
  });

  test('7. エラーハンドリングのテスト', () => {
    console.log('\n=== Test 7: Error Handling ===');
    
    try {
      // 存在しないユーザーの検索
      console.log('Testing non-existent user search...');
      const nonExistentUser = findUserById('non-existent-user-id');
      
      if (nonExistentUser !== null) {
        throw new Error('Non-existent user should return null');
      }
      
      console.log('✅ Non-existent user handling test passed');
      
      // 無効なメールでの検索
      console.log('Testing invalid email search...');
      const invalidEmailUser = findUserByEmail('invalid-email-format');
      
      if (invalidEmailUser !== null) {
        throw new Error('Invalid email search should return null');
      }
      
      console.log('✅ Invalid email handling test passed');
      
      // 安全なスプレッドシート操作テスト
      console.log('Testing safe spreadsheet operation...');
      const result = safeSpreadsheetOperation(() => {
        throw new Error('Test error');
      }, 'fallback-value');
      
      if (result !== 'fallback-value') {
        throw new Error('Safe operation fallback failed');
      }
      
      console.log('✅ Safe operation test passed');
      
    } catch (error) {
      console.error('❌ Error handling test failed:', error.message);
      throw error;
    }
  });

});

// テスト実行関数
function runTests() {
  ensureCodeEvaluated(); // Code.gsを評価
  console.log('🚀 Starting Service Account Database Architecture Tests...\n');
  
  const tests = [
    () => describe.tests[0](), // Setup
    () => describe.tests[1](), // Authentication  
    () => describe.tests[2](), // Sheets API
    () => describe.tests[3](), // Database Operations
    () => describe.tests[4](), // Cache Management
    () => describe.tests[5](), // Helper Functions
    () => describe.tests[6]()  // Error Handling
  ];
  
  let passedTests = 0;
  let totalTests = tests.length;
  
  for (let i = 0; i < tests.length; i++) {
    try {
      tests[i]();
      passedTests++;
    } catch (error) {
      console.error(`\n❌ Test ${i + 1} failed:`, error.message);
    }
  }
  
  console.log('\n' + '='.repeat(60));
  console.log(`🏁 Test Results: ${passedTests}/${totalTests} tests passed`);
  
  if (passedTests === totalTests) {
    console.log('🎉 All tests passed! The new architecture is working correctly.');
    return true;
  } else {
    console.log('⚠️  Some tests failed. Please check the implementation.');
    return false;
  }
}

// 簡易テストフレームワーク
function describe(name, callback) {
  console.log(`\n📋 ${name}`);
  if (typeof callback === 'function') {
    callback();
  }
}

function test(name, callback) {
  if (!describe.tests) {
    describe.tests = [];
  }
  describe.tests.push(() => {
    console.log(`\n🧪 ${name}`);
    callback();
  });
}

// メインテスト実行
if (require.main === module) {
  runTests();
}

module.exports = { runTests };