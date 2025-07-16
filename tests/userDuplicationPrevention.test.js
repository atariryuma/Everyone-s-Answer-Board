/**
 * ユーザー重複登録防止のテスト
 * Phase 1-3 の実装を検証
 */

const fs = require('fs');
const vm = require('vm');

describe('User Duplication Prevention Tests', () => {
  const databaseCode = fs.readFileSync('src/database.gs', 'utf8');
  const coreCode = fs.readFileSync('src/Core.gs', 'utf8');
  let context;
  
  beforeEach(() => {
    // Mock環境の構築
    context = {
      console,
      LockService: {
        getScriptLock: jest.fn(() => ({
          waitLock: jest.fn(() => true),
          tryLock: jest.fn(() => true),
          releaseLock: jest.fn()
        }))
      },
      PropertiesService: {
        getScriptProperties: jest.fn(() => ({
          getProperty: jest.fn(() => 'test-db-id')
        }))
      },
      Session: {
        getActiveUser: jest.fn(() => ({
          getEmail: jest.fn(() => 'test@example.com')
        }))
      },
      Utilities: {
        computeDigest: jest.fn(() => [0x12, 0x34, 0x56, 0x78, 0x9a, 0xbc, 0xde, 0xf0]),
        DigestAlgorithm: { SHA_256: 'sha256' },
        Charset: { UTF_8: 'utf8' }
      },
      SCRIPT_PROPS_KEYS: {
        DATABASE_SPREADSHEET_ID: 'DATABASE_SPREADSHEET_ID'
      },
      DB_SHEET_CONFIG: {
        SHEET_NAME: 'Users',
        HEADERS: ['userId', 'adminEmail', 'spreadsheetId', 'spreadsheetUrl', 'createdAt', 'configJson', 'setupStatus', 'lastAccessedAt', 'isActive']
      },
      // Mock functions
      findUserByEmail: jest.fn(),
      updateUser: jest.fn(),
      getSheetsService: jest.fn(),
      appendSheetsData: jest.fn(),
      invalidateUserCache: jest.fn(),
      debugLog: jest.fn(),
      getDeployUserDomainInfo: jest.fn(() => ({ isDomainMatch: true })),
      generateAppUrls: jest.fn(() => ({
        adminUrl: 'https://admin.example.com',
        viewUrl: 'https://view.example.com'
      }))
    };
    
    vm.createContext(context);
    vm.runInContext(databaseCode, context);
    vm.runInContext(coreCode, context);
  });

  describe('findOrCreateUser - Atomic User Operations', () => {
    test('should create new user when not exists', () => {
      // Setup: ユーザーが存在しない
      context.findUserByEmail.mockReturnValue(null);
      
      // Execute: 新規ユーザー作成
      const result = context.findOrCreateUser('newuser@example.com');
      
      // Verify: 新規ユーザーが作成される
      expect(result).toHaveProperty('userId');
      expect(result).toHaveProperty('isNewUser', true);
      expect(result).toHaveProperty('userInfo');
      expect(context.appendSheetsData).toHaveBeenCalled();
      expect(context.debugLog).toHaveBeenCalledWith(
        'findOrCreateUser: 新規ユーザー作成完了', 
        expect.objectContaining({ adminEmail: 'newuser@example.com' })
      );
    });

    test('should return existing user when exists', () => {
      // Setup: 既存ユーザーが存在
      const existingUser = {
        userId: 'existing-user-123',
        adminEmail: 'existing@example.com',
        configJson: '{"test": true}'
      };
      context.findUserByEmail.mockReturnValue(existingUser);
      
      // Execute: 既存ユーザー取得
      const result = context.findOrCreateUser('existing@example.com');
      
      // Verify: 既存ユーザーが返される
      expect(result).toHaveProperty('userId', 'existing-user-123');
      expect(result).toHaveProperty('isNewUser', false);
      expect(result).toHaveProperty('userInfo', existingUser);
      expect(context.appendSheetsData).not.toHaveBeenCalled();
      expect(context.updateUser).toHaveBeenCalled();
    });

    test('should handle concurrent access with lock', () => {
      // Setup: ロックが成功する
      const mockLock = {
        waitLock: jest.fn(() => true),
        releaseLock: jest.fn()
      };
      context.LockService.getScriptLock.mockReturnValue(mockLock);
      context.findUserByEmail.mockReturnValue(null);
      
      // Execute
      context.findOrCreateUser('concurrent@example.com');
      
      // Verify: ロックが適切に取得・解放される
      expect(mockLock.waitLock).toHaveBeenCalledWith(10000);
      expect(mockLock.releaseLock).toHaveBeenCalled();
    });

    test('should throw error when lock fails', () => {
      // Setup: ロック取得失敗
      const mockLock = {
        waitLock: jest.fn(() => false),
        releaseLock: jest.fn()
      };
      context.LockService.getScriptLock.mockReturnValue(mockLock);
      
      // Execute & Verify: エラーが投げられる
      expect(() => {
        context.findOrCreateUser('lockfail@example.com');
      }).toThrow('システムが混雑しています');
    });

    test('should generate consistent user ID from email', () => {
      // Setup
      context.findUserByEmail.mockReturnValue(null);
      
      // Execute: 同じメールアドレスで複数回実行
      const result1 = context.findOrCreateUser('consistent@example.com');
      const result2 = context.findOrCreateUser('consistent@example.com');
      
      // Verify: 同じUserIDが生成される（決定論的）
      expect(result1.userId).toBe(result2.userId);
    });
  });

  describe('ensureUserExists - Improved Registration', () => {
    test('should use findOrCreateUser for safe registration', () => {
      // Setup
      context.findUserByEmail.mockReturnValue(null);
      
      // Execute
      const result = context.ensureUserExists('newuser@example.com');
      
      // Verify: findOrCreateUserが使用される
      expect(result).toHaveProperty('userId');
      expect(result).toHaveProperty('setupRequired', true);
      expect(result.message).toContain('ユーザー登録が完了しました');
    });

    test('should handle domain restrictions', () => {
      // Setup: ドメイン制限あり
      context.getDeployUserDomainInfo.mockReturnValue({
        deployDomain: 'allowed.com',
        currentDomain: 'forbidden.com',
        isDomainMatch: false
      });
      
      // Execute & Verify
      expect(() => {
        context.ensureUserExists('user@forbidden.com');
      }).toThrow('ドメインアクセスが制限されています');
    });

    test('should validate email authentication', () => {
      // Setup: セッションユーザーと異なるメール
      context.Session.getActiveUser.mockReturnValue({
        getEmail: () => 'session@example.com'
      });
      
      // Execute & Verify
      expect(() => {
        context.ensureUserExists('different@example.com');
      }).toThrow('認証エラー');
    });
  });

  describe('quickStartSetup - Simplified Flow', () => {
    test('should auto-detect user when no userId provided', () => {
      // Setup
      context.findUserByEmail.mockReturnValue(null);
      context.createStudyQuestForm = jest.fn(() => ({
        formUrl: 'https://form.example.com',
        spreadsheetId: 'sheet123',
        spreadsheetUrl: 'https://sheet.example.com',
        sheetName: 'Sheet1'
      }));
      context.createUserFolder = jest.fn();
      
      // Execute: userIdなしでクイックスタート
      const result = context.quickStartSetup();
      
      // Verify: 自動的にユーザーが作成される
      expect(context.debugLog).toHaveBeenCalledWith(
        'quickStartSetup: ユーザー確保完了',
        expect.objectContaining({ isNewUser: true })
      );
    });

    test('should verify existing userId when provided', () => {
      // Setup
      const existingUser = {
        userId: 'existing-123',
        adminEmail: 'test@example.com',
        configJson: '{}'
      };
      context.verifyUserAccess = jest.fn();
      context.findUserById = jest.fn(() => existingUser);
      context.createStudyQuestForm = jest.fn(() => ({
        formUrl: 'https://form.example.com',
        spreadsheetId: 'sheet123'
      }));
      context.createUserFolder = jest.fn();
      
      // Execute
      const result = context.quickStartSetup('existing-123');
      
      // Verify: 認証が実行される
      expect(context.verifyUserAccess).toHaveBeenCalledWith('existing-123');
      expect(context.findUserById).toHaveBeenCalledWith('existing-123');
    });

    test('should skip setup if already completed', () => {
      // Setup: セットアップ済みユーザー
      const completedUser = {
        userId: 'completed-123',
        adminEmail: 'test@example.com',
        configJson: '{"formCreated": true}',
        spreadsheetId: 'existing-sheet'
      };
      context.verifyUserAccess = jest.fn();
      context.findUserById = jest.fn(() => completedUser);
      
      // Execute
      const result = context.quickStartSetup('completed-123');
      
      // Verify: スキップされる
      expect(result.status).toBe('already_completed');
      expect(result.message).toContain('既に完了しています');
    });
  });

  describe('Race Condition Prevention', () => {
    test('should handle simultaneous user creation attempts', async () => {
      // Setup: 複数の同時リクエストをシミュレート
      context.findUserByEmail
        .mockReturnValueOnce(null) // 1回目: ユーザー存在しない
        .mockReturnValueOnce(null); // 2回目: ユーザー存在しない
      
      const mockLock = {
        waitLock: jest.fn()
          .mockReturnValueOnce(true)  // 1回目: ロック成功
          .mockReturnValueOnce(false), // 2回目: ロック失敗
        releaseLock: jest.fn()
      };
      context.LockService.getScriptLock.mockReturnValue(mockLock);
      
      // Execute: 並行実行をシミュレート
      const result1 = context.findOrCreateUser('race@example.com');
      
      expect(() => {
        context.findOrCreateUser('race@example.com');
      }).toThrow('システムが混雑しています');
      
      // Verify: 1つ目のみ成功
      expect(result1.isNewUser).toBe(true);
    });

    test('should validate lock timeout behavior', () => {
      // Setup: ロックタイムアウト
      const mockLock = {
        waitLock: jest.fn(() => false),
        releaseLock: jest.fn()
      };
      context.LockService.getScriptLock.mockReturnValue(mockLock);
      
      // Execute & Verify
      expect(() => {
        context.findOrCreateUser('timeout@example.com');
      }).toThrow('システムが混雑しています');
      
      expect(mockLock.waitLock).toHaveBeenCalledWith(10000); // 10秒
    });
  });

  describe('Error Recovery and Logging', () => {
    test('should log detailed debugging information', () => {
      // Setup
      context.findUserByEmail.mockReturnValue(null);
      
      // Execute
      context.findOrCreateUser('debug@example.com');
      
      // Verify: デバッグログが記録される
      expect(context.debugLog).toHaveBeenCalledWith(
        'findOrCreateUser: ロック取得成功',
        expect.objectContaining({ adminEmail: 'debug@example.com' })
      );
      expect(context.debugLog).toHaveBeenCalledWith(
        'findOrCreateUser: 新規ユーザー作成開始',
        expect.objectContaining({ adminEmail: 'debug@example.com' })
      );
    });

    test('should maintain backward compatibility', () => {
      // Setup
      context.findUserByEmail.mockReturnValue(null);
      
      // Execute: 旧関数を呼び出し
      const result = context.registerNewUser('legacy@example.com');
      
      // Verify: 非推奨警告が表示される
      expect(console.warn).toHaveBeenCalledWith(
        expect.stringContaining('registerNewUser は非推奨です')
      );
      expect(result).toHaveProperty('userId');
    });

    test('should handle database errors gracefully', () => {
      // Setup: データベースエラー
      context.findUserByEmail.mockReturnValue(null);
      context.appendSheetsData.mockImplementation(() => {
        throw new Error('Database connection failed');
      });
      
      // Execute & Verify
      expect(() => {
        context.findOrCreateUser('dberror@example.com');
      }).toThrow('Database connection failed');
    });
  });

  describe('Performance and Monitoring', () => {
    test('should complete user creation within acceptable time', () => {
      // Setup
      context.findUserByEmail.mockReturnValue(null);
      
      // Execute with timing
      const start = Date.now();
      context.findOrCreateUser('performance@example.com');
      const duration = Date.now() - start;
      
      // Verify: 妥当な時間内で完了（テスト環境では1秒以内）
      expect(duration).toBeLessThan(1000);
    });

    test('should track system health metrics', () => {
      // Execute multiple operations
      context.findUserByEmail
        .mockReturnValueOnce(null)
        .mockReturnValueOnce({ userId: 'existing', adminEmail: 'test@example.com' });
      
      context.findOrCreateUser('new@example.com');
      context.findOrCreateUser('existing@example.com');
      
      // Verify: 適切な操作が実行される
      expect(context.appendSheetsData).toHaveBeenCalledTimes(1); // 新規のみ
      expect(context.updateUser).toHaveBeenCalledTimes(1); // 既存のみ
    });
  });

  afterEach(() => {
    jest.clearAllMocks();
  });
});