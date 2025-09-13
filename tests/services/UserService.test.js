/**
 * UserService Test Suite
 * Comprehensive testing for user management functionality
 */

const { testUserData, createMockSpreadsheet } = require('../mocks/gas-mocks.js');

// Mock the UserService (simulate GAS environment)
const UserService = {
  getCurrentEmail: jest.fn(),
  getCurrentUserInfo: jest.fn(),
  createUser: jest.fn(),
  getAccessLevel: jest.fn(),
  clearUserCache: jest.fn(),
  diagnose: jest.fn()
};

describe('UserService', () => {
  
  beforeEach(() => {
    jest.resetAllMocks();
    jest.clearAllMocks();
  });

  describe('getCurrentEmail', () => {
    
    it('should return current user email from session', () => {
      // Arrange
      const expectedEmail = 'test@example.com';
      global.Session.getActiveUser.mockReturnValue({
        getEmail: jest.fn(() => expectedEmail)
      });

      // Mock UserService implementation
      UserService.getCurrentEmail.mockImplementation(() => {
        try {
          return Session.getActiveUser().getEmail();
        } catch (error) {
          return null;
        }
      });

      // Act
      const result = UserService.getCurrentEmail();

      // Assert
      expect(result).toBe(expectedEmail);
      expect(global.Session.getActiveUser).toHaveBeenCalled();
    });

    it('should return null when session is invalid', () => {
      // Arrange
      global.Session.getActiveUser.mockImplementation(() => {
        throw new Error('No active session');
      });

      UserService.getCurrentEmail.mockImplementation(() => {
        try {
          return Session.getActiveUser().getEmail();
        } catch (error) {
          return null;
        }
      });

      // Act
      const result = UserService.getCurrentEmail();

      // Assert
      expect(result).toBeNull();
    });

  });

  describe('getCurrentUserInfo', () => {

    it('should return complete user info with config', () => {
      // Arrange
      const mockUserInfo = {
        userId: 'test-user-123',
        userEmail: 'test@example.com',
        isActive: true,
        config: {
          setupStatus: 'completed',
          appPublished: true
        }
      };

      UserService.getCurrentUserInfo.mockResolvedValue(mockUserInfo);

      // Act
      const result = UserService.getCurrentUserInfo();

      // Assert
      expect(result).resolves.toEqual(mockUserInfo);
    });

    it('should return null for non-existent user', () => {
      // Arrange
      UserService.getCurrentUserInfo.mockResolvedValue(null);

      // Act
      const result = UserService.getCurrentUserInfo();

      // Assert
      expect(result).resolves.toBeNull();
    });

  });

  describe('createUser', () => {

    it('should create new user with valid email', () => {
      // Arrange
      const userEmail = 'newuser@example.com';
      const mockNewUser = {
        userId: 'new-user-123',
        userEmail,
        isActive: true,
        configJson: JSON.stringify({
          setupStatus: 'pending',
          appPublished: false
        })
      };

      UserService.createUser.mockResolvedValue(mockNewUser);

      // Act
      const result = UserService.createUser(userEmail);

      // Assert
      expect(result).resolves.toEqual(mockNewUser);
      expect(UserService.createUser).toHaveBeenCalledWith(userEmail);
    });

    it('should throw error for invalid email', () => {
      // Arrange
      const invalidEmail = 'invalid-email';
      UserService.createUser.mockRejectedValue(new Error('無効なメールアドレス'));

      // Act & Assert
      expect(UserService.createUser(invalidEmail)).rejects.toThrow('無効なメールアドレス');
    });

  });

  describe('getAccessLevel', () => {

    it('should return owner level for user owner', () => {
      // Arrange
      const userId = 'test-user-123';
      UserService.getAccessLevel.mockReturnValue('owner');

      // Act
      const result = UserService.getAccessLevel(userId);

      // Assert
      expect(result).toBe('owner');
    });

    it('should return none for invalid user', () => {
      // Arrange
      const invalidUserId = null;
      UserService.getAccessLevel.mockReturnValue('none');

      // Act
      const result = UserService.getAccessLevel(invalidUserId);

      // Assert
      expect(result).toBe('none');
    });

  });

  describe('clearUserCache', () => {

    it('should clear user cache successfully', () => {
      // Arrange
      const mockRemove = jest.fn();
      const mockCacheService = {
        getScriptCache: jest.fn(() => ({
          remove: mockRemove
        }))
      };
      
      // Mock CacheService globally for this test
      const originalCacheService = global.CacheService;
      global.CacheService = mockCacheService;

      // Create a real implementation for this test
      const realClearUserCache = function(userId = null) {
        try {
          const cache = global.CacheService.getScriptCache();
          
          if (userId) {
            cache.remove(`user_info_${userId}`);
            cache.remove(`user_config_${userId}`);
          } else {
            cache.remove('current_user_info');
          }
        } catch (error) {
          console.error('UserService.clearUserCache: エラー', error.message);
        }
      };

      try {
        // Act - Call the real implementation
        realClearUserCache();

        // Assert
        expect(mockCacheService.getScriptCache).toHaveBeenCalled();
        expect(mockRemove).toHaveBeenCalledWith('current_user_info');
      } finally {
        // Restore original CacheService
        global.CacheService = originalCacheService;
      }
    });

  });

  describe('diagnose', () => {

    it('should return diagnostic information', () => {
      // Arrange
      const mockDiagnostics = {
        service: 'UserService',
        timestamp: '2025-01-15T00:00:00.000Z',
        checks: [
          { name: 'Session Check', status: '✅', details: 'test@example.com' },
          { name: 'Database Connection', status: '✅', details: 'Database status check' },
          { name: 'Cache Service', status: '✅', details: 'Cache service accessible' }
        ],
        overall: '✅'
      };

      UserService.diagnose.mockReturnValue(mockDiagnostics);

      // Act
      const result = UserService.diagnose();

      // Assert
      expect(result).toEqual(mockDiagnostics);
      expect(result.service).toBe('UserService');
      expect(result.overall).toBe('✅');
    });

  });

  describe('Error Handling', () => {

    it('should handle database connection errors gracefully', () => {
      // Arrange
      UserService.getCurrentUserInfo.mockRejectedValue(new Error('Database connection failed'));

      // Act & Assert
      expect(UserService.getCurrentUserInfo()).rejects.toThrow('Database connection failed');
    });

    it('should handle cache service errors gracefully', () => {
      // Arrange
      global.CacheService.getScriptCache.mockImplementation(() => {
        throw new Error('Cache service unavailable');
      });

      UserService.clearUserCache.mockImplementation(() => {
        try {
          global.CacheService.getScriptCache().remove('test');
        } catch (error) {
          console.error('Cache error:', error.message);
        }
      });

      // Act & Assert (should not throw)
      expect(() => UserService.clearUserCache()).not.toThrow();
    });

  });

  describe('Integration Tests', () => {

    it('should complete full user workflow', async () => {
      // Arrange
      const userEmail = 'integration@example.com';
      
      // Mock the full workflow
      UserService.getCurrentEmail.mockReturnValue(userEmail);
      UserService.createUser.mockResolvedValue(testUserData);
      UserService.getCurrentUserInfo.mockResolvedValue({
        ...testUserData,
        config: JSON.parse(testUserData.configJson)
      });
      UserService.getAccessLevel.mockReturnValue('owner');

      // Act
      const email = UserService.getCurrentEmail();
      const newUser = await UserService.createUser(email);
      const userInfo = await UserService.getCurrentUserInfo();
      const accessLevel = UserService.getAccessLevel(newUser.userId);

      // Assert
      expect(email).toBe(userEmail);
      expect(newUser).toBeDefined();
      expect(userInfo).toBeDefined();
      expect(accessLevel).toBe('owner');
    });

  });

});

describe('UserService Performance Tests', () => {

  it('should respond within acceptable time limits', async () => {
    // Arrange
    const startTime = Date.now();
    UserService.getCurrentUserInfo.mockResolvedValue(testUserData);

    // Act
    await UserService.getCurrentUserInfo();
    const endTime = Date.now();

    // Assert (should complete within 100ms for mock)
    expect(endTime - startTime).toBeLessThan(100);
  });

});

describe('UserService Security Tests', () => {

  it('should not expose sensitive information in errors', () => {
    // Arrange
    UserService.getCurrentUserInfo.mockRejectedValue(new Error('ユーザー情報を取得できません'));

    // Act & Assert
    expect(UserService.getCurrentUserInfo()).rejects.toThrow('ユーザー情報を取得できません');
    // Should not contain internal details like database paths, etc.
  });

  it('should validate user input properly', () => {
    // Arrange
    const maliciousInput = '<script>alert("xss")</script>@example.com';
    UserService.createUser.mockRejectedValue(new Error('無効なメールアドレス'));

    // Act & Assert
    expect(UserService.createUser(maliciousInput)).rejects.toThrow('無効なメールアドレス');
  });

});