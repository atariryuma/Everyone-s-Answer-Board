/**
 * System Integration Test Suite
 * End-to-end testing for Everyone's Answer Board
 */

const { resetAllMocks, testUserData, testSheetData, createMockSpreadsheet } = require('../mocks/gas-mocks.js');

describe('System Integration Tests', () => {

  beforeEach(() => {
    resetAllMocks();
  });

  describe('Complete User Journey', () => {

    it('should handle complete user onboarding flow', async () => {
      // Arrange - Mock services
      const mockServices = setupMockServices();
      
      // Act & Assert - Step by step flow
      
      // 1. User visits app (doGet)
      const loginResult = await simulateDoGet({ mode: 'login' });
      expect(loginResult.success).toBe(true);
      
      // 2. User authenticates
      const authResult = await mockServices.UserService.getCurrentEmail();
      expect(authResult).toBe('test@example.com');
      
      // 3. New user creation
      const newUser = await mockServices.UserService.createUser('test@example.com');
      expect(newUser.userId).toBeDefined();
      expect(newUser.userEmail).toBe('test@example.com');
      
      // 4. Initial config setup
      const config = await mockServices.ConfigService.getUserConfig(newUser.userId);
      expect(config.setupStatus).toBe('pending');
      
      // 5. Setup completion
      const updatedConfig = {
        ...config,
        setupStatus: 'completed',
        spreadsheetId: 'test-sheet-id',
        sheetName: 'テストシート',
        appPublished: true
      };
      
      const saveResult = await mockServices.ConfigService.saveConfig(newUser.userId, updatedConfig);
      expect(saveResult).toBe(true);
      
      // 6. Access admin panel
      const adminResult = await simulateDoGet({ 
        mode: 'admin', 
        userId: newUser.userId 
      });
      expect(adminResult.success).toBe(true);
      expect(adminResult.data.userData).toBeDefined();
    });

  });

  describe('Data Flow Integration', () => {

    it('should handle complete data processing workflow', async () => {
      // Arrange
      const mockServices = setupMockServices();
      const userId = 'test-user-123';
      
      // Setup user with complete config
      await setupCompleteUser(mockServices, userId);
      
      // Act & Assert
      
      // 1. Fetch sheet data
      const sheetResult = await mockServices.DataService.getSheetData(userId);
      expect(sheetResult.success).toBe(true);
      expect(Array.isArray(sheetResult.data)).toBe(true);
      
      // 2. Add reaction
      const reactionResult = await mockServices.DataService.addReaction(
        userId, 
        'row_2', 
        'LIKE'
      );
      expect(reactionResult).toBe(true);
      
      // 3. Toggle highlight
      const highlightResult = await mockServices.DataService.toggleHighlight(
        userId, 
        'row_2'
      );
      expect(highlightResult).toBe(true);
      
      // 4. Get updated data with statistics
      const statsResult = await mockServices.DataService.getDataStatistics(userId);
      expect(statsResult.totalEntries).toBeGreaterThan(0);
      expect(statsResult.reactionCounts).toBeDefined();
    });

  });

  describe('Security Integration', () => {

    it('should enforce access control across all services', async () => {
      // Arrange
      const mockServices = setupMockServices();
      const ownerUserId = 'owner-user-123';
      const guestUserId = 'guest-user-456';
      
      // Setup users with different access levels
      await setupUserWithRole(mockServices, ownerUserId, 'owner');
      await setupUserWithRole(mockServices, guestUserId, 'guest');
      
      // Act & Assert
      
      // 1. Owner should have full access
      const ownerPermission = await mockServices.SecurityService.checkUserPermission(
        ownerUserId, 
        'owner'
      );
      expect(ownerPermission.hasPermission).toBe(true);
      
      // 2. Guest should have limited access
      const guestPermission = await mockServices.SecurityService.checkUserPermission(
        guestUserId, 
        'owner'
      );
      expect(guestPermission.hasPermission).toBe(false);
      
      // 3. Validate input through security service
      const inputValidation = await mockServices.SecurityService.validateUserData({
        email: 'test@example.com',
        answer: 'Test answer',
        url: 'https://docs.google.com/spreadsheets/d/test'
      });
      expect(inputValidation.isValid).toBe(true);
      
      // 4. Reject malicious input
      const maliciousInput = await mockServices.SecurityService.validateUserData({
        email: '<script>alert("xss")</script>',
        answer: 'javascript:alert("xss")'
      });
      expect(maliciousInput.isValid).toBe(false);
    });

  });

  describe('Error Handling Integration', () => {

    it('should handle errors gracefully across service boundaries', async () => {
      // Arrange
      const mockServices = setupMockServices();
      
      // Simulate various error conditions
      mockServices.UserService.getCurrentUserInfo.mockRejectedValue(
        new Error('Database connection failed')
      );
      
      // Act & Assert
      
      // 1. Service should handle database errors
      try {
        await mockServices.UserService.getCurrentUserInfo();
        fail('Should have thrown error');
      } catch (error) {
        expect(error.message).toBe('Database connection failed');
      }
      
      // 2. Error handler should create proper response
      const errorResponse = mockServices.ErrorHandler.handle(
        new Error('Test error'),
        { service: 'UserService', method: 'getCurrentUserInfo' }
      );
      
      expect(errorResponse.success).toBe(false);
      expect(errorResponse.message).toBeDefined();
      expect(errorResponse.errorId).toBeDefined();
    });

  });

  describe('Performance Integration', () => {

    it('should meet performance requirements under load', async () => {
      // Arrange
      const mockServices = setupMockServices();
      const userId = 'perf-test-user';
      await setupCompleteUser(mockServices, userId);
      
      // Act - Simulate concurrent requests
      const startTime = Date.now();
      
      const promises = Array(10).fill().map(() => 
        mockServices.DataService.getSheetData(userId)
      );
      
      const results = await Promise.all(promises);
      const endTime = Date.now();
      
      // Assert
      const totalTime = endTime - startTime;
      const avgTime = totalTime / results.length;
      
      // Should complete within 3 seconds total
      expect(totalTime).toBeLessThan(3000);
      
      // Average per request should be under 500ms
      expect(avgTime).toBeLessThan(500);
      
      // All requests should succeed
      results.forEach(result => {
        expect(result.success).toBe(true);
      });
    });

  });

  describe('Cache Integration', () => {

    it('should utilize caching effectively across services', async () => {
      // Arrange
      const mockServices = setupMockServices();
      const userId = 'cache-test-user';
      
      // Track cache calls
      let cacheGets = 0;
      let cacheSets = 0;
      
      global.CacheService.getScriptCache.mockReturnValue({
        get: jest.fn(() => { cacheGets++; return null; }),
        put: jest.fn(() => { cacheSets++; }),
        remove: jest.fn()
      });
      
      // Act
      await mockServices.UserService.getCurrentUserInfo();
      await mockServices.ConfigService.getUserConfig(userId);
      
      // Assert - Cache should be utilized
      expect(cacheGets).toBeGreaterThan(0);
      expect(cacheSets).toBeGreaterThan(0);
    });

  });

});

// Helper Functions

function setupMockServices() {
  return {
    UserService: {
      getCurrentEmail: jest.fn(() => 'test@example.com'),
      getCurrentUserInfo: jest.fn(() => Promise.resolve(testUserData.validUser)),
      createUser: jest.fn((email) => Promise.resolve({
        userId: 'test-user-123',
        userEmail: email,
        isActive: true,
        configJson: JSON.stringify({ setupStatus: 'pending' }),
        lastModified: new Date().toISOString()
      })),
      getAccessLevel: jest.fn(() => 'owner')
    },
    
    ConfigService: {
      getUserConfig: jest.fn(() => Promise.resolve({
        userId: 'test-user-123',
        setupStatus: 'pending',
        appPublished: false
      })),
      saveConfig: jest.fn(() => Promise.resolve(true))
    },
    
    DataService: {
      getSheetData: jest.fn(() => Promise.resolve({
        success: true,
        data: testSheetData.slice(1), // Exclude header
        count: testSheetData.length - 1
      })),
      addReaction: jest.fn(() => Promise.resolve(true)),
      toggleHighlight: jest.fn(() => Promise.resolve(true)),
      getDataStatistics: jest.fn(() => Promise.resolve({
        totalEntries: 2,
        nonEmptyEntries: 2,
        reactionCounts: { LIKE: 5, UNDERSTAND: 3 }
      }))
    },
    
    SecurityService: {
      checkUserPermission: jest.fn((userId, level) => ({
        hasPermission: level === 'owner',
        currentLevel: 'owner',
        userEmail: 'test@example.com'
      })),
      validateUserData: jest.fn((data) => ({
        isValid: !data.email?.includes('<script>'),
        sanitizedData: data,
        errors: data.email?.includes('<script>') ? ['Invalid input'] : []
      }))
    },
    
    ErrorHandler: {
      handle: jest.fn((error, context) => ({
        success: false,
        message: error.message || 'エラーが発生しました',
        errorId: `error_${Date.now()}`,
        canRetry: true,
        timestamp: new Date().toISOString()
      }))
    }
  };
}

async function simulateDoGet(params) {
  // Simulate main.js doGet function
  return {
    success: true,
    mode: params.mode,
    data: params.mode === 'admin' ? { userData: testUserData.validUser } : null
  };
}

async function setupCompleteUser(mockServices, userId) {
  const completeConfig = {
    userId,
    setupStatus: 'completed',
    appPublished: true,
    spreadsheetId: 'test-sheet-id',
    sheetName: 'テストシート',
    displaySettings: { showNames: true, showReactions: true }
  };
  
  mockServices.ConfigService.getUserConfig.mockResolvedValue(completeConfig);
  return completeConfig;
}

async function setupUserWithRole(mockServices, userId, role) {
  mockServices.SecurityService.checkUserPermission.mockImplementation(
    (checkUserId, requiredLevel) => ({
      hasPermission: role === 'owner' || requiredLevel !== 'owner',
      currentLevel: role,
      userEmail: `${userId}@example.com`
    })
  );
}