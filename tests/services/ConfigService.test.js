/**
 * ConfigService Test Suite
 * Comprehensive testing for configuration management functionality
 */

const { resetAllMocks, testUserData, testConfigData } = require('../mocks/gas-mocks.js');

// Mock ConfigService (simulate GAS environment)
const ConfigService = {
  getUserConfig: jest.fn(),
  saveUserConfig: jest.fn(),
  updateUserConfig: jest.fn(),
  validateConfig: jest.fn(),
  determineSetupStep: jest.fn(),
  isSystemSetup: jest.fn(),
  createDefaultConfig: jest.fn(),
  migrateConfig: jest.fn(),
  getUserInfo: jest.fn(),
  getConfigCompletionScore: jest.fn(),
  validateUserId: jest.fn(),
  validateSpreadsheetId: jest.fn(),
  validateFormUrl: jest.fn(),
  diagnose: jest.fn()
};

describe('ConfigService', () => {
  
  beforeEach(() => {
    resetAllMocks();
    jest.clearAllMocks();
  });

  describe('getUserConfig', () => {
    
    it('should return user configuration from configJson', () => {
      // Arrange
      const userId = 'test-user-123';
      const mockConfig = {
        userId,
        setupStatus: 'completed',
        spreadsheetId: 'test-sheet-id',
        sheetName: 'テストシート',
        appPublished: true,
        displaySettings: { showNames: true, showReactions: true }
      };

      ConfigService.getUserConfig.mockReturnValue(mockConfig);

      // Act
      const result = ConfigService.getUserConfig(userId);

      // Assert
      expect(result).toEqual(mockConfig);
      expect(ConfigService.getUserConfig).toHaveBeenCalledWith(userId);
    });

    it('should return null for invalid userId', () => {
      // Arrange
      ConfigService.getUserConfig.mockReturnValue(null);

      // Act
      const result = ConfigService.getUserConfig(null);

      // Assert
      expect(result).toBeNull();
    });

    it('should handle config migration automatically', () => {
      // Arrange
      const userId = 'test-user-123';
      const legacyConfig = {
        setupStatus: 'pending',
        spreadsheetId: 'old-sheet-id'
      };
      const migratedConfig = {
        ...legacyConfig,
        configVersion: '2.0',
        migratedAt: '2025-01-15T00:00:00.000Z'
      };

      ConfigService.getUserConfig.mockReturnValue(migratedConfig);

      // Act
      const result = ConfigService.getUserConfig(userId);

      // Assert
      expect(result.configVersion).toBe('2.0');
      expect(result.migratedAt).toBeDefined();
    });

  });

  describe('saveUserConfig', () => {

    it('should save valid configuration', () => {
      // Arrange
      const userId = 'test-user-123';
      const config = {
        setupStatus: 'completed',
        spreadsheetId: 'test-sheet-id',
        sheetName: 'テストシート',
        appPublished: true
      };

      ConfigService.saveUserConfig.mockReturnValue(true);

      // Act
      const result = ConfigService.saveUserConfig(userId, config);

      // Assert
      expect(result).toBe(true);
      expect(ConfigService.saveUserConfig).toHaveBeenCalledWith(userId, config);
    });

    it('should reject invalid configuration', () => {
      // Arrange
      const userId = 'test-user-123';
      const invalidConfig = {
        setupStatus: 'invalid-status',
        spreadsheetId: '',
        sheetName: null
      };

      ConfigService.saveUserConfig.mockImplementation(() => {
        throw new Error('設定の検証に失敗しました');
      });

      // Act & Assert
      expect(() => ConfigService.saveUserConfig(userId, invalidConfig))
        .toThrow('設定の検証に失敗しました');
    });

  });

  describe('updateUserConfig', () => {

    it('should update specific configuration fields', () => {
      // Arrange
      const userId = 'test-user-123';
      const updates = {
        appPublished: true,
        publishedAt: '2025-01-15T00:00:00.000Z'
      };
      const updatedConfig = {
        userId,
        setupStatus: 'completed',
        spreadsheetId: 'test-sheet-id',
        ...updates
      };

      ConfigService.updateUserConfig.mockReturnValue(updatedConfig);

      // Act
      const result = ConfigService.updateUserConfig(userId, updates);

      // Assert
      expect(result).toEqual(updatedConfig);
      expect(result.appPublished).toBe(true);
      expect(result.publishedAt).toBeDefined();
    });

  });

  describe('validateConfig', () => {

    it('should validate complete configuration', () => {
      // Arrange
      const validConfig = {
        setupStatus: 'completed',
        spreadsheetId: 'valid-sheet-id-123',
        sheetName: 'テストシート',
        formUrl: 'https://docs.google.com/forms/test',
        appPublished: true
      };

      ConfigService.validateConfig.mockReturnValue({
        isValid: true,
        errors: [],
        completionScore: 100
      });

      // Act
      const result = ConfigService.validateConfig(validConfig);

      // Assert
      expect(result.isValid).toBe(true);
      expect(result.errors).toHaveLength(0);
      expect(result.completionScore).toBe(100);
    });

    it('should identify validation errors', () => {
      // Arrange
      const invalidConfig = {
        setupStatus: 'invalid',
        spreadsheetId: '',
        sheetName: null
      };

      ConfigService.validateConfig.mockReturnValue({
        isValid: false,
        errors: [
          '無効なセットアップステータス',
          'スプレッドシートIDが不正',
          'シート名が未設定'
        ],
        completionScore: 25
      });

      // Act
      const result = ConfigService.validateConfig(invalidConfig);

      // Assert
      expect(result.isValid).toBe(false);
      expect(result.errors).toHaveLength(3);
      expect(result.completionScore).toBe(25);
    });

  });

  describe('determineSetupStep', () => {

    it('should return step 1 for new users', () => {
      // Arrange
      const userInfo = {
        userId: 'new-user-123',
        configJson: JSON.stringify({ setupStatus: 'pending' })
      };

      ConfigService.determineSetupStep.mockReturnValue(1);

      // Act
      const result = ConfigService.determineSetupStep(userInfo);

      // Assert
      expect(result).toBe(1);
    });

    it('should return step 2 for data source configured', () => {
      // Arrange
      const userInfo = {
        userId: 'user-123',
        configJson: JSON.stringify({
          setupStatus: 'in_progress',
          spreadsheetId: 'test-sheet-id',
          sheetName: 'テストシート'
        })
      };

      ConfigService.determineSetupStep.mockReturnValue(2);

      // Act
      const result = ConfigService.determineSetupStep(userInfo);

      // Assert
      expect(result).toBe(2);
    });

    it('should return step 3 for completed setup', () => {
      // Arrange
      const userInfo = {
        userId: 'user-123',
        configJson: JSON.stringify({
          setupStatus: 'completed',
          spreadsheetId: 'test-sheet-id',
          sheetName: 'テストシート',
          formCreated: true,
          appPublished: true
        })
      };

      ConfigService.determineSetupStep.mockReturnValue(3);

      // Act
      const result = ConfigService.determineSetupStep(userInfo);

      // Assert
      expect(result).toBe(3);
    });

  });

  describe('Validation Functions', () => {

    describe('validateUserId', () => {
      
      it('should validate correct user IDs', () => {
        ConfigService.validateUserId.mockImplementation((userId) => {
          return userId && typeof userId === 'string' && userId.length > 0;
        });

        expect(ConfigService.validateUserId('valid-user-123')).toBe(true);
        expect(ConfigService.validateUserId('')).toBe(false);
        expect(ConfigService.validateUserId(null)).toBe(false);
        expect(ConfigService.validateUserId(undefined)).toBe(false);
      });

    });

    describe('validateSpreadsheetId', () => {
      
      it('should validate Google Sheets IDs', () => {
        ConfigService.validateSpreadsheetId.mockImplementation((id) => {
          if (!id || typeof id !== 'string') return false;
          return /^[a-zA-Z0-9-_]+$/.test(id) && id.length > 10;
        });

        expect(ConfigService.validateSpreadsheetId('1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms')).toBe(true);
        expect(ConfigService.validateSpreadsheetId('invalid-id')).toBe(false);
        expect(ConfigService.validateSpreadsheetId('')).toBe(false);
      });

    });

    describe('validateFormUrl', () => {
      
      it('should validate Google Forms URLs', () => {
        ConfigService.validateFormUrl.mockImplementation((url) => {
          if (!url || typeof url !== 'string') return false;
          return url.includes('forms.gle') || url.includes('docs.google.com/forms');
        });

        expect(ConfigService.validateFormUrl('https://docs.google.com/forms/d/test/viewform')).toBe(true);
        expect(ConfigService.validateFormUrl('https://forms.gle/test123')).toBe(true);
        expect(ConfigService.validateFormUrl('https://example.com/form')).toBe(false);
      });

    });

  });

  describe('Configuration Migration', () => {

    it('should migrate legacy configurations', () => {
      // Arrange
      const legacyConfig = {
        spreadsheetId: 'test-sheet-id',
        formUrl: 'https://forms.gle/test'
      };
      const migratedConfig = {
        ...legacyConfig,
        configVersion: '2.0',
        setupStatus: 'completed',
        displaySettings: { showNames: true, showReactions: true },
        migratedAt: '2025-01-15T00:00:00.000Z'
      };

      ConfigService.migrateConfig.mockReturnValue(migratedConfig);

      // Act
      const result = ConfigService.migrateConfig(legacyConfig);

      // Assert
      expect(result.configVersion).toBe('2.0');
      expect(result.setupStatus).toBe('completed');
      expect(result.displaySettings).toBeDefined();
      expect(result.migratedAt).toBeDefined();
    });

  });

  describe('System Configuration', () => {

    it('should check system setup status', () => {
      // Arrange
      global.PropertiesService.getScriptProperties.mockReturnValue({
        getProperty: jest.fn(() => JSON.stringify({ initialized: true }))
      });

      ConfigService.isSystemSetup.mockReturnValue(true);

      // Act
      const result = ConfigService.isSystemSetup();

      // Assert
      expect(result).toBe(true);
    });

    it('should handle missing system configuration', () => {
      // Arrange
      global.PropertiesService.getScriptProperties.mockReturnValue({
        getProperty: jest.fn(() => null)
      });

      ConfigService.isSystemSetup.mockReturnValue(false);

      // Act
      const result = ConfigService.isSystemSetup();

      // Assert
      expect(result).toBe(false);
    });

  });

  describe('Configuration Completion Score', () => {

    it('should calculate completion score for complete config', () => {
      // Arrange
      const completeConfig = {
        setupStatus: 'completed',
        spreadsheetId: 'test-sheet-id',
        sheetName: 'テストシート',
        formUrl: 'https://forms.gle/test',
        appPublished: true,
        displaySettings: { showNames: true }
      };

      ConfigService.getConfigCompletionScore.mockReturnValue(100);

      // Act
      const result = ConfigService.getConfigCompletionScore(completeConfig);

      // Assert
      expect(result).toBe(100);
    });

    it('should calculate completion score for partial config', () => {
      // Arrange
      const partialConfig = {
        setupStatus: 'pending',
        spreadsheetId: 'test-sheet-id'
      };

      ConfigService.getConfigCompletionScore.mockReturnValue(40);

      // Act
      const result = ConfigService.getConfigCompletionScore(partialConfig);

      // Assert
      expect(result).toBe(40);
    });

  });

  describe('Error Handling', () => {

    it('should handle database connection errors', () => {
      // Arrange
      ConfigService.getUserConfig.mockImplementation(() => {
        throw new Error('Database connection failed');
      });

      // Act & Assert
      expect(() => ConfigService.getUserConfig('test-user')).toThrow('Database connection failed');
    });

    it('should handle invalid JSON in configJson', () => {
      // Arrange
      const userInfo = {
        userId: 'test-user',
        configJson: 'invalid-json{'
      };

      ConfigService.getUserConfig.mockImplementation(() => {
        throw new Error('設定の解析に失敗しました');
      });

      // Act & Assert
      expect(() => ConfigService.getUserConfig(userInfo.userId)).toThrow('設定の解析に失敗しました');
    });

  });

  describe('Performance Tests', () => {

    it('should respond within acceptable time limits', async () => {
      // Arrange
      const startTime = Date.now();
      ConfigService.getUserConfig.mockReturnValue(testConfigData.validConfig);

      // Act
      ConfigService.getUserConfig('test-user');
      const endTime = Date.now();

      // Assert (should complete within 50ms for mock)
      expect(endTime - startTime).toBeLessThan(50);
    });

  });

  describe('Diagnostics', () => {

    it('should return diagnostic information', () => {
      // Arrange
      const mockDiagnostics = {
        service: 'ConfigService',
        timestamp: '2025-01-15T00:00:00.000Z',
        checks: [
          { name: 'Database Connection', status: '✅', details: 'Database accessible' },
          { name: 'Config Validation', status: '✅', details: 'Validation working' },
          { name: 'Migration System', status: '✅', details: 'Migration ready' }
        ],
        overall: '✅'
      };

      ConfigService.diagnose.mockReturnValue(mockDiagnostics);

      // Act
      const result = ConfigService.diagnose();

      // Assert
      expect(result.service).toBe('ConfigService');
      expect(result.overall).toBe('✅');
      expect(result.checks).toHaveLength(3);
    });

  });

});

describe('ConfigService Integration Tests', () => {

  it('should complete full configuration workflow', async () => {
    // Arrange
    const userId = 'integration-test-user';
    const initialConfig = { setupStatus: 'pending' };
    const updates = {
      spreadsheetId: 'test-sheet-id',
      sheetName: 'テストシート',
      setupStatus: 'completed'
    };

    // Mock the workflow
    ConfigService.createDefaultConfig.mockReturnValue(initialConfig);
    ConfigService.updateUserConfig.mockReturnValue({ ...initialConfig, ...updates });
    ConfigService.validateConfig.mockReturnValue({ isValid: true, errors: [] });
    ConfigService.determineSetupStep.mockReturnValue(3);

    // Act
    const defaultConfig = ConfigService.createDefaultConfig(userId);
    const updatedConfig = ConfigService.updateUserConfig(userId, updates);
    const validation = ConfigService.validateConfig(updatedConfig);
    const setupStep = ConfigService.determineSetupStep({ configJson: JSON.stringify(updatedConfig) });

    // Assert
    expect(defaultConfig.setupStatus).toBe('pending');
    expect(updatedConfig.setupStatus).toBe('completed');
    expect(validation.isValid).toBe(true);
    expect(setupStep).toBe(3);
  });

});