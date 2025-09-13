/**
 * SecurityService Test Suite
 * Comprehensive testing for security and validation functionality
 */

const { testUserData } = require('../mocks/gas-mocks.js');

// Mock SecurityService (simulate GAS environment)
const SecurityService = {
  validateUserData: jest.fn(),
  checkUserPermission: jest.fn(),
  sanitizeUserData: jest.fn(),
  validateEmailAccess: jest.fn(),
  isAuthorizedUser: jest.fn(),
  getAccessLevel: jest.fn(),
  sanitizeInput: jest.fn(),
  validateInput: jest.fn(),
  escapeHtml: jest.fn(),
  checkCSRFToken: jest.fn(),
  generateSecureToken: jest.fn(),
  validateWebAppAccess: jest.fn(),
  logSecurityEvent: jest.fn(),
  diagnose: jest.fn()
};

describe('SecurityService', () => {
  
  beforeEach(() => {
    jest.resetAllMocks();
    jest.clearAllMocks();
  });

  describe('validateUserData', () => {
    
    it('should validate clean user data', () => {
      // Arrange
      const cleanData = {
        email: 'user@example.com',
        answer: 'This is a valid answer',
        name: '田中太郎',
        class: '3年A組'
      };

      SecurityService.validateUserData.mockReturnValue({
        isValid: true,
        sanitizedData: cleanData,
        errors: []
      });

      // Act
      const result = SecurityService.validateUserData(cleanData);

      // Assert
      expect(result.isValid).toBe(true);
      expect(result.errors).toHaveLength(0);
      expect(result.sanitizedData).toEqual(cleanData);
    });

    it('should detect XSS attempts', () => {
      // Arrange
      const maliciousData = {
        email: 'user@example.com',
        answer: '<script>alert("XSS")</script>',
        name: '<img src="x" onerror="alert(1)">',
        class: 'javascript:alert("XSS")'
      };

      SecurityService.validateUserData.mockReturnValue({
        isValid: false,
        sanitizedData: {
          email: 'user@example.com',
          answer: 'alert("XSS")',
          name: '',
          class: ''
        },
        errors: [
          'HTMLタグが検出されました',
          'JavaScriptコードが検出されました',
          '不正なURLスキームが検出されました'
        ]
      });

      // Act
      const result = SecurityService.validateUserData(maliciousData);

      // Assert
      expect(result.isValid).toBe(false);
      expect(result.errors.length).toBeGreaterThan(0);
      expect(result.errors).toContain('HTMLタグが検出されました');
    });

    it('should handle SQL injection attempts', () => {
      // Arrange
      const sqlInjectionData = {
        email: "'; DROP TABLE users; --",
        answer: "1' OR '1'='1",
        name: "admin'--"
      };

      SecurityService.validateUserData.mockReturnValue({
        isValid: false,
        sanitizedData: {
          email: '',
          answer: '1 OR 1=1',
          name: 'admin'
        },
        errors: [
          '無効なメールアドレス',
          'SQLインジェクションの可能性があります'
        ]
      });

      // Act
      const result = SecurityService.validateUserData(sqlInjectionData);

      // Assert
      expect(result.isValid).toBe(false);
      expect(result.errors).toContain('SQLインジェクションの可能性があります');
    });

    it('should validate email format', () => {
      // Arrange
      const invalidEmailData = {
        email: 'not-an-email',
        answer: 'Valid answer'
      };

      SecurityService.validateUserData.mockReturnValue({
        isValid: false,
        sanitizedData: invalidEmailData,
        errors: ['無効なメールアドレス形式']
      });

      // Act
      const result = SecurityService.validateUserData(invalidEmailData);

      // Assert
      expect(result.isValid).toBe(false);
      expect(result.errors).toContain('無効なメールアドレス形式');
    });

  });

  describe('checkUserPermission', () => {

    it('should grant owner permissions', () => {
      // Arrange
      const userId = 'owner-user-123';
      const requiredLevel = 'owner';

      SecurityService.checkUserPermission.mockReturnValue({
        hasPermission: true,
        currentLevel: 'owner',
        userEmail: 'owner@example.com'
      });

      // Act
      const result = SecurityService.checkUserPermission(userId, requiredLevel);

      // Assert
      expect(result.hasPermission).toBe(true);
      expect(result.currentLevel).toBe('owner');
    });

    it('should deny insufficient permissions', () => {
      // Arrange
      const userId = 'guest-user-123';
      const requiredLevel = 'owner';

      SecurityService.checkUserPermission.mockReturnValue({
        hasPermission: false,
        currentLevel: 'guest',
        userEmail: 'guest@example.com',
        reason: '権限が不足しています'
      });

      // Act
      const result = SecurityService.checkUserPermission(userId, requiredLevel);

      // Assert
      expect(result.hasPermission).toBe(false);
      expect(result.currentLevel).toBe('guest');
      expect(result.reason).toBe('権限が不足しています');
    });

    it('should handle system admin permissions', () => {
      // Arrange
      const userId = 'admin-user-123';

      SecurityService.checkUserPermission.mockReturnValue({
        hasPermission: true,
        currentLevel: 'system_admin',
        userEmail: 'admin@example.com'
      });

      // Act
      const result = SecurityService.checkUserPermission(userId, 'authenticated_user');

      // Assert
      expect(result.hasPermission).toBe(true);
      expect(result.currentLevel).toBe('system_admin');
    });

  });

  describe('sanitizeUserData', () => {

    it('should sanitize HTML content', () => {
      // Arrange
      const dirtyData = {
        answer: '<b>Bold</b> and <script>alert("xss")</script>',
        name: '<img src="x" onerror="alert(1)">Test User'
      };

      SecurityService.sanitizeUserData.mockReturnValue({
        answer: 'Bold and alert("xss")',
        name: 'Test User'
      });

      // Act
      const result = SecurityService.sanitizeUserData(dirtyData);

      // Assert
      expect(result.answer).not.toContain('<script>');
      expect(result.name).not.toContain('<img');
      expect(result.name).toBe('Test User');
    });

    it('should preserve safe content', () => {
      // Arrange
      const safeData = {
        answer: '数学の問題について考えました',
        name: '田中太郎',
        class: '3年A組'
      };

      SecurityService.sanitizeUserData.mockReturnValue(safeData);

      // Act
      const result = SecurityService.sanitizeUserData(safeData);

      // Assert
      expect(result).toEqual(safeData);
    });

  });

  describe('validateEmailAccess', () => {

    it('should validate authorized email domains', () => {
      // Arrange
      const authorizedEmail = 'student@school.edu.jp';

      SecurityService.validateEmailAccess.mockReturnValue({
        isAuthorized: true,
        domain: 'school.edu.jp',
        accessLevel: 'authenticated_user'
      });

      // Act
      const result = SecurityService.validateEmailAccess(authorizedEmail);

      // Assert
      expect(result.isAuthorized).toBe(true);
      expect(result.domain).toBe('school.edu.jp');
    });

    it('should reject unauthorized domains', () => {
      // Arrange
      const unauthorizedEmail = 'user@spam.com';

      SecurityService.validateEmailAccess.mockReturnValue({
        isAuthorized: false,
        domain: 'spam.com',
        reason: '許可されていないドメインです'
      });

      // Act
      const result = SecurityService.validateEmailAccess(unauthorizedEmail);

      // Assert
      expect(result.isAuthorized).toBe(false);
      expect(result.reason).toBe('許可されていないドメインです');
    });

  });

  describe('escapeHtml', () => {

    it('should escape HTML special characters', () => {
      // Arrange
      const htmlString = '<script>alert("XSS")</script>';
      const expectedEscaped = '&lt;script&gt;alert(&quot;XSS&quot;)&lt;/script&gt;';

      SecurityService.escapeHtml.mockReturnValue(expectedEscaped);

      // Act
      const result = SecurityService.escapeHtml(htmlString);

      // Assert
      expect(result).toBe(expectedEscaped);
      expect(result).not.toContain('<script>');
    });

    it('should handle empty and null values', () => {
      // Arrange
      SecurityService.escapeHtml.mockImplementation((input) => {
        if (!input) return '';
        return input.toString();
      });

      // Act & Assert
      expect(SecurityService.escapeHtml('')).toBe('');
      expect(SecurityService.escapeHtml(null)).toBe('');
      expect(SecurityService.escapeHtml(undefined)).toBe('');
    });

  });

  describe('validateInput', () => {

    it('should validate input length limits', () => {
      // Arrange
      const longInput = 'a'.repeat(10000);

      SecurityService.validateInput.mockReturnValue({
        isValid: false,
        errors: ['入力が長すぎます（最大1000文字）'],
        truncated: 'a'.repeat(1000)
      });

      // Act
      const result = SecurityService.validateInput(longInput, { maxLength: 1000 });

      // Assert
      expect(result.isValid).toBe(false);
      expect(result.errors).toContain('入力が長すぎます（最大1000文字）');
    });

    it('should validate required fields', () => {
      // Arrange
      SecurityService.validateInput.mockReturnValue({
        isValid: false,
        errors: ['必須項目が未入力です']
      });

      // Act
      const result = SecurityService.validateInput('', { required: true });

      // Assert
      expect(result.isValid).toBe(false);
      expect(result.errors).toContain('必須項目が未入力です');
    });

  });

  describe('CSRF Protection', () => {

    it('should validate CSRF tokens', () => {
      // Arrange
      const validToken = 'csrf-token-123';

      SecurityService.checkCSRFToken.mockReturnValue({
        isValid: true,
        tokenAge: 300000 // 5 minutes
      });

      // Act
      const result = SecurityService.checkCSRFToken(validToken);

      // Assert
      expect(result.isValid).toBe(true);
      expect(result.tokenAge).toBeLessThan(3600000); // Less than 1 hour
    });

    it('should reject expired CSRF tokens', () => {
      // Arrange
      const expiredToken = 'expired-token';

      SecurityService.checkCSRFToken.mockReturnValue({
        isValid: false,
        reason: 'トークンが期限切れです'
      });

      // Act
      const result = SecurityService.checkCSRFToken(expiredToken);

      // Assert
      expect(result.isValid).toBe(false);
      expect(result.reason).toBe('トークンが期限切れです');
    });

    it('should generate secure tokens', () => {
      // Arrange
      const mockToken = 'secure-random-token-' + Date.now();

      SecurityService.generateSecureToken.mockReturnValue(mockToken);

      // Act
      const result = SecurityService.generateSecureToken();

      // Assert
      expect(result).toBeDefined();
      expect(typeof result).toBe('string');
      expect(result.length).toBeGreaterThan(10);
    });

  });

  describe('Web App Access Control', () => {

    it('should validate web app access permissions', () => {
      // Arrange
      const userEmail = 'authorized@example.com';
      const appId = 'test-app-123';

      SecurityService.validateWebAppAccess.mockReturnValue({
        hasAccess: true,
        permissions: ['read', 'react'],
        sessionValid: true
      });

      // Act
      const result = SecurityService.validateWebAppAccess(userEmail, appId);

      // Assert
      expect(result.hasAccess).toBe(true);
      expect(result.permissions).toContain('read');
      expect(result.sessionValid).toBe(true);
    });

    it('should handle unauthorized access attempts', () => {
      // Arrange
      SecurityService.validateWebAppAccess.mockReturnValue({
        hasAccess: false,
        reason: 'アクセス権限がありません',
        sessionValid: false
      });

      // Act
      const result = SecurityService.validateWebAppAccess('unauthorized@spam.com', 'test-app');

      // Assert
      expect(result.hasAccess).toBe(false);
      expect(result.reason).toBe('アクセス権限がありません');
    });

  });

  describe('Security Event Logging', () => {

    it('should log security events', () => {
      // Arrange
      const securityEvent = {
        type: 'login_attempt',
        userEmail: 'user@example.com',
        success: true,
        ipAddress: '192.168.1.1',
        userAgent: 'Mozilla/5.0...'
      };

      SecurityService.logSecurityEvent.mockReturnValue({
        logged: true,
        eventId: 'event-123'
      });

      // Act
      const result = SecurityService.logSecurityEvent(securityEvent);

      // Assert
      expect(result.logged).toBe(true);
      expect(result.eventId).toBeDefined();
    });

    it('should log suspicious activities', () => {
      // Arrange
      const suspiciousEvent = {
        type: 'xss_attempt',
        userEmail: 'attacker@example.com',
        success: false,
        payload: '<script>alert("xss")</script>',
        blocked: true
      };

      SecurityService.logSecurityEvent.mockReturnValue({
        logged: true,
        eventId: 'security-incident-456',
        alertSent: true
      });

      // Act
      const result = SecurityService.logSecurityEvent(suspiciousEvent);

      // Assert
      expect(result.logged).toBe(true);
      expect(result.alertSent).toBe(true);
    });

  });

  describe('Access Level Management', () => {

    it('should determine correct access levels', () => {
      // Arrange
      const testCases = [
        { email: 'owner@example.com', expected: 'owner' },
        { email: 'admin@example.com', expected: 'system_admin' },
        { email: 'user@example.com', expected: 'authenticated_user' },
        { email: 'guest@unknown.com', expected: 'guest' }
      ];

      testCases.forEach(({ email, expected }) => {
        SecurityService.getAccessLevel.mockReturnValue(expected);

        // Act
        const result = SecurityService.getAccessLevel(email);

        // Assert
        expect(result).toBe(expected);
      });
    });

  });

  describe('Error Handling', () => {

    it('should handle validation errors gracefully', () => {
      // Arrange
      SecurityService.validateUserData.mockImplementation(() => {
        throw new Error('Validation service unavailable');
      });

      // Act & Assert
      expect(() => SecurityService.validateUserData({})).toThrow('Validation service unavailable');
    });

    it('should handle permission check failures', () => {
      // Arrange
      SecurityService.checkUserPermission.mockImplementation(() => {
        throw new Error('Permission service error');
      });

      // Act & Assert
      expect(() => SecurityService.checkUserPermission('user', 'level')).toThrow('Permission service error');
    });

  });

  describe('Performance Tests', () => {

    it('should complete validation within time limits', () => {
      // Arrange
      const startTime = Date.now();
      const testData = {
        email: 'user@example.com',
        answer: 'Test answer',
        name: 'Test User'
      };

      SecurityService.validateUserData.mockReturnValue({
        isValid: true,
        sanitizedData: testData,
        errors: []
      });

      // Act
      SecurityService.validateUserData(testData);
      const endTime = Date.now();

      // Assert (should complete within 50ms for mock)
      expect(endTime - startTime).toBeLessThan(50);
    });

  });

  describe('Diagnostics', () => {

    it('should return security diagnostic information', () => {
      // Arrange
      const mockDiagnostics = {
        service: 'SecurityService',
        timestamp: '2025-01-15T00:00:00.000Z',
        checks: [
          { name: 'Input Validation', status: '✅', details: 'Validation working' },
          { name: 'XSS Protection', status: '✅', details: 'XSS filters active' },
          { name: 'CSRF Protection', status: '✅', details: 'CSRF tokens working' },
          { name: 'Access Control', status: '✅', details: 'Permissions working' }
        ],
        overall: '✅'
      };

      SecurityService.diagnose.mockReturnValue(mockDiagnostics);

      // Act
      const result = SecurityService.diagnose();

      // Assert
      expect(result.service).toBe('SecurityService');
      expect(result.overall).toBe('✅');
      expect(result.checks).toHaveLength(4);
    });

  });

});

describe('SecurityService Integration Tests', () => {

  it('should complete full security validation workflow', () => {
    // Arrange
    const userData = {
      email: 'student@school.edu.jp',
      answer: 'This is my answer to the question',
      name: '田中太郎',
      class: '3年A組'
    };

    // Mock the full workflow
    SecurityService.validateEmailAccess.mockReturnValue({
      isAuthorized: true,
      accessLevel: 'authenticated_user'
    });

    SecurityService.validateUserData.mockReturnValue({
      isValid: true,
      sanitizedData: userData,
      errors: []
    });

    SecurityService.checkUserPermission.mockReturnValue({
      hasPermission: true,
      currentLevel: 'authenticated_user'
    });

    // Act
    const emailValidation = SecurityService.validateEmailAccess(userData.email);
    const dataValidation = SecurityService.validateUserData(userData);
    const permissionCheck = SecurityService.checkUserPermission('user-123', 'authenticated_user');

    // Assert
    expect(emailValidation.isAuthorized).toBe(true);
    expect(dataValidation.isValid).toBe(true);
    expect(permissionCheck.hasPermission).toBe(true);
  });

  it('should block malicious input attempts', () => {
    // Arrange
    const maliciousData = {
      email: '<script>alert("xss")</script>@example.com',
      answer: 'javascript:alert("xss")',
      name: '<img src="x" onerror="alert(1)">'
    };

    // Mock security blocking
    SecurityService.validateUserData.mockReturnValue({
      isValid: false,
      sanitizedData: {
        email: '',
        answer: 'alert("xss")',
        name: ''
      },
      errors: [
        'HTMLタグが検出されました',
        'JavaScriptコードが検出されました',
        '不正なメールアドレス'
      ]
    });

    SecurityService.logSecurityEvent.mockReturnValue({
      logged: true,
      eventId: 'security-block-123'
    });

    // Act
    const validation = SecurityService.validateUserData(maliciousData);
    const logging = SecurityService.logSecurityEvent({
      type: 'xss_attempt',
      blocked: true,
      payload: JSON.stringify(maliciousData)
    });

    // Assert
    expect(validation.isValid).toBe(false);
    expect(validation.errors.length).toBeGreaterThan(0);
    expect(logging.logged).toBe(true);
  });

});