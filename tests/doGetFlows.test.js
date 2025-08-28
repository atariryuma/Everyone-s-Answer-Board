const fs = require('fs');
const vm = require('vm');

describe.skip('doGet flows', () => {
  const code = fs.readFileSync('src/main.gs', 'utf8');
  let context;
  beforeEach(() => {
    context = {
      cacheManager: { get: (_k, fn) => fn() },
      debugLog: () => {},
      HtmlService: { createHtmlOutput: jest.fn(() => 'redirect') },
      Session: { getActiveUser: () => ({ getEmail: () => 'me@example.com' }) },
      PropertiesService: { getScriptProperties: () => ({ getProperty: jest.fn(() => 'id') }) },
    };
    vm.createContext(context);
    vm.runInContext(code, context);
    context.renderAnswerBoard = jest.fn(() => 'board');
    context.renderAdminPanel = jest.fn(() => 'admin');
    context.showRegistrationPage = jest.fn(() => 'register');
    context.handleSetupPages = context.handleSetupPages || (() => null);
    context.validateUserSession = jest.fn(() => ({
      userEmail: 'me@example.com',
      userInfo: { userId: '1', adminEmail: 'me@example.com' },
    }));
    context.parseRequestParams = jest.fn(() => ({ mode: 'admin', isDirectPageAccess: false }));
  });

  test('renders answer board for direct access', () => {
    context.parseRequestParams.mockReturnValue({
      userId: '1',
      mode: 'view',
      isDirectPageAccess: true,
    });
    const result = context.doGet({});
    expect(result).toBe('board');
  });

  test('renders admin panel for admin mode', () => {
    context.parseRequestParams.mockReturnValue({ mode: 'admin', isDirectPageAccess: false });
    const result = context.doGet({});
    expect(result).toBe('admin');
  });

  test('shows registration when userInfo missing', () => {
    context.validateUserSession.mockReturnValue({ userEmail: null, userInfo: null });
    const result = context.doGet({});
    expect(result).toBe('register');
  });
});
