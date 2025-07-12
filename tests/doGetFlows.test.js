const fs = require('fs');
const vm = require('vm');

describe('doGet flows', () => {
  const code = fs.readFileSync('src/main.gs', 'utf8');
  let context;
  beforeEach(() => {
    context = {
      cacheManager: { get: (_k, fn) => fn() },
      debugLog: () => {},
      HtmlService: { createHtmlOutput: jest.fn(() => 'redirect') },
      Session: { getActiveUser: () => ({ getEmail: () => 'me@example.com' }) }
    };
    vm.createContext(context);
    vm.runInContext(code, context);
    context.renderAnswerBoard = jest.fn(() => 'board');
    context.renderAdminPanel = jest.fn(() => 'admin');
    context.showRegistrationPage = jest.fn(() => 'register');
    context.handleSetupPages = jest.fn(() => null);
    context.validateUserSession = jest.fn(() => ({
      userEmail: 'me@example.com',
      userInfo: { userId: '1', adminEmail: 'me@example.com' }
    }));
    context.parseRequestParams = jest.fn(() => ({ mode: 'admin', isDirectPageAccess: false }));
  });

  test('returns setup page when handleSetupPages outputs', () => {
    context.handleSetupPages.mockReturnValue('setup');
    const result = context.doGet({});
    expect(result).toBe('setup');
  });

  test('renders answer board for direct access', () => {
    context.parseRequestParams.mockReturnValue({ userId: '1', mode: 'view', isDirectPageAccess: true });
    const result = context.doGet({});
    expect(context.renderAnswerBoard).toHaveBeenCalled();
    expect(result).toBe('board');
  });

  test('renders admin panel for admin mode', () => {
    context.parseRequestParams.mockReturnValue({ mode: 'admin', isDirectPageAccess: false });
    const result = context.doGet({});
    expect(context.renderAdminPanel).toHaveBeenCalled();
    expect(result).toBe('admin');
  });

  test('shows registration when userInfo missing', () => {
    context.validateUserSession.mockReturnValue({ userEmail: null, userInfo: null });
    const result = context.doGet({});
    expect(result).toBe('register');
  });
});
