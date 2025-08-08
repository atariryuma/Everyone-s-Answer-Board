const fs = require('fs');
const vm = require('vm');

describe('handleAdminMode fallback verification', () => {
  const mainCode = fs.readFileSync('src/main.gs', 'utf8');
  let context;

  beforeEach(() => {
    context = {
      console,
      debugLog: () => {},
      infoLog: () => {},
      verifyAdminAccess: jest.fn(() => false),
      getCurrentUserEmail: () => 'admin@example.com',
      findUserByEmail: jest.fn(() => ({ userId: 'U1', isActive: true })),
      findUserById: jest.fn(() => ({ userId: 'U1' })),
      renderAdminPanel: jest.fn(() => 'panel'),
      showErrorPage: jest.fn(() => 'error'),
      PropertiesService: {
        getUserProperties: () => ({
          setProperty: jest.fn(),
          getProperty: jest.fn(),
          deleteProperty: jest.fn(),
        }),
        getScriptProperties: () => ({ getProperty: jest.fn(() => '') }),
      },
    };
    vm.createContext(context);
    vm.runInContext(mainCode, context);
    // Overwrite renderAdminPanel to avoid HtmlService dependency
    context.renderAdminPanel = jest.fn(() => 'panel');
  });

  test('grants access when fallback validation succeeds', () => {
    const res = context.handleAdminMode({ userId: 'U1' });
    expect(context.verifyAdminAccess).toHaveBeenCalledWith('U1');
    expect(context.findUserByEmail).toHaveBeenCalledWith('admin@example.com');
    expect(context.renderAdminPanel).toHaveBeenCalled();
    expect(res).toBe('panel');
  });
});

