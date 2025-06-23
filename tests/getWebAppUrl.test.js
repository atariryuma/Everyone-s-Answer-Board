const { getWebAppUrl } = require('../src/Code.gs');

function setup(stored, current) {
  const propsObj = {
    getProperty: (key) => key === 'WEB_APP_URL' ? stored : null,
    setProperties: jest.fn(),
    setProperty: jest.fn()
  };
  global.PropertiesService = {
    getScriptProperties: () => propsObj,
    getUserProperties: () => ({
      getProperty: jest.fn(),
      setProperty: jest.fn(),
      setProperties: jest.fn()
    })
  };
  global.ScriptApp = { getService: () => ({ getUrl: () => current }) };
  return propsObj;
}

afterEach(() => {
  delete global.PropertiesService;
  delete global.ScriptApp;
});

test('getWebAppUrl returns stored url when origins match', () => {
  const props = setup('https://example.com/exec', 'https://example.com/exec');
  expect(getWebAppUrl()).toBe('https://example.com/exec');
  expect(props.setProperties).not.toHaveBeenCalled();
});

test('getWebAppUrl updates url when origin differs', () => {
  const props = setup('https://old.com/exec', 'https://new.com/exec');
  expect(getWebAppUrl()).toBe('https://new.com/exec');
  expect(props.setProperties).toHaveBeenCalledWith({ WEB_APP_URL: 'https://new.com/exec' });
});
