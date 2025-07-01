const { getWebAppUrl } = require('../src/Code.gs');

function setup(stored, current) {
  const propsObj = {
    getProperty: (key) => {
      if (key === 'WEB_APP_URL') return stored;
      return null;
    },
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

test('returns stored url when available', () => {
  const props = setup('https://example.com/exec', 'https://another.com/exec');
  expect(getWebAppUrl()).toBe('https://example.com/exec');
  expect(props.setProperties).not.toHaveBeenCalled();
});

test('falls back to ScriptApp url when stored value missing', () => {
  const props = setup('', 'https://current.com/exec');
  expect(getWebAppUrl()).toBe('https://current.com/exec');
  expect(props.setProperties).not.toHaveBeenCalled();
});
