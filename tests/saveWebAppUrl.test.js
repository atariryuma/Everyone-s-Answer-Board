const { saveWebAppUrl } = require('../src/Code.gs');

function setup() {
  const props = { setProperties: jest.fn(), deleteProperty: jest.fn() };
  global.PropertiesService = {
    getScriptProperties: () => props,
    getUserProperties: () => ({
      getProperty: jest.fn(),
      setProperty: jest.fn(),
      setProperties: jest.fn()
    })
  };
  return props;
}

afterEach(() => {
  delete global.PropertiesService;
});

test('saveWebAppUrl stores trimmed url', () => {
  const props = setup();
  saveWebAppUrl('  https://example.com/app  ');
  expect(props.setProperties).toHaveBeenCalledWith({ WEB_APP_URL: 'https://example.com/app' });
});

test('saveWebAppUrl handles empty value', () => {
  const props = setup();
  saveWebAppUrl('');
  expect(props.setProperties).toHaveBeenCalledWith({ WEB_APP_URL: '' });
});

