const { getWebAppUrl } = require('../src/Code.gs');

function setup(stored, current, deployId = 'deploy123') {
  const propsObj = {
    getProperty: (key) => {
      if (key === 'WEB_APP_URL') return stored;
      if (key === 'DEPLOY_ID') return deployId;
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

test('getWebAppUrl converts preview domain using deploy id', () => {
  const props = setup('', 'https://foo-1234-script.googleusercontent.com/userCodeAppPanel?x=1', 'AK123');
  expect(getWebAppUrl()).toBe('https://script.google.com/macros/s/AK123/exec?x=1');
  expect(props.setProperties).toHaveBeenCalledWith({ WEB_APP_URL: 'https://script.google.com/macros/s/AK123/exec?x=1' });
});

test('stored preview url is converted when deploy id provided', () => {
  const props = setup('https://foo-1234-script.googleusercontent.com/userCodeAppPanel?y=2', '', 'AK999');
  expect(getWebAppUrl()).toBe('https://script.google.com/macros/s/AK999/exec?y=2');
  expect(props.setProperties).toHaveBeenCalledWith({ WEB_APP_URL: 'https://script.google.com/macros/s/AK999/exec?y=2' });
});
