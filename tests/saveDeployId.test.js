const { saveDeployId } = require('../src/Code.gs');

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

test('saveDeployId stores trimmed id', () => {
  const props = setup();
  saveDeployId('  abc123  ');
  expect(props.setProperties).toHaveBeenCalledWith({ DEPLOY_ID: 'abc123' });
});

test('saveDeployId handles empty value', () => {
  const props = setup();
  saveDeployId('');
  expect(props.setProperties).toHaveBeenCalledWith({ DEPLOY_ID: '' });
});
