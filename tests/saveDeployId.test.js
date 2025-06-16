const { saveDeployId } = require('../src/Code.gs');

function setup() {
  const props = { setProperty: jest.fn(), deleteProperty: jest.fn() };
  global.PropertiesService = { getScriptProperties: () => props };
  return props;
}

afterEach(() => {
  delete global.PropertiesService;
});

test('saveDeployId stores trimmed id', () => {
  const props = setup();
  saveDeployId('  abc123  ');
  expect(props.setProperty).toHaveBeenCalledWith('DEPLOY_ID', 'abc123');
});

test('saveDeployId handles empty value', () => {
  const props = setup();
  saveDeployId('');
  expect(props.setProperty).toHaveBeenCalledWith('DEPLOY_ID', '');
});
