const { saveDisplayMode } = require('../src/Code.gs');

function setup() {
  const props = { setProperty: jest.fn(), deleteProperty: jest.fn() };
  global.PropertiesService = { getScriptProperties: () => props };
  return props;
}

afterEach(() => {
  delete global.PropertiesService;
});

test('saveDisplayMode stores named mode', () => {
  const props = setup();
  saveDisplayMode('named');
  expect(props.setProperty).toHaveBeenCalledWith('DISPLAY_MODE', 'named');
});

test('saveDisplayMode defaults to anonymous', () => {
  const props = setup();
  saveDisplayMode('other');
  expect(props.setProperty).toHaveBeenCalledWith('DISPLAY_MODE', 'anonymous');
});
