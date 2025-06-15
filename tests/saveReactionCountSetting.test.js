const { saveReactionCountSetting } = require('../src/Code.gs');

function setup() {
  const props = { setProperty: jest.fn(), deleteProperty: jest.fn() };
  global.PropertiesService = { getScriptProperties: () => props };
  return props;
}

afterEach(() => {
  delete global.PropertiesService;
});

test('saveReactionCountSetting stores boolean state', () => {
  const props = setup();
  saveReactionCountSetting(true);
  expect(props.setProperty).toHaveBeenCalledWith('REACTION_COUNT_ENABLED', 'true');
});

test('saveReactionCountSetting handles false', () => {
  const props = setup();
  saveReactionCountSetting(false);
  expect(props.setProperty).toHaveBeenCalledWith('REACTION_COUNT_ENABLED', 'false');
});
