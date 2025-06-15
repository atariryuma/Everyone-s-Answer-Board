const { saveScoreSortSetting } = require('../src/Code.gs');

function setup() {
  const props = { setProperty: jest.fn(), deleteProperty: jest.fn() };
  global.PropertiesService = { getScriptProperties: () => props };
  return props;
}

afterEach(() => {
  delete global.PropertiesService;
});

test('saveScoreSortSetting stores boolean state', () => {
  const props = setup();
  saveScoreSortSetting(true);
  expect(props.setProperty).toHaveBeenCalledWith('SCORE_SORT_ENABLED', 'true');
});

test('saveScoreSortSetting handles false', () => {
  const props = setup();
  saveScoreSortSetting(false);
  expect(props.setProperty).toHaveBeenCalledWith('SCORE_SORT_ENABLED', 'false');
});
