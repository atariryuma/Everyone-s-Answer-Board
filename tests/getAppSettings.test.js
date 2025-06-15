const { getAppSettings } = require('../src/Code.gs');

function setup() {
  global.PropertiesService = {
    getScriptProperties: () => ({
      getProperty: (key) => {
        switch (key) {
          case 'IS_PUBLISHED': return 'true';
          case 'ACTIVE_SHEET_NAME': return 'SheetA';
          case 'REACTION_COUNT_ENABLED': return 'false';
          default: return null;
        }
      }
    })
  };
}

afterEach(() => {
  delete global.PropertiesService;
});

test('getAppSettings includes reaction count flag', () => {
  setup();
  const result = getAppSettings();
  expect(result).toEqual({
    isPublished: true,
    activeSheetName: 'SheetA',
    reactionCountEnabled: false
  });
});
