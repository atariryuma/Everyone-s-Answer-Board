const { doGet } = require('../src/Code.gs');

function setup() {
  global.PropertiesService = {
    getScriptProperties: () => ({
      getProperty: (key) => {
        switch (key) {
          case 'IS_PUBLISHED': return 'true';
          case 'ACTIVE_SHEET_NAME': return 'Sheet1';
          case 'ADMIN_EMAILS': return 'admin@example.com';
          default: return null;
        }
      }
    })
  };
  global.Session = { getActiveUser: () => ({ getEmail: () => 'admin@example.com' }) };
  const htmlOutput = {};
  global.HtmlService = {
    createTemplateFromFile: jest.fn(() => ({
      evaluate: jest.fn(() => ({
        setTitle: jest.fn(() => ({ addMetaTag: jest.fn(() => htmlOutput) }))
      }))
    })),
    createHtmlOutput: jest.fn(() => htmlOutput),
    createHtmlOutputFromFile: jest.fn(() => ({ evaluate: jest.fn(() => htmlOutput) }))
  };
  return { htmlOutput };
}

afterEach(() => {
  delete global.PropertiesService;
  delete global.Session;
  delete global.HtmlService;
});

test('doGet respects view parameter for admin groups', () => {
  setup();
  doGet({ parameter: { admin: '1', view: 'groups' } });
  expect(HtmlService.createTemplateFromFile).toHaveBeenCalledWith('OpinionGroups');
});
