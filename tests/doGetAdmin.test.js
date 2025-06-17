const { doGetAdmin } = require('../src/Code.gs');

afterEach(() => {
  delete global.PropertiesService;
  delete global.Session;
  delete global.getActiveUserEmail;
  delete global.HtmlService;
});

test('non-admin users see access denied message', () => {
  global.PropertiesService = {
    getScriptProperties: () => ({
      getProperty: () => 'admin@example.com'
    })
  };
  global.Session = { getActiveUser: () => ({ getEmail: () => 'user@example.com' }) };
  global.getActiveUserEmail = () => 'user@example.com';

  const output = {
    content: '',
    setTitle: jest.fn(function() { return this; }),
    getContent() { return this.content; }
  };
  global.HtmlService = {
    createHtmlOutput: jest.fn(html => { output.content = html; return output; })
  };

  const result = doGetAdmin({});
  expect(result.getContent()).toContain('アクセス拒否');
});
