jest.mock('fs');
jest.mock('googleapis');

const fs = require('fs');
const { google } = require('googleapis');

test('updateWebAppUrl calls saveWebAppUrl', async () => {
  fs.readFileSync.mockReturnValue('{"scriptId":"123"}');

  const runMock = jest.fn().mockResolvedValue({});
  const listMock = jest.fn().mockResolvedValue({
    data: {
      deployments: [
        { entryPoints: [{ webApp: { url: 'https://example.com' } }] }
      ]
    }
  });

  google.auth = {
    GoogleAuth: jest.fn().mockImplementation(() => ({
      getClient: jest.fn().mockResolvedValue('client')
    }))
  };
  google.script = jest.fn(() => ({
    projects: { deployments: { list: listMock } },
    scripts: { run: runMock }
  }));

  jest.spyOn(console, 'log').mockImplementation(() => {});
  jest.spyOn(console, 'error').mockImplementation(() => {});
  const exitSpy = jest.spyOn(process, 'exit').mockImplementation(() => {});

  require('../scripts/updateWebAppUrl.js');
  await new Promise(setImmediate);

  expect(runMock).toHaveBeenCalledWith({
    scriptId: '123',
    requestBody: { function: 'saveWebAppUrl', parameters: ['https://example.com'] }
  });
  expect(exitSpy).not.toHaveBeenCalled();

  console.log.mockRestore();
  console.error.mockRestore();
  exitSpy.mockRestore();
});

