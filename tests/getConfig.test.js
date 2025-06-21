const { getConfig } = require('../src/config.gs');

function setup(values) {
  const sheet = {
    getDataRange: () => ({ getValues: () => values })
  };
  global.SpreadsheetApp = {
    getActiveSpreadsheet: () => ({
      getSheetByName: (name) => name === 'Config' ? sheet : null
    })
  };
}

afterEach(() => {
  delete global.SpreadsheetApp;
});

test('getConfig returns mapping for sheet', () => {
  const data = [
    ['表示シート名','Q','A','R','同一シート','氏名','クラス'],
    ['Sheet1','Q1','A1','R1','同一シート','名前','クラス']
  ];
  setup(data);
  const cfg = getConfig('Sheet1');
  expect(cfg).toEqual({
    questionHeader:'Q1',
    answerHeader:'A1',
    reasonHeader:'R1',
    nameMode:'同一シート',
    nameHeader:'名前',
    classHeader:'クラス'
  });
});

test('getConfig throws for missing sheet', () => {
  const data = [
    ['表示シート名','Q','A','R','同一シート','氏名','クラス']
  ];
  setup(data);
  expect(() => getConfig('Other')).toThrow('Config シートに設定がありません');
});
