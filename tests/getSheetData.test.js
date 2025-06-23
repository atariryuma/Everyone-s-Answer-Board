let getSheetData;
let COLUMN_HEADERS;

beforeEach(() => {
  jest.resetModules();
  delete global.getConfig;
});

function setupMocks(rows, userEmail, adminEmails = '') {
  const rosterHeaders = ['姓', '名', 'ニックネーム', 'Googleアカウント'];
  const rosterRows = [
    ['A', 'Alice', '', 'a@example.com'],
    ['B', 'Bob', '', 'b@example.com']
  ];

  // Mock SpreadsheetApp
  global.SpreadsheetApp = {
    getActiveSpreadsheet: () => ({
      getSheetByName: (name) => {
        if (name === 'Sheet1') {
          return { getDataRange: () => ({ getValues: () => rows }) };
        }
        if (name === 'roster') {
          return { getDataRange: () => ({ getValues: () => [rosterHeaders, ...rosterRows] }) };
        }
        return null;
      }
    })
  };

  // Mock other global services
  global.Session = { getActiveUser: () => ({ getEmail: () => userEmail }) };
  global.CacheService = { getScriptCache: () => ({ get: () => null, put: () => null }) };
  global.PropertiesService = {
    getScriptProperties: () => ({
      getProperty: (key) => {
        if (key === 'ADMIN_EMAILS') return adminEmails;
        return null;
      }
    }),
    getUserProperties: () => ({
      getProperty: jest.fn(),
      setProperty: jest.fn(),
      setProperties: jest.fn()
    })
  };
}

afterEach(() => {
  delete global.SpreadsheetApp;
  delete global.Session;
  delete global.CacheService;
  delete global.PropertiesService;
  delete global.getConfig;
});

test('getSheetData filters and scores rows', () => {
  ({ getSheetData, COLUMN_HEADERS } = require('../src/Code.gs'));
  const data = [
    [
      COLUMN_HEADERS.EMAIL,
      COLUMN_HEADERS.CLASS,
      COLUMN_HEADERS.OPINION,
      COLUMN_HEADERS.REASON,
      COLUMN_HEADERS.TIMESTAMP,
      COLUMN_HEADERS.UNDERSTAND,
      COLUMN_HEADERS.LIKE,
      COLUMN_HEADERS.CURIOUS,
      COLUMN_HEADERS.HIGHLIGHT
    ],
    ['a@example.com', '1-1', 'Opinion1', 'Reason1', '2024-01-01T00:00:00Z', 'b@example.com', '', '', 'false'],
    ['b@example.com', '1-1', 'Opinion2', 'Reason2', '2024-01-02T00:00:00Z', '', '', '', 'false'],
    ['', '', '', '', '', '', '', '', ''] // ignored
  ];
  setupMocks(data, 'b@example.com', 'b@example.com');

  ({ getSheetData, COLUMN_HEADERS } = require('../src/Code.gs'));

  const result = getSheetData('Sheet1', undefined, 'score');

  expect(result.header).toBe(COLUMN_HEADERS.OPINION);
  expect(result.rows).toHaveLength(2);
  expect(result.rows[0].name).toBe('A Alice');
  expect(result.rows[0].reactions.UNDERSTAND.reacted).toBe(true);
  expect(result.rows[1].name).toBe('B Bob');
  expect(result.rows[1].reactions.UNDERSTAND.reacted).toBe(false);
});

test('getSheetData sorts by newest when specified', () => {
  ({ getSheetData, COLUMN_HEADERS } = require('../src/Code.gs'));
  const data = [
    [
      COLUMN_HEADERS.EMAIL,
      COLUMN_HEADERS.CLASS,
      COLUMN_HEADERS.OPINION,
      COLUMN_HEADERS.REASON,
      COLUMN_HEADERS.TIMESTAMP,
      COLUMN_HEADERS.UNDERSTAND,
      COLUMN_HEADERS.LIKE,
      COLUMN_HEADERS.CURIOUS,
      COLUMN_HEADERS.HIGHLIGHT
    ],
    ['first@example.com', '1-1', 'Old', 'A', '2024-01-01T00:00:00Z', '', '', '', 'false'],
    ['second@example.com', '1-1', 'New', 'B', '2024-02-01T00:00:00Z', '', '', '', 'false']
  ];
  setupMocks(data, '', '');

  ({ getSheetData, COLUMN_HEADERS } = require('../src/Code.gs'));

  const result = getSheetData('Sheet1', undefined, 'newest');

  expect(result.rows.map(r => r.rowIndex)).toEqual([3, 2]);
});

test('getSheetData supports random sort', () => {
  ({ getSheetData, COLUMN_HEADERS } = require('../src/Code.gs'));
  const data = [
    [
      COLUMN_HEADERS.EMAIL,
      COLUMN_HEADERS.CLASS,
      COLUMN_HEADERS.OPINION,
      COLUMN_HEADERS.REASON,
      COLUMN_HEADERS.TIMESTAMP,
      COLUMN_HEADERS.UNDERSTAND,
      COLUMN_HEADERS.LIKE,
      COLUMN_HEADERS.CURIOUS,
      COLUMN_HEADERS.HIGHLIGHT
    ],
    ['first@example.com', '1-1', 'Old', 'A', '2024-03-01T00:00:00Z', '', '', '', 'false'],
    ['second@example.com', '1-1', 'New', 'B', '2024-04-01T00:00:00Z', '', '', '', 'false']
  ];
  setupMocks(data, '', '');

  ({ getSheetData, COLUMN_HEADERS } = require('../src/Code.gs'));

  const result = getSheetData('Sheet1', undefined, 'random');

  const indexes = result.rows.map(r => r.rowIndex).sort();
  expect(indexes).toEqual([2, 3]);
});

test('getSheetData omits names for non-admin', () => {
  ({ getSheetData, COLUMN_HEADERS } = require('../src/Code.gs'));
  const data = [
    [
      COLUMN_HEADERS.EMAIL,
      COLUMN_HEADERS.CLASS,
      COLUMN_HEADERS.OPINION,
      COLUMN_HEADERS.REASON,
      COLUMN_HEADERS.TIMESTAMP,
      COLUMN_HEADERS.UNDERSTAND,
      COLUMN_HEADERS.LIKE,
      COLUMN_HEADERS.CURIOUS,
      COLUMN_HEADERS.HIGHLIGHT
    ],
    ['a@example.com', '1-1', 'Opinion1', 'Reason1', '2024-05-01T00:00:00Z', '', '', '', 'false']
  ];
  setupMocks(data, 'a@example.com', '');

  ({ getSheetData, COLUMN_HEADERS } = require('../src/Code.gs'));

  const result = getSheetData('Sheet1');

  expect(result.rows[0].name).toBe('');
});

test('getSheetData returns names for admin users', () => {
  ({ getSheetData, COLUMN_HEADERS } = require('../src/Code.gs'));
  const data = [
    [
      COLUMN_HEADERS.EMAIL,
      COLUMN_HEADERS.CLASS,
      COLUMN_HEADERS.OPINION,
      COLUMN_HEADERS.REASON,
      COLUMN_HEADERS.TIMESTAMP,
      COLUMN_HEADERS.UNDERSTAND,
      COLUMN_HEADERS.LIKE,
      COLUMN_HEADERS.CURIOUS,
      COLUMN_HEADERS.HIGHLIGHT
    ],
    ['a@example.com', '1-1', 'Opinion1', 'Reason1', '2024-06-01T00:00:00Z', '', '', '', 'false']
  ];
  setupMocks(data, 'a@example.com', 'a@example.com');

  ({ getSheetData, COLUMN_HEADERS } = require('../src/Code.gs'));

  const result = getSheetData('Sheet1');

  expect(result.rows[0].name).toBe('A Alice');
});

test('reaction lists trim spaces and ignore empties', () => {
  ({ getSheetData, COLUMN_HEADERS } = require('../src/Code.gs'));
  const data = [
    [
      COLUMN_HEADERS.EMAIL,
      COLUMN_HEADERS.CLASS,
      COLUMN_HEADERS.OPINION,
      COLUMN_HEADERS.REASON,
      COLUMN_HEADERS.TIMESTAMP,
      COLUMN_HEADERS.UNDERSTAND,
      COLUMN_HEADERS.LIKE,
      COLUMN_HEADERS.CURIOUS,
      COLUMN_HEADERS.HIGHLIGHT
    ],
    [
      'a@example.com',
      '1-1',
      'Opinion',
      'Reason',
      '2024-07-01T00:00:00Z',
      ' a@example.com , b@example.com ',
      ' a@example.com , ',
      '   ',
      'false'
    ]
  ];
  setupMocks(data, 'a@example.com', 'a@example.com');

  ({ getSheetData, COLUMN_HEADERS } = require('../src/Code.gs'));

  const result = getSheetData('Sheet1');

  const row = result.rows[0];
  expect(row.reactions.UNDERSTAND.count).toBe(2);
  expect(row.reactions.UNDERSTAND.reacted).toBe(true);
  expect(row.reactions.LIKE.count).toBe(1);
  expect(row.reactions.LIKE.reacted).toBe(true);
  expect(row.reactions.CURIOUS.count).toBe(0);
});

test('getSheetData supports custom headers from config', () => {
  global.getConfig = () => ({
    questionHeader: 'Question',
    answerHeader: 'Ans',
    reasonHeader: 'Why',
    nameMode: '同一シート',
    nameHeader: 'Name',
    classHeader: 'Class'
  });
  ({ getSheetData, COLUMN_HEADERS } = require('../src/Code.gs'));
  const data = [
    [
      COLUMN_HEADERS.EMAIL,
      'Class',
      'Ans',
      'Why',
      COLUMN_HEADERS.TIMESTAMP,
      COLUMN_HEADERS.UNDERSTAND,
      COLUMN_HEADERS.LIKE,
      COLUMN_HEADERS.CURIOUS,
      COLUMN_HEADERS.HIGHLIGHT,
      'Name'
    ],
    ['a@example.com', '1-1', 'Answer1', 'Because', '2024-08-01T00:00:00Z', '', '', '', 'false', 'Alice']
  ];
  setupMocks(data, 'a@example.com', 'a@example.com');

  const result = getSheetData('Sheet1');

  expect(result.header).toBe('Question');
  expect(result.rows[0].opinion).toBe('Answer1');
  expect(result.rows[0].name).toBe('Alice');
});
