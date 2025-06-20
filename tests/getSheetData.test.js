const { getSheetData, COLUMN_HEADERS } = require('../src/Code.gs');

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
        if (name === 'sheet 1') {
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
    })
  };
}

afterEach(() => {
  delete global.SpreadsheetApp;
  delete global.Session;
  delete global.CacheService;
  delete global.PropertiesService;
});

test('getSheetData filters and scores rows', () => {
  const data = [
    [
      COLUMN_HEADERS.EMAIL,
      COLUMN_HEADERS.CLASS,
      COLUMN_HEADERS.OPINION,
      COLUMN_HEADERS.REASON,
      COLUMN_HEADERS.UNDERSTAND,
      COLUMN_HEADERS.LIKE,
      COLUMN_HEADERS.CURIOUS,
      COLUMN_HEADERS.HIGHLIGHT
    ],
    ['a@example.com', '1-1', 'Opinion1', 'Reason1', 'b@example.com', '', '', 'false'],
    ['b@example.com', '1-1', 'Opinion2', 'Reason2', '', '', '', 'false'],
    ['', '', '', '', '', '', '', ''] // ignored
  ];
  setupMocks(data, 'b@example.com', 'b@example.com');

  const result = getSheetData('Sheet1', undefined, 'score');

  expect(result.header).toBe(COLUMN_HEADERS.OPINION);
  expect(result.rows).toHaveLength(2);
  expect(result.rows[0].name).toBe('A Alice');
  expect(result.rows[0].reactions.UNDERSTAND.reacted).toBe(true);
  expect(result.rows[1].name).toBe('B Bob');
  expect(result.rows[1].reactions.UNDERSTAND.reacted).toBe(false);
});

test('getSheetData sorts by newest when specified', () => {
  const data = [
    [
      COLUMN_HEADERS.EMAIL,
      COLUMN_HEADERS.CLASS,
      COLUMN_HEADERS.OPINION,
      COLUMN_HEADERS.REASON,
      COLUMN_HEADERS.UNDERSTAND,
      COLUMN_HEADERS.LIKE,
      COLUMN_HEADERS.CURIOUS,
      COLUMN_HEADERS.HIGHLIGHT
    ],
    ['first@example.com', '1-1', 'Old', 'A', '', '', '', 'false'],
    ['second@example.com', '1-1', 'New', 'B', '', '', '', 'false']
  ];
  setupMocks(data, '', '');

  const result = getSheetData('Sheet1', undefined, 'newest');

  expect(result.rows.map(r => r.rowIndex)).toEqual([3, 2]);
});

test('getSheetData supports random sort', () => {
  const data = [
    [
      COLUMN_HEADERS.EMAIL,
      COLUMN_HEADERS.CLASS,
      COLUMN_HEADERS.OPINION,
      COLUMN_HEADERS.REASON,
      COLUMN_HEADERS.UNDERSTAND,
      COLUMN_HEADERS.LIKE,
      COLUMN_HEADERS.CURIOUS,
      COLUMN_HEADERS.HIGHLIGHT
    ],
    ['first@example.com', '1-1', 'Old', 'A', '', '', '', 'false'],
    ['second@example.com', '1-1', 'New', 'B', '', '', '', 'false']
  ];
  setupMocks(data, '', '');

  const result = getSheetData('Sheet1', undefined, 'random');

  const indexes = result.rows.map(r => r.rowIndex).sort();
  expect(indexes).toEqual([2, 3]);
});

test('getSheetData omits names for non-admin', () => {
  const data = [
    [
      COLUMN_HEADERS.EMAIL,
      COLUMN_HEADERS.CLASS,
      COLUMN_HEADERS.OPINION,
      COLUMN_HEADERS.REASON,
      COLUMN_HEADERS.UNDERSTAND,
      COLUMN_HEADERS.LIKE,
      COLUMN_HEADERS.CURIOUS,
      COLUMN_HEADERS.HIGHLIGHT
    ],
    ['a@example.com', '1-1', 'Opinion1', 'Reason1', '', '', '', 'false']
  ];
  setupMocks(data, 'a@example.com', '');

  const result = getSheetData('Sheet1');

  expect(result.rows[0].name).toBe('');
});

test('getSheetData returns names for admin users', () => {
  const data = [
    [
      COLUMN_HEADERS.EMAIL,
      COLUMN_HEADERS.CLASS,
      COLUMN_HEADERS.OPINION,
      COLUMN_HEADERS.REASON,
      COLUMN_HEADERS.UNDERSTAND,
      COLUMN_HEADERS.LIKE,
      COLUMN_HEADERS.CURIOUS,
      COLUMN_HEADERS.HIGHLIGHT
    ],
    ['a@example.com', '1-1', 'Opinion1', 'Reason1', '', '', '', 'false']
  ];
  setupMocks(data, 'a@example.com', 'a@example.com');

  const result = getSheetData('Sheet1');

  expect(result.rows[0].name).toBe('A Alice');
});

test('reaction lists trim spaces and ignore empties', () => {
  const data = [
    [
      COLUMN_HEADERS.EMAIL,
      COLUMN_HEADERS.CLASS,
      COLUMN_HEADERS.OPINION,
      COLUMN_HEADERS.REASON,
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
      ' a@example.com , b@example.com ',
      ' a@example.com , ',
      '   ',
      'false'
    ]
  ];
  setupMocks(data, 'a@example.com', 'a@example.com');

  const result = getSheetData('Sheet1');

  const row = result.rows[0];
  expect(row.reactions.UNDERSTAND.count).toBe(2);
  expect(row.reactions.UNDERSTAND.reacted).toBe(true);
  expect(row.reactions.LIKE.count).toBe(1);
  expect(row.reactions.LIKE.reacted).toBe(true);
  expect(row.reactions.CURIOUS.count).toBe(0);
});
