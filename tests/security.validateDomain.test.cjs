const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('path');
const vm = require('vm');

function loadSecurityContext(overrides = {}) {
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (key) => key === 'ADMIN_EMAIL' ? 'admin@example.com' : null
      })
    },
    getCachedProperty: (key) => key === 'ADMIN_EMAIL' ? 'admin@example.com' : null,
    hasCoreSystemProps: () => true,
    openSpreadsheet: () => null,
    getCurrentEmail: () => 'user@example.com',
    ...overrides
  };
  vm.createContext(context);
  const source = fs.readFileSync(path.resolve(__dirname, '../src/SecurityService.js'), 'utf8');
  vm.runInContext(source, context, { filename: 'SecurityService.js' });
  return context;
}

// --- extractDomainSafely ---

test('extractDomainSafely: extracts domain from valid email', () => {
  const ctx = loadSecurityContext();
  assert.equal(ctx.extractDomainSafely('user@example.com'), 'example.com');
});

test('extractDomainSafely: lowercases domain', () => {
  const ctx = loadSecurityContext();
  assert.equal(ctx.extractDomainSafely('User@EXAMPLE.COM'), 'example.com');
});

test('extractDomainSafely: returns null for invalid email', () => {
  const ctx = loadSecurityContext();
  assert.equal(ctx.extractDomainSafely('notanemail'), null);
  assert.equal(ctx.extractDomainSafely(''), null);
  assert.equal(ctx.extractDomainSafely(null), null);
  assert.equal(ctx.extractDomainSafely(undefined), null);
});

// --- validateDomainAccess ---

test('validateDomainAccess: allows same domain', () => {
  const ctx = loadSecurityContext();
  const result = ctx.validateDomainAccess('user@example.com');
  assert.equal(result.allowed, true);
  assert.equal(result.reason, 'domain_match');
});

test('validateDomainAccess: blocks different domain', () => {
  const ctx = loadSecurityContext();
  const result = ctx.validateDomainAccess('user@other.com');
  assert.equal(result.allowed, false);
  assert.equal(result.reason, 'domain_mismatch');
});

test('validateDomainAccess: case insensitive domain match', () => {
  const ctx = loadSecurityContext();
  const result = ctx.validateDomainAccess('User@EXAMPLE.COM');
  assert.equal(result.allowed, true);
});

test('validateDomainAccess: missing email with allowIfEmailMissing=true', () => {
  const ctx = loadSecurityContext();
  const result = ctx.validateDomainAccess(null, { allowIfEmailMissing: true });
  assert.equal(result.allowed, true);
  assert.equal(result.reason, 'missing_email');
});

test('validateDomainAccess: missing email with allowIfEmailMissing=false', () => {
  const ctx = loadSecurityContext();
  const result = ctx.validateDomainAccess(null, { allowIfEmailMissing: false });
  assert.equal(result.allowed, false);
  assert.equal(result.reason, 'missing_email');
});

test('validateDomainAccess: invalid email format', () => {
  const ctx = loadSecurityContext();
  const result = ctx.validateDomainAccess('notanemail');
  assert.equal(result.allowed, false);
  assert.equal(result.reason, 'invalid_email_format');
});

test('validateDomainAccess: admin unconfigured with allowIfAdminUnconfigured=true', () => {
  const ctx = loadSecurityContext({
    getCachedProperty: () => null
  });
  const result = ctx.validateDomainAccess('user@example.com', { allowIfAdminUnconfigured: true });
  assert.equal(result.allowed, true);
  assert.equal(result.reason, 'admin_domain_unconfigured');
});

test('validateDomainAccess: admin unconfigured with allowIfAdminUnconfigured=false', () => {
  const ctx = loadSecurityContext({
    getCachedProperty: () => null
  });
  const result = ctx.validateDomainAccess('user@example.com', { allowIfAdminUnconfigured: false });
  assert.equal(result.allowed, false);
});

test('validateDomainAccess: no success field in return value', () => {
  const ctx = loadSecurityContext();
  const result = ctx.validateDomainAccess('user@example.com');
  assert.equal(result.success, undefined);
});

// --- validateSpreadsheetAccess ---

test('validateSpreadsheetAccess: rejects empty spreadsheetId', () => {
  const ctx = loadSecurityContext();
  const result = ctx.validateSpreadsheetAccess('');
  assert.equal(result.success, false);
  assert.match(result.message, /スプレッドシート/);
  assert.deepEqual([...result.sheets], []);
});

test('validateSpreadsheetAccess: rejects null spreadsheetId', () => {
  const ctx = loadSecurityContext();
  const result = ctx.validateSpreadsheetAccess(null);
  assert.equal(result.success, false);
});

test('validateSpreadsheetAccess: returns error when openSpreadsheet throws', () => {
  const ctx = loadSecurityContext({
    openSpreadsheet: () => { throw new Error('access denied'); }
  });
  const result = ctx.validateSpreadsheetAccess('ss-1');
  assert.equal(result.success, false);
  assert.match(result.message, /アクセスに失敗/);
  assert.equal(result.error, 'access denied');
});

test('validateSpreadsheetAccess: returns error when metadata read fails', () => {
  const ctx = loadSecurityContext({
    openSpreadsheet: () => ({
      spreadsheet: {
        getName: () => { throw new Error('permission denied'); },
        getSheets: () => []
      }
    })
  });
  const result = ctx.validateSpreadsheetAccess('ss-1');
  assert.equal(result.success, false);
  assert.match(result.message, /情報の取得に失敗/);
});

test('validateSpreadsheetAccess: returns sheet list on success', () => {
  const mockSpreadsheet = {
    getName: () => 'My Board',
    getSheets: () => [
      { getName: () => 'Sheet1', getIndex: () => 0 },
      { getName: () => 'Responses', getIndex: () => 1 }
    ]
  };
  const ctx = loadSecurityContext({
    openSpreadsheet: () => ({ spreadsheet: mockSpreadsheet })
  });

  const result = ctx.validateSpreadsheetAccess('ss-1');
  assert.equal(result.success, true);
  assert.equal(result.name, 'My Board');
  assert.equal(result.sheets.length, 2);
  assert.equal(result.sheets[0].name, 'Sheet1');
  assert.equal(result.sheets[1].name, 'Responses');
});

test('validateSpreadsheetAccess: includes executionTime on all response paths', () => {
  const ctx = loadSecurityContext();
  const emptyResult = ctx.validateSpreadsheetAccess('');
  assert.ok(typeof emptyResult.executionTime === 'string');
  assert.match(emptyResult.executionTime, /^\d+ms$/);
});
