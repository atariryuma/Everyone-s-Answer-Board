const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function loadValidatorsContext(overrides = {}) {
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    URL,
    // SYSTEM_LIMITS は SystemController.js でも定義されるが、validators.js は
    // 別 vm context に単独でロードされるため、ここでミラーする必要がある。
    SYSTEM_LIMITS: {
      MAX_LOCK_ROWS: 100, PREVIEW_LENGTH: 50,
      DEFAULT_PAGE_SIZE: 20, MAX_PAGE_SIZE: 100,
      RADIX_DECIMAL: 10
    },
    getColumnAnalysis: () => ({ success: true }),
    getFormInfo: () => ({ success: true }),
    ...overrides
  };
  vm.createContext(context);
  const source = fs.readFileSync(path.resolve(__dirname, '../src/validators.js'), 'utf8');
  vm.runInContext(source, context, { filename: 'validators.js' });
  return context;
}

// =====================================================================
// validateEmail
// =====================================================================

test('validateEmail: accepts standard email and lowercases + trims', () => {
  const ctx = loadValidatorsContext();
  const r = ctx.validateEmail('  User@Example.COM  ');
  assert.equal(r.isValid, true);
  assert.equal(r.sanitized, 'user@example.com');
});

test('validateEmail: rejects missing or non-string input', () => {
  const ctx = loadValidatorsContext();
  assert.equal(ctx.validateEmail('').isValid, false);
  assert.equal(ctx.validateEmail(null).isValid, false);
  assert.equal(ctx.validateEmail(42).isValid, false);
  assert.equal(ctx.validateEmail({}).isValid, false);
});

test('validateEmail: rejects malformed email strings', () => {
  const ctx = loadValidatorsContext();
  assert.equal(ctx.validateEmail('notanemail').isValid, false);
  assert.equal(ctx.validateEmail('user@').isValid, false);
  assert.equal(ctx.validateEmail('@example.com').isValid, false);
  assert.equal(ctx.validateEmail('user @example.com').isValid, false);
});

// =====================================================================
// validateUrl
// =====================================================================

test('validateUrl: accepts docs.google.com URL', () => {
  const ctx = loadValidatorsContext();
  const r = ctx.validateUrl('https://docs.google.com/forms/d/e/abc/viewform');
  assert.equal(r.isValid, true);
  assert.equal(r.metadata.hostname, 'docs.google.com');
});

test('validateUrl: accepts subdomains of allowed domains', () => {
  const ctx = loadValidatorsContext();
  assert.equal(ctx.validateUrl('https://foo.docs.google.com/x').isValid, true);
});

test('validateUrl: rejects non-Google hostnames', () => {
  const ctx = loadValidatorsContext();
  assert.equal(ctx.validateUrl('https://evil.example.com/').isValid, false);
});

test('validateUrl: rejects javascript: / data: protocols', () => {
  const ctx = loadValidatorsContext();
  assert.equal(ctx.validateUrl('javascript:alert(1)').isValid, false);
  assert.equal(ctx.validateUrl('data:text/html,<script>').isValid, false);
});

test('validateUrl: rejects empty or non-string inputs', () => {
  const ctx = loadValidatorsContext();
  assert.equal(ctx.validateUrl('').isValid, false);
  assert.equal(ctx.validateUrl(null).isValid, false);
  assert.equal(ctx.validateUrl(42).isValid, false);
});

test('validateUrl: sanitized output preserves protocol, host, pathname', () => {
  const ctx = loadValidatorsContext();
  const r = ctx.validateUrl('https://docs.google.com/spreadsheets/d/abc/edit?foo=bar#gid=0');
  assert.equal(r.isValid, true);
  assert.ok(r.sanitized.includes('docs.google.com'));
  assert.ok(r.sanitized.includes('/spreadsheets/d/abc'));
});

// =====================================================================
// validateSpreadsheetId
// =====================================================================

test('validateSpreadsheetId: accepts 40-50 char alphanumeric + _ + -', () => {
  const ctx = loadValidatorsContext();
  assert.equal(ctx.validateSpreadsheetId('1'.repeat(44)).isValid, true);
  assert.equal(ctx.validateSpreadsheetId('a-b_c'.repeat(9)).isValid, true); // 45 chars
});

test('validateSpreadsheetId: rejects too short (< 40)', () => {
  const ctx = loadValidatorsContext();
  assert.equal(ctx.validateSpreadsheetId('short').isValid, false);
  assert.equal(ctx.validateSpreadsheetId('a'.repeat(39)).isValid, false);
});

test('validateSpreadsheetId: rejects too long (> 50)', () => {
  const ctx = loadValidatorsContext();
  assert.equal(ctx.validateSpreadsheetId('a'.repeat(51)).isValid, false);
});

test('validateSpreadsheetId: rejects invalid characters', () => {
  const ctx = loadValidatorsContext();
  // contains / which is invalid in the ID pattern
  assert.equal(ctx.validateSpreadsheetId('a/b' + 'x'.repeat(40)).isValid, false);
  // contains spaces
  assert.equal(ctx.validateSpreadsheetId(' ' + 'x'.repeat(43)).isValid, false);
});

test('validateSpreadsheetId: rejects empty or non-string', () => {
  const ctx = loadValidatorsContext();
  assert.equal(ctx.validateSpreadsheetId('').isValid, false);
  assert.equal(ctx.validateSpreadsheetId(null).isValid, false);
  assert.equal(ctx.validateSpreadsheetId(42).isValid, false);
});

// =====================================================================
// validateText
// =====================================================================

test('validateText: escapes HTML by default', () => {
  const ctx = loadValidatorsContext();
  const r = ctx.validateText('<script>alert(1)</script>');
  assert.equal(r.isValid, true);
  assert.ok(!r.sanitized.includes('<script>'));
  // Since the dangerous-pattern scrub replaces first, then escape runs,
  // final output should not contain executable markup
  assert.ok(!r.sanitized.toLowerCase().includes('<script'));
});

test('validateText: detects javascript: and flags security risk', () => {
  const ctx = loadValidatorsContext();
  const r = ctx.validateText('Click me: javascript:alert(1)');
  assert.equal(r.metadata.hadSecurityRisk, true);
  assert.ok(!r.sanitized.includes('javascript:'));
});

test('validateText: enforces maxLength', () => {
  const ctx = loadValidatorsContext();
  const r = ctx.validateText('x'.repeat(9000));
  assert.equal(r.isValid, false);
  assert.ok(r.errors.some((e) => e.includes('以下')));
});

test('validateText: enforces minLength', () => {
  const ctx = loadValidatorsContext();
  const r = ctx.validateText('ab', { minLength: 5 });
  assert.equal(r.isValid, false);
  assert.ok(r.errors.some((e) => e.includes('以上')));
});

test('validateText: rejects non-string input with type metadata', () => {
  const ctx = loadValidatorsContext();
  const r = ctx.validateText(null);
  assert.equal(r.isValid, false);
  assert.equal(r.metadata.inputType, 'null');
});

test('validateText: collapses multiple whitespace', () => {
  const ctx = loadValidatorsContext();
  const r = ctx.validateText('hello   world   ');
  assert.equal(r.sanitized, 'hello world');
});

test('validateText: strips newlines when allowNewlines=false', () => {
  const ctx = loadValidatorsContext();
  const r = ctx.validateText('line1\nline2\r\nline3', { allowNewlines: false });
  assert.ok(!r.sanitized.includes('\n'));
  assert.ok(!r.sanitized.includes('\r'));
});

test('validateText: normalizes CRLF to LF when allowNewlines=true', () => {
  const ctx = loadValidatorsContext();
  // Run with allowHtml:true so newlines survive (default would replace whitespace)
  const r = ctx.validateText('a\r\nb\rc', { allowNewlines: true, allowHtml: true });
  // \s+ collapse still runs, but that's OK — we just check no CR remains
  assert.ok(!r.sanitized.includes('\r'));
});

// =====================================================================
// validateMapping
// =====================================================================

test('validateMapping: requires answer column index', () => {
  const ctx = loadValidatorsContext();
  const r = ctx.validateMapping({ reason: 3 });
  assert.equal(r.isValid, false);
  assert.ok(r.errors.some((e) => e.includes('answer')));
});

test('validateMapping: accepts valid mapping', () => {
  const ctx = loadValidatorsContext();
  const r = ctx.validateMapping({ answer: 2, reason: 3, name: 0, class: 1 });
  assert.equal(r.isValid, true);
  assert.equal(r.errors.length, 0);
});

test('validateMapping: rejects duplicate indices', () => {
  const ctx = loadValidatorsContext();
  const r = ctx.validateMapping({ answer: 2, reason: 2 });
  assert.equal(r.isValid, false);
  assert.ok(r.errors.some((e) => e.includes('重複')));
});

test('validateMapping: rejects non-object input', () => {
  const ctx = loadValidatorsContext();
  assert.equal(ctx.validateMapping(null).isValid, false);
  assert.equal(ctx.validateMapping('string').isValid, false);
  assert.equal(ctx.validateMapping(42).isValid, false);
});

test('validateMapping: rejects empty object', () => {
  const ctx = loadValidatorsContext();
  const r = ctx.validateMapping({});
  assert.equal(r.isValid, false);
});

test('validateMapping: unwraps { mapping: {...} } shape', () => {
  const ctx = loadValidatorsContext();
  const r = ctx.validateMapping({ mapping: { answer: 2 } });
  assert.equal(r.isValid, true);
});

test('validateMapping: warns on unknown column types but stays valid', () => {
  const ctx = loadValidatorsContext();
  const r = ctx.validateMapping({ answer: 2, unknownField: 5 });
  assert.equal(r.isValid, true);
  assert.ok(r.warnings.some((w) => w.includes('未知')));
});

test('validateMapping: warns (not errors) on invalid optional column index', () => {
  const ctx = loadValidatorsContext();
  const r = ctx.validateMapping({ answer: 2, reason: 'not-a-number' });
  assert.equal(r.isValid, true);
  assert.ok(r.warnings.some((w) => w.includes('reason')));
});

test('validateMapping: errors on non-integer answer (float)', () => {
  const ctx = loadValidatorsContext();
  const r = ctx.validateMapping({ answer: 2.5 });
  assert.equal(r.isValid, false);
});

test('validateMapping: errors on negative answer index', () => {
  const ctx = loadValidatorsContext();
  const r = ctx.validateMapping({ answer: -1 });
  assert.equal(r.isValid, false);
});

// =====================================================================
// validateConfig: displaySettings boardMode pass-through
// Why: validateConfig が displaySettings を再構築する際に boardMode を含めないと、
//      後段のサニタイズで永遠に消える regression が発生する。これを防ぐ。
// =====================================================================

test('validateConfig: preserves boardMode "auto" in sanitized displaySettings', () => {
  const ctx = loadValidatorsContext();
  const r = ctx.validateConfig({ displaySettings: { boardMode: 'auto' } });
  assert.equal(r.sanitized.displaySettings.boardMode, 'auto');
});

test('validateConfig: preserves boardMode "numberline"', () => {
  const ctx = loadValidatorsContext();
  const r = ctx.validateConfig({ displaySettings: { boardMode: 'numberline' } });
  assert.equal(r.sanitized.displaySettings.boardMode, 'numberline');
});

test('validateConfig: preserves boardMode "matrix"', () => {
  const ctx = loadValidatorsContext();
  const r = ctx.validateConfig({ displaySettings: { boardMode: 'matrix' } });
  assert.equal(r.sanitized.displaySettings.boardMode, 'matrix');
});

test('validateConfig: drops invalid boardMode silently', () => {
  const ctx = loadValidatorsContext();
  const r = ctx.validateConfig({ displaySettings: { boardMode: 'panic' } });
  assert.equal(r.sanitized.displaySettings.boardMode, undefined);
  // 他のフィールドは保持されること
  assert.equal(r.sanitized.displaySettings.showNames, false);
  assert.equal(r.sanitized.displaySettings.pageSize, 20);
});

test('validateConfig: drops boardMode when type wrong', () => {
  const ctx = loadValidatorsContext();
  const r = ctx.validateConfig({ displaySettings: { boardMode: 123 } });
  assert.equal(r.sanitized.displaySettings.boardMode, undefined);
});

// =====================================================================
// validateMapping: numericX/Y overlap with answer is allowed
// Why: 「立場」列を answer(4) かつ numericX(4) として参照するのは意図的な重複。
//      これを「列インデックス重複」エラーにすると道徳指導案の設定が壊れる。
// =====================================================================

test('validateMapping: numericX same index as answer is OK', () => {
  const ctx = loadValidatorsContext();
  const r = ctx.validateMapping({ answer: 4, numericX: 4, reason: 5, class: 2, name: 3 });
  assert.equal(r.isValid, true);
  assert.equal(r.errors.length, 0);
});

test('validateMapping: numericY same index as numericX is OK (M2 1軸表示用法)', () => {
  const ctx = loadValidatorsContext();
  const r = ctx.validateMapping({ answer: 4, numericX: 5, numericY: 5 });
  assert.equal(r.isValid, true);
});

test('validateMapping: answer/reason real duplicate still rejected', () => {
  const ctx = loadValidatorsContext();
  const r = ctx.validateMapping({ answer: 4, reason: 4 });
  assert.equal(r.isValid, false);
  assert.ok(r.errors.some(e => e.includes('重複')));
});
