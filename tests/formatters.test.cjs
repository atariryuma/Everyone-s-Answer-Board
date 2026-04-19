const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function loadFormattersContext(overrides = {}) {
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    Utilities: {
      formatDate: (date, _tz, fmt) => {
        // Simplified deterministic formatter for tests: always 2026/04/19 style
        const yyyy = date.getUTCFullYear();
        const MM = String(date.getUTCMonth() + 1).padStart(2, '0');
        const dd = String(date.getUTCDate()).padStart(2, '0');
        const HH = String(date.getUTCHours()).padStart(2, '0');
        const mm = String(date.getUTCMinutes()).padStart(2, '0');
        if (fmt === 'yyyy/MM/dd HH:mm') return `${yyyy}/${MM}/${dd} ${HH}:${mm}`;
        return date.toISOString();
      }
    },
    Session: { getScriptTimeZone: () => 'UTC' },
    ...overrides
  };
  vm.createContext(context);
  const source = fs.readFileSync(path.resolve(__dirname, '../src/formatters.js'), 'utf8');
  vm.runInContext(source, context, { filename: 'formatters.js' });
  return context;
}

// =====================================================================
// formatTimestamp
// =====================================================================

test('formatTimestamp: formats ISO string into yyyy/MM/dd HH:mm', () => {
  const ctx = loadFormattersContext();
  const result = ctx.formatTimestamp('2026-04-19T08:30:00Z');
  assert.equal(result, '2026/04/19 08:30');
});

test('formatTimestamp: accepts Date object', () => {
  const ctx = loadFormattersContext();
  const result = ctx.formatTimestamp(new Date('2026-01-02T15:45:00Z'));
  assert.equal(result, '2026/01/02 15:45');
});

test('formatTimestamp: returns "-" for empty/null/undefined', () => {
  const ctx = loadFormattersContext();
  assert.equal(ctx.formatTimestamp(''), '-');
  assert.equal(ctx.formatTimestamp(null), '-');
  assert.equal(ctx.formatTimestamp(undefined), '-');
});

test('formatTimestamp: returns "-" for invalid date string', () => {
  const ctx = loadFormattersContext();
  assert.equal(ctx.formatTimestamp('not-a-date'), '-');
});

test('formatTimestamp: returns "-" when Utilities.formatDate throws', () => {
  const ctx = loadFormattersContext({
    Utilities: { formatDate: () => { throw new Error('tz error'); } },
    Session: { getScriptTimeZone: () => 'UTC' }
  });
  assert.equal(ctx.formatTimestamp('2026-04-19T00:00:00Z'), '-');
});

// =====================================================================
// htmlEncode
// =====================================================================

test('htmlEncode: escapes the five HTML-significant characters', () => {
  const ctx = loadFormattersContext();
  const result = ctx.htmlEncode(`<img src="x" onerror='alert(1)'> & "Q"`);
  assert.ok(!result.includes('<img'));
  assert.ok(result.includes('&lt;img'));
  assert.ok(result.includes('&quot;'));
  assert.ok(result.includes('&#39;'));
  assert.ok(result.includes('&amp;'));
});

test('htmlEncode: returns empty string for null/undefined/empty', () => {
  const ctx = loadFormattersContext();
  assert.equal(ctx.htmlEncode(null), '');
  assert.equal(ctx.htmlEncode(undefined), '');
  assert.equal(ctx.htmlEncode(''), '');
});

test('htmlEncode: coerces non-string input to string', () => {
  const ctx = loadFormattersContext();
  assert.equal(ctx.htmlEncode(42), '42');
  assert.equal(ctx.htmlEncode(true), 'true');
});

test('htmlEncode: preserves safe characters', () => {
  const ctx = loadFormattersContext();
  assert.equal(ctx.htmlEncode('Hello, 世界! 123'), 'Hello, 世界! 123');
});

test('htmlEncode: escapes ampersand first to avoid double-encoding', () => {
  const ctx = loadFormattersContext();
  // If & were escaped after <, we'd get &amp;lt; instead of &lt;
  const result = ctx.htmlEncode('<a&b>');
  assert.equal(result, '&lt;a&amp;b&gt;');
});
