const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function loadCtx() {
  const context = {
    console: { log: () => {}, warn: () => {}, error: () => {} },
    normalizeHeader: (h) => String(h || '').toLowerCase().trim()
  };
  vm.createContext(context);
  const source = fs.readFileSync(path.resolve(__dirname, '../src/ColumnMappingService.js'), 'utf8');
  vm.runInContext(source, context, { filename: 'ColumnMappingService.js' });
  return context;
}

// =====================================================================
// performIntegratedColumnDiagnostics: 統合返却に candidates が含まれること
// =====================================================================

test('performIntegratedColumnDiagnostics: returns numericScaleCandidates field', () => {
  const ctx = loadCtx();
  const headers = ['タイムスタンプ', 'メール', '回答', '評価'];
  const sample = [
    ['2026', 'a@x', 'はい', 1],
    ['2026', 'b@x', 'いいえ', 5]
  ];
  const result = ctx.performIntegratedColumnDiagnostics(headers, { sampleData: sample });
  assert.equal(result.success, true);
  assert.ok(Array.isArray(result.numericScaleCandidates));
  assert.equal(result.numericScaleCandidates.length, 1);
  assert.equal(result.numericScaleCandidates[0].header, '評価');
});
