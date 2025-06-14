const { findHeaderIndices } = require('../src/Code.gs');

test('findHeaderIndices returns indices for existing headers', () => {
  const headers = ['A', 'B', 'C'];
  const required = ['A', 'C'];
  expect(findHeaderIndices(headers, required)).toEqual({ A:0, C:2 });
});

test('findHeaderIndices throws for missing headers', () => {
  expect(() => findHeaderIndices(['X'], ['X', 'Y'])).toThrow('必須ヘッダーが見つかりません');
});

test('findHeaderIndices ignores whitespace differences', () => {
  const headers = ['A B', ' C\nD '];
  const required = ['AB', 'CD'];
  expect(findHeaderIndices(headers, required)).toEqual({ AB:0, CD:1 });
});
