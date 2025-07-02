const { loadCode } = require('./shared-mocks');
const { parseReactionString } = loadCode();

test('parseReactionString trims spaces and filters empties', () => {
  expect(parseReactionString(' a@example.com , b@example.com  ')).toEqual([
    'a@example.com',
    'b@example.com'
  ]);
  expect(parseReactionString('a@example.com , ')).toEqual(['a@example.com']);
  expect(parseReactionString('   ')).toEqual([]);
});
