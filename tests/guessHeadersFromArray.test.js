const { guessHeadersFromArray } = require('../src/Code.gs');

test('guessHeadersFromArray detects japanese synonyms', () => {
  const headers = ['問い', 'コメント', 'わけ', 'ニックネーム', '班'];
  const result = guessHeadersFromArray(headers);
  expect(result).toEqual({
    questionHeader: 'コメント',
    answerHeader: 'コメント',
    reasonHeader: 'わけ',
    nameHeader: 'ニックネーム',
    classHeader: '班'
  });
});

test('guessHeadersFromArray handles full width and spacing', () => {
  const headers = [' Ｑｕｅｓｔｉｏｎ ', 'Ａｎｓｗｅｒ', ' Reason ', ' NAME ', 'CLASSROOM'];
  const result = guessHeadersFromArray(headers);
  expect(result).toEqual({
    questionHeader: 'Ａｎｓｗｅｒ',
    answerHeader: 'Ａｎｓｗｅｒ',
    reasonHeader: ' Reason ',
    nameHeader: ' NAME ',
    classHeader: 'CLASSROOM'
  });
});
