/**
 * @fileoverview parseUserConfigのテスト
 */

const fs = require('fs');
const path = require('path');

function loadParseUserConfig() {
  const file = fs.readFileSync(path.join(__dirname, '../src/SharedUtilities.html'), 'utf8');
  const start = file.indexOf('function parseUserConfig');
  if (start === -1) {
    throw new Error('parseUserConfig not found');
  }
  let i = file.indexOf('{', start);
  let depth = 1;
  i += 1;
  while (i < file.length && depth > 0) {
    const char = file[i];
    if (char === '{') depth += 1;
    if (char === '}') depth -= 1;
    i += 1;
  }
  const fnStr = file.slice(start, i);
  const fn = new Function(`${fnStr}; return parseUserConfig;`);
  return fn();
}

describe('parseUserConfig', () => {
  const parseUserConfig = loadParseUserConfig();

  test('空文字列は空オブジェクトを返す', () => {
    expect(parseUserConfig('')).toEqual({});
  });

  test('不正なJSON文字列は空オブジェクトを返す', () => {
    expect(parseUserConfig('{invalid')).toEqual({});
  });

  test('正しいJSON文字列はオブジェクトを返す', () => {
    expect(parseUserConfig('{"a":1}')).toEqual({ a: 1 });
  });
});
