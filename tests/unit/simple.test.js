/**
 * シンプルテスト - CLAUDE.md準拠版基本テスト
 * 🚀 テスト環境の動作確認用
 */

describe('Simple Test - 基本動作確認', () => {
  
  test('基本的な数値計算', () => {
    expect(2 + 2).toBe(4);
    expect(10 - 5).toBe(5);
    expect(3 * 4).toBe(12);
  });
  
  test('文字列操作', () => {
    expect('Hello' + ' World').toBe('Hello World');
    expect('test'.toUpperCase()).toBe('TEST');
  });
  
  test('配列操作', () => {
    const arr = [1, 2, 3];
    expect(arr.length).toBe(3);
    expect(arr.includes(2)).toBe(true);
  });
  
  test('オブジェクト操作', () => {
    const obj = { name: 'test', value: 123 };
    expect(obj.name).toBe('test');
    expect(obj.value).toBe(123);
  });
  
  test('JSON操作（configJSON基本テスト）', () => {
    const config = {
      spreadsheetId: 'test-sheet-id',
      setupStatus: 'completed'
    };
    
    const json = JSON.stringify(config);
    const parsed = JSON.parse(json);
    
    expect(parsed.spreadsheetId).toBe('test-sheet-id');
    expect(parsed.setupStatus).toBe('completed');
  });
});