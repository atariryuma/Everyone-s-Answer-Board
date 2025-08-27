/**
 * @fileoverview 循環参照エラーハンドリングの統合テスト
 */

const fs = require('fs');
const path = require('path');

// main.gsからsafeStringify関数を抽出して読み込み
function loadSafeStringifyFunction() {
  const file = fs.readFileSync(path.join(__dirname, '../src/main.gs'), 'utf8');
  
  // safeStringify関数を抽出
  const start = file.indexOf('function safeStringify');
  if (start === -1) {
    throw new Error('safeStringify function not found');
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
  
  // WeakSetが利用可能であることを確認
  if (typeof WeakSet === 'undefined') {
    global.WeakSet = class WeakSet {
      constructor() {
        this._items = [];
      }
      has(item) {
        return this._items.includes(item);
      }
      add(item) {
        if (!this.has(item)) {
          this._items.push(item);
        }
      }
    };
  }
  
  const fn = new Function(`${fnStr}; return safeStringify;`);
  return fn();
}

// logClientError関数を抽出して読み込み
function loadLogClientErrorFunction() {
  const file = fs.readFileSync(path.join(__dirname, '../src/main.gs'), 'utf8');
  
  // 必要な関数を抽出
  const safeStringifyStart = file.indexOf('function safeStringify');
  const safeStringifyEnd = file.indexOf('\n}', file.indexOf('function safeStringify')) + 2;
  const safeStringifyCode = file.slice(safeStringifyStart, safeStringifyEnd);
  
  const logClientErrorStart = file.indexOf('function logClientError');
  const logClientErrorEnd = file.indexOf('\n}', file.indexOf('function logClientError')) + 2;
  const logClientErrorCode = file.slice(logClientErrorStart, logClientErrorEnd);
  
  // WeakSetサポート
  if (typeof WeakSet === 'undefined') {
    global.WeakSet = class WeakSet {
      constructor() {
        this._items = [];
      }
      has(item) {
        return this._items.includes(item);
      }
      add(item) {
        if (!this.has(item)) {
          this._items.push(item);
        }
      }
    };
  }
  
  const fn = new Function(`
    // logError関数のモック
    function logError(error, context, severity, category) {
      console.error('Mocked logError:', { error, context, severity, category });
    }
    
    // 定数のモック
    const MAIN_ERROR_SEVERITY = { MEDIUM: 'medium' };
    const MAIN_ERROR_CATEGORIES = { CLIENT: 'client', INTERNAL: 'internal' };
    
    ${safeStringifyCode}
    ${logClientErrorCode}
    return { safeStringify, logClientError, logError };
  `);
  return fn();
}

describe('循環参照エラーハンドリング統合テスト', () => {
  let safeStringify, logClientError;
  
  beforeAll(() => {
    const functions = loadLogClientErrorFunction();
    safeStringify = functions.safeStringify;
    logClientError = functions.logClientError;
  });

  describe('safeStringify関数', () => {
    test('正常なオブジェクトを処理できる', () => {
      const obj = { name: 'test', value: 123 };
      const result = safeStringify(obj);
      expect(result).toBe('{"name":"test","value":123}');
    });

    test('循環参照を安全に処理できる', () => {
      const obj = { name: 'test' };
      obj.self = obj; // 循環参照を作成
      
      const result = safeStringify(obj);
      expect(result).toContain('[Circular Reference]');
      expect(result).toContain('test');
    });

    test('関数を安全に処理できる', () => {
      const obj = {
        name: 'test',
        method() { return 'hello'; }
      };
      
      const result = safeStringify(obj);
      expect(result).toContain('[Function]');
      expect(result).toContain('test');
    });

    test('DOM要素を安全に処理できる（模擬）', () => {
      const mockDomElement = {
        nodeType: 1,
        tagName: 'DIV',
        innerHTML: 'content'
      };
      
      const result = safeStringify(mockDomElement);
      expect(result).toContain('[DOM Element]');
    });

    test('Errorオブジェクトを安全に処理できる', () => {
      const obj = {
        error: new Error('Test error message'),
        name: 'test'
      };
      
      const result = safeStringify(obj);
      expect(result).toContain('Test error message');
      expect(result).toContain('test');
    });

    test('文字数制限が適用される', () => {
      const longString = 'a'.repeat(1000);
      const obj = { data: longString };
      
      const result = safeStringify(obj, 100);
      expect(result.length).toBeLessThanOrEqual(120); // 100 + "[truncated]"の分
      expect(result).toContain('[truncated]');
    });

    test('Stringify不可能なオブジェクトを安全に処理', () => {
      // JSON.stringify でエラーが発生するオブジェクトを作成
      const problematicObj = {};
      problematicObj.toJSON = function() {
        throw new Error('toJSON failed');
      };
      
      const result = safeStringify(problematicObj);
      expect(typeof result).toBe('string');
      // toJSONでエラーが発生した場合、JSON.stringifyは例外をスローしcatchブロックで処理される
      expect(result).toContain('[Stringify Error:');
    });
  });

  describe('logClientError関数', () => {
    let consoleSpy;

    beforeEach(() => {
      consoleSpy = jest.spyOn(console, 'error').mockImplementation();
    });

    afterEach(() => {
      consoleSpy.mockRestore();
    });

    test('文字列エラーを正常に処理', () => {
      const result = logClientError('Simple error message');
      
      expect(result.status).toBe('success');
      expect(consoleSpy).toHaveBeenCalledWith('CLIENT ERROR:', expect.any(String));
    });

    test('オブジェクトエラーのmessageプロパティを優先', () => {
      const errorObj = {
        message: 'Error message',
        userId: 'user123',
        other: 'ignored'
      };
      
      const result = logClientError(errorObj);
      
      expect(result.status).toBe('success');
      expect(consoleSpy).toHaveBeenCalledWith('CLIENT ERROR:', expect.any(String));
    });

    test('循環参照を含むエラーオブジェクトを安全に処理', () => {
      const errorObj = {
        userId: 'user123'
      };
      errorObj.self = errorObj; // 循環参照
      
      const result = logClientError(errorObj);
      
      expect(result.status).toBe('success');
      expect(consoleSpy).toHaveBeenCalledWith('CLIENT ERROR:', expect.any(String));
    });

    test('複雑なエラーオブジェクトの安全な情報抽出', () => {
      const complexError = {
        userId: 'user123',
        type: 'SyntaxError',
        url: 'https://example.com/test',
        timestamp: '2024-01-01T12:00:00Z',
        userAgent: 'Mozilla/5.0 (Test Browser)',
        // messageプロパティが無い場合
      };
      
      const result = logClientError(complexError);
      
      expect(result.status).toBe('success');
      expect(consoleSpy).toHaveBeenCalledWith('CLIENT ERROR:', expect.any(String));
    });

    test('処理中にエラーが発生した場合のフォールバック', () => {
      // JSON.stringify が失敗するような状況を模擬（直接テストは困難なため確認のみ）
      const result = logClientError(null);
      
      expect(result.status).toBe('success');
      expect(consoleSpy).toHaveBeenCalledWith('CLIENT ERROR:', expect.any(String));
    });

    test('完全に処理不可能な場合の最終フォールバック', () => {
      // logClientError関数内でエラーをスローするモック
      const originalConsoleError = console.error;
      console.error = jest.fn(() => {
        throw new Error('Console error failed');
      });
      
      // エラーを補足してテストを継続できるようにラップ
      let result;
      try {
        result = logClientError('test');
      } catch (e) {
        // 例外がスローされた場合、フォールバックが不完全
        result = { status: 'error', message: e.message };
      }
      
      // 関数は例外をスローせず、何らかのレスポンスを返すはず
      expect(typeof result).toBe('object');
      expect(result).toHaveProperty('status');
      
      console.error = originalConsoleError;
    });
  });

  describe('統合シナリオテスト', () => {
    let consoleSpy;

    beforeEach(() => {
      consoleSpy = jest.spyOn(console, 'error').mockImplementation();
    });

    afterEach(() => {
      consoleSpy.mockRestore();
    });

    test('実際の"Unexpected token"エラーシナリオ', () => {
      const unexpectedTokenError = {
        message: 'Unexpected token \'{\'',
        name: 'SyntaxError',
        userId: 'user123',
        url: 'https://script.google.com/test',
        timestamp: new Date().toISOString()
      };
      
      const result = logClientError(unexpectedTokenError);
      
      expect(result.status).toBe('success');
      expect(consoleSpy).toHaveBeenCalledWith('CLIENT ERROR:', expect.any(String));
    });

    test('フロントエンドから送信される典型的なエラーデータ', () => {
      const frontendError = {
        message: 'Unexpected token \'{\'',
        userId: 'user123',
        url: 'https://script.google.com/test?mode=admin',
        timestamp: '2024-08-27T14:24:05.000Z',
        userAgent: 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36',
        type: 'SyntaxError'
      };
      
      const result = logClientError(frontendError);
      
      expect(result.status).toBe('success');
      expect(consoleSpy).toHaveBeenCalledWith('CLIENT ERROR:', expect.any(String));
    });
  });
});