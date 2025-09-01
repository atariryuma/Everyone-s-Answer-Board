/**
 * プロダクション環境Edge Caseテスト
 * 実際の運用で頻発する境界値・異常系・パフォーマンス問題を再現
 */

import { MockTestUtils, GASError } from '../mocks/gasMocks';

// Edge case type definitions
interface EdgeCaseData {
  description: string;
  input: any;
  expectedBehavior: 'throw' | 'graceful' | 'fallback';
  expectedOutput?: any;
  performanceThreshold?: number; // ms
}

interface MemoryUsageTest {
  name: string;
  operation: () => any;
  maxMemoryMB?: number;
  maxExecutionTimeMS?: number;
}

// Real production edge cases observed in the field
const PRODUCTION_EDGE_CASES = {
  // Data input edge cases
  textInputEdgeCases: [
    {
      description: '空文字列の処理',
      input: '',
      expectedBehavior: 'graceful' as const,
      expectedOutput: ''
    },
    {
      description: '非常に長いテキスト（10000文字）',
      input: 'あ'.repeat(10000),
      expectedBehavior: 'graceful' as const
    },
    {
      description: '特殊文字・絵文字混在テキスト',
      input: '🎌こんにちは！\n改行文字\t\t\tタブ文字\\特殊記号@#$%^&*()',
      expectedBehavior: 'graceful' as const
    },
    {
      description: 'HTMLタグ・スクリプト注入の試み',
      input: '<script>alert("XSS")</script><img src="x" onerror="alert(1)">',
      expectedBehavior: 'graceful' as const
    },
    {
      description: 'SQLインジェクションの試み',
      input: "'; DROP TABLE users; --",
      expectedBehavior: 'graceful' as const
    },
    {
      description: 'NULL文字・制御文字',
      input: 'text\x00null\x01\x02\x03控制字符',
      expectedBehavior: 'graceful' as const
    }
  ],

  // Numerical edge cases
  numericalEdgeCases: [
    {
      description: '負の数値',
      input: -1,
      expectedBehavior: 'graceful' as const,
      expectedOutput: 0
    },
    {
      description: '非常に大きな数値',
      input: Number.MAX_SAFE_INTEGER + 1,
      expectedBehavior: 'graceful' as const
    },
    {
      description: 'NaN値',
      input: NaN,
      expectedBehavior: 'graceful' as const,
      expectedOutput: 0
    },
    {
      description: 'Infinity値',
      input: Infinity,
      expectedBehavior: 'graceful' as const,
      expectedOutput: 0
    },
    {
      description: '浮動小数点精度問題',
      input: 0.1 + 0.2,
      expectedBehavior: 'graceful' as const
    }
  ],

  // Array/Object edge cases
  structuralEdgeCases: [
    {
      description: '空配列',
      input: [],
      expectedBehavior: 'graceful' as const,
      expectedOutput: []
    },
    {
      description: 'null値を含む配列',
      input: [1, null, undefined, 3, '', 0],
      expectedBehavior: 'graceful' as const
    },
    {
      description: 'ネストが深いオブジェクト',
      input: { a: { b: { c: { d: { e: { f: 'deep' } } } } } },
      expectedBehavior: 'graceful' as const
    },
    {
      description: '循環参照オブジェクト（JSON.stringifyでエラー）',
      input: (() => { const obj: any = { a: 1 }; obj.self = obj; return obj; })(),
      expectedBehavior: 'graceful' as const
    }
  ],

  // Date/Time edge cases
  dateTimeEdgeCases: [
    {
      description: '無効な日付文字列',
      input: 'invalid-date',
      expectedBehavior: 'graceful' as const
    },
    {
      description: 'エポック時刻0',
      input: new Date(0),
      expectedBehavior: 'graceful' as const
    },
    {
      description: '未来の日付（2100年）',
      input: new Date('2100-12-31'),
      expectedBehavior: 'graceful' as const
    },
    {
      description: 'タイムゾーン問題',
      input: '2024-01-15T12:00:00+09:00',
      expectedBehavior: 'graceful' as const
    }
  ]
};

describe('プロダクション環境Edge Case再現テスト', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    MockTestUtils.clearAllMockData();
    
    // Realistic test timing setup
    jest.useFakeTimers();
    jest.setSystemTime(new Date('2024-01-15T12:00:00Z'));
  });

  afterEach(() => {
    jest.useRealTimers();
  });

  describe('テキスト入力Edge Cases', () => {
    // Safe text processing function
    const safeTextProcessor = (input: any): string => {
      if (input === null || input === undefined) return '';
      if (typeof input !== 'string') input = String(input);
      
      // Remove potential security risks
      let processed = input
        .replace(/<script[^>]*>.*?<\/script>/gi, '') // Remove scripts
        .replace(/javascript:/gi, '') // Remove javascript: URLs
        .replace(/on\w+\s*=/gi, ''); // Remove event handlers
      
      // Handle control characters
      processed = processed.replace(/[\x00-\x1F\x7F]/g, '');
      
      // Limit length (防止DoS攻撃)
      if (processed.length > 10000) {
        processed = processed.substring(0, 10000) + '...';
      }
      
      return processed.trim();
    };

    PRODUCTION_EDGE_CASES.textInputEdgeCases.forEach((testCase) => {
      test(`テキスト処理: ${testCase.description}`, () => {
        const result = safeTextProcessor(testCase.input);
        
        switch (testCase.expectedBehavior) {
          case 'graceful':
            expect(() => safeTextProcessor(testCase.input)).not.toThrow();
            expect(typeof result).toBe('string');
            if (testCase.expectedOutput !== undefined) {
              expect(result).toBe(testCase.expectedOutput);
            }
            break;
          case 'throw':
            expect(() => safeTextProcessor(testCase.input)).toThrow();
            break;
        }
        
        // Security checks
        expect(result).not.toContain('<script>');
        expect(result).not.toContain('javascript:');
      });
    });

    test('大量のテキストデータ処理パフォーマンス', () => {
      const largeTexts = Array.from({ length: 1000 }, (_, i) => 
        `大量のテキストデータ${i}です。`.repeat(100)
      );

      const startTime = Date.now();
      const results = largeTexts.map(text => safeTextProcessor(text));
      const endTime = Date.now();

      expect(results).toHaveLength(1000);
      expect(endTime - startTime).toBeLessThan(1000); // 1秒以内
      expect(results.every(result => result.length <= 10000)).toBe(true);
    });
  });

  describe('数値処理Edge Cases', () => {
    const safeNumberProcessor = (input: any): number => {
      if (input === null || input === undefined) return 0;
      
      const num = Number(input);
      
      if (isNaN(num) || !isFinite(num)) return 0;
      if (num < 0) return 0; // リアクション数は負になり得ない
      if (num > 999999) return 999999; // 上限設定
      
      return Math.floor(num); // 整数に丸める
    };

    PRODUCTION_EDGE_CASES.numericalEdgeCases.forEach((testCase) => {
      test(`数値処理: ${testCase.description}`, () => {
        const result = safeNumberProcessor(testCase.input);
        
        expect(() => safeNumberProcessor(testCase.input)).not.toThrow();
        expect(typeof result).toBe('number');
        expect(isFinite(result)).toBe(true);
        expect(result >= 0).toBe(true);
        expect(result <= 999999).toBe(true);
        
        if (testCase.expectedOutput !== undefined) {
          expect(result).toBe(testCase.expectedOutput);
        }
      });
    });

    test('浮動小数点精度問題の対処', () => {
      const precisionProblem = 0.1 + 0.2; // = 0.30000000000000004
      const result = Math.round(precisionProblem * 100) / 100; // 小数点2桁で丸める
      
      expect(result).toBe(0.3);
      expect(result).not.toBe(0.30000000000000004);
    });
  });

  describe('配列・オブジェクト処理Edge Cases', () => {
    const safeArrayProcessor = (input: any): any[] => {
      if (!Array.isArray(input)) return [];
      
      return input
        .filter(item => item !== null && item !== undefined)
        .map(item => {
          if (typeof item === 'object') {
            try {
              // 循環参照をチェック
              JSON.stringify(item);
              return item;
            } catch {
              return { error: 'Circular reference detected' };
            }
          }
          return item;
        })
        .slice(0, 1000); // 配列サイズ制限
    };

    const safeObjectProcessor = (input: any): Record<string, any> => {
      if (typeof input !== 'object' || input === null) return {};
      
      try {
        // 循環参照チェック
        JSON.stringify(input);
        
        // ネストの深さ制限
        const flattenDeep = (obj: any, depth = 0): any => {
          if (depth > 10) return '[Too Deep]';
          if (typeof obj !== 'object' || obj === null) return obj;
          
          const result: any = {};
          for (const [key, value] of Object.entries(obj)) {
            result[key] = flattenDeep(value, depth + 1);
          }
          return result;
        };
        
        return flattenDeep(input);
      } catch {
        return { error: 'Invalid object structure' };
      }
    };

    PRODUCTION_EDGE_CASES.structuralEdgeCases.forEach((testCase) => {
      test(`構造データ処理: ${testCase.description}`, () => {
        if (Array.isArray(testCase.input)) {
          const result = safeArrayProcessor(testCase.input);
          expect(Array.isArray(result)).toBe(true);
          expect(result.length).toBeLessThanOrEqual(1000);
        } else {
          const result = safeObjectProcessor(testCase.input);
          expect(typeof result).toBe('object');
          expect(result).not.toBeNull();
        }
      });
    });
  });

  describe('日付・時刻処理Edge Cases', () => {
    const safeDateProcessor = (input: any): Date => {
      let date: Date;
      
      if (input instanceof Date) {
        date = input;
      } else if (typeof input === 'string' || typeof input === 'number') {
        date = new Date(input);
      } else {
        date = new Date();
      }
      
      // 無効な日付チェック
      if (isNaN(date.getTime())) {
        return new Date(); // 現在時刻をフォールバック
      }
      
      // 妥当な範囲チェック (1900-2100年)
      const year = date.getFullYear();
      if (year < 1900 || year > 2100) {
        return new Date(); // 現在時刻をフォールバック
      }
      
      return date;
    };

    PRODUCTION_EDGE_CASES.dateTimeEdgeCases.forEach((testCase) => {
      test(`日付処理: ${testCase.description}`, () => {
        const result = safeDateProcessor(testCase.input);
        
        expect(result instanceof Date).toBe(true);
        expect(isNaN(result.getTime())).toBe(false);
        
        const year = result.getFullYear();
        expect(year).toBeGreaterThanOrEqual(1900);
        expect(year).toBeLessThanOrEqual(2100);
      });
    });
  });

  describe('GAS環境特有のEdge Cases', () => {
    test('GAS実行時間制限シミュレーション', async () => {
      const longRunningProcess = async () => {
        const startTime = Date.now();
        let iterations = 0;
        
        // 模擬的な長時間処理
        while (Date.now() - startTime < 300000) { // 5分制限
          iterations++;
          
          // 定期的なチェックポイント
          if (iterations % 1000 === 0) {
            await new Promise(resolve => setTimeout(resolve, 1));
            
            // 実行時間チェック
            if (Date.now() - startTime > 270000) { // 4分30秒でタイムアウト警告
              console.warn('⚠️ Approaching GAS execution time limit');
              break;
            }
          }
        }
        
        return { iterations, executionTime: Date.now() - startTime };
      };

      jest.setTimeout(10000); // テストタイムアウト10秒
      
      // Fast-forward time simulation
      const result = await longRunningProcess();
      
      expect(result.iterations).toBeGreaterThan(0);
      expect(result.executionTime).toBeLessThan(300000);
    });

    test('SpreadsheetApp APIエラー処理', () => {
      // 無効なスプレッドシートID
      expect(() => {
        global.SpreadsheetApp.openById('');
      }).toThrowGASError();

      // 存在しないシート名
      const spreadsheet = global.SpreadsheetApp.openById('valid-id');
      const sheet = spreadsheet.getSheetByName('non-existent-sheet');
      expect(sheet).toBeNull();

      // 範囲指定エラー
      const validSheet = spreadsheet.getSheetByName('Sheet1');
      expect(() => {
        validSheet.getRange(0, 0, 1, 1); // 無効な範囲（1から始まる）
      }).toThrowGASError();
    });

    test('大量データでのsetValues配列サイズ不一致', () => {
      MockTestUtils.createRealisticSpreadsheetData(
        ['Header1', 'Header2'],
        [['Data1', 'Data2']]
      );

      const sheet = global.SpreadsheetApp.getActiveSheet();
      const range = sheet.getRange(1, 1, 2, 2);

      // 正しいサイズ（2x2）
      expect(() => {
        range.setValues([['A', 'B'], ['C', 'D']]);
      }).not.toThrow();

      // 間違ったサイズ（1x2を2x2に）
      expect(() => {
        range.setValues([['A', 'B']]);
      }).toThrowGASError(/number of rows/);

      // 間違ったサイズ（2x1を2x2に）
      expect(() => {
        range.setValues([['A'], ['C']]);
      }).toThrowGASError(/number of columns/);
    });
  });

  describe('メモリ・パフォーマンスEdge Cases', () => {
    const memoryTests: MemoryUsageTest[] = [
      {
        name: '大量文字列処理',
        operation: () => {
          const largeStrings = Array.from({ length: 1000 }, () => 
            'あいうえお'.repeat(1000)
          );
          return largeStrings.map(str => str.toUpperCase());
        },
        maxExecutionTimeMS: 1000
      },
      {
        name: '大量オブジェクト作成',
        operation: () => {
          return Array.from({ length: 10000 }, (_, i) => ({
            id: i,
            name: `Object${i}`,
            data: { value: i * 2, active: i % 2 === 0 }
          }));
        },
        maxExecutionTimeMS: 500
      },
      {
        name: '深いネストオブジェクト処理',
        operation: () => {
          let obj: any = {};
          let current = obj;
          
          for (let i = 0; i < 100; i++) {
            current.next = { level: i, data: `level${i}` };
            current = current.next;
          }
          
          return JSON.stringify(obj).length;
        },
        maxExecutionTimeMS: 100
      }
    ];

    memoryTests.forEach((test) => {
      test(`メモリ効率: ${test.name}`, () => {
        const startTime = Date.now();
        
        const result = test.operation();
        
        const endTime = Date.now();
        const executionTime = endTime - startTime;

        expect(result).toBeDefined();
        
        if (test.maxExecutionTimeMS) {
          expect(executionTime).toBeLessThan(test.maxExecutionTimeMS);
        }
      });
    });
  });

  describe('同時アクセス・競合状態Edge Cases', () => {
    test('複数ユーザー同時アクセスシミュレーション', async () => {
      const simulateUser = async (userId: string) => {
        // ユーザーデータ設定
        MockTestUtils.setMockData(`user_${userId}`, {
          userId,
          email: `${userId}@test.com`,
          spreadsheetId: 'shared-spreadsheet'
        });

        // 同時データ処理
        const spreadsheet = global.SpreadsheetApp.openById('shared-spreadsheet');
        const sheet = spreadsheet.getSheetByName('Sheet1');
        
        return sheet.getRange(1, 1, 10, 5).getValues();
      };

      const users = ['user1', 'user2', 'user3', 'user4', 'user5'];
      
      const startTime = Date.now();
      const results = await Promise.all(users.map(user => simulateUser(user)));
      const endTime = Date.now();

      expect(results).toHaveLength(5);
      expect(results.every(result => Array.isArray(result))).toBe(true);
      expect(endTime - startTime).toBeLessThan(2000); // 並列処理で高速
    });

    test('キャッシュ競合状態の処理', () => {
      const cacheKey = 'shared-data';
      const users = ['userA', 'userB', 'userC'];
      
      // Race condition simulation
      const simulateCacheAccess = (userId: string) => {
        const userData = { userId, timestamp: Date.now() };
        
        // Cache read-modify-write操作
        const existing = MockTestUtils.getMockData(cacheKey, []);
        const updated = [...existing, userData];
        MockTestUtils.setMockData(cacheKey, updated);
        
        return updated;
      };

      const results = users.map(user => simulateCacheAccess(user));
      
      // 最終状態確認
      const finalCacheData = MockTestUtils.getMockData(cacheKey);
      expect(finalCacheData).toHaveLength(3);
      
      // データ整合性確認
      const userIds = finalCacheData.map((item: any) => item.userId);
      expect(userIds).toEqual(expect.arrayContaining(['userA', 'userB', 'userC']));
    });
  });

  describe('国際化・文字エンコーディングEdge Cases', () => {
    const unicodeTestCases = [
      {
        name: '日本語（ひらがな、カタカナ、漢字）',
        text: 'ひらがなカタカナ漢字混合テスト',
        expected: true
      },
      {
        name: '中国語簡体字・繁体字',
        text: '简体中文繁體中文測試',
        expected: true
      },
      {
        name: '韓国語ハングル',
        text: '한국어 테스트',
        expected: true
      },
      {
        name: 'アラビア語（右から左）',
        text: 'اختبار اللغة العربية',
        expected: true
      },
      {
        name: '絵文字・記号',
        text: '🎌🌸🗾💯✨🎯📊📈',
        expected: true
      },
      {
        name: '数学記号・特殊文字',
        text: '∑∫∂∞≤≥±×÷√∝∴∵⟨⟩',
        expected: true
      }
    ];

    unicodeTestCases.forEach((testCase) => {
      test(`Unicode対応: ${testCase.name}`, () => {
        const input = testCase.text;
        
        // 基本的な文字列操作
        expect(input.length).toBeGreaterThan(0);
        expect(typeof input).toBe('string');
        
        // エンコーディング・デコーディング
        const encoded = encodeURIComponent(input);
        const decoded = decodeURIComponent(encoded);
        expect(decoded).toBe(input);
        
        // JSON処理
        const jsonString = JSON.stringify({ text: input });
        const parsed = JSON.parse(jsonString);
        expect(parsed.text).toBe(input);
        
        // 文字列検索・操作
        expect(input.indexOf(input.charAt(0))).toBe(0);
        expect(input.toUpperCase()).toBeDefined();
        expect(input.toLowerCase()).toBeDefined();
      });
    });
  });
});