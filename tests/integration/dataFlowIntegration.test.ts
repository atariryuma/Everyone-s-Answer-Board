/**
 * データフロー統合テストスイート
 * 実際のワークフロー全体での動作確認：読み込み→列検出→処理→書き込み
 */

import { MockTestUtils, GASError } from '../mocks/gasMocks';

// Type definitions for integration testing
interface SpreadsheetData {
  headers: string[];
  rows: any[][];
}

interface ColumnMapping {
  answer: number;
  reason: number;
  class: number;
  name: number;
  confidence: {
    answer: number;
    reason: number;
    class: number;
    name: number;
  };
}

interface ProcessedResponse {
  id: string;
  timestamp: string;
  email: string;
  answer: string;
  reason?: string;
  class?: string;
  name?: string;
  reactions: {
    UNDERSTAND: number;
    LIKE: number;
    CURIOUS: number;
  };
}

interface AdminPanelResult {
  success: boolean;
  data?: ProcessedResponse[];
  error?: string;
  stats?: {
    totalResponses: number;
    averageReactions: number;
    uniqueClasses: string[];
  };
}

// Real-world test data that mirrors production
const PRODUCTION_LIKE_DATA = {
  // Japanese Google Forms typical headers
  formResponseHeaders: [
    'タイムスタンプ',
    'メールアドレス', 
    'あなたの意見を教えてください',
    'その理由を教えてください',
    'あなたのクラスを教えてください',
    'お名前を教えてください（任意）',
    'なるほど！',
    'いいね！', 
    'もっと知りたい！'
  ],
  
  // Typical form response data
  formResponseData: [
    [
      '2024/01/15 14:23:45',
      'student1@school.ed.jp',
      'AIの発達により、今後は創造性がより重要になると思います',
      '単純作業はAIに任せて、人間は創造的な仕事に集中できるから',
      '3年A組',
      '田中太郎',
      '5', '3', '2'
    ],
    [
      '2024/01/15 15:12:33',
      'student2@school.ed.jp', 
      '環境問題への取り組みは個人レベルでも重要だと考えます',
      '小さな行動でも積み重なると大きな変化になるため',
      '2年B組',
      '佐藤花子',
      '8', '6', '4'
    ],
    [
      '2024/01/15 16:45:12',
      'student3@school.ed.jp',
      'デジタル化が進んでも、対面でのコミュニケーションは大切だと思う',
      '人と人の直接的なつながりが信頼関係を築くため',
      '1年C組', 
      '',  // Empty name (optional field)
      '12', '8', '6'
    ],
    [
      '2024/01/15 17:30:21',
      'student4@school.ed.jp',
      '部活動を通じて学んだチームワークが将来に活かされると感じています',
      '協力することで個人では達成できない成果が得られるから',
      '3年A組',
      '鈴木次郎', 
      '7', '9', '3'
    ]
  ],

  // Edge case data - problematic formats that occur in production
  edgeCaseData: [
    [
      '2024/01/15 18:15:30',
      'teacher@school.ed.jp',  // Teacher email (different domain pattern)
      '改行を含む\n意見です。\nこのような\n場合もあります。',  // Multi-line text
      '理由も\n複数行に\nわたる場合があります。',
      '教員',  // Non-standard class format
      '先生',
      '0', '0', '0'  // Zero reactions
    ],
    [
      '2024/01/15 19:22:18',
      'student5@school.ed.jp',
      'とても長い意見です。'.repeat(50),  // Very long text
      'とても長い理由です。'.repeat(30),
      '3年A組',
      '非常に長い名前の生徒です'.repeat(5),
      '999', '999', '999'  // Extremely high reaction counts
    ],
    [
      '2024/01/15 20:10:45',
      'incomplete@school.ed.jp',
      '不完全なデータ',  // Missing reason
      '',  // Empty reason
      '',  // Empty class
      '',  // Empty name
      '', '', ''  // Empty reaction counts
    ]
  ]
};

describe('データフロー統合テスト', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    MockTestUtils.clearAllMockData();
    
    // Set up realistic timing for tests
    jest.useFakeTimers();
    jest.setSystemTime(new Date('2024-01-15T12:00:00Z'));
  });

  afterEach(() => {
    jest.useRealTimers();
  });

  describe('Complete Data Flow: 読み込み → 列検出 → 処理 → 出力', () => {
    test('典型的なフォーム回答データの完全処理フロー', async () => {
      // 1. Setup realistic spreadsheet data
      MockTestUtils.createRealisticSpreadsheetData(
        PRODUCTION_LIKE_DATA.formResponseHeaders,
        PRODUCTION_LIKE_DATA.formResponseData
      );

      // 2. Simulate spreadsheet access
      const spreadsheetId = 'test-form-responses-spreadsheet';
      const sheetName = 'フォームの回答 1';
      
      MockTestUtils.setMockData('activeSpreadsheetId', spreadsheetId);
      MockTestUtils.setMockData('activeSheetName', sheetName);

      // 3. Test column detection (identifyHeadersAdvanced equivalent)
      const spreadsheet = global.SpreadsheetApp.openById(spreadsheetId);
      const sheet = spreadsheet.getSheetByName(sheetName);
      const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      
      expect(headerRow).toEqual(PRODUCTION_LIKE_DATA.formResponseHeaders);

      // 4. Test AI column mapping
      const detectColumnMapping = (headers: string[]): ColumnMapping => {
        const mapping: ColumnMapping = {
          answer: -1,
          reason: -1,
          class: -1,
          name: -1,
          confidence: { answer: 0, reason: 0, class: 0, name: 0 }
        };

        headers.forEach((header, index) => {
          // Answer column detection
          if (header.includes('意見') || header.includes('どうして') || header.includes('回答')) {
            mapping.answer = index;
            mapping.confidence.answer = 95;
          }
          // Reason column detection  
          if (header.includes('理由') || header.includes('根拠') || header.includes('なぜ')) {
            mapping.reason = index;
            mapping.confidence.reason = 90;
          }
          // Class column detection
          if (header.includes('クラス') || header.includes('学年')) {
            mapping.class = index;
            mapping.confidence.class = 85;
          }
          // Name column detection
          if (header.includes('名前') || header.includes('氏名') || header.includes('お名前')) {
            mapping.name = index;
            mapping.confidence.name = 80;
          }
        });

        return mapping;
      };

      const columnMapping = detectColumnMapping(headerRow);
      
      expect(columnMapping.answer).toBe(2); // "あなたの意見を教えてください"
      expect(columnMapping.reason).toBe(3); // "その理由を教えてください"
      expect(columnMapping.class).toBe(4);  // "あなたのクラスを教えてください"
      expect(columnMapping.name).toBe(5);   // "お名前を教えてください（任意）"
      expect(columnMapping).toHaveValidColumnMapping();

      // 5. Test data processing
      const allData = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
      const dataRows = allData.slice(1); // Skip header row

      const processedData: ProcessedResponse[] = dataRows.map((row, index) => ({
        id: `response-${index + 1}`,
        timestamp: row[0] as string,
        email: row[1] as string,
        answer: row[columnMapping.answer] as string,
        reason: columnMapping.reason >= 0 ? row[columnMapping.reason] as string : undefined,
        class: columnMapping.class >= 0 ? row[columnMapping.class] as string : undefined,
        name: columnMapping.name >= 0 ? row[columnMapping.name] as string : undefined,
        reactions: {
          UNDERSTAND: parseInt(row[6]) || 0,
          LIKE: parseInt(row[7]) || 0,
          CURIOUS: parseInt(row[8]) || 0,
        }
      }));

      expect(processedData).toHaveLength(4);
      expect(processedData[0]).toMatchObject({
        answer: 'AIの発達により、今後は創造性がより重要になると思います',
        reason: '単純作業はAIに任せて、人間は創造的な仕事に集中できるから',
        class: '3年A組',
        name: '田中太郎',
        reactions: { UNDERSTAND: 5, LIKE: 3, CURIOUS: 2 }
      });

      // 6. Test AdminPanel output formatting
      const adminPanelResult: AdminPanelResult = {
        success: true,
        data: processedData,
        stats: {
          totalResponses: processedData.length,
          averageReactions: processedData.reduce((sum, item) => 
            sum + item.reactions.UNDERSTAND + item.reactions.LIKE + item.reactions.CURIOUS, 0
          ) / processedData.length,
          uniqueClasses: [...new Set(processedData.map(item => item.class).filter(Boolean))]
        }
      };

      expect(adminPanelResult.success).toBe(true);
      expect(adminPanelResult.stats!.totalResponses).toBe(4);
      expect(adminPanelResult.stats!.uniqueClasses).toEqual(['3年A組', '2年B組', '1年C組']);
      expect(adminPanelResult.stats!.averageReactions).toBeCloseTo(8.25, 2);
    });

    test('エッジケースデータの処理フロー', () => {
      // Setup edge case data
      MockTestUtils.createRealisticSpreadsheetData(
        PRODUCTION_LIKE_DATA.formResponseHeaders,
        PRODUCTION_LIKE_DATA.edgeCaseData
      );

      const spreadsheet = global.SpreadsheetApp.openById('test-edge-cases');
      const sheet = spreadsheet.getSheetByName('Sheet1');
      const allData = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
      const dataRows = allData.slice(1);

      // Test handling of problematic data
      const processData = (rows: any[][]) => {
        return rows.map((row, index) => {
          // Handle empty/undefined values gracefully
          const safeGet = (arr: any[], idx: number, fallback = '') => 
            arr[idx] !== undefined && arr[idx] !== null ? String(arr[idx]).trim() : fallback;

          return {
            id: `edge-case-${index + 1}`,
            timestamp: safeGet(row, 0),
            email: safeGet(row, 1),
            answer: safeGet(row, 2),
            reason: safeGet(row, 3),
            class: safeGet(row, 4),
            name: safeGet(row, 5),
            reactions: {
              UNDERSTAND: Math.max(0, parseInt(safeGet(row, 6, '0')) || 0),
              LIKE: Math.max(0, parseInt(safeGet(row, 7, '0')) || 0),
              CURIOUS: Math.max(0, parseInt(safeGet(row, 8, '0')) || 0),
            }
          };
        });
      };

      const processedEdgeCases = processData(dataRows);

      // Verify handling of multiline text
      expect(processedEdgeCases[0].answer).toContain('\n');
      expect(processedEdgeCases[0].reason).toContain('\n');

      // Verify handling of very long text (should be preserved)
      expect(processedEdgeCases[1].answer.length).toBeGreaterThan(500);

      // Verify handling of empty fields
      expect(processedEdgeCases[2].reason).toBe('');
      expect(processedEdgeCases[2].class).toBe('');
      expect(processedEdgeCases[2].name).toBe('');
      expect(processedEdgeCases[2].reactions).toEqual({ UNDERSTAND: 0, LIKE: 0, CURIOUS: 0 });
    });
  });

  describe('Error Handling in Data Flow', () => {
    test('スプレッドシートアクセス権限エラーのハンドリング', () => {
      MockTestUtils.simulatePermissionError('spreadsheet_no-permission');

      const testDataFlow = () => {
        try {
          const spreadsheet = global.SpreadsheetApp.openById('no-permission');
          const sheet = spreadsheet.getSheetByName('Sheet1');
          return sheet.getRange(1, 1, 10, 10).getValues();
        } catch (error) {
          return { error: (error as Error).message };
        }
      };

      const result = testDataFlow();
      expect(result).toHaveProperty('error');
      expect((result as any).error).toContain('permission');
    });

    test('列マッピング失敗時のフォールバック処理', () => {
      // Setup data with unexpected headers
      const unexpectedHeaders = [
        'Date', 'User', 'Comment', 'Note', 'Group', 'DisplayName'
      ];
      
      MockTestUtils.createRealisticSpreadsheetData(
        unexpectedHeaders,
        [['2024-01-15', 'user1@example.com', 'Some comment', 'Some note', 'Group1', 'User One']]
      );

      const detectColumnMappingWithFallback = (headers: string[]): ColumnMapping => {
        const mapping: ColumnMapping = {
          answer: -1,
          reason: -1,
          class: -1,
          name: -1,
          confidence: { answer: 0, reason: 0, class: 0, name: 0 }
        };

        // Try exact matches first
        headers.forEach((header, index) => {
          const lowerHeader = header.toLowerCase();
          
          // Fallback to partial matches for international/different formats
          if (lowerHeader.includes('comment') || lowerHeader.includes('opinion') || lowerHeader.includes('answer')) {
            mapping.answer = index;
            mapping.confidence.answer = 60; // Lower confidence for fallback
          }
          if (lowerHeader.includes('note') || lowerHeader.includes('reason') || lowerHeader.includes('why')) {
            mapping.reason = index;
            mapping.confidence.reason = 50;
          }
          if (lowerHeader.includes('group') || lowerHeader.includes('class') || lowerHeader.includes('team')) {
            mapping.class = index;
            mapping.confidence.class = 55;
          }
          if (lowerHeader.includes('name') || lowerHeader.includes('displayname') || lowerHeader.includes('user')) {
            mapping.name = index;
            mapping.confidence.name = 45;
          }
        });

        // If no answer column found, use the first text-like column as fallback
        if (mapping.answer === -1 && headers.length >= 3) {
          mapping.answer = 2; // Assume third column is answer
          mapping.confidence.answer = 30;
        }

        return mapping;
      };

      const spreadsheet = global.SpreadsheetApp.openById('test-unexpected-headers');
      const sheet = spreadsheet.getSheetByName('Sheet1');
      const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      
      const mapping = detectColumnMappingWithFallback(headerRow);
      
      expect(mapping.answer).toBe(2); // 'Comment' column
      expect(mapping.reason).toBe(3); // 'Note' column  
      expect(mapping.class).toBe(4); // 'Group' column
      expect(mapping.confidence.answer).toBe(60); // Lower confidence
    });
  });

  describe('Performance and Scalability', () => {
    test('大量データ処理のパフォーマンステスト', () => {
      // Generate large dataset (1000 responses)
      const largeHeaders = PRODUCTION_LIKE_DATA.formResponseHeaders;
      const largeDataSet = Array.from({ length: 1000 }, (_, i) => [
        `2024/01/15 ${String(Math.floor(i / 60)).padStart(2, '0')}:${String(i % 60).padStart(2, '0')}:00`,
        `student${i}@school.ed.jp`,
        `意見${i}について詳細に述べます。`,
        `理由${i}について説明します。`,
        `${Math.floor(i / 100) + 1}年${String.fromCharCode(65 + (i % 26))}組`,
        `生徒${i}`,
        String(Math.floor(Math.random() * 20)),
        String(Math.floor(Math.random() * 20)), 
        String(Math.floor(Math.random() * 20))
      ]);

      MockTestUtils.createRealisticSpreadsheetData(largeHeaders, largeDataSet);
      
      const startTime = Date.now();
      
      const spreadsheet = global.SpreadsheetApp.openById('test-large-dataset');
      const sheet = spreadsheet.getSheetByName('Sheet1');
      
      // Simulate batch data retrieval (performance optimization)
      const batchSize = 100;
      const totalRows = sheet.getLastRow();
      const processedBatches: any[][] = [];
      
      for (let i = 2; i <= totalRows; i += batchSize) {
        const endRow = Math.min(i + batchSize - 1, totalRows);
        const batchData = sheet.getRange(i, 1, endRow - i + 1, sheet.getLastColumn()).getValues();
        processedBatches.push(batchData);
      }

      const endTime = Date.now();
      const processingTime = endTime - startTime;

      expect(processedBatches.length).toBeGreaterThan(0);
      expect(processingTime).toBeLessThan(5000); // Should complete within 5 seconds
      
      // Verify data integrity in batching
      const totalProcessedRows = processedBatches.reduce((sum, batch) => sum + batch.length, 0);
      expect(totalProcessedRows).toBe(1000);
    });

    test('メモリ効率的なデータ処理', () => {
      const memoryEfficientProcessor = function* (data: any[][], chunkSize = 50) {
        for (let i = 0; i < data.length; i += chunkSize) {
          yield data.slice(i, i + chunkSize);
        }
      };

      const testData = Array.from({ length: 200 }, (_, i) => [
        `data${i}`, `value${i}`, `info${i}`
      ]);

      const chunks: any[][] = [];
      const processor = memoryEfficientProcessor(testData, 50);
      
      for (const chunk of processor) {
        chunks.push(chunk);
        // Simulate processing each chunk
        expect(chunk.length).toBeLessThanOrEqual(50);
      }

      expect(chunks.length).toBe(4); // 200 / 50 = 4 chunks
      
      // Verify all data was processed
      const totalProcessed = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
      expect(totalProcessed).toBe(200);
    });
  });

  describe('Real-world Scenario Testing', () => {
    test('複数クラス混在データの処理', () => {
      const multiClassData = [
        ['2024/01/15 10:00:00', 'student1@school.ed.jp', '環境保護が重要', '地球温暖化対策のため', '1年A組', '山田太郎', '5', '3', '2'],
        ['2024/01/15 10:05:00', 'student2@school.ed.jp', 'デジタル化推進', '効率性向上のため', '2年B組', '佐藤花子', '8', '6', '4'],
        ['2024/01/15 10:10:00', 'student3@school.ed.jp', '国際理解教育', 'グローバル社会対応', '3年C組', '田中次郎', '7', '5', '6'],
        ['2024/01/15 10:15:00', 'teacher1@school.ed.jp', '教育改革必要', '時代に合わせた変化', '教員', '鈴木先生', '10', '8', '5'],
      ];

      MockTestUtils.createRealisticSpreadsheetData(
        PRODUCTION_LIKE_DATA.formResponseHeaders,
        multiClassData
      );

      const spreadsheet = global.SpreadsheetApp.openById('test-multi-class');
      const sheet = spreadsheet.getSheetByName('Sheet1');
      const allData = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
      const dataRows = allData.slice(1);

      // Group by class functionality
      const groupedByClass = dataRows.reduce((groups: Record<string, any[]>, row) => {
        const className = row[4] || 'その他';
        if (!groups[className]) {
          groups[className] = [];
        }
        groups[className].push(row);
        return groups;
      }, {});

      expect(Object.keys(groupedByClass)).toEqual(['1年A組', '2年B組', '3年C組', '教員']);
      expect(groupedByClass['1年A組']).toHaveLength(1);
      expect(groupedByClass['教員']).toHaveLength(1);

      // Statistical analysis
      const classStats = Object.entries(groupedByClass).map(([className, rows]) => ({
        className,
        count: rows.length,
        avgReactions: rows.reduce((sum, row) => 
          sum + (parseInt(row[6]) + parseInt(row[7]) + parseInt(row[8])), 0
        ) / rows.length
      }));

      expect(classStats).toEqual(
        expect.arrayContaining([
          expect.objectContaining({ className: '1年A組', count: 1, avgReactions: 10 }),
          expect.objectContaining({ className: '教員', count: 1, avgReactions: 23 })
        ])
      );
    });
  });
});