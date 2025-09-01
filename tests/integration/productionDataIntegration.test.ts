/**
 * プロダクションデータ統合テスト
 * 実際のプロダクション形式データを使用した包括的なシステムテスト
 */

import { MockTestUtils } from '../mocks/gasMocks';
import ProductionDataManager, { PRODUCTION_DATA_FIXTURES } from '../fixtures/productionDataFixtures';

// Test-specific types
interface ProcessedResponseData {
  id: string;
  timestamp: Date;
  email: string;
  answer: string;
  reason?: string;
  class?: string;
  name?: string;
  reactions: {
    understand: number;
    like: number;
    curious: number;
  };
  metadata: {
    wordCount: number;
    hasEmptyFields: boolean;
    isAnonymous: boolean;
  };
}

interface SystemPerformanceMetrics {
  dataProcessingTime: number;
  memoryUsage: number;
  errorRate: number;
  throughput: number; // responses per second
}

describe('プロダクションデータ統合テスト', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    MockTestUtils.clearAllMockData();
    
    // Set realistic test timing
    jest.useFakeTimers();
    jest.setSystemTime(new Date('2024-01-15T12:00:00Z'));
  });

  afterEach(() => {
    jest.useRealTimers();
  });

  describe('学校アンケートデータ処理', () => {
    test('2024年度学校アンケートの完全処理フロー', () => {
      // Setup production-like school survey data
      ProductionDataManager.setupProductionEnvironment('SCHOOL_SURVEY_2024');
      
      const spreadsheet = global.SpreadsheetApp.openById('test-spreadsheet');
      const sheet = spreadsheet.getSheetByName('Sheet1');
      
      // Verify data structure matches production
      const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const expectedHeaders = PRODUCTION_DATA_FIXTURES.SCHOOL_SURVEY_2024.headers;
      
      expect(headerRow).toEqual(expectedHeaders);
      expect(headerRow[0]).toBe('タイムスタンプ');
      expect(headerRow[2]).toContain('将来の夢');
      
      // Process all responses
      const allData = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
      const dataRows = allData.slice(1); // Skip header
      
      expect(dataRows).toHaveLength(6); // 6 students
      
      // Generate column mapping using production logic
      const columnMapping = ProductionDataManager.generateColumnMapping(headerRow);
      expect(columnMapping.answer).toBe(2); // "将来の夢について"
      expect(columnMapping.reason).toBe(3); // "なぜそう思うのか"
      expect(columnMapping.class).toBe(4);  // "あなたのクラス"
      
      // Process data with production-like logic
      const processedData: ProcessedResponseData[] = dataRows.map((row, index) => {
        const reactions = (columnMapping.reactions as number[]).reduce((acc, colIndex, i) => {
          const reactionKeys = ['understand', 'like', 'curious'] as const;
          acc[reactionKeys[i] || 'understand'] = parseInt(row[colIndex]) || 0;
          return acc;
        }, { understand: 0, like: 0, curious: 0 });

        return {
          id: `school_response_${index + 1}`,
          timestamp: new Date(row[columnMapping.timestamp]),
          email: row[columnMapping.email],
          answer: row[columnMapping.answer],
          reason: columnMapping.reason >= 0 ? row[columnMapping.reason] : undefined,
          class: columnMapping.class >= 0 ? row[columnMapping.class] : undefined,
          name: columnMapping.name >= 0 ? row[columnMapping.name] : undefined,
          reactions,
          metadata: {
            wordCount: (row[columnMapping.answer] || '').length,
            hasEmptyFields: !row[columnMapping.reason] || !row[columnMapping.class],
            isAnonymous: !row[columnMapping.name] || row[columnMapping.name].trim() === ''
          }
        };
      });

      // Validate processed data quality
      expect(processedData.every(item => item.answer.length > 0)).toBe(true);
      expect(processedData.every(item => item.email.includes('@'))).toBe(true);
      expect(processedData.every(item => typeof item.reactions.understand === 'number')).toBe(true);

      // Statistical analysis
      const totalReactions = processedData.reduce((sum, item) => 
        sum + item.reactions.understand + item.reactions.like + item.reactions.curious, 0);
      const averageReactions = totalReactions / processedData.length;
      
      expect(averageReactions).toBeGreaterThan(10); // School data shows engagement
      expect(averageReactions).toBeLessThan(200); // Reasonable upper limit

      // Class distribution analysis
      const classDistribution = processedData.reduce((dist: Record<string, number>, item) => {
        if (item.class) {
          dist[item.class] = (dist[item.class] || 0) + 1;
        }
        return dist;
      }, {});

      expect(Object.keys(classDistribution)).toContain('3年A組');
      expect(Object.keys(classDistribution)).toContain('2年C組');
    });

    test('大量学校データの処理パフォーマンス', () => {
      // Generate large dataset (1000 responses)
      const largeDataset = ProductionDataManager.generateLargeDataset(1000, 'SCHOOL_SURVEY_2024');
      
      expect(ProductionDataManager.validateDataIntegrity(largeDataset)).toBe(true);
      
      MockTestUtils.createRealisticSpreadsheetData(largeDataset.headers, largeDataset.rows);
      
      const startTime = Date.now();
      
      // Simulate batch processing
      const batchSize = 100;
      const batches: any[][] = [];
      
      for (let i = 0; i < largeDataset.rows.length; i += batchSize) {
        const batch = largeDataset.rows.slice(i, i + batchSize);
        batches.push(batch);
      }
      
      const endTime = Date.now();
      const processingTime = endTime - startTime;

      expect(batches).toHaveLength(10); // 1000 / 100 = 10 batches
      expect(processingTime).toBeLessThan(2000); // Should complete within 2 seconds
      
      // Memory efficiency check
      const totalProcessed = batches.reduce((sum, batch) => sum + batch.length, 0);
      expect(totalProcessed).toBe(1000);
    });
  });

  describe('企業研修フィードバック処理', () => {
    test('企業研修データの匿名化処理', () => {
      ProductionDataManager.setupProductionEnvironment('CORPORATE_TRAINING_2024');
      
      const spreadsheet = global.SpreadsheetApp.openById('test-spreadsheet');
      const sheet = spreadsheet.getSheetByName('Sheet1');
      const allData = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
      const headerRow = allData[0];
      const dataRows = allData.slice(1);

      // Corporate data often requires anonymization
      const anonymizeResponse = (row: any[], columnMapping: Record<string, number>) => {
        const anonymized = [...row];
        
        // Remove or hash email for privacy
        if (columnMapping.email >= 0) {
          const email = row[columnMapping.email];
          anonymized[columnMapping.email] = email.replace(/^([^@])([^@]*?)([^@])(@.+)$/, '$1***$3$4');
        }
        
        // Remove or anonymize name
        if (columnMapping.name >= 0) {
          anonymized[columnMapping.name] = '匿名';
        }
        
        return anonymized;
      };

      const columnMapping = ProductionDataManager.generateColumnMapping(headerRow);
      const anonymizedData = dataRows.map(row => anonymizeResponse(row, columnMapping));

      // Verify anonymization
      expect(anonymizedData.every(row => 
        row[columnMapping.email].includes('***')
      )).toBe(true);
      
      expect(anonymizedData.every(row => 
        row[columnMapping.name] === '匿名'
      )).toBe(true);

      // Verify data utility is preserved
      expect(anonymizedData.every(row => 
        row[columnMapping.answer] && row[columnMapping.answer].length > 0
      )).toBe(true);
    });
  });

  describe('国際学会フィードバック処理', () => {
    test('多言語データの処理', () => {
      ProductionDataManager.setupProductionEnvironment('ACADEMIC_CONFERENCE_2024');
      
      const spreadsheet = global.SpreadsheetApp.openById('test-spreadsheet');
      const sheet = spreadsheet.getSheetByName('Sheet1');
      const allData = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
      const headerRow = allData[0];
      const dataRows = allData.slice(1);

      // International data has mixed languages
      const detectLanguage = (text: string): string => {
        if (/[\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FAF]/.test(text)) return 'ja';
        if (/[a-zA-Z]/.test(text)) return 'en';
        return 'unknown';
      };

      const languageDistribution = dataRows.reduce((dist: Record<string, number>, row) => {
        const columnMapping = ProductionDataManager.generateColumnMapping(headerRow);
        const answerText = row[columnMapping.answer] || '';
        const lang = detectLanguage(answerText);
        dist[lang] = (dist[lang] || 0) + 1;
        return dist;
      }, {});

      expect(languageDistribution).toHaveProperty('en');
      expect(languageDistribution).toHaveProperty('ja');
    });
  });

  describe('Edge Case データ処理', () => {
    test('問題のあるデータの堅牢な処理', () => {
      ProductionDataManager.setupEdgeCaseData();
      
      const spreadsheet = global.SpreadsheetApp.openById('test-spreadsheet');
      const sheet = spreadsheet.getSheetByName('Sheet1');
      const allData = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
      const dataRows = allData.slice(1);

      // Robust data processing
      const processRobustly = (row: any[]) => {
        const safeGetString = (value: any): string => {
          if (value === null || value === undefined) return '';
          return String(value).trim();
        };

        const safeGetNumber = (value: any): number => {
          if (value === null || value === undefined || value === '') return 0;
          const num = Number(value);
          if (isNaN(num) || !isFinite(num)) return 0;
          return Math.max(0, Math.min(999, Math.floor(num))); // Clamp to reasonable range
        };

        return {
          timestamp: safeGetString(row[0]),
          email: safeGetString(row[1]),
          answer: safeGetString(row[2]),
          reason: safeGetString(row[3]),
          class: safeGetString(row[4]),
          name: safeGetString(row[5]),
          reactions: {
            understand: safeGetNumber(row[6]),
            like: safeGetNumber(row[7]),
            curious: safeGetNumber(row[8])
          }
        };
      };

      const robustResults = dataRows.map(row => processRobustly(row));

      // Verify robust processing handles edge cases
      expect(robustResults).toHaveLength(4); // All problematic rows processed
      
      // Check empty data handling
      expect(robustResults[0].answer).toBe('');
      expect(robustResults[0].reactions.understand).toBe(0);
      
      // Check very long text handling (preserved but manageable)
      expect(robustResults[1].answer.length).toBeGreaterThan(1000);
      expect(typeof robustResults[1].answer).toBe('string');
      
      // Check security issue handling (input preserved but can be sanitized later)
      expect(robustResults[2].answer).toContain('<script>');
      expect(robustResults[2].reactions.understand).toBe(0); // NaN converted to 0
      
      // Check unicode handling
      expect(robustResults[3].answer).toContain('مرحبا');
      expect(robustResults[3].reason).toContain('中文');
      expect(robustResults[3].class).toContain('한국어');
    });
  });

  describe('システム統合パフォーマンステスト', () => {
    test('複数データセット同時処理', async () => {
      const datasets = ['SCHOOL_SURVEY_2024', 'CORPORATE_TRAINING_2024', 'ACADEMIC_CONFERENCE_2024'] as const;
      
      const processDataset = async (datasetName: typeof datasets[number]) => {
        ProductionDataManager.setupProductionEnvironment(datasetName);
        
        const spreadsheet = global.SpreadsheetApp.openById(`test-${datasetName.toLowerCase()}`);
        const sheet = spreadsheet.getSheetByName('Sheet1');
        const data = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
        
        return {
          dataset: datasetName,
          rowCount: data.length - 1, // Exclude header
          columnCount: data[0].length,
          processedAt: new Date().toISOString()
        };
      };

      const startTime = Date.now();
      const results = await Promise.all(datasets.map(dataset => processDataset(dataset)));
      const endTime = Date.now();

      expect(results).toHaveLength(3);
      expect(results.every(result => result.rowCount > 0)).toBe(true);
      expect(endTime - startTime).toBeLessThan(3000); // Parallel processing should be fast
    });

    test('メモリ効率的な大量データ処理', () => {
      // Test with progressively larger datasets
      const dataSizes = [100, 500, 1000, 2000];
      const performanceMetrics: SystemPerformanceMetrics[] = [];

      dataSizes.forEach(size => {
        const startTime = Date.now();
        const startMemory = process.memoryUsage().heapUsed;

        // Generate and process large dataset
        const largeDataset = ProductionDataManager.generateLargeDataset(size, 'SCHOOL_SURVEY_2024');
        MockTestUtils.createRealisticSpreadsheetData(largeDataset.headers, largeDataset.rows);

        // Simulate memory-efficient processing
        const processInChunks = (data: any[][], chunkSize = 50) => {
          let processedCount = 0;
          let errorCount = 0;

          for (let i = 0; i < data.length; i += chunkSize) {
            const chunk = data.slice(i, i + chunkSize);
            
            try {
              // Simulate processing
              chunk.forEach(row => {
                if (row.length > 0) processedCount++;
              });
            } catch (error) {
              errorCount++;
            }
          }

          return { processedCount, errorCount };
        };

        const result = processInChunks(largeDataset.rows);
        
        const endTime = Date.now();
        const endMemory = process.memoryUsage().heapUsed;

        const metrics: SystemPerformanceMetrics = {
          dataProcessingTime: endTime - startTime,
          memoryUsage: endMemory - startMemory,
          errorRate: result.errorCount / result.processedCount,
          throughput: result.processedCount / ((endTime - startTime) / 1000)
        };

        performanceMetrics.push(metrics);

        // Performance assertions
        expect(metrics.dataProcessingTime).toBeLessThan(size * 2); // Linear time complexity
        expect(metrics.errorRate).toBe(0); // No errors expected
        expect(metrics.throughput).toBeGreaterThan(50); // At least 50 responses/second
      });

      // Verify scalability
      const timeComplexities = performanceMetrics.map((metric, i) => 
        metric.dataProcessingTime / dataSizes[i]
      );

      // Time complexity should remain relatively constant (good scalability)
      const avgComplexity = timeComplexities.reduce((sum, complexity) => sum + complexity, 0) / timeComplexities.length;
      const complexityVariance = timeComplexities.reduce((sum, complexity) => sum + Math.pow(complexity - avgComplexity, 2), 0) / timeComplexities.length;
      
      expect(complexityVariance).toBeLessThan(avgComplexity * 0.5); // Less than 50% variance
    });
  });

  describe('データ品質保証テスト', () => {
    test('データ整合性自動検証', () => {
      Object.entries(PRODUCTION_DATA_FIXTURES).forEach(([datasetName, dataset]) => {
        const spreadsheetData = {
          headers: dataset.headers,
          rows: dataset.responses,
          metadata: {
            totalResponses: dataset.responses.length,
            dateRange: {
              start: '2024-01-01T00:00:00Z',
              end: '2024-12-31T23:59:59Z'
            },
            schools: ['Test School'],
            classes: ['Test Class']
          }
        };

        const isValid = ProductionDataManager.validateDataIntegrity(spreadsheetData);
        expect(isValid).toBe(true);

        // Verify each row has correct number of columns
        dataset.responses.forEach((row, index) => {
          expect(row).toHaveLength(dataset.headers.length);
        });
      });
    });

    test('異常データの自動検出', () => {
      const anomalyDetector = (data: any[][], headers: string[]) => {
        const anomalies: Array<{row: number, column: number, issue: string}> = [];
        
        data.forEach((row, rowIndex) => {
          row.forEach((cell, colIndex) => {
            const header = headers[colIndex];
            
            // Check for suspiciously long text
            if (typeof cell === 'string' && cell.length > 1000) {
              anomalies.push({
                row: rowIndex,
                column: colIndex,
                issue: 'Suspiciously long text'
              });
            }
            
            // Check for potential injection attempts
            if (typeof cell === 'string' && (
              cell.includes('<script>') ||
              cell.includes('DROP TABLE') ||
              cell.includes('javascript:')
            )) {
              anomalies.push({
                row: rowIndex,
                column: colIndex,
                issue: 'Potential security threat'
              });
            }
            
            // Check for numeric anomalies in reaction columns
            if (header && (header.includes('なるほど') || header.includes('いいね') || header.includes('知りたい'))) {
              const num = Number(cell);
              if (num > 1000) {
                anomalies.push({
                  row: rowIndex,
                  column: colIndex,
                  issue: 'Unusually high reaction count'
                });
              }
            }
          });
        });
        
        return anomalies;
      };

      ProductionDataManager.setupEdgeCaseData();
      const spreadsheet = global.SpreadsheetApp.openById('test-spreadsheet');
      const sheet = spreadsheet.getSheetByName('Sheet1');
      const allData = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
      const headers = allData[0];
      const dataRows = allData.slice(1);

      const detectedAnomalies = anomalyDetector(dataRows, headers);
      
      // Should detect the problematic data we set up
      expect(detectedAnomalies.length).toBeGreaterThan(0);
      expect(detectedAnomalies.some(anomaly => anomaly.issue === 'Suspiciously long text')).toBe(true);
      expect(detectedAnomalies.some(anomaly => anomaly.issue === 'Potential security threat')).toBe(true);
      expect(detectedAnomalies.some(anomaly => anomaly.issue === 'Unusually high reaction count')).toBe(true);
    });
  });
});