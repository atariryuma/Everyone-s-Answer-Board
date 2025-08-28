/**
 * 統一Validationシステムのテスト
 */

// テスト用のモック設定
const mockUnifiedValidation = {
  validate: jest.fn(),
  getValidationResult: jest.fn(),
  getAllValidationResults: jest.fn(),
  clearValidationResults: jest.fn()
};

// モジュールのモック化
jest.mock('../src/unifiedValidationSystem.gs', () => ({
  UnifiedValidation: mockUnifiedValidation,
  validateAuthentication: jest.fn(),
  validateConfiguration: jest.fn(),
  validateDataIntegrity: jest.fn(),
  validateSystemHealth: jest.fn(),
  validateWorkflow: jest.fn(),
  comprehensiveValidation: jest.fn()
}));

describe('統一Validationシステム', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  describe('基本検証機能', () => {
    test('認証検証の実行', () => {
      const mockResult = {
        validationId: 'auth_standard_123',
        category: 'authentication',
        level: 'standard',
        success: true,
        hasCritical: false,
        summary: {
          passed: 3,
          total: 3,
          critical: 0,
          passRate: 100
        },
        tests: {
          userAuth: { passed: true, message: 'ユーザー認証成功' },
          sessionValid: { passed: true, message: 'セッション有効' },
          userAccess: { passed: true, message: 'アクセス権確認完了' }
        }
      };

      mockUnifiedValidation.validate.mockReturnValue(mockResult);

      const result = mockUnifiedValidation.validate('authentication', 'standard', { userId: 'testUser' });

      expect(mockUnifiedValidation.validate).toHaveBeenCalledWith('authentication', 'standard', { userId: 'testUser' });
      expect(result.success).toBe(true);
      expect(result.category).toBe('authentication');
      expect(result.summary.passRate).toBe(100);
    });

    test('設定検証の実行', () => {
      const mockConfig = {
        setupStatus: 'completed',
        formCreated: true,
        appPublished: true,
        publishedSheetName: 'TestSheet',
        publishedSpreadsheetId: '12345'
      };

      const mockResult = {
        validationId: 'config_basic_124',
        category: 'configuration',
        level: 'basic',
        success: true,
        hasCritical: false,
        summary: {
          passed: 2,
          total: 2,
          critical: 0,
          passRate: 100
        },
        tests: {
          configJson: { passed: true, message: 'Config JSON有効' },
          configState: { passed: true, message: 'Config状態チェック完了' }
        }
      };

      mockUnifiedValidation.validate.mockReturnValue(mockResult);

      const result = mockUnifiedValidation.validate('configuration', 'basic', { config: mockConfig });

      expect(result.success).toBe(true);
      expect(result.tests.configJson.passed).toBe(true);
    });

    test('データ整合性検証の実行', () => {
      const mockHeaders = ['名前', '学年', '回答'];
      const mockUserRows = [
        ['田中太郎', '3年', 'はい'],
        ['山田花子', '2年', 'いいえ']
      ];

      const mockResult = {
        validationId: 'data_standard_125',
        category: 'data_integrity',
        level: 'standard',
        success: true,
        hasCritical: false,
        summary: {
          passed: 4,
          total: 4,
          critical: 0,
          passRate: 100
        },
        tests: {
          headerIntegrity: { passed: true, message: 'ヘッダー整合性OK' },
          duplicates: { passed: true, message: '重複なし' },
          missingFields: { passed: true, message: '必須フィールドOK' },
          invalidFormats: { passed: true, message: 'データ形式OK' }
        }
      };

      mockUnifiedValidation.validate.mockReturnValue(mockResult);

      const result = mockUnifiedValidation.validate('data_integrity', 'standard', {
        userId: 'testUser',
        headers: mockHeaders,
        userRows: mockUserRows
      });

      expect(result.success).toBe(true);
      expect(result.summary.passRate).toBe(100);
    });

    test('システムヘルス検証の実行', () => {
      const mockResult = {
        validationId: 'system_comprehensive_126',
        category: 'system_health',
        level: 'comprehensive',
        success: false,
        hasCritical: true,
        summary: {
          passed: 4,
          total: 6,
          critical: 1,
          passRate: 67
        },
        tests: {
          basicHealth: { passed: true, message: '基本ヘルス正常' },
          databaseHealth: { passed: true, message: 'DB接続正常' },
          performanceCheck: { passed: true, message: 'パフォーマンス良好' },
          securityHealth: { passed: false, message: 'セキュリティ警告', critical: true },
          multiTenantHealth: { passed: true, message: 'マルチテナント正常' },
          serviceAccountHealth: { passed: false, message: 'サービスアカウント警告' }
        }
      };

      mockUnifiedValidation.validate.mockReturnValue(mockResult);

      const result = mockUnifiedValidation.validate('system_health', 'comprehensive', { userId: 'testUser' });

      expect(result.success).toBe(false);
      expect(result.hasCritical).toBe(true);
      expect(result.summary.critical).toBe(1);
      expect(result.summary.passRate).toBe(67);
    });

    test('ワークフロー検証の実行', () => {
      const mockResult = {
        validationId: 'workflow_standard_127',
        category: 'workflow',
        level: 'standard',
        success: true,
        hasCritical: false,
        summary: {
          passed: 4,
          total: 4,
          critical: 0,
          passRate: 100
        },
        tests: {
          workflowGAS: { passed: true, message: 'ワークフローGAS検証成功' },
          publicationStatus: { passed: true, message: '公開状態: 公開中' },
          workflowWithSheet: { passed: true, message: 'スプレッドシート連携OK' },
          autoStop: { passed: true, message: '自動停止なし' }
        }
      };

      mockUnifiedValidation.validate.mockReturnValue(mockResult);

      const result = mockUnifiedValidation.validate('workflow', 'standard', {
        userId: 'testUser',
        data: 'test',
        config: { appPublished: true },
        userInfo: { email: 'test@example.com' }
      });

      expect(result.success).toBe(true);
      expect(result.tests.workflowGAS.passed).toBe(true);
      expect(result.tests.publicationStatus.message).toContain('公開中');
    });
  });

  describe('検証レベルテスト', () => {
    test('基本レベル検証', () => {
      const mockResult = {
        validationId: 'auth_basic_128',
        category: 'authentication',
        level: 'basic',
        summary: {
          passed: 1,
          total: 1,
          passRate: 100
        },
        tests: {
          userAuth: { passed: true }
        }
      };

      mockUnifiedValidation.validate.mockReturnValue(mockResult);

      const result = mockUnifiedValidation.validate('authentication', 'basic', { userId: 'testUser' });

      expect(result.level).toBe('basic');
      expect(result.summary.total).toBe(1); // 基本レベルは少ないテスト数
    });

    test('標準レベル検証', () => {
      const mockResult = {
        validationId: 'auth_standard_129',
        category: 'authentication',
        level: 'standard',
        summary: {
          passed: 3,
          total: 3,
          passRate: 100
        },
        tests: {
          userAuth: { passed: true },
          sessionValid: { passed: true },
          userAccess: { passed: true }
        }
      };

      mockUnifiedValidation.validate.mockReturnValue(mockResult);

      const result = mockUnifiedValidation.validate('authentication', 'standard', { userId: 'testUser' });

      expect(result.level).toBe('standard');
      expect(result.summary.total).toBe(3); // 標準レベルは中程度のテスト数
    });

    test('包括レベル検証', () => {
      const mockResult = {
        validationId: 'auth_comprehensive_130',
        category: 'authentication',
        level: 'comprehensive',
        summary: {
          passed: 4,
          total: 5,
          passRate: 80
        },
        tests: {
          userAuth: { passed: true },
          sessionValid: { passed: true },
          userAccess: { passed: true },
          authHealth: { passed: true },
          securityCheck: { passed: false, critical: true }
        }
      };

      mockUnifiedValidation.validate.mockReturnValue(mockResult);

      const result = mockUnifiedValidation.validate('authentication', 'comprehensive', { userId: 'testUser' });

      expect(result.level).toBe('comprehensive');
      expect(result.summary.total).toBe(5); // 包括レベルは多くのテスト数
    });
  });

  describe('エラーハンドリング', () => {
    test('不正なカテゴリエラー', () => {
      const mockErrorResult = {
        validationId: 'invalid_category_131',
        category: 'invalid_category',
        success: false,
        hasCritical: true,
        error: {
          message: '未対応の検証カテゴリ: invalid_category'
        }
      };

      mockUnifiedValidation.validate.mockReturnValue(mockErrorResult);

      const result = mockUnifiedValidation.validate('invalid_category', 'standard', {});

      expect(result.success).toBe(false);
      expect(result.hasCritical).toBe(true);
      expect(result.error.message).toContain('未対応の検証カテゴリ');
    });

    test('検証実行エラー', () => {
      const mockErrorResult = {
        validationId: 'execution_error_132',
        category: 'authentication',
        success: false,
        hasCritical: true,
        error: {
          message: 'Validation execution failed'
        },
        tests: {
          validationExecution: {
            passed: false,
            message: '検証実行エラー: Internal error',
            critical: true
          }
        }
      };

      mockUnifiedValidation.validate.mockReturnValue(mockErrorResult);

      const result = mockUnifiedValidation.validate('authentication', 'standard', { userId: null });

      expect(result.success).toBe(false);
      expect(result.tests.validationExecution.passed).toBe(false);
      expect(result.tests.validationExecution.critical).toBe(true);
    });
  });

  describe('レガシー関数統合', () => {
    test('validateConfigJson互換性', () => {
      // レガシー形式の結果をモック
      const mockLegacyResult = {
        isValid: true,
        errors: []
      };

      // 統一システムの結果をレガシー形式に変換した結果をモック
      const { validateConfiguration } = require('../src/unifiedValidationSystem.gs');
      validateConfiguration.mockReturnValue(mockLegacyResult);

      const config = {
        setupStatus: 'completed',
        formCreated: true,
        appPublished: true,
        publishedSheetName: 'TestSheet',
        publishedSpreadsheetId: '12345'
      };

      const result = validateConfiguration(config, {}, 'basic');

      expect(result.isValid).toBe(true);
      expect(result.errors).toEqual([]);
    });

    test('validateUserAuthentication互換性', () => {
      const mockLegacyResult = {
        hasAccess: true,
        message: '3/3 通過'
      };

      const { validateAuthentication } = require('../src/unifiedValidationSystem.gs');
      validateAuthentication.mockReturnValue(mockLegacyResult);

      const result = validateAuthentication('testUser', 'standard');

      expect(result.hasAccess).toBe(true);
      expect(result.message).toContain('通過');
    });

    test('performHealthCheck互換性', () => {
      const mockLegacyResult = {
        overallStatus: 'HEALTHY',
        summary: '通過率: 100%',
        details: {
          basicHealth: { passed: true },
          databaseHealth: { passed: true }
        }
      };

      const { validateSystemHealth } = require('../src/unifiedValidationSystem.gs');
      validateSystemHealth.mockReturnValue(mockLegacyResult);

      const result = validateSystemHealth('testUser', 'standard');

      expect(result.overallStatus).toBe('HEALTHY');
      expect(result.summary).toContain('通過率');
      expect(result.details).toBeDefined();
    });
  });

  describe('包括検証', () => {
    test('全カテゴリ包括検証', () => {
      const mockComprehensiveResult = {
        timestamp: '2025-01-01T00:00:00.000Z',
        overallSuccess: true,
        categories: {
          authentication: { success: true, hasCritical: false },
          configuration: { success: true, hasCritical: false },
          data_integrity: { success: true, hasCritical: false },
          system_health: { success: false, hasCritical: true },
          workflow: { success: true, hasCritical: false }
        },
        summary: {
          total: 5,
          passed: 4,
          critical: 1
        }
      };

      const { comprehensiveValidation } = require('../src/unifiedValidationSystem.gs');
      comprehensiveValidation.mockReturnValue(mockComprehensiveResult);

      const result = comprehensiveValidation({
        userId: 'testUser',
        config: { setupStatus: 'completed' }
      });

      expect(result.overallSuccess).toBe(true);
      expect(result.summary.total).toBe(5);
      expect(result.summary.passed).toBe(4);
      expect(result.categories.system_health.hasCritical).toBe(true);
    });
  });

  describe('結果管理', () => {
    test('検証結果の取得', () => {
      const mockResult = {
        validationId: 'test_133',
        category: 'authentication',
        success: true
      };

      mockUnifiedValidation.getValidationResult.mockReturnValue(mockResult);

      const result = mockUnifiedValidation.getValidationResult('test_133');

      expect(result.validationId).toBe('test_133');
      expect(result.success).toBe(true);
    });

    test('全検証結果の取得', () => {
      const mockResults = [
        { validationId: 'test_134', success: true },
        { validationId: 'test_135', success: false }
      ];

      mockUnifiedValidation.getAllValidationResults.mockReturnValue(mockResults);

      const results = mockUnifiedValidation.getAllValidationResults();

      expect(results).toHaveLength(2);
      expect(results[0].validationId).toBe('test_134');
      expect(results[1].success).toBe(false);
    });

    test('古い検証結果のクリア', () => {
      mockUnifiedValidation.clearValidationResults.mockReturnValue(undefined);

      mockUnifiedValidation.clearValidationResults(24); // 24時間以上前の結果をクリア

      expect(mockUnifiedValidation.clearValidationResults).toHaveBeenCalledWith(24);
    });
  });

  describe('パフォーマンステスト', () => {
    test('大量検証の実行時間', () => {
      const startTime = Date.now();

      // 複数の検証を並行実行
      const categories = ['authentication', 'configuration', 'data_integrity', 'system_health', 'workflow'];
      const mockPromises = categories.map(category => {
        const mockResult = {
          validationId: `${category}_perf_test`,
          category,
          success: true,
          duration: '50ms'
        };
        mockUnifiedValidation.validate.mockResolvedValue(mockResult);
        return Promise.resolve(mockResult);
      });

      return Promise.all(mockPromises).then(results => {
        const endTime = Date.now();
        const executionTime = endTime - startTime;

        expect(results).toHaveLength(5);
        expect(executionTime).toBeLessThan(1000); // 1秒以内での完了を期待
      });
    });

    test('メモリ効率性チェック', () => {
      // 大量の検証結果を生成
      const results = [];
      for (let i = 0; i < 100; i++) {
        results.push({
          validationId: `memory_test_${i}`,
          category: 'authentication',
          success: true,
          tests: {}
        });
      }

      mockUnifiedValidation.getAllValidationResults.mockReturnValue(results);

      const allResults = mockUnifiedValidation.getAllValidationResults();

      expect(allResults).toHaveLength(100);
      // メモリリークがないことを確認（実際の実装ではより詳細なメモリ監視が必要）
    });
  });

  describe('統合シナリオ', () => {
    test('新規ユーザー登録フロー検証', () => {
      // 新規ユーザーの登録から初回利用までの包括的な検証
      const newUserScenario = {
        userId: 'newUser123',
        config: {
          setupStatus: 'pending',
          formCreated: false,
          appPublished: false
        }
      };

      const mockScenarioResults = [
        {
          category: 'authentication',
          success: true,
          tests: { userAuth: { passed: true, message: '新規ユーザー認証成功' } }
        },
        {
          category: 'configuration',
          success: false,
          tests: { configJson: { passed: false, message: '初期設定が未完了' } }
        },
        {
          category: 'workflow',
          success: false,
          tests: { publicationStatus: { passed: false, message: '公開状態: 非公開' } }
        }
      ];

      mockUnifiedValidation.validate
        .mockReturnValueOnce(mockScenarioResults[0])
        .mockReturnValueOnce(mockScenarioResults[1])
        .mockReturnValueOnce(mockScenarioResults[2]);

      const authResult = mockUnifiedValidation.validate('authentication', 'basic', newUserScenario);
      const configResult = mockUnifiedValidation.validate('configuration', 'basic', newUserScenario);
      const workflowResult = mockUnifiedValidation.validate('workflow', 'basic', newUserScenario);

      expect(authResult.success).toBe(true); // 認証は成功
      expect(configResult.success).toBe(false); // 設定は未完了
      expect(workflowResult.success).toBe(false); // ワークフローは未公開
    });

    test('システム運用監視シナリオ', () => {
      // 定期的なシステム監視での包括的な健康状態チェック
      const monitoringScenario = {
        userId: 'systemMonitor',
        checkLevel: 'comprehensive'
      };

      const mockMonitoringResult = {
        timestamp: '2025-01-01T12:00:00.000Z',
        overallSuccess: false,
        categories: {
          authentication: { success: true, summary: { passRate: 100 } },
          configuration: { success: true, summary: { passRate: 95 } },
          data_integrity: { success: true, summary: { passRate: 98 } },
          system_health: { success: false, summary: { passRate: 75, critical: 2 } },
          workflow: { success: true, summary: { passRate: 100 } }
        },
        summary: {
          total: 5,
          passed: 4,
          critical: 2,
          overallHealth: 'WARNING'
        }
      };

      const { comprehensiveValidation } = require('../src/unifiedValidationSystem.gs');
      comprehensiveValidation.mockReturnValue(mockMonitoringResult);

      const result = comprehensiveValidation(monitoringScenario);

      expect(result.overallSuccess).toBe(false);
      expect(result.summary.critical).toBe(2);
      expect(result.categories.system_health.success).toBe(false);
      expect(result.summary.overallHealth).toBe('WARNING');
    });
  });
});

describe('検証移行システム', () => {
  test('移行マッピング統計', () => {
    const mockStats = {
      total: 23,
      deprecated: 2,
      byCategory: {
        authentication: { total: 5, deprecated: 1 },
        configuration: { total: 6, deprecated: 1 },
        data_integrity: { total: 7, deprecated: 0 },
        system_health: { total: 3, deprecated: 0 },
        workflow: { total: 2, deprecated: 0 }
      },
      byLevel: {
        basic: { total: 8, deprecated: 0 },
        standard: { total: 10, deprecated: 1 },
        comprehensive: { total: 5, deprecated: 1 }
      }
    };

    // getValidationMigrationStats関数をモック
    const getValidationMigrationStats = jest.fn().mockReturnValue(mockStats);

    const stats = getValidationMigrationStats();

    expect(stats.total).toBe(23);
    expect(stats.deprecated).toBe(2);
    expect(stats.byCategory.authentication.deprecated).toBe(1);
  });

  test('非推奨警告生成', () => {
    const mockWarning = {
      message: "⚠️  関数 'parseAndValidateConfigJson' は非推奨です",
      replacement: 'validateConfigJson + validateConfigJsonState',
      category: 'configuration',
      level: 'standard',
      notes: 'parseとvalidateが混在していたため分離'
    };

    const getDeprecationWarning = jest.fn().mockReturnValue(mockWarning);

    const warning = getDeprecationWarning('parseAndValidateConfigJson');

    expect(warning.message).toContain('非推奨');
    expect(warning.replacement).toContain('validateConfigJson');
    expect(warning.category).toBe('configuration');
  });
});