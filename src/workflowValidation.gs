/**
 * ワークフロー検証用GAS関数
 * Claude Code開発環境のE2Eテスト用
 */

/**
 * 開発ワークフローの検証を実行
 * @param {string} data 検証データ
 * @return {Object} 検証結果
 */
function validateWorkflowGAS(data) {
  // 統一バリデーション使用
  const validation = UValidate.validateRequired({ data }, ['data']);
  if (!validation.isValid) {
    return {
      valid: false,
      message: validation.errors[0],
      data: null,
      timestamp: new Date().toISOString(),
    };
  }

  // GAS API を使用した実際の処理例
  const properties = PropertiesService.getScriptProperties();
  const cacheKey = `workflow_validation_${Date.now()}`;

  try {
    // プロパティに保存
    properties.setProperty('last_validation', data);

    // キャッシュにも保存（実際のGAS環境でのテスト用）
    const cache = CacheService.getScriptCache();
    cache.put(
      cacheKey,
      JSON.stringify({
        data: data,
        validated: true,
        timestamp: new Date().toISOString(),
      }),
      300
    ); // 5分キャッシュ

    ULog('INFO', 'validateWorkflowGAS', 'Workflow validation completed', { data });

    return {
      valid: true,
      message: 'Workflow validation successful with GAS integration',
      data: data,
      timestamp: new Date().toISOString(),
      cacheKey: cacheKey,
    };
  } catch (error) {
    return UError(
      error,
      'validateWorkflowGAS',
      UNIFIED_CONSTANTS.ERROR.SEVERITY.MEDIUM,
      UNIFIED_CONSTANTS.ERROR.CATEGORIES.VALIDATION,
      { data }
    );
  }
}

/**
 * スプレッドシートを使用したワークフロー検証
 * @return {Object} 検証結果
 */
function validateWorkflowWithSheet() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const testRange = sheet.getRange(1, 1, 1, 3);

    // テストデータを書き込み
    const testData = [['Workflow', 'Validation', new Date().toISOString()]];
    testRange.setValues(testData);

    // 読み込みテスト
    const readData = testRange.getValues();

    ULog('INFO', 'validateWorkflowWithSheet', 'Sheet validation completed', { readData });

    return {
      valid: true,
      message: 'Sheet-based workflow validation successful',
      sheetData: readData[0],
      timestamp: new Date().toISOString(),
    };
  } catch (error) {
    return UError(
      error,
      'validateWorkflowWithSheet',
      UNIFIED_CONSTANTS.ERROR.SEVERITY.MEDIUM,
      UNIFIED_CONSTANTS.ERROR.CATEGORIES.DATABASE
    );
  }
}

/**
 * 包括的なワークフロー検証
 * 複数のGAS APIを使用した統合テスト
 * @return {Object} 総合検証結果
 */
function comprehensiveWorkflowValidation() {
  const results = {
    timestamp: new Date().toISOString(),
    tests: {},
    overall: { passed: 0, total: 0 },
  };

  // 1. プロパティサービステスト
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('test_workflow', 'active');
    const retrieved = props.getProperty('test_workflow');

    results.tests.properties = {
      passed: retrieved === 'active',
      message: retrieved === 'active' ? 'Properties service working' : 'Properties service failed',
    };
  } catch (error) {
    ULog('ERROR', 'comprehensiveWorkflowValidation', 'Properties service test failed', {
      error: error.message,
    });
    results.tests.properties = { passed: false, message: error.message };
  }

  // 2. キャッシュサービステスト
  try {
    const cache = CacheService.getScriptCache();
    const testKey = 'workflow_test';
    cache.put(testKey, 'test_value', 60);
    const cached = cache.get(testKey);

    results.tests.cache = {
      passed: cached === 'test_value',
      message: cached === 'test_value' ? 'Cache service working' : 'Cache service failed',
    };
  } catch (error) {
    ULog('ERROR', 'comprehensiveWorkflowValidation', 'Cache service test failed', {
      error: error.message,
    });
    results.tests.cache = { passed: false, message: error.message };
  }

  // 3. ユーティリティテスト
  try {
    const uuid = Utilities.getUuid();
    const formatted = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

    results.tests.utilities = {
      passed: uuid && formatted,
      message: uuid && formatted ? 'Utilities working' : 'Utilities failed',
      uuid: uuid,
      formatted: formatted,
    };
  } catch (error) {
    ULog('ERROR', 'comprehensiveWorkflowValidation', 'Utilities test failed', {
      error: error.message,
    });
    results.tests.utilities = { passed: false, message: error.message };
  }

  // 結果集計
  results.overall.total = Object.keys(results.tests).length;
  results.overall.passed = Object.values(results.tests).filter((test) => test.passed).length;
  results.overall.success = results.overall.passed === results.overall.total;

  ULog(
    'INFO',
    'comprehensiveWorkflowValidation',
    'Comprehensive workflow validation completed',
    results
  );
  return results;
}
