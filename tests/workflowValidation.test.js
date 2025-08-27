/**
 * ワークフロー検証用テスト
 * Claude Code開発環境の動作確認
 */

describe('ワークフロー検証', () => {
  describe('環境検証', () => {
    test('Jestが正常に動作する', () => {
      expect(true).toBe(true);
    });

    test('モック環境が利用可能', () => {
      expect(SpreadsheetApp).toBeDefined();
      expect(PropertiesService).toBeDefined();
      expect(CacheService).toBeDefined();
    });

    test('新機能: ユーティリティ関数の動作確認', () => {
      const result = validateWorkflow('test-data');
      expect(result).toEqual({
        valid: true,
        message: 'Workflow validation successful',
        data: 'test-data'
      });
    });
  });

  describe('GAS連携テスト', () => {
    test('SpreadsheetApp モックが動作する', () => {
      const sheet = SpreadsheetApp.getActiveSheet();
      const range = sheet.getRange(1, 1, 5, 3);
      
      expect(range.getValues).toBeDefined();
      expect(typeof range.getValues()).toBe('object');
    });

    test('PropertiesService モックが動作する', () => {
      const props = PropertiesService.getScriptProperties();
      props.setProperty('test-key', 'test-value');
      
      expect(props.setProperty).toHaveBeenCalledWith('test-key', 'test-value');
    });
  });
});

/**
 * ワークフロー検証用のユーティリティ関数
 * @param {string} data テストデータ
 * @returns {Object} 検証結果
 */
function validateWorkflow(data) {
  if (!data) {
    return {
      valid: false,
      message: 'Data is required',
      data: null
    };
  }

  return {
    valid: true,
    message: 'Workflow validation successful',
    data: data
  };
}

// グローバル関数として定義（テスト用）
global.validateWorkflow = validateWorkflow;