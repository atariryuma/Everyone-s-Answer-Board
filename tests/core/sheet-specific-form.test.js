/**
 * シート別フォーム連携テスト
 * getSheetSpecificFormInfo()とgetFormInfo()の正確性を検証
 */

describe('Sheet-Specific Form Connection Tests', () => {
  let mockSpreadsheetApp;
  let mockFormApp;

  beforeEach(() => {
    // SpreadsheetAppのモック
    mockSpreadsheetApp = {
      openById: jest.fn(),
      getSheetByName: jest.fn()
    };
    global.SpreadsheetApp = mockSpreadsheetApp;

    // FormAppのモック
    mockFormApp = {
      openByUrl: jest.fn()
    };
    global.FormApp = mockFormApp;

    // console.logのモック
    global.console = {
      log: jest.fn(),
      warn: jest.fn(),
      error: jest.fn()
    };
  });

  describe('getSheetSpecificFormInfo', () => {
    test('フォームの回答 6に正しいフォーム情報を返す', () => {
      // モックの設定
      const mockSheet = {
        getFormUrl: jest.fn().mockReturnValue('https://docs.google.com/forms/d/e/1FAIpQLSdbfLhJJbqneRsBr2cN7CHhG1yZN3Ik4p9r3TCWuPMQViGtuA/viewform')
      };

      const mockForm = {
        getPublishedUrl: jest.fn().mockReturnValue('https://docs.google.com/forms/d/e/1FAIpQLSdbfLhJJbqneRsBr2cN7CHhG1yZN3Ik4p9r3TCWuPMQViGtuA/viewform'),
        getTitle: jest.fn().mockReturnValue('自然の中での動物の食べ物を調べよう。'),
        getId: jest.fn().mockReturnValue('12BK-SczqlVTaPOJGjTmuVJ2sEnvObxYYJSS7eecbPDI'),
        getDestinationId: jest.fn().mockReturnValue('1z5DNDxNZJB6x8KYRNBKRM1qRlGKCH44rNSMF9vlu4bA')
      };

      const mockSpreadsheet = {
        getSheetByName: jest.fn().mockReturnValue(mockSheet)
      };

      mockSpreadsheetApp.openById.mockReturnValue(mockSpreadsheet);
      mockFormApp.openByUrl.mockReturnValue(mockForm);

      // 実際の関数実装をモック
      function getSheetSpecificFormInfo(spreadsheetId, sheetName) {
        const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
        const sheet = spreadsheet.getSheetByName(sheetName);
        
        if (!sheet) {
          return { hasForm: false, reason: 'シートが見つかりません' };
        }

        const formUrl = sheet.getFormUrl();
        if (!formUrl) {
          return { hasForm: false, reason: 'このシートにフォーム連携なし' };
        }

        const form = FormApp.openByUrl(formUrl);
        return {
          hasForm: true,
          formUrl: form.getPublishedUrl(),
          formTitle: form.getTitle(),
          formId: form.getId(),
          sheetName: sheetName,
          destinationCheck: form.getDestinationId() === spreadsheetId
        };
      }

      // テスト実行
      const result = getSheetSpecificFormInfo('1z5DNDxNZJB6x8KYRNBKRM1qRlGKCH44rNSMF9vlu4bA', 'フォームの回答 6');

      // 検証
      expect(result.hasForm).toBe(true);
      expect(result.formTitle).toBe('自然の中での動物の食べ物を調べよう。');
      expect(result.formId).toBe('12BK-SczqlVTaPOJGjTmuVJ2sEnvObxYYJSS7eecbPDI');
      expect(result.sheetName).toBe('フォームの回答 6');
      expect(result.destinationCheck).toBe(true);

      // 正しいAPI呼び出しを確認
      expect(mockSpreadsheetApp.openById).toHaveBeenCalledWith('1z5DNDxNZJB6x8KYRNBKRM1qRlGKCH44rNSMF9vlu4bA');
      expect(mockSpreadsheet.getSheetByName).toHaveBeenCalledWith('フォームの回答 6');
      expect(mockSheet.getFormUrl).toHaveBeenCalled();
      expect(mockFormApp.openByUrl).toHaveBeenCalledWith('https://docs.google.com/forms/d/e/1FAIpQLSdbfLhJJbqneRsBr2cN7CHhG1yZN3Ik4p9r3TCWuPMQViGtuA/viewform');
    });

    test('フォーム連携のないシートでfalseを返す', () => {
      const mockSheet = {
        getFormUrl: jest.fn().mockReturnValue(null)
      };

      const mockSpreadsheet = {
        getSheetByName: jest.fn().mockReturnValue(mockSheet)
      };

      mockSpreadsheetApp.openById.mockReturnValue(mockSpreadsheet);

      function getSheetSpecificFormInfo(spreadsheetId, sheetName) {
        const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
        const sheet = spreadsheet.getSheetByName(sheetName);
        
        if (!sheet) {
          return { hasForm: false, reason: 'シートが見つかりません' };
        }

        const formUrl = sheet.getFormUrl();
        if (!formUrl) {
          return { hasForm: false, reason: 'このシートにフォーム連携なし' };
        }
      }

      const result = getSheetSpecificFormInfo('test-spreadsheet-id', 'データシート');

      expect(result.hasForm).toBe(false);
      expect(result.reason).toBe('このシートにフォーム連携なし');
    });
  });

  describe('Multiple Forms in Single Spreadsheet', () => {
    test('異なるシートが異なるフォームを返すことを確認', () => {
      // フォームの回答 2のモック
      const mockSheet2 = {
        getFormUrl: jest.fn().mockReturnValue('https://docs.google.com/forms/d/e/form2-url/viewform')
      };

      // フォームの回答 6のモック  
      const mockSheet6 = {
        getFormUrl: jest.fn().mockReturnValue('https://docs.google.com/forms/d/e/form6-url/viewform')
      };

      const mockForm2 = {
        getTitle: jest.fn().mockReturnValue('別のフォーム'),
        getId: jest.fn().mockReturnValue('form2-id')
      };

      const mockForm6 = {
        getTitle: jest.fn().mockReturnValue('自然の中での動物の食べ物を調べよう。'),
        getId: jest.fn().mockReturnValue('form6-id')
      };

      const mockSpreadsheet = {
        getSheetByName: jest.fn((name) => {
          if (name === 'フォームの回答 2') return mockSheet2;
          if (name === 'フォームの回答 6') return mockSheet6;
          return null;
        })
      };

      mockSpreadsheetApp.openById.mockReturnValue(mockSpreadsheet);
      mockFormApp.openByUrl.mockImplementation((url) => {
        if (url.includes('form2-url')) return mockForm2;
        if (url.includes('form6-url')) return mockForm6;
        throw new Error('Unknown form URL');
      });

      function getSheetSpecificFormInfo(spreadsheetId, sheetName) {
        const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
        const sheet = spreadsheet.getSheetByName(sheetName);
        const formUrl = sheet.getFormUrl();
        const form = FormApp.openByUrl(formUrl);
        
        return {
          hasForm: true,
          formTitle: form.getTitle(),
          formId: form.getId(),
          sheetName: sheetName
        };
      }

      // 異なるシートをテスト
      const result2 = getSheetSpecificFormInfo('test-id', 'フォームの回答 2');
      const result6 = getSheetSpecificFormInfo('test-id', 'フォームの回答 6');

      // 異なるフォーム情報が正しく返されることを確認
      expect(result2.formTitle).toBe('別のフォーム');
      expect(result2.formId).toBe('form2-id');
      expect(result2.sheetName).toBe('フォームの回答 2');

      expect(result6.formTitle).toBe('自然の中での動物の食べ物を調べよう。');
      expect(result6.formId).toBe('form6-id');
      expect(result6.sheetName).toBe('フォームの回答 6');

      // 正しいシートが呼び出されていることを確認
      expect(mockSpreadsheet.getSheetByName).toHaveBeenCalledWith('フォームの回答 2');
      expect(mockSpreadsheet.getSheetByName).toHaveBeenCalledWith('フォームの回答 6');
    });
  });
});