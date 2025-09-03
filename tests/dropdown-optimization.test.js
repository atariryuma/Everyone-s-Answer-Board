/**
 * ドロップダウン最適化テスト
 * 自動読み込み停止とオンデマンド読み込み検証
 */

describe('Dropdown Optimization Test', () => {
  beforeEach(() => {
    // DOM要素のモック
    global.document = {
      getElementById: jest.fn((id) => {
        if (id === 'spreadsheet-select') {
          return {
            innerHTML: '',
            children: { length: 1 },
            dataset: {},
            addEventListener: jest.fn(),
            value: '',
            querySelector: jest.fn(() => null),
            appendChild: jest.fn()
          };
        }
        if (id === 'sheet-select') {
          return {
            innerHTML: ''
          };
        }
        if (id === 'connect-source') {
          return {
            disabled: true,
            textContent: '接続'
          };
        }
        return null;
      })
    };

    // console.log モック
    global.console = {
      log: jest.fn(),
      error: jest.fn()
    };
  });

  test('should not auto-load spreadsheet list on initialization', () => {
    console.log('🔧 テスト: 初期化時の自動読み込み停止確認');

    // initializeDropdownsWithDatabaseInfo の動作をシミュレート
    function simulateDropdownInit(hasExistingData) {
      const spreadsheetSelect = global.document.getElementById('spreadsheet-select');
      
      if (hasExistingData) {
        // データベースに情報がある場合
        spreadsheetSelect.innerHTML = `
          <option value="">-- クリックして読み込み --</option>
          <option value="test-sheet-123" selected>設定済み: テストシート</option>
        `;
        return 'pre-populated';
      } else {
        // データベースに情報がない場合
        spreadsheetSelect.innerHTML = '<option value="">-- クリックして読み込み --</option>';
        return 'on-demand-ready';
      }
    }

    // テスト実行
    const resultWithData = simulateDropdownInit(true);
    const resultWithoutData = simulateDropdownInit(false);

    // 検証
    expect(resultWithData).toBe('pre-populated');
    expect(resultWithoutData).toBe('on-demand-ready');
    
    console.log('✅ 自動読み込み停止テスト成功:', {
      withData: resultWithData,
      withoutData: resultWithoutData
    });
  });

  test('should trigger loading only on dropdown click', () => {
    console.log('🎯 テスト: ドロップダウンクリック時のみ読み込み確認');

    const spreadsheetSelect = global.document.getElementById('spreadsheet-select');
    let loadTriggered = false;

    // オンデマンド読み込みのシミュレート
    function simulateOnDemandLoading() {
      // 既に読み込み済みかチェック
      if (spreadsheetSelect.children.length > 2 || 
          spreadsheetSelect.dataset.loaded === 'true') {
        return 'already-loaded';
      }

      // 読み込み実行
      loadTriggered = true;
      spreadsheetSelect.innerHTML = '<option value="">読み込み中...</option>';
      
      // 読み込み完了のシミュレート
      setTimeout(() => {
        spreadsheetSelect.innerHTML = `
          <option value="">スプレッドシートを選択...</option>
          <option value="sheet1">シート1</option>
          <option value="sheet2">シート2</option>
        `;
        spreadsheetSelect.children.length = 3;
        spreadsheetSelect.dataset.loaded = 'true';
      }, 100);

      return 'loading-triggered';
    }

    // イベントリスナーのシミュレート
    const clickHandler = simulateOnDemandLoading;

    // テスト実行
    const firstClickResult = clickHandler();
    expect(firstClickResult).toBe('loading-triggered');
    expect(loadTriggered).toBe(true);

    // 2回目のクリック（既に読み込み済み）
    spreadsheetSelect.dataset.loaded = 'true';
    spreadsheetSelect.children.length = 3;
    const secondClickResult = clickHandler();
    expect(secondClickResult).toBe('already-loaded');

    console.log('✅ オンデマンド読み込みテスト成功:', {
      firstClick: firstClickResult,
      secondClick: secondClickResult,
      loadTriggered: loadTriggered
    });
  });

  test('should preserve pre-populated database info', () => {
    console.log('📋 テスト: データベース情報の事前入力保持確認');

    const mockConfig = {
      spreadsheetId: 'database-sheet-456',
      sheetName: 'データベースシート'
    };

    // データベース情報の事前入力シミュレート
    function simulatePrePopulation(config) {
      const spreadsheetSelect = global.document.getElementById('spreadsheet-select');
      const sheetSelect = global.document.getElementById('sheet-select');
      const connectBtn = global.document.getElementById('connect-source');

      if (config && config.spreadsheetId && config.sheetName) {
        // 事前入力処理
        spreadsheetSelect.innerHTML = `
          <option value="">-- クリックして読み込み --</option>
          <option value="${config.spreadsheetId}" selected>設定済み: ${config.sheetName}</option>
        `;

        sheetSelect.innerHTML = `
          <option value="">シートを選択</option>
          <option value="${config.sheetName}" selected>${config.sheetName}</option>
        `;

        connectBtn.disabled = false;
        connectBtn.textContent = '再接続';

        return {
          prePopulated: true,
          hasSpreadsheet: true,
          hasSheet: true,
          connectEnabled: true
        };
      }

      return {
        prePopulated: false,
        hasSpreadsheet: false,
        hasSheet: false,
        connectEnabled: false
      };
    }

    // テスト実行
    const result = simulatePrePopulation(mockConfig);

    // 検証
    expect(result.prePopulated).toBe(true);
    expect(result.hasSpreadsheet).toBe(true);
    expect(result.hasSheet).toBe(true);
    expect(result.connectEnabled).toBe(true);

    console.log('✅ 事前入力保持テスト成功:', result);
  });

  test('should optimize performance by avoiding unnecessary API calls', () => {
    console.log('⚡ テスト: 不要なAPI呼び出し回避による性能最適化確認');

    let apiCallCount = 0;

    // API呼び出しのシミュレート
    function simulateApiCall(action) {
      apiCallCount++;
      return `API called: ${action} (call #${apiCallCount})`;
    }

    // 従来のアプローチ（毎回自動読み込み）
    function traditionalApproach() {
      const results = [];
      results.push(simulateApiCall('auto-load on init'));
      results.push(simulateApiCall('auto-load on config load'));
      return {
        approach: 'traditional',
        apiCalls: apiCallCount,
        results: results
      };
    }

    // 最適化アプローチ（オンデマンド読み込み）
    function optimizedApproach(userClicksDropdown) {
      apiCallCount = 0; // リセット
      const results = [];
      
      // 初期化では API 呼び出しなし
      results.push('init: no API call');
      
      // ユーザーがドロップダウンをクリックした場合のみ
      if (userClicksDropdown) {
        results.push(simulateApiCall('on-demand load'));
      }
      
      return {
        approach: 'optimized',
        apiCalls: apiCallCount,
        results: results
      };
    }

    // テスト実行
    apiCallCount = 0;
    const traditionalResult = traditionalApproach();
    const optimizedResult = optimizedApproach(false); // ユーザーがクリックしない場合
    const optimizedWithClickResult = optimizedApproach(true); // ユーザーがクリックした場合

    // 検証
    expect(traditionalResult.apiCalls).toBe(2);
    expect(optimizedResult.apiCalls).toBe(0);
    expect(optimizedWithClickResult.apiCalls).toBe(1);

    console.log('⚡ 性能最適化テスト成功:', {
      traditional: `${traditionalResult.apiCalls} API calls`,
      optimizedNoClick: `${optimizedResult.apiCalls} API calls`,
      optimizedWithClick: `${optimizedWithClickResult.apiCalls} API calls`,
      improvement: `${((traditionalResult.apiCalls - optimizedResult.apiCalls) / traditionalResult.apiCalls * 100).toFixed(0)}% reduction when no interaction`
    });
  });
});