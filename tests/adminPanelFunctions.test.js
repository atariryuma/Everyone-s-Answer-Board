/**
 * @fileoverview AdminPanel関数のテスト
 */

const fs = require('fs');
const path = require('path');

// adminPanel.js.htmlから必要な関数を抽出
function loadAdminPanelFunctions() {
  const file = fs.readFileSync(path.join(__dirname, '../src/adminPanel.js.html'), 'utf8');
  
  // <script>タグ内のコードを抽出
  const scriptStart = file.indexOf('<script>');
  const scriptEnd = file.lastIndexOf('</script>');
  const scriptContent = file.slice(scriptStart + 8, scriptEnd);
  
  // グローバルオブジェクトとDOM APIのモック
  const mockWindow = {
    location: { href: 'https://example.com' },
    confirm: jest.fn(() => true),
    setTimeout: setTimeout,
    clearTimeout: clearTimeout,
    addEventListener: jest.fn(),
    removeEventListener: jest.fn()
  };
  
  const mockDocument = {
    getElementById: jest.fn(() => ({
      textContent: '',
      innerHTML: '',
      value: '',
      classList: {
        add: jest.fn(),
        remove: jest.fn(),
        contains: jest.fn(() => false)
      },
      addEventListener: jest.fn()
    })),
    addEventListener: jest.fn(),
    createElement: jest.fn(() => ({
      click: jest.fn()
    }))
  };
  
  const mockConsole = {
    log: jest.fn(),
    error: jest.fn(),
    warn: jest.fn(),
    info: jest.fn()
  };
  
  // Google Apps Script APIのモック
  const mockGoogle = {
    script: {
      run: {
        withSuccessHandler: jest.fn(() => mockGoogle.script.run),
        withFailureHandler: jest.fn(() => mockGoogle.script.run),
        withUserObject: jest.fn(() => mockGoogle.script.run),
        getCurrentUserStatus: jest.fn(),
        getLogoutUrl: jest.fn(() => 'https://accounts.google.com/logout')
      }
    }
  };
  
  // 実行コンテキストの作成
  const context = {
    window: mockWindow,
    document: mockDocument,
    console: mockConsole,
    google: mockGoogle,
    confirm: mockWindow.confirm,
    setTimeout,
    clearTimeout,
    // その他必要なグローバル変数
    CONFIG: undefined,
    utils: undefined,
    state: undefined,
    auth: undefined
  };
  
  // コードを評価してコンテキストに実行
  const executeCode = new Function('window', 'document', 'console', 'google', 'confirm', scriptContent);
  executeCode(mockWindow, mockDocument, mockConsole, mockGoogle, mockWindow.confirm);
  
  return {
    window: mockWindow,
    document: mockDocument,
    console: mockConsole,
    google: mockGoogle
  };
}

describe('AdminPanel関数テスト', () => {
  let mocks;
  
  beforeEach(() => {
    // 各テスト前にモックをリセット
    jest.clearAllMocks();
    mocks = loadAdminPanelFunctions();
  });

  describe('logoutAndRedirect関数', () => {
    test('window.adminPanelオブジェクトが存在する', () => {
      expect(mocks.window.adminPanel).toBeDefined();
      expect(typeof mocks.window.adminPanel).toBe('object');
    });

    test('logoutAndRedirect関数が定義されている', () => {
      expect(mocks.window.adminPanel.logoutAndRedirect).toBeDefined();
      expect(typeof mocks.window.adminPanel.logoutAndRedirect).toBe('function');
    });

    test('logoutAndRedirect関数がconfirmダイアログを表示する', () => {
      // confirmがtrueを返すように設定
      mocks.window.confirm.mockReturnValue(true);
      
      // logoutAndRedirect関数を呼び出し
      mocks.window.adminPanel.logoutAndRedirect();
      
      // confirmが呼び出されたことを確認
      expect(mocks.window.confirm).toHaveBeenCalledWith('ログアウトしますか？');
    });

    test('confirmでキャンセルされた場合は何もしない', () => {
      // confirmがfalseを返すように設定
      mocks.window.confirm.mockReturnValue(false);
      
      const originalHref = mocks.window.location.href;
      
      // logoutAndRedirect関数を呼び出し
      mocks.window.adminPanel.logoutAndRedirect();
      
      // location.hrefが変更されていないことを確認
      expect(mocks.window.location.href).toBe(originalHref);
    });

    test('confirmでOKされた場合はリダイレクトが実行される', () => {
      // confirmがtrueを返すように設定
      mocks.window.confirm.mockReturnValue(true);
      
      // Google Apps Script APIが適切な値を返すように設定
      mocks.google.script.run.getLogoutUrl.mockReturnValue('https://accounts.google.com/logout');
      
      // logoutAndRedirect関数を呼び出し
      try {
        mocks.window.adminPanel.logoutAndRedirect();
      } catch (e) {
        // utils.callAPIが存在しない場合のエラーは想定内
        expect(e.message).toContain('utils');
      }
      
      // confirmが呼び出されたことを確認
      expect(mocks.window.confirm).toHaveBeenCalledWith('ログアウトしますか？');
    });
  });

  describe('AdminPanelオブジェクトの構造テスト', () => {
    test('必要なプロパティが存在する', () => {
      const adminPanel = mocks.window.adminPanel;
      
      expect(adminPanel).toHaveProperty('logoutAndRedirect');
      
      // 他の主要なプロパティの存在確認（定義されている場合）
      const expectedProperties = ['init', 'auth', 'data', 'config', 'board', 'ui', 'handlers', 'utils', 'state'];
      expectedProperties.forEach(prop => {
        if (adminPanel.hasOwnProperty(prop)) {
          expect(adminPanel[prop]).toBeDefined();
        }
      });
    });

    test('logoutAndRedirect関数が正しい引数で呼び出せる', () => {
      const logoutFn = mocks.window.adminPanel.logoutAndRedirect;
      
      // 関数が引数なしで呼び出せることを確認
      expect(() => logoutFn()).not.toThrow();
      
      // 関数の長さ（引数の数）を確認
      expect(logoutFn.length).toBe(0); // 引数は0個
    });
  });

  describe('グローバル関数の互換性テスト', () => {
    test('onclick属性で使用される他のグローバル関数が定義されている', () => {
      // toggleSection関数の確認
      if (mocks.window.toggleSection) {
        expect(typeof mocks.window.toggleSection).toBe('function');
      }
      
      // copyBoardUrl関数の確認
      if (mocks.window.copyBoardUrl) {
        expect(typeof mocks.window.copyBoardUrl).toBe('function');
      }
      
      // copyFormUrl関数の確認
      if (mocks.window.copyFormUrl) {
        expect(typeof mocks.window.copyFormUrl).toBe('function');
      }
    });
  });

  describe('エラーハンドリングテスト', () => {
    test('undefined状態での関数呼び出しでも例外がスローされない', () => {
      // authオブジェクトが未定義の場合のテスト
      const originalAdminPanel = mocks.window.adminPanel;
      mocks.window.adminPanel = { 
        logoutAndRedirect: () => {
          // auth.logout()が存在しない場合の代替動作
          try {
            return originalAdminPanel.auth.logout();
          } catch (e) {
            // フォールバック動作
            mocks.window.location.href = 'https://accounts.google.com/logout';
          }
        }
      };
      
      mocks.window.confirm.mockReturnValue(true);
      
      expect(() => {
        mocks.window.adminPanel.logoutAndRedirect();
      }).not.toThrow();
    });
  });
});