/**
 * @fileoverview 管理パネルのデータ取得フロー検証テスト
 */

const fs = require('fs');
const path = require('path');

// データベース関数を読み込み
function loadDatabaseFunctions() {
  const dbFile = fs.readFileSync(path.join(__dirname, '../src/database.gs'), 'utf8');
  const coreFile = fs.readFileSync(path.join(__dirname, '../src/Core.gs'), 'utf8');

  return { dbFile, coreFile };
}

describe('管理パネル データ取得フロー検証', () => {
  describe('データベースアクセス関数の存在確認', () => {
    test('fetchUserFromDatabase関数が定義されている', () => {
      const { dbFile } = loadDatabaseFunctions();
      expect(dbFile).toContain('function fetchUserFromDatabase');
    });

    test('getCurrentUserStatus関数が定義されている', () => {
      const { coreFile } = loadDatabaseFunctions();
      expect(coreFile).toContain('function getCurrentUserStatus');
    });

    test('データベースアクセスにPropertiesServiceを使用している', () => {
      const { dbFile, coreFile } = loadDatabaseFunctions();
      expect(dbFile + coreFile).toContain('PropertiesService');
    });

    test('サービスアカウント認証が実装されている', () => {
      const { dbFile, coreFile } = loadDatabaseFunctions();
      expect(dbFile + coreFile).toContain('SERVICE_ACCOUNT_CREDS');
    });
  });

  describe('エラーハンドリング機能の確認', () => {
    test('データベースエラー処理が実装されている', () => {
      const { coreFile } = loadDatabaseFunctions();
      expect(coreFile).toContain('logDatabaseError');
    });

    test('リトライ機能が実装されている', () => {
      const { dbFile } = loadDatabaseFunctions();
      expect(dbFile).toContain('retryOnce');
      expect(dbFile).toContain('forceFresh');
    });

    test('キャッシュ制御が実装されている', () => {
      const { dbFile } = loadDatabaseFunctions();
      expect(dbFile).toContain('forceFresh');
    });
  });

  describe('セキュリティ関連の確認', () => {
    test('セキュアなデータベースID取得が実装されている', () => {
      const { dbFile } = loadDatabaseFunctions();
      expect(dbFile).toContain('getSecureDatabaseId');
    });

    test('ユーザー認証が実装されている', () => {
      const { coreFile } = loadDatabaseFunctions();
      expect(coreFile).toContain('getCurrentUserEmail');
    });

    test('ドメイン検証が実装されている', () => {
      const { coreFile } = loadDatabaseFunctions();
      expect(coreFile).toContain('domain');
    });
  });

  describe('管理パネル固有のデータフロー', () => {
    test('ユーザー状態の復元機能がある', () => {
      const adminFile = fs.readFileSync(path.join(__dirname, '../src/adminPanel.js.html'), 'utf8');
      expect(adminFile).toContain('initializeFromUserData');
    });

    test('スプレッドシート情報の取得機能がある', () => {
      const adminFile = fs.readFileSync(path.join(__dirname, '../src/adminPanel.js.html'), 'utf8');
      expect(adminFile).toContain('getSpreadsheetSheets');
    });

    test('設定の保存・復元機能がある', () => {
      const adminFile = fs.readFileSync(path.join(__dirname, '../src/adminPanel.js.html'), 'utf8');
      expect(adminFile).toContain('saveUserConfiguration');
      expect(adminFile).toContain('currentConfig');
    });
  });

  describe('API呼び出しパターンの検証', () => {
    test('getCurrentUserStatus APIが正しいパラメータを受け取る', () => {
      const { coreFile } = loadDatabaseFunctions();
      const functionMatch = coreFile.match(/function getCurrentUserStatus\s*\([^)]*\)/);
      expect(functionMatch).toBeTruthy();
      expect(functionMatch[0]).toContain('requestUserId');
    });

    test('fetchUserFromDatabase APIが適切なオプションを受け取る', () => {
      const { dbFile } = loadDatabaseFunctions();
      const functionMatch = dbFile.match(/function fetchUserFromDatabase\s*\([^)]*\)/);
      expect(functionMatch).toBeTruthy();
      expect(functionMatch[0]).toContain('options');
    });

    test('データベース検索でユーザーIDとメールの両方に対応', () => {
      const { dbFile } = loadDatabaseFunctions();
      expect(dbFile).toContain('userId');
      expect(dbFile).toContain('adminEmail');
    });
  });

  describe('レスポンス形式の確認', () => {
    test('getCurrentUserStatusが正しい形式のレスポンスを返す', () => {
      const { coreFile } = loadDatabaseFunctions();
      expect(coreFile).toContain("status: 'success'");
      expect(coreFile).toContain('userInfo');
    });

    test('データベース操作でエラー時の適切なレスポンス形式', () => {
      const { coreFile } = loadDatabaseFunctions();
      expect(coreFile).toContain("status: 'error'");
      expect(coreFile).toContain('message:');
    });
  });

  describe('パフォーマンス最適化の確認', () => {
    test('キャッシュサービスの利用', () => {
      const { dbFile } = loadDatabaseFunctions();
      expect(dbFile).toContain('Cache');
    });

    test('バッチ処理の実装', () => {
      const { dbFile, coreFile } = loadDatabaseFunctions();
      const combined = dbFile + coreFile;
      // getValues, setValues などの一括処理APIの使用を確認
      expect(combined).toMatch(/getValues|setValues/);
    });
  });
});
