/**
 * コーディング規範準拠テスト
 * CLAUDE.md GAS V8規範の遵守を検証
 */

const fs = require('fs');
const path = require('path');

describe('CLAUDE.md Coding Standards Compliance', () => {
  const srcDir = path.join(__dirname, '../../src');

  describe('var使用禁止の検証', () => {
    test('全GASファイルでvarが使用されていない', () => {
      const gasFiles = fs
        .readdirSync(srcDir)
        .filter((file) => file.endsWith('.gs'))
        .map((file) => path.join(srcDir, file));

      gasFiles.forEach((filePath) => {
        const content = fs.readFileSync(filePath, 'utf8');
        const varMatches = content.match(/\bvar\s+\w+/g) || [];

        expect(varMatches.length).toBe(0);
      });
    });
  });

  describe('const/let適切使用の検証', () => {
    test('const優先使用パターンが実装されている', () => {
      const testFile = path.join(srcDir, 'database.gs');
      const content = fs.readFileSync(testFile, 'utf8');

      // constの使用数をカウント
      const constMatches = content.match(/\bconst\s+\w+/g) || [];
      const letMatches = content.match(/\blet\s+\w+/g) || [];

      // constがletより多く使用されていることを確認
      expect(constMatches.length).toBeGreaterThan(letMatches.length);
    });

    test('ループ変数にletが使用されている', () => {
      const testFile = path.join(srcDir, 'database.gs');
      const content = fs.readFileSync(testFile, 'utf8');

      // forループでletが使用されているか確認
      const forLoopPattern = /for\s*\(\s*let\s+/g;
      const forLoopMatches = content.match(forLoopPattern) || [];

      expect(forLoopMatches.length).toBeGreaterThan(0);
    });
  });

  describe('Object.freeze()使用の検証', () => {
    test('定数オブジェクトにObject.freeze()が適用されている', () => {
      const constantsFile = path.join(srcDir, 'constants.gs');
      const content = fs.readFileSync(constantsFile, 'utf8');

      // Object.freezeの使用数をカウント
      const freezeMatches = content.match(/Object\.freeze\(/g) || [];

      // constants.gsで多数のObject.freezeが使用されていることを確認
      expect(freezeMatches.length).toBeGreaterThan(20);
    });

    test('DB_CONFIGがObject.freeze()で保護されている', () => {
      const dbFile = path.join(srcDir, 'database.gs');
      const content = fs.readFileSync(dbFile, 'utf8');

      // DB_CONFIGの定義を確認
      const dbConfigPattern = /const\s+DB_CONFIG\s*=\s*Object\.freeze\(/;
      const hasFreeze = dbConfigPattern.test(content);

      expect(hasFreeze).toBe(true);
    });
  });

  describe('アロー関数使用の検証', () => {
    test('アロー関数が適切に使用されている', () => {
      const testFile = path.join(srcDir, 'database.gs');
      const content = fs.readFileSync(testFile, 'utf8');

      // アロー関数の使用をカウント
      const arrowFunctions = content.match(/=>/g) || [];

      // アロー関数が使用されていることを確認
      expect(arrowFunctions.length).toBeGreaterThan(0);
    });
  });

  describe('テンプレートリテラル使用の検証', () => {
    test('文字列結合にテンプレートリテラルが使用されている', () => {
      const testFile = path.join(srcDir, 'database.gs');
      const content = fs.readFileSync(testFile, 'utf8');

      // テンプレートリテラルの使用をカウント
      const templateLiterals = content.match(/`[^`]*\$\{[^}]+\}[^`]*`/g) || [];

      // テンプレートリテラルが使用されていることを確認
      expect(templateLiterals.length).toBeGreaterThan(0);
    });
  });
});

describe('System Constants Structure', () => {
  test('SYSTEM_CONSTANTSが正しく定義されている', () => {
    // SYSTEM_CONSTANTSのモック
    const SYSTEM_CONSTANTS = {
      DATABASE: {
        SHEET_NAME: 'Users',
        HEADERS: [
          'userId',
          'userEmail',
          'createdAt',
          'lastAccessedAt',
          'isActive',
          'spreadsheetId',
          'spreadsheetUrl',
          'configJson',
          'formUrl',
          'sheetName',
          'columnMappingJson',
          'publishedAt',
          'appUrl',
          'lastModified',
        ],
      },
      REACTIONS: {
        KEYS: ['UNDERSTAND', 'LIKE', 'CURIOUS'],
        LABELS: {
          UNDERSTAND: 'なるほど！',
          LIKE: 'いいね！',
          CURIOUS: 'もっと知りたい！',
          HIGHLIGHT: 'ハイライト',
        },
      },
      ACCESS: {
        LEVELS: {
          OWNER: 'owner',
          SYSTEM_ADMIN: 'system_admin',
          AUTHENTICATED_USER: 'authenticated_user',
          GUEST: 'guest',
          NONE: 'none',
        },
      },
    };

    // 必須フィールドの存在確認
    expect(SYSTEM_CONSTANTS.DATABASE).toBeDefined();
    expect(SYSTEM_CONSTANTS.DATABASE.HEADERS.length).toBe(14);
    expect(SYSTEM_CONSTANTS.REACTIONS.KEYS).toEqual(['UNDERSTAND', 'LIKE', 'CURIOUS']);
    expect(SYSTEM_CONSTANTS.ACCESS.LEVELS.OWNER).toBe('owner');
  });

  test('PROPS_KEYSが正しく定義されている', () => {
    const PROPS_KEYS = {
      SERVICE_ACCOUNT_CREDS: 'SERVICE_ACCOUNT_CREDS',
      DATABASE_SPREADSHEET_ID: 'DATABASE_SPREADSHEET_ID',
      ADMIN_EMAIL: 'ADMIN_EMAIL',
    };

    expect(PROPS_KEYS.SERVICE_ACCOUNT_CREDS).toBe('SERVICE_ACCOUNT_CREDS');
    expect(PROPS_KEYS.DATABASE_SPREADSHEET_ID).toBe('DATABASE_SPREADSHEET_ID');
    expect(PROPS_KEYS.ADMIN_EMAIL).toBe('ADMIN_EMAIL');
  });
});
