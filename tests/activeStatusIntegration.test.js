/**
 * アクティブステータス統合関数のテスト（統合確認）
 */

describe('アクティブステータス統合確認テスト', () => {
  test('統合処理が完了している', () => {
    // 統合関数の存在確認
    const fs = require('fs');
    const path = require('path');

    const mainGsPath = path.join(__dirname, '../src/main.gs');
    const mainContent = fs.readFileSync(mainGsPath, 'utf8');

    // updateUserActiveStatusCore が存在することを確認
    expect(mainContent).toContain('function updateUserActiveStatusCore(');

    // 既存の関数が統合関数を使用していることを確認
    expect(mainContent).toContain('return updateUserActiveStatusCore(targetUserId, isActive, {');

    // 統合パターンの確認
    expect(mainContent).toContain("callerType: 'self'");
    expect(mainContent).toContain("callerType: 'admin'");
  });

  test('Core.gs の関数も統合関数を使用している', () => {
    const fs = require('fs');
    const path = require('path');

    const coreGsPath = path.join(__dirname, '../src/Core.gs');
    const coreContent = fs.readFileSync(coreGsPath, 'utf8');

    // updateIsActiveStatus と updateUserActiveStatusForUI が統合関数を使用
    expect(coreContent).toContain('updateUserActiveStatusCore(');
    expect(coreContent).toContain("callerType: 'api'");
    expect(coreContent).toContain("callerType: 'ui'");
  });

  test('統合前の重複コードが削除されている', () => {
    const fs = require('fs');
    const path = require('path');

    const mainGsPath = path.join(__dirname, '../src/main.gs');
    const coreGsPath = path.join(__dirname, '../src/Core.gs');

    const mainContent = fs.readFileSync(mainGsPath, 'utf8');
    const coreContent = fs.readFileSync(coreGsPath, 'utf8');

    // main.gs: updateUserActiveStatus と updateSelfActiveStatus が簡潔になっている
    const updateUserActiveStatusMatch = mainContent.match(
      /function updateUserActiveStatus\([\s\S]*?\n}/
    );
    if (updateUserActiveStatusMatch) {
      const functionLength = updateUserActiveStatusMatch[0].split('\n').length;
      expect(functionLength).toBeLessThan(10); // 統合前は約40行だったが、統合後は5行程度
    }

    const updateSelfActiveStatusMatch = mainContent.match(
      /function updateSelfActiveStatus\([\s\S]*?\n}/
    );
    if (updateSelfActiveStatusMatch) {
      const functionLength = updateSelfActiveStatusMatch[0].split('\n').length;
      expect(functionLength).toBeLessThan(10); // 統合前は約40行だったが、統合後は5行程度
    }
  });

  test('統一された権限チェック機能', () => {
    const fs = require('fs');
    const path = require('path');

    const mainGsPath = path.join(__dirname, '../src/main.gs');
    const mainContent = fs.readFileSync(mainGsPath, 'utf8');

    // 統一関数に各権限チェックパターンが含まれている
    expect(mainContent).toContain("case 'self':");
    expect(mainContent).toContain("case 'admin':");
    expect(mainContent).toContain("case 'api':");
    expect(mainContent).toContain("case 'ui':");
  });

  test('統一されたエラーハンドリング', () => {
    const fs = require('fs');
    const path = require('path');

    const mainGsPath = path.join(__dirname, '../src/main.gs');
    const mainContent = fs.readFileSync(mainGsPath, 'utf8');

    // 統一関数に適切なエラーハンドリングが含まれている
    expect(mainContent).toContain('有効なユーザーIDが必要です');
    expect(mainContent).toContain('指定されたユーザーが見つかりません');
    expect(mainContent).toContain('updateUserActiveStatusCore error:');
  });

  test('統一されたレスポンス形式', () => {
    const fs = require('fs');
    const path = require('path');

    const mainGsPath = path.join(__dirname, '../src/main.gs');
    const mainContent = fs.readFileSync(mainGsPath, 'utf8');

    // 統一関数のレスポンス形式
    expect(mainContent).toContain('success: true');
    expect(mainContent).toContain('changed: false');
    expect(mainContent).toContain('changed: true');
    expect(mainContent).toContain('timestamp: new Date().toISOString()');
  });
});
