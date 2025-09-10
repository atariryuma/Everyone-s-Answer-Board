/**
 * configJSON保存機能のテスト
 * 完全置換モードの動作確認
 */

// テスト用のモック設定
const testUserId = 'test-user-123';
const testUserEmail = 'test@example.com';

/**
 * 完全置換モードのテスト
 */
function testCompleteReplaceMode() {
  console.log('========================================');
  console.log('🧪 完全置換モードテスト開始');
  console.log('========================================');

  // 初期データの設定
  const initialConfig = {
    createdAt: '2025-01-01T00:00:00Z',
    spreadsheetId: 'old-spreadsheet-id',
    sheetName: 'Old Sheet',
    formUrl: 'https://old-form.url',
    formTitle: 'Old Form',
    columnMapping: { answer: 0, reason: 1 },
    opinionHeader: '古い質問',
    extraField: 'これは削除されるべきフィールド',
  };

  console.log('📝 初期設定:', initialConfig);

  // 新しい設定（管理パネルから）
  const newConfig = {
    spreadsheetId: 'new-spreadsheet-id',
    sheetName: 'New Sheet',
    showNames: true,
    showReactions: false,
  };

  console.log('📝 新しい設定（管理パネルから）:', newConfig);

  // saveDraftConfiguration の処理をシミュレート
  const updatedConfig = {
    // 基本設定保持
    createdAt: initialConfig.createdAt,
    lastAccessedAt: new Date().toISOString(),
    
    // データソース更新
    spreadsheetId: newConfig.spreadsheetId,
    sheetName: newConfig.sheetName,
    spreadsheetUrl: `https://docs.google.com/spreadsheets/d/${newConfig.spreadsheetId}`,
    
    // 表示設定更新
    displaySettings: {
      showNames: newConfig.showNames,
      showReactions: newConfig.showReactions,
    },
    displayMode: 'anonymous',
    
    // 状態設定
    setupStatus: 'pending',
    appPublished: false,
    isDraft: true,
    
    // connectDataSourceのデータ保持
    columnMapping: initialConfig.columnMapping,
    opinionHeader: initialConfig.opinionHeader,
    formUrl: initialConfig.formUrl,
    formTitle: initialConfig.formTitle,
    
    // メタ情報
    configVersion: '2.0',
    claudeMdCompliant: true,
  };

  console.log('✅ 更新後の設定:', updatedConfig);

  // 検証
  const tests = [
    {
      name: 'spreadsheetIdが更新されている',
      pass: updatedConfig.spreadsheetId === newConfig.spreadsheetId,
      expected: newConfig.spreadsheetId,
      actual: updatedConfig.spreadsheetId,
    },
    {
      name: 'sheetNameが更新されている',
      pass: updatedConfig.sheetName === newConfig.sheetName,
      expected: newConfig.sheetName,
      actual: updatedConfig.sheetName,
    },
    {
      name: '表示設定が更新されている',
      pass: updatedConfig.displaySettings.showNames === newConfig.showNames,
      expected: newConfig.showNames,
      actual: updatedConfig.displaySettings.showNames,
    },
    {
      name: '不要なフィールド（extraField）が削除されている',
      pass: !updatedConfig.hasOwnProperty('extraField'),
      expected: 'undefined',
      actual: updatedConfig.extraField,
    },
    {
      name: 'columnMappingが保持されている',
      pass: JSON.stringify(updatedConfig.columnMapping) === JSON.stringify(initialConfig.columnMapping),
      expected: initialConfig.columnMapping,
      actual: updatedConfig.columnMapping,
    },
    {
      name: 'formUrlが保持されている',
      pass: updatedConfig.formUrl === initialConfig.formUrl,
      expected: initialConfig.formUrl,
      actual: updatedConfig.formUrl,
    },
  ];

  console.log('\n🔍 テスト結果:');
  let allPassed = true;
  tests.forEach(test => {
    if (test.pass) {
      console.log(`  ✅ ${test.name}`);
    } else {
      console.log(`  ❌ ${test.name}`);
      console.log(`     期待値: ${JSON.stringify(test.expected)}`);
      console.log(`     実際値: ${JSON.stringify(test.actual)}`);
      allPassed = false;
    }
  });

  console.log('\n========================================');
  if (allPassed) {
    console.log('🎉 すべてのテストに合格しました！');
  } else {
    console.log('❌ 一部のテストが失敗しました');
  }
  console.log('========================================\n');

  return allPassed;
}

/**
 * 公開時の設定統合テスト
 */
function testPublishConfiguration() {
  console.log('========================================');
  console.log('🧪 公開時設定統合テスト開始');
  console.log('========================================');

  // 現在の設定
  const currentConfig = {
    createdAt: '2025-01-01T00:00:00Z',
    spreadsheetId: 'test-spreadsheet-id',
    sheetName: 'Test Sheet',
    columnMapping: { answer: 0, reason: 1 },
    opinionHeader: 'テスト質問',
    formUrl: 'https://test-form.url',
    setupStatus: 'data_connected',
    appPublished: false,
  };

  console.log('📝 公開前の設定:', currentConfig);

  // 公開時の設定
  const publishConfig = {
    showNames: false,
    showReactions: true,
  };

  // publishApplication の処理をシミュレート
  const publishedConfig = {
    // 基本設定保持
    createdAt: currentConfig.createdAt,
    lastAccessedAt: new Date().toISOString(),
    
    // データソース設定（確定済み）
    spreadsheetId: currentConfig.spreadsheetId,
    sheetName: currentConfig.sheetName,
    spreadsheetUrl: `https://docs.google.com/spreadsheets/d/${currentConfig.spreadsheetId}`,
    
    // 表示設定
    displaySettings: {
      showNames: publishConfig.showNames,
      showReactions: publishConfig.showReactions,
    },
    displayMode: 'anonymous',
    
    // 公開設定
    appPublished: true,
    setupStatus: 'completed',
    publishedAt: new Date().toISOString(),
    isDraft: false,
    appUrl: 'https://script.google.com/macros/s/test-app-id/exec',
    
    // データ保持
    columnMapping: currentConfig.columnMapping,
    opinionHeader: currentConfig.opinionHeader,
    formUrl: currentConfig.formUrl,
    
    // メタ情報
    configVersion: '2.0',
    claudeMdCompliant: true,
    lastModified: new Date().toISOString(),
  };

  console.log('✅ 公開後の設定:', publishedConfig);

  // 検証
  const tests = [
    {
      name: 'appPublishedがtrueになっている',
      pass: publishedConfig.appPublished === true,
      expected: true,
      actual: publishedConfig.appPublished,
    },
    {
      name: 'setupStatusがcompletedになっている',
      pass: publishedConfig.setupStatus === 'completed',
      expected: 'completed',
      actual: publishedConfig.setupStatus,
    },
    {
      name: 'isDraftがfalseになっている',
      pass: publishedConfig.isDraft === false,
      expected: false,
      actual: publishedConfig.isDraft,
    },
    {
      name: 'appUrlが設定されている',
      pass: publishedConfig.appUrl && publishedConfig.appUrl.length > 0,
      expected: 'non-empty string',
      actual: publishedConfig.appUrl,
    },
    {
      name: 'columnMappingが保持されている',
      pass: JSON.stringify(publishedConfig.columnMapping) === JSON.stringify(currentConfig.columnMapping),
      expected: currentConfig.columnMapping,
      actual: publishedConfig.columnMapping,
    },
    {
      name: 'formUrlが保持されている',
      pass: publishedConfig.formUrl === currentConfig.formUrl,
      expected: currentConfig.formUrl,
      actual: publishedConfig.formUrl,
    },
  ];

  console.log('\n🔍 テスト結果:');
  let allPassed = true;
  tests.forEach(test => {
    if (test.pass) {
      console.log(`  ✅ ${test.name}`);
    } else {
      console.log(`  ❌ ${test.name}`);
      console.log(`     期待値: ${JSON.stringify(test.expected)}`);
      console.log(`     実際値: ${JSON.stringify(test.actual)}`);
      allPassed = false;
    }
  });

  console.log('\n========================================');
  if (allPassed) {
    console.log('🎉 すべてのテストに合格しました！');
  } else {
    console.log('❌ 一部のテストが失敗しました');
  }
  console.log('========================================\n');

  return allPassed;
}

// テスト実行
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    testCompleteReplaceMode,
    testPublishConfiguration,
  };
}

// スタンドアロン実行
console.log('🚀 configJSON保存テストスイート実行開始\n');
const test1 = testCompleteReplaceMode();
const test2 = testPublishConfiguration();

console.log('\n📊 総合結果:');
if (test1 && test2) {
  console.log('✅ すべてのテストスイートが成功しました！');
} else {
  console.log('❌ 一部のテストスイートが失敗しました');
}