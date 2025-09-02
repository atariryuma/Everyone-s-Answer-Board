/**
 * ヘッダーマッピング修正のテストコード
 */

// モック設定
const COLUMN_HEADERS = {
  REASON: '理由',
  OPINION: '回答',
};

// 未定義関数のモック定義
function validateRequiredHeaders(indices) {
  if (!indices || typeof indices !== 'object') {
    return { success: false, missing: ['すべて'], hasReasonColumn: false, hasOpinionColumn: false };
  }
  
  const hasReason = indices[COLUMN_HEADERS.REASON] !== undefined;
  const hasOpinion = indices[COLUMN_HEADERS.OPINION] !== undefined;
  const missing = [];
  
  if (!hasReason) missing.push(COLUMN_HEADERS.REASON);
  if (!hasOpinion) missing.push(COLUMN_HEADERS.OPINION);
  
  return {
    success: hasReason && hasOpinion,
    missing,
    hasReasonColumn: hasReason,
    hasOpinionColumn: hasOpinion
  };
}

function convertIndicesToMapping(headerIndices, headerRow) {
  if (!headerIndices || !headerRow) {
    return { answer: null, reason: null };
  }
  
  const mapping = {};
  
  // SYSTEM_CONSTANTS.COLUMN_MAPPINGを使用してマッピング作成
  Object.values(SYSTEM_CONSTANTS.COLUMN_MAPPING).forEach((column) => {
    const headerName = column.header;
    const index = Object.keys(headerIndices).find(key => 
      headerIndices[key] !== undefined && headerRow[headerIndices[key]] === headerName
    );
    
    mapping[column.key] = index !== undefined ? headerIndices[index] : null;
  });
  
  return mapping;
}

const SYSTEM_CONSTANTS = {
  COLUMN_MAPPING: {
    answer: {
      key: 'answer',
      header: '回答',
      alternates: ['どうして', '質問', '問題', '意見', '答え'],
      required: true,
    },
    reason: {
      key: 'reason',
      header: '理由',
      alternates: ['理由', '根拠', '体験', 'なぜ'],
      required: false,
    },
    class: {
      key: 'class',
      header: 'クラス',
      alternates: ['クラス', '学年'],
      required: false,
    },
    name: {
      key: 'name',
      header: '名前',
      alternates: ['名前', '氏名', 'お名前'],
      required: false,
    },
  },
};

// validateRequiredHeaders関数のテスト
function testValidateRequiredHeaders() {
  console.log('=== validateRequiredHeaders テスト開始 ===');

  // テストケース1: 質問文がヘッダーになっている場合
  const test1 = {
    'どうして、メダカと一緒に、水草、ミジンコを入れると思いますか？観察していて、気づいたことを書きましょう。': 4,
    'そう考える理由や体験があれば教えてください。': 5,
    'なるほど！': 6,
    'いいね！': 7,
  };

  const result1 = validateRequiredHeaders(test1);
  console.log('Test1 - 質問文ヘッダー:', {
    isValid: result1.isValid,
    hasQuestionColumn: result1.hasQuestionColumn,
    hasReasonColumn: result1.hasReasonColumn,
    hasOpinionColumn: result1.hasOpinionColumn,
  });

  // テストケース2: 固定列名の場合
  const test2 = {
    回答: 1,
    理由: 2,
    クラス: 3,
    名前: 4,
  };

  const result2 = validateRequiredHeaders(test2);
  console.log('Test2 - 固定列名:', {
    isValid: result2.isValid,
    hasQuestionColumn: result2.hasQuestionColumn,
    hasReasonColumn: result2.hasReasonColumn,
    hasOpinionColumn: result2.hasOpinionColumn,
  });

  // テストケース3: 空のインデックス
  const test3 = {};

  const result3 = validateRequiredHeaders(test3);
  console.log('Test3 - 空のインデックス:', {
    isValid: result3.isValid,
    missing: result3.missing,
  });

  console.log('=== validateRequiredHeaders テスト完了 ===\n');
}

// convertIndicesToMapping関数のテスト
function testConvertIndicesToMapping() {
  console.log('=== convertIndicesToMapping テスト開始 ===');

  // テストケース1: 質問文ヘッダーの場合
  const headerIndices1 = {
    タイムスタンプ: 0,
    メールアドレス: 1,
    クラス: 2,
    名前: 3,
    'どうして、メダカと一緒に、水草、ミジンコを入れると思いますか？観察していて、気づいたことを書きましょう。': 4,
    'そう考える理由や体験があれば教えてください。': 5,
    'なるほど！': 6,
    'いいね！': 7,
    'もっと知りたい！': 8,
    ハイライト: 9,
  };

  const headerRow1 = Object.keys(headerIndices1);

  const result1 = convertIndicesToMapping(headerIndices1, headerRow1);
  console.log('Test1 - 質問文ヘッダーマッピング:', result1);

  // テストケース2: 空のheaderIndicesの場合
  const headerIndices2 = {};
  const headerRow2 = [
    'タイムスタンプ',
    'メールアドレス',
    'クラス',
    '名前',
    'どうして、メダカと一緒に、水草、ミジンコを入れると思いますか？',
    'そう考える理由や体験があれば教えてください。',
  ];

  const result2 = convertIndicesToMapping(headerIndices2, headerRow2);
  console.log('Test2 - 空のheaderIndices（AI判定フォールバック）:', result2);

  // テストケース3: nullのheaderIndicesの場合
  const result3 = convertIndicesToMapping(null, headerRow2);
  console.log('Test3 - nullのheaderIndices（AI判定フォールバック）:', result3);

  console.log('=== convertIndicesToMapping テスト完了 ===\n');
}

// 統合テスト
function runIntegrationTest() {
  console.log('=== 統合テスト開始 ===');

  // 実際のエラーケースをシミュレート
  const spreadsheetHeaders = [
    'タイムスタンプ',
    'メールアドレス',
    'クラス',
    '名前',
    'どうして、メダカと一緒に、水草、ミジンコを入れると思いますか？観察していて、気づいたことを書きましょう。',
    'そう考える理由や体験があれば教えてください。',
    'なるほど！',
    'いいね！',
    'もっと知りたい！',
    'ハイライト',
  ];

  console.log('入力ヘッダー:', spreadsheetHeaders);

  // ヘッダーインデックスを作成
  const headerIndices = {};
  spreadsheetHeaders.forEach((header, index) => {
    headerIndices[header] = index;
  });

  // 検証
  const validation = validateRequiredHeaders(headerIndices);
  console.log('検証結果:', validation);

  // マッピング
  const mapping = convertIndicesToMapping(headerIndices, spreadsheetHeaders);
  console.log('マッピング結果:', mapping);

  // 結果の検証
  const success = mapping.answer !== null && mapping.reason !== null;
  console.log(`\n統合テスト結果: ${success ? '✅ 成功' : '❌ 失敗'}`);
  console.log(
    '- answer列:',
    mapping.answer !== null ? `検出（index: ${mapping.answer}）` : '未検出'
  );
  console.log(
    '- reason列:',
    mapping.reason !== null ? `検出（index: ${mapping.reason}）` : '未検出'
  );
  console.log('- class列:', mapping.class !== null ? `検出（index: ${mapping.class}）` : '未検出');
  console.log('- name列:', mapping.name !== null ? `検出（index: ${mapping.name}）` : '未検出');

  console.log('=== 統合テスト完了 ===\n');
}

// テスト実行
console.log('🔧 ヘッダーマッピング修正テスト\n');
testValidateRequiredHeaders();
testConvertIndicesToMapping();
runIntegrationTest();
console.log('✅ すべてのテスト完了');
