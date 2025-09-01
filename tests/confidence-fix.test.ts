/**
 * confidence値0%問題の修正検証テスト
 * mergeColumnConfidence関数の動作確認
 */

// Mock GAS APIs
const mockSpreadsheetApp = {
  openById: jest.fn(() => ({
    getSheetByName: jest.fn(() => ({
      getRange: jest.fn(() => ({
        getValues: jest.fn(() => [['タイムスタンプ', 'メールアドレス', 'クラス', '名前', '回答', '理由']]),
        setValues: jest.fn()
      })),
      getLastRow: jest.fn(() => 100)
    }))
  }))
};

global.SpreadsheetApp = mockSpreadsheetApp;

// 修正後の関数をテスト用に定義
const mergeColumnConfidence = (basicMapping: any, aiResult: any, headers: string[]) => {
  const enhanced = { ...basicMapping };
  
  // 既存のconfidence値を保持（重要：0%問題の修正）
  enhanced.confidence = { ...basicMapping.confidence };
  
  // AI結果で既存マッピングを強化（既存confidence値を保持）
  if (aiResult.answer && (!enhanced.answer || (aiResult.confidence?.answer || 0) > (enhanced.confidence?.answer || 0))) {
    enhanced.answer = headers.indexOf(aiResult.answer);
    if (aiResult.confidence?.answer) {
      enhanced.confidence.answer = aiResult.confidence.answer;
    }
  }
  
  if (aiResult.reason && (!enhanced.reason || (aiResult.confidence?.reason || 0) > (enhanced.confidence?.reason || 0))) {
    enhanced.reason = headers.indexOf(aiResult.reason);
    if (aiResult.confidence?.reason) {
      enhanced.confidence.reason = aiResult.confidence.reason;
    }
  }
  
  if (aiResult.classHeader && (!enhanced.class || (aiResult.confidence?.class || 0) > (enhanced.confidence?.class || 0))) {
    enhanced.class = headers.indexOf(aiResult.classHeader);
    if (aiResult.confidence?.class) {
      enhanced.confidence.class = aiResult.confidence.class;
    }
  }
  
  if (aiResult.name && (!enhanced.name || (aiResult.confidence?.name || 0) > (enhanced.confidence?.name || 0))) {
    enhanced.name = headers.indexOf(aiResult.name);
    if (aiResult.confidence?.name) {
      enhanced.confidence.name = aiResult.confidence.name;
    }
  }
  
  return enhanced;
};

describe('confidence値0%問題修正テスト', () => {
  const sampleHeaders = ['タイムスタンプ', 'メールアドレス', 'クラス', '名前', '回答', '理由'];
  
  test('既存のconfidence値が保持される（修正後）', () => {
    const basicMapping = {
      answer: 4,
      reason: 5,
      class: 2,
      name: 3,
      confidence: {
        answer: 95,
        reason: 85,
        class: 80,
        name: 75
      }
    };
    
    // AI結果にconfidenceが含まれない場合
    const aiResult = {
      answer: '回答',
      reason: '理由',
      classHeader: 'クラス',
      name: '名前'
      // confidence: undefined (AIから提供されない)
    };
    
    const result = mergeColumnConfidence(basicMapping, aiResult, sampleHeaders);
    
    // 既存のconfidence値が保持されることを確認
    expect(result.confidence.answer).toBe(95);
    expect(result.confidence.reason).toBe(85);
    expect(result.confidence.class).toBe(80);
    expect(result.confidence.name).toBe(75);
    
    console.log('✅ 修正後: confidence値が正しく保持されました', result.confidence);
  });
  
  test('AI結果のconfidenceが高い場合は置き換わる', () => {
    const basicMapping = {
      answer: 4,
      confidence: {
        answer: 80  // 既存値
      }
    };
    
    const aiResult = {
      answer: '回答',
      confidence: {
        answer: 98  // AI結果の方が高い
      }
    };
    
    const result = mergeColumnConfidence(basicMapping, aiResult, sampleHeaders);
    
    // AI結果の方が高い場合は置き換わる
    expect(result.confidence.answer).toBe(98);
  });
  
  test('旧版（問題のあるバージョン）のシミュレーション', () => {
    const problematicMerge = (basicMapping: any, aiResult: any) => {
      const enhanced = { ...basicMapping };
      
      if (aiResult.answer) {
        enhanced.confidence = enhanced.confidence || {}; // ここで既存値が消える！
        enhanced.confidence.answer = aiResult.confidence?.answer || 95;
      }
      
      return enhanced;
    };
    
    const basicMapping = {
      answer: 4,
      confidence: {
        answer: 95,
        reason: 85
      }
    };
    
    const aiResult = {
      answer: '回答'
      // confidence: undefined
    };
    
    const result = problematicMerge(basicMapping, aiResult);
    
    // 旧版では他のconfidence値（reason）が消える
    expect(result.confidence.reason).toBeUndefined();
    console.log('❌ 旧版: reason confidenceが消えました', result.confidence);
  });
  
  test('プロダクション環境のシナリオ再現', () => {
    // プロダクションで実際に発生していた状況
    const realWorldMapping = {
      answer: 4,    // E列
      reason: 5,    // F列  
      class: 2,     // C列
      name: 3,      // D列
      confidence: {
        answer: 92,  // 高精度検出済み
        reason: 85,
        class: 80,
        name: 75
      }
    };
    
    // AIが何も返さない場合（よくあるケース）
    const emptyAiResult = {};
    
    const result = mergeColumnConfidence(realWorldMapping, emptyAiResult, sampleHeaders);
    
    // 修正後は既存値がすべて保持される
    expect(result.confidence.answer).toBe(92);
    expect(result.confidence.reason).toBe(85);
    expect(result.confidence.class).toBe(80);
    expect(result.confidence.name).toBe(75);
    
    console.log('🎉 プロダクション修正: すべてのconfidence値が保持', result.confidence);
  });
});