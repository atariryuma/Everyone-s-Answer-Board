/* global generateRecommendedMapping, resolveColumnIndex */

/**
 * Quick test for the optimized column detection system
 * Tests the specific user-reported problem cases
 */
function testOptimizedSystem() {
  console.log('🚀 Testing Optimized Column Detection System');
  console.log('='.repeat(50));

  // Test the exact user-reported problem scenario
  const problemHeaders = [
    'タイムスタンプ',
    'メールアドレス',
    'クラス',
    '名前',
    'どうして、メダカと一緒に、水草、ミジンコを入れると思いますか？観察していて、気づいたことを書きましょう。'
  ];

  const sampleData = [
    ['2024-01-15 10:30:00', 'student1@school.jp', '3-A', '山田太郎', 'メダカが元気になるから'],
    ['2024-01-15 10:31:00', 'student2@school.jp', '3-B', '佐藤花子', '自然の環境に近づけるため'],
    ['2024-01-15 10:32:00', 'student3@school.jp', '3-A', '田中一郎', '水をきれいにしてくれるから']
  ];

  console.log('\\n🎯 Testing User-Reported Problem Cases:');

  try {
    // Test Answer Column (E列, index 4) - User reported 62%
    const answerResult = resolveColumnIndex(problemHeaders, 'answer', {}, { sampleData });
    const answerSuccess = answerResult.index === 4 && answerResult.confidence > 80;
    console.log(`✓ Answer Column: ${answerResult.confidence.toFixed(1)}% confidence ${answerSuccess ? '✅' : '❌'}`);
    console.log(`  Previous: 62% → New: ${answerResult.confidence.toFixed(1)}% (${answerResult.confidence > 62 ? '+' : ''}${(answerResult.confidence - 62).toFixed(1)}%)`);

    // Test Class Column (C列, index 2) - User reported 25%
    const classResult = resolveColumnIndex(problemHeaders, 'class', {}, { sampleData });
    const classSuccess = classResult.index === 2 && classResult.confidence > 85;
    console.log(`✓ Class Column: ${classResult.confidence.toFixed(1)}% confidence ${classSuccess ? '✅' : '❌'}`);
    console.log(`  Previous: 25% → New: ${classResult.confidence.toFixed(1)}% (${classResult.confidence > 25 ? '+' : ''}${(classResult.confidence - 25).toFixed(1)}%)`);

    // Test Name Column (D列, index 3) - User reported 31%
    const nameResult = resolveColumnIndex(problemHeaders, 'name', {}, { sampleData });
    const nameSuccess = nameResult.index === 3 && nameResult.confidence > 85;
    console.log(`✓ Name Column: ${nameResult.confidence.toFixed(1)}% confidence ${nameSuccess ? '✅' : '❌'}`);
    console.log(`  Previous: 31% → New: ${nameResult.confidence.toFixed(1)}% (${nameResult.confidence > 31 ? '+' : ''}${(nameResult.confidence - 31).toFixed(1)}%)`);

    // Test Email Column
    const emailResult = resolveColumnIndex(problemHeaders, 'email', {}, { sampleData });
    const emailSuccess = emailResult.index === 1 && emailResult.confidence > 85;
    console.log(`✓ Email Column: ${emailResult.confidence.toFixed(1)}% confidence ${emailSuccess ? '✅' : '❌'}`);

    // Summary
    const allSuccess = answerSuccess && classSuccess && nameSuccess && emailSuccess;
    const avgOldConfidence = (62 + 25 + 31) / 3;
    const avgNewConfidence = (answerResult.confidence + classResult.confidence + nameResult.confidence) / 3;
    const avgImprovement = avgNewConfidence - avgOldConfidence;

    console.log(`\\n${  '='.repeat(50)}`);
    console.log('🏁 OPTIMIZATION RESULTS:');
    console.log(`✓ All Fields Detected: ${allSuccess ? '✅ YES' : '❌ NO'}`);
    console.log(`✓ Average Confidence: ${avgOldConfidence.toFixed(1)}% → ${avgNewConfidence.toFixed(1)}%`);
    console.log(`✓ Average Improvement: +${avgImprovement.toFixed(1)}%`);
    console.log(`✓ Answer Target Met (>80%): ${answerResult.confidence > 80 ? '✅ YES' : '❌ NO'}`);
    console.log(`✓ Simple Fields >85%: ${classResult.confidence > 85 && nameResult.confidence > 85 ? '✅ YES' : '❌ NO'}`);

    if (allSuccess) {
      console.log('\\n🎉 RESULT: Optimization successful - All user issues resolved!');
      console.log('💡 RECOMMENDATION: Deploy optimized system');
    } else {
      console.log('\\n⚠️ RESULT: Some issues remain - further optimization needed');
    }

    return {
      success: allSuccess,
      results: { answerResult, classResult, nameResult, emailResult },
      avgImprovement
    };

  } catch (error) {
    console.error('❌ Test Error:', error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Test comprehensive mapping generation
 */
function testMappingGeneration() {
  console.log('\\n🧪 Testing Mapping Generation:');

  const headers = [
    'タイムスタンプ',
    'メールアドレス',
    'クラス',
    '名前',
    'どうして、メダカと一緒に、水草、ミジンコを入れると思いますか？観察していて、気づいたことを書きましょう。'
  ];

  const result = generateRecommendedMapping(headers);

  console.log('✓ Mapping Result:', result.success ? '✅' : '❌');
  console.log('✓ Fields Mapped:', Object.keys(result.recommendedMapping).length, '/', 5);
  console.log('✓ Average Confidence:', result.analysis?.overallScore, '%');
  console.log('✓ Detected Fields:', result.recommendedMapping);
  console.log('✓ Confidence Scores:', result.confidence);

  return result;
}

/**
 * Run all optimization tests
 */
function runAllOptimizationTests() {
  console.log('🚀 COMPLETE OPTIMIZATION TEST SUITE');
  console.log('='.repeat(60));
  console.log(`Started: ${new Date().toISOString()}`);

  // Run individual field test
  const fieldTest = testOptimizedSystem();

  // Run mapping generation test
  const mappingTest = testMappingGeneration();

  console.log(`\\n${  '='.repeat(60)}`);
  console.log('🏆 FINAL ASSESSMENT');
  console.log('='.repeat(60));

  const criteria = {
    allFieldsDetected: fieldTest.success,
    averageImprovement: fieldTest.avgImprovement > 20,
    answerHighConfidence: fieldTest.results?.answerResult.confidence > 80,
    simpleFieldsHigh: fieldTest.results?.classResult.confidence > 85 && fieldTest.results?.nameResult.confidence > 85,
    mappingWorks: mappingTest.success
  };

  const passedCriteria = Object.values(criteria).filter(Boolean).length;
  const totalCriteria = Object.keys(criteria).length;
  const successRate = (passedCriteria / totalCriteria) * 100;

  console.log('\\n📊 CRITERIA RESULTS:');
  Object.entries(criteria).forEach(([criterion, passed]) => {
    const status = passed ? '✅' : '❌';
    const descriptions = {
      allFieldsDetected: 'All problem fields correctly detected',
      averageImprovement: 'Average improvement >20%',
      answerHighConfidence: 'Answer column confidence >80%',
      simpleFieldsHigh: 'Simple fields confidence >85%',
      mappingWorks: 'Mapping generation successful'
    };
    console.log(`${status} ${descriptions[criterion]}`);
  });

  console.log(`\\n🎯 SUCCESS RATE: ${successRate.toFixed(1)}% (${passedCriteria}/${totalCriteria} criteria)`);

  if (successRate >= 100) {
    console.log('🏆 ASSESSMENT: PERFECT - All optimization goals achieved');
  } else if (successRate >= 80) {
    console.log('✅ ASSESSMENT: EXCELLENT - Ready for deployment');
  } else if (successRate >= 60) {
    console.log('⚠️ ASSESSMENT: GOOD - Minor improvements needed');
  } else {
    console.log('❌ ASSESSMENT: NEEDS WORK - Significant improvements required');
  }

  console.log('='.repeat(60));

  return {
    successRate,
    passedCriteria,
    totalCriteria,
    fieldTest,
    mappingTest
  };
}