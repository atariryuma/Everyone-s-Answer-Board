/* global generateRecommendedMapping, analyzeFieldRelationships */

/**
 * Test logical field ordering constraints (answer before reason)
 */
function testLogicalOrdering() {
  console.log('🧠 Testing Logical Field Ordering System');
  console.log('='.repeat(50));

  const testCases = [
    {
      name: 'Correct Order (Answer → Reason)',
      headers: [
        'タイムスタンプ',
        'どう思いますか？',  // Answer
        'その理由を教えてください',  // Reason
        'メール',
        '名前'
      ],
      expected: {
        answerIndex: 1,
        reasonIndex: 2,
        logical: true
      }
    },
    {
      name: 'Incorrect Order (Reason → Answer)',
      headers: [
        'タイムスタンプ',
        'なぜそう思うのですか？',  // Reason (appears first - illogical)
        'あなたの意見を述べてください',  // Answer
        'メール',
        '名前'
      ],
      expected: {
        answerIndex: 2,
        reasonIndex: 1,
        logical: false,
        shouldCorrect: true
      }
    },
    {
      name: 'Complex Question With Logical Order',
      headers: [
        'タイムスタンプ',
        'どうして、メダカと一緒に、水草、ミジンコを入れると思いますか？観察していて、気づいたことを書きましょう。',  // Answer
        'なぜそのように考えたのか理由を書いてください',  // Reason
        'クラス',
        '名前'
      ],
      expected: {
        answerIndex: 1,
        reasonIndex: 2,
        logical: true
      }
    },
    {
      name: 'Only Answer Field (No Reason)',
      headers: [
        'タイムスタンプ',
        'ご意見をお聞かせください',  // Answer only
        'メール',
        '名前'
      ],
      expected: {
        answerIndex: 1,
        reasonIndex: -1,
        logical: true
      }
    },
    {
      name: 'Ambiguous Case (Close Confidence)',
      headers: [
        'タイムスタンプ',
        'どうしてそう思いますか',  // Could be reason
        'あなたの考えを教えてください',  // Could be answer
        'メール'
      ],
      expected: {
        shouldAnalyze: true
      }
    }
  ];

  const results = [];

  testCases.forEach((testCase, index) => {
    console.log(`${index + 1}. ${testCase.name}:`);
    console.log(`   Headers: ${testCase.headers.slice(1, -1).join(' | ')}`);

    try {
      const result = generateRecommendedMapping(testCase.headers);
      const mapping = result.recommendedMapping;
      const {confidence} = result;
      const validation = result.analysis?.logicalValidation;

      const actualAnswerIndex = mapping.answer;
      const actualReasonIndex = mapping.reason;
      const hasLogicalOrder = actualReasonIndex === undefined || actualAnswerIndex < actualReasonIndex;

      // Check against expectations
      let testPassed = true;
      const details = [];

      if (testCase.expected.answerIndex !== undefined) {
        const answerCorrect = actualAnswerIndex === testCase.expected.answerIndex;
        details.push(`Answer: Expected ${testCase.expected.answerIndex}, Got ${actualAnswerIndex} ${answerCorrect ? '✅' : '❌'}`);
        if (!answerCorrect) testPassed = false;
      }

      if (testCase.expected.reasonIndex !== undefined && testCase.expected.reasonIndex !== -1) {
        const reasonCorrect = actualReasonIndex === testCase.expected.reasonIndex;
        details.push(`Reason: Expected ${testCase.expected.reasonIndex}, Got ${actualReasonIndex} ${reasonCorrect ? '✅' : '❌'}`);
        if (!reasonCorrect) testPassed = false;
      } else if (testCase.expected.reasonIndex === -1) {
        const noReason = actualReasonIndex === undefined;
        details.push(`Reason: Expected none, Got ${actualReasonIndex || 'none'} ${noReason ? '✅' : '❌'}`);
        if (!noReason) testPassed = false;
      }

      if (testCase.expected.logical !== undefined) {
        const logicalCorrect = hasLogicalOrder === testCase.expected.logical;
        details.push(`Logical Order: Expected ${testCase.expected.logical}, Got ${hasLogicalOrder} ${logicalCorrect ? '✅' : '❌'}`);
        if (!logicalCorrect && !testCase.expected.shouldCorrect) testPassed = false;
      }

      // Display confidence scores
      const answerConf = confidence.answer ? confidence.answer.toFixed(1) : 'N/A';
      const reasonConf = confidence.reason ? confidence.reason.toFixed(1) : 'N/A';
      details.push(`Confidence: Answer=${answerConf}%, Reason=${reasonConf}%`);

      // Display validation results
      if (validation) {
        if (validation.corrections && validation.corrections.length > 0) {
          details.push(`Corrections: ${validation.corrections.length} applied`);
        }
        if (validation.warnings && validation.warnings.length > 0) {
          details.push(`Warnings: ${validation.warnings.length} raised`);
        }
      }

      console.log(`   Result: ${testPassed ? '✅ PASS' : '❌ FAIL'}`);
      details.forEach(detail => console.log(`   ${detail}`));

      results.push({
        testName: testCase.name,
        passed: testPassed,
        actualAnswerIndex,
        actualReasonIndex,
        hasLogicalOrder,
        confidence,
        validation
      });

    } catch (error) {
      console.log(`   Result: ❌ ERROR - ${error.message}`);
      results.push({
        testName: testCase.name,
        passed: false,
        error: error.message
      });
    }
  });

  // Summary
  console.log('='.repeat(50));
  console.log('🏁 LOGICAL ORDERING TEST SUMMARY');
  console.log('='.repeat(50));

  const passedTests = results.filter(r => r.passed).length;
  const totalTests = results.length;
  const successRate = (passedTests / totalTests) * 100;

  console.log(`✓ Test Results: ${passedTests}/${totalTests} passed (${successRate.toFixed(1)}%)`);

  // Analyze logical ordering compliance
  const logicalTests = results.filter(r => r.actualAnswerIndex !== undefined && r.actualReasonIndex !== undefined);
  const logicallyCorrect = logicalTests.filter(r => r.hasLogicalOrder).length;
  const logicalRate = logicalTests.length > 0 ? (logicallyCorrect / logicalTests.length) * 100 : 100;

  console.log(`✓ Logical Ordering: ${logicallyCorrect}/${logicalTests.length} correct (${logicalRate.toFixed(1)}%)`);

  // Confidence analysis
  const avgAnswerConf = results
    .filter(r => r.confidence && r.confidence.answer)
    .reduce((sum, r, _, arr) => sum + r.confidence.answer / arr.length, 0);
  const avgReasonConf = results
    .filter(r => r.confidence && r.confidence.reason)
    .reduce((sum, r, _, arr) => sum + r.confidence.reason / arr.length, 0);

  console.log(`✓ Average Confidence: Answer=${avgAnswerConf.toFixed(1)}%, Reason=${avgReasonConf.toFixed(1)}%`);

  // Final assessment
  if (successRate >= 100 && logicalRate >= 100) {
    console.log('🏆 ASSESSMENT: PERFECT - All logical ordering constraints satisfied');
  } else if (successRate >= 80 && logicalRate >= 90) {
    console.log('✅ ASSESSMENT: EXCELLENT - Strong logical ordering compliance');
  } else if (successRate >= 60 && logicalRate >= 75) {
    console.log('⚠️ ASSESSMENT: GOOD - Some logical ordering improvements needed');
  } else {
    console.log('❌ ASSESSMENT: NEEDS WORK - Logical ordering system requires fixes');
  }

  console.log('='.repeat(50));

  return {
    successRate,
    logicalRate,
    avgAnswerConf,
    avgReasonConf,
    results
  };
}

/**
 * Test the relationship analysis function specifically
 */
function testRelationshipAnalysis() {
  console.log('🔍 Testing Field Relationship Analysis');
  console.log('='.repeat(40));

  const testHeaders = [
    'タイムスタンプ',
    'なぜそう思いますか？',  // Reason candidate (index 1)
    'あなたの意見を教えてください',  // Answer candidate (index 2)
    'どうしてそのように考えたのですか？',  // Another reason candidate (index 3)
    'メール'
  ];

  const relationships = analyzeFieldRelationships(testHeaders);

  console.log('Answer Candidates:', relationships.answerCandidates.map(a => `${a.index}:${a.score.toFixed(1)}%`));
  console.log('Reason Candidates:', relationships.reasonCandidates.map(r => `${r.index}:${r.score.toFixed(1)}%`));
  console.log('Logical Pairs:', relationships.answerReasonPairs.map(p =>
    `Answer[${p.answerIndex}] → Reason[${p.reasonIndex}] (${p.confidence.toFixed(1)}%)`
  ));

  const hasLogicalPairs = relationships.answerReasonPairs.length > 0;
  const allPairsLogical = relationships.answerReasonPairs.every(p => p.logicalOrder);

  console.log(`Analysis Result: ${hasLogicalPairs && allPairsLogical ? '✅ LOGICAL' : '❌ ILLOGICAL'}`);

  return relationships;
}

/**
 * Run comprehensive logical ordering validation
 */
function runLogicalOrderingValidation() {
  console.log('🚀 COMPREHENSIVE LOGICAL ORDERING VALIDATION');
  console.log('='.repeat(60));
  console.log(`Started: ${new Date().toISOString()}`);

  // Run main logical ordering test
  const orderingTest = testLogicalOrdering();

  // Run relationship analysis test
  const relationshipTest = testRelationshipAnalysis();

  console.log('='.repeat(60));
  console.log('🏆 FINAL LOGICAL ORDERING ASSESSMENT');
  console.log('='.repeat(60));

  const criteria = {
    basicOrdering: orderingTest.successRate >= 80,
    logicalConstraints: orderingTest.logicalRate >= 90,
    answerDetection: orderingTest.avgAnswerConf >= 85,
    reasonDetection: orderingTest.avgReasonConf >= 75,
    relationshipAnalysis: relationshipTest.answerCandidates.length > 0 || relationshipTest.reasonCandidates.length > 0
  };

  const passedCriteria = Object.values(criteria).filter(Boolean).length;
  const totalCriteria = Object.keys(criteria).length;
  const overallScore = (passedCriteria / totalCriteria) * 100;

  console.log('📊 CRITERIA RESULTS:');
  Object.entries(criteria).forEach(([criterion, passed]) => {
    const status = passed ? '✅' : '❌';
    const descriptions = {
      basicOrdering: 'Basic field detection accuracy >80%',
      logicalConstraints: 'Logical ordering compliance >90%',
      answerDetection: 'Answer field confidence >85%',
      reasonDetection: 'Reason field confidence >75%',
      relationshipAnalysis: 'Relationship analysis functional'
    };
    console.log(`${status} ${descriptions[criterion]}`);
  });

  console.log(`🎯 OVERALL SCORE: ${overallScore.toFixed(1)}% (${passedCriteria}/${totalCriteria} criteria)`);

  if (overallScore >= 100) {
    console.log('🏆 ASSESSMENT: PERFECT - Logical ordering system fully functional');
  } else if (overallScore >= 80) {
    console.log('✅ ASSESSMENT: EXCELLENT - Ready for deployment');
  } else if (overallScore >= 60) {
    console.log('⚠️ ASSESSMENT: GOOD - Minor improvements recommended');
  } else {
    console.log('❌ ASSESSMENT: NEEDS WORK - Significant logical ordering issues');
  }

  console.log('='.repeat(60));

  return {
    overallScore,
    passedCriteria,
    totalCriteria,
    orderingTest,
    relationshipTest
  };
}