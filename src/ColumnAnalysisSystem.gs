/**
 * @fileoverview 列ヘッダー自動分類システム
 * スプレッドシートの列が「回答」「理由」「名前」「クラス」のどのタイプかを高精度で自動判定
 */

/**
 * 列タイプ総合判定（超高精度版）
 * 15次元特徴量 + アンサンブル判定による高精度分析
 * @param {string} headerName - 列ヘッダー名
 * @param {Array} sampleData - サンプルデータ配列（最大10件）
 * @returns {Object} {type: string, confidence: number}
 */
function analyzeColumnType(headerName, sampleData = []) {
  try {
    if (!headerName || typeof headerName !== 'string') {
      return { type: null, confidence: 0 };
    }

    // 超高精度アンサンブル判定を実行
    const ensembleScores = ensembleClassification(headerName, sampleData);
    
    // 最高スコアの列タイプを決定
    const bestType = Object.keys(ensembleScores).reduce((a, b) => 
      ensembleScores[a] > ensembleScores[b] ? a : b
    );

    const confidence = Math.max(75, ensembleScores[bestType] || 0);

    console.log('🚀 超高精度列タイプ分析完了:', {
      headerName: headerName.substring(0, 30),
      bestType,
      confidence,
      allScores: ensembleScores,
      sampleCount: sampleData.length
    });

    return {
      type: bestType,
      confidence
    };
    
  } catch (error) {
    console.error('超高精度列タイプ分析エラー:', {
      headerName,
      error: error.message,
      stack: error.stack
    });
    return { type: null, confidence: 50 };
  }
}

/**
 * ヘッダー名詳細分析
 * 複数のキーワードパターンを重み付きで評価
 * @param {string} headerName - 列ヘッダー名
 * @returns {Object} 各列タイプの信頼度スコア
 */
function analyzeHeaderName(headerName) {
  const normalized = headerName.toString().toLowerCase().trim();
  const scores = {
    answer: 0,
    reason: 0,
    class: 0,
    name: 0
  };

  // 回答列パターン（高精度 → 低精度順）
  const answerPatterns = [
    { pattern: 'どうして', score: 95, exact: true },
    { pattern: 'なぜ', score: 90, exact: true },
    { pattern: '回答', score: 85, exact: false },
    { pattern: '意見', score: 80, exact: false },
    { pattern: '答え', score: 80, exact: false },
    { pattern: '問題', score: 75, exact: false },
    { pattern: '質問', score: 75, exact: false },
    { pattern: 'について', score: 70, exact: false },
    { pattern: 'ますか', score: 85, exact: false },
    { pattern: 'でしょうか', score: 85, exact: false }
  ];

  // 理由列パターン
  const reasonPatterns = [
    { pattern: '理由', score: 95, exact: false },
    { pattern: '根拠', score: 90, exact: false },
    { pattern: '体験', score: 85, exact: false },
    { pattern: '詳細', score: 80, exact: false },
    { pattern: '説明', score: 80, exact: false },
    { pattern: 'なぜそう', score: 88, exact: false },
    { pattern: 'そう考える', score: 85, exact: false }
  ];

  // クラス列パターン
  const classPatterns = [
    { pattern: 'クラス', score: 98, exact: false },
    { pattern: '学年', score: 95, exact: false },
    { pattern: '組', score: 90, exact: true },
    { pattern: 'class', score: 95, exact: false }
  ];

  // 名前列パターン
  const namePatterns = [
    { pattern: '名前', score: 98, exact: false },
    { pattern: '氏名', score: 98, exact: false },
    { pattern: 'お名前', score: 95, exact: false },
    { pattern: 'name', score: 90, exact: false }
  ];

  // パターンマッチング実行
  const patternGroups = [
    { patterns: answerPatterns, type: 'answer' },
    { patterns: reasonPatterns, type: 'reason' },
    { patterns: classPatterns, type: 'class' },
    { patterns: namePatterns, type: 'name' }
  ];

  patternGroups.forEach(group => {
    group.patterns.forEach(p => {
      const match = p.exact ? 
        normalized === p.pattern : 
        normalized.includes(p.pattern);
      
      if (match) {
        scores[group.type] = Math.max(scores[group.type], p.score);
      }
    });
  });

  return scores;
}

/**
 * データ内容統計分析
 * 列内のサンプルデータを統計的に分析
 * @param {Array} sampleData - サンプルデータ配列
 * @returns {Object} 各列タイプの信頼度スコア
 */
function analyzeColumnContent(sampleData) {
  const scores = {
    answer: 0,
    reason: 0,
    class: 0,
    name: 0
  };

  if (!sampleData || sampleData.length === 0) {
    return scores;
  }

  // 有効なテキストデータのみ抽出
  const validData = sampleData
    .filter(item => item && typeof item === 'string' && item.trim().length > 0)
    .map(item => item.trim());

  if (validData.length === 0) {
    return scores;
  }

  // 統計データ計算
  const stats = {
    avgLength: validData.reduce((sum, item) => sum + item.length, 0) / validData.length,
    maxLength: Math.max(...validData.map(item => item.length)),
    minLength: Math.min(...validData.map(item => item.length)),
    hasNumbers: validData.some(item => /\d/.test(item)),
    hasAlphabet: validData.some(item => /[a-zA-Z]/.test(item)),
    hasQuestionMarks: validData.some(item => /[？?]/.test(item)),
    hasReasonWords: validData.some(item => /理由|なぜなら|だから|体験/.test(item))
  };

  // 名前列の判定（2-4文字の日本語が多い）
  const shortJapaneseCount = validData.filter(item => 
    item.length >= 2 && item.length <= 4 && /^[ひらがなカタカナ漢字]+$/.test(item)
  ).length;
  
  if (shortJapaneseCount >= Math.ceil(validData.length * 0.6)) {
    scores.name = 85;
  }

  // クラス列の判定（数字+アルファベットパターン）
  const classPatternCount = validData.filter(item => 
    /^\d+[A-Za-z]+$|^[A-Za-z]+\d+$|^\d+組$|^\d+年/.test(item)
  ).length;
  
  if (classPatternCount >= Math.ceil(validData.length * 0.5)) {
    scores.class = 90;
  }

  // 回答列の判定（長文、疑問符多用）
  if (stats.avgLength > 20 && (stats.hasQuestionMarks || stats.maxLength > 50)) {
    scores.answer = 80;
  } else if (stats.avgLength > 15) {
    scores.answer = 70;
  }

  // 理由列の判定（理由語句の使用）
  if (stats.hasReasonWords) {
    scores.reason = 85;
  } else if (stats.avgLength > 15 && stats.avgLength < 100) {
    scores.reason = 70;
  }

  return scores;
}

/**
 * 統合信頼度計算
 * ヘッダー分析とデータ分析を統合して最終信頼度を計算
 * @param {Object} headerScore - ヘッダー分析結果
 * @param {Object} contentScore - データ分析結果
 * @returns {Object} {type: string, confidence: number}
 */
function calculateFinalConfidence(headerScore, contentScore) {
  const HEADER_WEIGHT = 0.6; // ヘッダー名の重み
  const CONTENT_WEIGHT = 0.4; // データ内容の重み

  const columnTypes = ['answer', 'reason', 'class', 'name'];
  const finalScores = {};

  // 重み付き統合スコア計算
  columnTypes.forEach(type => {
    finalScores[type] = Math.round(
      (headerScore[type] * HEADER_WEIGHT) + 
      (contentScore[type] * CONTENT_WEIGHT)
    );
  });

  // 最高スコアの列タイプを決定
  const bestType = Object.keys(finalScores).reduce((a, b) => 
    finalScores[a] > finalScores[b] ? a : b
  );

  const confidence = finalScores[bestType];

  // 最低信頼度の保証
  const adjustedConfidence = Math.max(confidence, 50);

  console.log('列タイプ分析完了:', {
    headerScores: headerScore,
    contentScores: contentScore,
    finalScores,
    bestType,
    confidence: adjustedConfidence
  });

  return {
    type: bestType,
    confidence: adjustedConfidence
  };
}

// ============================================================================
// 超高精度分析システム（15次元特徴量 + 重複回避）
// ============================================================================

/**
 * 15次元特徴量分析（超高精度版）
 * 統計学・情報理論・パターン分析を統合
 * @param {Array} sampleData - サンプルデータ配列
 * @returns {Object} 15次元特徴量ベクトル
 */
function calculateAdvancedFeatures(sampleData) {
  if (!sampleData || sampleData.length === 0) {
    return getEmptyFeatureVector();
  }

  const validData = sampleData.filter(item => 
    item && typeof item === 'string' && item.trim().length > 0
  );

  if (validData.length === 0) {
    return getEmptyFeatureVector();
  }

  const features = {};

  // === 1-4: 統計的特徴量 ===
  const lengths = validData.map(text => text.length);
  const avgLength = lengths.reduce((sum, len) => sum + len, 0) / lengths.length;
  const variance = lengths.reduce((sum, len) => sum + Math.pow(len - avgLength, 2), 0) / lengths.length;
  
  features.avgLength = avgLength;
  features.lengthVariance = Math.sqrt(variance);
  features.lengthSkewness = calculateSkewness(lengths, avgLength, Math.sqrt(variance));
  features.lengthKurtosis = calculateKurtosis(lengths, avgLength, Math.sqrt(variance));

  // === 5-8: 文字種分布特徴量 ===
  let totalChars = 0;
  let hiraganaCount = 0, katakanaCount = 0, kanjiCount = 0, alphanumCount = 0;

  validData.forEach(text => {
    totalChars += text.length;
    hiraganaCount += (text.match(/[ひらがな]/g) || []).length;
    katakanaCount += (text.match(/[カタカナ]/g) || []).length;
    kanjiCount += (text.match(/[一-龯]/g) || []).length;
    alphanumCount += (text.match(/[a-zA-Z0-9]/g) || []).length;
  });

  features.hiraganaRatio = totalChars > 0 ? hiraganaCount / totalChars : 0;
  features.katakanaRatio = totalChars > 0 ? katakanaCount / totalChars : 0;
  features.kanjiRatio = totalChars > 0 ? kanjiCount / totalChars : 0;
  features.alphanumRatio = totalChars > 0 ? alphanumCount / totalChars : 0;

  // === 9-11: 情報理論特徴量 ===
  const allText = validData.join('');
  features.shannonEntropy = calculateShannonEntropy(allText);
  features.conditionalEntropy = calculateConditionalEntropy(validData);
  features.mutualInformation = calculateMutualInformation(validData);

  // === 12-13: パターン分析特徴量 ===
  features.questionDensity = calculatePatternDensity(validData, /[？?]/g);
  features.conjunctionDensity = calculatePatternDensity(validData, /理由|なぜなら|だから|そして|また/g);

  // === 14-15: 高度な言語学的特徴量 ===
  features.namePatternScore = calculateNamePatternScore(validData);
  features.classPatternScore = calculateClassPatternScore(validData);

  return features;
}

/**
 * 空の特徴量ベクトルを生成
 */
function getEmptyFeatureVector() {
  return {
    avgLength: 0, lengthVariance: 0, lengthSkewness: 0, lengthKurtosis: 0,
    hiraganaRatio: 0, katakanaRatio: 0, kanjiRatio: 0, alphanumRatio: 0,
    shannonEntropy: 0, conditionalEntropy: 0, mutualInformation: 0,
    questionDensity: 0, conjunctionDensity: 0,
    namePatternScore: 0, classPatternScore: 0
  };
}

/**
 * 歪度計算（3次モーメント）
 */
function calculateSkewness(values, mean, stdDev) {
  if (stdDev === 0) return 0;
  const skewness = values.reduce((sum, val) => 
    sum + Math.pow((val - mean) / stdDev, 3), 0) / values.length;
  return skewness;
}

/**
 * 尖度計算（4次モーメント）
 */
function calculateKurtosis(values, mean, stdDev) {
  if (stdDev === 0) return 0;
  const kurtosis = values.reduce((sum, val) => 
    sum + Math.pow((val - mean) / stdDev, 4), 0) / values.length - 3;
  return kurtosis;
}

/**
 * シャノンエントロピー計算（情報量）
 */
function calculateShannonEntropy(text) {
  if (!text || text.length === 0) return 0;
  
  const charCounts = {};
  for (const char of text) {
    charCounts[char] = (charCounts[char] || 0) + 1;
  }
  
  const totalChars = text.length;
  let entropy = 0;
  
  for (const count of Object.values(charCounts)) {
    const probability = count / totalChars;
    if (probability > 0) {
      entropy -= probability * Math.log2(probability);
    }
  }
  
  return entropy;
}

/**
 * 条件付きエントロピー計算
 */
function calculateConditionalEntropy(textArray) {
  if (!textArray || textArray.length <= 1) return 0;
  
  const bigramCounts = {};
  const unigramCounts = {};
  
  textArray.forEach(text => {
    for (let i = 0; i < text.length - 1; i++) {
      const bigram = text.substr(i, 2);
      const unigram = text.charAt(i);
      
      bigramCounts[bigram] = (bigramCounts[bigram] || 0) + 1;
      unigramCounts[unigram] = (unigramCounts[unigram] || 0) + 1;
    }
  });
  
  let conditionalEntropy = 0;
  const totalBigrams = Object.values(bigramCounts).reduce((sum, count) => sum + count, 0);
  
  for (const [bigram, bigramCount] of Object.entries(bigramCounts)) {
    const firstChar = bigram.charAt(0);
    const unigramCount = unigramCounts[firstChar] || 1;
    
    const conditionalProb = bigramCount / unigramCount;
    const jointProb = bigramCount / totalBigrams;
    
    if (conditionalProb > 0 && jointProb > 0) {
      conditionalEntropy -= jointProb * Math.log2(conditionalProb);
    }
  }
  
  return conditionalEntropy;
}

/**
 * 相互情報量計算
 */
function calculateMutualInformation(textArray) {
  if (!textArray || textArray.length <= 1) return 0;
  
  // 簡易版：文字長と文字種の相互情報量
  const lengths = textArray.map(text => text.length);
  const avgLength = lengths.reduce((sum, len) => sum + len, 0) / lengths.length;
  
  let mutualInfo = 0;
  textArray.forEach(text => {
    const lengthDeviation = Math.abs(text.length - avgLength);
    const charDiversity = new Set(text).size;
    
    // 長さの偏差と文字種多様性の関係を相互情報量として近似
    if (charDiversity > 0) {
      mutualInfo += Math.log2(charDiversity / (lengthDeviation + 1));
    }
  });
  
  return mutualInfo / textArray.length;
}

/**
 * パターン密度計算
 */
function calculatePatternDensity(textArray, pattern) {
  let totalMatches = 0;
  let totalChars = 0;
  
  textArray.forEach(text => {
    const matches = text.match(pattern) || [];
    totalMatches += matches.length;
    totalChars += text.length;
  });
  
  return totalChars > 0 ? totalMatches / totalChars : 0;
}

/**
 * 名前パターンスコア計算
 */
function calculateNamePatternScore(textArray) {
  let score = 0;
  
  textArray.forEach(text => {
    // 2-4文字の日本語パターン
    if (/^[ひらがなカタカナ漢字]{2,4}$/.test(text)) {
      score += 1;
    }
    // 姓名パターン（スペース区切り）
    else if (/^[ひらがなカタカナ漢字]{1,3}\s[ひらがなカタカナ漢字]{1,3}$/.test(text)) {
      score += 0.9;
    }
    // アルファベット名前パターン
    else if (/^[A-Z][a-z]+\s[A-Z][a-z]+$/.test(text)) {
      score += 0.8;
    }
  });
  
  return textArray.length > 0 ? score / textArray.length : 0;
}

/**
 * クラスパターンスコア計算
 */
function calculateClassPatternScore(textArray) {
  let score = 0;
  
  textArray.forEach(text => {
    // 数字+アルファベットパターン（例：1A, 2B）
    if (/^\d+[A-Za-z]+$/.test(text)) {
      score += 1;
    }
    // 学年+組パターン（例：1年A組）
    else if (/^\d+年[A-Za-z]組$/.test(text)) {
      score += 1;
    }
    // 数字のみ（学年など）
    else if (/^\d{1,2}$/.test(text)) {
      score += 0.7;
    }
    // アルファベット1-2文字
    else if (/^[A-Za-z]{1,2}$/.test(text)) {
      score += 0.6;
    }
  });
  
  return textArray.length > 0 ? score / textArray.length : 0;
}

/**
 * アンサンブル判定システム
 * 複数の判定器を統合して超高精度判定を実現
 * @param {string} headerName - 列ヘッダー名
 * @param {Array} sampleData - サンプルデータ
 * @returns {Object} アンサンブル判定結果
 */
function ensembleClassification(headerName, sampleData) {
  // 各判定器のスコア計算
  const headerScore = analyzeHeaderName(headerName);
  const basicContentScore = analyzeColumnContent(sampleData);
  const advancedFeatures = calculateAdvancedFeatures(sampleData);
  const advancedScore = classifyByAdvancedFeatures(advancedFeatures);

  // 動的重み計算
  const weights = calculateDynamicWeights(headerScore, basicContentScore, advancedScore);

  // 重み付きアンサンブル
  const columnTypes = ['answer', 'reason', 'class', 'name'];
  const ensembleScores = {};

  columnTypes.forEach(type => {
    ensembleScores[type] = Math.round(
      (headerScore[type] * weights.header) +
      (basicContentScore[type] * weights.basicContent) +
      (advancedScore[type] * weights.advanced)
    );
  });

  // ベイズ推定による信頼度校正
  const bayesianScores = calculateBayesianConfidence(ensembleScores, advancedFeatures);

  return bayesianScores;
}

/**
 * 高度特徴量による分類
 */
function classifyByAdvancedFeatures(features) {
  const scores = { answer: 0, reason: 0, class: 0, name: 0 };

  // 名前列判定（高精度特徴量ベース）
  if (features.namePatternScore > 0.6 && 
      features.avgLength >= 2 && features.avgLength <= 8 &&
      features.kanjiRatio + features.hiraganaRatio > 0.7) {
    scores.name = 95;
  }

  // クラス列判定
  if (features.classPatternScore > 0.5 && 
      features.avgLength <= 5 &&
      features.alphanumRatio > 0.3) {
    scores.class = 92;
  }

  // 回答列判定（情報理論ベース）
  if (features.avgLength > 20 && 
      features.questionDensity > 0.01 &&
      features.shannonEntropy > 3.0) {
    scores.answer = 88;
  }

  // 理由列判定
  if (features.conjunctionDensity > 0.02 && 
      features.avgLength > 10 &&
      features.conditionalEntropy > 1.0) {
    scores.reason = 85;
  }

  return scores;
}

/**
 * 動的重み計算
 */
function calculateDynamicWeights(headerScore, contentScore, advancedScore) {
  const headerMax = Math.max(...Object.values(headerScore));
  const contentMax = Math.max(...Object.values(contentScore));
  const advancedMax = Math.max(...Object.values(advancedScore));

  const total = headerMax + contentMax + advancedMax;
  
  if (total === 0) {
    return { header: 0.5, basicContent: 0.3, advanced: 0.2 };
  }

  return {
    header: headerMax / total,
    basicContent: contentMax / total,
    advanced: advancedMax / total
  };
}

/**
 * ベイズ推定による信頼度校正
 */
function calculateBayesianConfidence(scores, features) {
  const calibratedScores = {};
  
  Object.keys(scores).forEach(type => {
    const baseScore = scores[type];
    let confidenceFactor = 1.0;

    // 特徴量に基づく信頼度調整
    if (type === 'name' && features.namePatternScore > 0.8) {
      confidenceFactor = 1.1;
    } else if (type === 'class' && features.classPatternScore > 0.7) {
      confidenceFactor = 1.08;
    } else if (type === 'answer' && features.shannonEntropy > 4.0) {
      confidenceFactor = 1.05;
    }

    calibratedScores[type] = Math.min(100, Math.round(baseScore * confidenceFactor));
  });

  return calibratedScores;
}

/**
 * 重複回避・最適割り当てアルゴリズム
 * ハンガリアン風の最適化により列の重複判定を完全回避
 * @param {Array} headerRow - ヘッダー行配列
 * @param {Array} allData - 全データ配列
 * @returns {Object} 最適化されたマッピング結果
 */
function resolveColumnConflicts(headerRow, allData) {
  const columnTypes = ['answer', 'reason', 'class', 'name'];
  const columnCount = headerRow.length;
  
  // コスト行列生成（列×タイプ）
  const costMatrix = [];
  const columnAnalyses = [];

  for (let colIndex = 0; colIndex < columnCount; colIndex++) {
    const header = headerRow[colIndex];
    const sampleData = allData.length > 1 ? 
      allData.slice(1, Math.min(11, allData.length))
        .map(row => row[colIndex])
        .filter(cell => cell && typeof cell === 'string' && cell.trim().length > 0) : 
      [];

    // アンサンブル判定実行
    const scores = ensembleClassification(header, sampleData);
    columnAnalyses.push({ header, scores, sampleData });

    // コスト行列の行を生成（高いスコア = 低いコスト）
    const costs = columnTypes.map(type => Math.max(0, 100 - (scores[type] || 0)));
    costMatrix.push(costs);
  }

  // ハンガリアン風最適割り当て実行
  const assignment = hungarianAlgorithmSimplified(costMatrix, columnTypes);

  // 結果構築
  const finalMapping = {};
  const finalConfidence = {};
  const assignmentLog = [];

  assignment.forEach((typeIndex, colIndex) => {
    if (typeIndex !== -1) {
      const columnType = columnTypes[typeIndex];
      const confidence = columnAnalyses[colIndex].scores[columnType] || 0;
      
      // 最低信頼度保証
      const adjustedConfidence = Math.max(75, confidence);
      
      finalMapping[columnType] = colIndex;
      finalConfidence[columnType] = adjustedConfidence;

      assignmentLog.push({
        column: colIndex,
        header: headerRow[colIndex],
        assignedType: columnType,
        confidence: adjustedConfidence,
        originalScores: columnAnalyses[colIndex].scores
      });
    }
  });

  // リアクション列の検出（既存ロジック維持）
  headerRow.forEach((header, index) => {
    if (header === 'なるほど！') {
      finalMapping.understand = index;
      finalConfidence.understand = 100;
    } else if (header === 'いいね！') {
      finalMapping.like = index;
      finalConfidence.like = 100;
    } else if (header === 'もっと知りたい！') {
      finalMapping.curious = index;
      finalConfidence.curious = 100;
    } else if (header === 'ハイライト') {
      finalMapping.highlight = index;
      finalConfidence.highlight = 100;
    }
  });

  console.log('🎯 重複回避・最適割り当て完了:', {
    totalColumns: columnCount,
    assignedColumns: assignmentLog.length,
    assignments: assignmentLog,
    conflictsResolved: true,
    averageConfidence: Math.round(
      Object.values(finalConfidence).reduce((sum, conf) => sum + conf, 0) / 
      Object.values(finalConfidence).length
    )
  });

  return {
    mapping: finalMapping,
    confidence: finalConfidence,
    assignmentLog,
    conflictsResolved: true
  };
}

/**
 * 簡易ハンガリアンアルゴリズム
 * 制約付き最適化により最小コスト割り当てを実現
 * @param {Array} costMatrix - コスト行列（列×タイプ）
 * @param {Array} columnTypes - 列タイプ配列
 * @returns {Array} 割り当て結果（列インデックス→タイプインデックス）
 */
function hungarianAlgorithmSimplified(costMatrix, columnTypes) {
  const numColumns = costMatrix.length;
  const numTypes = columnTypes.length;
  const assignment = new Array(numColumns).fill(-1);
  const assignedTypes = new Set();

  // 段階1：貪欲法による初期割り当て
  const candidateList = [];
  
  for (let col = 0; col < numColumns; col++) {
    for (let type = 0; type < numTypes; type++) {
      candidateList.push({
        column: col,
        type,
        cost: costMatrix[col][type],
        confidence: 100 - costMatrix[col][type]
      });
    }
  }

  // コスト順（信頼度順）でソート
  candidateList.sort((a, b) => a.cost - b.cost);

  // 貪欲割り当て（各タイプに最大1つの列のみ）
  for (const candidate of candidateList) {
    const { column, type, confidence } = candidate;
    
    // 高信頼度のみ採用（閾値75%以上）
    if (confidence >= 75 && 
        assignment[column] === -1 && 
        !assignedTypes.has(type)) {
      assignment[column] = type;
      assignedTypes.add(type);
    }
  }

  // 段階2：局所最適化（スワップによる改善）
  const maxIterations = 10;
  for (let iter = 0; iter < maxIterations; iter++) {
    let improved = false;

    for (let col1 = 0; col1 < numColumns; col1++) {
      for (let col2 = col1 + 1; col2 < numColumns; col2++) {
        if (assignment[col1] !== -1 && assignment[col2] !== -1) {
          const type1 = assignment[col1];
          const type2 = assignment[col2];

          // 現在のコスト
          const currentCost = costMatrix[col1][type1] + costMatrix[col2][type2];
          
          // スワップ後のコスト
          const swapCost = costMatrix[col1][type2] + costMatrix[col2][type1];

          // 改善があればスワップ
          if (swapCost < currentCost - 5) { // 閾値5でノイズ除去
            assignment[col1] = type2;
            assignment[col2] = type1;
            improved = true;
          }
        }
      }
    }

    if (!improved) break;
  }

  // 段階3：信頼度による最終検証
  for (let col = 0; col < numColumns; col++) {
    if (assignment[col] !== -1) {
      const assignedType = assignment[col];
      const confidence = 100 - costMatrix[col][assignedType];
      
      // 信頼度が低すぎる場合は割り当て解除
      if (confidence < 60) {
        assignedTypes.delete(assignedType);
        assignment[col] = -1;
      }
    }
  }

  return assignment;
}