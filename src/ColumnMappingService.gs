/**
 * @fileoverview ColumnMappingService - 統一列判定・マッピングサービス
 *
 * 🎯 責任範囲:
 * - 列インデックス解決（マッピング・パターン・位置ベース）
 * - ヘッダーパターン判定
 * - フィールド値抽出（統一API）
 * - 列診断・レポート生成
 * - システム健全性チェック
 *
 * 🔄 CLAUDE.md Best Practices準拠:
 * - Zero-Dependency Architecture（完全独立）
 * - GAS-Native Pattern（直接API）
 * - 単一責任原則（列マッピング専用）
 * - V8ランタイム最適化
 */

// ===========================================
// 🎯 統一列判定システム - CLAUDE.md準拠
// ===========================================

/**
 * 統一列判定関数 - 最適化されたAI判定システム
 * @param {Array} headers - ヘッダー配列
 * @param {string} fieldType - フィールドタイプ
 * @param {Object} columnMapping - 列マッピング（優先）
 * @param {Object} options - オプション設定（sampleData含む）
 * @returns {Object} { index: number, confidence: number, method: string, debug: Object }
 */
function resolveColumnIndex(headers, fieldType, columnMapping = {}, options = {}) {
  const debugInfo = {
    fieldType,
    searchMethods: [],
    candidateHeaders: [],
    finalSelection: null,
    scoringDetails: {}
  };

  try {
    // 🛡️ 入力検証強化
    if (!headers || !Array.isArray(headers) || headers.length === 0) {
      debugInfo.error = 'Invalid or empty headers array';
      return { index: -1, confidence: 0, method: 'validation_failed', debug: debugInfo };
    }

    if (!fieldType || typeof fieldType !== 'string' || fieldType.trim() === '') {
      debugInfo.error = 'Invalid or empty fieldType';
      return { index: -1, confidence: 0, method: 'validation_failed', debug: debugInfo };
    }

    // 1. 明示的マッピング（最高優先度）
    if (columnMapping && typeof columnMapping === 'object' && columnMapping[fieldType] !== undefined && columnMapping[fieldType] !== null) {
      const mappedIndex = columnMapping[fieldType];
      if (typeof mappedIndex === 'number' && Number.isInteger(mappedIndex) && mappedIndex >= 0 && mappedIndex < headers.length) {
        debugInfo.searchMethods.push({ method: 'explicit_mapping', index: mappedIndex, confidence: 100 });
        debugInfo.finalSelection = { method: 'explicit_mapping', index: mappedIndex };
        return { index: mappedIndex, confidence: 100, method: 'explicit_mapping', debug: debugInfo };
      }
    }

    // 2. 🧠 AI強化パターンマッチング - 全候補を評価してベストマッチを選択
    const candidates = [];
    const headerPatterns = getHeaderPatterns();
    const patterns = headerPatterns[fieldType] || [];

    // 各ヘッダーを候補として評価
    headers.forEach((header, index) => {
      if (!header || typeof header !== 'string') return;

      const candidate = evaluateHeaderCandidate(header, index, fieldType, patterns, options);
      if (candidate.totalScore > 0) {
        candidates.push(candidate);
      }
    });

    debugInfo.candidateHeaders = candidates.slice(0, 3); // 上位3候補をデバッグ表示

    // 最高スコアの候補を選択
    if (candidates.length > 0) {
      const bestCandidate = candidates.reduce((best, current) =>
        current.totalScore > best.totalScore ? current : best
      );

      // 🎯 信頼度基準: 50以上で採用、80以上で高信頼
      if (bestCandidate.totalScore >= 50) {
        debugInfo.scoringDetails = bestCandidate.scoreBreakdown;
        debugInfo.finalSelection = {
          method: 'ai_enhanced_pattern',
          index: bestCandidate.index,
          score: bestCandidate.totalScore
        };

        return {
          index: bestCandidate.index,
          confidence: Math.min(bestCandidate.totalScore, 99), // 99%上限
          method: 'ai_enhanced_pattern',
          debug: debugInfo
        };
      }
    }

    // 3. 位置ベースフォールバック（改良版）
    if (options.allowPositionalFallback !== false) {
      const positionalResult = getSmartPositionalFallback(fieldType, headers, options);
      if (positionalResult.index !== -1) {
        debugInfo.searchMethods.push({ method: 'smart_positional', ...positionalResult });
        debugInfo.finalSelection = { method: 'smart_positional', ...positionalResult };
        return {
          index: positionalResult.index,
          confidence: positionalResult.confidence,
          method: 'smart_positional',
          debug: debugInfo
        };
      }
    }

    // 4. 見つからない場合
    debugInfo.finalSelection = { method: 'not_found', index: -1 };
    return { index: -1, confidence: 0, method: 'not_found', debug: debugInfo };

  } catch (error) {
    console.error('[ERROR] ColumnMappingService.resolveColumnIndex:', error.message || 'Unknown error');
    debugInfo.error = error.message;
    return { index: -1, confidence: 0, method: 'error', debug: debugInfo };
  }
}

/**
 * ヘッダーパターン定義（フロントエンドと統一）
 * @returns {Object} パターン定義オブジェクト
 */
function getHeaderPatterns() {
  return {
    timestamp: ['タイムスタンプ', 'timestamp', '投稿日時', '回答日時', '記録時刻'],
    email: ['メール', 'email', 'メールアドレス', 'mail', 'e-mail'],
    answer: ['回答', '答え', '意見', 'answer', 'opinion', 'response'],
    reason: ['理由', '根拠', 'reason', '説明', 'explanation'],
    class: ['クラス', '学年', 'class', '組', '学級'],
    name: ['名前', '氏名', 'name', '名', 'full_name'],

    // リアクション列（extractReactions用）
    understand: ['understand', '理解', 'UNDERSTAND'],
    like: ['like', 'いいね', 'LIKE'],
    curious: ['curious', '気になる', 'CURIOUS'],

    // ハイライト列（extractHighlight用）
    highlight: ['highlight', 'ハイライト', 'HIGHLIGHT']
  };
}

/**
 * 位置ベースフォールバック（Google Formsの典型的な列順序）
 * @param {string} fieldType - フィールドタイプ
 * @param {number|Array} columnCountOrRow - 列数またはデータ行
 * @returns {number|*} インデックスまたはフォールバック値
 */
function getPositionalFallback(fieldType, columnCountOrRow) {
  // Google Formsの典型的な列順序: タイムスタンプ, 質問1, 質問2, ...
  const typicalPositions = {
    timestamp: 0,
    answer: 1,
    reason: 2,
    class: 3,
    name: 4,
    email: 5
  };

  const position = typicalPositions[fieldType];

  // データ行が渡された場合（getPhysicalPositionFallback相当）
  if (Array.isArray(columnCountOrRow)) {
    const row = columnCountOrRow;
    if (!row || row.length === 0) return '';

    // 実際の行長に合わせて調整
    const adjustedPosition = position !== undefined ? Math.min(position, row.length - 1) : -1;
    return adjustedPosition !== -1 ? (row[adjustedPosition] || '') : '';
  }

  // 列数が渡された場合（従来のgetPositionalFallback）
  const columnCount = columnCountOrRow;
  return (position !== undefined && position < columnCount) ? position : -1;
}

/**
 * 🧠 AI強化ヘッダー候補評価システム - 多要素スコアリング
 * @param {string} header - ヘッダー文字列
 * @param {number} index - ヘッダーのインデックス
 * @param {string} fieldType - 対象フィールドタイプ
 * @param {Array} patterns - パターン配列
 * @param {Object} options - オプション設定
 * @returns {Object} { index, totalScore, scoreBreakdown }
 */
function evaluateHeaderCandidate(header, index, fieldType, patterns, options = {}) {
  const candidate = {
    index,
    header,
    totalScore: 0,
    scoreBreakdown: {
      patternMatch: 0,
      semanticSimilarity: 0,
      positionalScore: 0,
      contentValidation: 0,
      lengthPenalty: 0
    }
  };

  try {
    const normalizedHeader = header.toLowerCase().trim();

    // 1. 🎯 パターンマッチングスコア（重み40%）
    let patternScore = 0;
    let bestPattern = null;

    for (const pattern of patterns) {
      const normalizedPattern = pattern.toLowerCase().trim();

      // 完全一致: 100点
      if (normalizedHeader === normalizedPattern) {
        patternScore = 100;
        bestPattern = pattern;
        break;
      }

      // 部分一致の詳細評価
      if (normalizedHeader.includes(normalizedPattern)) {
        // 含有度によるスコア調整
        const containmentRatio = normalizedPattern.length / normalizedHeader.length;
        const containmentScore = 60 + (containmentRatio * 30); // 60-90点

        if (containmentScore > patternScore) {
          patternScore = containmentScore;
          bestPattern = pattern;
        }
      }

      // 逆含有（パターンがヘッダーを含む場合）
      if (normalizedPattern.includes(normalizedHeader)) {
        const reverseScore = 40 + ((normalizedHeader.length / normalizedPattern.length) * 20); // 40-60点
        if (reverseScore > patternScore) {
          patternScore = reverseScore;
          bestPattern = pattern;
        }
      }
    }

    candidate.scoreBreakdown.patternMatch = patternScore;
    candidate.bestPattern = bestPattern;

    // 2. 🔤 意味的類似性スコア（重み20%）
    candidate.scoreBreakdown.semanticSimilarity = calculateSemanticSimilarity(normalizedHeader, fieldType);

    // 3. 📍 位置的適合性スコア（重み20%）
    candidate.scoreBreakdown.positionalScore = calculatePositionalScore(index, fieldType);

    // 4. 📊 コンテンツ検証スコア（重み15%）
    if (options.sampleData && Array.isArray(options.sampleData) && options.sampleData.length > 0) {
      candidate.scoreBreakdown.contentValidation = validateContentType(options.sampleData, index, fieldType);
    }

    // 5. 📏 長さペナルティ（重み5%）
    candidate.scoreBreakdown.lengthPenalty = calculateLengthPenalty(normalizedHeader);

    // 🧮 適応的重み付き総合スコア計算
    const adaptiveWeights = getAdaptiveWeights(fieldType, options);
    candidate.totalScore = Math.round(
      (candidate.scoreBreakdown.patternMatch * adaptiveWeights.patternMatch) +
      (candidate.scoreBreakdown.semanticSimilarity * adaptiveWeights.semanticSimilarity) +
      (candidate.scoreBreakdown.positionalScore * adaptiveWeights.positionalScore) +
      (candidate.scoreBreakdown.contentValidation * adaptiveWeights.contentValidation) +
      (candidate.scoreBreakdown.lengthPenalty * adaptiveWeights.lengthPenalty)
    );

    return candidate;

  } catch (error) {
    console.error('evaluateHeaderCandidate エラー:', error.message);
    return { index, header, totalScore: 0, scoreBreakdown: {}, error: error.message };
  }
}

/**
 * 🔤 意味的類似性計算（統計的文字列距離アルゴリズム強化版）
 * @param {string} header - 正規化済みヘッダー
 * @param {string} fieldType - フィールドタイプ
 * @returns {number} 類似性スコア（0-100）
 */
function calculateSemanticSimilarity(header, fieldType) {
  // フィールドタイプ別の関連キーワード（優先度付き）
  const semanticKeywords = {
    timestamp: [
      { term: '日時', weight: 30, priority: 1 },
      { term: 'timestamp', weight: 35, priority: 1 },
      { term: 'タイムスタンプ', weight: 35, priority: 1 },
      { term: '時間', weight: 25, priority: 2 },
      { term: '投稿', weight: 20, priority: 2 },
      { term: 'time', weight: 25, priority: 2 },
      { term: 'date', weight: 25, priority: 2 }
    ],
    email: [
      { term: 'email', weight: 35, priority: 1 },
      { term: 'メール', weight: 35, priority: 1 },
      { term: 'mail', weight: 30, priority: 1 },
      { term: 'アドレス', weight: 25, priority: 2 },
      { term: 'address', weight: 25, priority: 2 }
    ],
    answer: [
      { term: '回答', weight: 35, priority: 1 },
      { term: 'answer', weight: 35, priority: 1 },
      { term: '答え', weight: 30, priority: 1 },
      { term: '意見', weight: 25, priority: 2 },
      { term: 'response', weight: 30, priority: 2 },
      { term: 'opinion', weight: 25, priority: 2 }
    ],
    reason: [
      { term: '理由', weight: 35, priority: 1 },
      { term: 'reason', weight: 35, priority: 1 },
      { term: '根拠', weight: 30, priority: 1 },
      { term: '説明', weight: 25, priority: 2 },
      { term: 'explain', weight: 25, priority: 2 },
      { term: 'why', weight: 30, priority: 2 }
    ],
    class: [
      { term: 'クラス', weight: 35, priority: 1 },
      { term: 'class', weight: 35, priority: 1 },
      { term: '学年', weight: 30, priority: 1 },
      { term: '組', weight: 25, priority: 2 },
      { term: 'grade', weight: 25, priority: 2 }
    ],
    name: [
      { term: '名前', weight: 35, priority: 1 },
      { term: 'name', weight: 35, priority: 1 },
      { term: '氏名', weight: 30, priority: 1 },
      { term: '名', weight: 20, priority: 2 },
      { term: 'user', weight: 15, priority: 3 }
    ]
  };

  const keywords = semanticKeywords[fieldType] || [];
  let totalScore = 0;
  let bestStringDistance = 0;

  // 🎯 段階1: 高精度文字列類似度計算
  keywords.forEach(keywordObj => {
    const { term, weight, priority } = keywordObj;

    // 完全一致チェック
    if (header === term.toLowerCase()) {
      totalScore += weight;
      return;
    }

    // レーベンシュタイン距離による類似度
    const levenshteinSimilarity = calculateLevenshteinSimilarity(header, term.toLowerCase());
    if (levenshteinSimilarity > 0.7) {
      const distanceScore = Math.round(weight * levenshteinSimilarity * (priority === 1 ? 1.0 : 0.8));
      totalScore += distanceScore;
      bestStringDistance = Math.max(bestStringDistance, levenshteinSimilarity * 100);
    }

    // ジャロ・ウィンクラー類似度（短い文字列に効果的）
    if (term.length <= 8) {
      const jaroSimilarity = calculateJaroWinklerSimilarity(header, term.toLowerCase());
      if (jaroSimilarity > 0.8) {
        const jaroScore = Math.round(weight * jaroSimilarity * 0.9);
        totalScore += Math.max(0, jaroScore);
      }
    }

    // n-gram類似度（部分一致向上）
    const ngramSimilarity = calculateNGramSimilarity(header, term.toLowerCase(), 2);
    if (ngramSimilarity > 0.6) {
      const ngramScore = Math.round(weight * ngramSimilarity * 0.7);
      totalScore += Math.max(0, ngramScore);
    }

    // 従来の含有チェック（基本ベース）
    if (header.includes(term.toLowerCase())) {
      const containmentScore = priority === 1 ? weight * 0.8 : weight * 0.6;
      totalScore += Math.round(containmentScore);
    }
  });

  // 🎯 段階2: フィールドタイプ特有パターン検証（統計的パターン強化）
  const patternScore = calculateFieldSpecificPatterns(header, fieldType);
  totalScore += patternScore;

  // 🎯 段階3: 統計的正規化・異常値検出
  const normalizedScore = normalizeSemanticScore(totalScore, bestStringDistance, keywords.length);

  return Math.min(normalizedScore, 100);
}

/**
 * 🧮 レーベンシュタイン距離による類似度計算
 * @param {string} str1 - 文字列1
 * @param {string} str2 - 文字列2
 * @returns {number} 類似度（0-1）
 */
function calculateLevenshteinSimilarity(str1, str2) {
  try {
    if (str1 === str2) return 1.0;
    if (!str1 || !str2) return 0.0;

    const maxLength = Math.max(str1.length, str2.length);
    if (maxLength === 0) return 1.0;

    const distance = calculateLevenshteinDistance(str1, str2);
    return (maxLength - distance) / maxLength;
  } catch (error) {
    console.warn('Levenshtein similarity calculation error:', error.message);
    return 0.0;
  }
}

/**
 * 🧮 レーベンシュタイン距離計算（動的プログラミング）
 * @param {string} str1 - 文字列1
 * @param {string} str2 - 文字列2
 * @returns {number} 編集距離
 */
function calculateLevenshteinDistance(str1, str2) {
  const matrix = Array.from({length: str1.length + 1}, () =>
    Array.from({length: str2.length + 1}, () => 0)
  );

  // 初期化
  for (let i = 0; i <= str1.length; i++) matrix[i][0] = i;
  for (let j = 0; j <= str2.length; j++) matrix[0][j] = j;

  // 動的プログラミング
  for (let i = 1; i <= str1.length; i++) {
    for (let j = 1; j <= str2.length; j++) {
      if (str1[i - 1] === str2[j - 1]) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j] + 1,     // 削除
          matrix[i][j - 1] + 1,     // 挿入
          matrix[i - 1][j - 1] + 1  // 置換
        );
      }
    }
  }

  return matrix[str1.length][str2.length];
}

/**
 * 🧮 ジャロ・ウィンクラー類似度計算
 * @param {string} str1 - 文字列1
 * @param {string} str2 - 文字列2
 * @returns {number} 類似度（0-1）
 */
function calculateJaroWinklerSimilarity(str1, str2) {
  try {
    if (str1 === str2) return 1.0;
    if (!str1 || !str2) return 0.0;

    const maxDistance = Math.floor(Math.max(str1.length, str2.length) / 2) - 1;
    const matches1 = Array(str1.length).fill(false);
    const matches2 = Array(str2.length).fill(false);
    let matches = 0;
    let transpositions = 0;

    // マッチング検出
    for (let i = 0; i < str1.length; i++) {
      const start = Math.max(0, i - maxDistance);
      const end = Math.min(i + maxDistance + 1, str2.length);

      for (let j = start; j < end; j++) {
        if (matches2[j] || str1[i] !== str2[j]) continue;
        matches1[i] = matches2[j] = true;
        matches++;
        break;
      }
    }

    if (matches === 0) return 0.0;

    // 転置カウント
    let k = 0;
    for (let i = 0; i < str1.length; i++) {
      if (!matches1[i]) continue;
      while (!matches2[k]) k++;
      if (str1[i] !== str2[k]) transpositions++;
      k++;
    }

    // ジャロ類似度計算
    const jaro = (matches / str1.length + matches / str2.length +
                  (matches - transpositions / 2) / matches) / 3;

    // ウィンクラー調整（共通プレフィックス）
    let prefixLength = 0;
    for (let i = 0; i < Math.min(str1.length, str2.length, 4); i++) {
      if (str1[i] === str2[i]) prefixLength++;
      else break;
    }

    return jaro + (0.1 * prefixLength * (1 - jaro));
  } catch (error) {
    console.warn('Jaro-Winkler similarity calculation error:', error.message);
    return 0.0;
  }
}

/**
 * 🧮 n-gram類似度計算
 * @param {string} str1 - 文字列1
 * @param {string} str2 - 文字列2
 * @param {number} n - n-gramサイズ
 * @returns {number} 類似度（0-1）
 */
function calculateNGramSimilarity(str1, str2, n = 2) {
  try {
    if (str1 === str2) return 1.0;
    if (!str1 || !str2) return 0.0;

    const ngrams1 = generateNGrams(str1, n);
    const ngrams2 = generateNGrams(str2, n);

    if (ngrams1.length === 0 && ngrams2.length === 0) return 1.0;
    if (ngrams1.length === 0 || ngrams2.length === 0) return 0.0;

    // ジャッカード係数計算
    const set1 = new Set(ngrams1);
    const set2 = new Set(ngrams2);
    const intersection = new Set([...set1].filter(x => set2.has(x)));
    const union = new Set([...set1, ...set2]);

    return intersection.size / union.size;
  } catch (error) {
    console.warn('N-gram similarity calculation error:', error.message);
    return 0.0;
  }
}

/**
 * 🔤 n-gram生成
 * @param {string} str - 文字列
 * @param {number} n - n-gramサイズ
 * @returns {Array} n-gram配列
 */
function generateNGrams(str, n) {
  if (str.length < n) return [str];
  const ngrams = [];
  for (let i = 0; i <= str.length - n; i++) {
    ngrams.push(str.slice(i, i + n));
  }
  return ngrams;
}

/**
 * 📊 フィールド特有パターン検証（統計的パターン強化版）
 * @param {string} header - ヘッダー文字列
 * @param {string} fieldType - フィールドタイプ
 * @returns {number} パターンスコア（0-30）
 */
function calculateFieldSpecificPatterns(header, fieldType) {
  let patternScore = 0;

  try {
    switch (fieldType) {
      case 'email':
        // より厳密なメール形式検証
        if (/@/.test(header)) patternScore += 15;
        if (/mail|メール/.test(header)) patternScore += 10;
        if (/address|アドレス/.test(header)) patternScore += 8;
        break;

      case 'timestamp':
        // タイムスタンプパターンの多様性対応
        if (/\d{4}[-/]\d{1,2}[-/]\d{1,2}/.test(header)) patternScore += 20;
        if (/時|time|date|日/.test(header)) patternScore += 12;
        if (/記録|投稿|作成/.test(header)) patternScore += 10;
        break;

      case 'class':
        // クラス表記の多様性
        if (/[1-6][年組AB]|[1-9]\d*[組年クラス]/.test(header)) patternScore += 18;
        if (/\d+/.test(header)) patternScore += 10;
        if (/class|クラス|学年|組/.test(header)) patternScore += 8;
        break;

      case 'name':
        // 名前フィールドの特徴
        if (header.length >= 2 && header.length <= 10) patternScore += 8;
        if (/名|name|user/.test(header)) patternScore += 12;
        break;

      case 'answer':
        // 回答フィールドの特徴
        if (/回答|答|answer|response|意見/.test(header)) patternScore += 15;
        if (header.length >= 2 && header.length <= 15) patternScore += 5;
        break;

      case 'reason':
        // 理由フィールドの特徴
        if (/理由|根拠|reason|why|説明/.test(header)) patternScore += 15;
        if (header.includes('？') || header.includes('?')) patternScore += 8;
        break;
    }

    return Math.min(patternScore, 30);
  } catch (error) {
    console.warn('Field-specific pattern calculation error:', error.message);
    return 0;
  }
}

/**
 * 📊 意味スコア正規化・統計的調整
 * @param {number} rawScore - 生スコア
 * @param {number} bestDistance - 最高文字列距離
 * @param {number} keywordCount - キーワード数
 * @returns {number} 正規化スコア（0-100）
 */
function normalizeSemanticScore(rawScore, bestDistance, keywordCount) {
  try {
    // 基本正規化（キーワード数基準）
    const maxPossibleScore = keywordCount * 35; // 最高重み35点×キーワード数
    let normalizedScore = maxPossibleScore > 0 ? (rawScore / maxPossibleScore) * 100 : 0;

    // 文字列距離ボーナス（高精度一致への追加重み）
    if (bestDistance > 80) {
      normalizedScore *= 1.1; // 10%ボーナス
    } else if (bestDistance > 90) {
      normalizedScore *= 1.15; // 15%ボーナス
    }

    // 統計的上限・下限調整
    return Math.max(0, Math.min(100, Math.round(normalizedScore)));
  } catch (error) {
    console.warn('Semantic score normalization error:', error.message);
    return Math.min(rawScore, 100);
  }
}

/**
 * 📍 位置的適合性スコア計算
 * @param {number} index - ヘッダーのインデックス
 * @param {string} fieldType - フィールドタイプ
 * @returns {number} 位置スコア（0-100）
 */
function calculatePositionalScore(index, fieldType) {
  // Google Forms典型的位置
  const idealPositions = {
    timestamp: 0,
    answer: 1,
    reason: 2,
    class: 3,
    name: 4,
    email: 5
  };

  const idealPos = idealPositions[fieldType];
  if (idealPos === undefined) return 50; // 不明フィールドは中立

  // 位置差によるペナルティ（距離に比例）
  const positionDiff = Math.abs(index - idealPos);
  let positionScore = Math.max(0, 100 - (positionDiff * 15));

  // 特殊調整
  switch (fieldType) {
    case 'timestamp':
      // タイムスタンプは最初の列が最も理想的
      positionScore = index === 0 ? 100 : Math.max(20, 100 - (index * 20));
      break;
    case 'answer':
      // 回答は1-3列目が理想的
      if (index >= 1 && index <= 3) positionScore = Math.max(positionScore, 80);
      break;
  }

  return Math.min(positionScore, 100);
}

/**
 * 📊 統計的コンテンツ検証システム（ベイズ推定＋信頼区間）
 * @param {Array} sampleData - サンプルデータ配列
 * @param {number} index - 列インデックス
 * @param {string} fieldType - フィールドタイプ
 * @returns {number} コンテンツ検証スコア（0-100）
 */
function validateContentType(sampleData, index, fieldType) {
  try {
    if (!sampleData || sampleData.length === 0) return 0;

    // 🎯 段階1: 統計的サンプルサイズ最適化
    const optimalSampleSize = calculateOptimalSampleSize(sampleData.length, fieldType);
    const samples = extractStatisticalSamples(sampleData, index, optimalSampleSize);

    if (samples.length === 0) return 0;

    // 🎯 段階2: ベイズ推定ベース検証
    const bayesianResult = performBayesianValidation(samples, fieldType);

    // 🎯 段階3: 統計的信頼区間計算
    const confidenceInterval = calculateValidationConfidenceInterval(
      bayesianResult.validCount,
      samples.length,
      0.95 // 95%信頼区間
    );

    // 🎯 段階4: 異常値検出・除去
    const filteredSamples = detectAndRemoveOutliers(samples, fieldType);
    const outlierPenalty = (samples.length - filteredSamples.length) * 5; // 異常値ペナルティ

    // 🎯 段階5: 統計的スコア計算
    const baseScore = calculateStatisticalValidityScore(
      bayesianResult.posteriorProbability,
      confidenceInterval,
      samples.length,
      optimalSampleSize
    );

    // 異常値ペナルティ適用
    const finalScore = Math.max(0, Math.min(100, baseScore - outlierPenalty));

    return Math.round(finalScore);

  } catch (error) {
    console.error('validateContentType エラー:', error.message);
    return 0;
  }
}

/**
 * 📊 統計的サンプルサイズ最適化
 * @param {number} totalDataSize - 全データサイズ
 * @param {string} fieldType - フィールドタイプ
 * @returns {number} 最適サンプルサイズ
 */
function calculateOptimalSampleSize(totalDataSize, fieldType) {
  try {
    // フィールドタイプ別の必要精度
    const precisionRequirements = {
      email: 0.95,      // メール形式は高精度が必要
      timestamp: 0.90,  // 日時形式も高精度
      answer: 0.80,     // 回答は中程度
      reason: 0.80,     // 理由も中程度
      class: 0.85,      // クラスは比較的高精度
      name: 0.75        // 名前は中程度
    };

    const requiredPrecision = precisionRequirements[fieldType] || 0.80;

    // サンプルサイズ計算（Wilson Score Interval基準）
    // n = (z²p(1-p)) / e² (z=1.96 for 95% confidence, e=margin of error)
    const z = 1.96; // 95%信頼区間
    const p = requiredPrecision;
    const marginError = 0.1; // 10%誤差許容

    const theoreticalSampleSize = Math.ceil(
      (z * z * p * (1 - p)) / (marginError * marginError)
    );

    // 実データサイズを考慮した調整
    const minSample = Math.max(3, Math.min(10, Math.floor(totalDataSize * 0.3)));
    const maxSample = Math.min(20, totalDataSize);

    return Math.max(minSample, Math.min(maxSample, theoreticalSampleSize));

  } catch (error) {
    console.warn('Optimal sample size calculation error:', error.message);
    return Math.min(8, totalDataSize); // フォールバック
  }
}

/**
 * 📊 統計的サンプル抽出
 * @param {Array} sampleData - 元データ配列
 * @param {number} index - 列インデックス
 * @param {number} sampleSize - サンプルサイズ
 * @returns {Array} 抽出されたサンプル
 */
function extractStatisticalSamples(sampleData, index, sampleSize) {
  const samples = [];
  const dataSize = sampleData.length;

  try {
    // 等間隔サンプリング（systematic sampling）で偏りを減少
    const interval = Math.max(1, Math.floor(dataSize / sampleSize));

    for (let i = 0; i < dataSize && samples.length < sampleSize; i += interval) {
      const row = sampleData[i];
      if (!row || !Array.isArray(row) || index >= row.length) continue;

      const cellValue = row[index];
      if (cellValue !== null && cellValue !== undefined) {
        const cleanValue = String(cellValue).trim();
        if (cleanValue.length > 0) {
          samples.push(cleanValue);
        }
      }
    }

    return samples;
  } catch (error) {
    console.warn('Statistical sampling error:', error.message);
    return [];
  }
}

/**
 * 🧠 ベイズ推定ベース検証
 * @param {Array} samples - サンプル配列
 * @param {string} fieldType - フィールドタイプ
 * @returns {Object} { validCount, posteriorProbability, priorProbability }
 */
function performBayesianValidation(samples, fieldType) {
  try {
    // 🎯 事前確率（Prior Probability）設定
    const priorProbabilities = {
      email: 0.85,      // メール形式は比較的明確
      timestamp: 0.90,  // 日時形式は識別しやすい
      answer: 0.70,     // 回答は多様性がある
      reason: 0.70,     // 理由も多様
      class: 0.80,      // クラスは比較的明確
      name: 0.65        // 名前は最も多様
    };

    const priorProbability = priorProbabilities[fieldType] || 0.70;

    // 🎯 尤度計算（Likelihood）
    let validCount = 0;
    const validationResults = samples.map(sample => validateSingleSample(sample, fieldType));
    validCount = validationResults.filter(result => result.isValid).length;

    const observedSuccessRate = samples.length > 0 ? validCount / samples.length : 0;

    // 🎯 ベイズ更新（Beta分布での事後確率計算）
    // Beta(α + successes, β + failures)
    const alpha = priorProbability * 10; // 事前情報の重み
    const beta = (1 - priorProbability) * 10;

    const posteriorAlpha = alpha + validCount;
    const posteriorBeta = beta + (samples.length - validCount);

    // 事後確率の期待値
    const posteriorProbability = posteriorAlpha / (posteriorAlpha + posteriorBeta);

    return {
      validCount,
      posteriorProbability,
      priorProbability,
      observedSuccessRate,
      confidenceInValidation: Math.abs(posteriorProbability - 0.5) * 2 // 0.5からの距離で信頼度
    };

  } catch (error) {
    console.warn('Bayesian validation error:', error.message);
    return { validCount: 0, posteriorProbability: 0.5, priorProbability: 0.5 };
  }
}

/**
 * 🔍 単一サンプル検証（高度パターンマッチング）
 * @param {string} sample - サンプル値
 * @param {string} fieldType - フィールドタイプ
 * @returns {Object} { isValid, confidence, reasons }
 */
function validateSingleSample(sample, fieldType) {
  const result = {
    isValid: false,
    confidence: 0,
    reasons: []
  };

  try {
    switch (fieldType) {
      case 'email': {
        // RFC5322準拠の詳細メール検証
        const emailRegex = /^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)*$/;
        if (emailRegex.test(sample)) {
          result.isValid = true;
          result.confidence = 95;
          result.reasons.push('正規表現マッチ');
        }
        break;
      }

      case 'timestamp': {
        // 多様なタイムスタンプ形式対応
        const timePatterns = [
          /^\d{4}[-/]\d{1,2}[-/]\d{1,2}/,  // YYYY-MM-DD, YYYY/MM/DD
          /^\d{1,2}[-/]\d{1,2}[-/]\d{4}/,  // MM-DD-YYYY, MM/DD/YYYY
          /^\d{4}年\d{1,2}月\d{1,2}日/,     // 和暦形式
          /^\d{1,2}:\d{2}(:\d{2})?/          // 時刻形式
        ];

        const parsedDate = new Date(sample);
        const isValidDate = !isNaN(parsedDate.getTime()) && parsedDate.getTime() > 0;
        const matchesPattern = timePatterns.some(pattern => pattern.test(sample));

        if (isValidDate && matchesPattern) {
          result.isValid = true;
          result.confidence = 90;
          result.reasons.push('日時パース成功', 'パターンマッチ');
        } else if (isValidDate) {
          result.isValid = true;
          result.confidence = 70;
          result.reasons.push('日時パース成功');
        }
        break;
      }

      case 'answer':
      case 'reason': {
        // 文章品質評価
        const hasText = sample.length >= 3;
        const hasVariedChars = /[あ-ん]|[ア-ン]|[a-zA-Z]/.test(sample);
        const hasReasonableLength = sample.length <= 500;
        const notOnlyNumbers = !/^\d+$/.test(sample);

        if (hasText && hasVariedChars && hasReasonableLength && notOnlyNumbers) {
          result.isValid = true;
          result.confidence = 80;
          result.reasons.push('文章構造適切');
        }
        break;
      }

      case 'name': {
        // 名前パターン検証
        const nameLength = sample.length >= 1 && sample.length <= 30;
        const noSpecialChars = !/[@#$%^&*()_+={}[\]|\\:";'<>?,./]/.test(sample);
        const notOnlySpaces = /\S/.test(sample);

        if (nameLength && noSpecialChars && notOnlySpaces) {
          result.isValid = true;
          result.confidence = 75;
          result.reasons.push('名前形式適合');
        }
        break;
      }

      case 'class': {
        // クラス表記検証
        const classPatterns = [
          /^[1-6][年組AB]?$/,               // 1年A組形式
          /^[1-9]\d?[組年クラス]?$/,       // 数字+組/年/クラス
          /^\d+$/                            // 純粋な数字
        ];

        if (classPatterns.some(pattern => pattern.test(sample))) {
          result.isValid = true;
          result.confidence = 85;
          result.reasons.push('クラス形式マッチ');
        }
        break;
      }

      default:
        result.isValid = sample.length >= 1;
        result.confidence = 50;
        result.reasons.push('基本検証のみ');
    }

    return result;

  } catch (error) {
    console.warn('Single sample validation error:', error.message);
    return { isValid: false, confidence: 0, reasons: ['検証エラー'] };
  }
}

/**
 * 📊 統計的信頼区間計算
 * @param {number} successes - 成功数
 * @param {number} trials - 試行数
 * @param {number} confidenceLevel - 信頼水準（0-1）
 * @returns {Object} { lowerBound, upperBound, margin }
 */
function calculateValidationConfidenceInterval(successes, trials, confidenceLevel = 0.95) {
  try {
    if (trials === 0) return { lowerBound: 0, upperBound: 0, margin: 0 };

    const p = successes / trials;
    const z = confidenceLevel === 0.95 ? 1.96 : 2.58; // 95% or 99%

    // Wilson Score Interval（より正確）
    const denominator = 1 + (z * z) / trials;
    const center = (p + (z * z) / (2 * trials)) / denominator;
    const margin = (z / denominator) * Math.sqrt((p * (1 - p) / trials) + ((z * z) / (4 * trials * trials)));

    return {
      lowerBound: Math.max(0, center - margin),
      upperBound: Math.min(1, center + margin),
      margin
    };

  } catch (error) {
    console.warn('Confidence interval calculation error:', error.message);
    return { lowerBound: 0, upperBound: 1, margin: 0.5 };
  }
}

/**
 * 🚫 異常値検出・除去
 * @param {Array} samples - サンプル配列
 * @param {string} fieldType - フィールドタイプ
 * @returns {Array} 異常値を除去したサンプル配列
 */
function detectAndRemoveOutliers(samples, fieldType) {
  try {
    const outlierDetectors = {
      email: (sample) => sample.length > 100 || sample.split('@').length !== 2,
      timestamp: (sample) => sample.length > 50 || sample.includes('undefined'),
      answer: (sample) => sample.length > 1000 || /^(.)\1{10,}$/.test(sample), // 同じ文字の繰り返し
      reason: (sample) => sample.length > 1000 || /^(.)\1{10,}$/.test(sample),
      class: (sample) => sample.length > 20 || /[!@#$%^&*()_+={}[\]|\\:";'<>?,./]/.test(sample),
      name: (sample) => sample.length > 50 || /\d{5,}/.test(sample) // 長い数字列
    };

    const detector = outlierDetectors[fieldType] || (() => false);
    return samples.filter(sample => !detector(sample));

  } catch (error) {
    console.warn('Outlier detection error:', error.message);
    return samples;
  }
}

/**
 * 📊 統計的妥当性スコア計算
 * @param {number} posteriorProbability - 事後確率
 * @param {Object} confidenceInterval - 信頼区間
 * @param {number} actualSampleSize - 実際のサンプルサイズ
 * @param {number} optimalSampleSize - 最適サンプルサイズ
 * @returns {number} 統計的スコア（0-100）
 */
function calculateStatisticalValidityScore(posteriorProbability, confidenceInterval, actualSampleSize, optimalSampleSize) {
  try {
    // 🎯 ベイズ確率スコア（重み60%）
    const bayesScore = posteriorProbability * 100;

    // 🎯 信頼区間品質スコア（重み25%）
    const intervalWidth = confidenceInterval.upperBound - confidenceInterval.lowerBound;
    const precisionScore = Math.max(0, (1 - intervalWidth) * 100); // 狭い区間ほど高スコア

    // 🎯 サンプルサイズ適切性スコア（重み15%）
    const sampleRatio = actualSampleSize / optimalSampleSize;
    const sampleScore = sampleRatio >= 1 ? 100 : Math.max(50, sampleRatio * 100);

    // 重み付き統合スコア
    const totalScore = (bayesScore * 0.60) + (precisionScore * 0.25) + (sampleScore * 0.15);

    return Math.max(0, Math.min(100, totalScore));

  } catch (error) {
    console.warn('Statistical validity score calculation error:', error.message);
    return 50; // 中立値
  }
}

/**
 * 🧠 適応的重み調整システム（学習ベース最適化）
 * @param {string} fieldType - フィールドタイプ
 * @param {Object} options - オプション設定
 * @returns {Object} 適応的重みオブジェクト
 */
function getAdaptiveWeights(fieldType, options = {}) {
  try {
    // 🎯 デフォルト重み（ベースライン）
    const baseWeights = {
      patternMatch: 0.40,
      semanticSimilarity: 0.20,
      positionalScore: 0.20,
      contentValidation: 0.15,
      lengthPenalty: 0.05
    };

    // 🎯 段階1: フィールドタイプ別重み調整
    const fieldSpecificWeights = getFieldSpecificWeights(fieldType, baseWeights);

    // 🎯 段階2: データ環境適応調整
    const environmentAdaptedWeights = adaptWeightsToEnvironment(
      fieldSpecificWeights,
      options.sampleData || [],
      options.headers || []
    );

    // 🎯 段階3: 過去成功パターン学習調整（キャッシュベース）
    const learningAdaptedWeights = applyLearningBasedAdjustments(
      environmentAdaptedWeights,
      fieldType
    );

    // 🎯 段階4: 重み正規化（合計が1.0になるよう調整）
    const normalizedWeights = normalizeWeights(learningAdaptedWeights);

    return normalizedWeights;

  } catch (error) {
    console.warn('Adaptive weights calculation error:', error.message);
    return {
      patternMatch: 0.40,
      semanticSimilarity: 0.20,
      positionalScore: 0.20,
      contentValidation: 0.15,
      lengthPenalty: 0.05
    };
  }
}

/**
 * 📊 フィールドタイプ別重み特化
 * @param {string} fieldType - フィールドタイプ
 * @param {Object} baseWeights - ベース重み
 * @returns {Object} 特化重み
 */
function getFieldSpecificWeights(fieldType, baseWeights) {
  const weightAdjustments = {
    email: {
      // メール形式は明確なパターンがあるため、パターンマッチングを重視
      patternMatch: 0.50,      // +0.10
      semanticSimilarity: 0.15, // -0.05
      positionalScore: 0.15,   // -0.05
      contentValidation: 0.15, // 変更なし
      lengthPenalty: 0.05      // 変更なし
    },

    timestamp: {
      // タイムスタンプは位置（最初の列）とパターンが重要
      patternMatch: 0.35,      // -0.05
      semanticSimilarity: 0.20, // 変更なし
      positionalScore: 0.25,   // +0.05
      contentValidation: 0.15, // 変更なし
      lengthPenalty: 0.05      // 変更なし
    },

    answer: {
      // 回答は内容とパターンの組み合わせが重要
      patternMatch: 0.35,      // -0.05
      semanticSimilarity: 0.25, // +0.05
      positionalScore: 0.15,   // -0.05
      contentValidation: 0.20, // +0.05
      lengthPenalty: 0.05      // 変更なし
    },

    reason: {
      // 理由も回答と似た調整
      patternMatch: 0.35,      // -0.05
      semanticSimilarity: 0.25, // +0.05
      positionalScore: 0.15,   // -0.05
      contentValidation: 0.20, // +0.05
      lengthPenalty: 0.05      // 変更なし
    },

    class: {
      // クラスはパターンと位置が重要
      patternMatch: 0.45,      // +0.05
      semanticSimilarity: 0.15, // -0.05
      positionalScore: 0.25,   // +0.05
      contentValidation: 0.10, // -0.05
      lengthPenalty: 0.05      // 変更なし
    },

    name: {
      // 名前は最も多様なため、内容検証を重視
      patternMatch: 0.30,      // -0.10
      semanticSimilarity: 0.20, // 変更なし
      positionalScore: 0.15,   // -0.05
      contentValidation: 0.25, // +0.10
      lengthPenalty: 0.10      // +0.05（長さも重要）
    }
  };

  return weightAdjustments[fieldType] || baseWeights;
}

/**
 * 🌍 データ環境適応調整
 * @param {Object} weights - 現在の重み
 * @param {Array} sampleData - サンプルデータ
 * @param {Array} headers - ヘッダー配列
 * @returns {Object} 環境適応重み
 */
function adaptWeightsToEnvironment(weights, sampleData, headers) {
  try {
    const adaptedWeights = { ...weights };

    // 🔍 データ環境分析
    const environment = analyzeDataEnvironment(sampleData, headers);

    // 🎯 環境特性に基づく重み調整

    // 小規模データセット（<10行）: コンテンツ検証の重みを下げる
    if (environment.isSmallDataset) {
      adaptedWeights.contentValidation *= 0.7;
      adaptedWeights.patternMatch *= 1.1;
    }

    // ヘッダー品質が低い場合: 意味的類似性を重視
    if (environment.hasLowQualityHeaders) {
      adaptedWeights.semanticSimilarity *= 1.2;
      adaptedWeights.patternMatch *= 0.9;
    }

    // 非標準的なヘッダー構造: 位置ベース判定の重みを下げる
    if (environment.hasNonStandardStructure) {
      adaptedWeights.positionalScore *= 0.8;
      adaptedWeights.semanticSimilarity *= 1.1;
    }

    // 多様なデータタイプが混在: コンテンツ検証を重視
    if (environment.hasMixedDataTypes) {
      adaptedWeights.contentValidation *= 1.3;
      adaptedWeights.positionalScore *= 0.9;
    }

    return adaptedWeights;

  } catch (error) {
    console.warn('Environment adaptation error:', error.message);
    return weights;
  }
}

/**
 * 🔍 データ環境分析
 * @param {Array} sampleData - サンプルデータ
 * @param {Array} headers - ヘッダー配列
 * @returns {Object} 環境特性オブジェクト
 */
function analyzeDataEnvironment(sampleData, headers) {
  const analysis = {
    isSmallDataset: false,
    hasLowQualityHeaders: false,
    hasNonStandardStructure: false,
    hasMixedDataTypes: false
  };

  try {
    // 小規模データセット判定
    analysis.isSmallDataset = sampleData.length < 10;

    // ヘッダー品質評価
    if (headers && headers.length > 0) {
      const avgHeaderLength = headers.reduce((sum, h) => sum + (h ? h.length : 0), 0) / headers.length;
      const shortHeaders = headers.filter(h => h && h.length < 3).length;
      const longHeaders = headers.filter(h => h && h.length > 20).length;

      analysis.hasLowQualityHeaders =
        avgHeaderLength < 4 ||
        (shortHeaders / headers.length) > 0.3 ||
        (longHeaders / headers.length) > 0.2;
    }

    // 非標準構造判定（タイムスタンプが最初の列でない等）
    if (headers && headers.length > 0) {
      const firstHeader = headers[0] ? headers[0].toLowerCase() : '';
      const hasTimestampFirst = /time|日時|timestamp/.test(firstHeader);
      analysis.hasNonStandardStructure = !hasTimestampFirst && headers.length > 3;
    }

    // データタイプ多様性分析
    if (sampleData.length > 0) {
      const firstRow = sampleData[0] || [];
      const dataTypes = firstRow.map(cell => {
        if (!cell) return 'empty';
        const str = String(cell);
        if (/^\d+$/.test(str)) return 'number';
        if (/@/.test(str)) return 'email';
        if (/\d{4}[-/]\d{1,2}[-/]\d{1,2}/.test(str)) return 'date';
        return 'text';
      });

      const uniqueTypes = new Set(dataTypes);
      analysis.hasMixedDataTypes = uniqueTypes.size > 3;
    }

    return analysis;

  } catch (error) {
    console.warn('Data environment analysis error:', error.message);
    return analysis;
  }
}

/**
 * 🎓 過去成功パターン学習調整（シンプルキャッシュベース）
 * @param {Object} weights - 現在の重み
 * @param {string} fieldType - フィールドタイプ
 * @returns {Object} 学習調整済み重み
 */
function applyLearningBasedAdjustments(weights, fieldType) {
  try {
    // 🎯 学習データ取得（キャッシュから）
    const learningData = getLearningData(fieldType);

    if (!learningData || learningData.samples < 5) {
      return weights; // 学習データが不足している場合はそのまま
    }

    const adjustedWeights = { ...weights };

    // 🎯 成功パターンに基づく重み調整
    const successPatterns = learningData.successPatterns || {};

    // パターンマッチング成功率が高い場合
    if (successPatterns.patternMatchSuccessRate > 0.8) {
      adjustedWeights.patternMatch *= 1.1;
      adjustedWeights.semanticSimilarity *= 0.95;
    }

    // 意味的類似性による成功が多い場合
    if (successPatterns.semanticSuccessRate > 0.7) {
      adjustedWeights.semanticSimilarity *= 1.1;
      adjustedWeights.patternMatch *= 0.95;
    }

    // 位置ベース予測が外れることが多い場合
    if (successPatterns.positionalFailureRate > 0.6) {
      adjustedWeights.positionalScore *= 0.8;
      adjustedWeights.semanticSimilarity *= 1.1;
    }

    return adjustedWeights;

  } catch (error) {
    console.warn('Learning-based adjustment error:', error.message);
    return weights;
  }
}

/**
 * 📚 学習データ取得（シンプル実装）
 * @param {string} fieldType - フィールドタイプ
 * @returns {Object|null} 学習データ
 */
function getLearningData(fieldType) {
  try {
    // 🎯 PropertiesServiceから学習データを取得
    const cacheKey = `aiColumnLearning_${fieldType}`;
    const cachedData = PropertiesService.getScriptProperties().getProperty(cacheKey);

    if (cachedData) {
      return JSON.parse(cachedData);
    }

    return null;
  } catch (error) {
    console.warn('Learning data retrieval error:', error.message);
    return null;
  }
}

/**
 * 📊 重み正規化（合計1.0に調整）
 * @param {Object} weights - 重みオブジェクト
 * @returns {Object} 正規化済み重み
 */
function normalizeWeights(weights) {
  try {
    const total = Object.values(weights).reduce((sum, w) => sum + w, 0);

    if (total === 0) {
      // 全て0の場合はデフォルトに戻す
      return {
        patternMatch: 0.40,
        semanticSimilarity: 0.20,
        positionalScore: 0.20,
        contentValidation: 0.15,
        lengthPenalty: 0.05
      };
    }

    const normalized = {};
    Object.keys(weights).forEach(key => {
      normalized[key] = weights[key] / total;
    });

    return normalized;

  } catch (error) {
    console.warn('Weight normalization error:', error.message);
    return weights;
  }
}

// 学習データ更新機能は現在未使用のためコメントアウト
// 将来的な機械学習機能拡張時に再利用予定

/**
 * 📏 ヘッダー長さペナルティ計算
 * @param {string} header - 正規化済みヘッダー
 * @returns {number} 長さスコア（0-100）
 */
function calculateLengthPenalty(header) {
  const {length} = header;

  // 理想的な長さ: 3-15文字
  if (length >= 3 && length <= 15) return 100;

  // 短すぎる（1-2文字）
  if (length < 3) return Math.max(0, 50 - ((3 - length) * 20));

  // 長すぎる（16文字以上）
  return Math.max(20, 100 - ((length - 15) * 5));
}

/**
 * 🎯 スマート位置ベースフォールバック（改良版）
 * @param {string} fieldType - フィールドタイプ
 * @param {Array} headers - ヘッダー配列
 * @returns {Object} { index: number, confidence: number }
 */
function getSmartPositionalFallback(fieldType, headers) {
  try {
    const typicalPositions = {
      timestamp: [0],
      answer: [1, 2, 3],
      reason: [2, 3, 4],
      class: [3, 4, 5],
      name: [4, 5, 6],
      email: [5, 6, 7]
    };

    const candidatePositions = typicalPositions[fieldType] || [];
    if (candidatePositions.length === 0) return { index: -1, confidence: 0 };

    // 有効な位置から最も適切なものを選択
    for (const pos of candidatePositions) {
      if (pos < headers.length && headers[pos]) {
        // 位置による信頼度調整
        let confidence;
        if (pos === candidatePositions[0]) {
          confidence = 40; // 第一候補
        } else if (pos === candidatePositions[1]) {
          confidence = 30; // 第二候補
        } else {
          confidence = 20; // その他
        }

        // ヘッダーの妥当性による微調整
        const headerQuality = evaluateHeaderQuality(headers[pos], fieldType);
        confidence = Math.min(50, confidence + headerQuality);

        return { index: pos, confidence };
      }
    }

    return { index: -1, confidence: 0 };

  } catch (error) {
    console.error('getSmartPositionalFallback エラー:', error.message);
    return { index: -1, confidence: 0 };
  }
}

/**
 * 📈 ヘッダー品質評価（位置フォールバック用）
 * @param {string} header - ヘッダー文字列
 * @param {string} fieldType - フィールドタイプ
 * @returns {number} 品質調整値（-10 to +10）
 */
function evaluateHeaderQuality(header, fieldType) {
  if (!header || typeof header !== 'string') return -10;

  let quality = 0;
  const normalizedHeader = header.toLowerCase().trim();

  // 長さチェック
  if (normalizedHeader.length >= 2 && normalizedHeader.length <= 20) quality += 3;

  // フィールドタイプ関連語句
  const relatedTerms = {
    timestamp: ['時', 'time', 'date'],
    email: ['mail', 'メール'],
    answer: ['答', 'answer'],
    reason: ['理由', 'reason'],
    class: ['class', 'クラス'],
    name: ['name', '名']
  };

  const terms = relatedTerms[fieldType] || [];
  for (const term of terms) {
    if (normalizedHeader.includes(term)) {
      quality += 5;
      break;
    }
  }

  // 意味のない文字列のペナルティ
  if (/^[0-9]+$|^[!@#$%^&*()]+$|^[\s]+$/.test(normalizedHeader)) {
    quality -= 8;
  }

  return Math.max(-10, Math.min(10, quality));
}

/**
 * 統一フィールド値抽出関数（エラーハンドリング強化版）
 * @param {Array} row - データ行
 * @param {Array} headers - ヘッダー配列
 * @param {string} fieldType - フィールドタイプ
 * @param {Object} columnMapping - 列マッピング
 * @param {Object} options - オプション設定
 * @returns {*} フィールド値
 */
function extractFieldValueUnified(row, headers, fieldType, columnMapping = {}, options = {}) {
  try {
    // 🛡️ 入力検証強化 - 完全なnull/undefined/型チェック
    if (!row || !Array.isArray(row) || row.length === 0) {
      if (options.enableDebug) {
        console.warn(`[WARN] ColumnMappingService.extractFieldValueUnified (${fieldType}): Invalid or empty row data`, {
          row: row ? 'defined' : 'null/undefined',
          isArray: Array.isArray(row),
          length: row ? row.length : 'N/A'
        });
      }
      return options.defaultValue || '';
    }

    // headers配列の検証も追加
    if (!headers || !Array.isArray(headers)) {
      if (options.enableDebug) {
        console.warn(`[WARN] ColumnMappingService.extractFieldValueUnified (${fieldType}): Invalid headers data`);
      }
      return handleColumnNotFound(fieldType, row, headers, options);
    }

    const columnResult = resolveColumnIndex(headers, fieldType, columnMapping, options);

    // 🔍 詳細なデバッグ情報
    if (options.enableDebug) {
      console.log(`[DEBUG] ColumnMappingService.extractFieldValueUnified (${fieldType}):`, {
        columnIndex: columnResult.index,
        confidence: columnResult.confidence,
        method: columnResult.method,
        rowLength: row.length,
        headersLength: headers ? headers.length : 0
      });
    }

    // 列が見つからない場合の処理
    if (columnResult.index === -1) {
      return handleColumnNotFound(fieldType, row, headers, options);
    }

    // 🛡️ 範囲外チェック強化 - 負の値や非整数も検証
    if (columnResult.index < 0 || columnResult.index >= row.length || !Number.isInteger(columnResult.index)) {
      if (options.enableDebug) {
        console.warn(`[WARN] ColumnMappingService.extractFieldValueUnified (${fieldType}): Invalid column index`, {
          index: columnResult.index,
          rowLength: row.length,
          isInteger: Number.isInteger(columnResult.index)
        });
      }
      return options.defaultValue || '';
    }

    // 🎯 値抽出実行
    const extractedValue = row[columnResult.index];

    // 🧹 値のクリーンアップ（V8最適化）
    if (extractedValue === null || extractedValue === undefined) {
      return options.defaultValue || '';
    }

    const cleanedValue = String(extractedValue).trim();
    return cleanedValue || (options.defaultValue || '');

  } catch (error) {
    const errorMessage = error && error.message ? error.message : 'Unknown extraction error';
    console.error(`[ERROR] ColumnMappingService.extractFieldValueUnified (${fieldType}):`, errorMessage);
    return options.defaultValue || '';
  }
}

/**
 * 列が見つからない場合の処理
 * @param {string} fieldType - フィールドタイプ
 * @param {Array} row - データ行
 * @param {Array} headers - ヘッダー配列
 * @param {Object} options - オプション設定
 * @returns {*} フォールバック値
 */
function handleColumnNotFound(fieldType, row, headers, options = {}) {
  if (options.enableDebug) {
    console.warn(`[WARN] ColumnMappingService.handleColumnNotFound: Column not found for ${fieldType}`, {
      availableHeaders: headers || [],
      rowLength: row.length
    });
  }

  // フィールドタイプ別のフォールバック戦略
  switch (fieldType) {
    case 'timestamp':
      // タイムスタンプは通常最初の列
      return row[0] || options.defaultValue || '';

    case 'answer':
    case 'opinion':
      // 回答フィールドが見つからない場合は物理位置フォールバック
      return getPositionalFallback('answer', row) || options.defaultValue || '';

    case 'email': {
      // メールアドレスパターンを検索
      const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
      const emailCandidate = row.find(cell =>
        cell && typeof cell === 'string' && emailPattern.test(cell)
      );
      return emailCandidate || options.defaultValue || '';
    }

    default:
      return options.defaultValue || '';
  }
}



// ===========================================
// 🔍 列診断・レポート生成システム
// ===========================================

/**
 * 列診断レポート生成（統合版）
 * @param {Array} headers - ヘッダー配列
 * @param {Object} columnMapping - 列マッピング
 * @param {Array} requiredFields - 必須フィールド配列
 * @returns {Object} 診断レポート
 */
function generateColumnDiagnosticReport(headers, columnMapping = {}, requiredFields = ['answer', 'reason', 'class', 'name']) {
  const report = {
    timestamp: new Date().toISOString(),
    headers: headers || [],
    requiredFields,
    optionalFields: ['timestamp', 'email'],
    fieldAnalysis: {},
    issues: [],
    recommendations: [],
    summary: {
      totalFields: 0,
      resolvedFields: 0,
      unresolvedFields: 0,
      confidence: {
        high: 0,
        medium: 0,
        low: 0
      },
      overallScore: 0
    }
  };

  try {
    // 全フィールドを分析対象に
    const allFields = [...requiredFields, ...report.optionalFields];
    report.summary.totalFields = allFields.length;

    // フィールド別分析
    allFields.forEach(fieldType => {
      const columnResult = resolveColumnIndex(headers, fieldType, columnMapping);

      const analysis = {
        fieldType,
        resolved: columnResult.index !== -1,
        columnIndex: columnResult.index,
        confidence: columnResult.confidence,
        method: columnResult.method,
        header: columnResult.index !== -1 ? headers[columnResult.index] : null,
        required: requiredFields.includes(fieldType),
        issues: [],
        recommendations: []
      };

      // 解決状況の評価
      if (analysis.resolved) {
        report.summary.resolvedFields++;

        // 信頼度評価
        if (analysis.confidence >= 80) {
          report.summary.confidence.high++;
        } else if (analysis.confidence >= 50) {
          report.summary.confidence.medium++;
        } else {
          report.summary.confidence.low++;
          analysis.issues.push(`信頼度が低い (${analysis.confidence}%)`);
          analysis.recommendations.push('より明確なヘッダー名の使用を推奨');
        }
      } else {
        report.summary.unresolvedFields++;

        if (analysis.required) {
          analysis.issues.push('必須フィールドが見つかりません');
          analysis.recommendations.push('列マッピングを手動設定してください');
          report.issues.push(`必須フィールド '${fieldType}' が見つかりません`);
        } else {
          analysis.issues.push('オプションフィールドが見つかりません');
          analysis.recommendations.push('必要に応じて列マッピングを設定してください');
        }
      }

      report.fieldAnalysis[fieldType] = analysis;
    });

    // 全体スコア計算
    const resolvedWeight = (report.summary.resolvedFields / report.summary.totalFields) * 100;
    const confidenceWeight =
      (report.summary.confidence.high * 1.0 +
       report.summary.confidence.medium * 0.7 +
       report.summary.confidence.low * 0.4) / report.summary.totalFields * 100;

    report.summary.overallScore = Math.round((resolvedWeight + confidenceWeight) / 2);

    // 全体推奨事項
    report.recommendations = generateColumnRecommendations(report);

    return report;

  } catch (error) {
    console.error('generateColumnDiagnosticReport エラー:', error.message);
    report.issues.push(`診断中にエラーが発生しました: ${error.message}`);
    return report;
  }
}

/**
 * 列推奨事項生成
 * @param {Object} report - 診断レポート
 * @returns {Array} 推奨事項配列
 */
function generateColumnRecommendations(report) {
  const recommendations = [];

  try {
    // 未解決必須フィールドの推奨事項
    const unresolvedRequired = Object.values(report.fieldAnalysis)
      .filter(field => field.required && !field.resolved);

    if (unresolvedRequired.length > 0) {
      recommendations.push({
        priority: 'critical',
        type: 'missing_required_fields',
        message: `必須フィールド (${unresolvedRequired.map(f => f.fieldType).join(', ')}) が見つかりません。列マッピングを確認してください。`,
        fields: unresolvedRequired.map(f => f.fieldType)
      });
    }

    // 低信頼度フィールドの推奨事項
    const lowConfidenceFields = Object.values(report.fieldAnalysis)
      .filter(field => field.resolved && field.confidence < 70);

    if (lowConfidenceFields.length > 0) {
      recommendations.push({
        priority: 'medium',
        type: 'low_confidence_fields',
        message: `一部フィールドの判定信頼度が低いです。より明確なヘッダー名を使用することを推奨します。`,
        fields: lowConfidenceFields.map(f => ({ fieldType: f.fieldType, confidence: f.confidence }))
      });
    }

    // 全体スコアベースの推奨事項
    if (report.summary.overallScore < 60) {
      recommendations.push({
        priority: 'high',
        type: 'overall_improvement',
        message: `全体的な列判定精度が低いです (${report.summary.overallScore}%)。スプレッドシートのヘッダー構造を見直すことを推奨します。`
      });
    } else if (report.summary.overallScore < 80) {
      recommendations.push({
        priority: 'medium',
        type: 'partial_improvement',
        message: `列判定精度に改善の余地があります (${report.summary.overallScore}%)。一部のヘッダー名をより具体的にすることを推奨します。`
      });
    }

    return recommendations;

  } catch (error) {
    console.error('generateColumnRecommendations エラー:', error);
    return [{
      priority: 'low',
      type: 'error',
      message: '推奨事項の生成中にエラーが発生しました'
    }];
  }
}


// ===========================================
// 🔬 統合診断システム（高度）
// ===========================================

/**
 * 🧠 AI強化列分析：ヘッダーから推奨マッピングと信頼度を自動生成
 * @param {Array} headers - ヘッダー配列
 * @param {Object} options - 分析オプション（sampleData含む）
 * @returns {Object} { recommendedMapping, confidence, analysis }
 */
function generateRecommendedMapping(headers, options = {}) {
  const analysis = {
    timestamp: new Date().toISOString(),
    headers: headers || [],
    fieldResults: {},
    conflictResolution: {},
    qualityMetrics: {},
    overallScore: 0,
    aiEnhancementUsed: true
  };

  try {
    const targetFields = options.fields || ['answer', 'reason', 'class', 'name', 'timestamp', 'email'];
    const sampleData = options.sampleData || [];
    const recommendedMapping = {};
    const confidence = {};
    const usedIndices = new Set(); // 重複チェック用

    // 🧠 AI強化分析オプション設定
    const aiOptions = {
      allowPositionalFallback: true,
      sampleData: sampleData.slice(0, 5), // 最大5行のサンプル
      enableContentValidation: sampleData.length > 0
    };

    let totalConfidence = 0;
    let resolvedFields = 0;
    const fieldAnalysisResults = [];

    // 🎯 段階1: 各フィールドのAI分析実行（重複チェックなし）
    targetFields.forEach(fieldType => {
      const result = resolveColumnIndex(headers, fieldType, {}, aiOptions);

      const fieldAnalysis = {
        fieldType,
        resolved: result.index !== -1,
        index: result.index,
        confidence: result.confidence,
        method: result.method,
        header: result.index !== -1 ? headers[result.index] : null,
        scoringDetails: result.debug?.scoringDetails || null
      };

      analysis.fieldResults[fieldType] = fieldAnalysis;
      fieldAnalysisResults.push({ fieldType, result, analysis: fieldAnalysis });
    });

    // 🎯 段階2: 重複解決アルゴリズム（信頼度ベース優先順位）
    fieldAnalysisResults
      .filter(item => item.result.index !== -1)
      .sort((a, b) => b.result.confidence - a.result.confidence) // 信頼度降順
      .forEach(item => {
        const { fieldType, result } = item;

        if (!usedIndices.has(result.index)) {
          // 重複なし: 採用
          recommendedMapping[fieldType] = result.index;
          confidence[fieldType] = result.confidence;
          usedIndices.add(result.index);
          totalConfidence += result.confidence;
          resolvedFields++;

          analysis.fieldResults[fieldType].adopted = true;
        } else {
          // 🔧 重複解決: 代替候補を探索
          const alternativeResult = findAlternativeColumn(headers, fieldType, usedIndices, aiOptions);

          if (alternativeResult.index !== -1) {
            recommendedMapping[fieldType] = alternativeResult.index;
            confidence[fieldType] = alternativeResult.confidence;
            usedIndices.add(alternativeResult.index);
            totalConfidence += alternativeResult.confidence;
            resolvedFields++;

            analysis.fieldResults[fieldType].adopted = true;
            analysis.fieldResults[fieldType].alternativeUsed = true;
            analysis.conflictResolution[fieldType] = {
              originalIndex: result.index,
              alternativeIndex: alternativeResult.index,
              reason: 'index_conflict_resolved'
            };
          } else {
            analysis.fieldResults[fieldType].adopted = false;
            analysis.conflictResolution[fieldType] = {
              originalIndex: result.index,
              reason: 'no_alternative_found'
            };
          }
        }
      });

    // 🎯 段階3: 品質メトリクス計算
    analysis.qualityMetrics = calculateMappingQuality(recommendedMapping, headers, sampleData);

    // 全体スコア計算（品質メトリクスを考慮）
    const baseScore = resolvedFields > 0 ? Math.round(totalConfidence / resolvedFields) : 0;
    const qualityBonus = Math.round(analysis.qualityMetrics.overallQuality * 0.1); // 10%まで品質ボーナス
    analysis.overallScore = Math.min(99, baseScore + qualityBonus);

    // 🎯 段階4: 論理整合性チェック
    const consistencyCheck = validateLogicalConsistency(recommendedMapping, headers);
    analysis.consistencyCheck = consistencyCheck;

    if (!consistencyCheck.isConsistent) {
      analysis.overallScore = Math.max(50, analysis.overallScore - 15); // 整合性ペナルティ
    }

    console.log('✅ AI強化列分析完了:', {
      resolvedFields: `${resolvedFields}/${targetFields.length}`,
      overallScore: analysis.overallScore,
      qualityScore: analysis.qualityMetrics.overallQuality,
      consistencyCheck: consistencyCheck.isConsistent,
      conflicts: Object.keys(analysis.conflictResolution).length
    });

    return {
      recommendedMapping,
      confidence,
      analysis,
      success: true
    };

  } catch (error) {
    console.error('generateRecommendedMapping エラー:', error.message);
    return {
      recommendedMapping: {},
      confidence: {},
      analysis: { ...analysis, error: error.message },
      success: false
    };
  }
}

/**
 * 🔍 代替列候補探索（重複解決用）
 * @param {Array} headers - ヘッダー配列
 * @param {string} fieldType - フィールドタイプ
 * @param {Set} usedIndices - 使用済みインデックス
 * @param {Object} options - オプション設定
 * @returns {Object} { index: number, confidence: number }
 */
function findAlternativeColumn(headers, fieldType, usedIndices, options = {}) {
  try {
    const headerPatterns = getHeaderPatterns();
    const patterns = headerPatterns[fieldType] || [];
    let bestAlternative = { index: -1, confidence: 0 };

    // 各ヘッダーを代替候補として評価（使用済み除外）
    headers.forEach((header, index) => {
      if (usedIndices.has(index) || !header || typeof header !== 'string') return;

      const candidate = evaluateHeaderCandidate(header, index, fieldType, patterns, options);

      // より良い代替候補があれば更新
      if (candidate.totalScore > bestAlternative.confidence) {
        bestAlternative = {
          index: candidate.index,
          confidence: Math.min(candidate.totalScore - 5, 90), // 代替候補は少しペナルティ
          method: 'alternative_search'
        };
      }
    });

    return bestAlternative;

  } catch (error) {
    console.error('findAlternativeColumn エラー:', error.message);
    return { index: -1, confidence: 0 };
  }
}

/**
 * 📊 マッピング品質評価
 * @param {Object} mapping - 推奨マッピング
 * @param {Array} headers - ヘッダー配列
 * @param {Array} sampleData - サンプルデータ
 * @returns {Object} 品質メトリクス
 */
function calculateMappingQuality(mapping, headers, sampleData = []) {
  const metrics = {
    mappingCoverage: 0,
    headerQuality: 0,
    contentConsistency: 0,
    duplicateCheck: 0,
    overallQuality: 0
  };

  try {
    const requiredFields = ['answer', 'reason'];
    const mappedFields = Object.keys(mapping);

    // 1. マッピングカバレッジ（必須フィールドの解決率）
    const resolvedRequired = requiredFields.filter(field => mappedFields.includes(field));
    metrics.mappingCoverage = (resolvedRequired.length / requiredFields.length) * 100;

    // 2. ヘッダー品質（平均的なヘッダーの明確さ）
    const headerQualities = mappedFields.map(field => {
      const index = mapping[field];
      return evaluateHeaderQuality(headers[index], field) + 50; // -10~+10 を 40~60 に変換
    });
    metrics.headerQuality = headerQualities.length > 0 ?
      headerQualities.reduce((sum, q) => sum + q, 0) / headerQualities.length : 50;

    // 3. コンテンツ整合性（サンプルデータがある場合）
    if (sampleData.length > 0) {
      const contentScores = mappedFields.map(field => {
        const index = mapping[field];
        return validateContentType(sampleData, index, field);
      });
      metrics.contentConsistency = contentScores.length > 0 ?
        contentScores.reduce((sum, score) => sum + score, 0) / contentScores.length : 50;
    } else {
      metrics.contentConsistency = 50; // 中立値
    }

    // 4. 重複チェック（すべて異なるインデックスか）
    const indices = Object.values(mapping);
    const uniqueIndices = new Set(indices);
    metrics.duplicateCheck = (uniqueIndices.size === indices.length) ? 100 : 0;

    // 総合品質スコア
    metrics.overallQuality = Math.round(
      (metrics.mappingCoverage * 0.35) +
      (metrics.headerQuality * 0.25) +
      (metrics.contentConsistency * 0.25) +
      (metrics.duplicateCheck * 0.15)
    );

    return metrics;

  } catch (error) {
    console.error('calculateMappingQuality エラー:', error.message);
    return { ...metrics, error: error.message };
  }
}

/**
 * 🧩 論理整合性検証
 * @param {Object} mapping - 推奨マッピング
 * @param {Array} headers - ヘッダー配列
 * @returns {Object} 整合性チェック結果
 */
function validateLogicalConsistency(mapping, headers) {
  const result = {
    isConsistent: true,
    issues: [],
    warnings: []
  };

  try {
    // 1. 重複インデックスチェック
    const indices = Object.values(mapping);
    const uniqueIndices = new Set(indices);
    if (uniqueIndices.size !== indices.length) {
      result.isConsistent = false;
      result.issues.push('重複するインデックスが検出されました');
    }

    // 2. 範囲外インデックスチェック
    indices.forEach(index => {
      if (index < 0 || index >= headers.length) {
        result.isConsistent = false;
        result.issues.push(`範囲外インデックス: ${index}`);
      }
    });

    // 3. 必須フィールドの存在チェック
    const requiredFields = ['answer'];
    const missingRequired = requiredFields.filter(field => !mapping[field]);
    if (missingRequired.length > 0) {
      result.warnings.push(`必須フィールド未解決: ${missingRequired.join(', ')}`);
    }

    // 4. 論理的順序チェック（timestamp < answer など）
    if (mapping.timestamp !== undefined && mapping.answer !== undefined) {
      if (mapping.timestamp > mapping.answer) {
        result.warnings.push('タイムスタンプが回答フィールドより後にあります');
      }
    }

    return result;

  } catch (error) {
    console.error('validateLogicalConsistency エラー:', error.message);
    return {
      isConsistent: false,
      issues: [`整合性チェック中にエラー: ${error.message}`],
      warnings: []
    };
  }
}

/**
 * 統合列診断（サンプルデータ付き高精度版）
 * @param {Array} headers - ヘッダー配列
 * @param {Object} columnMapping - 列マッピング
 * @param {Array} sampleData - サンプルデータ（オプション）
 * @returns {Object} 統合診断結果
 */
function performIntegratedColumnDiagnostics(headers, columnMapping = {}, sampleData = []) {
  const startTime = Date.now();
  const diagnostics = {
    timestamp: new Date().toISOString(),
    executionTime: 0,

    // 基本診断
    basicReport: generateColumnDiagnosticReport(headers, columnMapping),

    // システム健全性
    systemHealth: {
      backend: diagnoseBackendColumnSystem(headers),
      frontend: diagnoseFrontendColumnSystem(columnMapping),
      integration: null // 後で設定
    },

    // 統合テスト結果
    integrationTests: [],

    // 総合評価
    summary: {
      overallScore: 0,
      criticalIssues: 0,
      warnings: 0,
      recommendations: []
    }
  };

  try {
    // ✅ AI列分析実行：推奨マッピングと信頼度を自動生成
    const aiAnalysis = generateRecommendedMapping(headers);
    diagnostics.recommendedMapping = aiAnalysis.recommendedMapping;
    diagnostics.confidence = aiAnalysis.confidence;
    diagnostics.aiAnalysis = aiAnalysis.analysis;

    // 統合テスト実行
    if (sampleData.length > 0) {
      diagnostics.integrationTests = performIntegrationTests(headers, columnMapping, sampleData);
    }

    // システム健全性統合評価
    diagnostics.systemHealth.integration = {
      backendFrontendSync: Math.abs(diagnostics.systemHealth.backend.score - diagnostics.systemHealth.frontend.score) <= 20,
      overallSystemScore: (diagnostics.systemHealth.backend.score + diagnostics.systemHealth.frontend.score) / 2
    };

    // 総合スコア計算
    const scores = [
      diagnostics.basicReport.summary.overallScore,
      diagnostics.systemHealth.backend.score,
      diagnostics.systemHealth.frontend.score
    ];
    diagnostics.summary.overallScore = Math.round(scores.reduce((sum, score) => sum + score, 0) / scores.length);

    // 重大問題・警告カウント
    diagnostics.summary.criticalIssues = diagnostics.basicReport.issues.length;
    diagnostics.summary.warnings = Object.values(diagnostics.basicReport.fieldAnalysis)
      .filter(field => field.issues.length > 0).length;

    // 統合推奨事項
    diagnostics.summary.recommendations = generateSystemRecommendations(diagnostics);

    diagnostics.executionTime = Date.now() - startTime;
    return diagnostics;

  } catch (error) {
    console.error('performIntegratedColumnDiagnostics エラー:', error);
    diagnostics.error = error.message;
    diagnostics.executionTime = Date.now() - startTime;
    return diagnostics;
  }
}

/**
 * バックエンド列システム診断（簡素化版）
 * @param {Array} headers - ヘッダー配列
 * @returns {Object} バックエンド診断結果
 */
function diagnoseBackendColumnSystem(headers) {
  const diagnosis = {
    system: 'backend',
    score: 0,
    issues: [],
    strengths: [],
    details: { totalHeaders: headers.length }
  };

  try {
    // 必須フィールドの存在チェック
    const essentialFields = ['answer', 'timestamp'];
    let foundEssential = 0;

    essentialFields.forEach(fieldType => {
      const result = resolveColumnIndex(headers, fieldType);
      if (result.index !== -1) foundEssential++;
    });

    // スコア計算（シンプル化）
    diagnosis.score = Math.round((foundEssential / essentialFields.length) * 100);

    // 問題・強みの判定
    if (foundEssential === essentialFields.length) {
      diagnosis.strengths.push('必須フィールドがすべて認識されている');
    } else {
      diagnosis.issues.push(`必須フィールドが不足 (${foundEssential}/${essentialFields.length})`);
    }

    diagnosis.details.foundEssential = foundEssential;
    return diagnosis;

  } catch (error) {
    diagnosis.score = 0;
    diagnosis.issues.push(`バックエンド診断エラー: ${error.message}`);
    return diagnosis;
  }
}

/**
 * フロントエンド列システム診断（簡素化版）
 * @param {Object} columnMapping - 列マッピング
 * @returns {Object} フロントエンド診断結果
 */
function diagnoseFrontendColumnSystem(columnMapping) {
  const diagnosis = {
    system: 'frontend',
    score: 0,
    issues: [],
    strengths: [],
    details: {}
  };

  try {
    const hasMappings = columnMapping && Object.keys(columnMapping).length > 0;
    const mappingCount = hasMappings ? Object.keys(columnMapping).length : 0;

    if (hasMappings) {
      // 有効マッピング数をカウント
      const validMappings = Object.values(columnMapping)
        .filter(value => typeof value === 'number' && value >= 0).length;

      // スコア計算（シンプル化）
      diagnosis.score = Math.round((validMappings / mappingCount) * 100);

      // 問題・強みの判定
      if (validMappings === mappingCount) {
        diagnosis.strengths.push('すべての列マッピングが有効');
      } else {
        diagnosis.issues.push(`無効な列マッピングがあります`);
      }

      diagnosis.details = { mappingCount, validMappings };
    } else {
      diagnosis.score = 0;
      diagnosis.issues.push('列マッピングが設定されていません');
      diagnosis.details = { mappingCount: 0, validMappings: 0 };
    }

    return diagnosis;

  } catch (error) {
    diagnosis.score = 0;
    diagnosis.issues.push(`フロントエンド診断エラー: ${error.message}`);
    return diagnosis;
  }
}

/**
 * 統合テスト実行
 * @param {Array} headers - ヘッダー配列
 * @param {Object} columnMapping - 列マッピング
 * @param {Array} sampleData - サンプルデータ
 * @returns {Array} テスト結果配列
 */
function performIntegrationTests(headers, columnMapping, sampleData) {
  const tests = [];

  try {
    const fieldsToTest = ['answer', 'reason', 'class', 'name', 'timestamp', 'email'];

    fieldsToTest.forEach(fieldType => {
      const test = testFieldResolution(headers, fieldType, columnMapping, sampleData);
      tests.push(test);
    });

    return tests;

  } catch (error) {
    console.error('performIntegrationTests エラー:', error);
    return [{
      fieldType: 'error',
      resolved: false,
      severity: 'critical',
      issue: `統合テスト実行エラー: ${error.message}`
    }];
  }
}

/**
 * フィールド解決テスト
 * @param {Array} headers - ヘッダー配列
 * @param {string} fieldType - フィールドタイプ
 * @param {Object} columnMapping - 列マッピング
 * @param {Array} sampleData - サンプルデータ
 * @returns {Object} テスト結果
 */
function testFieldResolution(headers, fieldType, columnMapping, sampleData) {
  try {
    const result = resolveColumnIndex(headers, fieldType, columnMapping);

    const test = {
      fieldType,
      resolved: result.index !== -1,
      confidence: result.confidence,
      method: result.method,
      severity: 'info'
    };

    if (!test.resolved) {
      test.severity = fieldType === 'answer' ? 'critical' : 'warning';
      test.issue = `${fieldType}フィールドが解決できません`;
    } else if (test.confidence < 50) {
      test.severity = 'warning';
      test.issue = `${fieldType}フィールドの信頼度が低いです (${test.confidence}%)`;
    }

    // 実データでの抽出テスト
    if (test.resolved && sampleData.length > 0) {
      const extractedValue = extractFieldValueUnified(sampleData[0], headers, fieldType, columnMapping);
      test.extractionSuccess = extractedValue !== '';

      if (!test.extractionSuccess) {
        test.severity = 'warning';
        test.issue = `${fieldType}フィールドからデータを抽出できません`;
      }
    }

    return test;

  } catch (error) {
    return {
      fieldType,
      resolved: false,
      severity: 'critical',
      issue: `${fieldType}フィールドテスト中にエラー: ${error.message}`
    };
  }
}

/**
 * システム推奨事項生成
 * @param {Object} diagnostics - 診断結果
 * @returns {Array} 推奨事項リスト
 */
function generateSystemRecommendations(diagnostics) {
  const recommendations = [];

  try {
    // 重大な問題の推奨事項
    if (diagnostics.summary.criticalIssues > 0) {
      recommendations.push({
        priority: 'critical',
        type: 'immediate_action',
        message: `${diagnostics.summary.criticalIssues}件の重大な問題があります。列設定を確認してください。`
      });
    }

    // システム別推奨事項
    if (diagnostics.systemHealth.backend.score < 70) {
      recommendations.push({
        priority: 'high',
        type: 'backend_improvement',
        message: 'バックエンドの列判定システムに問題があります。ヘッダーパターンや列マッピングを確認してください。'
      });
    }

    if (diagnostics.systemHealth.frontend.score < 70) {
      recommendations.push({
        priority: 'high',
        type: 'frontend_improvement',
        message: 'フロントエンドの列設定に問題があります。管理パネルで設定を確認してください。'
      });
    }

    if (diagnostics.summary.overallScore < 80) {
      recommendations.push({
        priority: 'medium',
        type: 'general_improvement',
        message: 'システム全体の列判定精度向上のため、より明確なヘッダー名の使用を推奨します。'
      });
    }

    return recommendations;

  } catch (error) {
    console.error('generateSystemRecommendations エラー:', error);
    return [{
      priority: 'low',
      type: 'error',
      message: '推奨事項の生成中にエラーが発生しました'
    }];
  }
}