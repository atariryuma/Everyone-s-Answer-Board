/**
 * @fileoverview ColumnMappingService - High-Precision AI Column Detection (2025)
 *
 * 🎯 Design Principles:
 * - Direct pattern matching with statistical validation
 * - Efficient single-pass detection algorithm
 * - Optimized for educational forms and surveys
 * - Zero-dependency GAS-native implementation
 */

/**
 * Main column detection engine - optimized for precision and efficiency
 * @param {Array} headers - Header array
 * @param {string} fieldType - Field type to detect
 * @param {Object} columnMapping - Existing mapping (takes priority)
 * @param {Object} options - Detection options
 * @returns {Object} { index: number, confidence: number, method: string }
 */
function resolveColumnIndex(headers, fieldType, columnMapping = {}, options = {}) {
  try {
    // 1. Existing mapping takes priority
    if (columnMapping && columnMapping[fieldType] !== undefined) {
      const mappedIndex = columnMapping[fieldType];
      if (typeof mappedIndex === 'number' && mappedIndex >= 0 && mappedIndex < headers.length) {
        return { index: mappedIndex, confidence: 100, method: 'existing_mapping' };
      }
    }

    // 2. Filter system columns
    const { cleanHeaders, indexMap } = filterSystemColumns(headers);
    if (cleanHeaders.length === 0) {
      return { index: -1, confidence: 0, method: 'no_valid_headers' };
    }

    // 3. Run optimized detection
    const detection = detectColumn(cleanHeaders, fieldType, options);
    const originalIndex = detection.index !== -1 ? indexMap[detection.index] : -1;

    return {
      index: originalIndex,
      confidence: detection.confidence,
      method: detection.method
    };

  } catch (error) {
    console.error(`resolveColumnIndex error for ${fieldType}:`, error.message);
    return { index: -1, confidence: 0, method: 'error', error: error.message };
  }
}

/**
 * Filter out system columns that should not be detected
 */
function filterSystemColumns(headers) {
  const systemPatterns = [
    /^タイムスタンプ$/i, /^timestamp$/i, /^日時$/i, /^日付$/i,
    /^UNDERSTAND$/i, /^LIKE$/i, /^CURIOUS$/i, /^HIGHLIGHT$/i,
    /^理解$/i, /^いいね$/i, /^気になる$/i, /^ハイライト$/i,
    /^_/  // Internal columns
  ];

  const cleanHeaders = [];
  const indexMap = [];

  headers.forEach((header, originalIndex) => {
    if (header && typeof header === 'string' && header.trim()) {
      const isSystem = systemPatterns.some(pattern => pattern.test(header.trim()));
      if (!isSystem) {
        cleanHeaders.push(header);
        indexMap.push(originalIndex);
      }
    }
  });

  return { cleanHeaders, indexMap };
}

/**
 * Enhanced column detection with logical field ordering
 */
function detectColumn(headers, fieldType, options = {}) {
  const patterns = getFieldPatterns()[fieldType];
  if (!patterns) {
    return { index: -1, confidence: 0, method: 'unknown_field' };
  }

  // Analyze field relationships for logical ordering
  const fieldRelationships = analyzeFieldRelationships(headers);

  let bestMatch = { index: -1, confidence: 0, method: 'no_match' };

  headers.forEach((header, index) => {
    if (!header || typeof header !== 'string') return;

    const score = calculateScoreWithLogic(header, fieldType, patterns, index, fieldRelationships, options);

    if (score > bestMatch.confidence) {
      bestMatch = {
        index,
        confidence: Math.min(score, 95),
        method: 'pattern_match_with_logic'
      };
    }
  });

  return bestMatch;
}

/**
 * Analyze field relationships and logical constraints across all headers
 */
function analyzeFieldRelationships(headers) {
  const relationships = {
    answerCandidates: [],
    reasonCandidates: [],
    answerReasonPairs: []
  };

  const patterns = getFieldPatterns();

  // Identify potential answer and reason columns
  headers.forEach((header, index) => {
    if (!header || typeof header !== 'string') return;

    const normalizedHeader = header.toLowerCase().trim();

    // Check for answer patterns
    const answerScore = calculateBaseScore(header, 'answer', patterns.answer);
    if (answerScore > 60) {
      relationships.answerCandidates.push({ index, header, score: answerScore });
    }

    // Check for reason patterns
    const reasonScore = calculateBaseScore(header, 'reason', patterns.reason);
    if (reasonScore > 60) {
      relationships.reasonCandidates.push({ index, header, score: reasonScore });
    }
  });

  // Identify logical answer → reason pairs
  relationships.answerCandidates.forEach(answer => {
    relationships.reasonCandidates.forEach(reason => {
      if (answer.index < reason.index) { // Answer comes before reason (logical)
        relationships.answerReasonPairs.push({
          answerIndex: answer.index,
          reasonIndex: reason.index,
          logicalOrder: true,
          confidence: (answer.score + reason.score) / 2
        });
      }
    });
  });

  return relationships;
}

/**
 * Calculate base score without positional logic (for relationship analysis)
 */
function calculateBaseScore(header, fieldType, patterns) {
  const normalizedHeader = header.toLowerCase().trim();
  let maxScore = 0;

  // Direct pattern matching
  if (patterns.exact && patterns.exact.some(keyword => normalizedHeader === keyword.toLowerCase())) {
    maxScore = 95;
  }

  // Keyword matching
  if (patterns.keywords) {
    for (const keyword of patterns.keywords) {
      if (normalizedHeader.includes(keyword.toLowerCase())) {
        maxScore = Math.max(maxScore, 90);
        break;
      }
    }
  }

  // Question pattern matching (for answer fields)
  if (fieldType === 'answer' && patterns.questionPatterns) {
    for (const pattern of patterns.questionPatterns) {
      if (pattern.test(header)) {
        maxScore = Math.max(maxScore, 87);
        break;
      }
    }
  }

  // Regex patterns
  if (patterns.regex && patterns.regex.test(header)) {
    maxScore = Math.max(maxScore, 85);
  }

  return maxScore;
}

/**
 * Calculate header score with logical field ordering constraints
 */
function calculateScoreWithLogic(header, fieldType, patterns, index, relationships, options = {}) {
  const normalizedHeader = header.toLowerCase().trim();
  const baseScore = calculateBaseScore(header, fieldType, patterns);

  // Apply logical ordering constraints
  let logicalBonus = 0;
  let logicalPenalty = 0;

  if (fieldType === 'answer') {
    // Boost answer fields that appear before reason candidates
    const hasReasonAfter = relationships.reasonCandidates.some(reason => reason.index > index);
    if (hasReasonAfter && baseScore > 70) {
      logicalBonus = 8; // Significant boost for logical answer placement
    }

    // Check if this answer is part of a logical pair
    const isInLogicalPair = relationships.answerReasonPairs.some(pair => pair.answerIndex === index);
    if (isInLogicalPair) {
      logicalBonus += 5; // Additional boost for confirmed logical pairs
    }
  }

  if (fieldType === 'reason') {
    // Penalize reason fields that appear before answer candidates
    const hasAnswerBefore = relationships.answerCandidates.some(answer => answer.index < index);
    if (!hasAnswerBefore && baseScore > 70) {
      logicalPenalty = 15; // Strong penalty for illogical reason placement
    }

    // Boost reason fields that appear after answer candidates
    if (hasAnswerBefore) {
      logicalBonus = 5; // Moderate boost for logical reason placement
    }

    // Check if this reason is part of a logical pair
    const isInLogicalPair = relationships.answerReasonPairs.some(pair => pair.reasonIndex === index);
    if (isInLogicalPair) {
      logicalBonus += 3; // Additional boost for confirmed logical pairs
    }
  }

  // Apply sample data validation bonus
  let sampleBonus = 0;
  if (options.sampleData && options.sampleData.length > 0 && baseScore > 0) {
    sampleBonus = validateSampleData(options.sampleData, normalizedHeader, fieldType);
  }

  const finalScore = Math.min(baseScore + logicalBonus + sampleBonus - logicalPenalty, 95);
  return Math.max(finalScore, 0); // Ensure score doesn't go negative
}

/**
 * Validate sample data for field type consistency
 */
function validateSampleData(sampleData, header, fieldType) {
  if (!sampleData || sampleData.length < 2) return 0;

  // Extract column values (assume header is at a specific index)
  const columnValues = sampleData.slice(0, Math.min(5, sampleData.length))
    .map(row => row.find(cell => String(cell || '').toLowerCase().includes(header.split(' ')[0])))
    .filter(val => val != null && String(val).trim());

  if (columnValues.length === 0) return 0;

  const validators = {
    email: val => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(val)),
    answer: val => String(val).length > 5 && String(val).length < 1000,
    reason: val => String(val).length > 3 && String(val).length < 500,
    name: val => String(val).length > 0 && String(val).length < 100 && !/[@.]/.test(String(val)),
    class: val => /^[0-9一-九\w-]+$/.test(String(val)) && String(val).length < 20
  };

  const validator = validators[fieldType];
  if (!validator) return 0;

  const validCount = columnValues.filter(validator).length;
  const validationRatio = validCount / columnValues.length;

  return validationRatio > 0.7 ? 5 : 0;
}

/**
 * Streamlined field patterns focused on high precision detection
 */
function getFieldPatterns() {
  return {
    answer: {
      exact: ['回答', 'answer'],
      keywords: ['回答', '答え', 'answer', '意見', '考え', '感想', 'コメント'],
      questionPatterns: [
        /どう.*思い?.*ますか/i,
        /.*と思い?.*ますか/i,
        /.*書きましょう/i,
        /.*述べ/i,
        /.*説明して/i,
        /.*気づいたこと/i,
        /.*観察して/i,
        /what.*do you think/i,
        /how.*do you feel/i,
        /explain.*your/i
      ],
      regex: /(回答|答え|answer|意見|考え)/i
    },
    reason: {
      exact: ['理由', 'reason'],
      keywords: ['理由', 'reason', 'なぜ', 'どうして', '根拠', '原因'],
      regex: /(理由|根拠|reason|なぜ|どうして)/i
    },
    name: {
      exact: ['名前', 'name'],
      keywords: ['名前', 'name', '氏名', 'お名前'],
      regex: /(名前|氏名|name)/i
    },
    class: {
      exact: ['クラス', 'class'],
      keywords: ['クラス', 'class', '組', '学年'],
      regex: /(クラス|class|組)/i
    },
    email: {
      exact: ['メール', 'email'],
      keywords: ['メール', 'email', 'mail', 'アドレス'],
      regex: /(メール|email|mail|アドレス)/i
    }
  };
}

/**
 * Validate and correct field ordering logic in final mapping
 */
function validateFieldOrderLogic(mapping, confidence, headers) {
  const validatedMapping = { ...mapping };
  const validatedConfidence = { ...confidence };
  const validation = {
    checks: [],
    corrections: [],
    warnings: []
  };

  // Check answer → reason ordering constraint
  if (validatedMapping.answer !== undefined && validatedMapping.reason !== undefined) {
    const answerIndex = validatedMapping.answer;
    const reasonIndex = validatedMapping.reason;

    validation.checks.push({
      rule: 'answer_before_reason',
      answerIndex,
      reasonIndex,
      logical: answerIndex < reasonIndex
    });

    // If reason comes before answer (illogical), apply correction
    if (reasonIndex < answerIndex) {
      const answerHeader = headers[answerIndex];
      const reasonHeader = headers[reasonIndex];

      // Check if we should swap or remove the weaker field
      if (validatedConfidence.reason < validatedConfidence.answer - 10) {
        // Remove reason field if it's significantly weaker
        delete validatedMapping.reason;
        delete validatedConfidence.reason;
        validation.corrections.push({
          action: 'removed_illogical_reason',
          field: 'reason',
          index: reasonIndex,
          header: reasonHeader,
          reason: 'Reason field appears before answer field with low confidence'
        });
      } else if (validatedConfidence.answer < validatedConfidence.reason - 10) {
        // Consider if the "answer" might actually be a reason
        const answerAsReasonScore = calculateBaseScore(answerHeader, 'reason', getFieldPatterns().reason);
        if (answerAsReasonScore > validatedConfidence.answer - 20) {
          validation.warnings.push({
            warning: 'potential_field_swap',
            message: 'Answer field may actually be a reason field',
            answerIndex,
            reasonIndex,
            suggestion: 'Manual review recommended'
          });
        }
      } else {
        // Similar confidence levels - prefer logical ordering
        validation.corrections.push({
          action: 'preserved_logical_order',
          message: 'Maintained answer before reason despite close confidence scores',
          answerIndex,
          reasonIndex
        });
      }
    } else {
      validation.checks[validation.checks.length - 1].status = 'passed';
    }
  }

  // Additional logical constraints can be added here
  // Example: timestamp should be first, email/name should be last, etc.

  return {
    mapping: validatedMapping,
    confidence: validatedConfidence,
    validation
  };
}

/**
 * Generate recommended column mapping with high precision
 */
function generateRecommendedMapping(headers, options = {}) {
  try {
    const targetFields = options.fields || ['answer', 'reason', 'class', 'name', 'email'];
    const mapping = {};
    const confidence = {};
    const usedIndices = new Set();

    // Sort fields by priority (answer first)
    const sortedFields = targetFields.sort((a, b) => {
      const priority = { answer: 10, reason: 8, email: 6, name: 4, class: 2 };
      return (priority[b] || 0) - (priority[a] || 0);
    });

    // Resolve each field
    for (const fieldType of sortedFields) {
      const result = resolveColumnIndex(headers, fieldType, {}, options);

      if (result.index !== -1 && !usedIndices.has(result.index)) {
        mapping[fieldType] = result.index;
        confidence[fieldType] = result.confidence;
        usedIndices.add(result.index);
      }
    }

    // Apply cross-field logical validation
    const validatedMapping = validateFieldOrderLogic(mapping, confidence, headers);

    const avgConfidence = Object.keys(validatedMapping.mapping).length > 0 ?
      Math.round(Object.values(validatedMapping.confidence).reduce((sum, c) => sum + c, 0) / Object.keys(validatedMapping.mapping).length) : 0;

    return {
      recommendedMapping: validatedMapping.mapping,
      confidence: validatedMapping.confidence,
      analysis: {
        resolvedFields: Object.keys(validatedMapping.mapping).length,
        totalFields: targetFields.length,
        overallScore: avgConfidence,
        logicalValidation: validatedMapping.validation
      },
      success: true
    };

  } catch (error) {
    console.error('generateRecommendedMapping error:', error.message);
    return {
      recommendedMapping: {},
      confidence: {},
      analysis: { error: error.message },
      success: false
    };
  }
}

/**
 * Integrated column diagnostics for frontend
 */
function performIntegratedColumnDiagnostics(originalHeaders, options = {}) {
  try {
    const result = generateRecommendedMapping(originalHeaders, options);

    return {
      success: true,
      headers: originalHeaders,
      recommendedMapping: result.recommendedMapping,
      confidence: result.confidence,
      aiAnalysis: result.analysis,
      timestamp: new Date().toISOString()
    };

  } catch (error) {
    console.error('performIntegratedColumnDiagnostics error:', error.message);
    return {
      success: false,
      error: error.message,
      headers: originalHeaders,
      recommendedMapping: {},
      confidence: {}
    };
  }
}

/**
 * Extract field value from row data
 */
function extractFieldValueUnified(row, originalHeaders, fieldType, columnMapping = {}, options = {}) {
  try {
    const result = resolveColumnIndex(originalHeaders, fieldType, columnMapping, options);

    // 厳密な検証: インデックス、行データ、値の存在をチェック
    if (result.index !== -1 && row && result.index < row.length && row[result.index] !== undefined) {
      const extractedValue = row[result.index];

      // 値の妥当性チェック: 空文字、null、undefinedは null として扱う
      if (extractedValue === null || extractedValue === undefined || String(extractedValue).trim() === '') {
        return { value: null, index: -1, confidence: 0, method: 'empty_value' };
      }


      return {
        value: extractedValue,
        index: result.index,
        confidence: result.confidence,
        method: result.method
      };
    }

    // 検出失敗時は確実に null を返す（フォールバックなし）
    return { value: null, index: -1, confidence: 0, method: 'not_found' };

  } catch (error) {
    console.error(`extractFieldValueUnified error for ${fieldType}:`, error.message);
    return { value: null, index: -1, confidence: 0, method: 'error', error: error.message };
  }
}

