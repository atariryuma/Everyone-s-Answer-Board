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
 * Optimized column detection engine focused on precision and efficiency
 */
function detectColumn(headers, fieldType, options = {}) {
  const patterns = getFieldPatterns()[fieldType];
  if (!patterns) {
    return { index: -1, confidence: 0, method: 'unknown_field' };
  }

  let bestMatch = { index: -1, confidence: 0, method: 'no_match' };

  headers.forEach((header, index) => {
    if (!header || typeof header !== 'string') return;

    const score = calculateScore(header, fieldType, patterns, options);

    if (score > bestMatch.confidence) {
      bestMatch = {
        index,
        confidence: Math.min(score, 95),
        method: 'pattern_match'
      };
    }
  });

  return bestMatch;
}

/**
 * Calculate header score using streamlined algorithm
 */
function calculateScore(header, fieldType, patterns, options = {}) {
  const normalizedHeader = header.toLowerCase().trim();
  let maxScore = 0;

  // 1. Direct pattern matching (most important)
  if (patterns.exact && patterns.exact.some(keyword => normalizedHeader === keyword.toLowerCase())) {
    maxScore = 95;
  }

  // 2. Keyword matching
  if (patterns.keywords) {
    for (const keyword of patterns.keywords) {
      if (normalizedHeader.includes(keyword.toLowerCase())) {
        maxScore = Math.max(maxScore, 90);
        break;
      }
    }
  }

  // 3. Question pattern matching (for answer fields)
  if (fieldType === 'answer' && patterns.questionPatterns) {
    for (const pattern of patterns.questionPatterns) {
      if (pattern.test(header)) {
        maxScore = Math.max(maxScore, 87);
        break;
      }
    }
  }

  // 4. Regex patterns
  if (patterns.regex && patterns.regex.test(header)) {
    maxScore = Math.max(maxScore, 85);
  }

  // 5. Sample data validation (if available)
  if (options.sampleData && options.sampleData.length > 0 && maxScore > 0) {
    const validationBonus = validateSampleData(options.sampleData, normalizedHeader, fieldType);
    maxScore = Math.min(maxScore + validationBonus, 95);
  }

  return maxScore;
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

    const avgConfidence = Object.keys(mapping).length > 0 ?
      Math.round(Object.values(confidence).reduce((sum, c) => sum + c, 0) / Object.keys(mapping).length) : 0;

    return {
      recommendedMapping: mapping,
      confidence,
      analysis: {
        resolvedFields: Object.keys(mapping).length,
        totalFields: targetFields.length,
        overallScore: avgConfidence
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
function extractFieldValueUnified(row, originalHeaders, fieldType, options = {}) {
  try {
    const result = resolveColumnIndex(originalHeaders, fieldType, {}, options);

    if (result.index !== -1 && row && row[result.index] !== undefined) {
      return {
        value: row[result.index],
        index: result.index,
        confidence: result.confidence,
        method: result.method
      };
    }

    return { value: null, index: -1, confidence: 0, method: 'not_found' };

  } catch (error) {
    console.error(`extractFieldValueUnified error for ${fieldType}:`, error.message);
    return { value: null, index: -1, confidence: 0, method: 'error', error: error.message };
  }
}