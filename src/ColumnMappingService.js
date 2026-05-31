/**
 * @fileoverview ColumnMappingService - フォームヘッダー + サンプルデータから列役割を推論する。
 *
 * 統合スコアリング (L1 + L2 + L3):
 *   L1 Header pattern: ヘッダー文字列 vs 役割辞書 (exact/keywords/questionPatterns/regex)
 *   L2 Data shape:     サンプル値の型・分布から役割を補強
 *                      (integer 1..N → numericScale, email format → email,
 *                       avg length → answer/reason vs name/class)
 *   L3 Board mode:     表示モードが要求する役割への bias (pie → answer 重視 等)
 *
 * 制約:
 *   - answer は reason より前にあるべき (sequential)
 *   - 1 列 = 1 役割 (greedy / 高スコア優先で衝突解消)
 *
 * Entry points:
 *   inferColumnRoles(headers, sampleData, options)        ← 統合分析 (内部 + 直接呼び)
 *   performIntegratedColumnDiagnostics(headers, options)  ← getColumnAnalysis 用
 *   detectNumericScaleColumns(headers, sampleData)        ← 線形尺度のみ (互換)
 *   resolveColumnIndex(headers, fieldType, mapping)       ← 行処理 hot path (DataService)
 *   filterSystemColumns(headers)                          ← resolveColumnIndex 内部 + テスト
 */

/* global normalizeHeader, logError_ */

// =====================================================================
// 辞書 (L1)
// =====================================================================

const __SYSTEM_HEADER_PATTERNS = [
  /^タイムスタンプ$/i, /^timestamp$/i, /^日時$/i, /^日付$/i,
  /^UNDERSTAND$/i, /^LIKE$/i, /^CURIOUS$/i, /^HIGHLIGHT$/i,
  /^理解$/i, /^いいね$/i, /^気になる$/i, /^ハイライト$/i,
  /^_/
];

const __ROLE_PATTERNS = {
  answer: {
    exact: ['回答', 'answer'],
    keywords: ['回答', '答え', 'answer', '意見', '考え', '感想', 'コメント'],
    // 「質問らしさ」を表すパターン群。フォーム作成者が自由記述で書いた質問文を answer 列として捕捉する。
    // 末尾の ？/? と「どう...する」「なぜ」「何」も含める (児童向けフォームで頻出)。
    questionPatterns: [
      /どう.*思い?.*ますか/i, /.*と思い?.*ますか/i,
      /.*書きましょう/i, /.*書いてください/, /.*書こう/, /.*書いて$/,
      /.*述べ/i, /.*説明して/i,
      /.*気づいたこと/i, /.*観察して/i,
      /.*学んだこと/, /.*感じたこと/, /.*心に残った/,  // 振り返り系の頻出
      /.*どう.*(する|します)/, /.*なぜ/, /.*何を/, /.*何が/,
      /[？?]\s*$/,   // ヘッダー末尾の質問符 (児童向けフォームの自由質問)
      /what.*do you think/i, /how.*do you feel/i, /explain.*your/i
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

const __DEFAULT_ROLES = ['answer', 'reason', 'email', 'name', 'class'];
// 優先度: answer/reason が割り当てを先に取り、email > name > class で衝突を解消する。
const __ROLE_PRIORITY = { answer: 10, reason: 8, email: 6, name: 4, class: 2 };

// =====================================================================
// system column filtering
// =====================================================================

function __isSystemHeader(header) {
  if (!header || typeof header !== 'string') return true;
  const trimmed = header.trim();
  if (!trimmed) return true;
  return __SYSTEM_HEADER_PATTERNS.some(p => p.test(trimmed));
}

function filterSystemColumns(headers) {
  const cleanHeaders = [];
  const indexMap = [];
  if (!Array.isArray(headers)) return { cleanHeaders, indexMap };
  for (let i = 0; i < headers.length; i++) {
    const h = headers[i];
    if (!h || typeof h !== 'string') continue;
    if (!h.trim()) continue;
    if (__SYSTEM_HEADER_PATTERNS.some(p => p.test(h.trim()))) continue;
    cleanHeaders.push(h);
    indexMap.push(i);
  }
  return { cleanHeaders, indexMap };
}

// =====================================================================
// L1: ヘッダーパターンスコア (0..95)
// =====================================================================

function __headerPatternScore(header, role) {
  const patterns = __ROLE_PATTERNS[role];
  if (!patterns || !header || typeof header !== 'string') return 0;
  const normalized = normalizeHeader(header);
  let score = 0;

  if (patterns.exact.some(k => normalized === k.toLowerCase())) score = 95;
  const hasOwnKeyword = patterns.keywords.some(k => normalized.includes(k.toLowerCase()));
  if (hasOwnKeyword) score = Math.max(score, 90);
  if (role === 'answer' && patterns.questionPatterns) {
    if (patterns.questionPatterns.some(p => p.test(header))) score = Math.max(score, 87);
  }
  if (patterns.regex.test(header)) score = Math.max(score, 85);

  // Why (v2865 / M7): answer の questionPatterns / regex (例 /.*なぜ/) は reason の keyword
  //   (なぜ/どうして/理由/根拠/原因) と語彙が重なる。 ヘッダーが reason の distinctive keyword に
  //   当たり、 かつ answer 固有 keyword (回答/答え/意見/考え 等) を含まないなら、 それは「理由」列の
  //   可能性が高い。 answer スコアを reason keyword(90) より十分下げ (=55)、 理由列が誤って answer に
  //   固定されるのを防ぐ (相互排他)。 exact 一致(95)・answer 固有 keyword(90) は確実なので維持。
  //   55 に残すのは、 自由記述が 1 列だけのフォーム (例「理由を書こう」のみ) で answer が
  //   未割当にならないようにするため (greedy で唯一候補なら拾える)。
  if (role === 'answer' && score > 55 && score < 95 && !hasOwnKeyword) {
    const reasonKeywords = __ROLE_PATTERNS.reason.keywords;
    const hasReasonKeyword = reasonKeywords.some(k => normalized.includes(k.toLowerCase()));
    if (hasReasonKeyword) score = 55;
  }
  return score;
}

// =====================================================================
// L2: サンプルデータの形状分析
// =====================================================================

const __EMAIL_REGEX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

/**
 * 列の非空サンプル値から型・分布を抽出する。
 * Why: dataType を 1 回計算しておけば、複数役割への boost 判定で再走査不要。
 */
function __analyzeColumnData(values) {
  const stats = {
    sampleCount: 0,
    dataType: 'empty',
    allInteger: false,
    integerMin: Infinity, integerMax: -Infinity,
    emailRatio: 0,
    avgLength: 0, maxLength: 0,
    cardinality: 0
  };
  if (!Array.isArray(values)) return stats;

  const nonEmpty = [];
  for (const v of values) {
    if (v === '' || v == null) continue;
    nonEmpty.push(v);
  }
  if (nonEmpty.length === 0) return stats;
  stats.sampleCount = nonEmpty.length;

  let intCount = 0;
  let emailCount = 0;
  let lenSum = 0;
  const unique = new Set();

  for (const v of nonEmpty) {
    const str = String(v).trim();
    const n = typeof v === 'number' ? v : Number(str);
    if (Number.isFinite(n) && Number.isInteger(n) && n >= -100 && n <= 100) {
      intCount++;
      if (n < stats.integerMin) stats.integerMin = n;
      if (n > stats.integerMax) stats.integerMax = n;
    }
    if (__EMAIL_REGEX.test(str)) emailCount++;
    lenSum += str.length;
    if (str.length > stats.maxLength) stats.maxLength = str.length;
    unique.add(str);
  }

  stats.allInteger = (intCount === nonEmpty.length);
  stats.emailRatio = emailCount / nonEmpty.length;
  stats.avgLength = lenSum / nonEmpty.length;
  stats.cardinality = unique.size;

  // dataType 判定. integer-scale は範囲 1..9 の整数列のみ採用。
  if (stats.allInteger && stats.sampleCount >= 2
      && stats.integerMax - stats.integerMin >= 1
      && stats.integerMax - stats.integerMin <= 9) {
    stats.dataType = 'integer-scale';
  } else if (stats.emailRatio >= 0.7) {
    stats.dataType = 'email';
  } else if (stats.avgLength >= 20) {
    stats.dataType = 'long-text';
  } else if (stats.avgLength <= 8) {
    stats.dataType = 'short-text';
  } else {
    stats.dataType = 'medium-text';
  }
  return stats;
}

/**
 * データ形状から役割への boost / penalty を返す (-15..+15)。
 * Why: L1 だけだと「同名 keyword が複数列に出る」「短文 vs 長文の区別」ができない。
 *      実データの dataType と組み合わせて衝突を解消する。
 */
function __dataTypeBoost(stats, role) {
  if (!stats || stats.sampleCount === 0) return 0;
  const t = stats.dataType;
  switch (role) {
    case 'email':
      if (t === 'email') return 15;
      if (stats.emailRatio > 0) return 0;
      return -10;
    case 'answer':
      if (t === 'long-text') return 8;
      if (t === 'medium-text') return 3;
      if (t === 'integer-scale' || t === 'email') return -15;
      return 0;
    case 'reason':
      if (t === 'long-text') return 6;
      if (t === 'medium-text') return 3;
      if (t === 'integer-scale' || t === 'email') return -15;
      return 0;
    case 'name':
      if (t === 'short-text' && stats.avgLength <= 8) return 8;
      if (t === 'email' || t === 'integer-scale') return -15;
      return 0;
    case 'class':
      if (t === 'short-text' && stats.avgLength <= 10) return 5;
      // クラスは値数が少ない (1-1, 1-2, …) ことが多い → 低カーディナリティを軽く後押し。
      if (stats.cardinality > 0 && stats.cardinality <= 10 && stats.sampleCount >= 5) return 5;
      if (t === 'integer-scale') return -10;
      return 0;
    default:
      return 0;
  }
}

// =====================================================================
// L3: 表示モード bias
// =====================================================================

function __boardModeBoost(role, boardMode) {
  // numericX/Y は numericScaleCandidates 経由で別系統に確定するので、ここでは扱わない。
  // 表示モードによる「主軸となる役割」への弱い bias のみ。
  switch (boardMode) {
    case 'pie':
    case 'wordcloud':
      return role === 'answer' ? 5 : 0;
    case 'numberline':
    case 'matrix':
    case 'board':
    case 'auto':
    default:
      return 0;
  }
}

// =====================================================================
// 統合: inferColumnRoles
// =====================================================================

/**
 * ヘッダー + サンプルデータから役割マッピングを推論する。
 *
 * @param {Array<string>} headers
 * @param {Array<Array>} sampleData - 2D 配列 (ヘッダー除く)
 * @param {Object} [options]
 * @param {string} [options.boardMode='auto']
 * @param {Array<string>} [options.fields]
 * @param {Object} [options.existingMapping] - 既存マッピング (固定)
 * @param {boolean} [options.includeNumericScale=true] - numericX/Y を mapping に含めるか
 * @returns {{ mapping, confidence, columns, numericScaleCandidates }}
 */
function inferColumnRoles(headers, sampleData = [], options = {}) {
  const roles = (options.fields && options.fields.length) ? options.fields : __DEFAULT_ROLES;
  const boardMode = options.boardMode || 'auto';
  const existingMapping = options.existingMapping || {};
  const includeNumericScale = options.includeNumericScale !== false;

  if (!Array.isArray(headers) || headers.length === 0) {
    return { mapping: {}, confidence: {}, columns: [], numericScaleCandidates: [] };
  }

  // 各列のサンプル統計を 1 回計算しキャッシュ
  const columns = headers.map((header, index) => {
    const isSystem = __isSystemHeader(header);
    const values = [];
    if (!isSystem && Array.isArray(sampleData)) {
      for (const row of sampleData) {
        if (Array.isArray(row)) values.push(row[index]);
      }
    }
    return {
      index, header,
      isSystem,
      stats: __analyzeColumnData(values),
      role: null,
      confidence: 0
    };
  });

  // 役割×列 のスコア候補
  // Why: L1 (ヘッダー) で 0 でも L2 (データ形状) が役割を支持するなら候補化する。
  //      例:「あなたなら、ポスターをどうする？」は L1 辞書に該当しないが、データが
  //      medium/long-text であれば answer 候補として 40 + L2 を base に拾う。
  //      L1>0 時は最大 95、L1=0 (L2 のみ) 時は最大 60 で「弱い推定」を表現する。
  const scoresByRole = {};
  for (const role of roles) {
    const cands = [];
    for (const col of columns) {
      if (col.isSystem) continue;
      const headerScore = __headerPatternScore(col.header, role);
      const dataBoost = __dataTypeBoost(col.stats, role);
      const modeBoost = __boardModeBoost(role, boardMode);
      let composite = 0;
      if (headerScore > 0) {
        composite = Math.max(0, Math.min(95, headerScore + dataBoost + modeBoost));
      } else if (dataBoost > 0) {
        // L1 miss + L2 hit: 弱い候補 (上限 60)
        composite = Math.max(0, Math.min(60, 40 + dataBoost + modeBoost));
      }
      if (composite > 0) cands.push({ index: col.index, score: composite });
    }
    cands.sort((a, b) => b.score - a.score);
    scoresByRole[role] = cands;
  }

  // 既存マッピング適用 (固定 / 信頼度 100)
  const mapping = {};
  const confidence = {};
  const usedIndices = new Set();
  for (const role of roles) {
    const fixedIdx = existingMapping[role];
    if (typeof fixedIdx === 'number' && fixedIdx >= 0 && fixedIdx < headers.length) {
      mapping[role] = fixedIdx;
      confidence[role] = 100;
      usedIndices.add(fixedIdx);
    }
  }

  // answer / reason ペアの先行最適化。
  // Why: answer は優先度が高く greedy だと先に高スコア列を取るが、 その列が reason 候補
  //   より後ろだと「answer-before-reason 制約」 で reason が行き場を失い未割当になる
  //   (理由→意見 の順に並んだフォーム等)。 両者が未固定のときは制約 (answerIdx < reasonIdx)
  //   を満たす組合せでスコア合計が最大のペアを先に確定させ、 reason の取りこぼしを防ぐ。
  if (mapping.answer === undefined && mapping.reason === undefined
      && Array.isArray(scoresByRole.answer) && Array.isArray(scoresByRole.reason)) {
    let best = null;
    for (const a of scoresByRole.answer) {
      if (usedIndices.has(a.index)) continue;
      for (const r of scoresByRole.reason) {
        if (usedIndices.has(r.index) || r.index === a.index) continue;
        if (a.index >= r.index) continue; // answer は reason より前
        const sum = a.score + r.score;
        if (!best || sum > best.sum) {
          best = { aIdx: a.index, aScore: a.score, rIdx: r.index, rScore: r.score, sum };
        }
      }
    }
    if (best) {
      mapping.answer = best.aIdx; confidence.answer = best.aScore; usedIndices.add(best.aIdx);
      mapping.reason = best.rIdx; confidence.reason = best.rScore; usedIndices.add(best.rIdx);
    }
  }

  // Greedy 割当 (役割優先度順)
  const remaining = roles
    .filter(r => mapping[r] === undefined)
    .sort((a, b) => (__ROLE_PRIORITY[b] || 0) - (__ROLE_PRIORITY[a] || 0));

  for (const role of remaining) {
    for (const cand of scoresByRole[role]) {
      if (usedIndices.has(cand.index)) continue;
      // answer-before-reason 制約: reason は answer より後の列でなければならない
      if (role === 'reason' && typeof mapping.answer === 'number' && cand.index <= mapping.answer) continue;
      if (role === 'answer' && typeof mapping.reason === 'number' && cand.index >= mapping.reason) continue;
      mapping[role] = cand.index;
      confidence[role] = cand.score;
      usedIndices.add(cand.index);
      break;
    }
  }

  // 線形尺度候補 (L2 で integer-scale 判定された列)
  const numericScaleCandidates = __collectNumericScaleCandidates(columns);

  // numericX/Y を mapping に自動マージ (Greedy で未割当 + 信頼度 >= 80)
  if (includeNumericScale) {
    if (numericScaleCandidates.length >= 1
        && typeof mapping.numericX !== 'number'
        && numericScaleCandidates[0].confidence >= 80) {
      mapping.numericX = numericScaleCandidates[0].index;
      confidence.numericX = numericScaleCandidates[0].confidence;
    }
    if (numericScaleCandidates.length >= 2
        && typeof mapping.numericY !== 'number'
        && numericScaleCandidates[1].confidence >= 80) {
      mapping.numericY = numericScaleCandidates[1].index;
      confidence.numericY = numericScaleCandidates[1].confidence;
    }
  }

  // 列情報に役割を埋め込む (UI 表示用)
  for (const role of Object.keys(mapping)) {
    const col = columns.find(c => c.index === mapping[role]);
    if (col) {
      col.role = role;
      col.confidence = confidence[role];
    }
  }

  return { mapping, confidence, columns, numericScaleCandidates };
}

// =====================================================================
// 線形尺度検出 (公開 + 内部共用)
// =====================================================================

function __collectNumericScaleCandidates(columns) {
  const candidates = [];
  for (const col of columns) {
    if (col.isSystem) continue;
    const s = col.stats;
    if (!s || s.dataType !== 'integer-scale') continue;
    const min = s.integerMin;
    const max = s.integerMax;

    let conf = 60;
    if (s.sampleCount >= 5) conf += 10;
    if (s.sampleCount >= 15) conf += 10;
    if (min === 1 && (max === 5 || max === 10)) conf += 15;
    if (min === 0 && (max === 4 || max === 10)) conf += 10;

    candidates.push({
      index: col.index,
      header: col.header.trim(),
      min, max,
      sampleCount: s.sampleCount,
      confidence: Math.min(conf, 95)
    });
  }
  candidates.sort((a, b) => b.confidence - a.confidence);
  return candidates;
}

function detectNumericScaleColumns(headers, sampleData) {
  if (!Array.isArray(headers) || !Array.isArray(sampleData)) return [];
  const columns = headers.map((header, index) => {
    const isSystem = __isSystemHeader(header);
    const values = [];
    if (!isSystem) {
      for (const row of sampleData) {
        if (Array.isArray(row)) values.push(row[index]);
      }
    }
    return { index, header, isSystem, stats: __analyzeColumnData(values) };
  });
  return __collectNumericScaleCandidates(columns);
}

// =====================================================================
// 行処理 hot path: 単一役割解決 (DataService から per-field 呼び出し)
// =====================================================================

/**
 * 行処理 hot path 用の単一役割解決。columnMapping が指定されていればそれを返す軽量パス。
 */
function resolveColumnIndex(headers, fieldType, columnMapping = {}, options = {}) {
  try {
    if (columnMapping && typeof columnMapping[fieldType] === 'number'
        && columnMapping[fieldType] >= 0 && columnMapping[fieldType] < headers.length) {
      return { index: columnMapping[fieldType], confidence: 100, method: 'existing_mapping' };
    }

    const { cleanHeaders, indexMap } = filterSystemColumns(headers);
    if (cleanHeaders.length === 0) return { index: -1, confidence: 0, method: 'no_valid_headers' };
    if (!__ROLE_PATTERNS[fieldType]) return { index: -1, confidence: 0, method: 'unknown_field' };

    let bestIdx = -1, bestScore = 0;
    for (let i = 0; i < cleanHeaders.length; i++) {
      const score = __headerPatternScore(cleanHeaders[i], fieldType);
      if (score > bestScore) { bestScore = score; bestIdx = i; }
    }
    return bestIdx === -1
      ? { index: -1, confidence: 0, method: 'no_match' }
      : { index: indexMap[bestIdx], confidence: Math.min(bestScore, 95), method: 'pattern_match' };
  } catch (error) {
    logError_('resolveColumnIndex', error, { fieldType });
    return { index: -1, confidence: 0, method: 'error', error: error.message };
  }
}

/**
 * getColumnAnalysis から呼ばれる統合分析。
 *
 * @param {Array<string>} originalHeaders
 * @param {Object} [options]
 * @param {Array<Array>} [options.sampleData]
 * @param {string} [options.boardMode]
 * @param {boolean} [options.includeNumericScale=false] - true で numericX/Y も
 *        recommendedMapping に含める (旧 API 互換のためデフォルト false)。
 */
function performIntegratedColumnDiagnostics(originalHeaders, options = {}) {
  try {
    const includeNumericScale = options.includeNumericScale === true;
    const inferred = inferColumnRoles(originalHeaders, options.sampleData || [], {
      fields: options.fields || __DEFAULT_ROLES,
      boardMode: options.boardMode || 'auto',
      includeNumericScale
    });
    const recommendedMapping = {};
    const conf = {};
    const keys = includeNumericScale
      ? [...__DEFAULT_ROLES, 'numericX', 'numericY']
      : __DEFAULT_ROLES;
    for (const r of keys) {
      if (typeof inferred.mapping[r] === 'number') {
        recommendedMapping[r] = inferred.mapping[r];
        conf[r] = inferred.confidence[r];
      }
    }
    const resolvedCount = Object.keys(recommendedMapping).length;
    const overallScore = resolvedCount > 0
      ? Math.round(Object.values(conf).reduce((s, c) => s + c, 0) / resolvedCount)
      : 0;

    return {
      success: true,
      headers: originalHeaders,
      recommendedMapping,
      confidence: conf,
      numericScaleCandidates: inferred.numericScaleCandidates,
      columnIntelligence: inferred.columns.map(c => ({
        index: c.index, header: c.header,
        role: c.role, confidence: c.confidence,
        dataType: c.stats.dataType
      })),
      aiAnalysis: {
        resolvedFields: resolvedCount,
        totalFields: __DEFAULT_ROLES.length,
        overallScore
      },
      timestamp: new Date().toISOString()
    };
  } catch (error) {
    if (typeof logError_ === 'function') logError_('performIntegratedColumnDiagnostics', error);
    return {
      success: false,
      error: error.message,
      headers: originalHeaders,
      recommendedMapping: {},
      numericScaleCandidates: [],
      confidence: {}
    };
  }
}
