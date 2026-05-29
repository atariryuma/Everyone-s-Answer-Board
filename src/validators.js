/**
 * @fileoverview 入力検証・サニタイズ（email / URL / spreadsheet ID / config 等）。
 */

/* global SYSTEM_LIMITS */

// Why module-level: validateConfig が boardMode を sanitize する際に使う enum。
//   ConfigService.js:285 の VALID_BOARD_MODES とミラー。GAS single-global-scope で
//   実行時には ConfigService 側の定義と統合されるが、ユニットテストでは
//   validators.js を単独で vm.runInContext するため、こちらにも宣言が必要。
//   ⚠️ 値を変える際は ConfigService.js:285 も同時更新すること。
//   詳細は docs/SPEC_visualization_modes.md §F-1 参照。
const VALIDATOR_BOARD_MODES = Object.freeze(['auto', 'board', 'numberline', 'matrix', 'wordcloud', 'pie']);


/**
 * メールアドレス検証
 * @param {string} email - メールアドレス
 * @returns {Object} 検証結果
 */
function validateEmail(email) {
  const result = {
    isValid: false,
    sanitized: null,
    errors: []
  };

  try {
    if (!email || typeof email !== 'string') {
      result.errors.push('メールアドレスが必要です');
      return result;
    }

    const trimmed = email.trim();
    // Why: RFC 5321 では local part 64 chars、domain 255 chars、合計 320 chars が上限。
    //   3-char domain (例: "a@b.c") は formal RFC 上 valid だが Google Workspace では
    //   通らないので、現実的に min 6 chars 全体で reject する（"a@b.cd" 以上を要求）。
    //   上限は安全マージンで 254 (Sheets セル最大値を考慮)。
    if (trimmed.length < 6 || trimmed.length > 254) {
      result.errors.push('メールアドレスの長さが不正です (6〜254 文字)');
      return result;
    }

    // 各セグメントを区切り直し、 連続ドット (a@b..co) / 先頭末尾ドット (.a@b.com / a@.com /
    //   a@b.com.) を拒否する。 local は dot 区切りの非空セグメント、 domain は 1 つ以上の
    //   サブドメイン + TLD(2 文字以上)。 旧 /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/ は `[^\s@]+` が
    //   `..` や先頭/末尾ドットを素通ししていた。
    const emailRegex = /^[^\s@.]+(?:\.[^\s@.]+)*@[^\s@.]+(?:\.[^\s@.]+)*\.[^\s@.]{2,}$/;
    if (!emailRegex.test(trimmed)) {
      result.errors.push('無効なメールアドレス形式です');
      return result;
    }

    result.isValid = true;
    result.sanitized = trimmed.toLowerCase();
    return result;
  } catch (error) {
    result.errors.push('メール検証エラー');
    return result;
  }
}

/**
 * URL検証（Google関連のみ許可）
 * @param {string} url - URL
 * @returns {Object} 検証結果
 */
function validateUrl(url) {
    const result = {
      isValid: false,
      sanitized: '',
      errors: [],
      metadata: {}
    };

    if (!url || typeof url !== 'string') {
      result.errors.push('URLが必要です');
      return result;
    }

    try {
      let parsed = null;
      if (typeof URL !== 'undefined') {
        try {
          const u = new URL(url);
          parsed = {
            protocol: u.protocol,
            hostname: u.hostname,
            pathname: u.pathname,
            search: u.search || '',
            hash: u.hash || ''
          };
        } catch {
        }
      }

      if (!parsed) {
        const m = String(url).match(new RegExp('^(https?)://([^/?#]+)([^?#]*)([?][^#]*)?(#.*)?$', 'i'));
        if (!m) throw new Error('invalid');
        parsed = {
          protocol: (`${m[1]  }:`).toLowerCase(),
          hostname: m[2].toLowerCase(),
          pathname: m[3] || '/',
          search: m[4] || '',
          hash: m[5] || ''
        };
      }

      const allowedProtocols = ['https:', 'http:'];
      if (!allowedProtocols.includes(parsed.protocol)) {
        result.errors.push('HTTPSまたはHTTPプロトコルが必要です');
        return result;
      }

      const allowedDomains = [
        'docs.google.com',
        'forms.gle',
        'script.google.com',
        'drive.google.com',
        'sheets.google.com'
      ];

      const isAllowedDomain = allowedDomains.some(domain => 
        parsed.hostname === domain || parsed.hostname.endsWith(`.${domain}`)
      );

      if (!isAllowedDomain) {
        result.errors.push('許可されていないドメインです（Google関連サービスのみ）');
        return result;
      }

      result.isValid = true;
      result.sanitized = `${parsed.protocol  }//${  parsed.hostname  }${parsed.pathname  }${parsed.search  }${parsed.hash}`;
      result.metadata = {
        protocol: parsed.protocol,
        hostname: parsed.hostname,
        pathname: parsed.pathname
      };

    } catch {
      result.errors.push('無効なURL形式');
    }

    return result;
}

/**
 * テキスト検証・サニタイズ（XSS対策込み）
 * @param {string} text - テキスト
 * @param {Object} options - オプション
 * @returns {Object} 検証結果
 */
function validateText(text, options = {}) {
    const {
      maxLength = 8192,
      minLength = 0,
      allowHtml = false,
      allowNewlines = true
    } = options;

    const result = {
      isValid: false,
      sanitized: '',
      errors: [],
      warnings: [],
      metadata: {}
    };

    if (text === null || text === undefined || typeof text !== 'string') {
      result.errors.push('有効なテキスト文字列が必要です');
      const inputType = text === null ? 'null' : text === undefined ? 'undefined' : typeof text;
      result.metadata.inputType = inputType;
      result.metadata.inputValue = text == null ? inputType : String(text);
      return result;
    }

    let sanitized = text;
    const originalLength = text.length;

    if (sanitized.length < minLength) {
      result.errors.push(`${minLength  }文字以上である必要があります`);
      return result;
    }

    if (sanitized.length > maxLength) {
      result.errors.push(`${maxLength  }文字以下である必要があります`);
      return result;
    }

    const dangerousPatterns = [
      /<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi,
      /javascript:/gi,
      /on\w+\s*=/gi,
      /<iframe\b[^>]*>/gi,
      /<object\b[^>]*>/gi,
      /<embed\b[^>]*>/gi,
      /<link\b[^>]*>/gi,
      /<meta\b[^>]*>/gi
    ];

    let hasSecurityRisk = false;
    dangerousPatterns.forEach(pattern => {
      if (pattern.test(sanitized)) {
        hasSecurityRisk = true;
        sanitized = sanitized.replace(pattern, '[REMOVED_FOR_SECURITY]');
        result.warnings.push('セキュリティリスクのあるコンテンツを除去しました');
      }
    });

    if (!allowHtml) {
      sanitized = sanitized
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#x27;');
    }

    if (!allowNewlines) {
      sanitized = sanitized.replace(/[\r\n]/g, ' ');
      // 改行なしモードは全空白を 1 つに圧縮。
      sanitized = sanitized.replace(/\s+/g, ' ').trim();
    } else {
      // 改行を保持するモード。 旧実装は直後の \s+→space で保持したはずの \n を潰しており
      //   allowNewlines が事実上 no-op だった。 行内の連続空白のみ圧縮し、 改行は維持する
      //   (過剰な空行は 2 連までに抑制)。
      sanitized = sanitized
        .replace(/\r\n/g, '\n')
        .replace(/\r/g, '\n')
        .replace(/[ \t]+/g, ' ')
        .replace(/[ \t]*\n[ \t]*/g, '\n')
        .replace(/\n{3,}/g, '\n\n')
        .trim();
    }

    result.isValid = true;
    result.sanitized = sanitized;
    result.metadata = {
      originalLength,
      sanitizedLength: sanitized.length,
      hadSecurityRisk: hasSecurityRisk,
      wasModified: originalLength !== sanitized.length || text !== sanitized
    };

    return result;
}




/**
 * スプレッドシートID検証
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {Object} 検証結果
 */
function validateSpreadsheetId(spreadsheetId) {
    const result = {
      isValid: false,
      errors: []
    };

    if (!spreadsheetId || typeof spreadsheetId !== 'string') {
      result.errors.push('スプレッドシートIDが必要です');
      return result;
    }

    const idPattern = /^[a-zA-Z0-9-_]{40,50}$/;
    if (!idPattern.test(spreadsheetId)) {
      result.errors.push('無効なスプレッドシートID形式');
      return result;
    }

    result.isValid = true;
    return result;
}

/**
 * 列マッピング検証
 * @param {Object} columnMapping - 列マッピング
 * @returns {Object} 検証結果
 */
function validateMapping(columnMapping) {
    const result = {
      isValid: false,
      errors: [],
      warnings: []
    };

    if (!columnMapping || typeof columnMapping !== 'object') {
      const errorMsg = '列マッピングが必要です';
      result.errors.push(errorMsg);
      return result;
    }

    let actualMapping = columnMapping;
    if (columnMapping.mapping && typeof columnMapping.mapping === 'object') {
      actualMapping = columnMapping.mapping;
    }

    if (Object.keys(actualMapping).length === 0) {
      const errorMsg = '列マッピングデータが必要です';
      result.errors.push(errorMsg);
      return result;
    }

    const requiredColumns = ['answer'];
    // Why: numericX/numericY は可視化モード（M1/M2）専用のオプション列指標。
    //      未指定でも答え列が機能するので optional。重複検出と未知列警告の対象にも入れる。
    const optionalColumns = ['reason', 'class', 'name', 'numericX', 'numericY'];
    const allColumns = [...requiredColumns, ...optionalColumns];

    for (const col of requiredColumns) {
      const index = actualMapping[col];
      if (index === undefined || typeof index !== 'number' || index < 0 || !Number.isInteger(index)) {
        const errorMsg = `必須列 '${col}' のインデックスが無効です（値: ${index}）`;
        result.errors.push(errorMsg);
      }
    }

    for (const col of optionalColumns) {
      const index = actualMapping[col];
      if (index !== undefined) {
        if (typeof index !== 'number' || index < 0 || !Number.isInteger(index)) {
          const warningMsg = `オプション列 '${col}' のインデックスが無効です（値: ${index}）`;
          result.warnings.push(warningMsg);
        }
      }
    }

    // Why: numericX/numericY は「線形尺度値の数値解釈」を意味する仮想的な別軸で、
    //      物理的には answer 列（または別列）と同じインデックスを指してよい。
    //      例: 「立場」列 = answer(4) としても見れるし numericX(4) としても見れる。
    //      この重複は意図的なので uniqueness チェックから除外する。
    const DEDUP_EXEMPT = new Set(['numericX', 'numericY']);
    const dedupColumns = Object.keys(actualMapping)
      .filter(key => allColumns.includes(key) && !DEDUP_EXEMPT.has(key));
    const usedIndices = dedupColumns
      .map(col => actualMapping[col])
      .filter(index => typeof index === 'number');
    const uniqueIndices = [...new Set(usedIndices)];
    if (usedIndices.length !== uniqueIndices.length) {
      result.errors.push('列インデックスに重複があります');
    }

    for (const col of Object.keys(actualMapping)) {
      if (!allColumns.includes(col)) {
        result.warnings.push(`未知の列タイプ '${col}' が含まれています`);
      }
    }

    result.isValid = result.errors.length === 0;
    return result;
}

/**
 * 設定オブジェクト総合検証
 * @param {Object} config - 設定オブジェクト
 * @returns {Object} 検証結果
 */
function validateConfig(config) {
    const result = {
      isValid: false,
      errors: [],
      warnings: [],
      sanitized: {}
    };

    if (!config || typeof config !== 'object') {
      result.errors.push('設定オブジェクトが必要です');
      return result;
    }

    const fields = {
      spreadsheetId: { validator: validateSpreadsheetId, required: false },
      formUrl: { validator: validateUrl, required: false },
      sheetName: { validator: (v) => validateText(v, { maxLength: 100 }), required: false }
    };

    for (const [field, { validator, required }] of Object.entries(fields)) {
      const value = config[field];
      
      if (required && !value) {
        result.errors.push(`必須フィールド '${  field  }' が不足しています`);
        continue;
      }

      if (value) {
        const validation = validator(value);
        if (!validation.isValid) {
          result.errors.push(`${field  }: ${  validation.errors.join(', ')}`);
        } else {
          result.sanitized[field] = validation.sanitized || value;
        }
      }
    }

    if (config.columnMapping) {
      const mappingValidation = validateMapping(config.columnMapping);
      if (!mappingValidation.isValid) {
        result.errors.push(...mappingValidation.errors);
      }
      if (mappingValidation.warnings.length > 0) {
        result.warnings.push(...mappingValidation.warnings);
      }
      if (mappingValidation.isValid) {
        result.sanitized.columnMapping = config.columnMapping;
      }
    }

    if (config.isPublished !== undefined) {
      result.sanitized.isPublished = Boolean(config.isPublished);
    }

    if (config.displaySettings && typeof config.displaySettings === 'object') {
      const sanitizedDS = {
        showNames: Boolean(config.displaySettings.showNames),
        showReactions: Boolean(config.displaySettings.showReactions),
        theme: String(config.displaySettings.theme || 'default').substring(0, 50),
        pageSize: Math.min(Math.max(Number(config.displaySettings.pageSize) || SYSTEM_LIMITS.DEFAULT_PAGE_SIZE, 1), SYSTEM_LIMITS.MAX_PAGE_SIZE)
      };
      // Why: 可視化モード設定。VALIDATOR_BOARD_MODES (module top) の enum のみ許可。
      //      ここで保持しないと、後段の ConfigService.sanitizeDisplaySettings は
      //      validateConfig が再構築した object を受け取るため永遠に boardMode が消える。
      const mode = config.displaySettings.boardMode;
      if (typeof mode === 'string' && VALIDATOR_BOARD_MODES.includes(mode)) {
        sanitizedDS.boardMode = mode;
      }
      result.sanitized.displaySettings = sanitizedDS;
    }

    const basicFields = ['userId', 'setupStatus', 'etag', 'lastAccessedAt'];
    basicFields.forEach(field => {
      if (config[field] !== undefined) {
        result.sanitized[field] = config[field];
      }
    });

    const configSize = JSON.stringify(result.sanitized).length;
    if (configSize > SYSTEM_LIMITS.CONFIG_JSON_MAX_CHARS) {
      result.errors.push('設定データが大きすぎます（32KB制限）');
    }

    result.isValid = result.errors.length === 0;
    return result;
}




