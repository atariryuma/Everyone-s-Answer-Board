/**
 * @fileoverview Input Validation & Sanitization
 *
 * 🎯 責任範囲:
 * - 入力データの検証・サニタイズ
 * - セキュリティ関連バリデーション
 * - データ整合性チェック
 * - フォーマット検証
 *
 * 🔄 GAS Best Practices準拠:
 * - フラット関数構造 (Object.freeze削除)
 * - 直接的な関数エクスポート
 * - シンプルなユーティリティ関数群
 */

/* global URL, getColumnAnalysis, getFormInfo */

// ===========================================
// 🔒 基本データ型検証関数群
// ===========================================

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

    // 基本的なメールアドレス形式チェック
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(email.trim())) {
      result.errors.push('無効なメールアドレス形式です');
      return result;
    }

    result.isValid = true;
    result.sanitized = email.trim().toLowerCase();
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
          // fall through to regex parser
        }
      }

      if (!parsed) {
        // Regex-based parser for GAS backend where URL may be undefined
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

      // 許可されたプロトコル
      const allowedProtocols = ['https:', 'http:'];
      if (!allowedProtocols.includes(parsed.protocol)) {
        result.errors.push('HTTPSまたはHTTPプロトコルが必要です');
        return result;
      }

      // 許可されたドメイン（Google関連のみ）
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

    // 🛡️ 型チェック強化 - null/undefined/非string型の完全検証
    if (text === null || text === undefined || typeof text !== 'string') {
      result.errors.push('有効なテキスト文字列が必要です');
      result.metadata.inputType = typeof text;
      result.metadata.inputValue = text === null ? 'null' : text === undefined ? 'undefined' : String(text);
      return result;
    }

    let sanitized = text;
    const originalLength = text.length;

    // 長さチェック
    if (sanitized.length < minLength) {
      result.errors.push(`${minLength  }文字以上である必要があります`);
      return result;
    }

    if (sanitized.length > maxLength) {
      result.errors.push(`${maxLength  }文字以下である必要があります`);
      return result;
    }

    // セキュリティチェック（XSS対策）
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

    // HTMLサニタイズ（許可されていない場合）
    if (!allowHtml) {
      sanitized = sanitized
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#x27;');
    }

    // 改行処理
    if (!allowNewlines) {
      sanitized = sanitized.replace(/[\r\n]/g, ' ');
    } else {
      sanitized = sanitized.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    }

    // 連続空白の正規化
    sanitized = sanitized.replace(/\s+/g, ' ').trim();

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

// ===========================================
// 📊 設定・構造検証
// ===========================================

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

    // Google Sheets IDの形式チェック
    const idPattern = /^[a-zA-Z0-9-_]{44}$/;
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
    console.log('🔍 validateMapping開始:', JSON.stringify(columnMapping, null, 2));

    const result = {
      isValid: false,
      errors: [],
      warnings: []
    };

    if (!columnMapping || typeof columnMapping !== 'object') {
      const errorMsg = '列マッピングが必要です';
      console.log(`❌ ${  errorMsg}`);
      result.errors.push(errorMsg);
      return result;
    }

    // ✅ 構造判定: 複雑構造 {mapping: {...}} vs シンプル構造 {answer: 4, class: 2}
    let actualMapping = columnMapping;
    if (columnMapping.mapping && typeof columnMapping.mapping === 'object') {
      console.log('🔄 validateMapping: 複雑構造を検出 - mapping プロパティを使用');
      actualMapping = columnMapping.mapping;
    } else {
      console.log('🔄 validateMapping: シンプル構造を検出 - 直接使用');
    }

    if (Object.keys(actualMapping).length === 0) {
      const errorMsg = '列マッピングデータが必要です';
      console.log(`❌ ${errorMsg}`);
      result.errors.push(errorMsg);
      return result;
    }

    console.log('✅ validateMapping: 使用するマッピング:', JSON.stringify(actualMapping, null, 2));

    // ✅ 構造に対応した検証
    const requiredColumns = ['answer'];
    const optionalColumns = ['reason', 'class', 'name'];
    const allColumns = [...requiredColumns, ...optionalColumns];

    // 必須列チェック
    for (const col of requiredColumns) {
      const index = actualMapping[col];
      console.log(`🔍 validateMapping: ${col} = ${index} (type: ${typeof index})`);
      if (typeof index !== 'number' || index < 0 || !Number.isInteger(index)) {
        const errorMsg = `必須列 '${col}' のインデックスが無効です（値: ${index}）`;
        console.log(`❌ ${errorMsg}`);
        result.errors.push(errorMsg);
      }
    }

    // オプション列チェック
    for (const col of optionalColumns) {
      const index = actualMapping[col];
      if (index !== undefined) {
        console.log(`🔍 validateMapping (optional): ${col} = ${index} (type: ${typeof index})`);
        if (typeof index !== 'number' || index < 0 || !Number.isInteger(index)) {
          const warningMsg = `オプション列 '${col}' のインデックスが無効です（値: ${index}）`;
          console.log(`⚠️ ${warningMsg}`);
          result.warnings.push(warningMsg);
        }
      }
    }

    // 重複チェック
    const validColumns = Object.keys(actualMapping).filter(key => allColumns.includes(key));
    const usedIndices = validColumns
      .map(col => actualMapping[col])
      .filter(index => typeof index === 'number');
    const uniqueIndices = [...new Set(usedIndices)];
    if (usedIndices.length !== uniqueIndices.length) {
      result.errors.push('列インデックスに重複があります');
    }

    // 未知の列チェック
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

    // 基本フィールド検証
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

    // 列マッピング検証
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

    // ブール値フィールド
    const booleanFields = ['isPublished', 'showNames', 'showReactions'];
    booleanFields.forEach(field => {
      if (config[field] !== undefined) {
        result.sanitized[field] = Boolean(config[field]);
      }
    });

    // 設定サイズチェック（32KB制限）
    const configSize = JSON.stringify(result.sanitized).length;
    if (configSize > 32000) {
      result.errors.push('設定データが大きすぎます（32KB制限）');
    }

    result.isValid = result.errors.length === 0;
    return result;
}

// ===========================================
// 🔧 ユーティリティ・診断
// ===========================================

/**
 * バリデーター診断
 * @returns {Object} 診断結果
 */
function diagnoseValidators() {
  const result = {
    service: 'Validators',
    timestamp: new Date().toISOString(),
    checks: []
  };

  // 基本バリデーター関数のテスト
  try {
    const emailTest = validateEmail('test@example.com');
    result.checks.push({
      name: 'Email Validator',
      status: emailTest.isValid ? '✅' : '❌',
      details: emailTest.isValid ? 'Email validation working' : 'Email validation failed'
    });
  } catch (error) {
    result.checks.push({
      name: 'Email Validator',
      status: '❌',
      details: error.message
    });
  }

  // URL バリデーターのテスト
  try {
    const urlTest = validateUrl('https://docs.google.com/spreadsheets/d/test');
    result.checks.push({
      name: 'URL Validator',
      status: urlTest.isValid ? '✅' : '❌',
      details: urlTest.isValid ? 'URL validation working' : 'URL validation failed'
    });
  } catch (error) {
    result.checks.push({
      name: 'URL Validator',
      status: '❌',
      details: error.message
    });
  }

  // テキストバリデーターのテスト
  try {
    const textTest = validateText('テストテキスト');
    result.checks.push({
      name: 'Text Validator',
      status: textTest.isValid ? '✅' : '❌',
      details: textTest.isValid ? 'Text validation working' : 'Text validation failed'
    });
  } catch (error) {
    result.checks.push({
      name: 'Text Validator',
      status: '❌',
      details: error.message
    });
  }

  result.overall = result.checks.every(check => check.status === '✅') ? '✅' : '⚠️';
  return result;
}

/**
 * フォーム-スプレッドシート整合性検証
 * @param {string} formUrl - Google Forms URL
 * @param {string} spreadsheetId - Spreadsheet ID
 * @returns {Object} 検証結果
 */
function validateFormLink(formUrl, spreadsheetId) {
  const result = {
    isValid: false,
    consistent: false,
    errors: [],
    details: {}
  };

  try {
    // 基本URL検証
    const formValidation = validateUrl(formUrl);
    const sheetValidation = validateSpreadsheetId(spreadsheetId);

    if (!formValidation.isValid) {
      result.errors.push('Invalid form URL format');
      return result;
    }

    if (!sheetValidation.isValid) {
      result.errors.push('Invalid spreadsheet ID format');
      return result;
    }

    // Google Forms URL pattern check
    const formsPattern = /^https:\/\/docs\.google\.com\/forms\/d\/([a-zA-Z0-9-_]+)/;
    const formMatch = formUrl.match(formsPattern);

    if (!formMatch) {
      result.errors.push('URL is not a valid Google Forms URL');
      return result;
    }

    const [, formId] = formMatch;
    result.details.formId = formId;
    result.details.spreadsheetId = spreadsheetId;

    // Basic validation passed
    result.isValid = true;

    // Consistency check (if both IDs are different, it might be intentional)
    // This is a warning rather than an error for flexibility
    if (formId !== spreadsheetId) {
      result.details.warning = 'Form ID and Spreadsheet ID are different - ensure this is intentional';
      result.consistent = false;
    } else {
      result.consistent = true;
    }

    // 実際の接続テスト（より確実な検証）
    try {
      // 🛡️ Zero-Dependency安全: スプレッドシート接続確認
      const hasColumnAnalysis = typeof getColumnAnalysis === 'function';
      if (hasColumnAnalysis) {
        try {
          const connectionTest = getColumnAnalysis(spreadsheetId, 'フォームの回答 1');
          if (connectionTest && connectionTest.success) {
            result.details.connectionVerified = true;
            result.details.sheetAccessible = true;
          } else {
            result.details.connectionVerified = false;
            result.details.connectionError = connectionTest?.errorResponse?.message || connectionTest?.message || 'Connection failed';
          }
        } catch (connectionError) {
          result.details.connectionVerified = false;
          result.details.connectionError = connectionError && connectionError.message ? connectionError.message : '接続テスト実行エラー';
        }
      } else {
        result.details.connectionVerified = false;
        result.details.connectionError = 'getColumnAnalysis関数が利用できません';
      }

      // 🛡️ Zero-Dependency安全: フォーム情報取得・検証
      const hasFormInfo = typeof getFormInfo === 'function';
      if (hasFormInfo) {
        try {
          const formInfoTest = getFormInfo(spreadsheetId, 'フォームの回答 1');
          if (formInfoTest && formInfoTest.success) {
            result.details.formInfoVerified = true;
            result.details.formData = formInfoTest.formData;

            // フォームURLが取得できた場合、URLの一致確認
            if (formInfoTest.formData && formInfoTest.formData.formUrl) {
              const detectedFormUrl = formInfoTest.formData.formUrl;
              result.details.detectedFormUrl = detectedFormUrl;
              result.details.formUrlMatches = (detectedFormUrl === formUrl);
            }
          } else {
            result.details.formInfoVerified = false;
            result.details.formInfoError = formInfoTest?.message || 'Form info retrieval failed';
          }
        } catch (formInfoError) {
          result.details.formInfoVerified = false;
          result.details.formInfoError = formInfoError && formInfoError.message ? formInfoError.message : 'フォーム情報取得実行エラー';
        }
      } else {
        result.details.formInfoVerified = false;
        result.details.formInfoError = 'getFormInfo関数が利用できません';
      }
    } catch (connectionError) {
      result.details.connectionVerified = false;
      result.details.connectionError = connectionError.message;
    }

    return result;
  } catch (error) {
    result.errors.push(error && error.message ? `Validation error: ${error.message}` : 'Validation error: 詳細不明');
    return result;
  }
}
