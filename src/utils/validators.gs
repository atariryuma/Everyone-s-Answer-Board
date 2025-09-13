/**
 * @fileoverview Input Validation & Sanitization
 * 
 * 🎯 責任範囲:
 * - 入力データの検証・サニタイズ
 * - セキュリティ関連バリデーション
 * - データ整合性チェック
 * - フォーマット検証
 * 
 * 🔄 移行元:
 * - SecurityService の検証機能
 * - ConfigService の設定検証
 * - 各所に分散している検証処理
 */

/* global CONSTANTS, SECURITY, SecurityValidator, SecurityService, URL */

/**
 * InputValidator - 統一入力検証システム
 * SecurityServiceとの連携でセキュアな検証を実現
 */
const InputValidator = Object.freeze({

  // ===========================================
  // 🔒 基本データ型検証
  // ===========================================

  /**
   * メールアドレス検証（SecurityServiceに統一）
   * @param {string} email - メールアドレス
   * @returns {Object} 検証結果
   */
  validateEmail(email) {
    // SecurityServiceに統一 - レガシー互換性維持
    return SecurityService.validateEmail(email);
  },

  /**
   * URL検証（Google関連のみ許可）
   * @param {string} url - URL
   * @returns {Object} 検証結果
   */
  validateUrl(url) {
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
        } catch (e1) {
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
      result.sanitized = `${parsed.protocol}//${parsed.hostname}${parsed.pathname}${parsed.search}${parsed.hash}`;
      result.metadata = {
        protocol: parsed.protocol,
        hostname: parsed.hostname,
        pathname: parsed.pathname
      };

    } catch (urlError) {
      result.errors.push('無効なURL形式');
    }

    return result;
  },

  /**
   * テキスト検証・サニタイズ（XSS対策込み）
   * @param {string} text - テキスト
   * @param {Object} options - オプション
   * @returns {Object} 検証結果
   */
  validateText(text, options = {}) {
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

    if (typeof text !== 'string') {
      result.errors.push('テキストが必要です');
      return result;
    }

    let sanitized = text;
    const originalLength = text.length;

    // 長さチェック
    if (sanitized.length < minLength) {
      result.errors.push(`${minLength}文字以上である必要があります`);
      return result;
    }

    if (sanitized.length > maxLength) {
      result.errors.push(`${maxLength}文字以下である必要があります`);
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
  },

  // ===========================================
  // 📊 設定・構造検証
  // ===========================================

  /**
   * スプレッドシートID検証
   * @param {string} spreadsheetId - スプレッドシートID
   * @returns {Object} 検証結果
   */
  validateSpreadsheetId(spreadsheetId) {
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
  },

  /**
   * 列マッピング検証
   * @param {Object} columnMapping - 列マッピング
   * @returns {Object} 検証結果
   */
  validateColumnMapping(columnMapping) {
    const result = {
      isValid: false,
      errors: [],
      warnings: []
    };

    if (!columnMapping || typeof columnMapping !== 'object') {
      result.errors.push('列マッピングが必要です');
      return result;
    }

    if (!columnMapping.mapping || typeof columnMapping.mapping !== 'object') {
      result.errors.push('列マッピングの mapping プロパティが必要です');
      return result;
    }

    const {mapping} = columnMapping;
    const requiredColumns = ['answer'];
    const optionalColumns = ['reason', 'class', 'name'];
    const allColumns = [...requiredColumns, ...optionalColumns];

    // 必須列チェック
    for (const col of requiredColumns) {
      const index = mapping[col];
      if (typeof index !== 'number' || index < 0 || !Number.isInteger(index)) {
        result.errors.push(`必須列 '${col}' が正しくマッピングされていません`);
      }
    }

    // オプション列チェック
    for (const col of optionalColumns) {
      const index = mapping[col];
      if (index !== undefined) {
        if (typeof index !== 'number' || index < 0 || !Number.isInteger(index)) {
          result.warnings.push(`オプション列 '${col}' のマッピングが無効です`);
        }
      }
    }

    // 重複チェック
    const usedIndices = Object.values(mapping).filter(index => typeof index === 'number');
    const uniqueIndices = [...new Set(usedIndices)];
    if (usedIndices.length !== uniqueIndices.length) {
      result.errors.push('列インデックスに重複があります');
    }

    // 未知の列チェック
    for (const col of Object.keys(mapping)) {
      if (!allColumns.includes(col)) {
        result.warnings.push(`未知の列タイプ '${col}' が含まれています`);
      }
    }

    result.isValid = result.errors.length === 0;
    return result;
  },

  /**
   * 設定オブジェクト総合検証
   * @param {Object} config - 設定オブジェクト
   * @returns {Object} 検証結果
   */
  validateConfig(config) {
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
      spreadsheetId: { validator: this.validateSpreadsheetId.bind(this), required: false },
      formUrl: { validator: this.validateUrl.bind(this), required: false },
      sheetName: { validator: (v) => this.validateText(v, { maxLength: 100 }), required: false }
    };

    for (const [field, { validator, required }] of Object.entries(fields)) {
      const value = config[field];
      
      if (required && !value) {
        result.errors.push(`必須フィールド '${field}' が不足しています`);
        continue;
      }

      if (value) {
        const validation = validator(value);
        if (!validation.isValid) {
          result.errors.push(`${field}: ${validation.errors.join(', ')}`);
        } else {
          result.sanitized[field] = validation.sanitized || value;
        }
      }
    }

    // 列マッピング検証
    if (config.columnMapping) {
      const mappingValidation = this.validateColumnMapping(config.columnMapping);
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
    const booleanFields = ['appPublished', 'showNames', 'showReactions'];
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
  },

  // ===========================================
  // 🔧 ユーティリティ・診断
  // ===========================================

  /**
   * バリデーター診断
   * @returns {Object} 診断結果
   */
  diagnose() {
    const tests = [
      {
        name: 'Email Validation',
        test: () => this.validateEmail('test@example.com').isValid
      },
      {
        name: 'URL Validation',
        test: () => this.validateUrl('https://docs.google.com/spreadsheets/d/test').isValid
      },
      {
        name: 'Text Sanitization',
        test: () => this.validateText('<script>alert("test")</script>').sanitized.includes('[REMOVED_FOR_SECURITY]')
      },
      {
        name: 'Spreadsheet ID Validation',
        test: () => this.validateSpreadsheetId('1234567890123456789012345678901234567890abcd').isValid
      }
    ];

    const results = tests.map(({ name, test }) => {
      try {
        return { name, status: test() ? '✅' : '❌', error: null };
      } catch (error) {
        return { name, status: '❌', error: error.message };
      }
    });

    return {
      service: 'InputValidator',
      timestamp: new Date().toISOString(),
      tests: results,
      overall: results.every(r => r.status === '✅') ? '✅' : '⚠️'
    };
  }

});

/**
 * レガシー互換関数
 * 既存コードとの互換性維持
 */

// SecurityServiceからの移行
function validateUserData(userData) {
  return InputValidator.validateConfig(userData);
}

// validateEmail - SecurityServiceに統一 (グローバル関数削除済み)

function validateUrl(url) {
  return InputValidator.validateUrl(url);
}

// ConfigServiceからの移行
function validateAndSanitizeConfig(config) {
  return InputValidator.validateConfig(config);
}
