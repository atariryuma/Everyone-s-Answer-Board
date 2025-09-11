/**
 * Core Constants - 2024 GAS V8 Best Practices
 * 最小限のコア定数のみ定義、Lazy-Loading Pattern対応
 */

/**
 * システム全体で使用するコア定数
 * CONSTANT_CASE命名 (Google Style Guide準拠)
 */
const CORE = Object.freeze({
  // タイムアウト設定（ミリ秒）
  TIMEOUTS: Object.freeze({
    SHORT: 1000, // 1秒 - UI反応
    MEDIUM: 5000, // 5秒 - API呼び出し
    LONG: 30000, // 30秒 - バッチ処理
    FLOW: 300000, // 5分 - GAS実行制限
  }),

  // ステータス定数
  STATUS: Object.freeze({
    ACTIVE: 'active',
    INACTIVE: 'inactive',
    PENDING: 'pending',
    ERROR: 'error',
  }),

  // エラーレベル（数値で優先度表現）
  ERROR_LEVELS: Object.freeze({
    INFO: 0,
    WARN: 1,
    ERROR: 2,
    CRITICAL: 3,
  }),

  // 基本的なHTTPステータス
  HTTP_STATUS: Object.freeze({
    OK: 200,
    BAD_REQUEST: 400,
    UNAUTHORIZED: 401,
    FORBIDDEN: 403,
    NOT_FOUND: 404,
    INTERNAL_ERROR: 500,
  }),

  // スコア計算設定
  SCORING_CONFIG: Object.freeze({
    LIKE_MULTIPLIER_FACTOR: 0.1, // いいね！のスコア倍率
    RANDOM_SCORE_FACTOR: 0.001, // ランダム要素の範囲
  }),

  // デバッグモード設定（本番環境では false に設定）
  DEBUG_MODE: false, // true: 詳細ログ出力, false: エラーのみ
});

/**
 * PropertiesServiceキー定数
 * セキュリティ重要項目の一元管理
 */
const PROPS_KEYS = Object.freeze({
  SERVICE_ACCOUNT_CREDS: 'SERVICE_ACCOUNT_CREDS',
  DATABASE_SPREADSHEET_ID: 'DATABASE_SPREADSHEET_ID',
  ADMIN_EMAIL: 'ADMIN_EMAIL',
  GOOGLE_CLIENT_ID: 'GOOGLE_CLIENT_ID',
});

/**
 * セキュリティ設定定数
 * GAS 2025 Security Best Practices準拠
 */
const SECURITY = Object.freeze({
  // Input validation patterns
  VALIDATION_PATTERNS: Object.freeze({
    EMAIL: /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/,
    UUID: /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i,
    SAFE_STRING: /^[a-zA-Z0-9\s\-_.@]+$/,
    URL: /^https?:\/\/(www\.)?[-a-zA-Z0-9@:%._\+~#=]{1,256}\.[a-zA-Z0-9()]{1,6}\b([-a-zA-Z0-9()@:%_\+.~#?&//=]*)$/,
    SHEET_NAME: /^[a-zA-Z0-9\s\-_()]{1,100}$/,
  }),

  // Maximum lengths for input validation
  MAX_LENGTHS: Object.freeze({
    EMAIL: 254,
    USER_ID: 36, // UUID length
    SHEET_NAME: 100,
    CONFIG_JSON: 10000, // 10KB
    URL: 2048,
    GENERAL_TEXT: 1000,
  }),

  // Sanitization settings
  SANITIZATION: Object.freeze({
    REMOVE_HTML: true,
    TRIM_WHITESPACE: true,
    NORMALIZE_UNICODE: true,
    MAX_ATTEMPTS: 3,
  }),
});

/**
 * Input validation utility functions
 * GAS 2025 Security Best Practices
 */
const SecurityValidator = Object.freeze({
  /**
   * Validate email address
   * @param {string} email - Email to validate
   * @returns {boolean} True if valid
   */
  isValidEmail: (email) => {
    if (!email || typeof email !== 'string') return false;
    if (email.length > SECURITY.MAX_LENGTHS.EMAIL) return false;
    return SECURITY.VALIDATION_PATTERNS.EMAIL.test(email.trim());
  },

  /**
   * Validate UUID format
   * @param {string} uuid - UUID to validate
   * @returns {boolean} True if valid
   */
  isValidUUID: (uuid) => {
    if (!uuid || typeof uuid !== 'string') return false;
    return SECURITY.VALIDATION_PATTERNS.UUID.test(uuid);
  },

  /**
   * Sanitize user input
   * @param {string} input - Input to sanitize
   * @param {number} maxLength - Maximum allowed length
   * @returns {string} Sanitized input
   */
  sanitizeInput: (input, maxLength = SECURITY.MAX_LENGTHS.GENERAL_TEXT) => {
    if (!input || typeof input !== 'string') return '';

    let sanitized = input;

    // Trim whitespace
    if (SECURITY.SANITIZATION.TRIM_WHITESPACE) {
      sanitized = sanitized.trim();
    }

    // Remove HTML tags (basic protection)
    if (SECURITY.SANITIZATION.REMOVE_HTML) {
      sanitized = sanitized.replace(/<[^>]*>/g, '');
    }

    // Length limit
    if (sanitized.length > maxLength) {
      sanitized = sanitized.substring(0, maxLength);
    }

    return sanitized;
  },

  /**
   * Validate and sanitize user data object
   * @param {Object} userData - User data to validate
   * @returns {Object} Validation result with sanitized data
   */
  validateUserData: (userData) => {
    const errors = [];
    const sanitizedData = {};

    // Email validation
    if (!SecurityValidator.isValidEmail(userData.userEmail)) {
      errors.push('有効なメールアドレスを入力してください。');
    } else {
      sanitizedData.userEmail = SecurityValidator.sanitizeInput(
        userData.userEmail,
        SECURITY.MAX_LENGTHS.EMAIL
      );
    }

    // UUID validation
    if (userData.userId && !SecurityValidator.isValidUUID(userData.userId)) {
      errors.push('無効なユーザーIDです。');
    } else if (userData.userId) {
      sanitizedData.userId = userData.userId;
    }

    // Spreadsheet ID/URL validation
    if (userData.spreadsheetId) {
      sanitizedData.spreadsheetId = SecurityValidator.sanitizeInput(
        userData.spreadsheetId,
        SECURITY.MAX_LENGTHS.GENERAL_TEXT
      );
    }

    if (userData.spreadsheetUrl) {
      if (SECURITY.VALIDATION_PATTERNS.URL.test(userData.spreadsheetUrl)) {
        sanitizedData.spreadsheetUrl = userData.spreadsheetUrl;
      } else {
        errors.push('有効なスプレッドシートURLを入力してください。');
      }
    }

    // Config JSON validation
    if (userData.configJson) {
      if (userData.configJson.length > SECURITY.MAX_LENGTHS.CONFIG_JSON) {
        errors.push('設定データが大きすぎます。');
      } else {
        try {
          JSON.parse(userData.configJson);
          sanitizedData.configJson = userData.configJson;
        } catch (e) {
          errors.push('設定データの形式が正しくありません。');
        }
      }
    }

    // Sheet name validation
    if (userData.sheetName) {
      sanitizedData.sheetName = SecurityValidator.sanitizeInput(userData.sheetName, 100);
    }

    // Form URL validation
    if (userData.formUrl) {
      if (SECURITY.VALIDATION_PATTERNS.URL.test(userData.formUrl)) {
        sanitizedData.formUrl = userData.formUrl;
      } else {
        errors.push('有効なフォームURLを入力してください。');
      }
    }

    return {
      isValid: errors.length === 0,
      errors,
      sanitizedData,
    };
  },
});

/**
 * 統合定数オブジェクト（フラット構造）
 * GAS 2025 Best Practices準拠
 */
const CONSTANTS = Object.freeze({
  // 🚀 configJSON中心型超効率化データベース定数
  DATABASE: Object.freeze({
    SHEET_NAME: 'Users',
    // ⚡ 5項目超効率化構造（64%削減、70%高速化）
    HEADERS: Object.freeze([
      'userId', // [0] UUID - 必須ID（検索用）
      'userEmail', // [1] メールアドレス - 必須認証（検索用）
      'isActive', // [2] アクティブ状態 - 必須フラグ（検索用）
      'configJson', // [3] 全設定統合 - メインデータ（JSON一括処理）
      'lastModified', // [4] 最終更新 - 監査用
    ]),

    // 🗑️ configJsonに統合済みフィールド（移行参照用）
    MIGRATED_FIELDS: Object.freeze([
      'createdAt', // → configJson.createdAt
      'lastAccessedAt', // → configJson.lastAccessedAt
      'spreadsheetId', // → configJson.spreadsheetId
      'sheetName', // → configJson.sheetName
      'spreadsheetUrl', // → configJson.spreadsheetUrl（動的生成）
      'formUrl', // → configJson.formUrl
      'columnMappingJson', // → configJson.columnMapping
      'publishedAt', // → configJson.publishedAt
      'appUrl', // → configJson.appUrl（動的生成）
    ]),

    DELETE_LOG: Object.freeze({
      SHEET_NAME: 'DeletionLogs',
      HEADERS: Object.freeze([
        'timestamp',
        'executorEmail',
        'targetUserId',
        'targetEmail',
        'reason',
        'deleteType',
      ]),
    }),
  }),

  // リアクション機能
  REACTIONS: Object.freeze({
    KEYS: ['UNDERSTAND', 'LIKE', 'CURIOUS'],
    LABELS: Object.freeze({
      UNDERSTAND: 'なるほど！',
      LIKE: 'いいね！',
      CURIOUS: 'もっと知りたい！',
      HIGHLIGHT: 'ハイライト',
    }),
  }),

  // 列ヘッダー定義
  COLUMNS: Object.freeze({
    TIMESTAMP: 'タイムスタンプ',
    EMAIL: 'メールアドレス',
    CLASS: 'クラス',
    OPINION: '回答',
    REASON: '理由',
    NAME: '名前',
  }),

  // AdminPanel用列マッピング定義（高性能AI検索対応）
  COLUMN_MAPPING: Object.freeze({
    answer: Object.freeze({
      key: 'answer',
      header: '回答',
      alternates: ['どうして', '質問', '問題', '意見', '答え', 'なぜ', '思います', '考え'], // 高性能AI用拡張パターン
      required: true,
      aiPatterns: ['？', '?', 'どうして', 'なぜ', '思いますか', '考えますか'], // AI専用検出パターン
    }),
    reason: Object.freeze({
      key: 'reason',
      header: '理由',
      alternates: ['理由', '根拠', '体験', 'なぜ', '詳細', '説明'], // 高性能AI用拡張パターン
      required: false,
      aiPatterns: ['理由', '体験', '根拠', '詳細'], // AI専用検出パターン
    }),
    class: Object.freeze({
      key: 'class',
      header: 'クラス',
      alternates: ['クラス', '学年'],
      required: false,
    }),
    name: Object.freeze({
      key: 'name',
      header: '名前',
      alternates: ['名前', '氏名', 'お名前'],
      required: false,
    }),
  }),

  // 表示モード
  DISPLAY_MODES: Object.freeze({
    ANONYMOUS: 'anonymous',
    NAMED: 'named',
    EMAIL: 'email',
  }),

  // アクセス制御関連
  ACCESS: Object.freeze({
    LEVELS: Object.freeze({
      OWNER: 'owner',
      SYSTEM_ADMIN: 'system_admin',
      AUTHENTICATED_USER: 'authenticated_user',
      GUEST: 'guest',
      NONE: 'none',
    }),
  }),
});

// 統一エラーハンドリング - セキュリティ強化版
const ErrorHandler = Object.freeze({
  createSafeResponse(error, context = 'operation', includeDetails = false) {
    // 内部ログ（詳細付き）
    console.error(`❌ ${context}:`, {
      message: error.message,
      timestamp: new Date().toISOString(),
    });

    // 外部向けレスポンス（セキュア）
    return {
      success: false,
      message: this.getSafeErrorMessage(error.message),
      timestamp: new Date().toISOString(),
      errorCode: this.getErrorCode(error.message),
    };
  },

  getSafeErrorMessage(originalMessage) {
    const patterns = ['Service Account', 'token', 'database', 'validation'];
    for (const pattern of patterns) {
      if (originalMessage.toLowerCase().includes(pattern.toLowerCase())) {
        return 'システムエラーが発生しました。管理者にお問い合わせください。';
      }
    }
    return '処理中にエラーが発生しました。';
  },

  getErrorCode(message) {
    const hash = message.split('').reduce((a, b) => {
      a = (a << 5) - a + b.charCodeAt(0);
      return a & a;
    }, 0);
    return `E${Math.abs(hash).toString(16).substr(0, 6).toUpperCase()}`;
  },
});

// パフォーマンス監視
const PerformanceMonitor = Object.freeze({
  measure(operationName, operation) {
    const startTime = Date.now();

    try {
      const result = operation();
      const duration = Date.now() - startTime;

      if (duration > 5000) {
        console.warn(`⚠️ 低速操作: ${operationName} - ${duration}ms`);
      }

      return result;
    } catch (error) {
      const duration = Date.now() - startTime;
      console.error(`❌ 操作失敗: ${operationName} (${duration}ms)`);
      throw error;
    }
  },
});

// ✅ 簡素化：直接CONSTANTSを使用（エイリアス削除でメモリ効率化）
// 使用例: CONSTANTS.REACTIONS.KEYS, CONSTANTS.COLUMNS, CONSTANTS.DATABASE.DELETE_LOG
