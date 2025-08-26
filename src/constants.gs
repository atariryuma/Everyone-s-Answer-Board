/**
 * 統合システム定数管理
 * 全システムで使用する定数を一元管理し、重複を排除
 */

/**
 * @typedef {Object} ErrorSeverity
 * @property {string} LOW - 軽微なエラー
 * @property {string} MEDIUM - 中程度のエラー
 * @property {string} HIGH - 重要なエラー
 * @property {string} CRITICAL - 致命的なエラー
 */

/**
 * @typedef {Object} ErrorCategories
 * @property {string} AUTHENTICATION - 認証エラー
 * @property {string} AUTHORIZATION - 認可エラー
 * @property {string} DATABASE - データベースエラー
 * @property {string} CACHE - キャッシュエラー
 * @property {string} NETWORK - ネットワークエラー
 * @property {string} VALIDATION - バリデーションエラー
 * @property {string} SYSTEM - システムエラー
 * @property {string} USER_INPUT - ユーザー入力エラー
 * @property {string} EXTERNAL - 外部サービスエラー
 * @property {string} CONFIG - 設定エラー
 * @property {string} SECURITY - セキュリティエラー
 */

/**
 * 統合システム定数
 */
const UNIFIED_CONSTANTS = {
  // エラー処理定数
  ERROR: {
    /** @type {ErrorSeverity} エラー重大度 */
    SEVERITY: {
      LOW: 'low',
      MEDIUM: 'medium',
      HIGH: 'high', 
      CRITICAL: 'critical'
    },
    
    /** @type {ErrorCategories} エラーカテゴリ */
    CATEGORIES: {
      AUTHENTICATION: 'authentication',
      AUTHORIZATION: 'authorization',
      DATABASE: 'database',
      CACHE: 'cache',
      NETWORK: 'network',
      VALIDATION: 'validation',
      SYSTEM: 'system',
      USER_INPUT: 'user_input',
      EXTERNAL: 'external',
      CONFIG: 'config',
      SECURITY: 'security'
    }
  },

  // キャッシュ・タイミング定数
  CACHE: {
    /** @type {Object} TTL設定（秒） */
    TTL: {
      SHORT: 300,      // 5分
      MEDIUM: 600,     // 10分  
      LONG: 3600,      // 1時間
      EXTENDED: 21600  // 6時間
    },
    
    /** @type {Object} バッチサイズ */
    BATCH_SIZE: {
      SMALL: 20,
      MEDIUM: 50,
      LARGE: 100,
      XLARGE: 200
    }
  },

  // タイムアウト定数（ミリ秒）
  TIMEOUTS: {
    SHORT: 1000,       // 1秒
    MEDIUM: 2000,      // 2秒
    LONG: 5000,        // 5秒
    POLLING: 30000,    // 30秒
    QUEUE: 60000,      // 1分
    EXECUTION_MAX: 300000  // 5分
  },

  // データ構造定数
  COLUMNS: {
    TIMESTAMP: 'タイムスタンプ',
    EMAIL: 'メールアドレス',
    CLASS: 'クラス',
    OPINION: '回答',
    REASON: '理由',
    NAME: '名前',
    UNDERSTAND: 'なるほど！',
    LIKE: 'いいね！',
    CURIOUS: 'もっと知りたい！',
    HIGHLIGHT: 'ハイライト'
  },

  // 表示モード
  DISPLAY: {
    MODES: {
      ANONYMOUS: 'anonymous',
      NAMED: 'named',
      EMAIL: 'email'
    }
  },

  // シート設定
  SHEETS: {
    /** @type {Object} ユーザーデータベースシート設定 */
    DATABASE: {
      NAME: 'Users',
      HEADERS: [
        'userId', 'adminEmail', 'spreadsheetId', 'spreadsheetUrl',
        'createdAt', 'configJson', 'lastAccessedAt', 'isActive'
      ]
    },
    
    /** @type {Object} 削除ログシート設定 */
    DELETE_LOG: {
      NAME: 'DeleteLogs',
      HEADERS: [
        'timestamp', 'executorEmail', 'targetUserId', 'targetEmail',
        'reason', 'deleteType'
      ]
    },
    
    /** @type {Object} 診断ログシート設定 */
    DIAGNOSTIC_LOG: {
      NAME: 'DiagnosticLogs',
      HEADERS: [
        'timestamp', 'functionName', 'status', 'problemCount',
        'repairCount', 'successfulRepairs', 'details', 'executor', 'summary'
      ]
    },

    /** @type {Object} 一般ログシート設定 */
    LOG: {
      NAME: 'Logs',
      HEADERS: ['timestamp', 'severity', 'category', 'message', 'userId', 'context']
    }
  },

  // フォームデフォルト値
  FORMS: {
    DEFAULT_MAIN_QUESTION: 'あなたの考えや気づいたことを教えてください',
    DEFAULT_REASON_QUESTION: 'そう考える理由や体験があれば教えてください（任意）',
    
    /** @type {Object} フォームプリセット */
    PRESETS: {
      OPINION_SURVEY: {
        title: '意見調査',
        description: 'ご意見をお聞かせください',
        mainQuestion: 'あなたの意見をお聞かせください',
        reasonQuestion: 'そう思う理由があれば教えてください'
      },
      REFLECTION: {
        title: '振り返り',
        description: '今回の学習について振り返ってみましょう',
        mainQuestion: '今回学んだことで印象に残ったことを教えてください',
        reasonQuestion: 'なぜそれが印象に残ったのですか？'
      },
      FEEDBACK: {
        title: 'フィードバック',
        description: 'フィードバックをお願いします',
        mainQuestion: 'ご感想やご意見をお聞かせください',
        reasonQuestion: '具体的な理由やエピソードがあれば教えてください'
      }
    }
  },

  // システム制限値
  LIMITS: {
    STRING_SHORT: 50,
    STRING_MEDIUM: 100,
    STRING_LONG: 500,
    STRING_MAX: 1000,
    HISTORY_ITEMS: 50,
    RETRY_COUNT: 3,
    MAX_CONCURRENT_FLOWS: 5
  },

  // スコアリング設定
  SCORING: {
    LIKE_MULTIPLIER: 0.1,
    RANDOM_FACTOR: 0.01,
    REACTION_KEYS: ['UNDERSTAND', 'LIKE', 'CURIOUS'],
    
    /** @type {Object} 重み設定 */
    WEIGHTS: {
      LIKE_WEIGHT: 1.0,
      LENGTH_WEIGHT: 0.1,
      RECENCY_WEIGHT: 0.05
    },
    
    /** @type {Object} ボーナス設定 */
    BONUSES: {
      MULTIPLE_REACTIONS: 0.2,
      DETAILED_REASON: 0.15
    }
  },

  // スクリプトプロパティキー
  SCRIPT_PROPS: {
    SERVICE_ACCOUNT_CREDS: 'SERVICE_ACCOUNT_CREDS',
    DATABASE_SPREADSHEET_ID: 'DATABASE_SPREADSHEET_ID',
    CLIENT_ID: 'CLIENT_ID',
    DEBUG_MODE: 'DEBUG_MODE'
  },

  // デバッグ設定
  DEBUG: {
    ENABLED: PropertiesService.getScriptProperties().getProperty('DEBUG_MODE') === 'true',
    LOG_LEVEL: 'INFO',
    TRACE_ENABLED: false
  },

  // 正規表現パターン
  REGEX: {
    EMAIL: /^[^\n@]+@[^\n@]+\.[^\n@]+$/,
    SPREADSHEET_ID: /^[a-zA-Z0-9-_]{44}$/,
    USER_ID: /^[a-f0-9]{32}$/
  }
};

// 読み取り専用にする
Object.freeze(UNIFIED_CONSTANTS);
Object.freeze(UNIFIED_CONSTANTS.ERROR);
Object.freeze(UNIFIED_CONSTANTS.ERROR.SEVERITY);
Object.freeze(UNIFIED_CONSTANTS.ERROR.CATEGORIES);
Object.freeze(UNIFIED_CONSTANTS.CACHE);
Object.freeze(UNIFIED_CONSTANTS.CACHE.TTL);
Object.freeze(UNIFIED_CONSTANTS.CACHE.BATCH_SIZE);
Object.freeze(UNIFIED_CONSTANTS.TIMEOUTS);
Object.freeze(UNIFIED_CONSTANTS.COLUMNS);
Object.freeze(UNIFIED_CONSTANTS.DISPLAY);
Object.freeze(UNIFIED_CONSTANTS.DISPLAY.MODES);
Object.freeze(UNIFIED_CONSTANTS.SHEETS);
Object.freeze(UNIFIED_CONSTANTS.SHEETS.DATABASE);
Object.freeze(UNIFIED_CONSTANTS.SHEETS.DELETE_LOG);
Object.freeze(UNIFIED_CONSTANTS.SHEETS.DIAGNOSTIC_LOG);
Object.freeze(UNIFIED_CONSTANTS.SHEETS.LOG);
Object.freeze(UNIFIED_CONSTANTS.FORMS);
Object.freeze(UNIFIED_CONSTANTS.FORMS.PRESETS);
Object.freeze(UNIFIED_CONSTANTS.LIMITS);
Object.freeze(UNIFIED_CONSTANTS.SCORING);
Object.freeze(UNIFIED_CONSTANTS.SCORING.WEIGHTS);
Object.freeze(UNIFIED_CONSTANTS.SCORING.BONUSES);
Object.freeze(UNIFIED_CONSTANTS.SCRIPT_PROPS);
Object.freeze(UNIFIED_CONSTANTS.DEBUG);
Object.freeze(UNIFIED_CONSTANTS.REGEX);

// 後方互換性のための別名エクスポート
/** @deprecated Use UNIFIED_CONSTANTS.ERROR.SEVERITY instead */
const ERROR_SEVERITY = UNIFIED_CONSTANTS.ERROR.SEVERITY;

/** @deprecated Use UNIFIED_CONSTANTS.ERROR.CATEGORIES instead */
const ERROR_CATEGORIES = UNIFIED_CONSTANTS.ERROR.CATEGORIES;

/** @deprecated Use UNIFIED_CONSTANTS.COLUMNS instead */
const COLUMN_HEADERS = UNIFIED_CONSTANTS.COLUMNS;

/** @deprecated Use UNIFIED_CONSTANTS.DISPLAY.MODES instead */
const DISPLAY_MODES = UNIFIED_CONSTANTS.DISPLAY.MODES;

/** @deprecated Use UNIFIED_CONSTANTS.SCORING.REACTION_KEYS instead */
const REACTION_KEYS = UNIFIED_CONSTANTS.SCORING.REACTION_KEYS;

/** @deprecated Use UNIFIED_CONSTANTS.CACHE.TTL.SHORT instead */
const USER_CACHE_TTL = UNIFIED_CONSTANTS.CACHE.TTL.SHORT;

/** @deprecated Use UNIFIED_CONSTANTS.TIMEOUTS.EXECUTION_MAX instead */
const EXECUTION_MAX_LIFETIME = UNIFIED_CONSTANTS.TIMEOUTS.EXECUTION_MAX;

/** @deprecated Use UNIFIED_CONSTANTS.CACHE.BATCH_SIZE.LARGE instead */
const DB_BATCH_SIZE = UNIFIED_CONSTANTS.CACHE.BATCH_SIZE.LARGE;

/** @deprecated Use UNIFIED_CONSTANTS.LIMITS.HISTORY_ITEMS instead */
const MAX_HISTORY_ITEMS = UNIFIED_CONSTANTS.LIMITS.HISTORY_ITEMS;

/** @deprecated Use UNIFIED_CONSTANTS.FORMS.DEFAULT_MAIN_QUESTION instead */
const DEFAULT_MAIN_QUESTION = UNIFIED_CONSTANTS.FORMS.DEFAULT_MAIN_QUESTION;

/** @deprecated Use UNIFIED_CONSTANTS.FORMS.DEFAULT_REASON_QUESTION instead */
const DEFAULT_REASON_QUESTION = UNIFIED_CONSTANTS.FORMS.DEFAULT_REASON_QUESTION;

/** @deprecated Use UNIFIED_CONSTANTS.SHEETS.DATABASE instead */
const DB_SHEET_CONFIG = {
  SHEET_NAME: UNIFIED_CONSTANTS.SHEETS.DATABASE.NAME,
  HEADERS: UNIFIED_CONSTANTS.SHEETS.DATABASE.HEADERS
};

/** @deprecated Use UNIFIED_CONSTANTS.SHEETS.DELETE_LOG instead */
const DELETE_LOG_SHEET_CONFIG = {
  SHEET_NAME: UNIFIED_CONSTANTS.SHEETS.DELETE_LOG.NAME,
  HEADERS: UNIFIED_CONSTANTS.SHEETS.DELETE_LOG.HEADERS
};

/** @deprecated Use UNIFIED_CONSTANTS.SHEETS.DIAGNOSTIC_LOG instead */
const DIAGNOSTIC_LOG_SHEET_CONFIG = {
  SHEET_NAME: UNIFIED_CONSTANTS.SHEETS.DIAGNOSTIC_LOG.NAME,
  HEADERS: UNIFIED_CONSTANTS.SHEETS.DIAGNOSTIC_LOG.HEADERS
};

/** @deprecated Use UNIFIED_CONSTANTS.SHEETS.LOG instead */
const LOG_SHEET_CONFIG = {
  SHEET_NAME: UNIFIED_CONSTANTS.SHEETS.LOG.NAME,
  HEADERS: UNIFIED_CONSTANTS.SHEETS.LOG.HEADERS
};

/** @deprecated Use UNIFIED_CONSTANTS.REGEX.EMAIL instead */
const EMAIL_REGEX = UNIFIED_CONSTANTS.REGEX.EMAIL;

/** @deprecated Use UNIFIED_CONSTANTS.DEBUG.ENABLED instead */
const DEBUG = UNIFIED_CONSTANTS.DEBUG.ENABLED;