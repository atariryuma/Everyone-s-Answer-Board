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

// 後方互換性のための別名エクスポート（直接値定義で循環参照を回避）
/** @deprecated Use UNIFIED_CONSTANTS.ERROR.SEVERITY instead */
const ERROR_SEVERITY = {
  LOW: 'low',
  MEDIUM: 'medium',
  HIGH: 'high', 
  CRITICAL: 'critical'
};

/** @deprecated Use UNIFIED_CONSTANTS.ERROR.CATEGORIES instead */
const ERROR_CATEGORIES = {
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
};

/** @deprecated Use UNIFIED_CONSTANTS.COLUMNS instead */
const COLUMN_HEADERS = {
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
};

/** @deprecated Use UNIFIED_CONSTANTS.DISPLAY.MODES instead */
const DISPLAY_MODES = {
  ANONYMOUS: 'anonymous',
  NAMED: 'named',
  EMAIL: 'email'
};

/** @deprecated Use UNIFIED_CONSTANTS.SCORING.REACTION_KEYS instead */
const REACTION_KEYS = ['UNDERSTAND', 'LIKE', 'CURIOUS'];

/** @deprecated Use UNIFIED_CONSTANTS.CACHE.TTL.SHORT instead */
const USER_CACHE_TTL = 300;

/** @deprecated Use UNIFIED_CONSTANTS.TIMEOUTS.EXECUTION_MAX instead */
const EXECUTION_MAX_LIFETIME = 300000;

/** @deprecated Use UNIFIED_CONSTANTS.CACHE.BATCH_SIZE.LARGE instead */
const DB_BATCH_SIZE = 100;

/** @deprecated Use UNIFIED_CONSTANTS.LIMITS.HISTORY_ITEMS instead */
const MAX_HISTORY_ITEMS = 50;

/** @deprecated Use UNIFIED_CONSTANTS.FORMS.DEFAULT_MAIN_QUESTION instead */
const DEFAULT_MAIN_QUESTION = 'あなたの考えや気づいたことを教えてください';

/** @deprecated Use UNIFIED_CONSTANTS.FORMS.DEFAULT_REASON_QUESTION instead */
const DEFAULT_REASON_QUESTION = 'そう考える理由や体験があれば教えてください（任意）';

/** @deprecated Use UNIFIED_CONSTANTS.SHEETS.DATABASE instead */
const DB_SHEET_CONFIG = {
  SHEET_NAME: 'Users',
  HEADERS: ['userId', 'adminEmail', 'spreadsheetId', 'spreadsheetUrl', 'createdAt', 'configJson', 'lastAccessedAt', 'isActive']
};

/** @deprecated Use UNIFIED_CONSTANTS.SHEETS.DELETE_LOG instead */
const DELETE_LOG_SHEET_CONFIG = {
  SHEET_NAME: 'DeleteLogs',
  HEADERS: ['timestamp', 'executorEmail', 'targetUserId', 'targetEmail', 'reason', 'deleteType']
};

/** @deprecated Use UNIFIED_CONSTANTS.SHEETS.DIAGNOSTIC_LOG instead */
const DIAGNOSTIC_LOG_SHEET_CONFIG = {
  SHEET_NAME: 'DiagnosticLogs',
  HEADERS: ['timestamp', 'functionName', 'status', 'problemCount', 'repairCount', 'successfulRepairs', 'details', 'executor', 'summary']
};

/** @deprecated Use UNIFIED_CONSTANTS.SHEETS.LOG instead */
const LOG_SHEET_CONFIG = {
  SHEET_NAME: 'Logs',
  HEADERS: ['timestamp', 'severity', 'category', 'message', 'userId', 'context']
};

/** @deprecated Use UNIFIED_CONSTANTS.REGEX.EMAIL instead */
const EMAIL_REGEX = /^[^\n@]+@[^\n@]+\.[^\n@]+$/;

/** @deprecated Use UNIFIED_CONSTANTS.DEBUG.ENABLED instead */
const DEBUG = PropertiesService.getScriptProperties().getProperty('DEBUG_MODE') === 'true';

// =============================================================================
// 実行時定数アクセス関数（フォールバック機能付き）
// =============================================================================

/**
 * 実行時に定数を安全に取得する関数
 * @param {string} path - 定数パス（例: 'ERROR.SEVERITY.HIGH'）
 * @param {*} fallback - フォールバック値
 * @returns {*} 定数値またはフォールバック値
 */
function getConstant(path, fallback = null) {
  try {
    if (typeof UNIFIED_CONSTANTS === 'undefined') {
      return fallback;
    }
    
    const parts = path.split('.');
    let value = UNIFIED_CONSTANTS;
    
    for (const part of parts) {
      if (value && typeof value === 'object' && part in value) {
        value = value[part];
      } else {
        return fallback;
      }
    }
    
    return value !== undefined ? value : fallback;
  } catch (error) {
    return fallback;
  }
}

/**
 * エラー重大度を安全に取得
 * @param {string} level - レベル（LOW, MEDIUM, HIGH, CRITICAL）
 * @returns {string} エラーレベル文字列
 */
function getErrorSeverity(level) {
  const mapping = {
    LOW: 'low',
    MEDIUM: 'medium',
    HIGH: 'high',
    CRITICAL: 'critical'
  };
  return mapping[level] || 'medium';
}

/**
 * エラーカテゴリを安全に取得
 * @param {string} category - カテゴリ名
 * @returns {string} エラーカテゴリ文字列
 */
function getErrorCategory(category) {
  const mapping = {
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
  };
  return mapping[category] || 'system';
}

/**
 * キャッシュTTLを安全に取得
 * @param {string} duration - 期間（SHORT, MEDIUM, LONG, EXTENDED）
 * @returns {number} TTL秒数
 */
function getCacheTTL(duration) {
  const mapping = {
    SHORT: 300,      // 5分
    MEDIUM: 600,     // 10分
    LONG: 3600,      // 1時間
    EXTENDED: 21600  // 6時間
  };
  return mapping[duration] || 300;
}

/**
 * タイムアウト値を安全に取得
 * @param {string} type - タイムアウトタイプ
 * @returns {number} タイムアウト値（ミリ秒）
 */
function getTimeout(type) {
  const mapping = {
    SHORT: 1000,       // 1秒
    MEDIUM: 2000,      // 2秒
    LONG: 5000,        // 5秒
    POLLING: 30000,    // 30秒
    QUEUE: 60000,      // 1分
    EXECUTION_MAX: 300000  // 5分
  };
  return mapping[type] || 1000;
}

/**
 * 実行時間制限を安全に取得
 * @returns {number} 実行時間制限（ミリ秒）
 */
function getExecutionMaxLifetime() {
  return getTimeout('EXECUTION_MAX');
}

/**
 * ユーザーキャッシュTTLを安全に取得
 * @returns {number} TTL秒数
 */
function getUserCacheTTL() {
  return getCacheTTL('SHORT');
}

/**
 * データベースバッチサイズを安全に取得
 * @returns {number} バッチサイズ
 */
function getDbBatchSize() {
  return getConstant('CACHE.BATCH_SIZE.LARGE', 100);
}

/**
 * デフォルト質問文を安全に取得
 * @param {string} type - 質問タイプ（MAIN, REASON）
 * @returns {string} 質問文
 */
function getDefaultQuestion(type) {
  const mapping = {
    MAIN: 'あなたの考えや気づいたことを教えてください',
    REASON: 'そう考える理由や体験があれば教えてください（任意）'
  };
  return mapping[type] || mapping.MAIN;
}

// 実行時定数アクセス関数ログ
console.log('✅ 実行時定数アクセス関数が利用可能です');