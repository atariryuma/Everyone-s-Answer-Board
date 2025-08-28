/**
 * デバッグ設定管理システム
 * 本番環境での debugLog 制御とロギングレベル管理
 */

// デバッグ設定 - 本番環境では false に設定
const DEBUG_CONFIG = {
  // 本番環境判定: PropertiesService で制御可能
  isProduction: false, // デプロイ時に true に変更

  // ロギングレベル設定
  logLevels: {
    ERROR: 0, // エラーログ（常に出力）
    WARN: 1, // 警告ログ
    INFO: 2, // 情報ログ
    DEBUG: 3, // デバッグログ
  },

  // 現在のログレベル（本番では ERROR のみ）
  currentLogLevel: 2, // 開発時は INFO、本番では 0

  // カテゴリ別デバッグ制御
  categories: {
    CACHE: true, // キャッシュ関連
    AUTH: true, // 認証関連
    DATABASE: true, // データベース操作
    UI: false, // UI更新関連（頻繁なため通常は無効）
    PERFORMANCE: true, // パフォーマンス計測
  },
};

/**
 * 本番環境かどうかを判定
 * @returns {boolean} 本番環境の場合 true
 */
function isProductionEnvironment() {
  try {
    // PropertiesService での制御を優先
    const envSetting = PropertiesService.getScriptProperties().getProperty('ENVIRONMENT');
    if (envSetting) {
      return envSetting === 'production';
    }

    // fallback: DEBUG_CONFIG の設定を使用
    return DEBUG_CONFIG.isProduction;
  } catch (error) {
    // エラー時は安全側に倒して本番扱い
    return true;
  }
}

/**
 * ログレベルに応じた出力制御
 * @param {string} level - ログレベル (ERROR, WARN, INFO, DEBUG)
 * @param {string} category - カテゴリ (CACHE, AUTH, DATABASE, UI, PERFORMANCE)
 * @param {string} message - ログメッセージ
 * @param {...any} args - 追加引数
 */
function controlledLog(level, category, message, ...args) {
  // DEBUG_MODE設定をチェック（AppSetupPageでの設定を優先）
  try {
    const debugModeEnabled =
      PropertiesService.getScriptProperties().getProperty('DEBUG_MODE') === 'true';

    // DEBUG_MODEが無効の場合はERRORレベルのみ出力
    if (!debugModeEnabled && level !== 'ERROR') {
      return;
    }
  } catch (error) {
    // PropertiesService エラー時は従来の本番環境判定にフォールバック
    if (isProductionEnvironment() && level !== 'ERROR') {
      return;
    }
  }

  // ログレベルチェック（DEBUG_MODEが有効な場合のみ適用）
  const levelValue = DEBUG_CONFIG.logLevels[level] || DEBUG_CONFIG.logLevels.DEBUG;
  if (levelValue > DEBUG_CONFIG.currentLogLevel) {
    return;
  }

  // カテゴリ別制御チェック（DEBUG_MODEが有効な場合のみ適用）
  if (category && DEBUG_CONFIG.categories[category] === false) {
    return;
  }

  // ログ出力
  const timestamp = new Date().toISOString();
  const formattedMessage = `[${timestamp}] ${level}${category ? `[${category}]` : ''}: ${message}`;
  const logData = args.length > 0 ? { additionalData: args } : {};
  
  // ULogのカテゴリにマッピング
  const ulogCategory = category ? ULog.CATEGORIES[category] || ULog.CATEGORIES.SYSTEM : ULog.CATEGORIES.SYSTEM;

  switch (level) {
    case 'ERROR':
      ULog.error(formattedMessage, logData, ulogCategory);
      break;
    case 'WARN':
      ULog.warn(formattedMessage, logData, ulogCategory);
      break;
    case 'INFO':
      ULog.info(formattedMessage, logData, ulogCategory);
      break;
    default:
      ULog.debug(formattedMessage, logData, ulogCategory);
  }
}

/**
 * debugLog 関数の改良版 - 本番環境制御対応
 * @param {string} message - ログメッセージ
 * @param {...any} args - 追加引数
 */
function debugLog(message, ...args) {
  // カテゴリを自動判定
  let category = null;
  if (message.includes('キャッシュ') || message.includes('cache') || message.includes('Cache')) {
    category = 'CACHE';
  } else if (message.includes('認証') || message.includes('auth') || message.includes('Auth')) {
    category = 'AUTH';
  } else if (
    message.includes('データベース') ||
    message.includes('database') ||
    message.includes('DB')
  ) {
    category = 'DATABASE';
  } else if (message.includes('UI') || message.includes('表示') || message.includes('更新')) {
    category = 'UI';
  } else if (
    message.includes('パフォーマンス') ||
    message.includes('performance') ||
    message.includes('時間')
  ) {
    category = 'PERFORMANCE';
  }

  controlledLog('DEBUG', category, message, ...args);
}

// errorLog統合: Core.gsのlogErrorに統一。この関数は削除し、Core.gsを使用

// warnLog - 重複削除済み（Core.gsの実装を使用）

// infoLog - 重複削除済み（Core.gsの実装を使用）

/**
 * パフォーマンス計測用のタイマー
 */
