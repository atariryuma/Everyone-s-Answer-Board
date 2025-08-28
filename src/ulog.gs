/**
 * 統合ログシステム (ULog)
 * 全GASファイルで統一されたログ記録を提供
 */

class ULog {
  
  /**
   * 汎用ログメソッド
   * @param {string} level - ログレベル (DEBUG, INFO, WARN, ERROR, CRITICAL)
   * @param {string} functionName - 呼び出し元の関数名
   * @param {string} message - ログメッセージ
   * @param {Object} data - 追加データ（オプション）
   * @param {string} category - ログカテゴリ（オプション）
   */
  static log(level, functionName, message, data = {}, category = null) {
    const levelValue = ULog.LEVELS[level] || ULog.LEVELS.INFO;
    
    // ログレベルフィルタリング
    if (levelValue < ULog.currentLevel) {
      return;
    }
    
    const timestamp = new Date().toISOString();
    const categoryTag = category ? `[${category}]` : '';
    const formattedMessage = `${ULog.prefix}${categoryTag} ${level} - ${functionName}: ${message}`;
    
    // データがある場合は追加情報として表示
    if (data && Object.keys(data).length > 0) {
      const dataString = typeof data === 'object' ? JSON.stringify(data, null, 2) : String(data);
      
      switch (level) {
        case 'DEBUG':
          console.log(`${formattedMessage}\n${dataString}`);
          break;
        case 'INFO':
          console.info(`${formattedMessage}\n${dataString}`);
          break;
        case 'WARN':
          console.warn(`${formattedMessage}\n${dataString}`);
          break;
        case 'ERROR':
        case 'CRITICAL':
          console.error(`${formattedMessage}\n${dataString}`);
          // 重要なエラーは Logger にも記録
          Logger.log(`${formattedMessage}\n${dataString}`);
          break;
      }
    } else {
      // データがない場合のシンプル出力
      switch (level) {
        case 'DEBUG':
          console.log(formattedMessage);
          break;
        case 'INFO':
          console.info(formattedMessage);
          break;
        case 'WARN':
          console.warn(formattedMessage);
          break;
        case 'ERROR':
        case 'CRITICAL':
          console.error(formattedMessage);
          Logger.log(formattedMessage);
          break;
      }
    }
  }
  
  /**
   * デバッグログ
   * @param {string} message - メッセージ
   * @param {Object} data - 追加データ（オプション）
   * @param {string} category - ログカテゴリ（オプション）
   */
  static debug(message, data = {}, category = null) {
    const caller = ULog._getCallerInfo();
    ULog.log('DEBUG', caller, message, data, category);
  }
  
  /**
   * 情報ログ
   * @param {string} message - メッセージ
   * @param {Object} data - 追加データ（オプション）
   * @param {string} category - ログカテゴリ（オプション）
   */
  static info(message, data = {}, category = null) {
    const caller = ULog._getCallerInfo();
    ULog.log('INFO', caller, message, data, category);
  }
  
  /**
   * 警告ログ
   * @param {string} message - メッセージ
   * @param {Object} data - 追加データ（オプション）
   * @param {string} category - ログカテゴリ（オプション）
   */
  static warn(message, data = {}, category = null) {
    const caller = ULog._getCallerInfo();
    ULog.log('WARN', caller, message, data, category);
  }
  
  /**
   * エラーログ
   * @param {string} message - メッセージ
   * @param {Object} data - 追加データ（オプション）
   * @param {string} category - ログカテゴリ（オプション）
   */
  static error(message, data = {}, category = null) {
    const caller = ULog._getCallerInfo();
    ULog.log('ERROR', caller, message, data, category);
  }
  
  /**
   * 重大エラーログ
   * @param {string} message - メッセージ
   * @param {Object} data - 追加データ（オプション）
   * @param {string} category - ログカテゴリ（オプション）
   */
  static critical(message, data = {}, category = null) {
    const caller = ULog._getCallerInfo();
    ULog.log('CRITICAL', caller, message, data, category);
  }
  
  /**
   * 呼び出し元関数名を取得（スタックトレースから推測）
   * @return {string} 関数名
   * @private
   */
  static _getCallerInfo() {
    try {
      const stack = new Error().stack;
      const stackLines = stack.split('\n');
      
      // スタックの3番目の行に呼び出し元があることが多い
      // 0: Error
      // 1: _getCallerInfo
      // 2: debug/info/warn/error/critical
      // 3: 実際の呼び出し元
      if (stackLines.length > 3) {
        const callerLine = stackLines[3];
        const functionMatch = callerLine.match(/at\s+(\w+)/);
        if (functionMatch) {
          return functionMatch[1];
        }
      }
      return 'unknown';
    } catch (e) {
      return 'unknown';
    }
  }
  
  /**
   * ログレベルを設定
   * @param {string} level - 新しいログレベル (DEBUG, INFO, WARN, ERROR, CRITICAL)
   */
  static setLevel(level) {
    if (ULog.LEVELS.hasOwnProperty(level)) {
      ULog.currentLevel = ULog.LEVELS[level];
      console.log(`${ULog.prefix} ログレベルを ${level} に設定しました`);
    } else {
      console.warn(`${ULog.prefix} 無効なログレベル: ${level}`);
    }
  }
  
  /**
   * プリフィックスを設定
   * @param {string} newPrefix - 新しいプリフィックス
   */
  static setPrefix(newPrefix) {
    ULog.prefix = newPrefix;
  }
  
  /**
   * パフォーマンス測定開始
   * @param {string} label - 測定ラベル
   * @return {Object} タイマーオブジェクト
   */
  static startTimer(label) {
    const startTime = new Date().getTime();
    return {
      label,
      startTime,
      end: function() {
        const endTime = new Date().getTime();
        const duration = endTime - startTime;
        ULog.info(`Performance: ${label} took ${duration}ms`, { duration }, ULog.CATEGORIES.PERFORMANCE);
        return duration;
      }
    };
  }
}

// ログレベル定数 (GAS互換性のため静的プロパティとして定義)
ULog.LEVELS = {
  DEBUG: 0,
  INFO: 1,
  WARN: 2,
  ERROR: 3,
  CRITICAL: 4
};

// ログカテゴリ定数
ULog.CATEGORIES = {
  AUTH: 'AUTH',
  API: 'API', 
  DB: 'DB',
  CACHE: 'CACHE',
  SECURITY: 'SECURITY',
  VALIDATION: 'VALIDATION',
  PERFORMANCE: 'PERFORMANCE',
  SYSTEM: 'SYSTEM',
  UI: 'UI',
  INTEGRATION: 'INTEGRATION'
};

// 現在のログレベル (本番環境では INFO 以上)
ULog.currentLevel = ULog.LEVELS.DEBUG;

// プリフィックス設定
ULog.prefix = '[StudyQuest]';

// グローバル関数としてもエクスポート（レガシーサポート用）
function ulogFunction(level, functionName, message, data, category) {
  return ULog.log(level, functionName, message, data, category);
}