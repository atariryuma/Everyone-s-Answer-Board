/**
 * 統合ログシステム (ULog) - 最適化版
 * 全GASファイルで統一されたログ記録を提供
 * パフォーマンス重視で設計された軽量ログシステム
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
    
    // ログレベルフィルタリング（パフォーマンス最適化）
    if (levelValue < ULog.currentLevel) {
      return;
    }
    
    const categoryTag = category ? `[${category}]` : '';
    const formattedMessage = `${ULog.prefix}${categoryTag} ${level} - ${functionName}: ${message}`;
    
    // データがある場合の処理を効率化
    const hasData = data && Object.keys(data).length > 0;
    const finalMessage = hasData 
      ? `${formattedMessage}\n${typeof data === 'object' ? JSON.stringify(data, null, 2) : String(data)}`
      : formattedMessage;
    
    // レベル別出力（最適化済み）
    switch (level) {
      case 'DEBUG':
        console.log(finalMessage);
        break;
      case 'INFO':
        console.log(finalMessage);
        break;
      case 'WARN':
        console.warn(finalMessage);
        break;
      case 'ERROR':
      case 'CRITICAL':
        console.error(finalMessage);
        // 重要なエラーのみLogger記録（パフォーマンス考慮）
        if (level === 'CRITICAL') {
          Logger.log(finalMessage);
        }
        break;
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
   * 呼び出し元関数名を取得（パフォーマンス最適化版）
   * @return {string} 関数名
   * @private
   */
  static _getCallerInfo() {
    try {
      const stack = new Error().stack;
      const stackLines = stack.split('\n');
      
      // 最適化：スタックの3-4番目の行を効率的に検索
      for (let i = 3; i <= 4 && i < stackLines.length; i++) {
        const line = stackLines[i];
        const functionMatch = line.match(/at\s+(\w+)/);
        if (functionMatch && functionMatch[1] !== '_getCallerInfo') {
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
    const startTime = Date.now();
    return {
      label: label,
      startTime: startTime,
      end: () => {
        const duration = Date.now() - startTime;
        console.log(`Performance: ${label}`, { duration: `${duration}ms` });
        return duration;
      }
    };
  }
  
  /**
   * 本番モード設定（パフォーマンス最適化）
   */
  static setProductionMode() {
    ULog.setLevel('ERROR');
    ULog.setPrefix('[App]');
    console.log('ULog: 本番モード設定完了（ERROR以上のみ出力）');
  }
  
  /**
   * 開発モード設定
   */
  static setDevelopmentMode() {
    ULog.setLevel('DEBUG');
    ULog.setPrefix('[Dev]');
    console.log('ULog: 開発モード設定完了（全レベル出力）');
  }
}

// ログレベル定数
ULog.LEVELS = {
  DEBUG: 0,
  INFO: 1,
  WARN: 2,
  ERROR: 3,
  CRITICAL: 4
};

// ログカテゴリ定数（最適化済み）
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

// デフォルト設定（本番最適化）
ULog.currentLevel = ULog.LEVELS.ERROR; // 本番では ERROR 以上のみ
ULog.prefix = '[EAB]'; // Everyone's Answer Board

// 環境に応じた自動設定
if (typeof Session !== 'undefined') {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    // 開発者メール（必要に応じて調整）
    if (userEmail && userEmail.includes('ryuma')) {
      ULog.setDevelopmentMode();
    }
  } catch (e) {
    // 本番環境ではSession取得エラーが発生する可能性があるため無視
  }
}

// グローバル関数としてもエクスポート（レガシーサポート用）
function ulogFunction(level, functionName, message, data, category) {
  return ULog.log(level, functionName, message, data, category);
}

// 後方互換性のための簡易アクセサ
const Log = ULog;