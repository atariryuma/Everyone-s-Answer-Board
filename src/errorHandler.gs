/**
 * @fileoverview エラーハンドリング最適化 - ユーザーフレンドリーなエラーメッセージ
 */

/**
 * エラー分類とメッセージマッピング
 */
const ERROR_CATEGORIES = {
  AUTHENTICATION: {
    pattern: /(authentication|unauthorized|権限|認証)/i,
    userMessage: 'アクセス権限に問題があります。ログイン状態を確認してください。',
    technical: true
  },
  
  NETWORK: {
    pattern: /(timeout|network|connection|接続|タイムアウト)/i,
    userMessage: 'ネットワーク接続に問題があります。しばらく時間をおいて再試行してください。',
    technical: false
  },
  
  QUOTA_EXCEEDED: {
    pattern: /(quota|limit|exceeded|制限|上限)/i,
    userMessage: 'APIの使用制限に達しています。しばらく時間をおいて再試行してください。',
    technical: false
  },
  
  PERMISSION: {
    pattern: /(permission|forbidden|アクセス|許可)/i,
    userMessage: 'スプレッドシートへのアクセス権限がありません。共有設定を確認してください。',
    technical: true
  },
  
  DATA_NOT_FOUND: {
    pattern: /(not found|見つかりません|存在しません)/i,
    userMessage: '対象のデータが見つかりません。設定を確認してください。',
    technical: false
  },
  
  VALIDATION: {
    pattern: /(validation|invalid|不正|無効)/i,
    userMessage: '入力内容に問題があります。値を確認して再試行してください。',
    technical: false
  },
  
  LOCK_TIMEOUT: {
    pattern: /(lock|ロック|競合|同時)/i,
    userMessage: '他の処理と競合しています。少し待ってから再試行してください。',
    technical: false
  },
  
  SERVICE_UNAVAILABLE: {
    pattern: /(service unavailable|サービス|利用できません)/i,
    userMessage: 'サービスが一時的に利用できません。しばらく時間をおいて再試行してください。',
    technical: false
  }
};

/**
 * エラーを分析してユーザーフレンドリーなメッセージを生成
 * @param {Error|string} error - エラーオブジェクトまたはメッセージ
 * @param {object} context - エラーコンテキスト
 * @returns {object} 整形されたエラー情報
 */
function processError(error, context = {}) {
  const errorMessage = typeof error === 'string' ? error : error.message || 'Unknown error';
  const errorStack = error.stack || '';
  
  // エラーカテゴリを判定
  let category = null;
  let userMessage = '予期しないエラーが発生しました。';
  
  for (const [key, config] of Object.entries(ERROR_CATEGORIES)) {
    if (config.pattern.test(errorMessage)) {
      category = key;
      userMessage = config.userMessage;
      break;
    }
  }
  
  // エラーレベルを判定
  const level = determineErrorLevel(errorMessage, category);
  
  // ユーザー向けアクションを提案
  const suggestedActions = getSuggestedActions(category, context);
  
  // 技術的詳細を含めるかどうか
  const includeTechnicalDetails = category && ERROR_CATEGORIES[category].technical;
  
  const processedError = {
    userMessage: userMessage,
    category: category || 'UNKNOWN',
    level: level,
    suggestedActions: suggestedActions,
    timestamp: new Date().toISOString(),
    context: context
  };
  
  // 技術的詳細（開発者向け）
  if (includeTechnicalDetails && isDeployUser()) {
    processedError.technicalDetails = {
      originalMessage: errorMessage,
      stack: errorStack,
      function: context.function || 'unknown'
    };
  }
  
  // ログ出力
  logError(processedError, errorMessage);
  
  return processedError;
}

/**
 * エラーレベルを判定
 * @param {string} errorMessage - エラーメッセージ
 * @param {string} category - エラーカテゴリ
 * @returns {string} エラーレベル
 */
function determineErrorLevel(errorMessage, category) {
  if (category === 'AUTHENTICATION' || category === 'PERMISSION') {
    return 'HIGH';
  }
  
  if (category === 'QUOTA_EXCEEDED' || category === 'SERVICE_UNAVAILABLE') {
    return 'MEDIUM';
  }
  
  if (errorMessage.includes('致命的') || errorMessage.includes('critical')) {
    return 'CRITICAL';
  }
  
  return 'LOW';
}

/**
 * 推奨アクションを取得
 * @param {string} category - エラーカテゴリ
 * @param {object} context - エラーコンテキスト
 * @returns {array} 推奨アクション配列
 */
function getSuggestedActions(category, context) {
  const actions = [];
  
  switch (category) {
    case 'AUTHENTICATION':
      actions.push('ログアウトして再ログインしてください');
      actions.push('ブラウザのキャッシュをクリアしてください');
      break;
      
    case 'PERMISSION':
      actions.push('スプレッドシートの共有設定を確認してください');
      if (context.serviceAccount) {
        actions.push(`サービスアカウント (${context.serviceAccount}) を編集者として追加してください`);
      }
      break;
      
    case 'NETWORK':
      actions.push('インターネット接続を確認してください');
      actions.push('少し時間をおいて再試行してください');
      break;
      
    case 'QUOTA_EXCEEDED':
      actions.push('1分程度待ってから再試行してください');
      actions.push('同時実行している他の操作があれば終了してください');
      break;
      
    case 'DATA_NOT_FOUND':
      actions.push('スプレッドシートURLが正しいか確認してください');
      actions.push('シート名や設定を確認してください');
      break;
      
    case 'LOCK_TIMEOUT':
      actions.push('30秒程度待ってから再試行してください');
      actions.push('他のユーザーが同じ機能を使用していないか確認してください');
      break;
      
    default:
      actions.push('ページを再読み込みして再試行してください');
      actions.push('問題が続く場合は管理者にお問い合わせください');
  }
  
  return actions;
}

/**
 * エラーをログに記録
 * @param {object} processedError - 処理済みエラー情報
 * @param {string} originalMessage - 元のエラーメッセージ
 */
function logError(processedError, originalMessage) {
  const logLevel = processedError.level;
  const logMessage = `[${logLevel}] ${processedError.category}: ${processedError.userMessage}`;
  
  switch (logLevel) {
    case 'CRITICAL':
      console.error('🚨', logMessage);
      console.error('Original:', originalMessage);
      break;
    case 'HIGH':
      console.error('❌', logMessage);
      break;
    case 'MEDIUM':
      console.warn('⚠️', logMessage);
      break;
    default:
      console.log('ℹ️', logMessage);
  }
  
  // 高レベルエラーは詳細ログも出力
  if (logLevel === 'CRITICAL' || logLevel === 'HIGH') {
    console.log('Error Context:', JSON.stringify(processedError.context, null, 2));
  }
}

/**
 * 安全なエラー実行ラッパー
 * @param {function} fn - 実行する関数
 * @param {object} context - エラーコンテキスト
 * @returns {any} 関数の実行結果
 */
function safeExecute(fn, context = {}) {
  try {
    return fn();
  } catch (error) {
    const processedError = processError(error, context);
    
    // ユーザーフレンドリーなエラーを再スロー
    const userError = new Error(processedError.userMessage);
    userError.category = processedError.category;
    userError.level = processedError.level;
    userError.suggestedActions = processedError.suggestedActions;
    userError.processedError = processedError;
    
    throw userError;
  }
}

/**
 * 非同期処理用の安全なエラーハンドリング
 * @param {function} asyncFn - 非同期関数
 * @param {object} context - エラーコンテキスト
 * @returns {Promise} Promise結果
 */
async function safeExecuteAsync(asyncFn, context = {}) {
  try {
    return await asyncFn();
  } catch (error) {
    const processedError = processError(error, context);
    
    const userError = new Error(processedError.userMessage);
    userError.category = processedError.category;
    userError.level = processedError.level;
    userError.suggestedActions = processedError.suggestedActions;
    userError.processedError = processedError;
    
    throw userError;
  }
}

/**
 * Sheets API エラー専用ハンドラー
 * @param {Error} error - Sheets API エラー
 * @param {object} context - APIコンテキスト
 * @returns {object} 処理済みエラー情報
 */
function handleSheetsApiError(error, context = {}) {
  const errorMessage = error.message || error.toString();
  
  // Sheets API 特有のエラーパターンをチェック
  if (errorMessage.includes('403')) {
    context.suggestedAction = 'スプレッドシートの共有設定を確認してください';
    return processError('スプレッドシートへのアクセス権限がありません', context);
  }
  
  if (errorMessage.includes('404')) {
    context.suggestedAction = 'スプレッドシートURLまたはシート名を確認してください';
    return processError('指定されたスプレッドシートまたはシートが見つかりません', context);
  }
  
  if (errorMessage.includes('429')) {
    context.suggestedAction = 'APIの使用制限に達しました。しばらく待って再試行してください';
    return processError('API使用制限に達しました', context);
  }
  
  if (errorMessage.includes('500') || errorMessage.includes('502') || errorMessage.includes('503')) {
    context.suggestedAction = 'Google のサービスが一時的に利用できません';
    return processError('Google Sheets サービスが一時的に利用できません', context);
  }
  
  // その他のAPIエラー
  return processError(error, context);
}

/**
 * ユーザー向けエラー表示用のHTMLフォーマット
 * @param {object} processedError - 処理済みエラー情報
 * @returns {string} HTML形式のエラーメッセージ
 */
function formatErrorForDisplay(processedError) {
  let html = `<div class="error-container alert alert-${processedError.level.toLowerCase()}">`;
  html += `<h4 class="error-title">⚠️ ${processedError.userMessage}</h4>`;
  
  if (processedError.suggestedActions && processedError.suggestedActions.length > 0) {
    html += '<div class="error-actions"><p><strong>対処方法:</strong></p><ul>';
    processedError.suggestedActions.forEach(action => {
      html += `<li>${action}</li>`;
    });
    html += '</ul></div>';
  }
  
  if (processedError.technicalDetails && isDeployUser()) {
    html += '<details class="error-technical"><summary>技術的詳細 (管理者用)</summary>';
    html += `<pre>${JSON.stringify(processedError.technicalDetails, null, 2)}</pre>`;
    html += '</details>';
  }
  
  html += '</div>';
  return html;
}

/**
 * エラー回復の試行
 * @param {function} fn - 再試行する関数
 * @param {object} options - リトライオプション
 * @returns {any} 関数の実行結果
 */
function retryWithBackoff(fn, options = {}) {
  const maxRetries = options.maxRetries || 3;
  const baseDelay = options.baseDelay || 1000;
  const maxDelay = options.maxDelay || 10000;
  
  let lastError;
  
  for (let attempt = 0; attempt < maxRetries; attempt++) {
    try {
      return fn();
    } catch (error) {
      lastError = error;
      
      // リトライしないエラーカテゴリ
      if (error.category === 'AUTHENTICATION' || error.category === 'PERMISSION') {
        throw error;
      }
      
      if (attempt === maxRetries - 1) {
        // 最後の試行で失敗
        throw error;
      }
      
      // 指数バックオフ
      const delay = Math.min(baseDelay * Math.pow(2, attempt), maxDelay);
      console.log(`リトライ ${attempt + 1}/${maxRetries} (${delay}ms待機)...`);
      Utilities.sleep(delay);
    }
  }
  
  throw lastError;
}