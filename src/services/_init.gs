/**
 * @fileoverview 初期化ファイル - サービスの読み込み順序と依存関係を管理
 * 
 * このファイルは最初に読み込まれ、既存のコードとの競合を防ぎます。
 */

/**
 * グローバル名前空間の保護
 * 既存の関数との競合を防ぐ
 */
const _ORIGINAL_FUNCTIONS = {};

// 既存の関数を保存
if (typeof resilientExecutor !== 'undefined') {
  _ORIGINAL_FUNCTIONS.resilientExecutor = resilientExecutor;
}

if (typeof logError !== 'undefined') {
  _ORIGINAL_FUNCTIONS.logError = logError;
}

if (typeof debugLog !== 'undefined') {
  _ORIGINAL_FUNCTIONS.debugLog = debugLog;
}

if (typeof cacheManager !== 'undefined') {
  _ORIGINAL_FUNCTIONS.cacheManager = cacheManager;
}

/**
 * 新アーキテクチャの有効化フラグ
 * @returns {boolean} 新アーキテクチャが有効か
 */
function isNewArchitectureEnabled() {
  try {
    const props = PropertiesService.getScriptProperties();
    const flag = props.getProperty('USE_NEW_ARCHITECTURE');
    return flag === 'true';
  } catch (error) {
    // デフォルトは無効（安全側に倒す）
    return false;
  }
}

/**
 * サービスの初期化
 * 新アーキテクチャが有効な場合のみ実行
 */
function initializeServices() {
  if (!isNewArchitectureEnabled()) {
    console.log('📦 Legacy architecture is active');
    return false;
  }
  
  console.log('🚀 New architecture is active');
  
  // 新サービスの初期化フラグを設定
  global._NEW_SERVICES_INITIALIZED = true;
  
  return true;
}

/**
 * 既存関数の復元
 * 必要に応じて既存の関数を復元
 */
function restoreOriginalFunctions() {
  Object.keys(_ORIGINAL_FUNCTIONS).forEach(funcName => {
    if (_ORIGINAL_FUNCTIONS[funcName]) {
      global[funcName] = _ORIGINAL_FUNCTIONS[funcName];
    }
  });
}

/**
 * デバッグ用：現在の状態を出力
 */
function debugArchitectureStatus() {
  const status = {
    newArchitectureEnabled: isNewArchitectureEnabled(),
    servicesInitialized: global._NEW_SERVICES_INITIALIZED || false,
    originalFunctionsSaved: Object.keys(_ORIGINAL_FUNCTIONS).length,
    timestamp: new Date().toISOString()
  };
  
  console.log('Architecture Status:', JSON.stringify(status, null, 2));
  return status;
}

// 初期化を実行
const _INIT_RESULT = initializeServices();