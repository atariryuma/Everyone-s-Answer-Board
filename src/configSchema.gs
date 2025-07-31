/**
 * @fileoverview ConfigJSON 統一スキーマ（簡素版）
 * シンプルで実用的なConfigJSON処理システム
 */

/**
 * ConfigJSON のデフォルト値とスキーマ定義
 */
const CONFIG_SCHEMA = {
  // 必須フィールド
  setupStatus: { type: 'string', default: 'pending', values: ['pending', 'completed', 'error', 'reconfiguring'] },
  formCreated: { type: 'boolean', default: false },
  appPublished: { type: 'boolean', default: false },
  publishedSheetName: { type: 'string', default: '' },
  publishedSpreadsheetId: { type: 'string', default: '' },
  createdAt: { type: 'string', default: null },
  lastModified: { type: 'string', default: null },
  
  // オプショナルフィールド
  formUrl: { type: 'string', default: '' },
  editFormUrl: { type: 'string', default: '' },
  displayMode: { type: 'string', default: 'anonymous', values: ['anonymous', 'named'] },
  showCounts: { type: 'boolean', default: false },
  showNames: { type: 'boolean', default: false },
  sortOrder: { type: 'string', default: 'newest', values: ['newest', 'oldest'] },
  version: { type: 'string', default: '1.0.0' },
  
  // 自動管理フィールド
  autoStopEnabled: { type: 'boolean', default: false },
  autoStopMinutes: { type: 'number', default: 360 },
  totalPublishCount: { type: 'number', default: 0 }
};

/**
 * 統一ConfigJSON取得関数（全システム共通エントリーポイント）
 * @param {Object} userInfo - ユーザー情報オブジェクト
 * @returns {Object} 正規化されたConfigJSON
 */
function getConfigJSON(userInfo) {
  if (!userInfo || !userInfo.configJson) {
    return createDefaultConfigJSON();
  }
  
  let parsed = {};
  try {
    parsed = JSON.parse(userInfo.configJson);
  } catch (error) {
    console.warn('ConfigJSON解析エラー:', error.message);
    return createDefaultConfigJSON();
  }
  
  return normalizeConfigJSON(parsed);
}

/**
 * デフォルトConfigJSONを作成
 * @returns {Object} デフォルト設定
 */
function createDefaultConfigJSON() {
  const timestamp = new Date().toISOString();
  const config = {};
  
  // スキーマからデフォルト値を設定
  for (const [field, schema] of Object.entries(CONFIG_SCHEMA)) {
    config[field] = schema.default;
  }
  
  config.createdAt = timestamp;
  config.lastModified = timestamp;
  
  return config;
}

/**
 * ConfigJSONの正規化処理
 * @param {Object} config - 入力設定
 * @returns {Object} 正規化された設定
 */
function normalizeConfigJSON(config) {
  if (!config || typeof config !== 'object') {
    return createDefaultConfigJSON();
  }
  
  const normalized = {};
  const timestamp = new Date().toISOString();
  
  // スキーマ定義フィールドの処理
  for (const [field, schema] of Object.entries(CONFIG_SCHEMA)) {
    if (config[field] !== undefined) {
      normalized[field] = validateAndConvertValue(config[field], schema);
    } else {
      normalized[field] = schema.default;
    }
  }
  
  // シート固有設定の処理
  for (const [key, value] of Object.entries(config)) {
    if (key.startsWith('sheet_') && typeof value === 'object') {
      normalized[key] = normalizeSheetConfig(value);
    }
  }
  
  // タイムスタンプの更新
  if (!normalized.createdAt) {
    normalized.createdAt = timestamp;
  }
  normalized.lastModified = timestamp;
  
  return normalized;
}

/**
 * 値の検証と変換
 * @param {*} value - 値
 * @param {Object} schema - スキーマ定義
 * @returns {*} 変換された値
 */
function validateAndConvertValue(value, schema) {
  // 型変換
  switch (schema.type) {
    case 'boolean':
      return typeof value === 'string' ? value.toLowerCase() === 'true' : Boolean(value);
    case 'number':
      const num = Number(value);
      return isNaN(num) ? schema.default : num;
    case 'string':
      return String(value);
    default:
      return value;
  }
}

/**
 * シート固有設定の正規化
 * @param {Object} sheetConfig - シート設定
 * @returns {Object} 正規化されたシート設定
 */
function normalizeSheetConfig(sheetConfig) {
  const defaults = {
    opinionHeader: '',
    nameHeader: '',
    classHeader: '',
    reasonHeader: '',
    formCreated: false,
    setupStatus: 'pending',
    isConfigured: false,
    createdAt: '',
    lastModified: new Date().toISOString()
  };
  
  return { ...defaults, ...sheetConfig };
}

/**
 * ConfigJSONの基本検証
 * @param {Object} config - 検証対象
 * @returns {Object} 検証結果
 */
function validateConfigJSON(config) {
  const errors = [];
  
  if (!config || typeof config !== 'object') {
    return { isValid: false, errors: ['ConfigJSONが無効です'] };
  }
  
  // 必須フィールドチェック
  const requiredFields = ['setupStatus', 'formCreated', 'appPublished', 'publishedSheetName', 'publishedSpreadsheetId'];
  for (const field of requiredFields) {
    if (config[field] === undefined) {
      errors.push(`必須フィールド '${field}' が未定義です`);
    }
  }
  
  // 公開状態の整合性チェック
  if (config.appPublished && !config.publishedSheetName) {
    errors.push('appPublished=trueですが、publishedSheetNameが未設定です');
  }
  
  return {
    isValid: errors.length === 0,
    errors
  };
}

/**
 * ConfigJSONの文字列化（保存用）
 * @param {Object} config - 設定オブジェクト
 * @returns {string} JSON文字列
 */
function stringifyConfigJSON(config) {
  const normalized = normalizeConfigJSON(config);
  return JSON.stringify(normalized);
}

/**
 * シート固有設定の統一処理
 * @param {Object} config - ConfigJSON
 * @param {string} sheetName - シート名
 * @param {Object} sheetConfig - シート設定
 * @returns {Object} 更新されたConfigJSON
 */
function setSheetConfig(config, sheetName, sheetConfig) {
  if (!config || !sheetName) return config;
  
  const sheetKey = `sheet_${sheetName}`;
  const normalized = normalizeConfigJSON(config);
  normalized[sheetKey] = normalizeSheetConfig(sheetConfig);
  
  return normalized;
}

/**
 * シート固有設定の取得
 * @param {Object} config - ConfigJSON
 * @param {string} sheetName - シート名
 * @returns {Object} シート固有設定
 */
function getSheetConfig(config, sheetName) {
  if (!config || !sheetName) return {};
  
  const sheetKey = `sheet_${sheetName}`;
  return config[sheetKey] || {};
}