#!/usr/bin/env node

/**
 * システムスペック仕様書と実装の整合性検証スクリプト
 * 
 * 検証項目:
 * 1. ファイル構成の一致性
 * 2. モジュール記載の完全性
 * 3. API関数の存在確認
 * 4. 定数・設定値の一致性
 * 5. パフォーマンス数値の検証
 */

const fs = require('fs');
const path = require('path');

// カラー出力用のANSIコード
const colors = {
  reset: '\x1b[0m',
  red: '\x1b[31m',
  green: '\x1b[32m',
  yellow: '\x1b[33m',
  blue: '\x1b[34m',
  cyan: '\x1b[36m'
};

// ヘルパー関数
const log = {
  info: (msg) => console.log(`${colors.cyan}ℹ${colors.reset} ${msg}`),
  success: (msg) => console.log(`${colors.green}✓${colors.reset} ${msg}`),
  warning: (msg) => console.log(`${colors.yellow}⚠${colors.reset} ${msg}`),
  error: (msg) => console.log(`${colors.red}✗${colors.reset} ${msg}`),
  section: (msg) => console.log(`\n${colors.blue}═══ ${msg} ═══${colors.reset}\n`)
};

// プロジェクトルート
const PROJECT_ROOT = path.resolve(__dirname, '..');
const SRC_DIR = path.join(PROJECT_ROOT, 'src');
const SPEC_FILE = path.join(PROJECT_ROOT, 'SYSTEM_SPECIFICATIONS.md');

// 検証結果
const validationResults = {
  totalChecks: 0,
  passed: 0,
  warnings: 0,
  errors: 0,
  details: []
};

/**
 * ファイルリストを取得
 */
function getFiles(dir, extension) {
  try {
    return fs.readdirSync(dir)
      .filter(file => file.endsWith(extension))
      .sort();
  } catch (error) {
    return [];
  }
}

/**
 * 仕様書の内容を読み込み
 */
function loadSpecification() {
  try {
    return fs.readFileSync(SPEC_FILE, 'utf8');
  } catch (error) {
    log.error(`仕様書ファイルが見つかりません: ${SPEC_FILE}`);
    process.exit(1);
  }
}

/**
 * 1. ファイル構成の検証
 */
function validateFileStructure(specContent) {
  log.section('ファイル構成の検証');
  
  // 実際のファイル
  const actualGsFiles = getFiles(SRC_DIR, '.gs');
  const actualHtmlFiles = getFiles(SRC_DIR, '.html');
  
  // 仕様書に記載されているファイル（正規表現で抽出）
  const gsPattern = /\*\*([^*]+\.gs)\*\*/g;
  const htmlPattern = /\*\*([^*]+\.html)\*\*/g;
  
  const specGsFiles = new Set();
  const specHtmlFiles = new Set();
  
  let match;
  while ((match = gsPattern.exec(specContent)) !== null) {
    specGsFiles.add(match[1]);
  }
  while ((match = htmlPattern.exec(specContent)) !== null) {
    specHtmlFiles.add(match[1]);
  }
  
  // GSファイルの検証
  log.info(`GSファイル検証: 実装 ${actualGsFiles.length}個 vs 仕様書 ${specGsFiles.size}個`);
  
  actualGsFiles.forEach(file => {
    validationResults.totalChecks++;
    if (specGsFiles.has(file)) {
      log.success(`${file} - 仕様書に記載あり`);
      validationResults.passed++;
    } else {
      log.warning(`${file} - 仕様書に記載なし`);
      validationResults.warnings++;
      validationResults.details.push({
        type: 'warning',
        category: 'file',
        message: `GSファイル '${file}' が仕様書に記載されていません`
      });
    }
  });
  
  // HTMLファイルの検証
  log.info(`\nHTMLファイル検証: 実装 ${actualHtmlFiles.length}個 vs 仕様書 ${specHtmlFiles.size}個`);
  
  actualHtmlFiles.forEach(file => {
    validationResults.totalChecks++;
    if (specHtmlFiles.has(file)) {
      log.success(`${file} - 仕様書に記載あり`);
      validationResults.passed++;
    } else {
      log.warning(`${file} - 仕様書に記載なし`);
      validationResults.warnings++;
      validationResults.details.push({
        type: 'warning',
        category: 'file',
        message: `HTMLファイル '${file}' が仕様書に記載されていません`
      });
    }
  });
}

/**
 * 2. API関数の存在確認
 */
function validateAPIFunctions(specContent) {
  log.section('API関数の存在確認');
  
  // 仕様書からAPI関数を抽出
  const apiFunctionPattern = /^function\s+(\w+)\([^)]*\)\s*→/gm;
  const specFunctions = new Set();
  
  let match;
  while ((match = apiFunctionPattern.exec(specContent)) !== null) {
    specFunctions.add(match[1]);
  }
  
  log.info(`仕様書に記載されているAPI関数: ${specFunctions.size}個`);
  
  // 実際のコードから関数を検索
  specFunctions.forEach(funcName => {
    validationResults.totalChecks++;
    let found = false;
    
    // GSファイルから検索
    const gsFiles = getFiles(SRC_DIR, '.gs');
    for (const file of gsFiles) {
      const content = fs.readFileSync(path.join(SRC_DIR, file), 'utf8');
      const funcPattern = new RegExp(`function\\s+${funcName}\\s*\\(`, 'm');
      
      if (funcPattern.test(content)) {
        log.success(`${funcName}() - ${file} に実装あり`);
        validationResults.passed++;
        found = true;
        break;
      }
    }
    
    if (!found) {
      log.error(`${funcName}() - 実装が見つかりません`);
      validationResults.errors++;
      validationResults.details.push({
        type: 'error',
        category: 'api',
        message: `API関数 '${funcName}()' の実装が見つかりません`
      });
    }
  });
}

/**
 * 3. データベース構造の検証
 */
function validateDatabaseStructure(specContent) {
  log.section('データベース構造の検証');
  
  // database.gsを読み込み
  const dbFile = path.join(SRC_DIR, 'database.gs');
  let dbContent = '';
  
  try {
    dbContent = fs.readFileSync(dbFile, 'utf8');
  } catch (error) {
    log.error('database.gsが見つかりません');
    validationResults.errors++;
    return;
  }
  
  // DB_CONFIGの検証
  const dbConfigPattern = /const\s+DB_CONFIG\s*=\s*Object\.freeze\(\{[\s\S]*?HEADERS:\s*Object\.freeze\(\[([\s\S]*?)\]\)/;
  const match = dbConfigPattern.exec(dbContent);
  
  if (match) {
    const headers = match[1]
      .split(',')
      .map(h => h.trim().replace(/['"]/g, '').replace(/\/\/.*$/, '').trim())
      .filter(h => h);
    
    const expectedHeaders = ['userId', 'userEmail', 'isActive', 'configJson', 'lastModified'];
    
    validationResults.totalChecks++;
    const isMatch = JSON.stringify(headers) === JSON.stringify(expectedHeaders);
    
    if (isMatch) {
      log.success('DB_CONFIG.HEADERS - 5フィールド構造が仕様書と一致');
      validationResults.passed++;
    } else {
      log.error(`DB_CONFIG.HEADERS - 不一致\n  実装: ${headers.join(', ')}\n  仕様: ${expectedHeaders.join(', ')}`);
      validationResults.errors++;
      validationResults.details.push({
        type: 'error',
        category: 'database',
        message: 'データベース構造が仕様書と一致しません'
      });
    }
  } else {
    log.error('DB_CONFIGが見つかりません');
    validationResults.errors++;
  }
}

/**
 * 4. キャッシュ設定の検証
 */
function validateCacheSettings(specContent) {
  log.section('キャッシュ設定の検証');
  
  // cache.gsを読み込み
  const cacheFile = path.join(SRC_DIR, 'cache.gs');
  let cacheContent = '';
  
  try {
    cacheContent = fs.readFileSync(cacheFile, 'utf8');
  } catch (error) {
    log.error('cache.gsが見つかりません');
    validationResults.errors++;
    return;
  }
  
  // TTL設定の検証
  const ttlPattern = /ttl:\s*(\d+)/g;
  const ttlValues = [];
  let match;
  
  while ((match = ttlPattern.exec(cacheContent)) !== null) {
    ttlValues.push(parseInt(match[1]));
  }
  
  // 仕様書のTTL値
  const specTTLs = {
    'userInfo': 300,
    'sheetsService': 3500,
    'headerIndices': 1800
  };
  
  log.info(`キャッシュTTL値: ${ttlValues.length}個検出`);
  
  // 3500秒の検証（sheetsService）
  validationResults.totalChecks++;
  if (ttlValues.includes(3500)) {
    log.success('sheetsService TTL (3500秒) - 実装あり');
    validationResults.passed++;
  } else {
    log.warning('sheetsService TTL (3500秒) - 実装で見つからない');
    validationResults.warnings++;
  }
}

/**
 * 5. セキュリティ設定の検証
 */
function validateSecuritySettings(specContent) {
  log.section('セキュリティ設定の検証');
  
  // SharedSecurityHeaders.htmlを読み込み
  const securityFile = path.join(SRC_DIR, 'SharedSecurityHeaders.html');
  let securityContent = '';
  
  try {
    securityContent = fs.readFileSync(securityFile, 'utf8');
  } catch (error) {
    log.error('SharedSecurityHeaders.htmlが見つかりません');
    validationResults.errors++;
    return;
  }
  
  // 必須セキュリティヘッダーの検証
  const requiredHeaders = [
    'Content-Security-Policy',
    'Permissions-Policy',
    'Cache-Control',
    'X-Content-Type-Options'
  ];
  
  requiredHeaders.forEach(header => {
    validationResults.totalChecks++;
    if (securityContent.includes(header)) {
      log.success(`${header} - 実装あり`);
      validationResults.passed++;
    } else {
      log.error(`${header} - 実装なし`);
      validationResults.errors++;
      validationResults.details.push({
        type: 'error',
        category: 'security',
        message: `セキュリティヘッダー '${header}' が実装されていません`
      });
    }
  });
}

/**
 * 6. パフォーマンス数値の検証
 */
function validatePerformanceMetrics(specContent) {
  log.section('パフォーマンス数値の検証');
  
  // 仕様書のパフォーマンス数値
  const performanceMetrics = [
    { metric: '60%高速化', context: '取得速度' },
    { metric: '70%効率化', context: '更新効率' },
    { metric: '40%削減', context: 'メモリ使用' }
  ];
  
  performanceMetrics.forEach(({ metric, context }) => {
    validationResults.totalChecks++;
    if (specContent.includes(metric)) {
      log.info(`${context}: ${metric} - 仕様書に記載あり`);
      validationResults.passed++;
      
      // 測定条件の記載確認
      const hasCondition = specContent.includes('測定') || specContent.includes('N=');
      if (!hasCondition) {
        log.warning(`  └─ 測定条件の記載なし`);
        validationResults.warnings++;
        validationResults.details.push({
          type: 'warning',
          category: 'performance',
          message: `パフォーマンス数値 '${metric}' の測定条件が記載されていません`
        });
      }
    }
  });
}

/**
 * 7. 定数定義の検証
 */
function validateConstants(specContent) {
  log.section('定数定義の検証');
  
  // constants.gsを読み込み
  const constFile = path.join(SRC_DIR, 'constants.gs');
  let constContent = '';
  
  try {
    constContent = fs.readFileSync(constFile, 'utf8');
  } catch (error) {
    log.error('constants.gsが見つかりません');
    validationResults.errors++;
    return;
  }
  
  // SYSTEM_CONSTANTSの主要項目
  const requiredConstants = [
    'DATABASE',
    'REACTIONS',
    'COLUMNS',
    'COLUMN_MAPPING',
    'DISPLAY_MODES',
    'ACCESS'
  ];
  
  requiredConstants.forEach(constant => {
    validationResults.totalChecks++;
    if (constContent.includes(`${constant}:`)) {
      log.success(`SYSTEM_CONSTANTS.${constant} - 実装あり`);
      validationResults.passed++;
    } else {
      log.error(`SYSTEM_CONSTANTS.${constant} - 実装なし`);
      validationResults.errors++;
      validationResults.details.push({
        type: 'error',
        category: 'constants',
        message: `定数 'SYSTEM_CONSTANTS.${constant}' が実装されていません`
      });
    }
  });
}

/**
 * 結果サマリーの表示
 */
function displaySummary() {
  log.section('検証結果サマリー');
  
  const passRate = ((validationResults.passed / validationResults.totalChecks) * 100).toFixed(1);
  
  console.log(`総チェック数: ${validationResults.totalChecks}`);
  console.log(`${colors.green}成功: ${validationResults.passed}${colors.reset}`);
  console.log(`${colors.yellow}警告: ${validationResults.warnings}${colors.reset}`);
  console.log(`${colors.red}エラー: ${validationResults.errors}${colors.reset}`);
  console.log(`\n合格率: ${passRate}%`);
  
  // 詳細レポート
  if (validationResults.details.length > 0) {
    log.section('詳細レポート');
    
    const errors = validationResults.details.filter(d => d.type === 'error');
    const warnings = validationResults.details.filter(d => d.type === 'warning');
    
    if (errors.length > 0) {
      console.log(`\n${colors.red}エラー項目:${colors.reset}`);
      errors.forEach(e => console.log(`  • [${e.category}] ${e.message}`));
    }
    
    if (warnings.length > 0) {
      console.log(`\n${colors.yellow}警告項目:${colors.reset}`);
      warnings.forEach(w => console.log(`  • [${w.category}] ${w.message}`));
    }
  }
  
  // 最終評価
  console.log('\n' + '='.repeat(50));
  if (validationResults.errors === 0) {
    if (validationResults.warnings === 0) {
      console.log(`${colors.green}✨ 完璧！仕様書と実装が完全に一致しています${colors.reset}`);
    } else {
      console.log(`${colors.green}✓ 合格: 重大な問題はありません${colors.reset}`);
      console.log(`${colors.yellow}  ただし、${validationResults.warnings}個の警告があります${colors.reset}`);
    }
  } else {
    console.log(`${colors.red}✗ 要修正: ${validationResults.errors}個のエラーがあります${colors.reset}`);
  }
  
  // 終了コード
  process.exit(validationResults.errors > 0 ? 1 : 0);
}

/**
 * メイン実行
 */
function main() {
  console.log('🔍 システムスペック仕様書検証スクリプト');
  console.log('='.repeat(50));
  
  const specContent = loadSpecification();
  
  // 各検証を実行
  validateFileStructure(specContent);
  validateAPIFunctions(specContent);
  validateDatabaseStructure(specContent);
  validateCacheSettings(specContent);
  validateSecuritySettings(specContent);
  validatePerformanceMetrics(specContent);
  validateConstants(specContent);
  
  // サマリー表示
  displaySummary();
}

// 実行
main();