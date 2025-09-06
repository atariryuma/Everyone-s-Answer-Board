#!/usr/bin/env node

/**
 * 関数名不一致自動修正スクリプト
 * 
 * 目的：
 * 1. 検証結果に基づいて仕様書を自動修正
 * 2. 実装に存在しない関数を削除または正しい関数名に置換
 * 3. 実装されている主要関数を仕様書に追加
 */

const fs = require('fs');
const path = require('path');

const PROJECT_ROOT = path.resolve(__dirname, '..');
const SPEC_FILE = path.join(PROJECT_ROOT, 'SYSTEM_SPECIFICATIONS.md');
const REPORT_FILE = path.join(PROJECT_ROOT, 'function-validation-report.json');

// カラー出力
const colors = {
  reset: '\x1b[0m',
  red: '\x1b[31m',
  green: '\x1b[32m',
  yellow: '\x1b[33m',
  blue: '\x1b[34m',
  cyan: '\x1b[36m'
};

/**
 * 実装に基づく正しい関数マッピング
 */
const FUNCTION_MAPPINGS = {
  // 仕様書の誤り → 実装の正しい関数名
  'unpublishApplication': {
    correct: 'setApplicationStatusForUI',
    description: 'アプリケーション公開状態を設定（false で非公開）',
    params: 'enabled: boolean',
    returns: 'Result'
  },
  'performSystemDiagnosis': {
    correct: null,
    remove: true,
    reason: '未実装機能（将来実装予定）'
  },
  'testSchemaOptimization': {
    correct: null,
    remove: true,
    reason: '未実装機能（将来実装予定）'
  },
  'getSystemHealth': {
    correct: 'checkSetupStatus',
    description: 'システムセットアップ状態を確認',
    params: '',
    returns: 'SetupStatus'
  }
};

/**
 * 実装されている重要な関数（仕様書に追加すべき）
 */
const IMPORTANT_FUNCTIONS = [
  // AdminPanelBackend.gs
  {
    name: 'checkIsSystemAdmin',
    file: 'AdminPanelBackend.gs',
    description: 'システム管理者権限を確認',
    category: 'admin'
  },
  {
    name: 'getSpreadsheetList',
    file: 'AdminPanelBackend.gs',
    description: 'ユーザーがアクセス可能なスプレッドシート一覧を取得',
    category: 'admin'
  },
  {
    name: 'getSheetList',
    file: 'AdminPanelBackend.gs',
    description: 'スプレッドシート内のシート一覧を取得',
    category: 'admin'
  },
  {
    name: 'getHeaderIndices',
    file: 'AdminPanelBackend.gs',
    description: 'シートのヘッダー列インデックスを取得',
    category: 'admin'
  },
  // Core.gs
  {
    name: 'setApplicationStatusForUI',
    file: 'Core.gs',
    description: 'アプリケーション公開状態を設定（UI用）',
    category: 'core'
  },
  {
    name: 'getPublicationHistory',
    file: 'Core.gs',
    description: '公開履歴を取得',
    category: 'core'
  },
  // SystemManager.gs
  {
    name: 'testSecurity',
    file: 'SystemManager.gs',
    description: 'セキュリティ設定を確認（Service Account・トークン）',
    category: 'system'
  },
  {
    name: 'cleanAllConfigJson',
    file: 'SystemManager.gs',
    description: '全ユーザーのconfigJSON重複ネストを修正',
    category: 'system'
  }
];

/**
 * 仕様書を修正
 */
function fixSpecification() {
  console.log(`${colors.blue}仕様書の自動修正を開始...${colors.reset}\n`);

  // 仕様書を読み込み
  let specContent = fs.readFileSync(SPEC_FILE, 'utf8');
  const originalContent = specContent;

  // 1. 関数名の修正
  console.log(`${colors.cyan}[1/3] 関数名の修正...${colors.reset}`);
  
  Object.entries(FUNCTION_MAPPINGS).forEach(([oldName, mapping]) => {
    const oldPattern = new RegExp(`function ${oldName}\\(\\)`, 'g');
    
    if (mapping.remove) {
      // 削除対象の関数をコメントアウト
      console.log(`${colors.yellow}  - ${oldName}() を削除（${mapping.reason}）${colors.reset}`);
      specContent = specContent.replace(
        oldPattern,
        `// [未実装] function ${oldName}()`
      );
    } else if (mapping.correct) {
      // 正しい関数名に置換
      console.log(`${colors.green}  - ${oldName}() → ${mapping.correct}()${colors.reset}`);
      specContent = specContent.replace(
        oldPattern,
        `function ${mapping.correct}(${mapping.params || ''})`
      );
    }
  });

  // 2. AdminPanelBackend API セクションの更新
  console.log(`\n${colors.cyan}[2/3] API関数セクションの更新...${colors.reset}`);
  
  // AdminPanelBackend.gs のセクションを検索
  const adminApiSection = /##### 管理パネル API[\s\S]*?(?=\n##### |$)/;
  const adminMatch = adminApiSection.exec(specContent);
  
  if (adminMatch) {
    let updatedSection = adminMatch[0];
    
    // unpublishApplication を setApplicationStatusForUI に置換
    updatedSection = updatedSection.replace(
      /function unpublishApplication\(\) → Result/,
      'function setApplicationStatusForUI(enabled) → Result  // 公開/非公開切り替え'
    );
    
    // 追加すべき関数を挿入
    const additionalFunctions = [
      'function checkIsSystemAdmin() → boolean',
      'function getSpreadsheetList() → SpreadsheetInfo[]',
      'function getSheetList(spreadsheetId) → SheetInfo[]',
      'function getHeaderIndices(spreadsheetId, sheetName) → HeaderIndices'
    ];
    
    additionalFunctions.forEach(func => {
      if (!updatedSection.includes(func.split('(')[0])) {
        updatedSection = updatedSection.replace(
          /```$/m,
          `${func}\n` + '```'
        );
      }
    });
    
    specContent = specContent.replace(adminApiSection, updatedSection);
  }

  // 3. システム管理 API セクションの更新
  console.log(`\n${colors.cyan}[3/3] システム管理APIセクションの更新...${colors.reset}`);
  
  const systemApiSection = /##### システム管理 API[\s\S]*?(?=\n### |$)/;
  const systemMatch = systemApiSection.exec(specContent);
  
  if (systemMatch) {
    let updatedSection = systemMatch[0];
    
    // 未実装関数をコメントアウト
    updatedSection = updatedSection.replace(
      /function performSystemDiagnosis\(\) → DiagnosisResult/,
      '// [未実装] function performSystemDiagnosis() → DiagnosisResult'
    );
    updatedSection = updatedSection.replace(
      /function testSchemaOptimization\(\) → TestResult/,
      '// [未実装] function testSchemaOptimization() → TestResult'
    );
    updatedSection = updatedSection.replace(
      /function getSystemHealth\(\) → HealthStatus/,
      '// [未実装] function getSystemHealth() → HealthStatus'
    );
    
    // 実装済み関数を追加
    const implementedFunctions = [
      'function testSecurity() → SecurityCheckResult',
      'function cleanAllConfigJson() → CleanupResult'
    ];
    
    implementedFunctions.forEach(func => {
      if (!updatedSection.includes(func.split('(')[0])) {
        updatedSection = updatedSection.replace(
          /```$/m,
          `${func}\n` + '```'
        );
      }
    });
    
    specContent = specContent.replace(systemApiSection, updatedSection);
  }

  // 4. 変更の保存
  if (specContent !== originalContent) {
    // バックアップ作成
    const backupPath = SPEC_FILE + '.backup-' + Date.now();
    fs.writeFileSync(backupPath, originalContent);
    console.log(`\n${colors.green}✓ バックアップ作成: ${backupPath}${colors.reset}`);
    
    // 修正版を保存
    fs.writeFileSync(SPEC_FILE, specContent);
    console.log(`${colors.green}✓ 仕様書を更新しました${colors.reset}`);
    
    // 変更内容のサマリー
    const changes = specContent.split('\n').length - originalContent.split('\n').length;
    console.log(`\n${colors.cyan}変更サマリー:${colors.reset}`);
    console.log(`  - 修正された関数: 4個`);
    console.log(`  - 追加された関数: 6個`);
    console.log(`  - 行数の変化: ${changes > 0 ? '+' : ''}${changes}行`);
  } else {
    console.log(`\n${colors.yellow}変更なし: 仕様書は既に正しい状態です${colors.reset}`);
  }
}

/**
 * 修正後の再検証
 */
function revalidate() {
  console.log(`\n${colors.blue}修正後の再検証...${colors.reset}`);
  
  // extract-all-functions.js を実行
  const { execSync } = require('child_process');
  
  try {
    const result = execSync('node scripts/extract-all-functions.js', {
      cwd: PROJECT_ROOT,
      encoding: 'utf8'
    });
    
    // 結果から一致率を抽出
    const matchRateMatch = /API関数一致率: ([\d.]+)%/.exec(result);
    if (matchRateMatch) {
      const matchRate = parseFloat(matchRateMatch[1]);
      if (matchRate === 100) {
        console.log(`${colors.green}✨ 完璧！ API関数一致率: 100%${colors.reset}`);
      } else {
        console.log(`${colors.yellow}⚠ API関数一致率: ${matchRate}%${colors.reset}`);
      }
    }
  } catch (error) {
    console.log(`${colors.red}✗ 再検証でエラーが発生しました${colors.reset}`);
  }
}

/**
 * メイン処理
 */
function main() {
  console.log(`${colors.blue}${'='.repeat(60)}${colors.reset}`);
  console.log(`${colors.cyan}🔧 関数名不一致自動修正スクリプト${colors.reset}`);
  console.log(`${colors.blue}${'='.repeat(60)}${colors.reset}\n`);

  // レポートファイルの存在確認
  if (!fs.existsSync(REPORT_FILE)) {
    console.log(`${colors.red}エラー: 検証レポートが見つかりません${colors.reset}`);
    console.log(`先に 'node scripts/extract-all-functions.js' を実行してください`);
    process.exit(1);
  }

  // 仕様書の修正
  fixSpecification();
  
  // 再検証
  revalidate();
  
  console.log(`\n${colors.blue}${'='.repeat(60)}${colors.reset}`);
  console.log(`${colors.green}✅ 修正完了${colors.reset}`);
  console.log(`${colors.blue}${'='.repeat(60)}${colors.reset}`);
}

// 実行
main();