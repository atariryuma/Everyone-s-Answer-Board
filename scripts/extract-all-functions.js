#!/usr/bin/env node

/**
 * 全関数抽出・完全検証スクリプト
 * 
 * 目的：
 * 1. 実装されている全関数を抽出
 * 2. 仕様書に記載されている関数と完全照合
 * 3. 不一致を特定し、正しいマッピングを作成
 */

const fs = require('fs');
const path = require('path');

const PROJECT_ROOT = path.resolve(__dirname, '..');
const SRC_DIR = path.join(PROJECT_ROOT, 'src');
const SPEC_FILE = path.join(PROJECT_ROOT, 'SYSTEM_SPECIFICATIONS.md');

// カラー出力
const colors = {
  reset: '\x1b[0m',
  red: '\x1b[31m',
  green: '\x1b[32m',
  yellow: '\x1b[33m',
  blue: '\x1b[34m',
  cyan: '\x1b[36m',
  magenta: '\x1b[35m'
};

/**
 * GSファイルから全関数を抽出
 */
function extractAllFunctions() {
  const functions = {};
  const gsFiles = fs.readdirSync(SRC_DIR)
    .filter(file => file.endsWith('.gs'))
    .sort();

  gsFiles.forEach(file => {
    const filePath = path.join(SRC_DIR, file);
    const content = fs.readFileSync(filePath, 'utf8');
    
    // 関数定義パターン（複数形式対応）
    const patterns = [
      // 通常の関数定義（行頭の空白を許可）
      /^\s*function\s+(\w+)\s*\([^)]*\)\s*{/gm,
      // constで定義された関数
      /^\s*const\s+(\w+)\s*=\s*function\s*\([^)]*\)\s*{/gm,
      // アロー関数
      /^\s*const\s+(\w+)\s*=\s*\([^)]*\)\s*=>\s*{/gm,
      // アロー関数（引数なし）
      /^\s*const\s+(\w+)\s*=\s*\(\)\s*=>\s*{/gm,
      // オブジェクト内のメソッド（例：App.init）
      /^\s{2,}(\w+)\s*\([^)]*\)\s*{/gm,
      // オブジェクト内のメソッド（プロパティ形式）
      /^\s{2,}(\w+):\s*function\s*\([^)]*\)\s*{/gm,
      // オブジェクト内のアロー関数
      /^\s{2,}(\w+):\s*\([^)]*\)\s*=>\s*{/gm
    ];

    const fileFunctions = new Set();
    
    patterns.forEach(pattern => {
      let match;
      while ((match = pattern.exec(content)) !== null) {
        const funcName = match[1];
        // テスト関数やプライベート関数を除外
        if (!funcName.startsWith('_') && !funcName.startsWith('test')) {
          fileFunctions.add(funcName);
        }
      }
    });

    if (fileFunctions.size > 0) {
      functions[file] = Array.from(fileFunctions).sort();
    }
  });

  return functions;
}

/**
 * 名前空間オブジェクトのメソッドを抽出
 */
function extractNamespaceMethods() {
  const namespaces = {};
  
  const namespacePatternsMap = {
    'App.gs': ['App'],
    'database.gs': ['DB'],
    'ConfigManager.gs': ['ConfigManager'],
    'SystemManager.gs': ['SystemManager'],
    'Base.gs': ['UserIdResolver', 'UserManager'],
    'security.gs': ['SecurityValidator']
  };

  Object.entries(namespacePatternsMap).forEach(([fileName, namespaceNames]) => {
    const filePath = path.join(SRC_DIR, fileName);
    if (!fs.existsSync(filePath)) return;
    
    const content = fs.readFileSync(filePath, 'utf8');
    
    namespaceNames.forEach(namespace => {
      const nsPattern = new RegExp(`const\\s+${namespace}\\s*=\\s*(?:Object\\.freeze\\()?{([^}]+)}`, 's');
      const nsMatch = nsPattern.exec(content);
      
      if (nsMatch) {
        const nsContent = nsMatch[1];
        const methodPattern = /(\w+)\s*(?:\([^)]*\)|:)/g;
        const methods = [];
        let methodMatch;
        
        while ((methodMatch = methodPattern.exec(nsContent)) !== null) {
          const methodName = methodMatch[1];
          if (methodName !== 'freeze' && methodName !== 'Object') {
            methods.push(`${namespace}.${methodName}`);
          }
        }
        
        if (methods.length > 0) {
          namespaces[fileName] = namespaces[fileName] || [];
          namespaces[fileName].push(...methods);
        }
      }
    });
  });

  return namespaces;
}

/**
 * 仕様書から関数を抽出
 */
function extractSpecFunctions() {
  const specContent = fs.readFileSync(SPEC_FILE, 'utf8');
  const functions = {
    api: new Set(),
    internal: new Set(),
    mentioned: new Set()
  };

  // API関数（→ を含む関数定義）
  const apiPattern = /^function\s+(\w+)\s*\([^)]*\)\s*→/gm;
  let match;
  while ((match = apiPattern.exec(specContent)) !== null) {
    functions.api.add(match[1]);
  }

  // コードブロック内の関数
  const codeBlockPattern = /```(?:javascript|js)?\s*([\s\S]*?)```/g;
  while ((match = codeBlockPattern.exec(specContent)) !== null) {
    const codeContent = match[1];
    const funcPattern = /function\s+(\w+)\s*\(/g;
    let funcMatch;
    while ((funcMatch = funcPattern.exec(codeContent)) !== null) {
      functions.internal.add(funcMatch[1]);
    }
  }

  // 本文中で言及されている関数（関数名()形式）
  const mentionPattern = /`(\w+)\(\)`/g;
  while ((match = mentionPattern.exec(specContent)) !== null) {
    functions.mentioned.add(match[1]);
  }

  return functions;
}

/**
 * 類似度計算（レーベンシュタイン距離）
 */
function levenshteinDistance(str1, str2) {
  const m = str1.length;
  const n = str2.length;
  const dp = Array(m + 1).fill(null).map(() => Array(n + 1).fill(0));

  for (let i = 0; i <= m; i++) dp[i][0] = i;
  for (let j = 0; j <= n; j++) dp[0][j] = j;

  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      if (str1[i - 1] === str2[j - 1]) {
        dp[i][j] = dp[i - 1][j - 1];
      } else {
        dp[i][j] = Math.min(
          dp[i - 1][j] + 1,
          dp[i][j - 1] + 1,
          dp[i - 1][j - 1] + 1
        );
      }
    }
  }

  return dp[m][n];
}

/**
 * 類似関数を検索
 */
function findSimilarFunction(targetFunc, allFunctions) {
  let bestMatch = null;
  let bestScore = Infinity;
  let bestFile = null;

  Object.entries(allFunctions).forEach(([file, funcs]) => {
    funcs.forEach(func => {
      const distance = levenshteinDistance(targetFunc.toLowerCase(), func.toLowerCase());
      if (distance < bestScore) {
        bestScore = distance;
        bestMatch = func;
        bestFile = file;
      }
    });
  });

  // 類似度が3文字以内の差なら候補として返す
  if (bestScore <= 3) {
    return { function: bestMatch, file: bestFile, score: bestScore };
  }

  return null;
}

/**
 * メイン検証処理
 */
function main() {
  console.log(`${colors.blue}${'='.repeat(60)}${colors.reset}`);
  console.log(`${colors.cyan}🔍 全関数抽出・完全検証スクリプト${colors.reset}`);
  console.log(`${colors.blue}${'='.repeat(60)}${colors.reset}\n`);

  // 1. 実装から全関数を抽出
  console.log(`${colors.blue}[1/5] 実装されている全関数を抽出中...${colors.reset}`);
  const implementedFunctions = extractAllFunctions();
  const namespaceMethods = extractNamespaceMethods();
  
  // 名前空間メソッドを統合
  Object.entries(namespaceMethods).forEach(([file, methods]) => {
    implementedFunctions[file] = implementedFunctions[file] || [];
    implementedFunctions[file].push(...methods);
  });

  let totalImplFuncs = 0;
  Object.values(implementedFunctions).forEach(funcs => {
    totalImplFuncs += funcs.length;
  });
  console.log(`${colors.green}✓ ${totalImplFuncs}個の関数を検出${colors.reset}\n`);

  // 2. 仕様書から関数を抽出
  console.log(`${colors.blue}[2/5] 仕様書から関数を抽出中...${colors.reset}`);
  const specFunctions = extractSpecFunctions();
  console.log(`${colors.green}✓ API関数: ${specFunctions.api.size}個${colors.reset}`);
  console.log(`${colors.green}✓ 内部関数: ${specFunctions.internal.size}個${colors.reset}`);
  console.log(`${colors.green}✓ 言及関数: ${specFunctions.mentioned.size}個${colors.reset}\n`);

  // 3. API関数の完全照合
  console.log(`${colors.blue}[3/5] API関数の完全照合...${colors.reset}`);
  const apiMismatches = [];
  const apiMatches = [];

  specFunctions.api.forEach(specFunc => {
    let found = false;
    let foundIn = null;

    // 完全一致を検索（名前空間メソッドも考慮）
    for (const [file, funcs] of Object.entries(implementedFunctions)) {
      if (funcs.includes(specFunc) || 
          funcs.some(func => func.split('.').pop() === specFunc)) {
        found = true;
        foundIn = file;
        break; // 見つかったらすぐに終了
      }
    }

    if (found) {
      apiMatches.push({ func: specFunc, file: foundIn });
      console.log(`${colors.green}✓ ${specFunc}() - ${foundIn}${colors.reset}`);
    } else {
      // 類似関数を検索
      const similar = findSimilarFunction(specFunc, implementedFunctions);
      apiMismatches.push({ spec: specFunc, similar });
      
      if (similar) {
        console.log(`${colors.yellow}⚠ ${specFunc}() - 未実装（類似: ${similar.function} in ${similar.file}）${colors.reset}`);
      } else {
        console.log(`${colors.red}✗ ${specFunc}() - 完全に未実装${colors.reset}`);
      }
    }
  });

  // 4. 実装されているが仕様書にない主要関数
  console.log(`\n${colors.blue}[4/5] 仕様書に記載されていない主要関数...${colors.reset}`);
  const undocumentedFunctions = [];

  Object.entries(implementedFunctions).forEach(([file, funcs]) => {
    funcs.forEach(func => {
      // 名前空間メソッドはスキップ
      if (func.includes('.')) return;
      
      // 主要な関数のみ（doで始まる、handle、get、set、create、update、deleteなど）
      const isMainFunction = /^(do|handle|get|set|create|update|delete|save|load|render|validate|check)/i.test(func);
      
      if (isMainFunction) {
        const isInSpec = specFunctions.api.has(func) || 
                         specFunctions.internal.has(func) || 
                         specFunctions.mentioned.has(func);
        
        if (!isInSpec) {
          undocumentedFunctions.push({ func, file });
        }
      }
    });
  });

  if (undocumentedFunctions.length > 0) {
    console.log(`${colors.yellow}以下の主要関数が仕様書に記載されていません:${colors.reset}`);
    undocumentedFunctions.slice(0, 10).forEach(({ func, file }) => {
      console.log(`  - ${func}() in ${file}`);
    });
    if (undocumentedFunctions.length > 10) {
      console.log(`  ... 他 ${undocumentedFunctions.length - 10} 件`);
    }
  }

  // 5. レポート生成
  console.log(`\n${colors.blue}[5/5] 検証レポート生成...${colors.reset}`);
  
  const report = {
    timestamp: new Date().toISOString(),
    summary: {
      totalImplemented: totalImplFuncs,
      totalSpecAPI: specFunctions.api.size,
      matchedAPI: apiMatches.length,
      mismatchedAPI: apiMismatches.length,
      undocumentedMain: undocumentedFunctions.length
    },
    apiMismatches: apiMismatches,
    undocumentedFunctions: undocumentedFunctions,
    implementedFunctions: implementedFunctions,
    specFunctions: {
      api: Array.from(specFunctions.api),
      internal: Array.from(specFunctions.internal),
      mentioned: Array.from(specFunctions.mentioned)
    }
  };

  // レポートファイル出力
  const reportPath = path.join(PROJECT_ROOT, 'function-validation-report.json');
  fs.writeFileSync(reportPath, JSON.stringify(report, null, 2));
  console.log(`${colors.green}✓ レポート生成完了: ${reportPath}${colors.reset}`);

  // 最終サマリー
  console.log(`\n${colors.blue}${'='.repeat(60)}${colors.reset}`);
  console.log(`${colors.cyan}📊 検証結果サマリー${colors.reset}`);
  console.log(`${colors.blue}${'='.repeat(60)}${colors.reset}\n`);

  const matchRate = ((apiMatches.length / specFunctions.api.size) * 100).toFixed(1);
  console.log(`実装関数総数: ${totalImplFuncs}`);
  console.log(`仕様書API関数: ${specFunctions.api.size}`);
  console.log(`${colors.green}✓ 一致: ${apiMatches.length}${colors.reset}`);
  console.log(`${colors.red}✗ 不一致: ${apiMismatches.length}${colors.reset}`);
  console.log(`${colors.yellow}⚠ 未文書化: ${undocumentedFunctions.length}${colors.reset}`);
  console.log(`\nAPI関数一致率: ${matchRate}%`);

  if (apiMismatches.length > 0) {
    console.log(`\n${colors.red}⚠️ 重要: ${apiMismatches.length}個のAPI関数名が不一致です${colors.reset}`);
    console.log('修正が必要な関数:');
    apiMismatches.forEach(({ spec, similar }) => {
      if (similar) {
        console.log(`  ${colors.yellow}仕様書: ${spec}() → 実装: ${similar.function}()${colors.reset}`);
      } else {
        console.log(`  ${colors.red}仕様書: ${spec}() → 実装なし${colors.reset}`);
      }
    });
  }

  // 終了コード
  process.exit(apiMismatches.length > 0 ? 1 : 0);
}

// 実行
main();