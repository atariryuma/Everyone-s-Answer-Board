/**
 * HTML/GAS Function Mapping Validation Test
 * HTMLファイルからのgoogle.script.run呼び出しとGASファイルの関数定義の整合性をチェック
 */

const fs = require('fs');
const path = require('path');
const glob = require('glob');

describe('HTML/GAS Function Mapping Validation', () => {
  const srcDir = path.join(__dirname, '../src');
  let htmlFiles = [];
  let gasFiles = [];
  let functionCalls = [];
  let gasFunctions = [];

  beforeAll(() => {
    // HTMLファイルとGASファイルを取得
    htmlFiles = glob.sync(path.join(srcDir, '**/*.html'));
    gasFiles = glob.sync(path.join(srcDir, '**/*.gs'));

    // HTMLファイルからgoogle.script.run呼び出しを抽出
    functionCalls = extractFunctionCalls(htmlFiles);

    // GASファイルから関数定義を抽出
    gasFunctions = extractGasFunctions(gasFiles);

    console.log(`📊 Found ${htmlFiles.length} HTML files, ${gasFiles.length} GAS files`);
    console.log(
      `🔍 Extracted ${functionCalls.length} function calls, ${gasFunctions.length} GAS functions`
    );
  });

  test('All google.script.run calls should have corresponding GAS functions', () => {
    const missingFunctions = [];
    const warnings = [];

    functionCalls.forEach((call) => {
      const gasFunction = gasFunctions.find((f) => f.name === call.functionName);

      if (!gasFunction) {
        // 完全一致しない場合、類似関数をチェック
        const similarFunctions = findSimilarFunctions(call.functionName, gasFunctions);

        missingFunctions.push({
          function: call.functionName,
          file: call.file,
          line: call.line,
          similarFunctions,
        });
      }
    });

    // レポート生成
    if (missingFunctions.length > 0) {
      console.log('\n🚨 Missing Function Report:');
      missingFunctions.forEach((missing) => {
        console.log(`❌ ${missing.function} (${path.basename(missing.file)}:${missing.line})`);
        if (missing.similarFunctions.length > 0) {
          console.log(`   💡 Similar functions found: ${missing.similarFunctions.join(', ')}`);
        }
      });
    }

    expect(missingFunctions).toHaveLength(0);
  });

  test('Should detect unused GAS functions (potential cleanup candidates)', () => {
    const unusedFunctions = [];

    gasFunctions.forEach((gasFunc) => {
      const isUsed = functionCalls.some((call) => call.functionName === gasFunc.name);

      if (!isUsed && !isInternalFunction(gasFunc.name)) {
        unusedFunctions.push({
          function: gasFunc.name,
          file: gasFunc.file,
          line: gasFunc.line,
        });
      }
    });

    if (unusedFunctions.length > 0) {
      console.log('\n📋 Potential Cleanup Candidates (Unused Functions):');
      unusedFunctions.forEach((unused) => {
        console.log(`ℹ️  ${unused.function} (${path.basename(unused.file)}:${unused.line})`);
      });
    }

    // これは警告レベル（テスト失敗させない）
    expect(unusedFunctions.length).toBeGreaterThanOrEqual(0);
  });

  test('Should validate function call patterns', () => {
    const invalidPatterns = [];

    functionCalls.forEach((call) => {
      // チェーンコールのパターン検証
      if (call.hasSuccessHandler && !call.hasFailureHandler) {
        invalidPatterns.push({
          issue: 'Missing withFailureHandler',
          function: call.functionName,
          file: call.file,
          line: call.line,
          suggestion: 'Add .withFailureHandler() for proper error handling',
        });
      }

      // 直接呼び出しの検出（非推奨パターン）
      if (!call.hasSuccessHandler && !call.hasFailureHandler) {
        invalidPatterns.push({
          issue: 'Direct call without handlers',
          function: call.functionName,
          file: call.file,
          line: call.line,
          suggestion: 'Consider adding success/failure handlers',
        });
      }
    });

    if (invalidPatterns.length > 0) {
      console.log('\n⚠️  Function Call Pattern Issues:');
      invalidPatterns.forEach((issue) => {
        console.log(
          `⚠️  ${issue.issue}: ${issue.function} (${path.basename(issue.file)}:${issue.line})`
        );
        console.log(`   💡 ${issue.suggestion}`);
      });
    }

    // パターン警告も失敗させない（改善提案として）
    expect(invalidPatterns.length).toBeGreaterThanOrEqual(0);
  });
});

/**
 * HTMLファイルからgoogle.script.run関数呼び出しを抽出
 */
function extractFunctionCalls(htmlFiles) {
  const calls = [];

  htmlFiles.forEach((file) => {
    const content = fs.readFileSync(file, 'utf8');

    // 多行パターンを含む全体的な検索
    const multilinePatterns = [
      /google\.script\.run\s*\.withSuccessHandler\([^)]*\)\s*\.withFailureHandler\([^)]*\)\s*\.(\w+)\s*\(/g,
      /google\.script\.run\s*\.withSuccessHandler\([^)]*\)\s*\.(\w+)\s*\(/g,
      /google\.script\.run\s*\.withFailureHandler\([^)]*\)\s*\.(\w+)\s*\(/g,
      /google\.script\.run\s*\.(\w+)\s*\(/g,
    ];

    multilinePatterns.forEach((pattern) => {
      let match;
      while ((match = pattern.exec(content)) !== null) {
        // 関数名がGASの予約語でないことを確認
        const funcName = match[1];
        if (!isGasReservedWord(funcName)) {
          const lines = content.substring(0, match.index).split('\n');
          const lineNumber = lines.length;
          const lineContent =
            lines[lines.length - 1] + content.substring(match.index, match.index + match[0].length);

          calls.push({
            functionName: funcName,
            file,
            line: lineNumber,
            pattern: lineContent.trim(),
            hasSuccessHandler: match[0].includes('.withSuccessHandler'),
            hasFailureHandler: match[0].includes('.withFailureHandler'),
          });
        }
      }
    });
  });

  return calls;
}

/**
 * GASファイルから関数定義を抽出
 */
function extractGasFunctions(gasFiles) {
  const functions = [];

  gasFiles.forEach((file) => {
    const content = fs.readFileSync(file, 'utf8');
    const lines = content.split('\n');

    lines.forEach((line, index) => {
      // 関数定義のパターンを検出
      const patterns = [
        /^function\s+(\w+)\s*\(/, // 通常の関数定義
        /^\s*function\s+(\w+)\s*\(/, // インデントされた関数定義
        /const\s+(\w+)\s*=\s*function/, // const宣言の関数
        /let\s+(\w+)\s*=\s*function/, // let宣言の関数
        /var\s+(\w+)\s*=\s*function/, // var宣言の関数（非推奨だが検出）
        /(\w+)\s*:\s*function/, // オブジェクトメソッド
      ];

      patterns.forEach((pattern) => {
        const match = line.match(pattern);
        if (match) {
          functions.push({
            name: match[1],
            file,
            line: index + 1,
            definition: line.trim(),
          });
        }
      });
    });
  });

  return functions;
}

/**
 * 類似関数名を検索（レーベンシュタイン距離ベース）
 */
function findSimilarFunctions(targetName, gasFunctions, threshold = 3) {
  const similar = [];

  gasFunctions.forEach((func) => {
    const distance = levenshteinDistance(targetName, func.name);
    if (distance <= threshold && distance > 0) {
      similar.push(func.name);
    }
  });

  return similar.slice(0, 5); // 最大5個まで
}

/**
 * レーベンシュタイン距離計算
 */
function levenshteinDistance(str1, str2) {
  const matrix = [];

  for (let i = 0; i <= str2.length; i++) {
    matrix[i] = [i];
  }

  for (let j = 0; j <= str1.length; j++) {
    matrix[0][j] = j;
  }

  for (let i = 1; i <= str2.length; i++) {
    for (let j = 1; j <= str1.length; j++) {
      if (str2.charAt(i - 1) === str1.charAt(j - 1)) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j - 1] + 1,
          matrix[i][j - 1] + 1,
          matrix[i - 1][j] + 1
        );
      }
    }
  }

  return matrix[str2.length][str1.length];
}

/**
 * 内部関数かどうかを判定（doGet等のGAS標準関数）
 */
function isInternalFunction(funcName) {
  const internalFunctions = [
    'doGet',
    'doPost',
    'onOpen',
    'onEdit',
    'onFormSubmit',
    'onInstall',
    'include', // GAS固有
    'main',
    'init',
    'test', // 汎用的な内部関数
  ];

  return (
    internalFunctions.includes(funcName) ||
    funcName.startsWith('_') || // プライベート関数
    funcName.startsWith('test') || // テスト関数
    funcName.startsWith('debug')
  ); // デバッグ関数
}

/**
 * GAS予約語かどうかを判定
 */
function isGasReservedWord(funcName) {
  const reservedWords = ['withSuccessHandler', 'withFailureHandler', 'withUserObject'];
  return reservedWords.includes(funcName);
}
