/**
 * システム全体の汎用名への一括リネーミングスクリプト
 * プログラミング的アプローチで漏れのない変更を実現
 */

// =============================================================================
// リネーミング設定とマッピング
// =============================================================================

const RENAMING_MAP = {
  // 関数名の統一
  getCurrentUserInfo: 'getActiveUserInfo',
  getUserInfo: 'getActiveUserInfo',

  // 列判定系の統一
  identifyHeaders: 'identifyHeadersAdvanced', // 新しいAI強化版に統一
  mapColumns: 'detectColumnMapping',

  // その他の汎用化（必要に応じて追加）
  connectToDataSource: 'connectDataSource',
  analyzeColumns: 'analyzeColumnMapping',
};

const FILES_TO_PROCESS = [
  'src/main.gs',
  'src/Core.gs',
  'src/AdminPanelBackend.gs',
  'src/AdminPanel.js.html',
  'src/auth.gs',
];

// =============================================================================
// リネーミング分析・実行機能
// =============================================================================

/**
 * システム全体のリネーミング分析
 * 実行前に全参照箇所を洗い出し
 */
function analyzeSystemwideRenaming() {
  console.log('🔍 システム全体のリネーミング分析開始');

  const results = {
    processed: [],
    errors: [],
    summary: {},
    totalReferences: 0,
  };

  // 各リネーミング対象の分析
  Object.entries(RENAMING_MAP).forEach(([oldName, newName]) => {
    console.log(`分析中: ${oldName} -> ${newName}`);

    const references = findAllReferences(oldName);
    results.summary[oldName] = {
      newName,
      referenceCount: references.length,
      files: [...new Set(references.map((r) => r.file))],
      references: references,
    };

    results.totalReferences += references.length;
    console.log(`  発見: ${references.length}箇所の参照`);
  });

  // 分析結果のサマリー出力
  console.log('📊 リネーミング分析結果:');
  console.log(`  総参照数: ${results.totalReferences}`);
  console.log(`  対象ファイル数: ${FILES_TO_PROCESS.length}`);

  return results;
}

/**
 * 指定された関数名の全参照箇所を検索
 * grep相当の機能をGASで実現
 */
function findAllReferences(functionName) {
  const references = [];

  // 正規表現パターンの生成
  const patterns = [
    new RegExp(`\\b${functionName}\\s*\\(`, 'g'), // 関数呼び出し
    new RegExp(`function\\s+${functionName}\\b`, 'g'), // 関数定義
    new RegExp(`const\\s+${functionName}\\s*=`, 'g'), // const定義
    new RegExp(`let\\s+${functionName}\\s*=`, 'g'), // let定義
    new RegExp(`\\.${functionName}\\b`, 'g'), // メソッド呼び出し
  ];

  // 注意：実際のファイル読み込みはGAS環境では制限があるため、
  // この例では概念的な実装を示す
  FILES_TO_PROCESS.forEach((filePath) => {
    try {
      // ファイル内容の仮想読み込み（実装時は適切なファイル読み込み方法を使用）
      const content = getFileContent(filePath);

      patterns.forEach((pattern) => {
        let match;
        while ((match = pattern.exec(content)) !== null) {
          references.push({
            file: filePath,
            line: getLineNumber(content, match.index),
            column: match.index,
            context: getContextLine(content, match.index),
            pattern: pattern.source,
          });
        }
      });
    } catch (error) {
      console.error(`ファイル処理エラー: ${filePath}`, error);
    }
  });

  return references;
}

/**
 * 段階的リネーミング実行
 */
function executeSystemwideRenaming(dryRun = true) {
  console.log(`🚀 ${dryRun ? '[ドライラン]' : '[実行]'} システム全体リネーミング開始`);

  const results = {
    success: [],
    errors: [],
    skipped: [],
  };

  // 段階的に各リネーミングを実行
  Object.entries(RENAMING_MAP).forEach(([oldName, newName]) => {
    try {
      const result = executeRenaming(oldName, newName, dryRun);

      if (result.success) {
        results.success.push({
          oldName,
          newName,
          changesCount: result.changes.length,
          files: result.files,
        });
      } else {
        results.errors.push({
          oldName,
          newName,
          error: result.error,
        });
      }
    } catch (error) {
      console.error(`リネーミングエラー: ${oldName} -> ${newName}`, error);
      results.errors.push({
        oldName,
        newName,
        error: error.message,
      });
    }
  });

  // 結果のサマリー
  console.log('📋 リネーミング結果:');
  console.log(`  成功: ${results.success.length}件`);
  console.log(`  エラー: ${results.errors.length}件`);
  console.log(`  スキップ: ${results.skipped.length}件`);

  return results;
}

/**
 * 個別のリネーミング実行
 */
function executeRenaming(oldName, newName, dryRun = true) {
  console.log(`${dryRun ? '[ドライラン]' : '[実行]'} ${oldName} -> ${newName}`);

  const changes = [];
  const affectedFiles = [];

  FILES_TO_PROCESS.forEach((filePath) => {
    try {
      const originalContent = getFileContent(filePath);
      const newContent = performRenaming(originalContent, oldName, newName);

      if (originalContent !== newContent) {
        changes.push({
          file: filePath,
          oldContent: originalContent,
          newContent: newContent,
          changeCount: countChanges(originalContent, newContent, oldName),
        });

        affectedFiles.push(filePath);

        if (!dryRun) {
          // 実際のファイル書き込み（実装時に適切な書き込み方法を使用）
          writeFileContent(filePath, newContent);
          console.log(`✅ 更新完了: ${filePath}`);
        } else {
          console.log(`📝 変更予定: ${filePath} (${changes[changes.length - 1].changeCount}箇所)`);
        }
      }
    } catch (error) {
      console.error(`ファイル処理エラー: ${filePath}`, error);
      return { success: false, error: error.message };
    }
  });

  return {
    success: true,
    changes: changes,
    files: affectedFiles,
  };
}

/**
 * 文字列内でのリネーミング実行
 */
function performRenaming(content, oldName, newName) {
  // 安全なリネーミングのための複数パターン対応
  const replacementPatterns = [
    {
      // 関数呼び出し: functionName( -> newName(
      pattern: new RegExp(`\\b${escapeRegExp(oldName)}\\s*\\(`, 'g'),
      replacement: `${newName}(`,
    },
    {
      // 関数定義: function functionName -> function newName
      pattern: new RegExp(`function\\s+${escapeRegExp(oldName)}\\b`, 'g'),
      replacement: `function ${newName}`,
    },
    {
      // 変数定義: const/let functionName = -> const/let newName =
      pattern: new RegExp(`(const|let)\\s+${escapeRegExp(oldName)}\\s*=`, 'g'),
      replacement: `$1 ${newName} =`,
    },
    {
      // メソッド呼び出し: .functionName( -> .newName(
      pattern: new RegExp(`\\.${escapeRegExp(oldName)}\\s*\\(`, 'g'),
      replacement: `.${newName}(`,
    },
  ];

  let newContent = content;

  replacementPatterns.forEach(({ pattern, replacement }) => {
    newContent = newContent.replace(pattern, replacement);
  });

  return newContent;
}

/**
 * リネーミング後の整合性チェック
 */
function validateRenamingResults() {
  console.log('🔍 リネーミング結果の整合性チェック開始');

  const validationResults = {
    missingReferences: [],
    undefinedFunctions: [],
    syntaxErrors: [],
    success: true,
  };

  // 1. 未定義関数のチェック
  Object.values(RENAMING_MAP).forEach((newName) => {
    const references = findAllReferences(newName);
    if (references.length === 0) {
      validationResults.missingReferences.push(newName);
      validationResults.success = false;
    }
  });

  // 2. 旧関数名の残存チェック
  Object.keys(RENAMING_MAP).forEach((oldName) => {
    const references = findAllReferences(oldName);
    if (references.length > 0) {
      validationResults.undefinedFunctions.push({
        oldName,
        remainingReferences: references.length,
      });
      validationResults.success = false;
    }
  });

  console.log('📋 整合性チェック結果:');
  console.log(`  未定義関数: ${validationResults.undefinedFunctions.length}件`);
  console.log(`  参照不足: ${validationResults.missingReferences.length}件`);
  console.log(`  総合判定: ${validationResults.success ? '✅ 成功' : '❌ 要修正'}`);

  return validationResults;
}

// =============================================================================
// ヘルパー関数群
// =============================================================================

/**
 * 正規表現用エスケープ
 */
function escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\\]\\\\]/g, '\\\\$&');
}

/**
 * ファイル内容取得（概念的実装）
 */
function getFileContent(filePath) {
  // 実装時は適切なファイル読み込み方法を使用
  // GAS環境では DriveApp や直接的なファイルアクセスを利用
  console.log(`ファイル読み込み: ${filePath}`);
  return '// ファイル内容のプレースホルダー';
}

/**
 * ファイル内容書き込み（概念的実装）
 */
function writeFileContent(filePath, content) {
  // 実装時は適切なファイル書き込み方法を使用
  console.log(`ファイル書き込み: ${filePath}`);
}

/**
 * 行番号取得
 */
function getLineNumber(content, index) {
  return content.substring(0, index).split('\\n').length;
}

/**
 * 文脈行取得
 */
function getContextLine(content, index) {
  const lines = content.split('\\n');
  const lineNumber = getLineNumber(content, index);
  return lines[lineNumber - 1] || '';
}

/**
 * 変更箇所数カウント
 */
function countChanges(oldContent, newContent, searchTerm) {
  const oldMatches = (oldContent.match(new RegExp(escapeRegExp(searchTerm), 'g')) || []).length;
  const newMatches = (newContent.match(new RegExp(escapeRegExp(searchTerm), 'g')) || []).length;
  return oldMatches - newMatches;
}

// =============================================================================
// 実行用エントリーポイント
// =============================================================================

/**
 * リネーミングスクリプトのメイン実行関数
 */
function runRenamingScript() {
  console.log('🚀 システム全体リネーミングスクリプト開始');

  try {
    // Step 1: 分析実行
    const analysis = analyzeSystemwideRenaming();

    // Step 2: ドライラン実行
    console.log('\\n--- ドライラン実行 ---');
    const dryRunResults = executeSystemwideRenaming(true);

    // Step 3: 実行可否の確認（手動判断）
    console.log('\\n--- 実行準備完了 ---');
    console.log('ドライラン結果を確認し、問題がなければ以下を実行:');
    console.log('executeSystemwideRenaming(false); // 実際のリネーミング実行');
    console.log('validateRenamingResults(); // 整合性チェック');

    return {
      analysis,
      dryRunResults,
      readyForExecution: dryRunResults.errors.length === 0,
    };
  } catch (error) {
    console.error('リネーミングスクリプト実行エラー:', error);
    return { error: error.message };
  }
}
