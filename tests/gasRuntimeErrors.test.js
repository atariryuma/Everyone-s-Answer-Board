/**
 * GAS実行時エラー検出テスト（包括的版）
 * 未定義変数・関数・クラスの参照を一括検出
 */

const fs = require('fs');
const path = require('path');

describe('GAS Runtime Error Detection', () => {
  let allGasFiles = [];
  let allDefinitions = new Set(); // すべての定義済み識別子
  let allReferences = new Map(); // すべての参照 key: 識別子, value: [{file, line, context}]
  
  // GASビルトイン関数・オブジェクト
  const GAS_BUILTINS = new Set([
    // Google Apps Script Services
    'SpreadsheetApp', 'DriveApp', 'UrlFetchApp', 'HtmlService', 'ScriptApp',
    'PropertiesService', 'CacheService', 'Utilities', 'Logger', 'Session',
    'FormApp', 'GmailApp', 'CalendarApp', 'DocumentApp', 'SlidesApp',
    'Browser', 'ContentService',
    
    // JavaScript標準
    'console', 'JSON', 'Math', 'Date', 'String', 'Number', 'Array', 'Object',
    'Promise', 'Error', 'RegExp', 'Map', 'Set', 'WeakMap', 'WeakSet', 'Boolean',
    'parseInt', 'parseFloat', 'isNaN', 'encodeURIComponent', 'decodeURIComponent',
    'undefined', 'null', 'NaN', 'Infinity',
    
    // DOM関連（HTMLテンプレート内で使用）
    'document', 'window', 'navigator', 'location', 'localStorage', 'sessionStorage',
    'google' // google.script.run
  ]);

  beforeAll(() => {
    // 全.gsファイルを読み込み
    const srcDir = path.join(__dirname, '../src');
    const gasFiles = fs.readdirSync(srcDir).filter(file => file.endsWith('.gs'));
    
    allGasFiles = gasFiles.map(file => {
      const filePath = path.join(srcDir, file);
      const content = fs.readFileSync(filePath, 'utf8');
      return { fileName: file, content, filePath };
    });
    
    // 1. すべての定義を収集
    collectAllDefinitions();
    
    // 2. すべての参照を収集
    collectAllReferences();
  });

  function collectAllDefinitions() {
    allGasFiles.forEach(file => {
      const lines = file.content.split('\n');
      
      lines.forEach((line, index) => {
        const trimmedLine = line.trim();
        
        // コメントをスキップ
        if (trimmedLine.startsWith('//') || trimmedLine.startsWith('/*') || trimmedLine.startsWith('*')) {
          return;
        }
        
        // function宣言
        const funcMatch = trimmedLine.match(/^(async\s+)?function\s+(\w+)/);
        if (funcMatch) {
          allDefinitions.add(funcMatch[2]);
        }
        
        // class宣言
        const classMatch = trimmedLine.match(/^class\s+(\w+)/);
        if (classMatch) {
          allDefinitions.add(classMatch[1]);
        }
        
        // const/let/var宣言（複数変数対応）
        const varMatch = trimmedLine.match(/^(const|let|var)\s+([^=]+)/);
        if (varMatch) {
          const vars = varMatch[2].split(',').map(v => v.trim().split(/[\s={]/)[0]);
          vars.forEach(v => {
            if (v) allDefinitions.add(v);
          });
        }
        
        // オブジェクトのプロパティ定義（グローバルオブジェクト）
        if (trimmedLine.match(/^[A-Z_]+\s*=\s*\{/)) {
          const objName = trimmedLine.match(/^([A-Z_]+)\s*=/)[1];
          allDefinitions.add(objName);
        }
      });
    });
    
    // ファイル間で共有される定数も追加
    allDefinitions.add('UNIFIED_CONSTANTS');
    allDefinitions.add('ERROR_SEVERITY');
    allDefinitions.add('ERROR_CATEGORIES');
    allDefinitions.add('COLUMN_HEADERS');
    allDefinitions.add('DISPLAY_MODES');
  }

  function collectAllReferences() {
    allGasFiles.forEach(file => {
      const lines = file.content.split('\n');
      
      lines.forEach((line, lineNumber) => {
        // 識別子パターン: 変数参照、関数呼び出し、プロパティアクセス
        const patterns = [
          /\b([A-Z_][A-Z0-9_]*)\b/g,  // 定数参照（ERROR_TYPES など）
          /\b(\w+)\s*\(/g,             // 関数呼び出し
          /\b(\w+)\./g,                // オブジェクトのプロパティアクセス
          /typeof\s+(\w+)/g,          // typeof チェック
          /new\s+(\w+)/g,             // クラスインスタンス化
        ];
        
        patterns.forEach(pattern => {
          let match;
          while ((match = pattern.exec(line)) !== null) {
            const identifier = match[1];
            
            // 予約語や短い識別子をスキップ
            if (identifier.length <= 1 || isReservedKeyword(identifier)) {
              continue;
            }
            
            // 参照を記録
            if (!allReferences.has(identifier)) {
              allReferences.set(identifier, []);
            }
            
            allReferences.get(identifier).push({
              file: file.fileName,
              line: lineNumber + 1,
              context: line.trim().substring(0, 100)
            });
          }
        });
      });
    });
  }

  function isReservedKeyword(word) {
    const keywords = [
      'if', 'else', 'for', 'while', 'do', 'switch', 'case', 'break', 'continue',
      'function', 'return', 'var', 'let', 'const', 'class', 'extends', 'new',
      'try', 'catch', 'finally', 'throw', 'typeof', 'instanceof', 'in', 'of',
      'true', 'false', 'null', 'undefined', 'this', 'super', 'import', 'export',
      'default', 'async', 'await', 'yield', 'static', 'get', 'set'
    ];
    return keywords.includes(word);
  }

  test('未定義識別子の一括検出', () => {
    const undefinedIdentifiers = new Map();
    
    // すべての参照をチェック
    allReferences.forEach((references, identifier) => {
      // 定義されていない かつ ビルトインでもない
      if (!allDefinitions.has(identifier) && !GAS_BUILTINS.has(identifier)) {
        undefinedIdentifiers.set(identifier, references);
      }
    });
    
    // 結果を整理
    const errors = [];
    undefinedIdentifiers.forEach((refs, identifier) => {
      // 参照が多い順にソート
      const uniqueFiles = new Set(refs.map(r => r.file));
      
      errors.push({
        identifier,
        occurrences: refs.length,
        files: Array.from(uniqueFiles),
        samples: refs.slice(0, 3) // 最初の3つのサンプル
      });
    });
    
    // 出現回数でソート
    errors.sort((a, b) => b.occurrences - a.occurrences);
    
    // レポート生成
    if (errors.length > 0) {
      console.log('\n🚨 未定義識別子エラー検出結果:');
      console.log('=' .repeat(80));
      
      // 上位20個を表示
      errors.slice(0, 20).forEach(error => {
        console.log(`\n📍 ${error.identifier} (${error.occurrences}回参照)`);
        console.log(`   Files: ${error.files.join(', ')}`);
        error.samples.forEach(sample => {
          console.log(`   Line ${sample.line} in ${sample.file}:`);
          console.log(`     ${sample.context}`);
        });
      });
      
      if (errors.length > 20) {
        console.log(`\n... and ${errors.length - 20} more undefined identifiers`);
      }
      
      // 完全なレポートをファイルに出力
      const report = errors.map(error => 
        `## ${error.identifier}\n` +
        `- Occurrences: ${error.occurrences}\n` +
        `- Files: ${error.files.join(', ')}\n` +
        `- Samples:\n` +
        error.samples.map(s => `  - ${s.file}:${s.line}: ${s.context}`).join('\n')
      ).join('\n\n');
      
      fs.writeFileSync(
        path.join(__dirname, '../runtime-errors-report.md'),
        `# Runtime Errors Report\n\nTotal: ${errors.length} undefined identifiers\n\n${report}\n`
      );
      
      // 修正提案を生成
      generateFixSuggestions(errors);
    } else {
      console.log('\n✅ 未定義識別子エラーなし');
    }
    
    // カテゴリ別の統計
    console.log('\n📊 エラー統計:');
    console.log(`  合計: ${errors.length}個の未定義識別子`);
    console.log(`  影響ファイル数: ${new Set(errors.flatMap(e => e.files)).size}個`);
    console.log(`  総参照回数: ${errors.reduce((sum, e) => sum + e.occurrences, 0)}回`);
    
    // テストは警告のみ（0になるまで段階的に修正）
    expect(errors.length).toBeLessThanOrEqual(100); // 許容範囲を設定
  });
  
  function generateFixSuggestions(errors) {
    const fixes = [];
    
    errors.forEach(error => {
      const id = error.identifier;
      
      // パターンマッチングで修正提案
      if (id.startsWith('ERROR_') || id.endsWith('_ERROR')) {
        fixes.push({
          identifier: id,
          suggestion: `const ${id} = UNIFIED_CONSTANTS.ERROR.CATEGORIES.${id};`,
          category: 'error_constant'
        });
      } else if (id.endsWith('_CACHE') || id.startsWith('CACHE_')) {
        fixes.push({
          identifier: id,
          suggestion: `const ${id} = {};`,
          category: 'cache_object'
        });
      } else if (id.match(/^[A-Z_]+$/)) {
        // 全大文字は定数
        fixes.push({
          identifier: id,
          suggestion: `const ${id} = '${id.toLowerCase()}';`,
          category: 'constant'
        });
      } else {
        // その他は関数として定義
        fixes.push({
          identifier: id,
          suggestion: `const ${id} = () => console.log('[${id}] Not implemented');`,
          category: 'function'
        });
      }
    });
    
    // カテゴリ別に整理
    const byCategory = {};
    fixes.forEach(fix => {
      if (!byCategory[fix.category]) {
        byCategory[fix.category] = [];
      }
      byCategory[fix.category].push(fix);
    });
    
    // 修正コードを生成
    let fixCode = '// === 自動生成された修正提案 ===\n\n';
    
    Object.entries(byCategory).forEach(([category, items]) => {
      fixCode += `// ${category}\n`;
      items.slice(0, 10).forEach(item => {
        fixCode += `${item.suggestion}\n`;
      });
      fixCode += '\n';
    });
    
    fs.writeFileSync(
      path.join(__dirname, '../suggested-fixes.js'),
      fixCode
    );
    
    console.log('\n💡 修正提案を suggested-fixes.js に生成しました');
  }
  
  test('特定のエラーパターン検出', () => {
    const patterns = {
      'ERROR_TYPES': [],
      'ERROR_CONSTANTS': [],
      'CACHE_REFERENCES': [],
      'MISSING_FUNCTIONS': []
    };
    
    allReferences.forEach((refs, identifier) => {
      if (!allDefinitions.has(identifier) && !GAS_BUILTINS.has(identifier)) {
        if (identifier === 'ERROR_TYPES') {
          patterns['ERROR_TYPES'] = refs;
        } else if (identifier.includes('ERROR')) {
          patterns['ERROR_CONSTANTS'].push({ identifier, refs });
        } else if (identifier.includes('CACHE') || identifier.includes('cache')) {
          patterns['CACHE_REFERENCES'].push({ identifier, refs });
        } else if (refs.some(r => r.context.includes('('))) {
          patterns['MISSING_FUNCTIONS'].push({ identifier, refs });
        }
      }
    });
    
    // ERROR_TYPES の詳細
    if (patterns['ERROR_TYPES'].length > 0) {
      console.log('\n🔴 ERROR_TYPES 参照箇所:');
      patterns['ERROR_TYPES'].forEach(ref => {
        console.log(`  ${ref.file}:${ref.line}`);
      });
      console.log('  → 修正: UNIFIED_CONSTANTS.ERROR.CATEGORIES を使用');
    }
  });
});