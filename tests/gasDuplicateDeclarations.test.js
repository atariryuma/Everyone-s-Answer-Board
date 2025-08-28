/**
 * GAS重複宣言エラー検出テスト（包括的版）
 * 全ての.gsファイル間での重複宣言を検出
 */

const fs = require('fs');
const path = require('path');

describe('GAS Duplicate Declarations Test', () => {
  let allGasFiles = [];
  let globalDeclarations = new Map(); // key: 識別子名, value: [{file, line, type}]

  beforeAll(() => {
    // 全.gsファイルを読み込み
    const srcDir = path.join(__dirname, '../src');
    const gasFiles = fs.readdirSync(srcDir).filter(file => file.endsWith('.gs'));
    
    allGasFiles = gasFiles.map(file => {
      const filePath = path.join(srcDir, file);
      const content = fs.readFileSync(filePath, 'utf8');
      return { fileName: file, content, filePath };
    });
    
    // すべてのグローバル宣言を収集
    collectAllDeclarations();
  });

  function collectAllDeclarations() {
    allGasFiles.forEach(file => {
      const lines = file.content.split('\n');
      
      lines.forEach((line, index) => {
        const lineNumber = index + 1;
        const trimmedLine = line.trim();
        
        // コメント行をスキップ
        if (trimmedLine.startsWith('//') || trimmedLine.startsWith('/*') || trimmedLine.startsWith('*')) {
          return;
        }
        
        // function宣言
        const funcMatch = trimmedLine.match(/^(async\s+)?function\s+(\w+)/);
        if (funcMatch) {
          const funcName = funcMatch[2];
          addDeclaration(funcName, file.fileName, lineNumber, 'function');
        }
        
        // class宣言
        const classMatch = trimmedLine.match(/^class\s+(\w+)/);
        if (classMatch) {
          const className = classMatch[1];
          addDeclaration(className, file.fileName, lineNumber, 'class');
        }
        
        // const/let/var宣言
        const varMatch = trimmedLine.match(/^(const|let|var)\s+(\w+)/);
        if (varMatch) {
          const varName = varMatch[2];
          const varType = varMatch[1];
          addDeclaration(varName, file.fileName, lineNumber, varType);
        }
      });
    });
  }

  function addDeclaration(name, fileName, lineNumber, type) {
    if (!globalDeclarations.has(name)) {
      globalDeclarations.set(name, []);
    }
    globalDeclarations.get(name).push({ fileName, lineNumber, type });
  }

  test('グローバルスコープの重複宣言をチェック', () => {
    const duplicates = [];
    
    // 重複を検出
    globalDeclarations.forEach((declarations, name) => {
      if (declarations.length > 1) {
        // 同じファイル内の重複は除外（スコープが異なる可能性）
        const uniqueFiles = new Set(declarations.map(d => d.fileName));
        if (uniqueFiles.size > 1) {
          duplicates.push({
            name,
            declarations: declarations
          });
        }
      }
    });
    
    // 結果を表示
    if (duplicates.length > 0) {
      console.log('\n🚨 重複宣言エラーを検出:');
      console.log('=' .repeat(60));
      
      duplicates.forEach(dup => {
        console.log(`\n📍 識別子: ${dup.name}`);
        dup.declarations.forEach(decl => {
          console.log(`   ${decl.type.padEnd(8)} in ${decl.fileName}:${decl.lineNumber}`);
        });
      });
      
      console.log('\n' + '=' .repeat(60));
      console.log(`合計: ${duplicates.length}個の重複宣言\n`);
      
      // 詳細レポートを生成
      const report = duplicates.map(dup => {
        const locations = dup.declarations.map(d => 
          `  - ${d.type} in ${d.fileName}:${d.lineNumber}`
        ).join('\n');
        return `## ${dup.name}\n${locations}`;
      }).join('\n\n');
      
      fs.writeFileSync(
        path.join(__dirname, '../duplicate-declarations-report.md'),
        `# 重複宣言レポート\n\n${report}\n`
      );
    } else {
      console.log('\n✅ 重複宣言エラーなし\n');
    }
    
    expect(duplicates.length).toBe(0);
  });

  test('特定の問題パターンをチェック', () => {
    const issues = [];
    
    // よくある問題パターンを検出
    globalDeclarations.forEach((declarations, name) => {
      declarations.forEach(decl => {
        // クラスと同名の関数
        const hasClass = declarations.some(d => d.type === 'class');
        const hasFunction = declarations.some(d => d.type === 'function');
        if (hasClass && hasFunction && declarations.length > 1) {
          issues.push({
            type: 'class-function-conflict',
            name,
            message: `クラス ${name} と同名の関数が存在`
          });
        }
        
        // const と function の競合
        const hasConst = declarations.some(d => d.type === 'const');
        if (hasConst && hasFunction && declarations.length > 1) {
          issues.push({
            type: 'const-function-conflict',
            name,
            message: `const ${name} と function ${name} が競合`
          });
        }
      });
    });
    
    // 重複を除去
    const uniqueIssues = Array.from(
      new Map(issues.map(item => [`${item.type}-${item.name}`, item])).values()
    );
    
    if (uniqueIssues.length > 0) {
      console.log('\n⚠️ 問題のあるパターンを検出:');
      uniqueIssues.forEach(issue => {
        console.log(`  - ${issue.message}`);
      });
    }
  });

  test('予約語・ビルトイン関数との競合をチェック', () => {
    const gasBuiltins = [
      'SpreadsheetApp', 'DriveApp', 'UrlFetchApp', 'HtmlService', 'ScriptApp',
      'PropertiesService', 'CacheService', 'Utilities', 'Logger', 'Session',
      'FormApp', 'GmailApp', 'CalendarApp', 'DocumentApp', 'SlidesApp'
    ];
    
    const conflicts = [];
    
    globalDeclarations.forEach((declarations, name) => {
      if (gasBuiltins.includes(name)) {
        conflicts.push({
          name,
          type: 'GAS built-in override',
          declarations
        });
      }
    });
    
    if (conflicts.length > 0) {
      console.log('\n⚠️ GASビルトインとの競合:');
      conflicts.forEach(conflict => {
        console.log(`  - ${conflict.name}: ${conflict.type}`);
      });
    }
  });

  test('統計情報を表示', () => {
    const stats = {
      totalFiles: allGasFiles.length,
      totalDeclarations: globalDeclarations.size,
      functions: 0,
      classes: 0,
      constants: 0,
      variables: 0
    };
    
    globalDeclarations.forEach(declarations => {
      declarations.forEach(decl => {
        switch(decl.type) {
          case 'function': stats.functions++; break;
          case 'class': stats.classes++; break;
          case 'const': stats.constants++; break;
          case 'let':
          case 'var': stats.variables++; break;
        }
      });
    });
    
    console.log('\n📊 宣言統計:');
    console.log(`  ファイル数: ${stats.totalFiles}`);
    console.log(`  識別子数: ${stats.totalDeclarations}`);
    console.log(`  関数: ${stats.functions}個`);
    console.log(`  クラス: ${stats.classes}個`);
    console.log(`  定数: ${stats.constants}個`);
    console.log(`  変数: ${stats.variables}個`);
  });
});