/**
 * 重複宣言エラー検出テスト
 * コードベース内の定数や関数の重複宣言をチェック
 */

const fs = require('fs');
const path = require('path');

describe('重複宣言エラー検出', () => {
  let allGsFiles = [];
  const declarations = new Map(); // key: 識別子名, value: [{file, line, type}]

  beforeAll(() => {
    // 全.gsファイルを読み込み
    const srcDir = path.join(__dirname, '../src');
    const gasFiles = fs.readdirSync(srcDir).filter(file => file.endsWith('.gs'));
    
    allGsFiles = gasFiles.map(file => {
      const filePath = path.join(srcDir, file);
      const content = fs.readFileSync(filePath, 'utf8');
      return { fileName: file, content, filePath };
    });
    
    // 全ての宣言を収集
    collectAllDeclarations();
  });

  function collectAllDeclarations() {
    allGsFiles.forEach(file => {
      const lines = file.content.split('\n');
      
      lines.forEach((line, index) => {
        const lineNumber = index + 1;
        const trimmedLine = line.trim();
        
        // コメント行をスキップ
        if (trimmedLine.startsWith('//') || 
            trimmedLine.startsWith('/*') || 
            trimmedLine.startsWith('*') ||
            trimmedLine.length === 0) {
          return;
        }
        
        // function宣言
        const funcMatch = trimmedLine.match(/^(async\s+)?function\s+(\w+)/);
        if (funcMatch) {
          const funcName = funcMatch[2];
          addDeclaration(funcName, file.fileName, lineNumber, 'function');
          return;
        }
        
        // class宣言
        const classMatch = trimmedLine.match(/^class\s+(\w+)/);
        if (classMatch) {
          const className = classMatch[1];
          addDeclaration(className, file.fileName, lineNumber, 'class');
          return;
        }
        
        // const/let/var宣言
        const varMatch = trimmedLine.match(/^(const|let|var)\s+(\w+)/);
        if (varMatch) {
          const varName = varMatch[2];
          const varType = varMatch[1];
          // グローバルスコープの定数のみをチェック（関数内は除外）
          if (!isInsideFunction(lines, index)) {
            addDeclaration(varName, file.fileName, lineNumber, varType);
          }
        }
      });
    });
  }

  function isInsideFunction(lines, currentIndex) {
    // より単純なアプローチ：現在の行より前に関数宣言があり、
    // その関数がまだ閉じていない場合は関数内部と判定
    
    let depth = 0;
    let lastFunctionStartLine = -1;
    
    for (let i = 0; i <= currentIndex; i++) {
      const line = lines[i].trim();
      
      // コメント行はスキップ
      if (line.startsWith('//') || line.startsWith('/*') || line.startsWith('*')) {
        continue;
      }
      
      // 関数宣言を検出（function キーワードで始まる、またはアロー関数）
      if (line.match(/^(async\s+)?function\s+\w+/) || 
          line.match(/^\w+\s*[:=]\s*(async\s+)?function/) ||
          line.match(/^\w+\s*[:=]\s*\(.*\)\s*=>/)) {
        lastFunctionStartLine = i;
      }
      
      // クラスメソッドの検出（クラス内のメソッド定義）
      if (line.match(/^\w+\s*\(.*\)\s*\{/) && depth > 0) {
        lastFunctionStartLine = i;
      }
      
      // 波括弧のカウント
      for (let j = 0; j < line.length; j++) {
        if (line[j] === '{') {
          depth++;
        } else if (line[j] === '}') {
          depth--;
          // 関数が終了した可能性
          if (lastFunctionStartLine >= 0 && depth === 0) {
            lastFunctionStartLine = -1;
          }
        }
      }
    }
    
    // 関数宣言があり、まだ閉じていない場合は関数内部
    return lastFunctionStartLine >= 0 && depth > 0;
  }

  function addDeclaration(name, fileName, lineNumber, type) {
    if (!declarations.has(name)) {
      declarations.set(name, []);
    }
    declarations.get(name).push({ fileName, lineNumber, type });
  }

  test('定数の重複宣言をチェック', () => {
    const duplicateConstants = [];
    
    declarations.forEach((decls, name) => {
      // const/let/var の重複をチェック
      const constants = decls.filter(d => ['const', 'let', 'var'].includes(d.type));
      if (constants.length > 1) {
        // 異なるファイル間での重複のみを問題とする
        const uniqueFiles = new Set(constants.map(d => d.fileName));
        if (uniqueFiles.size > 1) {
          duplicateConstants.push({
            name,
            declarations: constants,
            files: Array.from(uniqueFiles)
          });
        }
      }
    });

    if (duplicateConstants.length > 0) {
      console.log('\n🚨 重複定数宣言を検出:');
      duplicateConstants.forEach(dup => {
        console.log(`\n📍 ${dup.name}:`);
        dup.declarations.forEach(decl => {
          console.log(`   ${decl.type} in ${decl.fileName}:${decl.lineNumber}`);
        });
      });
      console.log(`\n合計: ${duplicateConstants.length}個の重複定数`);
    } else {
      console.log('\n✅ 重複定数宣言なし');
    }

    // 重複があればテスト失敗
    expect(duplicateConstants).toHaveLength(0);
  });

  test('関数の重複宣言をチェック', () => {
    const duplicateFunctions = [];
    
    declarations.forEach((decls, name) => {
      // function の重複をチェック  
      const functions = decls.filter(d => d.type === 'function');
      if (functions.length > 1) {
        // 異なるファイル間での重複のみを問題とする
        const uniqueFiles = new Set(functions.map(d => d.fileName));
        if (uniqueFiles.size > 1) {
          duplicateFunctions.push({
            name,
            declarations: functions,
            files: Array.from(uniqueFiles)
          });
        }
      }
    });

    if (duplicateFunctions.length > 0) {
      console.log('\n🚨 重複関数宣言を検出:');
      duplicateFunctions.forEach(dup => {
        console.log(`\n📍 ${dup.name}:`);
        dup.declarations.forEach(decl => {
          console.log(`   ${decl.type} in ${decl.fileName}:${decl.lineNumber}`);
        });
      });
      console.log(`\n合計: ${duplicateFunctions.length}個の重複関数`);
    } else {
      console.log('\n✅ 重複関数宣言なし');
    }

    // 重複があればテスト失敗
    expect(duplicateFunctions).toHaveLength(0);
  });

  test('統計情報を表示', () => {
    const stats = {
      totalFiles: allGsFiles.length,
      totalDeclarations: declarations.size,
      functions: 0,
      classes: 0,
      constants: 0,
      variables: 0
    };
    
    declarations.forEach(decls => {
      decls.forEach(decl => {
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

    // 統計情報は常に成功
    expect(stats.totalFiles).toBeGreaterThan(0);
  });

  test('重要な定数が適切に定義されているかチェック', () => {
    const expectedConstants = [
      'COLUMN_HEADERS',
      'REACTION_KEYS',
      'USER_CACHE_TTL',
      'DB_BATCH_SIZE',
      'DISPLAY_MODES'
    ];

    const missingConstants = [];
    const properlyDefined = [];

    expectedConstants.forEach(constName => {
      if (declarations.has(constName)) {
        const decls = declarations.get(constName).filter(d => d.type === 'const');
        if (decls.length === 1) {
          properlyDefined.push({
            name: constName,
            file: decls[0].fileName,
            line: decls[0].lineNumber
          });
        }
      } else {
        missingConstants.push(constName);
      }
    });

    console.log('\n📝 重要な定数の定義状況:');
    properlyDefined.forEach(c => {
      console.log(`  ✅ ${c.name} - ${c.file}:${c.line}`);
    });
    
    if (missingConstants.length > 0) {
      console.log('\n❌ 未定義の定数:');
      missingConstants.forEach(name => console.log(`  - ${name}`));
    }

    // 重要な定数がすべて定義されていることを確認
    expect(missingConstants).toHaveLength(0);
  });
});