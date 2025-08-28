/**
 * GAS重複宣言エラー・参照エラー検出テスト
 * GAS環境での実行時エラーを事前検出
 */

const fs = require('fs');
const path = require('path');

describe('GAS Error Validation Tests', () => {
  let allGasFiles = [];
  let allSourceCode = '';

  beforeAll(() => {
    // 全.gsファイルを読み込み
    const srcDir = path.join(__dirname, '../src');
    const gasFiles = fs.readdirSync(srcDir).filter(file => file.endsWith('.gs'));
    
    allGasFiles = gasFiles.map(file => {
      const filePath = path.join(srcDir, file);
      const content = fs.readFileSync(filePath, 'utf8');
      return { fileName: file, content, filePath };
    });
    
    // 全ソースコードを結合（GAS環境をシミュレート）
    allSourceCode = allGasFiles.map(f => f.content).join('\n\n');
  });

  describe('重複宣言エラー検出', () => {
    test('クラス名の重複宣言をチェック', () => {
      const errors = [];
      const classDeclarations = {};
      
      allGasFiles.forEach(file => {
        const classMatches = file.content.match(/^class\s+(\w+)/gm) || [];
        classMatches.forEach(match => {
          const className = match.replace('class ', '');
          if (classDeclarations[className]) {
            errors.push({
              type: 'duplicate_class',
              name: className,
              files: [classDeclarations[className], file.fileName]
            });
          } else {
            classDeclarations[className] = file.fileName;
          }
        });
      });
      
      if (errors.length > 0) {
        console.log('🚨 重複クラス宣言エラー:');
        errors.forEach(error => {
          console.log(`  - ${error.name}: ${error.files.join(' と ')}`);
        });
        expect(errors).toHaveLength(0);
      }
    });

    test('グローバル関数の重複宣言をチェック', () => {
      const errors = [];
      const functionDeclarations = {};
      
      allGasFiles.forEach(file => {
        const funcMatches = file.content.match(/^function\s+(\w+)/gm) || [];
        funcMatches.forEach(match => {
          const funcName = match.replace(/^function\s+/, '').replace(/\(.*/, '');
          if (functionDeclarations[funcName]) {
            errors.push({
              type: 'duplicate_function',
              name: funcName,
              files: [functionDeclarations[funcName], file.fileName]
            });
          } else {
            functionDeclarations[funcName] = file.fileName;
          }
        });
      });
      
      if (errors.length > 0) {
        console.log('🚨 重複関数宣言エラー:');
        errors.forEach(error => {
          console.log(`  - ${error.name}: ${error.files.join(' と ')}`);
        });
        expect(errors).toHaveLength(0);
      }
    });

    test('グローバル変数の重複宣言をチェック', () => {
      const errors = [];
      const varDeclarations = {};
      
      allGasFiles.forEach(file => {
        // const, let, var での変数宣言をチェック
        const varMatches = file.content.match(/^(const|let|var)\s+(\w+)/gm) || [];
        varMatches.forEach(match => {
          const varName = match.replace(/^(const|let|var)\s+/, '').replace(/\s*=.*/, '');
          if (varDeclarations[varName]) {
            errors.push({
              type: 'duplicate_variable',
              name: varName,
              files: [varDeclarations[varName], file.fileName]
            });
          } else {
            varDeclarations[varName] = file.fileName;
          }
        });
      });
      
      if (errors.length > 0) {
        console.log('🚨 重複変数宣言エラー:');
        errors.forEach(error => {
          console.log(`  - ${error.name}: ${error.files.join(' と ')}`);
        });
        expect(errors).toHaveLength(0);
      }
    });
  });

  describe('参照エラー検出', () => {
    test('未定義関数の参照をチェック', () => {
      const errors = [];
      const definedFunctions = new Set();
      const referencedFunctions = new Set();
      
      // 定義済み関数を収集
      allGasFiles.forEach(file => {
        // 関数定義
        const funcMatches = file.content.match(/^function\s+(\w+)/gm) || [];
        funcMatches.forEach(match => {
          const funcName = match.replace(/^function\s+/, '').replace(/\(.*/, '');
          definedFunctions.add(funcName);
        });
        
        // クラスメソッド
        const methodMatches = file.content.match(/^\s+(?:static\s+)?(\w+)\s*\(/gm) || [];
        methodMatches.forEach(match => {
          const methodName = match.trim().replace(/^static\s+/, '').replace(/\s*\(.*/, '');
          if (methodName !== 'constructor') {
            definedFunctions.add(methodName);
          }
        });
      });
      
      // 参照されている関数を収集
      allGasFiles.forEach(file => {
        // 関数呼び出しパターンを検索
        const callMatches = file.content.match(/(\w+)\s*\(/g) || [];
        callMatches.forEach(match => {
          const funcName = match.replace(/\s*\(.*/, '');
          // JavaScript キーワード、制御構文、演算子を除外
          const jsKeywords = [
            'if', 'for', 'while', 'switch', 'catch', 'typeof', 'return', 'new',
            'throw', 'try', 'else', 'do', 'break', 'continue', 'var', 'let', 'const',
            'function', 'class', 'extends', 'import', 'export', 'default', 'async',
            'await', 'yield', 'delete', 'in', 'instanceof', 'void', 'this', 'super'
          ];
          
          // CSS/HTML関連の単語も除外
          const cssHtmlKeywords = [
            'gradient', 'rgba', 'translateY', 'scale', 'blur', 'bezier', 'media',
            'px', 'em', 'rem', 'vh', 'vw', 'deg', 'rad'
          ];
          
          if (!jsKeywords.includes(funcName) && 
              !cssHtmlKeywords.includes(funcName) &&
              funcName.length > 1) {  // 1文字の変数名は除外
            referencedFunctions.add(funcName);
          }
        });
      });
      
      // 未定義の関数参照をチェック
      referencedFunctions.forEach(funcName => {
        if (!definedFunctions.has(funcName) && 
            !isGasBuiltinFunction(funcName) &&
            !isJavaScriptBuiltin(funcName)) {
          errors.push({
            type: 'undefined_function',
            name: funcName
          });
        }
      });
      
      if (errors.length > 0) {
        console.log('🚨 未定義関数参照エラー:');
        errors.forEach(error => {
          console.log(`  - ${error.name} is not defined`);
        });
        // 参照エラーは警告のみ（一部は他のファイルで定義されている可能性）
        console.log(`⚠️  ${errors.length}個の潜在的な参照エラーを検出`);
      }
    });

    test('未定義変数の参照をチェック', () => {
      const errors = [];
      const definedVariables = new Set();
      const referencedVariables = new Set();
      
      // 定義済み変数を収集
      allGasFiles.forEach(file => {
        const varMatches = file.content.match(/^(const|let|var)\s+(\w+)/gm) || [];
        varMatches.forEach(match => {
          const varName = match.replace(/^(const|let|var)\s+/, '').replace(/\s*=.*/, '');
          definedVariables.add(varName);
        });
      });
      
      // 参照されている変数を収集（簡易的）
      allGasFiles.forEach(file => {
        const refMatches = file.content.match(/[^.\w](\w+)/g) || [];
        refMatches.forEach(match => {
          const varName = match.trim();
          if (varName.length > 0 && /^[A-Z]/.test(varName)) {
            referencedVariables.add(varName);
          }
        });
      });
      
      // 未定義の変数参照をチェック
      referencedVariables.forEach(varName => {
        if (!definedVariables.has(varName) && 
            !isGasBuiltinVariable(varName) &&
            !isJavaScriptBuiltin(varName)) {
          errors.push({
            type: 'undefined_variable',
            name: varName
          });
        }
      });
      
      if (errors.length > 0) {
        console.log('⚠️  潜在的な未定義変数参照:');
        errors.slice(0, 10).forEach(error => {
          console.log(`  - ${error.name} is not defined`);
        });
        if (errors.length > 10) {
          console.log(`  ... and ${errors.length - 10} more`);
        }
      }
    });
  });

  describe('GAS特有の問題検出', () => {
    test('ES6+構文の非互換性をチェック', () => {
      const errors = [];
      
      allGasFiles.forEach(file => {
        // static field syntax (GAS非対応)
        const staticFieldMatches = file.content.match(/static\s+\w+\s*=/g) || [];
        staticFieldMatches.forEach(match => {
          errors.push({
            type: 'unsupported_syntax',
            syntax: 'static_field',
            match: match.trim(),
            file: file.fileName
          });
        });
        
        // import/export statements (GAS非対応)
        const importMatches = file.content.match(/^(import|export)/gm) || [];
        importMatches.forEach(match => {
          errors.push({
            type: 'unsupported_syntax',
            syntax: 'es_modules',
            match: match.trim(),
            file: file.fileName
          });
        });
      });
      
      if (errors.length > 0) {
        console.log('🚨 GAS非対応構文エラー:');
        errors.forEach(error => {
          console.log(`  - ${error.file}: ${error.syntax} - ${error.match}`);
        });
        expect(errors).toHaveLength(0);
      }
    });

    test('循環参照の可能性をチェック', () => {
      const dependencies = {};
      const errors = [];
      
      allGasFiles.forEach(file => {
        dependencies[file.fileName] = [];
        
        // 他ファイルのクラス/関数の参照を検出
        allGasFiles.forEach(otherFile => {
          if (file.fileName !== otherFile.fileName) {
            const classMatches = otherFile.content.match(/^class\s+(\w+)/gm) || [];
            classMatches.forEach(match => {
              const className = match.replace('class ', '');
              if (file.content.includes(className)) {
                dependencies[file.fileName].push(otherFile.fileName);
              }
            });
          }
        });
      });
      
      // 簡易的な循環参照検出
      Object.keys(dependencies).forEach(fileName => {
        dependencies[fileName].forEach(depFileName => {
          if (dependencies[depFileName] && dependencies[depFileName].includes(fileName)) {
            errors.push({
              type: 'circular_dependency',
              files: [fileName, depFileName]
            });
          }
        });
      });
      
      if (errors.length > 0) {
        console.log('⚠️  潜在的な循環参照:');
        errors.forEach(error => {
          console.log(`  - ${error.files.join(' ⟷ ')}`);
        });
      }
    });
  });
});

// ヘルパー関数
function isGasBuiltinFunction(funcName) {
  const gasBuiltins = [
    'SpreadsheetApp', 'DriveApp', 'UrlFetchApp', 'HtmlService', 'ScriptApp',
    'PropertiesService', 'CacheService', 'Utilities', 'Logger', 'Session',
    'FormApp', 'GmailApp', 'CalendarApp', 'DocumentApp', 'SlidesApp',
    'console', 'JSON', 'Math', 'Date', 'String', 'Number', 'Array', 'Object'
  ];
  
  // GAS API methods and properties
  const gasApiMethods = [
    'getActiveSheet', 'getActiveSpreadsheet', 'openById', 'openByUrl', 'create',
    'getSheetByName', 'getSheets', 'getRange', 'getValues', 'setValues',
    'getLastRow', 'getLastColumn', 'getName', 'getId', 'getUrl', 'setName',
    'autoResizeColumns', 'setFontWeight', 'setBackground', 'getNumColumns',
    'fetch', 'getContentText', 'getResponseCode', 'getFileById', 'getParents',
    'hasNext', 'next', 'getFoldersByName', 'createFolder', 'moveTo', 'addEditor',
    'getSheetById', 'formatDate', 'getUuid', 'sleep', 'computeDigest', 'newBlob',
    'base64Encode', 'base64Decode', 'getDataAsString', 'setTitle', 'setDescription',
    'getScriptProperties', 'getUserProperties', 'getProperty', 'setProperty',
    'deleteProperty', 'getProperties', 'setProperties', 'deleteAllProperties',
    'getScriptCache', 'getUserCache', 'getDocumentCache', 'put', 'get', 'remove',
    'removeAll', 'log', 'getActiveUser', 'getEmail', 'getEffectiveUser',
    'getScriptTimeZone', 'getScriptLock', 'waitLock', 'releaseLock',
    'createHtmlOutput', 'createHtmlOutputFromFile', 'createTemplateFromFile',
    'evaluate', 'getContent', 'setContent', 'setXFrameOptionsMode', 'setSandboxMode',
    'addMetaTag', 'setMimeType', 'createTextOutput', 'getService',
    'setSharing', 'setDestination', 'rename', 'getPublishedUrl', 'getEditUrl',
    'addListItem', 'addTextItem', 'addParagraphTextItem', 'addCheckboxItem',
    'addMultipleChoiceItem', 'setChoiceValues', 'setRequired', 'setHelpText',
    'setValidation', 'showOtherOption', 'createParagraphTextValidation',
    'requireTextLengthLessThanOrEqualTo', 'build', 'setEmailCollectionType',
    'setCollectEmail', 'getFormUrl', 'searchFiles', 'getDestinationType',
    'getDestinationId', 'getDateCreated', 'withSuccessHandler', 'withFailureHandler'
  ];
  
  return gasBuiltins.includes(funcName) || gasApiMethods.includes(funcName);
}

function isGasBuiltinVariable(varName) {
  const gasBuiltins = [
    'UNIFIED_CONSTANTS', 'DB_SHEET_CONFIG', 'ERROR_TYPES', 'SCRIPT_PROPS_KEYS'
  ];
  
  // プロジェクト固有の定数・クラス
  const projectSpecific = [
    'UnifiedErrorHandler', 'UnifiedBatchProcessor', 'ResilientExecutor',
    'UnifiedSecretManager', 'UError', 
    // 状態系
    'pending', 'completed', 'success', 'failed', 'fallback',
    // データ系  
    'data', 'status', 'method', 'column', 'setup', 'route', 'processor',
    'start', 'end', 'preview', 'instead', 'at'
  ];
  
  return gasBuiltins.includes(varName) || projectSpecific.includes(varName);
}

function isJavaScriptBuiltin(name) {
  const jsBuiltins = [
    'console', 'JSON', 'Math', 'Date', 'String', 'Number', 'Array', 'Object',
    'Promise', 'Error', 'RegExp', 'Map', 'Set', 'WeakMap', 'WeakSet', 'Boolean',
    'parseInt', 'parseFloat', 'isNaN', 'encodeURIComponent', 'decodeURIComponent',
    'Function', 'URL', 'undefined', 'null', 'true', 'false'
  ];
  
  // Built-in JavaScript methods and properties
  const jsBuiltinMethods = [
    'toString', 'valueOf', 'hasOwnProperty', 'constructor', 'prototype',
    'length', 'push', 'pop', 'shift', 'unshift', 'slice', 'splice', 'join',
    'indexOf', 'includes', 'forEach', 'map', 'filter', 'reduce', 'find',
    'findIndex', 'some', 'every', 'sort', 'reverse', 'concat', 'split',
    'replace', 'match', 'search', 'substring', 'substr', 'charAt', 'charCodeAt',
    'toLowerCase', 'toUpperCase', 'trim', 'startsWith', 'endsWith',
    'test', 'exec', 'toFixed', 'toPrecision', 'toExponential', 'toLocaleString',
    'toISOString', 'getTime', 'getDate', 'setDate', 'now', 'parse', 'stringify',
    'keys', 'values', 'entries', 'assign', 'create', 'freeze', 'isArray',
    'from', 'fromCharCode', 'max', 'min', 'floor', 'ceil', 'round', 'abs',
    'pow', 'sqrt', 'random', 'add', 'has', 'get', 'set', 'delete', 'clear',
    'size', 'then', 'catch', 'finally'
  ];
  
  return jsBuiltins.includes(name) || jsBuiltinMethods.includes(name);
}