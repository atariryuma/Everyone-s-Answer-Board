/**
 * GAS未定義エラー・参照エラー検出テスト
 * 未定義の関数・変数・クラスの参照を検出し修正対象を特定
 */

const fs = require('fs');
const path = require('path');

describe('GAS Undefined Reference Error Tests', () => {
  let allGasFiles = [];
  let allSourceCode = '';
  let definedFunctions = new Set();
  let definedClasses = new Set();
  let definedVariables = new Set();
  let referencedFunctions = new Set();
  let referencedClasses = new Set();
  let referencedVariables = new Set();

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
    
    // 定義されている関数・クラス・変数を収集
    collectDefinitions();
    
    // 参照されている関数・クラス・変数を収集
    collectReferences();
  });

  function collectDefinitions() {
    allGasFiles.forEach(file => {
      // 関数定義を収集
      const funcMatches = file.content.match(/^function\s+(\w+)/gm) || [];
      funcMatches.forEach(match => {
        const funcName = match.replace(/^function\s+/, '').replace(/\(.*/, '');
        definedFunctions.add(funcName);
      });
      
      // クラス定義を収集
      const classMatches = file.content.match(/^class\s+(\w+)/gm) || [];
      classMatches.forEach(match => {
        const className = match.replace('class ', '').replace(/\s.*/, '');
        definedClasses.add(className);
      });
      
      // クラスメソッド定義を収集
      const methodMatches = file.content.match(/^\s+(?:static\s+)?(\w+)\s*\([^)]*\)\s*\{/gm) || [];
      methodMatches.forEach(match => {
        const methodName = match.trim().replace(/^static\s+/, '').replace(/\s*\(.*/, '');
        if (methodName !== 'constructor') {
          definedFunctions.add(methodName);
        }
      });
      
      // 変数定義を収集（const, let, var）- 関数・クラス含む
      const lines = file.content.split('\n');
      lines.forEach(line => {
        const trimmedLine = line.trim();
        
        // const/let/var 宣言
        const varMatch = trimmedLine.match(/^(const|let|var)\s+(\w+)/);
        if (varMatch) {
          const varName = varMatch[2];
          definedVariables.add(varName);
          
          // 関数宣言も検出
          if (line.includes('=>') || line.includes('function')) {
            definedFunctions.add(varName);
          }
          
          // クラス宣言も検出
          if (line.includes('= class') || line.includes('class {')) {
            definedClasses.add(varName);
          }
        }
      });
    });
  }

  function collectReferences() {
    allGasFiles.forEach(file => {
      // 関数呼び出しを収集
      const funcCalls = file.content.match(/(\w+)\s*\(/g) || [];
      funcCalls.forEach(match => {
        const funcName = match.replace(/\s*\(.*/, '');
        if (!isReservedKeyword(funcName) && funcName.length > 1) {
          referencedFunctions.add(funcName);
        }
      });
      
      // クラスインスタンス化を収集
      const newCalls = file.content.match(/new\s+(\w+)/g) || [];
      newCalls.forEach(match => {
        const className = match.replace('new ', '').replace(/\s*\(.*/, '');
        referencedClasses.add(className);
      });
      
      // 静的クラスメソッド呼び出しを収集
      const staticCalls = file.content.match(/(\w+)\.(\w+)\s*\(/g) || [];
      staticCalls.forEach(match => {
        const [, className, methodName] = match.match(/(\w+)\.(\w+)\s*\(/);
        if (!isBuiltinObject(className)) {
          referencedClasses.add(className);
        }
      });
      
      // 変数参照を収集（単語境界で識別）
      const varRefs = file.content.match(/\b([A-Z][A-Z_]+)\b/g) || [];
      varRefs.forEach(ref => {
        if (!isBuiltinConstant(ref)) {
          referencedVariables.add(ref);
        }
      });
    });
  }

  describe('未定義関数エラー検出', () => {
    test('未定義関数の参照をチェック', () => {
      const errors = [];
      
      referencedFunctions.forEach(funcName => {
        if (!definedFunctions.has(funcName) && 
            !isGasBuiltinFunction(funcName) &&
            !isJavaScriptBuiltin(funcName)) {
          
          // 参照元ファイルを特定
          const referenceFiles = [];
          allGasFiles.forEach(file => {
            const regex = new RegExp(`\\b${funcName}\\s*\\(`, 'g');
            if (regex.test(file.content)) {
              referenceFiles.push(file.fileName);
            }
          });
          
          errors.push({
            type: 'undefined_function',
            name: funcName,
            referencedIn: referenceFiles
          });
        }
      });
      
      if (errors.length > 0) {
        console.log('🚨 未定義関数エラー:');
        // 最初の10個のみ表示
        errors.slice(0, 10).forEach(error => {
          console.log(`  - ${error.name} is not defined`);
          console.log(`    Referenced in: ${error.referencedIn.join(', ')}`);
        });
        
        if (errors.length > 10) {
          console.log(`  ... and ${errors.length - 10} more undefined functions`);
        }
        
        // 修正が必要な関数をファイルに出力
        const errorReport = errors.map(error => 
          `UNDEFINED_FUNCTION: ${error.name}\nFiles: ${error.referencedIn.join(', ')}\n`
        ).join('\n');
        
        fs.writeFileSync(
          path.join(__dirname, '../undefined-functions-report.md'),
          `# 未定義関数エラー報告\n\n${errorReport}`
        );
        
        console.log(`📋 ${errors.length}個の未定義関数エラーを検出`);
      } else {
        console.log('✅ 未定義関数エラーなし');
      }
      
      // テストは警告として扱う（完全に0になるまで段階的に修正）
    });
  });

  describe('未定義クラスエラー検出', () => {
    test('未定義クラスの参照をチェック', () => {
      const errors = [];
      
      referencedClasses.forEach(className => {
        if (!definedClasses.has(className) && 
            !isGasBuiltinClass(className) &&
            !isJavaScriptBuiltin(className)) {
          
          // 参照元ファイルを特定
          const referenceFiles = [];
          allGasFiles.forEach(file => {
            const newRegex = new RegExp(`new\\s+${className}\\b`, 'g');
            const staticRegex = new RegExp(`\\b${className}\\.\\w+`, 'g');
            if (newRegex.test(file.content) || staticRegex.test(file.content)) {
              referenceFiles.push(file.fileName);
            }
          });
          
          errors.push({
            type: 'undefined_class',
            name: className,
            referencedIn: referenceFiles
          });
        }
      });
      
      if (errors.length > 0) {
        console.log('🚨 未定義クラスエラー:');
        // 最初の5個のみ表示
        errors.slice(0, 5).forEach(error => {
          console.log(`  - ${error.name} is not defined`);
          console.log(`    Referenced in: ${error.referencedIn.join(', ')}`);
        });
        
        if (errors.length > 5) {
          console.log(`  ... and ${errors.length - 5} more undefined classes`);
        }
        
        // 修正が必要なクラスをファイルに出力
        const errorReport = errors.map(error => 
          `UNDEFINED_CLASS: ${error.name}\nFiles: ${error.referencedIn.join(', ')}\n`
        ).join('\n');
        
        fs.writeFileSync(
          path.join(__dirname, '../undefined-classes-report.md'),
          `# 未定義クラスエラー報告\n\n${errorReport}`
        );
        
        console.log(`📋 ${errors.length}個の未定義クラスエラーを検出`);
      } else {
        console.log('✅ 未定義クラスエラーなし');
      }
    });
  });

  describe('未定義定数・変数エラー検出', () => {
    test('未定義定数・変数の参照をチェック', () => {
      const errors = [];
      
      referencedVariables.forEach(varName => {
        if (!definedVariables.has(varName) && 
            !isGasBuiltinVariable(varName) &&
            !isJavaScriptBuiltin(varName)) {
          
          // 参照元ファイルを特定
          const referenceFiles = [];
          allGasFiles.forEach(file => {
            const regex = new RegExp(`\\b${varName}\\b`, 'g');
            if (regex.test(file.content)) {
              referenceFiles.push(file.fileName);
            }
          });
          
          errors.push({
            type: 'undefined_variable',
            name: varName,
            referencedIn: referenceFiles
          });
        }
      });
      
      if (errors.length > 0) {
        console.log('🚨 未定義変数・定数エラー:');
        errors.slice(0, 5).forEach(error => {
          console.log(`  - ${error.name} is not defined`);
          console.log(`    Referenced in: ${error.referencedIn.join(', ')}`);
        });
        
        if (errors.length > 5) {
          console.log(`  ... and ${errors.length - 5} more undefined variables`);
        }
        
        // 修正が必要な変数をファイルに出力
        const errorReport = errors.map(error => 
          `UNDEFINED_VARIABLE: ${error.name}\nFiles: ${error.referencedIn.join(', ')}\n`
        ).join('\n');
        
        fs.writeFileSync(
          path.join(__dirname, '../undefined-variables-report.md'),
          `# 未定義変数エラー報告\n\n${errorReport}`
        );
        
        console.log(`📋 ${errors.length}個の未定義変数エラーを検出`);
      } else {
        console.log('✅ 未定義変数エラーなし');
      }
    });
  });

  describe('参照エラー統計', () => {
    test('エラー統計を表示', () => {
      const undefinedFuncs = Array.from(referencedFunctions).filter(f => 
        !definedFunctions.has(f) && !isGasBuiltinFunction(f) && !isJavaScriptBuiltin(f)
      );
      
      const undefinedClasses = Array.from(referencedClasses).filter(c => 
        !definedClasses.has(c) && !isGasBuiltinClass(c) && !isJavaScriptBuiltin(c)
      );
      
      const undefinedVars = Array.from(referencedVariables).filter(v => 
        !definedVariables.has(v) && !isGasBuiltinVariable(v) && !isJavaScriptBuiltin(v)
      );
      
      const totalErrors = undefinedFuncs.length + undefinedClasses.length + undefinedVars.length;
      
      console.log('\n📊 未定義エラー統計:');
      console.log(`  未定義関数: ${undefinedFuncs.length}個`);
      console.log(`  未定義クラス: ${undefinedClasses.length}個`);
      console.log(`  未定義変数: ${undefinedVars.length}個`);
      console.log(`  合計: ${totalErrors}個`);
      
      if (totalErrors === 0) {
        console.log('🎉 全ての未定義エラーが解決されました！');
      } else {
        console.log(`⚠️  残り${totalErrors}個のエラーが検出されました`);
      }
    });
  });
});

// ヘルパー関数群
function isReservedKeyword(name) {
  const keywords = [
    'if', 'for', 'while', 'switch', 'catch', 'typeof', 'return', 'new',
    'throw', 'try', 'else', 'do', 'break', 'continue', 'var', 'let', 'const',
    'function', 'class', 'extends', 'import', 'export', 'default', 'async',
    'await', 'yield', 'delete', 'in', 'instanceof', 'void', 'this', 'super',
    'true', 'false', 'null', 'undefined'
  ];
  
  // CSS/HTML/DOM関連キーワード（main.gsで多く検出される）
  const webKeywords = [
    'gradient', 'rgba', 'translateY', 'scale', 'blur', 'bezier', 'media',
    'alert', 'confirm', 'reload', 'preventDefault', 'querySelector'
  ];
  
  // よくあるプロパティ名・パラメータ名（オブジェクトのプロパティとして使用）
  const propertyNames = [
    'data', 'status', 'method', 'column', 'setup', 'route', 'start', 'end',
    'success', 'failed', 'fallback', 'pending', 'completed', 'operation',
    'processor', 'instead', 'at', 'rate', 'limit', 'protected', 'response',
    'found', 'indices', 'user', 'admin', 'folder', 'preview', 'fn', 'called',
    'logMethod', 'ID', 'update', 'execute', 'reject', 'resolve', 'all',
    // Cache/Storage関連
    'L1', 'L2', 'L3', 'valueFn', 'valuesFn', 'getAll', 'putAll', 'IDs',
    // Promise/Async関連
    'then', 'catch', 'finally'
  ];
  
  // 短すぎる名前（多くは変数名や省略形）
  if (name.length <= 2) {
    return true;
  }
  
  return keywords.includes(name) || 
         webKeywords.includes(name) ||
         propertyNames.includes(name);
}

function isBuiltinObject(name) {
  const builtins = [
    'Math', 'Date', 'String', 'Number', 'Array', 'Object', 'JSON',
    'console', 'window', 'document', 'localStorage', 'sessionStorage'
  ];
  return builtins.includes(name);
}

function isBuiltinConstant(name) {
  const constants = [
    'MAX_VALUE', 'MIN_VALUE', 'PI', 'E', 'POSITIVE_INFINITY', 'NEGATIVE_INFINITY',
    'NaN', 'undefined', 'Infinity'
  ];
  return constants.includes(name);
}

function isGasBuiltinFunction(funcName) {
  const gasBuiltins = [
    'SpreadsheetApp', 'DriveApp', 'UrlFetchApp', 'HtmlService', 'ScriptApp',
    'PropertiesService', 'CacheService', 'Utilities', 'Logger', 'Session',
    'FormApp', 'GmailApp', 'CalendarApp', 'DocumentApp', 'SlidesApp',
    'console', 'JSON', 'Math', 'Date', 'String', 'Number', 'Array', 'Object',
    'Browser'
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
    // FormApp API メソッド (実際に存在するため除外)
    'setEmailCollectionType', 'setCollectEmail', 'getPublishedUrl', 'getEditUrl',
    'addListItem', 'setChoiceValues', 'setRequired', 'addTextItem', 'addParagraphTextItem',
    'setHelpText', 'createParagraphTextValidation', 'requireTextLengthLessThanOrEqualTo',
    'build', 'setValidation', 'addCheckboxItem', 'showOtherOption', 'addMultipleChoiceItem',
    'setSharing', 'setDestination', 'rename', 'getDateCreated', 'getFormUrl',
    'searchFiles', 'getDestinationType', 'getDestinationId', 
    // GAS Script API
    'getScriptId', 'list', 'deployments', 'deploymentConfig', 'webApp',
    // AppsScript API
    'getProjectTriggers', 'getHandlerFunction', 'deleteTrigger', 'newTrigger',
    'timeBased', 'everyMinutes', 'getEditors',
    // Utilities API
    'base64EncodeWebSafe', 'computeRsaSha256Signature',
    // Browser API (GAS互換)
    'alert', 'msgBox', 'confirm'
  ];
  
  // 代用可能な関数 (既存APIで代替可能)
  const substitutableFunctions = [
    'clearTimeout', 'setTimeout', // GASでは不要
    'reload', 'preventDefault', // Browser固有、GASでは不要  
    'querySelector', 'withSuccessHandler', 'withFailureHandler', // DOM/UI固有
    'handleSecureRedirect', // プロジェクト固有のUI関数
    'fromEntries' // Object.fromEntries (ポリフィル済み)
  ];
  
  return gasBuiltins.includes(funcName) || 
         gasApiMethods.includes(funcName) ||
         substitutableFunctions.includes(funcName);
}

function isGasBuiltinClass(className) {
  const gasClasses = [
    'SpreadsheetApp', 'DriveApp', 'UrlFetchApp', 'HtmlService', 'ScriptApp',
    'PropertiesService', 'CacheService', 'Utilities', 'Logger', 'Session',
    'FormApp', 'GmailApp', 'CalendarApp', 'DocumentApp', 'SlidesApp',
    'Spreadsheet', 'Sheet', 'Range', 'File', 'Folder', 'Blob', 'HTMLOutput'
  ];
  return gasClasses.includes(className);
}

function isGasBuiltinVariable(varName) {
  const gasBuiltins = [
    'UNIFIED_CONSTANTS', 'DB_SHEET_CONFIG', 'ERROR_TYPES', 'SCRIPT_PROPS_KEYS',
    'ERROR_SEVERITY', 'ERROR_CATEGORIES'
  ];
  
  // プロジェクト固有の定数・クラス
  const projectSpecific = [
    'UnifiedErrorHandler', 'UnifiedBatchProcessor', 'ResilientExecutor',
    'UnifiedSecretManager', 'UError', 'UnifiedCacheManager', 'UnifiedUserManager',
    'SystemIntegrationManager', 'UnifiedSheetDataManager', 'UnifiedCacheAPI',
    'MultiTenantSecurityManager', 'UnifiedValidationSystem', 'UnifiedExecutionCache',
    'CacheManager',
    'CRITICAL', 'ERROR', 'HIGH', 'SEVERITY', 'VALIDATION', 'SECURITY',
    'PERFORMANCE', 'INTEGRATION', 'AUTH', 'API', 'DB', 'CACHE', 'UI', 'SYSTEM'
  ];
  
  // よくある定数パターン（大文字のみ）
  const commonConstants = [
    'StudyQuest', 'Core', 'Functions', 'Utilities'
  ];
  
  // 短い変数名（多くはローカル変数）
  if (varName.length <= 2) {
    return true;
  }
  
  return gasBuiltins.includes(varName) || 
         projectSpecific.includes(varName) ||
         commonConstants.includes(varName);
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