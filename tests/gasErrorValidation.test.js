/**
 * GAS重複宣言エラー・参照エラー検出テスト（簡素化版）
 * 主要なエラーのみをチェックし、テスト数を適切に制限
 */

const fs = require('fs');
const path = require('path');

describe('GAS Error Validation Tests', () => {
  let allGasFiles = [];

  beforeAll(() => {
    // 全.gsファイルを読み込み
    const srcDir = path.join(__dirname, '../src');
    const gasFiles = fs.readdirSync(srcDir).filter(file => file.endsWith('.gs'));
    
    allGasFiles = gasFiles.map(file => {
      const filePath = path.join(srcDir, file);
      const content = fs.readFileSync(filePath, 'utf8');
      return { fileName: file, content, filePath };
    });
  });

  describe('重要なエラー検出', () => {
    test('重複関数宣言をチェック', () => {
      const functionDeclarations = {};
      const errors = [];
      
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
        console.log(`🚨 ${errors.length}個の重複関数宣言を検出:`);
        errors.slice(0, 5).forEach(error => {
          console.log(`  - ${error.name}: ${error.files.join(' と ')}`);
        });
        if (errors.length > 5) {
          console.log(`  ... and ${errors.length - 5} more duplicates`);
        }
      } else {
        console.log('✅ 重複関数宣言エラーなし');
      }
      
      // 重複宣言は0個であることを期待
      expect(errors).toHaveLength(0);
    });

    test('GAS非対応構文をチェック', () => {
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
        console.log(`🚨 ${errors.length}個のGAS非対応構文を検出:`);
        errors.slice(0, 3).forEach(error => {
          console.log(`  - ${error.file}: ${error.syntax} - ${error.match}`);
        });
        if (errors.length > 3) {
          console.log(`  ... and ${errors.length - 3} more syntax issues`);
        }
      } else {
        console.log('✅ GAS非対応構文エラーなし');
      }
      
      // GAS非対応構文は0個であることを期待
      expect(errors).toHaveLength(0);
    });
  });
});