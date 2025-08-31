#!/usr/bin/env node

/**
 * Standalone Function Mapping Validator
 * HTMLとGASの関数マッピング検証をスタンドアロンで実行
 * 
 * Usage:
 *   node scripts/validate-functions.js
 *   npm run validate:functions
 */

const fs = require('fs');
const path = require('path');
const glob = require('glob');

// カラー出力用のANSIコード
const colors = {
  red: '\x1b[31m',
  green: '\x1b[32m',
  yellow: '\x1b[33m',
  blue: '\x1b[34m',
  magenta: '\x1b[35m',
  cyan: '\x1b[36m',
  reset: '\x1b[0m',
  bold: '\x1b[1m'
};

class FunctionMappingValidator {
  constructor(srcDir = './src') {
    this.srcDir = path.resolve(srcDir);
    this.htmlFiles = [];
    this.gasFiles = [];
    this.functionCalls = [];
    this.gasFunctions = [];
    this.issues = {
      missing: [],
      unused: [],
      patterns: []
    };
  }

  async validate() {
    console.log(`${colors.bold}🔍 HTML/GAS Function Mapping Validation${colors.reset}\n`);
    
    this.scanFiles();
    this.extractFunctionCalls();
    this.extractGasFunctions();
    this.checkMissingFunctions();
    this.checkUnusedFunctions();
    this.checkCallPatterns();
    this.generateReport();
    
    return this.issues.missing.length === 0;
  }

  scanFiles() {
    this.htmlFiles = glob.sync(path.join(this.srcDir, '**/*.html'));
    this.gasFiles = glob.sync(path.join(this.srcDir, '**/*.gs'));
    
    console.log(`📊 Scanned: ${colors.cyan}${this.htmlFiles.length}${colors.reset} HTML files, ${colors.cyan}${this.gasFiles.length}${colors.reset} GAS files`);
  }

  extractFunctionCalls() {
    this.functionCalls = [];
    
    this.htmlFiles.forEach(file => {
      const content = fs.readFileSync(file, 'utf8');
      
      // 多行パターンを含む全体的な検索
      const multilinePatterns = [
        /google\.script\.run\s*\.withSuccessHandler\([^)]*\)\s*\.withFailureHandler\([^)]*\)\s*\.(\w+)\s*\(/g,
        /google\.script\.run\s*\.withSuccessHandler\([^)]*\)\s*\.(\w+)\s*\(/g,
        /google\.script\.run\s*\.withFailureHandler\([^)]*\)\s*\.(\w+)\s*\(/g,
        /google\.script\.run\s*\.(\w+)\s*\(/g
      ];
      
      multilinePatterns.forEach(pattern => {
        let match;
        while ((match = pattern.exec(content)) !== null) {
          // 関数名がGASの予約語でないことを確認
          const funcName = match[1];
          if (!this.isGasReservedWord(funcName)) {
            const lines = content.substring(0, match.index).split('\n');
            const lineNumber = lines.length;
            const lineContent = lines[lines.length - 1] + content.substring(match.index, match.index + match[0].length);
            
            this.functionCalls.push({
              functionName: funcName,
              file: path.relative(process.cwd(), file),
              line: lineNumber,
              pattern: lineContent.trim(),
              hasSuccessHandler: match[0].includes('.withSuccessHandler'),
              hasFailureHandler: match[0].includes('.withFailureHandler')
            });
          }
        }
      });
    });
    
    console.log(`🔍 Found ${colors.cyan}${this.functionCalls.length}${colors.reset} function calls`);
  }

  extractGasFunctions() {
    this.gasFunctions = [];
    
    this.gasFiles.forEach(file => {
      const content = fs.readFileSync(file, 'utf8');
      const lines = content.split('\n');
      
      lines.forEach((line, index) => {
        const patterns = [
          /^function\s+(\w+)\s*\(/,
          /^\s*function\s+(\w+)\s*\(/,
          /const\s+(\w+)\s*=\s*function/,
          /let\s+(\w+)\s*=\s*function/,
          /(\w+)\s*:\s*function/,
        ];
        
        patterns.forEach(pattern => {
          const match = line.match(pattern);
          if (match) {
            this.gasFunctions.push({
              name: match[1],
              file: path.relative(process.cwd(), file),
              line: index + 1,
              definition: line.trim()
            });
          }
        });
      });
    });
    
    console.log(`📝 Found ${colors.cyan}${this.gasFunctions.length}${colors.reset} GAS functions\n`);
  }

  checkMissingFunctions() {
    this.functionCalls.forEach(call => {
      const gasFunction = this.gasFunctions.find(f => f.name === call.functionName);
      
      if (!gasFunction) {
        const similarFunctions = this.findSimilarFunctions(call.functionName);
        this.issues.missing.push({
          ...call,
          similarFunctions
        });
      }
    });
  }

  checkUnusedFunctions() {
    this.gasFunctions.forEach(gasFunc => {
      const isUsed = this.functionCalls.some(call => call.functionName === gasFunc.name);
      
      if (!isUsed && !this.isInternalFunction(gasFunc.name)) {
        this.issues.unused.push(gasFunc);
      }
    });
  }

  checkCallPatterns() {
    this.functionCalls.forEach(call => {
      if (call.hasSuccessHandler && !call.hasFailureHandler) {
        this.issues.patterns.push({
          type: 'missing-failure-handler',
          message: 'Missing withFailureHandler',
          suggestion: 'Add .withFailureHandler() for proper error handling',
          ...call
        });
      }
      
      if (!call.hasSuccessHandler && !call.hasFailureHandler) {
        this.issues.patterns.push({
          type: 'direct-call',
          message: 'Direct call without handlers',
          suggestion: 'Consider adding success/failure handlers',
          ...call
        });
      }
    });
  }

  findSimilarFunctions(targetName, threshold = 3) {
    const similar = [];
    
    this.gasFunctions.forEach(func => {
      const distance = this.levenshteinDistance(targetName, func.name);
      if (distance <= threshold && distance > 0) {
        similar.push(func.name);
      }
    });
    
    return similar.slice(0, 5);
  }

  levenshteinDistance(str1, str2) {
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

  isInternalFunction(funcName) {
    const internalFunctions = [
      'doGet', 'doPost', 'onOpen', 'onEdit', 'onFormSubmit',
      'onInstall', 'include', 'main', 'init'
    ];
    
    return internalFunctions.includes(funcName) || 
           funcName.startsWith('_') || 
           funcName.startsWith('test') || 
           funcName.startsWith('debug');
  }

  isGasReservedWord(funcName) {
    const reservedWords = [
      'withSuccessHandler', 'withFailureHandler', 'withUserObject'
    ];
    return reservedWords.includes(funcName);
  }

  generateReport() {
    console.log(`${colors.bold}📋 Validation Report${colors.reset}`);
    console.log('='.repeat(50));
    
    // Missing Functions (Critical)
    if (this.issues.missing.length > 0) {
      console.log(`\n${colors.red}❌ CRITICAL: Missing Functions (${this.issues.missing.length})${colors.reset}`);
      this.issues.missing.forEach(issue => {
        console.log(`   ${colors.red}•${colors.reset} ${colors.bold}${issue.functionName}${colors.reset} (${issue.file}:${issue.line})`);
        if (issue.similarFunctions.length > 0) {
          console.log(`     ${colors.yellow}💡 Similar: ${issue.similarFunctions.join(', ')}${colors.reset}`);
        }
      });
    } else {
      console.log(`\n${colors.green}✅ All function calls have corresponding implementations${colors.reset}`);
    }
    
    // Unused Functions (Info)
    if (this.issues.unused.length > 0) {
      console.log(`\n${colors.yellow}📋 Potential Cleanup: Unused Functions (${this.issues.unused.length})${colors.reset}`);
      this.issues.unused.forEach(func => {
        console.log(`   ${colors.yellow}•${colors.reset} ${func.name} (${func.file}:${func.line})`);
      });
    }
    
    // Pattern Issues (Warning)
    if (this.issues.patterns.length > 0) {
      console.log(`\n${colors.magenta}⚠️  Pattern Issues (${this.issues.patterns.length})${colors.reset}`);
      this.issues.patterns.forEach(issue => {
        console.log(`   ${colors.magenta}•${colors.reset} ${issue.message}: ${colors.bold}${issue.functionName}${colors.reset} (${issue.file}:${issue.line})`);
        console.log(`     ${colors.cyan}💡 ${issue.suggestion}${colors.reset}`);
      });
    }
    
    console.log('\n' + '='.repeat(50));
    
    if (this.issues.missing.length === 0) {
      console.log(`${colors.green}🎉 Function mapping validation passed!${colors.reset}`);
      return true;
    } else {
      console.log(`${colors.red}💥 Function mapping validation failed!${colors.reset}`);
      console.log(`${colors.red}   Fix ${this.issues.missing.length} missing function(s) before proceeding.${colors.reset}`);
      return false;
    }
  }
}

// CLI実行
if (require.main === module) {
  const validator = new FunctionMappingValidator();
  
  validator.validate().then(success => {
    process.exit(success ? 0 : 1);
  }).catch(error => {
    console.error(`${colors.red}Error during validation:${colors.reset}`, error);
    process.exit(1);
  });
}

module.exports = FunctionMappingValidator;