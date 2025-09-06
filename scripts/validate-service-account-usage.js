#!/usr/bin/env node

/**
 * 🔍 Service Account API使用パターン検証スクリプト
 * システム全体のService Account APIの一貫性と正確性を検証
 */

const fs = require('fs');
const path = require('path');

// 正しいAPI構造パターン
const CORRECT_API_PATTERNS = {
  // cache.gsで定義された正しい構造
  VALUES_API: {
    'batchGet': 'service.spreadsheets.values.batchGet',
    'append': 'service.spreadsheets.values.append', 
    'update': 'service.spreadsheets.values.update'
  },
  // 正しいSpreadsheets API
  SPREADSHEETS_API: {
    'batchUpdate': 'service.spreadsheets.batchUpdate'
  }
};

// 検証パターン
const VALIDATION_PATTERNS = {
  // Service Accountサービス取得
  SERVICE_ACQUISITION: /getSheetsServiceCached\(\)/g,
  
  // API呼び出しパターン
  VALUES_BATCHGET: /service\.spreadsheets\.values\.batchGet\s*\(/g,
  VALUES_APPEND: /service\.spreadsheets\.values\.append\s*\(/g,  
  VALUES_UPDATE: /service\.spreadsheets\.values\.update\s*\(/g,
  SPREADSHEETS_BATCHUPDATE: /service\.spreadsheets\.batchUpdate\s*\(/g,
  
  // 間違ったパターン（検出すべき問題）
  WRONG_NESTED: /service\.spreadsheets\.values\.spreadsheets/g,
  WRONG_ORDER: /service\.values\.spreadsheets/g,
  // 条件分岐での存在チェックは除外、実際の関数呼び出し漏れのみ検出
  MISSING_FUNCTION_CALL: /service\.spreadsheets\.values\.(?:batchGet|append|update)(?!\s*\()(?!\s*\))/g
};

class ServiceAccountUsageValidator {
  constructor() {
    this.results = {
      files: new Map(),
      issues: [],
      summary: {
        totalFiles: 0,
        apiCallCount: 0,
        issueCount: 0
      }
    };
  }

  /**
   * メイン検証実行
   */
  async validateServiceAccountUsage() {
    console.log('🔍 Service Account API使用パターン検証開始...');
    console.log('='.repeat(60));
    
    const srcDir = path.join(__dirname, '../src');
    const files = this.getGasFiles(srcDir);
    
    this.results.summary.totalFiles = files.length;
    console.log(`📁 検証対象ファイル数: ${files.length}`);
    
    // 各ファイルを検証
    for (const filePath of files) {
      const content = fs.readFileSync(filePath, 'utf8');
      const fileName = path.basename(filePath);
      
      await this.validateFile(content, fileName);
    }
    
    // 結果出力
    this.outputResults();
    
    return this.results;
  }
  
  /**
   * GASファイル取得
   */
  getGasFiles(dir) {
    const files = [];
    const entries = fs.readdirSync(dir, { withFileTypes: true });
    
    for (const entry of entries) {
      if (entry.isFile() && entry.name.endsWith('.gs')) {
        files.push(path.join(dir, entry.name));
      }
    }
    
    return files.sort();
  }
  
  /**
   * 単一ファイルの検証
   */
  async validateFile(content, fileName) {
    const fileResult = {
      fileName,
      serviceAcquisitions: [],
      apiCalls: [],
      issues: []
    };
    
    // Service Account取得箇所を記録
    const serviceMatches = [...content.matchAll(VALIDATION_PATTERNS.SERVICE_ACQUISITION)];
    for (const match of serviceMatches) {
      const line = this.getLineNumber(content, match.index);
      fileResult.serviceAcquisitions.push({ line, context: this.getContext(content, match.index) });
    }
    
    // 各API呼び出しパターンをチェック
    this.checkApiPatterns(content, fileName, fileResult);
    
    // 問題パターンをチェック
    this.checkIssuePatterns(content, fileName, fileResult);
    
    this.results.files.set(fileName, fileResult);
    this.results.summary.apiCallCount += fileResult.apiCalls.length;
    this.results.summary.issueCount += fileResult.issues.length;
  }
  
  /**
   * API呼び出しパターンをチェック
   */
  checkApiPatterns(content, fileName, fileResult) {
    // Values API
    for (const [apiName, pattern] of Object.entries(VALIDATION_PATTERNS)) {
      if (!apiName.startsWith('VALUES_') && !apiName.startsWith('SPREADSHEETS_')) continue;
      
      const matches = [...content.matchAll(pattern)];
      for (const match of matches) {
        const line = this.getLineNumber(content, match.index);
        const context = this.getContext(content, match.index);
        
        fileResult.apiCalls.push({
          type: apiName,
          line,
          match: match[0],
          context
        });
      }
    }
  }
  
  /**
   * 問題パターンをチェック
   */
  checkIssuePatterns(content, fileName, fileResult) {
    // 間違ったネストパターン
    const wrongNested = [...content.matchAll(VALIDATION_PATTERNS.WRONG_NESTED)];
    for (const match of wrongNested) {
      this.addIssue(fileResult, 'WRONG_NESTED', 
        `間違ったネスト構造: ${match[0]}`, 
        this.getLineNumber(content, match.index));
    }
    
    // 間違った順序
    const wrongOrder = [...content.matchAll(VALIDATION_PATTERNS.WRONG_ORDER)];
    for (const match of wrongOrder) {
      this.addIssue(fileResult, 'WRONG_ORDER',
        `間違ったプロパティ順序: ${match[0]}`,
        this.getLineNumber(content, match.index));
    }
    
    // 関数呼び出し忘れ
    const missingCall = [...content.matchAll(VALIDATION_PATTERNS.MISSING_FUNCTION_CALL)];
    for (const match of missingCall) {
      this.addIssue(fileResult, 'MISSING_FUNCTION_CALL',
        `関数呼び出し忘れ（括弧不足）: ${match[0]}`,
        this.getLineNumber(content, match.index));
    }
  }
  
  /**
   * 問題を追加
   */
  addIssue(fileResult, type, message, line) {
    const issue = { type, message, line };
    fileResult.issues.push(issue);
    this.results.issues.push({ ...issue, fileName: fileResult.fileName });
  }
  
  /**
   * 行番号取得
   */
  getLineNumber(content, index) {
    return content.substring(0, index).split('\n').length;
  }
  
  /**
   * コンテキスト取得
   */
  getContext(content, index) {
    const lines = content.split('\n');
    const lineIndex = this.getLineNumber(content, index) - 1;
    const start = Math.max(0, lineIndex - 1);
    const end = Math.min(lines.length, lineIndex + 2);
    
    return lines.slice(start, end).join('\n');
  }
  
  /**
   * 結果出力
   */
  outputResults() {
    console.log('\n📊 Service Account API検証結果');
    console.log('='.repeat(60));
    console.log(`総ファイル数: ${this.results.summary.totalFiles}`);
    console.log(`API呼び出し総数: ${this.results.summary.apiCallCount}`);
    console.log(`問題検出数: ${this.results.summary.issueCount}`);
    
    if (this.results.summary.issueCount === 0) {
      console.log('\n✅ Service Account API使用パターンに問題はありません！');
      console.log('🎉 システム全体の一貫性が確保されています。');
      this.printSummary();
      return;
    }
    
    console.log('\n🚨 検出された問題:');
    console.log('-'.repeat(50));
    
    for (const issue of this.results.issues) {
      console.log(`\n❌ [${issue.fileName}:${issue.line}] ${issue.type}`);
      console.log(`   ${issue.message}`);
    }
    
    // 修正提案
    console.log('\n🔧 修正提案:');
    console.log('-'.repeat(40));
    
    const issueTypes = [...new Set(this.results.issues.map(i => i.type))];
    
    if (issueTypes.includes('WRONG_NESTED')) {
      console.log('1. ネスト構造を正しい形式に修正');
    }
    if (issueTypes.includes('WRONG_ORDER')) {
      console.log('2. プロパティアクセス順序を修正');  
    }
    if (issueTypes.includes('MISSING_FUNCTION_CALL')) {
      console.log('3. 関数呼び出しに括弧を追加');
    }
    
    console.log('\n📋 修正後の再検証:');
    console.log('   - このスクリプトを再実行');
    console.log('   - clasp push でGASにデプロイ');
    console.log('   - 実際のAPI呼び出しをテスト');
  }
  
  /**
   * 概要情報出力
   */
  printSummary() {
    console.log('\n📋 API使用状況サマリー:');
    console.log('-'.repeat(40));
    
    // ファイル別のAPI使用状況
    for (const [fileName, fileResult] of this.results.files) {
      if (fileResult.apiCalls.length > 0 || fileResult.serviceAcquisitions.length > 0) {
        console.log(`\n📁 ${fileName}:`);
        
        if (fileResult.serviceAcquisitions.length > 0) {
          console.log(`   Service取得: ${fileResult.serviceAcquisitions.length}箇所`);
        }
        
        // API別の集計
        const apiTypes = {};
        for (const call of fileResult.apiCalls) {
          apiTypes[call.type] = (apiTypes[call.type] || 0) + 1;
        }
        
        for (const [type, count] of Object.entries(apiTypes)) {
          const apiName = type.replace('VALUES_', '').replace('SPREADSHEETS_', '').toLowerCase();
          console.log(`   ${apiName}: ${count}回`);
        }
      }
    }
    
    console.log('\n✨ 全てのService Account API呼び出しが正しく実装されています！');
  }
}

// メイン実行
if (require.main === module) {
  const validator = new ServiceAccountUsageValidator();
  validator.validateServiceAccountUsage().catch(console.error);
}

module.exports = { ServiceAccountUsageValidator };