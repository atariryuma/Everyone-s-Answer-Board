#!/usr/bin/env node

/**
 * @fileoverview GAS プロジェクト用依存関係解析器
 * Google Apps Script の特殊な依存関係パターンを解析し、
 * 使用・未使用のファイルと関数を特定する
 */

const fs = require('fs');
const path = require('path');
const glob = require('glob');

class GasDependencyAnalyzer {
  constructor(srcDir = './src') {
    this.srcDir = path.resolve(srcDir);
    this.dependencies = new Map();
    this.functions = new Map();
    this.files = new Set();
    this.usedFiles = new Set();
    this.usedFunctions = new Set();
    
    // GAS特有のパターン（強化版）
    this.patterns = {
      // include('FileName') パターン
      include: /include\(\s*['"`]([^'"`]+)['"`]\s*\)/g,
      // google.script.run.functionName() パターン
      gasCall: /google\.script\.run(?:\.\w+)*\.(\w+)\s*\(/g,
      // 関数定義パターン（複数行対応）
      functionDef: /function\s+(\w+)\s*\([^)]*\)\s*\{/g,
      // const/let/var functionName = パターン
      constFunctionDef: /(?:const|let|var)\s+(\w+)\s*=\s*(?:function|\([^)]*\)\s*=>)/g,
      // クラスメソッド定義パターン
      methodDef: /^\s*(\w+)\s*\([^)]*\)\s*\{/gm,
      // 直接関数呼び出しパターン（改良）
      directCall: /\b(\w+)\s*\(/g,
      // HTMLファイル内のJavaScript関数呼び出し
      htmlJsCall: /(?:onclick|onchange|onload|onerror)\s*=\s*['"`]([^'"`(]+)\s*\(/g,
      // google.script.host.close() パターン
      hostCall: /google\.script\.host\.(\w+)\s*\(/g,
      // UrlFetchApp, SpreadsheetApp等のGAS API使用
      gasApiUsage: /(?:UrlFetchApp|SpreadsheetApp|DriveApp|PropertiesService|CacheService|HtmlService|ScriptApp|Session|Utilities|Logger)\.(\w+)/g,
      // トリガー関数パターン
      triggerFunction: /ScriptApp\.newTrigger\s*\(\s*['"`](\w+)['"`]/g,
      // 動的関数呼び出し
      dynamicCall: /eval\s*\(\s*['"`](\w+)\s*\(/g,
      // コメント内の参照（無視すべき）
      commentPattern: /\/\*[\s\S]*?\*\/|\/\/.*$/gm
    };

    // エントリーポイント関数（必ず保持）
    this.entryPoints = new Set([
      'doGet', 'doPost', 'include', 'onOpen', 'onEdit', 'onFormSubmit',
      'installTrigger', 'uninstallTrigger', 'onInstall'
    ]);

    // 除外するファイルパターン
    this.excludeFiles = new Set([
      'appsscript.json'
    ]);
  }

  /**
   * 解析を実行
   */
  async analyze() {
    console.log('🔍 GAS依存関係解析を開始...');
    
    try {
      await this.scanFiles();
      await this.extractFunctions();
      await this.analyzeDependencies();
      await this.identifyUnused();
      
      console.log('✅ 解析完了');
      return this.generateReport();
    } catch (error) {
      console.error('❌ 解析エラー:', error);
      throw error;
    }
  }

  /**
   * ファイル一覧をスキャン
   */
  async scanFiles() {
    const pattern = path.join(this.srcDir, '**/*');
    const allFiles = glob.sync(pattern);
    
    for (const filePath of allFiles) {
      if (fs.statSync(filePath).isFile()) {
        const relativePath = path.relative(this.srcDir, filePath);
        const baseName = path.basename(relativePath, path.extname(relativePath));
        
        if (!this.excludeFiles.has(relativePath)) {
          this.files.add(relativePath);
          
          // ファイル情報を初期化
          this.dependencies.set(relativePath, {
            type: this.getFileType(relativePath),
            size: fs.statSync(filePath).size,
            functions: new Set(),
            dependencies: new Set(),
            dependents: new Set()
          });
        }
      }
    }
    
    console.log(`📂 ${this.files.size} ファイルをスキャンしました`);
  }

  /**
   * 各ファイルから関数を抽出（強化版）
   */
  async extractFunctions() {
    for (const filePath of this.files) {
      const fullPath = path.join(this.srcDir, filePath);
      const content = fs.readFileSync(fullPath, 'utf-8');
      const fileInfo = this.dependencies.get(filePath);

      // コメントを除去したコンテンツ
      const cleanContent = this.removeComments(content);

      // 関数定義を抽出（複数パターン対応）
      const funcMatches = [
        ...cleanContent.matchAll(this.patterns.functionDef),
        ...cleanContent.matchAll(this.patterns.constFunctionDef),
        ...cleanContent.matchAll(this.patterns.methodDef)
      ];

      for (const match of funcMatches) {
        const funcName = match[1];
        
        // 無効な関数名を除外
        if (this.isValidFunctionName(funcName)) {
          fileInfo.functions.add(funcName);
          
          if (!this.functions.has(funcName)) {
            this.functions.set(funcName, new Set());
          }
          this.functions.get(funcName).add(filePath);
        }
      }

      // HTMLファイルの場合、インラインJavaScriptも解析
      if (path.extname(filePath) === '.html') {
        this.extractHtmlInlineFunctions(cleanContent, fileInfo, filePath);
      }
    }
    
    console.log(`🔧 ${this.functions.size} 個の関数を発見しました`);
  }

  /**
   * HTMLファイル内のインラインJavaScript関数を抽出
   */
  extractHtmlInlineFunctions(content, fileInfo, filePath) {
    // <script>タグ内のJavaScriptを抽出
    const scriptTags = content.match(/<script[^>]*>([\s\S]*?)<\/script>/gi) || [];
    
    for (const scriptTag of scriptTags) {
      const jsContent = scriptTag.replace(/<script[^>]*>|<\/script>/gi, '');
      
      // JavaScript内の関数定義を抽出
      const funcMatches = [
        ...jsContent.matchAll(this.patterns.functionDef),
        ...jsContent.matchAll(this.patterns.constFunctionDef)
      ];

      for (const match of funcMatches) {
        const funcName = match[1];
        if (this.isValidFunctionName(funcName)) {
          fileInfo.functions.add(funcName);
          
          if (!this.functions.has(funcName)) {
            this.functions.set(funcName, new Set());
          }
          this.functions.get(funcName).add(filePath);
        }
      }
    }
  }

  /**
   * コメントを除去
   */
  removeComments(content) {
    return content.replace(this.patterns.commentPattern, '');
  }

  /**
   * 有効な関数名かどうかをチェック
   */
  isValidFunctionName(funcName) {
    // 予約語や不適切な名前を除外
    const reservedWords = new Set([
      'if', 'for', 'while', 'switch', 'try', 'catch', 'finally',
      'return', 'var', 'let', 'const', 'function', 'class',
      'true', 'false', 'null', 'undefined', 'this', 'new'
    ]);
    
    return !reservedWords.has(funcName) && 
           /^[a-zA-Z_$][a-zA-Z0-9_$]*$/.test(funcName) &&
           funcName.length > 1;
  }

  /**
   * 依存関係を解析（強化版）
   */
  async analyzeDependencies() {
    for (const filePath of this.files) {
      const fullPath = path.join(this.srcDir, filePath);
      const content = fs.readFileSync(fullPath, 'utf-8');
      const fileInfo = this.dependencies.get(filePath);

      // コメントを除去したコンテンツで解析
      const cleanContent = this.removeComments(content);

      // include() 依存関係
      const includeMatches = [...cleanContent.matchAll(this.patterns.include)];
      for (const match of includeMatches) {
        const includedFile = this.findFileByName(match[1]);
        if (includedFile) {
          fileInfo.dependencies.add(includedFile);
          this.dependencies.get(includedFile).dependents.add(filePath);
        }
      }

      // google.script.run 関数呼び出し
      const gasMatches = [...cleanContent.matchAll(this.patterns.gasCall)];
      for (const match of gasMatches) {
        const funcName = match[1];
        this.usedFunctions.add(funcName);
      }

      // google.script.host 関数呼び出し
      const hostMatches = [...cleanContent.matchAll(this.patterns.hostCall)];
      for (const match of hostMatches) {
        const funcName = match[1];
        this.usedFunctions.add(funcName);
      }

      // HTMLファイル内のイベントハンドラー
      if (path.extname(filePath) === '.html') {
        const htmlMatches = [...cleanContent.matchAll(this.patterns.htmlJsCall)];
        for (const match of htmlMatches) {
          const funcName = match[1];
          if (this.isValidFunctionCall(funcName)) {
            this.usedFunctions.add(funcName);
          }
        }
      }

      // トリガー関数の検出
      const triggerMatches = [...cleanContent.matchAll(this.patterns.triggerFunction)];
      for (const match of triggerMatches) {
        const funcName = match[1];
        this.usedFunctions.add(funcName);
        // トリガー関数は自動的にエントリーポイントに追加
        this.entryPoints.add(funcName);
      }

      // 動的関数呼び出しの検出
      const dynamicMatches = [...cleanContent.matchAll(this.patterns.dynamicCall)];
      for (const match of dynamicMatches) {
        const funcName = match[1];
        this.usedFunctions.add(funcName);
      }

      // 直接関数呼び出し（改良版）
      const directMatches = [...cleanContent.matchAll(this.patterns.directCall)];
      for (const match of directMatches) {
        const funcName = match[1];
        if (this.isValidFunctionCall(funcName)) {
          this.usedFunctions.add(funcName);
        }
      }

      // GAS API使用状況の検出（情報収集用）
      const apiMatches = [...cleanContent.matchAll(this.patterns.gasApiUsage)];
      if (apiMatches.length > 0) {
        fileInfo.usesGasApi = true;
      }
    }
  }

  /**
   * 未使用のファイルと関数を特定
   */
  async identifyUnused() {
    // エントリーポイントから到達可能なファイル・関数をマーク
    this.markReachableFromEntryPoints();

    // include()で参照されているファイルをマーク
    this.markIncludedFiles();

    console.log(`📊 使用中: ${this.usedFiles.size}ファイル, ${this.usedFunctions.size}関数`);
  }

  /**
   * エントリーポイントから到達可能な要素をマーク
   */
  markReachableFromEntryPoints() {
    const visited = new Set();
    
    // エントリーポイント関数を含むファイルを特定
    for (const [funcName, fileSet] of this.functions.entries()) {
      if (this.entryPoints.has(funcName)) {
        this.usedFunctions.add(funcName);
        for (const filePath of fileSet) {
          this.markFileAsUsed(filePath, visited);
        }
      }
    }
  }

  /**
   * ファイルを使用済みとしてマークし、依存関係を再帰的に追跡
   */
  markFileAsUsed(filePath, visited = new Set()) {
    if (visited.has(filePath)) return;
    visited.add(filePath);
    
    this.usedFiles.add(filePath);
    const fileInfo = this.dependencies.get(filePath);
    
    if (fileInfo) {
      // ファイル内の関数を使用済みとしてマーク
      for (const funcName of fileInfo.functions) {
        this.usedFunctions.add(funcName);
      }
      
      // 依存関係を再帰的にマーク
      for (const depFile of fileInfo.dependencies) {
        this.markFileAsUsed(depFile, visited);
      }
    }
  }

  /**
   * include()で参照されているファイルをマーク
   */
  markIncludedFiles() {
    for (const [filePath, fileInfo] of this.dependencies.entries()) {
      if (fileInfo.dependencies.size > 0) {
        this.usedFiles.add(filePath);
        for (const depFile of fileInfo.dependencies) {
          this.usedFiles.add(depFile);
        }
      }
    }
  }

  /**
   * ファイル名からファイルパスを検索
   */
  findFileByName(name) {
    for (const filePath of this.files) {
      const baseName = path.basename(filePath, path.extname(filePath));
      if (baseName === name) {
        return filePath;
      }
    }
    return null;
  }

  /**
   * ファイルタイプを判定
   */
  getFileType(filePath) {
    const ext = path.extname(filePath);
    switch (ext) {
      case '.gs': return 'gas-backend';
      case '.html': return 'html-frontend';
      case '.json': return 'config';
      default: return 'unknown';
    }
  }

  /**
   * 有効な関数呼び出しかどうかチェック
   */
  isValidFunctionCall(funcName) {
    // JavaScriptの組み込み関数やキーワードを除外
    const builtins = new Set([
      'console', 'if', 'for', 'while', 'switch', 'try', 'catch',
      'return', 'var', 'let', 'const', 'function', 'class',
      'parseInt', 'parseFloat', 'isNaN', 'isFinite'
    ]);
    
    return !builtins.has(funcName) && 
           /^[a-zA-Z_$][a-zA-Z0-9_$]*$/.test(funcName) &&
           funcName.length > 1;
  }

  /**
   * 解析レポートを生成
   */
  generateReport() {
    const unusedFiles = [...this.files].filter(f => !this.usedFiles.has(f));
    const unusedFunctions = [...this.functions.keys()].filter(f => !this.usedFunctions.has(f));

    return {
      summary: {
        totalFiles: this.files.size,
        usedFiles: this.usedFiles.size,
        unusedFiles: unusedFiles.length,
        totalFunctions: this.functions.size,
        usedFunctions: this.usedFunctions.size,
        unusedFunctions: unusedFunctions.length
      },
      details: {
        unusedFiles: unusedFiles.map(f => ({
          path: f,
          size: this.dependencies.get(f)?.size || 0,
          type: this.dependencies.get(f)?.type || 'unknown'
        })),
        unusedFunctions: unusedFunctions.map(f => ({
          name: f,
          definedIn: [...(this.functions.get(f) || [])]
        })),
        usedFiles: [...this.usedFiles],
        usedFunctions: [...this.usedFunctions]
      },
      dependencies: Object.fromEntries(
        [...this.dependencies.entries()].map(([path, info]) => [
          path,
          {
            ...info,
            functions: [...info.functions],
            dependencies: [...info.dependencies],
            dependents: [...info.dependents]
          }
        ])
      )
    };
  }
}

// スクリプトが直接実行された場合の処理
if (require.main === module) {
  async function main() {
    try {
      const analyzer = new GasDependencyAnalyzer('./src');
      const report = await analyzer.analyze();
      
      console.log('\n📊 解析結果:');
      console.log(`総ファイル数: ${report.summary.totalFiles}`);
      console.log(`使用中ファイル: ${report.summary.usedFiles}`);
      console.log(`未使用ファイル: ${report.summary.unusedFiles}`);
      console.log(`総関数数: ${report.summary.totalFunctions}`);
      console.log(`使用中関数: ${report.summary.usedFunctions}`);
      console.log(`未使用関数: ${report.summary.unusedFunctions}`);
      
      if (report.details.unusedFiles.length > 0) {
        console.log('\n🗑️ 未使用ファイル:');
        report.details.unusedFiles.forEach(file => {
          console.log(`  - ${file.path} (${file.size} bytes)`);
        });
      }
      
      if (report.details.unusedFunctions.length > 0) {
        console.log('\n🔧 未使用関数:');
        report.details.unusedFunctions.forEach(func => {
          console.log(`  - ${func.name} (定義: ${func.definedIn.join(', ')})`);
        });
      }
      
      // 詳細レポートをJSONで出力
      const reportPath = './scripts/analysis-report.json';
      fs.writeFileSync(reportPath, JSON.stringify(report, null, 2));
      console.log(`\n📝 詳細レポートを保存: ${reportPath}`);
      
    } catch (error) {
      console.error('❌ エラー:', error.message);
      process.exit(1);
    }
  }
  
  main();
}

module.exports = GasDependencyAnalyzer;