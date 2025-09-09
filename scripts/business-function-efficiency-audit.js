#!/usr/bin/env node

/**
 * ビジネス関数効率性監査スクリプト
 * 実際のビジネスロジック関数の効率性と使用状況を分析
 */

const fs = require('fs');
const path = require('path');

const SRC_DIR = path.join(__dirname, '../src');

class BusinessFunctionEfficiencyAuditor {
  constructor() {
    this.businessFunctions = new Map();
    this.functionMetrics = new Map();
    this.inefficiencies = [];
    this.optimizations = [];
    this.dataFlowIssues = [];
    
    // ビジネス関数の定義（CLAUDE.mdに基づく）
    this.businessFunctionCategories = {
      // データアクセス層
      database: ['createUser', 'findUserById', 'findUserByEmail', 'updateUser', 'deleteUser', 'getAllUsers'],
      
      // 設定管理
      config: ['getConfig', 'updateConfig', 'saveDraftConfiguration', 'publishApplication', 'getUserConfig'],
      
      // データ処理
      data: ['getData', 'getPublishedSheetData', 'connectDataSource', 'getDataCount', 'refreshBoardData'],
      
      // インタラクション
      interaction: ['addReaction', 'toggleHighlight', 'addReactionBatch'],
      
      // システム管理  
      system: ['checkIsSystemAdmin', 'getCurrentBoardInfoAndUrls', 'executeConfigCleanup'],
      
      // 認証・アクセス制御
      auth: ['verifyAccess', 'handleUserRegistration', 'getCurrentUserInfoSafely'],
      
      // UI・表示
      rendering: ['renderAnswerBoard', 'renderAdminPanel', 'doGet', 'doPost']
    };
  }

  async audit() {
    console.log('⚙️  ビジネス関数効率性監査開始...\n');
    
    // 1. ビジネス関数収集
    await this.collectBusinessFunctions();
    
    // 2. 関数メトリクス分析
    await this.analyzeMetrics();
    
    // 3. データフロー効率性分析
    await this.analyzeDataFlowEfficiency();
    
    // 4. パフォーマンスボトルネック特定
    await this.identifyPerformanceBottlenecks();
    
    // 5. 最適化機会特定
    await this.identifyOptimizationOpportunities();
    
    // 6. レポート生成
    this.generateEfficiencyReport();
    
    return {
      functions: this.businessFunctions,
      metrics: this.functionMetrics,
      inefficiencies: this.inefficiencies,
      optimizations: this.optimizations
    };
  }

  async collectBusinessFunctions() {
    const gsFiles = this.getFiles(SRC_DIR, ['.gs']);
    
    for (const filePath of gsFiles) {
      const content = fs.readFileSync(filePath, 'utf8');
      const relativePath = path.relative(SRC_DIR, filePath);
      
      // 関数定義パターン
      const functionPattern = /^(?:async\s+)?function\s+(\w+)\s*\([^)]*\)\s*\{/gm;
      let match;
      
      while ((match = functionPattern.exec(content)) !== null) {
        const functionName = match[1];
        const category = this.getFunctionCategory(functionName);
        
        if (category) {
          const functionBody = this.extractFunctionBody(content, match.index);
          const metrics = this.calculateFunctionMetrics(functionBody);
          
          this.businessFunctions.set(functionName, {
            file: relativePath,
            line: this.getLineNumber(content, match.index),
            category,
            body: functionBody,
            ...metrics
          });
        }
      }
    }
    
    console.log(`📋 ビジネス関数収集完了: ${this.businessFunctions.size}個の関数`);
  }

  getFunctionCategory(functionName) {
    for (const [category, functions] of Object.entries(this.businessFunctionCategories)) {
      if (functions.includes(functionName)) {
        return category;
      }
    }
    
    // パターンマッチングで追加カテゴリ検出
    if (functionName.includes('Config') || functionName.includes('config')) return 'config';
    if (functionName.includes('User') || functionName.includes('user')) return 'database';
    if (functionName.includes('Data') || functionName.includes('data')) return 'data';
    if (functionName.startsWith('get') || functionName.startsWith('set')) return 'data';
    if (functionName.startsWith('render') || functionName.startsWith('do')) return 'rendering';
    
    return null; // ビジネス関数でない
  }

  extractFunctionBody(content, startIndex) {
    let braceCount = 0;
    let i = startIndex;
    let foundFirstBrace = false;
    
    while (i < content.length) {
      if (content[i] === '{') {
        braceCount++;
        foundFirstBrace = true;
      } else if (content[i] === '}') {
        braceCount--;
        if (foundFirstBrace && braceCount === 0) {
          return content.substring(startIndex, i + 1);
        }
      }
      i++;
    }
    
    return content.substring(startIndex, Math.min(startIndex + 1000, content.length));
  }

  calculateFunctionMetrics(body) {
    return {
      lines: body.split('\n').length,
      complexity: this.calculateComplexity(body),
      apiCalls: this.countApiCalls(body),
      dbAccesses: this.countDbAccesses(body),
      cacheUsage: this.countCacheUsage(body),
      errorHandling: this.hasErrorHandling(body),
      configJsonUsage: this.countConfigJsonUsage(body)
    };
  }

  calculateComplexity(body) {
    // 単純な複雑度計算（if, for, while, catch, case文の数）
    const patterns = [/\bif\s*\(/g, /\bfor\s*\(/g, /\bwhile\s*\(/g, /\bcatch\s*\(/g, /\bcase\s+/g];
    let complexity = 1; // ベース複雑度
    
    for (const pattern of patterns) {
      const matches = body.match(pattern) || [];
      complexity += matches.length;
    }
    
    return complexity;
  }

  countApiCalls(body) {
    const apiPatterns = [
      /SpreadsheetApp\./g,
      /DriveApp\./g, 
      /UrlFetchApp\./g,
      /PropertiesService\./g,
      /CacheService\./g,
      /\.getRange\(/g,
      /\.getValues\(\)/g,
      /\.setValues\(/g
    ];
    
    let count = 0;
    for (const pattern of apiPatterns) {
      const matches = body.match(pattern) || [];
      count += matches.length;
    }
    
    return count;
  }

  countDbAccesses(body) {
    const dbPatterns = [
      /DB\.\w+/g,
      /findUserBy/g,
      /createUser/g,
      /updateUser/g,
      /deleteUser/g
    ];
    
    let count = 0;
    for (const pattern of dbPatterns) {
      const matches = body.match(pattern) || [];
      count += matches.length;
    }
    
    return count;
  }

  countCacheUsage(body) {
    const cachePatterns = [
      /cacheManager\./g,
      /CacheService\./g,
      /cache\.get/g,
      /cache\.put/g
    ];
    
    let count = 0;
    for (const pattern of cachePatterns) {
      const matches = body.match(pattern) || [];
      count += matches.length;
    }
    
    return count;
  }

  hasErrorHandling(body) {
    return body.includes('try') && body.includes('catch');
  }

  countConfigJsonUsage(body) {
    const configPatterns = [
      /JSON\.parse\(userInfo\.configJson/g,
      /ConfigManager\./g,
      /getUserConfig/g,
      /updateConfig/g
    ];
    
    let count = 0;
    for (const pattern of configPatterns) {
      const matches = body.match(pattern) || [];
      count += matches.length;
    }
    
    return count;
  }

  async analyzeMetrics() {
    console.log('📊 関数メトリクス分析中...');
    
    for (const [functionName, data] of this.businessFunctions.entries()) {
      // 複雑度の評価
      if (data.complexity > 15) {
        this.inefficiencies.push({
          type: 'HIGH_COMPLEXITY',
          function: functionName,
          file: data.file,
          severity: 'WARNING',
          value: data.complexity,
          message: `関数の複雑度が高い: ${data.complexity} (推奨: 10以下)`
        });
      }
      
      // 長大関数の検出
      if (data.lines > 100) {
        this.inefficiencies.push({
          type: 'LONG_FUNCTION',
          function: functionName,
          file: data.file,
          severity: 'WARNING',
          value: data.lines,
          message: `関数が長すぎる: ${data.lines}行 (推奨: 50行以下)`
        });
      }
      
      // エラーハンドリングの不備
      if (data.apiCalls > 0 && !data.errorHandling) {
        this.inefficiencies.push({
          type: 'MISSING_ERROR_HANDLING',
          function: functionName,
          file: data.file,
          severity: 'ERROR',
          message: `API呼び出し${data.apiCalls}回あるがエラーハンドリングなし`
        });
      }
      
      // configJSON使用の良好事例
      if (data.configJsonUsage > 0) {
        this.optimizations.push({
          type: 'CONFIG_JSON_COMPLIANCE',
          function: functionName,
          file: data.file,
          value: data.configJsonUsage,
          message: `configJSON中心設計に準拠: ${data.configJsonUsage}箇所`
        });
      }
    }
  }

  async analyzeDataFlowEfficiency() {
    console.log('🔄 データフロー効率性分析中...');
    
    const gsFiles = this.getFiles(SRC_DIR, ['.gs']);
    
    for (const filePath of gsFiles) {
      const content = fs.readFileSync(filePath, 'utf8');
      const relativePath = path.relative(SRC_DIR, filePath);
      
      // 非効率なデータアクセスパターン
      const inefficientPatterns = [
        {
          pattern: /userInfo\.(?!configJson)(?!userId)(?!userEmail)(?!isActive)(?!lastModified)\w+/g,
          message: '5フィールド構造違反: 直接フィールドアクセス'
        },
        {
          pattern: /for\s*\([^}]*DB\.\w+[^}]*\)/g,
          message: 'ループ内でDB操作: バッチ処理を検討'
        },
        {
          pattern: /JSON\.parse\(.*JSON\.stringify\(/g,
          message: '無駄なJSON変換: 不要なparse/stringify'
        }
      ];
      
      for (const {pattern, message} of inefficientPatterns) {
        let match;
        while ((match = pattern.exec(content)) !== null) {
          this.dataFlowIssues.push({
            type: 'DATA_FLOW_INEFFICIENCY',
            file: relativePath,
            line: this.getLineNumber(content, match.index),
            pattern: match[0],
            message,
            severity: 'WARNING'
          });
        }
      }
    }
  }

  async identifyPerformanceBottlenecks() {
    console.log('🚀 パフォーマンスボトルネック特定中...');
    
    // 高API呼び出し関数
    const highApiCallFunctions = Array.from(this.businessFunctions.entries())
      .filter(([, data]) => data.apiCalls > 5)
      .sort(([,a], [,b]) => b.apiCalls - a.apiCalls);
    
    for (const [functionName, data] of highApiCallFunctions) {
      this.inefficiencies.push({
        type: 'HIGH_API_USAGE',
        function: functionName,
        file: data.file,
        severity: 'WARNING',
        value: data.apiCalls,
        message: `API呼び出しが多い: ${data.apiCalls}回 (バッチ処理を検討)`
      });
    }
    
    // キャッシュ未使用で重い処理
    const heavyNonCachedFunctions = Array.from(this.businessFunctions.entries())
      .filter(([, data]) => data.apiCalls > 3 && data.cacheUsage === 0);
    
    for (const [functionName, data] of heavyNonCachedFunctions) {
      this.optimizations.push({
        type: 'CACHE_OPPORTUNITY',
        function: functionName,
        file: data.file,
        value: data.apiCalls,
        message: `キャッシュ導入推奨: ${data.apiCalls}回のAPI呼び出し`
      });
    }
  }

  async identifyOptimizationOpportunities() {
    console.log('💡 最適化機会特定中...');
    
    // カテゴリ別分析
    const categoryMetrics = new Map();
    
    for (const [functionName, data] of this.businessFunctions.entries()) {
      if (!categoryMetrics.has(data.category)) {
        categoryMetrics.set(data.category, {
          count: 0,
          totalComplexity: 0,
          totalApiCalls: 0,
          totalLines: 0,
          functions: []
        });
      }
      
      const metrics = categoryMetrics.get(data.category);
      metrics.count++;
      metrics.totalComplexity += data.complexity;
      metrics.totalApiCalls += data.apiCalls;
      metrics.totalLines += data.lines;
      metrics.functions.push(functionName);
    }
    
    // カテゴリ別最適化提案
    for (const [category, metrics] of categoryMetrics.entries()) {
      const avgComplexity = metrics.totalComplexity / metrics.count;
      const avgApiCalls = metrics.totalApiCalls / metrics.count;
      
      if (avgComplexity > 8) {
        this.optimizations.push({
          type: 'CATEGORY_COMPLEXITY_OPTIMIZATION',
          category,
          value: avgComplexity.toFixed(1),
          message: `${category}カテゴリの平均複雑度が高い: ${avgComplexity.toFixed(1)} (リファクタリング推奨)`
        });
      }
      
      if (avgApiCalls > 3) {
        this.optimizations.push({
          type: 'CATEGORY_API_OPTIMIZATION',
          category,
          value: avgApiCalls.toFixed(1),
          message: `${category}カテゴリのAPI使用量が多い: ${avgApiCalls.toFixed(1)} (効率化推奨)`
        });
      }
    }
  }

  generateEfficiencyReport() {
    console.log('\n' + '='.repeat(100));
    console.log('⚙️  ビジネス関数効率性監査レポート');
    console.log('='.repeat(100));

    // サマリー統計
    const totalFunctions = this.businessFunctions.size;
    const totalIssues = this.inefficiencies.length;
    const totalOptimizations = this.optimizations.length;
    
    console.log(`\n📊 サマリー統計:`);
    console.log(`  ビジネス関数数: ${totalFunctions}`);
    console.log(`  効率性問題: ${totalIssues}件`);
    console.log(`  最適化機会: ${totalOptimizations}件`);
    console.log(`  データフロー問題: ${this.dataFlowIssues.length}件`);

    // カテゴリ別統計
    const categories = new Map();
    this.businessFunctions.forEach((data) => {
      if (!categories.has(data.category)) {
        categories.set(data.category, 0);
      }
      categories.set(data.category, categories.get(data.category) + 1);
    });

    console.log(`\n📋 カテゴリ別関数数:`);
    categories.forEach((count, category) => {
      console.log(`  ${category}: ${count}個`);
    });

    // 重要な効率性問題
    const criticalIssues = this.inefficiencies.filter(issue => issue.severity === 'ERROR');
    if (criticalIssues.length > 0) {
      console.log(`\n❌ 重要効率性問題 (${criticalIssues.length}件):`);
      criticalIssues.forEach((issue, index) => {
        console.log(`  ${index + 1}. ${issue.function}(): ${issue.message}`);
        console.log(`     📁 ${issue.file}`);
      });
    }

    // 最適化機会トップ5
    const topOptimizations = this.optimizations
      .filter(opt => opt.function)
      .sort((a, b) => (b.value || 0) - (a.value || 0))
      .slice(0, 5);

    if (topOptimizations.length > 0) {
      console.log(`\n💡 最適化機会トップ5:`);
      topOptimizations.forEach((opt, index) => {
        console.log(`  ${index + 1}. ${opt.function}(): ${opt.message}`);
        console.log(`     📁 ${opt.file}`);
      });
    }

    // パフォーマンスメトリクス
    let totalComplexity = 0;
    let totalApiCalls = 0;
    let totalLines = 0;

    this.businessFunctions.forEach((data) => {
      totalComplexity += data.complexity;
      totalApiCalls += data.apiCalls;
      totalLines += data.lines;
    });

    console.log(`\n📈 パフォーマンスメトリクス:`);
    console.log(`  平均複雑度: ${(totalComplexity / totalFunctions).toFixed(1)}`);
    console.log(`  平均API呼び出し数: ${(totalApiCalls / totalFunctions).toFixed(1)}`);
    console.log(`  平均関数長: ${Math.round(totalLines / totalFunctions)}行`);

    // configJSON準拠状況
    const configCompliantFunctions = Array.from(this.businessFunctions.values())
      .filter(data => data.configJsonUsage > 0).length;
    const complianceRate = Math.round((configCompliantFunctions / totalFunctions) * 100);

    console.log(`\n🎯 CLAUDE.md準拠状況:`);
    console.log(`  configJSON使用関数: ${configCompliantFunctions}/${totalFunctions} (${complianceRate}%)`);

    // 推奨アクション
    console.log(`\n🚀 推奨アクション:`);
    if (criticalIssues.length > 0) {
      console.log(`  1. 重要効率性問題 ${criticalIssues.length} 件の即座修正`);
    }
    if (totalIssues > criticalIssues.length) {
      console.log(`  2. 効率性警告 ${totalIssues - criticalIssues.length} 件の段階的改善`);
    }
    if (complianceRate < 80) {
      console.log(`  3. configJSON準拠率向上（現在${complianceRate}% → 目標80%以上）`);
    }
    console.log(`  4. パフォーマンス最適化機会 ${totalOptimizations} 件の検討`);

    console.log('\n' + '='.repeat(100));
  }

  // ヘルパーメソッド
  getFiles(dir, extensions) {
    const files = [];
    const items = fs.readdirSync(dir);
    
    for (const item of items) {
      const fullPath = path.join(dir, item);
      const stat = fs.statSync(fullPath);
      
      if (stat.isDirectory()) {
        files.push(...this.getFiles(fullPath, extensions));
      } else if (extensions.some(ext => item.endsWith(ext))) {
        files.push(fullPath);
      }
    }
    
    return files;
  }

  getLineNumber(content, index) {
    return content.substring(0, index).split('\n').length;
  }
}

// メイン実行
async function main() {
  const auditor = new BusinessFunctionEfficiencyAuditor();
  
  try {
    const result = await auditor.audit();
    
    // 結果をJSONファイルに保存
    const outputPath = path.join(__dirname, '../business-efficiency-results.json');
    fs.writeFileSync(outputPath, JSON.stringify(result, null, 2));
    console.log(`\n📄 詳細結果は ${outputPath} に保存されました`);
    
  } catch (error) {
    console.error('❌ 監査中にエラーが発生しました:', error);
    process.exit(1);
  }
}

if (require.main === module) {
  main();
}

module.exports = { BusinessFunctionEfficiencyAuditor };