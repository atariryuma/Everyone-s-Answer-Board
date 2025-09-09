#!/usr/bin/env node

/**
 * 徹底的なCLAUDE.md準拠監査スクリプト
 * システム全体の設計原則遵守状況を詳細分析
 */

const fs = require('fs');
const path = require('path');

const SRC_DIR = path.join(__dirname, '../src');
const CLAUDE_MD_PATH = path.join(__dirname, '../CLAUDE.md');

class ComprehensiveClaudeMdAuditor {
  constructor() {
    this.violations = [];
    this.compliance = [];
    this.metrics = {
      totalFiles: 0,
      totalLines: 0,
      violationCount: 0,
      complianceScore: 0,
      architecturalIssues: 0,
      performanceIssues: 0,
      securityIssues: 0
    };
    this.claudeMdRules = this.loadClaudeMdRules();
    this.functionUsage = new Map();
    this.dataFlows = new Map();
  }

  loadClaudeMdRules() {
    if (!fs.existsSync(CLAUDE_MD_PATH)) {
      console.warn('⚠️  CLAUDE.md not found, using built-in rules');
      return this.getBuiltinRules();
    }

    const claudeMdContent = fs.readFileSync(CLAUDE_MD_PATH, 'utf8');
    return this.parseClaudeMdRules(claudeMdContent);
  }

  getBuiltinRules() {
    return {
      // configJSON中心型データベーススキーマ
      database: {
        required: ['userId', 'userEmail', 'isActive', 'configJson', 'lastModified'],
        forbidden: ['spreadsheetId', 'sheetName', 'parsedConfig', 'createdAt'],
        configJsonFields: ['spreadsheetId', 'sheetName', 'formUrl', 'columnMapping']
      },
      // アーキテクチャ原則
      architecture: {
        namespaces: ['App', 'DB', 'SecurityValidator', 'ConfigManager'],
        entryPoints: ['doGet', 'doPost'],
        dataFlow: 'configJSON -> config -> usage'
      },
      // パフォーマンス原則
      performance: {
        batchOperations: ['getValues', 'setValues'],
        cacheUsage: ['CacheService', 'PropertiesService'],
        jsonOperations: ['JSON.parse(userInfo.configJson)']
      },
      // セキュリティ原則
      security: {
        validation: ['SecurityValidator'],
        safePatterns: ['Object.freeze', 'const']
      }
    };
  }

  parseClaudeMdRules(content) {
    // CLAUDE.mdから実際のルールを抽出（簡略版）
    const rules = this.getBuiltinRules();
    
    // configJSON中心型の検出
    if (content.includes('configJSON中心型')) {
      rules.database.emphasis = 'configJSON中心設計';
    }
    
    // 5フィールド構造の検出
    if (content.includes('5フィールド構造')) {
      rules.database.strictMode = true;
    }

    return rules;
  }

  async audit() {
    console.log('🔍 徹底的CLAUDE.md準拠監査開始...\n');
    
    // 1. 基本メトリクス収集
    await this.collectBasicMetrics();
    
    // 2. データベーススキーマ監査
    await this.auditDatabaseSchema();
    
    // 3. アーキテクチャパターン監査
    await this.auditArchitecturePatterns();
    
    // 4. パフォーマンス最適化監査
    await this.auditPerformanceOptimizations();
    
    // 5. セキュリティ準拠監査
    await this.auditSecurityCompliance();
    
    // 6. 関数効率性監査
    await this.auditFunctionEfficiency();
    
    // 7. データフロー監査
    await this.auditDataFlows();
    
    // 8. 統計計算
    this.calculateComplianceMetrics();
    
    // 9. レポート生成
    this.generateComprehensiveReport();
    
    return {
      violations: this.violations,
      compliance: this.compliance,
      metrics: this.metrics
    };
  }

  async collectBasicMetrics() {
    const files = this.getFiles(SRC_DIR, ['.gs', '.html']);
    this.metrics.totalFiles = files.length;
    
    for (const filePath of files) {
      const content = fs.readFileSync(filePath, 'utf8');
      this.metrics.totalLines += content.split('\n').length;
    }
    
    console.log(`📊 基本メトリクス: ${this.metrics.totalFiles}ファイル, ${this.metrics.totalLines}行`);
  }

  async auditDatabaseSchema() {
    console.log('🗃️  データベーススキーマ監査...');
    
    const files = this.getFiles(SRC_DIR, ['.gs']);
    
    for (const filePath of files) {
      const content = fs.readFileSync(filePath, 'utf8');
      const relativePath = path.relative(SRC_DIR, filePath);
      
      // 5フィールド構造違反チェック
      const forbiddenDbFields = [
        /userInfo\.spreadsheetId(?!\s*=)/g,
        /userInfo\.sheetName(?!\s*=)/g,
        /userInfo\.parsedConfig(?!\s*=)/g,
        /userInfo\.createdAt(?!\s*=)/g,
        /userInfo\.formUrl(?!\s*=)/g
      ];
      
      for (const pattern of forbiddenDbFields) {
        let match;
        while ((match = pattern.exec(content)) !== null) {
          this.violations.push({
            type: 'DATABASE_SCHEMA_VIOLATION',
            severity: 'ERROR',
            file: relativePath,
            line: this.getLineNumber(content, match.index),
            pattern: match[0],
            message: `5フィールド構造違反: ${match[0]} は禁止。configJSON経由でアクセスしてください`,
            rule: 'CLAUDE.md 5フィールド構造'
          });
          this.metrics.architecturalIssues++;
        }
      }
      
      // configJSON中心設計の遵守チェック
      const configJsonUsage = content.match(/JSON\.parse\(userInfo\.configJson/g) || [];
      const directDbAccess = content.match(/userInfo\.\w+/g) || [];
      
      if (configJsonUsage.length > 0) {
        this.compliance.push({
          type: 'CONFIG_JSON_USAGE',
          file: relativePath,
          count: configJsonUsage.length,
          message: `configJSON中心設計遵守: ${configJsonUsage.length}箇所でJSON.parse使用`
        });
      }
    }
  }

  async auditArchitecturePatterns() {
    console.log('🏗️  アーキテクチャパターン監査...');
    
    const files = this.getFiles(SRC_DIR, ['.gs']);
    
    // 名前空間使用状況チェック
    const expectedNamespaces = ['App', 'DB', 'ConfigManager', 'SecurityValidator'];
    const namespaceUsage = new Map();
    
    for (const filePath of files) {
      const content = fs.readFileSync(filePath, 'utf8');
      const relativePath = path.relative(SRC_DIR, filePath);
      
      for (const namespace of expectedNamespaces) {
        const pattern = new RegExp(`\\b${namespace}\\.\\w+`, 'g');
        const matches = content.match(pattern) || [];
        
        if (!namespaceUsage.has(namespace)) {
          namespaceUsage.set(namespace, 0);
        }
        namespaceUsage.set(namespace, namespaceUsage.get(namespace) + matches.length);
      }
      
      // グローバル関数の過度な使用チェック
      const globalFunctionPattern = /^function\s+(\w+)\s*\(/gm;
      const globalFunctions = [];
      let match;
      
      while ((match = globalFunctionPattern.exec(content)) !== null) {
        globalFunctions.push(match[1]);
      }
      
      if (globalFunctions.length > 10) {
        this.violations.push({
          type: 'ARCHITECTURE_VIOLATION',
          severity: 'WARNING',
          file: relativePath,
          message: `グローバル関数過多: ${globalFunctions.length}個（推奨: 10個以下）`,
          details: globalFunctions.slice(0, 5).join(', ') + (globalFunctions.length > 5 ? '...' : ''),
          rule: 'CLAUDE.md モジュール化原則'
        });
      }
    }
    
    // 名前空間使用統計
    console.log('📋 名前空間使用状況:');
    namespaceUsage.forEach((count, namespace) => {
      console.log(`  ${namespace}: ${count}回`);
      if (count > 0) {
        this.compliance.push({
          type: 'NAMESPACE_USAGE',
          namespace,
          count,
          message: `${namespace}名前空間を${count}回使用`
        });
      }
    });
  }

  async auditPerformanceOptimizations() {
    console.log('⚡ パフォーマンス最適化監査...');
    
    const files = this.getFiles(SRC_DIR, ['.gs']);
    
    for (const filePath of files) {
      const content = fs.readFileSync(filePath, 'utf8');
      const relativePath = path.relative(SRC_DIR, filePath);
      
      // 非効率なループ内API呼び出しチェック
      const inefficientPatterns = [
        {
          pattern: /for\s*\([^}]+\{[^}]*\.getRange\([^}]*\}/g,
          message: 'ループ内でgetRange()使用 → バッチ処理に変更推奨'
        },
        {
          pattern: /for\s*\([^}]+\{[^}]*\.setValue\([^}]*\}/g,
          message: 'ループ内でsetValue()使用 → setValues()に変更推奨'
        },
        {
          pattern: /\.getValue\(\)/g,
          message: 'getValue()使用 → getValues()推奨（複数セル取得時）'
        }
      ];
      
      for (const {pattern, message} of inefficientPatterns) {
        const matches = content.match(pattern) || [];
        if (matches.length > 0) {
          this.violations.push({
            type: 'PERFORMANCE_VIOLATION',
            severity: 'WARNING',
            file: relativePath,
            count: matches.length,
            message: `${message} (${matches.length}箇所)`,
            rule: 'CLAUDE.md パフォーマンス最適化'
          });
          this.metrics.performanceIssues += matches.length;
        }
      }
      
      // バッチ処理の使用をポジティブ評価
      const batchPatterns = [
        /\.getValues\(\)/g,
        /\.setValues\(\)/g,
        /CacheService\./g,
        /PropertiesService\./g
      ];
      
      for (const pattern of batchPatterns) {
        const matches = content.match(pattern) || [];
        if (matches.length > 0) {
          this.compliance.push({
            type: 'PERFORMANCE_OPTIMIZATION',
            file: relativePath,
            pattern: pattern.source,
            count: matches.length,
            message: `効率的API使用: ${pattern.source.replace(/\\\./g, '.')} を${matches.length}回使用`
          });
        }
      }
    }
  }

  async auditSecurityCompliance() {
    console.log('🔒 セキュリティ準拠監査...');
    
    const files = this.getFiles(SRC_DIR, ['.gs']);
    
    for (const filePath of files) {
      const content = fs.readFileSync(filePath, 'utf8');
      const relativePath = path.relative(SRC_DIR, filePath);
      
      // セキュリティパターンチェック
      const securityIssues = [
        {
          pattern: /var\s+\w+/g,
          message: 'var使用 → const/let推奨',
          severity: 'WARNING'
        },
        {
          pattern: /eval\s*\(/g,
          message: 'eval()使用は危険',
          severity: 'ERROR'
        },
        {
          pattern: /document\.write\s*\(/g,
          message: 'document.write()使用は危険',
          severity: 'ERROR'
        }
      ];
      
      for (const {pattern, message, severity} of securityIssues) {
        const matches = content.match(pattern) || [];
        if (matches.length > 0) {
          this.violations.push({
            type: 'SECURITY_VIOLATION',
            severity,
            file: relativePath,
            count: matches.length,
            message: `${message} (${matches.length}箇所)`,
            rule: 'CLAUDE.md セキュリティ原則'
          });
          this.metrics.securityIssues += matches.length;
        }
      }
      
      // セキュリティベストプラクティス使用チェック
      const securityBestPractices = [
        /Object\.freeze\(/g,
        /SecurityValidator\./g,
        /const\s+\w+\s*=/g
      ];
      
      for (const pattern of securityBestPractices) {
        const matches = content.match(pattern) || [];
        if (matches.length > 0) {
          this.compliance.push({
            type: 'SECURITY_BEST_PRACTICE',
            file: relativePath,
            pattern: pattern.source,
            count: matches.length,
            message: `セキュリティベストプラクティス使用: ${matches.length}箇所`
          });
        }
      }
    }
  }

  async auditFunctionEfficiency() {
    console.log('⚙️  関数効率性監査...');
    
    const files = this.getFiles(SRC_DIR, ['.gs']);
    
    // 関数定義と使用状況を収集
    const functionDefinitions = new Map();
    const functionCalls = new Map();
    
    for (const filePath of files) {
      const content = fs.readFileSync(filePath, 'utf8');
      const relativePath = path.relative(SRC_DIR, filePath);
      
      // 関数定義収集
      const functionPattern = /^(?:async\s+)?function\s+(\w+)\s*\(/gm;
      let match;
      
      while ((match = functionPattern.exec(content)) !== null) {
        const functionName = match[1];
        if (!functionDefinitions.has(functionName)) {
          functionDefinitions.set(functionName, []);
        }
        functionDefinitions.get(functionName).push({
          file: relativePath,
          line: this.getLineNumber(content, match.index)
        });
      }
      
      // 関数呼び出し収集
      const callPattern = /(\w+)\s*\(/g;
      while ((match = callPattern.exec(content)) !== null) {
        const functionName = match[1];
        if (this.isValidFunctionName(functionName)) {
          if (!functionCalls.has(functionName)) {
            functionCalls.set(functionName, 0);
          }
          functionCalls.set(functionName, functionCalls.get(functionName) + 1);
        }
      }
    }
    
    // 未使用関数検出
    functionDefinitions.forEach((definitions, functionName) => {
      const callCount = functionCalls.get(functionName) || 0;
      const isEntryPoint = ['doGet', 'doPost', 'onInstall', 'onOpen'].includes(functionName);
      
      if (callCount === 0 && !isEntryPoint) {
        this.violations.push({
          type: 'FUNCTION_EFFICIENCY_VIOLATION',
          severity: 'WARNING',
          function: functionName,
          files: definitions.map(d => `${d.file}:${d.line}`),
          message: `未使用関数: ${functionName} (定義のみで呼び出しなし)`,
          rule: 'CLAUDE.md 効率性原則'
        });
      } else if (callCount > 0) {
        this.compliance.push({
          type: 'FUNCTION_USAGE',
          function: functionName,
          callCount,
          message: `関数使用: ${functionName} (${callCount}回呼び出し)`
        });
      }
    });
    
    this.functionUsage = functionCalls;
  }

  async auditDataFlows() {
    console.log('🔄 データフロー監査...');
    
    const files = this.getFiles(SRC_DIR, ['.gs']);
    
    for (const filePath of files) {
      const content = fs.readFileSync(filePath, 'utf8');
      const relativePath = path.relative(SRC_DIR, filePath);
      
      // データフローパターン検出
      const dataFlowPatterns = [
        {
          pattern: /userInfo\s*->\s*configJson\s*->\s*config/,
          flow: 'userInfo -> configJson -> config',
          compliant: true
        },
        {
          pattern: /userInfo\.configJson.*JSON\.parse/,
          flow: 'configJSON parsing',
          compliant: true
        },
        {
          pattern: /userInfo\.\w+(?!\.configJson)(?!\.userId)(?!\.userEmail)(?!\.isActive)(?!\.lastModified)/,
          flow: 'Direct userInfo field access',
          compliant: false
        }
      ];
      
      for (const {pattern, flow, compliant} of dataFlowPatterns) {
        const matches = content.match(pattern) || [];
        if (matches.length > 0) {
          if (compliant) {
            this.compliance.push({
              type: 'DATA_FLOW_COMPLIANCE',
              file: relativePath,
              flow,
              count: matches.length,
              message: `適切なデータフロー: ${flow} (${matches.length}箇所)`
            });
          } else {
            this.violations.push({
              type: 'DATA_FLOW_VIOLATION',
              severity: 'WARNING',
              file: relativePath,
              flow,
              count: matches.length,
              message: `不適切なデータフロー: ${flow} (${matches.length}箇所)`,
              rule: 'CLAUDE.md データフロー原則'
            });
          }
        }
      }
    }
  }

  calculateComplianceMetrics() {
    this.metrics.violationCount = this.violations.length;
    const totalChecks = this.violations.length + this.compliance.length;
    
    if (totalChecks > 0) {
      this.metrics.complianceScore = Math.round((this.compliance.length / totalChecks) * 100);
    }
  }

  generateComprehensiveReport() {
    console.log('\n' + '='.repeat(100));
    console.log('📋 徹底的CLAUDE.md準拠監査レポート');
    console.log('='.repeat(100));

    // 総合スコア
    console.log(`\n🎯 総合コンプライアンススコア: ${this.metrics.complianceScore}%`);
    console.log(`📊 基本統計:`);
    console.log(`  ファイル数: ${this.metrics.totalFiles}`);
    console.log(`  総行数: ${this.metrics.totalLines}`);
    console.log(`  違反数: ${this.metrics.violationCount}`);
    console.log(`  準拠項目: ${this.compliance.length}`);

    // カテゴリ別問題数
    console.log(`\n🔍 問題カテゴリ別統計:`);
    console.log(`  アーキテクチャ問題: ${this.metrics.architecturalIssues}件`);
    console.log(`  パフォーマンス問題: ${this.metrics.performanceIssues}件`);
    console.log(`  セキュリティ問題: ${this.metrics.securityIssues}件`);

    // 重大な違反
    const criticalViolations = this.violations.filter(v => v.severity === 'ERROR');
    if (criticalViolations.length > 0) {
      console.log(`\n❌ 重大違反 (${criticalViolations.length}件):`);
      criticalViolations.forEach((violation, index) => {
        console.log(`  ${index + 1}. ${violation.message}`);
        console.log(`     📁 ${violation.file}${violation.line ? `:${violation.line}` : ''}`);
        console.log(`     📋 ルール: ${violation.rule}`);
      });
    }

    // 警告
    const warnings = this.violations.filter(v => v.severity === 'WARNING');
    if (warnings.length > 0) {
      console.log(`\n⚠️  警告 (${warnings.length}件):`);
      warnings.slice(0, 10).forEach((warning, index) => {
        console.log(`  ${index + 1}. ${warning.message}`);
        console.log(`     📁 ${warning.file}`);
      });
      if (warnings.length > 10) {
        console.log(`  ... 他 ${warnings.length - 10} 件`);
      }
    }

    // 準拠状況ハイライト
    const complianceTypes = new Map();
    this.compliance.forEach(item => {
      if (!complianceTypes.has(item.type)) {
        complianceTypes.set(item.type, 0);
      }
      complianceTypes.set(item.type, complianceTypes.get(item.type) + 1);
    });

    if (complianceTypes.size > 0) {
      console.log(`\n✅ 準拠ハイライト:`);
      complianceTypes.forEach((count, type) => {
        console.log(`  ${type}: ${count}項目`);
      });
    }

    // 関数効率性サマリー
    console.log(`\n⚙️  関数効率性サマリー:`);
    console.log(`  定義済み関数: ${this.functionUsage.size}個`);
    const totalCalls = Array.from(this.functionUsage.values()).reduce((sum, calls) => sum + calls, 0);
    console.log(`  総呼び出し数: ${totalCalls}回`);
    
    // よく使用される関数トップ5
    const topFunctions = Array.from(this.functionUsage.entries())
      .sort(([,a], [,b]) => b - a)
      .slice(0, 5);
    
    if (topFunctions.length > 0) {
      console.log(`  📈 よく使用される関数:`);
      topFunctions.forEach(([name, count], index) => {
        console.log(`    ${index + 1}. ${name}: ${count}回`);
      });
    }

    // 推奨アクション
    console.log(`\n🚀 推奨アクション:`);
    if (criticalViolations.length > 0) {
      console.log(`  1. 重大違反 ${criticalViolations.length} 件の即座修正`);
    }
    if (this.metrics.performanceIssues > 0) {
      console.log(`  2. パフォーマンス問題 ${this.metrics.performanceIssues} 件の最適化`);
    }
    if (warnings.length > 0) {
      console.log(`  3. 警告 ${warnings.length} 件の段階的改善`);
    }
    if (this.metrics.complianceScore < 80) {
      console.log(`  4. コンプライアンススコア向上（現在${this.metrics.complianceScore}% → 目標80%以上）`);
    }

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

  isValidFunctionName(name) {
    return /^[a-zA-Z_$][a-zA-Z0-9_$]*$/.test(name) && 
           name.length > 1 && 
           !['if', 'for', 'while', 'function', 'const', 'let', 'var', 'console'].includes(name);
  }
}

// メイン実行
async function main() {
  const auditor = new ComprehensiveClaudeMdAuditor();
  
  try {
    const result = await auditor.audit();
    
    // 結果をJSONファイルに保存
    const outputPath = path.join(__dirname, '../comprehensive-audit-results.json');
    fs.writeFileSync(outputPath, JSON.stringify(result, null, 2));
    console.log(`\n📄 詳細監査結果は ${outputPath} に保存されました`);
    
    // 問題があれば終了コード1
    const hasCriticalIssues = result.violations.some(v => v.severity === 'ERROR');
    process.exit(hasCriticalIssues ? 1 : 0);
    
  } catch (error) {
    console.error('❌ 監査中にエラーが発生しました:', error);
    process.exit(1);
  }
}

if (require.main === module) {
  main();
}

module.exports = { ComprehensiveClaudeMdAuditor };