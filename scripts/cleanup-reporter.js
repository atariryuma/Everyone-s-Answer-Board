#!/usr/bin/env node

/**
 * @fileoverview 未使用コードレポート生成器
 * 依存関係解析結果から詳細な未使用コードレポートを生成
 */

const fs = require('fs');
const path = require('path');
const GasDependencyAnalyzer = require('./dependency-analyzer');

class CleanupReporter {
  constructor(srcDir = './src') {
    this.srcDir = path.resolve(srcDir);
    this.reportDir = path.resolve('./scripts/reports');
    this.analyzer = null;
    this.analysisResult = null;
  }

  /**
   * 完全な解析とレポート生成を実行
   */
  async generateReport() {
    console.log('📊 未使用コード解析レポート生成を開始...');
    
    try {
      // 依存関係解析を実行
      this.analyzer = new GasDependencyAnalyzer(this.srcDir);
      this.analysisResult = await this.analyzer.analyze();
      
      // レポートディレクトリを準備
      this.ensureReportDirectory();
      
      // 複数形式のレポートを生成
      const reports = await this.generateAllReports();
      
      console.log('✅ レポート生成完了');
      return reports;
      
    } catch (error) {
      console.error('❌ レポート生成エラー:', error.message);
      throw error;
    }
  }

  /**
   * レポートディレクトリの準備
   */
  ensureReportDirectory() {
    if (!fs.existsSync(this.reportDir)) {
      fs.mkdirSync(this.reportDir, { recursive: true });
    }
  }

  /**
   * 全レポート形式を生成
   */
  async generateAllReports() {
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const reportBaseName = `cleanup-report-${timestamp}`;
    
    const reports = {
      detailed: await this.generateDetailedReport(reportBaseName),
      summary: await this.generateSummaryReport(reportBaseName),
      actionable: await this.generateActionableReport(reportBaseName),
      visualTree: await this.generateDependencyTreeReport(reportBaseName),
      csv: await this.generateCsvReport(reportBaseName)
    };

    return reports;
  }

  /**
   * 詳細レポート生成（JSON形式）
   */
  async generateDetailedReport(baseName) {
    const reportPath = path.join(this.reportDir, `${baseName}-detailed.json`);
    
    const detailedReport = {
      metadata: {
        generatedAt: new Date().toISOString(),
        srcDirectory: this.srcDir,
        analysisEngine: 'GasDependencyAnalyzer v1.0',
        totalAnalysisTime: 'N/A' // TODO: 実測可能にする
      },
      summary: this.analysisResult.summary,
      unusedFiles: this.analysisResult.details.unusedFiles.map(file => ({
        ...file,
        fullPath: path.join(this.srcDir, file.path),
        recommendedAction: this.getFileRecommendation(file),
        riskLevel: this.assessFileDeletionRisk(file),
        dependencies: this.getFileDependencies(file.path)
      })),
      unusedFunctions: this.analysisResult.details.unusedFunctions.map(func => ({
        ...func,
        recommendedAction: this.getFunctionRecommendation(func),
        riskLevel: this.assessFunctionDeletionRisk(func),
        potentialImpact: this.assessFunctionImpact(func)
      })),
      dependencyGraph: this.analysisResult.dependencies,
      entryPoints: [...this.analyzer.entryPoints],
      riskAssessment: this.generateRiskAssessment(),
      cleanupPlan: this.generateCleanupPlan()
    };

    fs.writeFileSync(reportPath, JSON.stringify(detailedReport, null, 2));
    console.log(`📄 詳細レポート: ${reportPath}`);
    
    return { path: reportPath, content: detailedReport };
  }

  /**
   * サマリーレポート生成（Markdown形式）
   */
  async generateSummaryReport(baseName) {
    const reportPath = path.join(this.reportDir, `${baseName}-summary.md`);
    
    const summary = this.analysisResult.summary;
    const unusedFiles = this.analysisResult.details.unusedFiles;
    const unusedFunctions = this.analysisResult.details.unusedFunctions;
    
    const markdown = `# 未使用コード解析レポート

## 📊 概要

- **解析日時**: ${new Date().toLocaleString()}
- **対象ディレクトリ**: \`${this.srcDir}\`
- **総ファイル数**: ${summary.totalFiles}
- **未使用ファイル数**: ${summary.unusedFiles} (${((summary.unusedFiles / summary.totalFiles) * 100).toFixed(1)}%)
- **総関数数**: ${summary.totalFunctions}
- **未使用関数数**: ${summary.unusedFunctions} (${((summary.unusedFunctions / summary.totalFunctions) * 100).toFixed(1)}%)

## 🗑️ 削除可能な未使用ファイル

${unusedFiles.length > 0 ? unusedFiles.map(file => {
  const risk = this.assessFileDeletionRisk(file);
  const riskEmoji = risk === 'low' ? '🟢' : risk === 'medium' ? '🟡' : '🔴';
  return `- ${riskEmoji} \`${file.path}\` (${file.size} bytes, ${file.type})`;
}).join('\n') : '✅ 未使用ファイルはありません'}

## 🔧 削除可能な未使用関数

${unusedFunctions.length > 0 ? unusedFunctions.map(func => {
  const risk = this.assessFunctionDeletionRisk(func);
  const riskEmoji = risk === 'low' ? '🟢' : risk === 'medium' ? '🟡' : '🔴';
  return `- ${riskEmoji} \`${func.name}\` (定義: ${func.definedIn.join(', ')})`;
}).join('\n') : '✅ 未使用関数はありません'}

## ⚠️ 削除前の注意事項

1. **バックアップ作成**: 削除前に必ずバックアップを作成してください
2. **テスト実行**: 削除後にすべてのテストが通ることを確認してください
3. **段階的削除**: まず低リスクの項目から削除を開始してください

## 🚀 推奨削除手順

1. 低リスクファイルの削除
2. 低リスク関数の削除
3. テスト実行・動作確認
4. 中リスクファイル・関数の検討
5. 高リスク項目の慎重な検討

---

*このレポートは自動生成されました。削除前に必ず手動確認を行ってください。*
`;

    fs.writeFileSync(reportPath, markdown);
    console.log(`📄 サマリーレポート: ${reportPath}`);
    
    return { path: reportPath, content: markdown };
  }

  /**
   * アクション可能レポート生成（実行コマンド付き）
   */
  async generateActionableReport(baseName) {
    const reportPath = path.join(this.reportDir, `${baseName}-actions.sh`);
    
    const unusedFiles = this.analysisResult.details.unusedFiles;
    const unusedFunctions = this.analysisResult.details.unusedFunctions;
    
    let script = `#!/bin/bash

# 自動生成された未使用コード削除スクリプト
# 生成日時: ${new Date().toLocaleString()}
# 実行前に必ずバックアップを作成してください！

set -e  # エラー時に停止

echo "🚀 未使用コード削除スクリプトを開始します"
echo "⚠️  実行前にバックアップが作成されていることを確認してください"

# バックアップ確認
read -p "バックアップは作成済みですか？ (y/N): " backup_confirm
if [[ ! "$backup_confirm" =~ ^[Yy]$ ]]; then
    echo "❌ バックアップを作成してから再実行してください"
    exit 1
fi

# 低リスクファイルの削除
echo "🗑️ 低リスクファイルの削除中..."
`;

    // 低リスクファイルの削除コマンド
    const lowRiskFiles = unusedFiles.filter(file => this.assessFileDeletionRisk(file) === 'low');
    if (lowRiskFiles.length > 0) {
      for (const file of lowRiskFiles) {
        script += `echo "削除: ${file.path}"\n`;
        script += `rm -f "${path.join(this.srcDir, file.path)}"\n`;
      }
    } else {
      script += 'echo "低リスクの削除対象ファイルはありません"\n';
    }

    script += `
# テスト実行
echo "🧪 テスト実行中..."
if command -v npm &> /dev/null && [ -f "package.json" ]; then
    npm run test || {
        echo "❌ テストが失敗しました。削除を中止します"
        exit 1
    }
else
    echo "⚠️ npmテストが利用できません。手動でテストを実行してください"
fi

echo "✅ 低リスクファイルの削除が完了しました"
echo "📝 次に中・高リスクファイルを慎重に検討してください："
`;

    // 中・高リスクファイルの警告
    const higherRiskFiles = unusedFiles.filter(file => this.assessFileDeletionRisk(file) !== 'low');
    for (const file of higherRiskFiles) {
      const risk = this.assessFileDeletionRisk(file);
      script += `echo "  ${risk === 'high' ? '🔴' : '🟡'} ${file.path} (${risk}リスク)"\n`;
    }

    script += '\necho "完了！"\n';

    fs.writeFileSync(reportPath, script);
    fs.chmodSync(reportPath, '755'); // 実行権限を付与
    console.log(`📄 実行可能レポート: ${reportPath}`);
    
    return { path: reportPath, content: script };
  }

  /**
   * 依存関係ツリーレポート生成
   */
  async generateDependencyTreeReport(baseName) {
    const reportPath = path.join(this.reportDir, `${baseName}-tree.txt`);
    
    let treeContent = `依存関係ツリー
生成日時: ${new Date().toLocaleString()}
対象: ${this.srcDir}

エントリーポイント:
${[...this.analyzer.entryPoints].map(ep => `  📌 ${ep}`).join('\n')}

ファイル依存関係:
`;

    // 使用中ファイルのツリー構造を生成
    for (const filePath of this.analysisResult.details.usedFiles) {
      const fileInfo = this.analysisResult.dependencies[filePath];
      if (fileInfo) {
        treeContent += `\n📁 ${filePath}\n`;
        treeContent += `  タイプ: ${fileInfo.type}\n`;
        treeContent += `  関数数: ${fileInfo.functions.length}\n`;
        
        if (fileInfo.dependencies.length > 0) {
          treeContent += `  依存先:\n`;
          for (const dep of fileInfo.dependencies) {
            treeContent += `    ↳ ${dep}\n`;
          }
        }
        
        if (fileInfo.dependents.length > 0) {
          treeContent += `  依存元:\n`;
          for (const dependent of fileInfo.dependents) {
            treeContent += `    ↰ ${dependent}\n`;
          }
        }
      }
    }

    treeContent += `\n未使用ファイル:\n`;
    for (const unusedFile of this.analysisResult.details.unusedFiles) {
      treeContent += `  🗑️ ${unusedFile.path}\n`;
    }

    fs.writeFileSync(reportPath, treeContent);
    console.log(`📄 依存関係ツリー: ${reportPath}`);
    
    return { path: reportPath, content: treeContent };
  }

  /**
   * CSV形式レポート生成（表計算ソフト用）
   */
  async generateCsvReport(baseName) {
    const reportPath = path.join(this.reportDir, `${baseName}.csv`);
    
    const csvRows = [
      'Type,Name,Path,Size,Risk,Recommendation,Impact'
    ];

    // 未使用ファイルをCSVに追加
    for (const file of this.analysisResult.details.unusedFiles) {
      const risk = this.assessFileDeletionRisk(file);
      const recommendation = this.getFileRecommendation(file);
      csvRows.push(`File,"${file.path}","${file.path}",${file.size},"${risk}","${recommendation}","Medium"`);
    }

    // 未使用関数をCSVに追加
    for (const func of this.analysisResult.details.unusedFunctions) {
      const risk = this.assessFunctionDeletionRisk(func);
      const recommendation = this.getFunctionRecommendation(func);
      const impact = this.assessFunctionImpact(func);
      csvRows.push(`Function,"${func.name}","${func.definedIn.join(';')}",0,"${risk}","${recommendation}","${impact}"`);
    }

    const csvContent = csvRows.join('\n');
    fs.writeFileSync(reportPath, csvContent);
    console.log(`📄 CSVレポート: ${reportPath}`);
    
    return { path: reportPath, content: csvContent };
  }

  // リスク評価・推奨アクション生成メソッド群

  /**
   * ファイル削除リスクを評価
   */
  assessFileDeletionRisk(file) {
    // appsscript.jsonは削除リスクが非常に高い
    if (file.path === 'appsscript.json') return 'high';
    
    // HTMLファイルは通常重要
    if (file.type === 'html-frontend') return 'medium';
    
    // .gsファイルでサイズが大きい場合は慎重に
    if (file.type === 'gas-backend' && file.size > 1000) return 'medium';
    
    // 小さな設定ファイル等は低リスク
    if (file.size < 500) return 'low';
    
    return 'low';
  }

  /**
   * 関数削除リスクを評価
   */
  assessFunctionDeletionRisk(func) {
    // エントリーポイント関数は削除しない
    if (this.analyzer.entryPoints.has(func.name)) return 'high';
    
    // 複数ファイルで定義されている関数は慎重に
    if (func.definedIn.length > 1) return 'medium';
    
    // 一般的な名前の関数は慎重に
    const commonNames = ['init', 'setup', 'config', 'utils', 'helper'];
    if (commonNames.some(common => func.name.toLowerCase().includes(common))) {
      return 'medium';
    }
    
    return 'low';
  }

  /**
   * ファイル推奨アクションを生成
   */
  getFileRecommendation(file) {
    const risk = this.assessFileDeletionRisk(file);
    
    if (risk === 'high') {
      return '削除非推奨：手動で慎重に検証が必要';
    } else if (risk === 'medium') {
      return 'テスト後削除：機能テスト後に削除を検討';
    } else {
      return '安全に削除可能：バックアップ後に削除';
    }
  }

  /**
   * 関数推奨アクションを生成
   */
  getFunctionRecommendation(func) {
    const risk = this.assessFunctionDeletionRisk(func);
    
    if (risk === 'high') {
      return '削除非推奨：エントリーポイントまたは重要な関数';
    } else if (risk === 'medium') {
      return 'テスト後削除：動作確認後に削除を検討';
    } else {
      return '安全に削除可能：バックアップ後に削除';
    }
  }

  /**
   * 関数インパクトを評価
   */
  assessFunctionImpact(func) {
    if (func.definedIn.length > 1) return 'High';
    return 'Low';
  }

  /**
   * ファイル依存関係を取得
   */
  getFileDependencies(filePath) {
    const fileInfo = this.analysisResult.dependencies[filePath];
    if (!fileInfo) return [];
    
    return {
      dependencies: fileInfo.dependencies || [],
      dependents: fileInfo.dependents || []
    };
  }

  /**
   * 全体的なリスク評価を生成
   */
  generateRiskAssessment() {
    const unusedFiles = this.analysisResult.details.unusedFiles;
    const unusedFunctions = this.analysisResult.details.unusedFunctions;
    
    const fileRisks = unusedFiles.map(f => this.assessFileDeletionRisk(f));
    const functionRisks = unusedFunctions.map(f => this.assessFunctionDeletionRisk(f));
    
    return {
      files: {
        low: fileRisks.filter(r => r === 'low').length,
        medium: fileRisks.filter(r => r === 'medium').length,
        high: fileRisks.filter(r => r === 'high').length
      },
      functions: {
        low: functionRisks.filter(r => r === 'low').length,
        medium: functionRisks.filter(r => r === 'medium').length,
        high: functionRisks.filter(r => r === 'high').length
      }
    };
  }

  /**
   * クリーンアップ計画を生成
   */
  generateCleanupPlan() {
    return {
      phase1: '低リスクファイル・関数の削除',
      phase2: 'テスト実行と動作確認',
      phase3: '中リスク項目の慎重な検討',
      phase4: '高リスク項目の手動検証',
      rollbackPlan: 'バックアップからの復旧手順を準備'
    };
  }
}

// コマンドライン実行時の処理
if (require.main === module) {
  async function main() {
    try {
      const reporter = new CleanupReporter('./src');
      const reports = await reporter.generateReport();
      
      console.log('\n📊 生成されたレポート:');
      for (const [type, report] of Object.entries(reports)) {
        console.log(`  ${type}: ${report.path}`);
      }
      
    } catch (error) {
      console.error('❌ エラー:', error.message);
      process.exit(1);
    }
  }
  
  main();
}

module.exports = CleanupReporter;