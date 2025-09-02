#!/usr/bin/env node

/**
 * 最終var変換プログラム - 残りの複雑なケースを慎重に処理
 * 再代入パターンを詳細分析してconst/letを適切に選択
 */

const fs = require('fs');
const path = require('path');

class FinalVarConverter {
  constructor() {
    this.backupDir = path.join(__dirname, '../backups/final-var-conversion-' + Date.now());
  }

  createBackup(filePath) {
    if (!fs.existsSync(this.backupDir)) {
      fs.mkdirSync(this.backupDir, { recursive: true });
    }
    
    const fileName = path.basename(filePath);
    const backupPath = path.join(this.backupDir, fileName);
    fs.copyFileSync(filePath, backupPath);
    console.log(`📁 バックアップ作成: ${backupPath}`);
  }

  /**
   * 高精度な変数スコープ分析
   */
  analyzeComplexPatterns(content) {
    const lines = content.split('\n');
    const analysis = [];
    
    lines.forEach((line, index) => {
      if (line.match(/\bvar\s+\w+/)) {
        const varMatch = line.match(/var\s+(\w+)/);
        if (varMatch) {
          const varName = varMatch[1];
          const analysis_result = this.isVariableReassigned(lines, index, varName);
          analysis.push({
            line: line.trim(),
            lineNumber: index + 1,
            varName,
            shouldBeConst: !analysis_result.hasReassignment,
            reassignmentInfo: analysis_result
          });
        }
      }
    });
    
    return analysis;
  }

  /**
   * より精密な再代入検出
   */
  isVariableReassigned(lines, startIndex, varName) {
    let braceDepth = 0;
    let inFunction = false;
    let reassignments = [];
    
    // 変数宣言行の解析
    const declarationLine = lines[startIndex];
    const isInLoop = declarationLine.match(/for\s*\(/);
    const isConditionalAssignment = declarationLine.includes('?') && declarationLine.includes(':');
    
    // スコープ内での再代入を検索
    for (let i = startIndex; i < Math.min(startIndex + 100, lines.length); i++) {
      const line = lines[i];
      
      // ブレース深度の追跡
      braceDepth += (line.match(/{/g) || []).length;
      braceDepth -= (line.match(/}/g) || []).length;
      
      // 関数終了判定
      if (i > startIndex && braceDepth < 0) break;
      
      // 再代入パターンの検出（より精密に）
      const reassignmentRegex = new RegExp(`\\b${varName}\\s*=(?!=|==)`, 'g');
      if (i > startIndex && line.match(reassignmentRegex) && !line.includes('var') && !line.includes('const') && !line.includes('let')) {
        // ただし、比較演算子（==, ===, !=, !==）は除外
        if (!line.match(new RegExp(`\\b${varName}\\s*(===?|!==?)`, 'g'))) {
          reassignments.push({
            lineNumber: i + 1,
            line: line.trim()
          });
        }
      }
    }
    
    return {
      hasReassignment: reassignments.length > 0,
      reassignments,
      isInLoop,
      isConditionalAssignment,
      recommendation: this.getRecommendation(varName, reassignments, isInLoop, isConditionalAssignment, declarationLine)
    };
  }

  /**
   * 変数タイプに基づく推奨タイプ決定
   */
  getRecommendation(varName, reassignments, isInLoop, isConditional, declarationLine) {
    // 再代入がない場合はconst
    if (reassignments.length === 0) {
      return 'const';
    }
    
    // 明確にループカウンタやフラグ変数の場合はlet
    if (varName.includes('Index') && reassignments.length > 0) return 'let';
    if (varName.includes('Count') && reassignments.length > 0) return 'let';
    if (['i', 'j', 'k'].includes(varName)) return 'let';
    if (varName.match(/^(is|has|should|can)/i)) return 'let';
    if (declarationLine.includes('= false') || declarationLine.includes('= true')) return 'let';
    if (declarationLine.includes('= -1') || declarationLine.includes('= 0')) return 'let';
    if (declarationLine.includes('= null')) return 'let';
    if (declarationLine.includes('= {}') || declarationLine.includes('= []')) return 'let';
    
    // その他の複雑なケースはlet
    return 'let';
  }

  /**
   * 最終的なvar変換実行
   */
  performFinalConversion(content) {
    const analysis = this.analyzeComplexPatterns(content);
    let converted = content;
    const changes = [];
    
    // 分析結果に基づいて変換
    analysis.forEach(item => {
      const { line, varName, shouldBeConst, reassignmentInfo } = item;
      const targetType = reassignmentInfo.recommendation;
      
      // より安全なパターンマッチング
      const safePattern = new RegExp(
        `^(\\s*)var(\\s+${varName}\\s*=.*?);?\\s*$`, 
        'gm'
      );
      
      converted = converted.replace(safePattern, (match, indent, declaration) => {
        changes.push(`${targetType}変換: var ${varName} → ${targetType} ${varName} (再代入: ${reassignmentInfo.reassignments.length}回)`);
        return `${indent}${targetType}${declaration};`;
      });
    });
    
    return { converted, changes, analysis };
  }

  /**
   * 手動確認レポート生成
   */
  generateReport(analysis) {
    console.log(`\n📋 変数分析レポート:`);
    console.log(`   - 総分析対象: ${analysis.length}個`);
    
    const constCandidates = analysis.filter(a => a.reassignmentInfo.recommendation === 'const');
    const letCandidates = analysis.filter(a => a.reassignmentInfo.recommendation === 'let');
    
    console.log(`   - const変換推奨: ${constCandidates.length}個`);
    console.log(`   - let変換推奨: ${letCandidates.length}個`);
    
    // 複雑なケースの詳細表示
    const complexCases = analysis.filter(a => a.reassignmentInfo.reassignments.length > 2);
    if (complexCases.length > 0) {
      console.log(`\n⚠️ 複雑な再代入パターン (${complexCases.length}個):`);
      complexCases.slice(0, 5).forEach(item => {
        console.log(`   Line ${item.lineNumber}: ${item.varName} (${item.reassignmentInfo.reassignments.length}回再代入)`);
      });
    }
  }

  processFile(filePath) {
    console.log(`\n🔧 最終変換処理: ${path.basename(filePath)}`);
    
    const content = fs.readFileSync(filePath, 'utf8');
    const originalVarCount = (content.match(/\bvar\s+\w+/g) || []).length;
    
    console.log(`📊 処理前のvar数: ${originalVarCount}`);
    
    if (originalVarCount === 0) {
      console.log(`✅ 処理対象なし`);
      return;
    }

    this.createBackup(filePath);
    
    const { converted, changes, analysis } = this.performFinalConversion(content);
    const finalVarCount = (converted.match(/\bvar\s+\w+/g) || []).length;
    
    console.log(`📊 処理後のvar数: ${finalVarCount}`);
    console.log(`📊 変換数: ${originalVarCount - finalVarCount}`);
    
    this.generateReport(analysis);
    
    if (changes.length > 0) {
      console.log(`\n✅ 適用された変更 (${changes.length}件):`);
      changes.slice(0, 15).forEach(change => console.log(`   - ${change}`));
      if (changes.length > 15) {
        console.log(`   ... 他${changes.length - 15}件`);
      }
      
      fs.writeFileSync(filePath, converted);
      console.log(`💾 ファイル更新完了`);
      
      if (finalVarCount > 0) {
        console.log(`\n⚠️ 変換されなかったvar (${finalVarCount}個):`);
        const remainingVars = converted.match(/^\s*var\s+\w+.*/gm);
        if (remainingVars) {
          remainingVars.slice(0, 5).forEach(line => {
            console.log(`   ${line.trim()}`);
          });
        }
      }
    }
  }

  run(targetFiles) {
    console.log(`🏁 最終var変換プログラム開始`);
    console.log(`🎯 残りの複雑なvarパターンを高精度分析`);
    console.log(`🛡️ バックアップ: ${this.backupDir}`);
    
    targetFiles.forEach(filePath => {
      this.processFile(filePath);
    });
    
    console.log(`\n🎉 最終変換処理完了`);
  }
}

// メイン実行
if (require.main === module) {
  const targetFiles = [
    '/Users/ryuma/Everyone-s-Answer-Board/src/database.gs'
  ];
  
  const converter = new FinalVarConverter();
  converter.run(targetFiles);
}

module.exports = FinalVarConverter;