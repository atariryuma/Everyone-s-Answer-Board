#!/usr/bin/env node

/**
 * 安全なvar→const/let変換プログラム
 * CLAUDE.md準拠のGAS V8コーディング規範に基づく
 */

const fs = require('fs');
const path = require('path');

class VarToConstConverter {
  constructor() {
    this.safePatterns = [
      // 定数的な使用パターン（const変換対象）
      /^(\s*)var\s+(\w+)\s*=\s*([^;]+);?\s*$/gm,
      // 関数呼び出し結果の代入（const変換対象）
      /^(\s*)var\s+(\w+)\s*=\s*(\w+\([^)]*\))\s*;?\s*$/gm,
      // PropertiesService等のサービス呼び出し（const変換対象）
      /^(\s*)var\s+(\w+)\s*=\s*(PropertiesService\.[^;]+)\s*;?\s*$/gm,
    ];
    
    this.loopPatterns = [
      // forループ内のイテレータ変数（let変換対象）
      /(\s*for\s*\(\s*)var\s+(\w+)\s+(in|of|=)/g,
      /(\s*for\s*\(\s*)var\s+(\w+\s*=\s*[^;]+;[^;]+;[^)]+\))/g,
    ];
    
    this.excludePatterns = [
      // 関数宣言内で再代入される変数（手動確認が必要）
      /var\s+\w+[^=]*=.*[\r\n].*\1\s*=/, 
    ];
    
    this.backupDir = path.join(__dirname, '../backups/var-conversion-' + Date.now());
  }

  /**
   * バックアップを作成
   */
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
   * ファイル内容を分析してvar使用パターンを分類
   */
  analyzeVarUsage(content, filePath) {
    const lines = content.split('\n');
    const analysis = {
      totalVars: 0,
      constCandidates: [],
      letCandidates: [], 
      manualReview: [],
      excludedLines: []
    };

    lines.forEach((line, index) => {
      if (line.match(/var\s+\w+/)) {
        analysis.totalVars++;
        const lineNum = index + 1;
        const context = {
          line: line.trim(),
          lineNumber: lineNum,
          filePath: path.basename(filePath)
        };

        // ループ変数の検出
        if (line.match(/for\s*\(.*var\s+\w+/)) {
          analysis.letCandidates.push({
            ...context,
            reason: 'ループイテレータ変数'
          });
        }
        // 定数的な使用パターンの検出
        else if (line.match(/var\s+\w+\s*=\s*[^=]+$/) && 
                 !this.hasReassignment(lines, index, line.match(/var\s+(\w+)/)[1])) {
          analysis.constCandidates.push({
            ...context,
            reason: '再代入なしの変数'
          });
        }
        // その他は手動レビューが必要
        else {
          analysis.manualReview.push({
            ...context,
            reason: '複雑なパターン - 手動確認要'
          });
        }
      }
    });

    return analysis;
  }

  /**
   * 変数の再代入をチェック
   */
  hasReassignment(lines, startIndex, varName) {
    // 同じ関数スコープ内での再代入をチェック
    for (let i = startIndex + 1; i < lines.length && i < startIndex + 50; i++) {
      const line = lines[i];
      
      // 関数終了の検出
      if (line.match(/^\s*}\s*$/)) break;
      
      // 再代入パターンの検出
      if (line.match(new RegExp(`\\s*${varName}\\s*=`)) && 
          !line.match(/var\s+/)) {
        return true;
      }
    }
    return false;
  }

  /**
   * 安全な変換の実行
   */
  convertSafePatterns(content) {
    let converted = content;
    let changes = [];

    // ループ変数の変換（let）
    converted = converted.replace(/(\s*for\s*\(\s*)var(\s+\w+)/g, (match, prefix, varDecl) => {
      changes.push(`ループ変数: var${varDecl} → let${varDecl}`);
      return `${prefix}let${varDecl}`;
    });

    // 定数的な変数の変換（const） - より慎重なパターン
    const constPatterns = [
      // PropertiesServiceの呼び出し
      {
        pattern: /^(\s*)var(\s+\w+\s*=\s*PropertiesService\.[^;]+);?\s*$/gm,
        replacement: '$1const$2;'
      },
      // JSON.parseの呼び出し
      {
        pattern: /^(\s*)var(\s+\w+\s*=\s*JSON\.parse\([^;]+)\);?\s*$/gm, 
        replacement: '$1const$2;'
      },
      // 関数呼び出し結果（単純なもの）
      {
        pattern: /^(\s*)var(\s+\w+\s*=\s*\w+\([^)]*\))\s*;?\s*$/gm,
        replacement: '$1const$2;'
      }
    ];

    constPatterns.forEach(({ pattern, replacement }) => {
      converted = converted.replace(pattern, (match, ...groups) => {
        const varName = match.match(/var\s+(\w+)/)?.[1];
        changes.push(`定数化: var ${varName} → const ${varName}`);
        return match.replace(/var/, 'const');
      });
    });

    return { converted, changes };
  }

  /**
   * ファイルの処理
   */
  processFile(filePath) {
    console.log(`\n🔍 処理開始: ${path.basename(filePath)}`);
    
    if (!fs.existsSync(filePath)) {
      console.log(`❌ ファイルが見つかりません: ${filePath}`);
      return;
    }

    const content = fs.readFileSync(filePath, 'utf8');
    const analysis = this.analyzeVarUsage(content, filePath);
    
    console.log(`📊 分析結果:`);
    console.log(`   - 総var数: ${analysis.totalVars}`);
    console.log(`   - const変換候補: ${analysis.constCandidates.length}`); 
    console.log(`   - let変換候補: ${analysis.letCandidates.length}`);
    console.log(`   - 手動確認必要: ${analysis.manualReview.length}`);

    if (analysis.totalVars === 0) {
      console.log(`✅ varの使用なし`);
      return;
    }

    // 詳細レポート
    if (analysis.manualReview.length > 0) {
      console.log(`\n⚠️  手動確認が必要な箇所:`);
      analysis.manualReview.slice(0, 5).forEach(item => {
        console.log(`   Line ${item.lineNumber}: ${item.line}`);
      });
      if (analysis.manualReview.length > 5) {
        console.log(`   ... 他${analysis.manualReview.length - 5}件`);
      }
    }

    // バックアップ作成
    this.createBackup(filePath);

    // 安全な変換の実行
    const { converted, changes } = this.convertSafePatterns(content);
    
    if (changes.length > 0) {
      console.log(`\n✅ 適用された変更:`);
      changes.forEach(change => console.log(`   - ${change}`));
      
      // ファイルに書き込み
      fs.writeFileSync(filePath, converted);
      console.log(`💾 ファイル更新完了`);
    } else {
      console.log(`ℹ️  自動変換可能なパターンなし`);
    }
  }

  /**
   * メイン処理
   */
  run(targetFiles) {
    console.log(`🚀 var → const/let 変換プログラム開始`);
    console.log(`📝 CLAUDE.md準拠のGAS V8規範に基づく安全な変換`);
    console.log(`🛡️  バックアップディレクトリ: ${this.backupDir}`);
    
    targetFiles.forEach(filePath => {
      this.processFile(filePath);
    });
    
    console.log(`\n✅ 変換処理完了`);
    console.log(`📋 バックアップは ${this.backupDir} に保存されています`);
  }
}

// メイン実行
if (require.main === module) {
  const targetFiles = [
    '/Users/ryuma/Everyone-s-Answer-Board/src/database.gs',
    '/Users/ryuma/Everyone-s-Answer-Board/src/Core.gs'
  ];
  
  const converter = new VarToConstConverter();
  converter.run(targetFiles);
}

module.exports = VarToConstConverter;