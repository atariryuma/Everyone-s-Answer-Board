#!/usr/bin/env node

/**
 * CLAUDE.md準拠違反修正スクリプト
 * configJSON中心設計に準拠するよう非推奨パターンを自動修正
 */

const fs = require('fs');
const path = require('path');

const SRC_DIR = path.join(__dirname, '../src');

class ClaudeMdFixer {
  constructor() {
    this.fixes = [];
    this.errors = [];
  }

  async fix() {
    console.log('🔧 CLAUDE.md準拠違反の自動修正開始...\n');
    
    // .gsファイルのみを対象
    const gsFiles = this.getFiles(SRC_DIR, ['.gs']);
    
    for (const filePath of gsFiles) {
      try {
        await this.fixFile(filePath);
      } catch (error) {
        console.error(`❌ ${path.relative(SRC_DIR, filePath)}: ${error.message}`);
        this.errors.push({ file: filePath, error: error.message });
      }
    }
    
    this.generateReport();
    
    return {
      fixes: this.fixes,
      errors: this.errors
    };
  }

  async fixFile(filePath) {
    const relativePath = path.relative(SRC_DIR, filePath);
    let content = fs.readFileSync(filePath, 'utf8');
    const originalContent = content;
    let changeCount = 0;

    // 修正ルール定義
    const fixes = [
      // 非推奨な直接フィールドアクセスを修正
      {
        pattern: /userInfo\.spreadsheetId/g,
        replacement: 'config.spreadsheetId',
        description: 'userInfo.spreadsheetId → config.spreadsheetId',
        requireConfigDeclaration: true
      },
      {
        pattern: /userInfo\.sheetName/g,
        replacement: 'config.sheetName', 
        description: 'userInfo.sheetName → config.sheetName',
        requireConfigDeclaration: true
      },
      // parsedConfigの修正（最重要）
      {
        pattern: /userInfo\.parsedConfig/g,
        replacement: 'JSON.parse(userInfo.configJson || \'{}\')',
        description: 'userInfo.parsedConfig → JSON.parse(userInfo.configJson || \'{}\')',
        requireConfigDeclaration: false
      }
    ];

    // 各修正ルールを適用
    for (const fix of fixes) {
      const beforeContent = content;
      content = content.replace(fix.pattern, fix.replacement);
      
      if (content !== beforeContent) {
        const matchCount = (beforeContent.match(fix.pattern) || []).length;
        changeCount += matchCount;
        
        this.fixes.push({
          file: relativePath,
          description: fix.description,
          count: matchCount
        });
        
        console.log(`  ✅ ${relativePath}: ${fix.description} (${matchCount}箇所)`);
      }
    }

    // config変数宣言が必要な場合の処理
    if (changeCount > 0 && this.needsConfigDeclaration(content)) {
      content = this.addConfigDeclaration(content);
      console.log(`  📋 ${relativePath}: config変数宣言を追加`);
    }

    // ファイル更新
    if (content !== originalContent) {
      fs.writeFileSync(filePath, content, 'utf8');
      console.log(`  💾 ${relativePath}: ${changeCount}箇所を修正して保存`);
    }
  }

  needsConfigDeclaration(content) {
    // config.spreadsheetIdやconfig.sheetNameが使われているが、
    // const config = の宣言がない場合
    return (content.includes('config.spreadsheetId') || content.includes('config.sheetName')) &&
           !content.includes('const config =') &&
           !content.includes('let config =');
  }

  addConfigDeclaration(content) {
    // 関数の最初のuserInfoチェック後にconfig宣言を追加
    const patterns = [
      // if (!userInfo) の後に追加
      /(\s+if \(!userInfo\) \{[^}]+\}\s*)/,
      // const userInfo = の後に追加  
      /(\s+const userInfo = [^;]+;\s*)/,
      // 関数の開始付近に追加
      /(\s+console\.log\([^)]+\);\s*)/
    ];

    for (const pattern of patterns) {
      if (pattern.test(content)) {
        content = content.replace(pattern, (match) => {
          return match + '\n    // 設定データをconfigJSONから取得\n    const config = JSON.parse(userInfo.configJson || \'{}\');\n';
        });
        break;
      }
    }

    return content;
  }

  generateReport() {
    console.log('\n' + '='.repeat(60));
    console.log('🔧 CLAUDE.md準拠違反修正レポート');
    console.log('='.repeat(60));

    if (this.fixes.length === 0) {
      console.log('✅ 修正対象は見つかりませんでした');
      return;
    }

    // 修正統計
    const fixesByType = new Map();
    let totalFixes = 0;
    
    this.fixes.forEach(fix => {
      if (!fixesByType.has(fix.description)) {
        fixesByType.set(fix.description, { count: 0, files: [] });
      }
      fixesByType.get(fix.description).count += fix.count;
      fixesByType.get(fix.description).files.push(`${fix.file}(${fix.count})`);
      totalFixes += fix.count;
    });

    console.log(`\n📈 修正統計:`);
    console.log(`  総修正数: ${totalFixes}箇所`);
    console.log(`  対象ファイル数: ${new Set(this.fixes.map(f => f.file)).size}個`);

    console.log(`\n🔄 修正内容:`);
    fixesByType.forEach((info, description) => {
      console.log(`  ${description}: ${info.count}箇所`);
      console.log(`    ${info.files.slice(0, 3).join(', ')}${info.files.length > 3 ? ` ...他${info.files.length - 3}` : ''}`);
    });

    if (this.errors.length > 0) {
      console.log(`\n❌ エラー (${this.errors.length}件):`);
      this.errors.forEach((error, index) => {
        console.log(`  ${index + 1}. ${path.relative(SRC_DIR, error.file)}: ${error.error}`);
      });
    }

    console.log('\n✅ 修正完了！GASへのデプロイを推奨します。');
    console.log('='.repeat(60));
  }

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
}

// メイン実行
async function main() {
  const fixer = new ClaudeMdFixer();
  
  try {
    const result = await fixer.fix();
    
    // 結果をJSONファイルに保存
    const outputPath = path.join(__dirname, '../claude-md-fixes.json');
    fs.writeFileSync(outputPath, JSON.stringify(result, null, 2));
    console.log(`\n📄 修正結果は ${outputPath} に保存されました`);
    
    // エラーがある場合は終了コード1で終了
    process.exit(result.errors.length > 0 ? 1 : 0);
    
  } catch (error) {
    console.error('❌ 修正処理中にエラーが発生しました:', error);
    process.exit(1);
  }
}

if (require.main === module) {
  main();
}

module.exports = { ClaudeMdFixer };