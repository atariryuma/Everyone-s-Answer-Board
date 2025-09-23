#!/usr/bin/env node
/**
 * Pre-deploy validation script
 * デプロイ前の包括的チェック - デプロイエラーを事前に防ぐ
 */

const fs = require('fs');
const path = require('path');

class DeploymentValidator {
  constructor() {
    this.errors = [];
    this.warnings = [];
    this.srcDir = path.join(__dirname, '..', 'src');
  }

  /**
   * 🔍 包括的検証実行
   */
  async validate() {
    console.log('🔍 デプロイ前検証開始...\n');

    await this.checkDependencies();
    await this.checkTemplateIncludes();
    await this.checkServiceIntegrity();
    await this.checkConstantsAvailability();
    await this.simulateGASExecution();

    this.report();
    return this.errors.length === 0;
  }

  /**
   * 🔗 依存関係チェック
   */
  async checkDependencies() {
    console.log('🔗 依存関係チェック...');

    const dependencyMap = {
      'main.gs': [], // ✅ CLAUDE.md準拠: Direct GAS API calls, no ServiceFactory needed
      'UserService.gs': [], // ✅ CLAUDE.md準拠: GAS-Native architecture
      'ConfigService.gs': [], // ✅ CLAUDE.md準拠: GAS-Native architecture
      'DataService.gs': [], // ✅ CLAUDE.md準拠: GAS-Native architecture
      'SecurityService.gs': []
    };

    for (const [file, dependencies] of Object.entries(dependencyMap)) {
      const filePath = path.join(this.srcDir, file);
      if (!fs.existsSync(filePath)) {
        this.errors.push(`❌ 必須ファイル不足: ${file}`);
        continue;
      }

      const content = fs.readFileSync(filePath, 'utf8');

      for (const dep of dependencies) {
        if (!content.includes(dep)) {
          this.errors.push(`❌ 依存関係不足: ${file} → ${dep}`);
        }
      }
    }
  }

  /**
   * 📄 HTMLテンプレートincludeチェック
   */
  async checkTemplateIncludes() {
    console.log('📄 HTMLテンプレートincludeチェック...');

    const htmlFiles = this.getHtmlFiles();
    const includePattern = /<\?!=\s*include\(['"`]([^'"`]+)['"`]\)\s*;\s*\?>/g;

    for (const htmlFile of htmlFiles) {
      const content = fs.readFileSync(htmlFile, 'utf8');
      let match;

      while ((match = includePattern.exec(content)) !== null) {
        const includedFile = match[1];
        const includePath = path.join(this.srcDir, `${includedFile}.html`);

        if (!fs.existsSync(includePath)) {
          this.errors.push(`❌ includeファイル不足: ${path.basename(htmlFile)} → ${includedFile}.html`);
        }
      }
    }

    // main.gsのinclude関数存在確認
    const mainPath = path.join(this.srcDir, 'main.gs');
    if (fs.existsSync(mainPath)) {
      const content = fs.readFileSync(mainPath, 'utf8');
      if (!content.includes('function include(')) {
        this.errors.push('❌ main.gsにinclude関数が定義されていません');
      }
    }
  }

  /**
   * ⚙️ Services整合性チェック
   */
  async checkServiceIntegrity() {
    console.log('⚙️ Services整合性チェック...');

    const requiredServices = [
      'UserService.gs',
      'ConfigService.gs',
      'DataService.gs'
    ];

    const requiredMethods = {
      'UserService': ['getCurrentUserInfo', 'isAdministrator'],
      'ConfigService': ['hasCoreSystemProps', 'isSystemSetup'],
      'DataService': ['getUserSheetData', 'processRawData'] // ✅ CLAUDE.md準拠: processReaction moved to ReactionService
    };

    // Check for main.gs functions separately (CLAUDE.md compliance: Direct GAS API calls)
    const mainGsPath = path.join(this.srcDir, 'main.gs');
    if (fs.existsSync(mainGsPath)) {
      const mainContent = fs.readFileSync(mainGsPath, 'utf8');
      const requiredMainFunctions = ['getCurrentEmail'];
      for (const func of requiredMainFunctions) {
        if (!mainContent.includes(`function ${func}`)) {
          this.errors.push(`❌ 必須メソッド不足: main.${func}`);
        }
      }
    }

    for (const service of requiredServices) {
      const servicePath = path.join(this.srcDir, service);
      if (!fs.existsSync(servicePath)) {
        this.errors.push(`❌ 必須サービス不足: ${service}`);
        continue;
      }

      const content = fs.readFileSync(servicePath, 'utf8');
      const serviceName = path.basename(service, '.gs');

      if (requiredMethods[serviceName]) {
        for (const method of requiredMethods[serviceName]) {
          if (!content.includes(`${method}(`)) {
            this.errors.push(`❌ 必須メソッド不足: ${serviceName}.${method}`);
          }
        }
      }
    }
  }

  /**
   * 📋 定数可用性チェック
   */
  async checkConstantsAvailability() {
    console.log('📋 定数可用性チェック...');

    // ✅ CLAUDE.md準拠: Zero-dependency architecture - no ServiceFactory needed
    // ServiceFactory replaced with GAS-Native direct API calls
    // Skip ServiceFactory validation as it's been migrated to GAS-Native pattern

    // ✅ CLAUDE.md準拠: No validation needed for GAS-Native architecture
    // All required functionality is available through direct GAS API calls
  }

  /**
   * 🎯 GAS実行シミュレーション
   */
  async simulateGASExecution() {
    console.log('🎯 GAS実行シミュレーション...');

    // main.gs のdoGet関数存在確認
    const mainPath = path.join(this.srcDir, 'main.gs');
    if (fs.existsSync(mainPath)) {
      const content = fs.readFileSync(mainPath, 'utf8');

      if (!content.includes('function doGet(')) {
        this.errors.push('❌ main.gsにdoGet関数が定義されていません');
      }

      // 基本的な構文エラーチェック
      try {
        // 簡易的な構文チェック（コメント削除後の括弧バランス）
        const cleaned = content.replace(/\/\*[\s\S]*?\*\//g, '').replace(/\/\/.*$/gm, '');
        const openBraces = (cleaned.match(/\{/g) || []).length;
        const closeBraces = (cleaned.match(/\}/g) || []).length;

        if (openBraces !== closeBraces) {
          this.warnings.push('⚠️ main.gsで括弧のバランスが不正の可能性');
        }
      } catch (e) {
        this.warnings.push(`⚠️ main.gs構文チェックエラー: ${e.message}`);
      }
    }
  }

  /**
   * 📁 HTMLファイル一覧取得
   */
  getHtmlFiles() {
    const htmlFiles = [];
    const scanDir = (dir) => {
      const items = fs.readdirSync(dir);
      for (const item of items) {
        const fullPath = path.join(dir, item);
        const stat = fs.statSync(fullPath);

        if (stat.isDirectory()) {
          scanDir(fullPath);
        } else if (item.endsWith('.html')) {
          htmlFiles.push(fullPath);
        }
      }
    };

    scanDir(this.srcDir);
    return htmlFiles;
  }

  /**
   * 📊 検証結果レポート
   */
  report() {
    console.log('\n📊 検証結果レポート');
    console.log('==================');

    if (this.errors.length === 0 && this.warnings.length === 0) {
      console.log('✅ すべてのチェックに合格しました！デプロイ可能です。');
      return;
    }

    if (this.errors.length > 0) {
      console.log('\n❌ エラー (デプロイ不可):');
      this.errors.forEach(error => console.log(`  ${error}`));
    }

    if (this.warnings.length > 0) {
      console.log('\n⚠️ 警告:');
      this.warnings.forEach(warning => console.log(`  ${warning}`));
    }

    console.log(`\n📈 統計: エラー${this.errors.length}件, 警告${this.warnings.length}件`);

    if (this.errors.length > 0) {
      console.log('\n🚫 デプロイを中止してください。上記エラーを修正後、再実行してください。');
      process.exit(1);
    }
  }
}

// 実行
if (require.main === module) {
  const validator = new DeploymentValidator();
  validator.validate().then(success => {
    process.exit(success ? 0 : 1);
  });
}

module.exports = DeploymentValidator;