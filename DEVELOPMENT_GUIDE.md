# 🎯 Development Guide - Claude Code 2025 Optimized

> **Everyone's Answer Board 開発者ガイド**
>
> **Claude Code 2025 最適化開発環境での効率的な開発フロー**

---

## 🚀 **開発環境セットアップ**

### **必要なツール**

```bash
# 1. Claude Code CLI
npm install -g @anthropic-ai/claude-code

# 2. Google Apps Script CLI  
npm install -g @google/clasp

# 3. プロジェクト依存関係
npm install
```

### **初期設定**

```bash
# 1. Google Apps Script認証
npx clasp login

# 2. プロジェクトプル
npx clasp pull

# 3. Claude Code開始
claude
```

---

## 📋 **Claude Code 2025 開発フロー**

### **Phase 1: プロジェクト開始（毎回）**

```bash
# コンテキストクリア
/clear

# TDD監視モード開始
npm run test:watch

# Claude Codeが自動実行する内容:
# 1. CLAUDE.md読み込み（プロジェクト理解）
# 2. TodoWrite作成（タスク計画）
# 3. Git branch作成（安全性確保）
```

### **Phase 2: 戦略・実行分離開発**

#### **戦略レベル（Claude Code得意領域）**

- 要件分析・設計判断
- アーキテクチャ決定
- テスト設計
- パフォーマンス最適化計画

#### **実行レベル（Claude Code自動化）**

- テストコード生成
- 実装コード作成
- リファクタリング
- ドキュメント更新

### **Phase 3: 品質確保（ゼロトレラント）**

```bash
# 必須品質チェック
npm run check             # テスト + リント + 型チェック

# Claude Code専用コマンド
/test-architecture        # アーキテクチャ検証
/review-security         # セキュリティ監査
/deploy-safe            # デプロイ安全性確認
```

---

## 🏗️ **アーキテクチャ原則**

### **サービス層設計（SOLID準拠）**

```javascript
// ✅ 正しい実装パターン
const UserService = Object.freeze({
  getCurrentEmail() {
    // 単一責任: ユーザー情報取得のみ
    return Session.getActiveUser().getEmail();
  },
  
  diagnose() {
    // 診断機能: 全サービス必須
    return { service: 'UserService', status: '✅' };
  }
});

// ❌ 避けるべきパターン
const BadService = {
  getUserAndProcessData() {
    // 複数責任: 取得と処理が混在
  }
};
```

### **configJSON中心設計**

```javascript
// ✅ 統一データソース原則
const config = JSON.parse(userInfo.configJson);
const spreadsheetId = config.spreadsheetId;

// ❌ レガシーアプローチ（廃止済み）
const spreadsheetId = userInfo.spreadsheetId; // 直接アクセス禁止
```

### **エラーハンドリング統一**

```javascript
// ✅ 統一エラーハンドラー使用
try {
  const result = DataService.processData(params);
  return result;
} catch (error) {
  return ErrorHandler.handle(error, {
    service: 'DataService',
    method: 'processData',
    userId: params.userId
  });
}
```

---

## 🧪 **テスト駆動開発（TDD-First）**

### **必須TDDワークフロー**

```bash
# 1. テスト監視開始
npm run test:watch

# 2. テスト作成（実装前）
# tests/services/NewService.test.js

# 3. 最小実装（テスト通過）
# src/services/NewService.gs

# 4. リファクタリング（グリーンを保持）
# 5. ドキュメント更新
```

### **テストパターン**

```javascript
// ✅ 正しいテスト構造
describe('NewService', () => {
  beforeEach(() => {
    resetAllMocks();
    jest.clearAllMocks();
  });

  describe('mainFunction', () => {
    it('should handle valid input', () => {
      // Arrange
      const validInput = { userId: 'test-123' };
      NewService.mainFunction.mockReturnValue({ success: true });

      // Act
      const result = NewService.mainFunction(validInput);

      // Assert
      expect(result.success).toBe(true);
    });

    it('should handle invalid input', () => {
      // エラーケースのテスト
    });
  });

  describe('Integration Tests', () => {
    it('should complete full workflow', async () => {
      // 統合テスト
    });
  });
});
```

---

## 🔧 **開発ツール・コマンド**

### **品質管理コマンド**

```bash
# 基本開発コマンド
npm test                 # テスト実行
npm run test:watch       # TDD監視モード
npm run test:coverage    # カバレッジ確認（90%以上必須）
npm run lint            # コード品質チェック
npm run format          # コード整形
npm run check           # 統合品質チェック

# Claude Code専用コマンド
npm run architecture:test    # アーキテクチャ検証
npm run security:review     # セキュリティ監査
```

### **Claude Code統合コマンド**

```bash
# アーキテクチャ検証
/test-architecture
# 実行内容:
# - サービス層適合性チェック
# - ファイル構造検証
# - レガシーコード検出
# - Object.freeze()使用確認

# セキュリティ監査
/review-security  
# 実行内容:
# - 入力検証確認
# - XSS/SQLインジェクション対策
# - 認証・認可チェック
# - 機密情報露出検査

# 安全デプロイ
/deploy-safe
# 実行内容:
# - 全品質ゲート実行
# - バックアップ作成
# - ステージング確認
# - ロールバック準備

# サービス リファクタリング
/refactor-service UserService analyze
# 実行内容:
# - サービス複雑度分析
# - SOLID原則適合性確認
# - 改善提案生成
```

---

## 📊 **パフォーマンス最適化**

### **Google Sheets API最適化**

```javascript
// ✅ バッチ操作（推奨）
const values = data.map(item => [item.field1, item.field2]);
sheet.getRange(1, 1, values.length, 2).setValues(values);

// ❌ 個別操作（禁止）
data.forEach(item => {
  sheet.getRange(row, 1).setValue(item.field1); // 遅い！
});
```

### **キャッシュ戦略**

```javascript
// ✅ 階層化キャッシュ
const cacheManager = {
  // Level 1: 実行キャッシュ（メモリ）
  execution: new Map(),
  
  // Level 2: スクリプトキャッシュ（6時間）
  script: CacheService.getScriptCache(),
  
  // Level 3: プロパティ（永続）
  persistent: PropertiesService.getScriptProperties()
};

// 使用例
function getCachedData(key) {
  // Level 1チェック
  if (cacheManager.execution.has(key)) {
    return cacheManager.execution.get(key);
  }
  
  // Level 2チェック
  const cached = cacheManager.script.get(key);
  if (cached) {
    cacheManager.execution.set(key, cached);
    return cached;
  }
  
  // 実データ取得とキャッシュ
  const data = fetchRealData(key);
  cacheManager.script.put(key, data, 300); // 5分
  cacheManager.execution.set(key, data);
  return data;
}
```

---

## 🛡️ **セキュリティベストプラクティス**

### **入力検証（必須）**

```javascript
// ✅ 必ず実行
const validation = SecurityService.validateUserData(userData);
if (!validation.isValid) {
  throw new Error(validation.errors.join(', '));
}
const safeData = validation.sanitizedData;
```

### **認証・認可（必須）**

```javascript
// ✅ 全API操作前に実行
const userEmail = UserService.getCurrentEmail();
if (!userEmail) {
  throw new Error('認証が必要です');
}

const accessLevel = SecurityService.getAccessLevel(userEmail);
if (accessLevel !== 'owner') {
  throw new Error('権限が不足しています');
}
```

### **機密情報管理**

```javascript
// ✅ PropertiesService使用
const apiKey = PropertiesService.getScriptProperties().getProperty('API_KEY');

// ❌ ハードコード禁止
const apiKey = 'sk-1234567890abcdef'; // 絶対禁止！
```

---

## 🎯 **開発ワークフロー例**

### **新機能開発の典型的フロー**

```bash
# 1. Claude Code開始
claude
/clear

# 2. 要件分析（Claude Code）
"ユーザーがデータをエクスポートできる機能を追加したい"

# 3. TodoWrite自動作成（Claude Code）
# - 要件分析・設計
# - テスト作成
# - サービス実装
# - API統合
# - 品質確認

# 4. TDD開始
npm run test:watch

# 5. テスト作成（Claude Code）
# tests/services/ExportService.test.js

# 6. 最小実装（Claude Code）
# src/services/ExportService.gs

# 7. 統合・テスト（Claude Code）
# src/main.gs（API追加）

# 8. 品質確認
npm run check
/test-architecture
/review-security

# 9. コミット・デプロイ
git add .
git commit -m "feat: add data export functionality 🤖 Generated with Claude Code"
npm run deploy:staging
```

### **バグ修正ワークフロー**

```bash
# 1. 問題再現テスト作成
# tests/bugs/specific-bug.test.js

# 2. テスト失敗確認
npm test

# 3. 最小修正実装
# 該当サービスを修正

# 4. テスト成功確認
npm test

# 5. リグレッション確認
npm run test:coverage

# 6. 品質ゲート
npm run check
/review-security

# 7. デプロイ
npm run deploy:staging
```

---

## 📚 **開発リソース**

### **必読ドキュメント**

1. **CLAUDE.md** - AI assistant technical specifications
2. **README.md** - Project overview and setup
3. **ARCHITECTURE_ANALYSIS.md** - System architecture details

### **参考資料**

- [Google Apps Script V8 Runtime](https://developers.google.com/apps-script/guides/v8-runtime)
- [Jest Testing Framework](https://jestjs.io/docs/getting-started)
- [ESLint Rules](https://eslint.org/docs/rules/)

### **Claude Code コマンド一覧**

```bash
# アーキテクチャ・品質
/test-architecture           # システム構成検証
/review-security            # セキュリティ監査
/deploy-safe               # 安全デプロイ確認
/refactor-service          # サービスリファクタリング

# 開発支援
/clear                     # コンテキストクリア
/help                      # ヘルプ表示
```

---

## 🚨 **開発時の注意事項**

### **❌ 禁止事項**

1. **直接列アクセス**: `userInfo.spreadsheetId` → `config.spreadsheetId`を使用
2. **var使用**: `const`/`let`のみ使用
3. **レガシーマネージャー**: `UserManager`等は使用禁止
4. **テストスキップ**: 新機能は必ずテスト作成
5. **ハードコード**: 機密情報はPropertiesService使用

### **✅ 推奨事項**

1. **Object.freeze()**: 全サービスでimmutable設計
2. **JSDoc**: 関数の詳細ドキュメント
3. **診断機能**: 全サービスに`diagnose()`実装
4. **エラーハンドリング**: 統一`ErrorHandler`使用
5. **TDD**: テスト→実装→リファクタの順序厳守

---

## 🎓 **継続学習**

### **開発スキル向上**

```bash
# 週次: アーキテクチャ理解度確認
npm run architecture:test

# 月次: セキュリティ知識更新
npm run security:review

# 四半期: パフォーマンス分析
node scripts/performance-analysis.js
```

### **Claude Code習熟**

- TodoWrite駆動開発の実践
- セキュリティ First approach
- アーキテクチャ適合性の理解
- 自動化ワークフローの活用

---

*🎯 この開発ガイドは Claude Code 2025 に最適化されており、高品質・高速開発を実現するためのベストプラクティスを含んでいます。*
