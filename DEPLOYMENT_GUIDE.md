# 🚀 Deployment Guide - Everyone's Answer Board

> **Claude Code 2025 最適化デプロイメントガイド**
> 
> **安全かつ効率的なプロダクション展開手順**

---

## 📋 デプロイメント前チェックリスト

### ✅ **必須確認事項**

```bash
# 1. 品質ゲート確認
npm run check                    # テスト + リント + 型チェック
npm run architecture:test        # アーキテクチャ検証
npm run security:review          # セキュリティ監査

# 2. Claude Code統合確認
/test-architecture              # アーキテクチャ適合性
/review-security               # セキュリティ適合性
/deploy-safe                   # デプロイ安全性確認
```

### 📊 **品質基準（すべて満たす必要あり）**

| 項目 | 基準 | コマンド |
|------|------|----------|
| テストカバレッジ | 90%+ | `npm run test:coverage` |
| ESLintエラー | 0件 | `npm run lint` |
| セキュリティスコア | 70+ | `npm run security:review` |
| アーキテクチャ適合性 | 100% | `npm run architecture:test` |

---

## 🏗️ **ステージング環境デプロイ**

### **Step 1: 環境準備**

```bash
# Google Apps Script CLI設定
npx clasp login
npx clasp pull

# 依存関係確認
npm install
npm audit fix
```

### **Step 2: ステージング展開**

```bash
# ステージング用デプロイ（自動品質チェック付き）
npm run deploy:staging

# 実行内容:
# 1. npm run check（品質ゲート）
# 2. クラスププッシュ
# 3. ステージング環境での動作確認
```

### **Step 3: ステージング検証**

```bash
# ステージング環境での統合テスト
curl -X GET "https://script.google.com/macros/s/STAGING_SCRIPT_ID/exec?mode=debug"

# 期待レスポンス:
# {
#   "app": { "name": "Everyone's Answer Board", "version": "2.0.0" },
#   "services": { "user": "✅", "config": "✅", "data": "✅", "security": "✅" }
# }
```

---

## 🎯 **プロダクション環境デプロイ**

### **Step 1: 最終品質確認**

```bash
# 包括的品質監査
npm run test:coverage           # カバレッジ90%以上確認
npm run architecture:test       # アーキテクチャ100%適合確認
npm run security:review         # セキュリティ70点以上確認
npm run lint                    # コード品質0エラー確認
```

### **Step 2: プロダクションデプロイ**

```bash
# プロダクション用デプロイ（厳格チェック付き）
npm run deploy:prod

# 実行内容:
# 1. 全品質ゲート実行
# 2. セキュリティ監査
# 3. バックアップ作成
# 4. プロダクション展開
# 5. ヘルスチェック
```

### **Step 3: デプロイ後確認**

```bash
# 1. アプリケーション起動確認
curl -X GET "https://script.google.com/macros/s/PROD_SCRIPT_ID/exec?mode=debug"

# 2. 認証フロー確認
curl -X GET "https://script.google.com/macros/s/PROD_SCRIPT_ID/exec?mode=login"

# 3. API機能確認
curl -X POST "https://script.google.com/macros/s/PROD_SCRIPT_ID/exec" \
  -H "Content-Type: application/json" \
  -d '{"action": "add_reaction", "userId": "test", "rowId": "row_2", "reactionType": "LIKE"}'
```

---

## 🔧 **環境設定**

### **Google Apps Script設定**

#### **1. スクリプトプロパティ設定**

```javascript
// PropertiesService.getScriptProperties()
{
  "SYSTEM_CONFIG": JSON.stringify({
    "initialized": true,
    "version": "2.0.0",
    "environment": "production"
  }),
  "DATABASE_SPREADSHEET_ID": "your-database-sheet-id",
  "ADMIN_EMAIL": "admin@your-domain.com"
}
```

#### **2. ウェブアプリケーション設定**

- **実行ユーザー**: `自分（デベロッパー）`
- **アクセス権限**: `すべてのユーザー`
- **バージョン**: `新バージョン（每回デプロイ時）`

#### **3. OAuth スコープ設定**

```json
{
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/forms",
    "https://www.googleapis.com/auth/script.external_request"
  ]
}
```

---

## 🛡️ **セキュリティ設定**

### **必須セキュリティ措置**

```bash
# 1. 機密情報の環境変数化
# PropertiesServiceを使用（ハードコード禁止）

# 2. アクセス制御の有効化
# Email-based認証の確認

# 3. 入力検証の確認
# XSS/SQLインジェクション対策の動作確認
```

### **監査ログ設定**

```javascript
// システム操作の監査ログ
const auditLog = {
  timestamp: new Date().toISOString(),
  operation: 'user_action',
  userId: 'hashed_user_id',
  success: true,
  ip: 'masked_ip_address'
};
```

---

## 📊 **パフォーマンス最適化**

### **キャッシュ戦略**

```javascript
// 階層化キャッシュの設定
const cacheConfig = {
  // Level 1: 実行キャッシュ（インメモリ）
  execution: new Map(),
  
  // Level 2: スクリプトキャッシュ（6時間）
  script: CacheService.getScriptCache(),
  
  // Level 3: プロパティ（永続）
  persistent: PropertiesService.getScriptProperties()
};
```

### **データベース最適化**

```javascript
// configJSON中心型5フィールド設計
const optimizedDB = {
  HEADERS: ['userId', 'userEmail', 'isActive', 'configJson', 'lastModified'],
  BATCH_SIZE: 1000,           // バッチ処理サイズ
  CACHE_TTL: 300             // キャッシュ有効期限（秒）
};
```

---

## 🔍 **監視・メンテナンス**

### **ヘルスチェック**

```bash
# 定期ヘルスチェックURL
https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec?mode=debug

# 期待レスポンス構造:
{
  "app": { "name": "Everyone's Answer Board", "version": "2.0.0" },
  "services": {
    "user": "✅",      # UserService診断
    "config": "✅",    # ConfigService診断  
    "data": "✅",      # DataService診断
    "security": "✅"   # SecurityService診断
  },
  "overall": "✅"
}
```

### **パフォーマンス監視**

```javascript
// レスポンス時間監視
const performanceThresholds = {
  dataRetrieval: 3000,      // データ取得: 3秒以下
  reactionProcessing: 1000, // リアクション処理: 1秒以下
  configUpdate: 2000        // 設定更新: 2秒以下
};
```

### **エラー監視**

```javascript
// エラーログ監視
const errorPatterns = [
  'SpreadsheetApp.*permission',  // 権限エラー
  'Database.*connection',        // DB接続エラー
  'Authentication.*failed',      // 認証エラー
  'SecurityService.*violation'   // セキュリティ違反
];
```

---

## 🚨 **トラブルシューティング**

### **一般的な問題と対処法**

#### **1. 認証エラー**
```bash
# 問題: "Authentication required"
# 対処: OAuth再認証
npx clasp login
```

#### **2. 権限エラー**
```bash
# 問題: "Spreadsheet permission denied"
# 対処: スプレッドシート共有設定確認
```

#### **3. パフォーマンス問題**
```bash
# 問題: レスポンス時間遅延
# 対処: キャッシュクリア + 最適化確認
npm run architecture:test
```

#### **4. セキュリティ警告**
```bash
# 問題: セキュリティスコア低下
# 対処: セキュリティ監査実行
npm run security:review
```

---

## 📚 **ロールバック手順**

### **緊急ロールバック**

```bash
# 1. 前バージョンに即座に切り戻し
npx clasp deployments        # バージョン一覧確認
npx clasp deploy --deploymentId PREVIOUS_DEPLOYMENT_ID

# 2. 設定データの復旧（必要に応じて）
# PropertiesServiceから前バージョン設定を復元

# 3. ヘルスチェック実行
curl -X GET "https://script.google.com/macros/s/SCRIPT_ID/exec?mode=debug"
```

### **データ整合性確認**

```bash
# ロールバック後のデータ整合性確認
npm run test:integration
npm run architecture:test
```

---

## 🎓 **継続的改善**

### **定期メンテナンス**

```bash
# 週次: 依存関係更新
npm update
npm audit fix

# 月次: 包括的品質監査
npm run test:coverage
npm run security:review
npm run architecture:test
```

### **パフォーマンス分析**

```bash
# パフォーマンス分析レポート
node scripts/performance-analysis.js

# セキュリティ監査レポート  
node scripts/security-audit.js
```

---

*🎯 このガイドは Claude Code 2025 最適化開発環境に最適化されており、継続的な品質向上とセキュリティ確保を実現します。*