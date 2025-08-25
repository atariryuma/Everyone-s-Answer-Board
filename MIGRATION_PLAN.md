# 新アーキテクチャへのマイグレーション計画

## 概要

このドキュメントは、既存の複雑なコードベース（約58,000行）から新しいシンプルなアーキテクチャ（約2,500行）への段階的な移行計画を説明します。

## 新アーキテクチャの構成

### コアサービス（`/src/services/`）
- **ErrorService.gs** - 統一エラーハンドリング（138行）
- **CacheService.gs** - キャッシュ管理（201行）
- **GASWrapper.gs** - Google Apps Script APIラッパー（330行）
- **DatabaseService.gs** - データアクセス層（355行）
- **SpreadsheetService.gs** - スプレッドシート操作（258行）
- **APIService.gs** - APIハンドラー（407行）

### エントリーポイント
- **main-new.gs** - シンプルなメインエントリーポイント（242行）

### マイグレーション支援
- **migration/bridge.gs** - 既存コードとの互換性ブリッジ（230行）

## マイグレーション戦略

### フェーズ1: 並行稼働（現在）
1. ✅ 新サービスの実装完了
2. ✅ ブリッジファイルによる互換性確保
3. ✅ 統合テストの作成

### フェーズ2: 段階的切り替え（次のステップ）

#### ステップ1: 新サービスの有効化
```javascript
// appsscript.json のファイル順序を更新
{
  "files": [
    // 新サービスを最初に読み込み
    {"source": "src/services/ErrorService.gs", "type": "SERVER_JS"},
    {"source": "src/services/CacheService.gs", "type": "SERVER_JS"},
    {"source": "src/services/GASWrapper.gs", "type": "SERVER_JS"},
    {"source": "src/services/DatabaseService.gs", "type": "SERVER_JS"},
    {"source": "src/services/SpreadsheetService.gs", "type": "SERVER_JS"},
    {"source": "src/services/APIService.gs", "type": "SERVER_JS"},
    {"source": "src/migration/bridge.gs", "type": "SERVER_JS"},
    // 既存ファイルは後から読み込み（必要なもののみ）
    {"source": "src/main.gs", "type": "SERVER_JS"}
  ]
}
```

#### ステップ2: エントリーポイントの切り替え
```javascript
// main.gs の doGet/doPost を新しい実装にリダイレクト
function doGet(e) {
  // 新しいmain-new.gsのdoGetを呼び出す
  return doGetNew(e);
}
```

#### ステップ3: 既存機能の段階的無効化
- 不要な unified 系マネージャーを削除
- 重複したキャッシュシステムを削除
- 冗長なエラーハンドリングを削除

### フェーズ3: 完全移行

#### 削除対象ファイル
- `unifiedCacheManager.gs` → `CacheService.gs` に置き換え
- `unifiedSecurityManager.gs` → 必要な部分のみ `GASWrapper.gs` に統合
- `unifiedBatchProcessor.gs` → 削除（過度に複雑）
- `unifiedUtilities.gs` → 必要な部分のみ各サービスに統合
- `resilientExecutor.gs` → `bridge.gs` の簡易実装で代替
- `systemIntegrationManager.gs` → 削除（不要な複雑性）

#### 統合対象ファイル
- `Core.gs` (6,783行) → 機能を各サービスに分散
  - データ取得 → `SpreadsheetService.gs`
  - API処理 → `APIService.gs`
  - 設定管理 → `DatabaseService.gs`
- `config.gs` (3,836行) → 必要な部分のみ各サービスに統合
- `database.gs` (3,574行) → `DatabaseService.gs` に簡素化して統合

## 実装ガイドライン

### 1. 新規機能の追加
新規機能は**必ず新サービス**に追加してください：
```javascript
// ❌ 悪い例：既存の巨大ファイルに追加
// Core.gs に新機能を追加

// ✅ 良い例：適切なサービスに追加
// src/services/SpreadsheetService.gs に追加
function newSpreadsheetFeature() {
  // シンプルで単一責任の実装
}
```

### 2. エラーハンドリング
```javascript
// 統一されたエラーハンドリングを使用
try {
  // 処理
} catch (error) {
  logError(error, 'functionName', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.DATABASE);
}
```

### 3. キャッシュ管理
```javascript
// シンプルなキャッシュAPIを使用
const data = getCacheValue('key') || fetchAndCache();

function fetchAndCache() {
  const value = fetchData();
  setCacheValue('key', value, 300); // 5分TTL
  return value;
}
```

### 4. GAS API呼び出し
```javascript
// 必ずラッパーを使用
const spreadsheet = Spreadsheet.openById(id);  // ✅
// const spreadsheet = SpreadsheetApp.openById(id);  // ❌
```

## テスト実行

```bash
# 統合テストの実行
npm test tests/integration/newArchitectureIntegration.test.js

# 既存テストとの互換性確認
npm test
```

## チェックリスト

### 移行前の確認
- [ ] 本番環境のバックアップ作成
- [ ] ステージング環境でのテスト完了
- [ ] パフォーマンステストの実施
- [ ] エラーログの監視体制確立

### 移行中の確認
- [ ] 新サービスの動作確認
- [ ] ブリッジ関数の動作確認
- [ ] キャッシュの正常動作確認
- [ ] APIレスポンスの互換性確認

### 移行後の確認
- [ ] 不要ファイルの削除
- [ ] ドキュメントの更新
- [ ] チーム向けトレーニングの実施
- [ ] パフォーマンス改善の測定

## メリット

### コード量の削減
- **Before**: 約58,000行
- **After**: 約2,500行
- **削減率**: 95.7%

### 保守性の向上
- ファイル数: 18個 → 8個
- 平均ファイルサイズ: 3,200行 → 250行
- 責任の明確化: 各サービスが単一責任

### パフォーマンスの改善
- API呼び出し数の削減
- キャッシュ戦略の統一
- 不要な処理の削除

## リスクと対策

### リスク1: 既存機能の動作不良
**対策**: ブリッジファイルによる完全な後方互換性の確保

### リスク2: パフォーマンス低下
**対策**: 段階的な移行とモニタリング

### リスク3: チームの学習コスト
**対策**: シンプルな設計により学習曲線を緩和

## サポート

問題が発生した場合は、以下を確認してください：

1. `MIGRATION_PLAN.md`（このドキュメント）
2. `tests/integration/newArchitectureIntegration.test.js`
3. `src/migration/bridge.gs` のマッピング

## 次のアクション

1. **開発環境でのテスト** - 新アーキテクチャの動作確認
2. **段階的デプロイ** - 一部のユーザーで新システムをテスト
3. **完全移行** - 全ユーザーを新システムに移行
4. **クリーンアップ** - 不要なコードの削除