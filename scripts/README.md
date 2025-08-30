# 未使用コード自動削除システム

Google Apps Script (GAS) プロジェクト用の包括的な未使用コード検出・削除システムです。安全なバックアップとロールバック機能を備えています。

## 🚀 クイックスタート

### 1. ドライラン実行（推奨）

最初に必ずドライランで動作を確認してください：

```bash
npm run cleanup:dry
```

### 2. 詳細レポート生成

削除対象を詳しく確認したい場合：

```bash
npm run cleanup:report
```

### 3. 実際のクリーンアップ実行

確認後、実際のクリーンアップを実行：

```bash
npm run cleanup
```

## 📊 利用可能なコマンド

### NPMスクリプト

```bash
# ドライラン（実際の削除は行わない）
npm run cleanup:dry

# 詳細レポートのみ生成
npm run cleanup:report

# 依存関係解析のみ実行
npm run analyze

# インタラクティブクリーンアップ
npm run cleanup
```

### 直接実行

```bash
# 基本実行
node scripts/cleanup-unused-code.js

# オプション付き実行
node scripts/cleanup-unused-code.js --dry-run --risk-level medium

# バックアップ一覧表示
node scripts/rollback.js --list

# ロールバック実行
node scripts/rollback.js "scripts/backups/backup-2025-01-15T10-30-00-000Z"
```

## ⚙️ オプション

### `--dry-run`
実際のファイル変更を行わず、削除対象を表示のみ

### `--risk-level <level>`
削除リスクレベルを指定：
- `low` (デフォルト): 低リスクの項目のみ
- `medium`: 中リスク以下の項目
- `high`: すべての項目（非推奨）

### `--non-interactive`
ユーザー確認なしで実行

### `--auto`
完全自動実行（⚠️ 慎重に使用）

## 🔍 システム構成

### 1. 依存関係解析器 (`dependency-analyzer.js`)
- GAS特有のパターンを解析
- `include()`, `google.script.run.*`, HTML内のJavaScriptイベントを検出
- 関数定義と呼び出しを追跡

### 2. 安全削除システム (`safe-delete.js`)
- 削除前の自動バックアップ作成
- ファイル・関数の段階的削除
- 削除ログと詳細レポート

### 3. レポート生成器 (`cleanup-reporter.js`)
- 複数形式のレポート（JSON, Markdown, CSV, Shell Script）
- リスク評価と推奨アクション
- 実行可能な削除コマンド生成

### 4. ロールバックシステム (`rollback.js`)
- バックアップからの完全復旧
- 削除前状態の再現
- 検証機能付きロールバック

### 5. メインオーケストレーター (`cleanup-unused-code.js`)
- 全システムの統合・制御
- インタラクティブな実行フロー
- エラーハンドリングと復旧

## 📁 生成されるファイル

### バックアップ
```
scripts/backups/
├── backup-2025-01-15T10-30-00-000Z/
│   ├── src/                    # 完全なソースバックアップ
│   ├── backup-metadata.json    # バックアップ情報
│   └── delete-report.json      # 削除内容詳細
```

### レポート
```
scripts/reports/
├── cleanup-report-2025-01-15T10-30-00-000Z-detailed.json    # 詳細JSON
├── cleanup-report-2025-01-15T10-30-00-000Z-summary.md       # サマリーMarkdown
├── cleanup-report-2025-01-15T10-30-00-000Z-actions.sh       # 実行スクリプト
├── cleanup-report-2025-01-15T10-30-00-000Z-tree.txt         # 依存関係ツリー
└── cleanup-report-2025-01-15T10-30-00-000Z.csv              # CSV形式
```

## ⚠️ 重要な注意事項

### 実行前の準備
1. **Gitコミット**: 変更をコミットしてください
2. **テストの確認**: すべてのテストが通ることを確認
3. **バックアップ確認**: 重要なデータのバックアップを取る

### 実行後の確認
1. **テスト実行**: 削除後にテストを実行
2. **動作確認**: アプリケーションの主要機能を確認
3. **デプロイテスト**: 可能であればステージング環境でテスト

### 削除対象の理解
- **エントリーポイント関数**: `doGet`, `doPost`, `onOpen`等は削除されません
- **HTML内のイベント**: `onclick`等で呼ばれる関数は保持されます
- **動的呼び出し**: `eval()`等で呼ばれる関数は検出困難

## 🔄 ロールバック手順

### 1. バックアップ確認
```bash
node scripts/rollback.js --list
```

### 2. ロールバック実行
```bash
node scripts/rollback.js "scripts/backups/backup-2025-01-15T10-30-00-000Z"
```

### 3. 動作確認
```bash
npm test
npm run lint
```

## 🐛 トラブルシューティング

### よくある問題

#### 1. 必要な関数が削除された
**解決**: ロールバックを実行し、その関数を手動で保護リストに追加

#### 2. テストが失敗する
**解決**: ロールバック実行後、削除対象を再検討

#### 3. HTMLイベントハンドラーが動作しない
**解決**: HTMLファイル内のイベント定義を確認し、パターンを調整

### デバッグ方法

#### 詳細ログ有効化
```bash
DEBUG=true node scripts/cleanup-unused-code.js --dry-run
```

#### 特定ファイルの依存関係確認
```bash
node scripts/dependency-analyzer.js | grep "特定のファイル名"
```

## 📈 カスタマイズ

### エントリーポイントの追加
`dependency-analyzer.js`の`entryPoints`配列に追加：

```javascript
this.entryPoints = new Set([
  'doGet', 'doPost', 'include', 'onOpen', 'onEdit', 'onFormSubmit',
  'installTrigger', 'uninstallTrigger', 'onInstall',
  'yourCustomEntryPoint'  // 追加
]);
```

### 検出パターンの調整
`patterns`オブジェクトを修正して新しいパターンを追加：

```javascript
this.patterns = {
  // 既存のパターン...
  customPattern: /your-regex-here/g
};
```

### リスク評価のカスタマイズ
`cleanup-reporter.js`の`assessFileDeletionRisk`メソッドを修正

## 🤝 貢献

バグ報告や機能要求はGitHubのIssueでお願いします。

## 📜 ライセンス

このプロジェクトのライセンスに準拠します。

---

**⚡ 重要**: 本番環境で使用する前に、必ずテスト環境で十分な検証を行ってください。