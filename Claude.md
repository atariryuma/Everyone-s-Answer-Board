# GAS V8/ES6 コーディング方針（AI向け）

> 目的：Google Apps Script（GAS）の **V8 ランタイム** 前提で、AIが安定したコードを生成できるようにする実務ガイド。  
> 対象：コード生成AI（Claude/ChatGPT等）、レビュー用 LLM、社内自動化ボット。

---

# ⚠️ システム破壊防止ルール（これを破ると全システム停止）

## 🚨 絶対遵守：データベーススキーマ
- ✅ **唯一使用**: `database.gs` の `DB_CONFIG`
- ✅ **正式フィールド**: `tenantId`, `ownerEmail`, `createdAt`, `lastAccessedAt`, `status`
- ❌ **使用禁止**: `userId`, `adminEmail`（旧フィールド名）
- ❌ **削除済み**: `constants.gs` の `DB_SHEET_CONFIG`（古い定義）

## 🎯 必須定数（src/constants.gs）
### リアクション定数
```javascript
const REACTION_KEYS = ['UNDERSTAND', 'LIKE', 'CURIOUS'];
const COLUMN_HEADERS = {
  OPINION: '回答', REASON: '理由', UNDERSTAND: 'なるほど！', 
  LIKE: 'いいね！', CURIOUS: 'もっと知りたい！', HIGHLIGHT: 'ハイライト'
};
```

### 統一システム定数
```javascript
const SYSTEM_CONSTANTS = {
  DATABASE: { SHEET_NAME: 'Users', HEADERS: ['tenantId', 'ownerEmail', ...] },
  REACTIONS: { KEYS: REACTION_KEYS, COLUMNS: {...} },
  ACCESS_LEVELS: { OWNER: 'owner', SYSTEM_ADMIN: 'system_admin', ... }
};
```

## 🔄 システムフロー
### 管理パネル作成フロー
```
管理パネル → connectDataSource → ConfigurationManager → Database.configJson
```

### 閲覧フロー  
```
doGet → verifyAccess → getUserConfig → renderAnswerBoard → HTMLテンプレート
```

### リアクションフロー
```
addReaction → LockService → processReaction → バッチ更新 → リアルタイム反映
```

## 🏗️ アーキテクチャ階層
1. **PropertiesService**: システム設定（SERVICE_ACCOUNT_CREDS, DATABASE_SPREADSHEET_ID, ADMIN_EMAIL）
2. **Database**: テナント管理（tenantId, ownerEmail, configJson）
3. **ConfigurationManager**: 設定管理（PropertiesService + Cache）
4. **AccessController**: アクセス制御（owner > system_admin > authenticated_user > guest）
5. **ReactionManager**: リアルタイム処理（LockService + バッチ処理）

---

## 0) ランタイム前提

- GAS は **V8**（Chrome/Node と同系エンジン）で動作し、**モダンな ECMAScript 構文**を利用可能。  
- `let/const`、アロー関数、テンプレートリテラル、分割代入、クラス、`Map/Set` などが使える。  
- 旧 Rhino も選べるが、**V8 を強く推奨**。

---

## 1) まず守るコーディング規範

- **`var` を禁止**。**`const` 優先**、必要時のみ `let`。  
- **関数は原則アロー関数**（`this` を必要とするクラスメソッドは通常のメソッド記法）。  
- **テンプレートリテラル**で文字列結合を可読化。  
- **分割代入 / スプレッド**で引数・配列・オブジェクト操作を簡潔に。  
- **不変データ**志向：オブジェクトの直接破壊より新オブジェクトを返す。

---

## 2) GAS ならではの設計ルール

- **エントリーポイントはグローバル関数**（トリガーやメニュー登録はトップレベル）。  
- **Apps Script のサービス API は同期的**。`UrlFetchApp` などはブロッキング呼び出し。  
- **バッチ処理**：Spreadsheet/Drive などはまとめて取得・更新。  
- **状態は PropertiesService/CacheService** に格納。  
- **ログ**は `console.log` または `Logger.log`。  
- **例外設計**：`throw new Error()` を用い、`try/catch` でハンドリング。

---

## 3) ファイル構成とモジュール化

- GAS は **ES Modules 非対応**。`import/export` はそのままでは使えない。  
- 多ファイルに分け、**グローバル名前空間を汚さない命名**で整理。  
- npm を利用する場合はローカルでバンドル（Webpack など）、`clasp` でアップロード。

```

/src               # ES6/TS 源コード
/dist/code.js      # バンドル出力（GAS 用）
/appsscript.json   # マニフェスト

````

---

## 4) 主要 ES6+ 機能の使い分け

- **クラス**：責務分離に利用。  
- **`Map/Set`**：高速検索や非文字列キー管理に有効。  
- **イテレータ/ジェネレータ**：大規模データの段階処理に。  
- **Promise/`async`**：構文は使えるが GAS API は同期。基本は不要。

---

## 5) I/O とパフォーマンス

- **Spreadsheet**：`getValues()`→配列処理→`setValues()` 一括。  
- **Drive/UrlFetch**：必要最小限。レスポンスは即 `JSON.parse`。  
- **キャッシュ**：`CacheService`/`PropertiesService` を積極利用。  
- **Utilities.sleep** は最小限。

---

## 6) 例外・リトライ・検証

- **入力検証**：Public 関数は JSDoc で型・必須性を明記。  
- **外部呼び出し**は指数バックオフを実装。  
- **失敗時の痕跡**：`console.error`＋簡潔な要約を残す。

---

## 7) ローカル開発・デプロイ

- **`clasp`** でローカル開発／デプロイ。  
- **ライブラリ**は Script ID 指定でバージョン固定。  
- **npm** は「バンドル → 単一ファイル」の原則。

---

## 8) HTML Service/フロント連携

- クライアント JS はブラウザの ES Modules 利用可。  
- サーバ側とは別。`google.script.run` で呼び出す。

---

## 9) 生成AI向けプロンプト指示（サンプル）

1. **`const`優先、`let`のみ許可、`var`禁止**。  
2. **エントリーポイントはトップレベル関数**。  
3. **バッチI/O・最小呼び出し**を強制。  
4. **JSDoc**を必須。  
5. **例外処理・ログ方針**を明記。  
6. **`async/await`は不要**。  
7. npm が必要な場合のみバンドル指示。  
8. 依存がなければ **単一ファイルに収束**。

---

## 10) 最小テンプレート（雛形）

```js
/** @OnlyCurrentDoc */

/**
 * シートのA列を集計してログ出力する
 * @return {void}
 */
function main() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const values = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
  const nonEmpty = values.flat().filter(v => String(v).trim() !== "");
  const counts = countBy(nonEmpty);
  console.log(summaryToLines(counts).join("\n"));
}

/** @param {string[]} arr @return {Map<string, number>} */
const countBy = (arr) => {
  const m = new Map();
  for (const v of arr) m.set(v, (m.get(v) ?? 0) + 1);
  return m;
};

/** @param {Map<string, number>} m @return {string[]} */
const summaryToLines = (m) =>
  Array.from(m.entries()).map(([k, n]) => `${k}: ${n}`);
````

---

## AI にありがちな誤りと対策

* **誤**「`import/export` を .gs に書く」
  → **正**「バンドルして単一 JS にしてからデプロイ」。
* **誤**「`await UrlFetchApp.fetch()`」
  → **正**「同期API。`await` は意味がない」。
* **誤**「セルを1件ずつ `setValue`」
  → **正**「`getValues`/`setValues` で一括」。
* **誤**「`var` 多用」
  → **正**「`const`/`let`」。

---

# Claude Code AI開発ワークフロー

> このセクションはClaude Code（AI）を使った効率的なGAS開発プロセスを定義します。

---

## 11) Claude Code を使った反復開発

Claude CodeはAI支援によるコーディング試行錯誤を大幅に効率化できます。

### TDD（テスト駆動開発）アプローチ
1. **テスト先行開発**：新機能実装時は先にテストケースを作成
   ```bash
   npm run test:watch  # テスト監視モード
   ```
2. **AI提案→テスト実行→改善サイクル**：
   - Claudeに実装を提案させる
   - `npm run test` で即座に検証
   - 失敗箇所をClaudeに修正依頼
   - 継続的に品質向上

### コード選択リファクタリング
- VS Code上でコード選択 → Claudeに「この部分をリファクタリングして」
- 局所的な改善で安全なリファクタリング実現
- GAS特有の制約も考慮したリファクタリング提案

### プロジェクトガイドライン参照
- ClaudeはCLAUDE.md、README.mdを自動参照
- プロジェクト固有のルールや意図を自動反映
- チーム開発での一貫性保持

---

## 12) コードフォーマット・Lint自動化

開発効率とコード品質を両立する自動化環境を構築済み。

### 利用可能コマンド
```bash
# コード整形（Prettier）
npm run format

# 構文チェック・自動修正（ESLint）  
npm run lint

# テスト実行
npm run test

# 品質チェック一括実行
npm run check

# デプロイ前総合チェック
npm run deploy
```

### VS Code連携設定推奨
```json
// .vscode/settings.json
{
  "editor.formatOnSave": true,
  "editor.codeActionsOnSave": {
    "source.fixAll.eslint": true
  },
  "eslint.validate": ["javascript", "gas"]
}
```

### AI生成コードの品質保証
- AIが生成したコードも自動的にプロジェクト標準に整形
- ESLintによるGAS特有のベストプラクティスチェック
- 保存時自動修正でヒューマンエラー防止

---

## 13) GAS最適化ベストプラクティス

GAS環境の制約と特性を活かした最適化指針。

### 実行時間・処理分割戦略
- **長時間処理は分割必須**：6分制限対策
- **トリガー・スケジューラ活用**：段階的な大量処理
- **プロセス状態管理**：PropertiesServiceで処理継続

### API呼び出し最適化
```js
// ❌ 非効率：ループ内でAPI呼び出し  
for(const row of rows) {
  sheet.getRange(row, 1).setValue(data);
}

// ✅ 効率的：バッチ処理
const values = rows.map(row => [data]);
sheet.getRange(1, 1, values.length, 1).setValues(values);
```

### キャッシュ戦略
- **CacheService**：短期間（最大6時間）の高速アクセス
- **PropertiesService**：永続化が必要なデータ
- **繰り返し処理のキャッシュ化**：API呼び出し削減

### エラーハンドリング・ログ戦略
```js
const handleWithRetry = (operation, maxRetries = 3) => {
  for(let i = 0; i < maxRetries; i++) {
    try {
      return operation();
    } catch (error) {
      console.error(`Attempt ${i + 1} failed:`, error.message);
      if (i === maxRetries - 1) throw error;
      Utilities.sleep(Math.pow(2, i) * 1000); // 指数バックオフ
    }
  }
};
```

### V8ランタイム活用
- **Map/Set**：高性能データ構造
- **Array.prototype.flat()**：配列展開
- **テンプレートリテラル**：文字列組み立て
- **アロー関数**：簡潔な関数記法

---

## 14) GAS API モックテスト

本格的な単体テストでバグ予防・品質向上を実現。

### テスト実行環境
```bash
# 単体テスト実行
npm run test

# ウォッチモード（開発時）
npm run test:watch

# カバレッジ付きテスト
npm run test -- --coverage
```

### GAS API モック利用例
```js
// tests/example.test.js
describe('スプレッドシート処理', () => {
  beforeEach(() => {
    // モックの初期化
    SpreadsheetApp.getActiveSheet.mockReturnValue({
      getRange: jest.fn(() => ({
        getValues: jest.fn(() => [['data1'], ['data2']]),
        setValues: jest.fn()
      })),
      getLastRow: jest.fn(() => 10)
    });
  });

  test('データ集計処理', () => {
    const result = countData(); // あなたのGAS関数
    expect(result).toEqual(expectedResult);
    // SpreadsheetApp呼び出し検証
    expect(SpreadsheetApp.getActiveSheet).toHaveBeenCalled();
  });
});
```

### Claude連携テスト開発
1. **Claudeにテストケース作成依頼**
   ```
   "この関数のテストケースを作成してください"
   ```
2. **モック設定をClaude提案**
   ```  
   "SpreadsheetAppをモックしてテストしてください"
   ```
3. **カバレッジ向上支援**
   ```
   "テストカバレッジを向上させるケースを追加してください"
   ```

### 利用可能なモック
- **SpreadsheetApp**：シート操作
- **PropertiesService**：設定管理  
- **CacheService**：キャッシュ操作
- **UrlFetchApp**：外部API呼び出し
- **HtmlService**：HTML出力
- **Utilities**：ユーティリティ関数
- **Logger/console**：ログ出力

---

## 15) AI開発での注意点・制約

### Claude提案コードの必須レビュー観点
1. **GAS実行時間制限**：6分以内で完了するか？
2. **API呼び出し効率**：バッチ処理になっているか？  
3. **権限スコープ**：必要最小限の権限か？
4. **エラー処理**：適切な例外ハンドリングがあるか？
5. **ログ出力**：デバッグに必要な情報を出力しているか？

### プロジェクト固有ルール遵守確認
- **命名規則**：既存コードベースとの統一性
- **ファイル構成**：適切な場所への配置  
- **依存関係**：不要なライブラリ追加の回避
- **セキュリティ**：秘匿情報の取り扱い

### 本番デプロイ前チェックリスト
```bash
# 1. 全テスト通過確認
npm run test

# 2. コード品質チェック
npm run lint

# 3. フォーマット統一
npm run format  

# 4. 統合チェック
npm run check

# 5. GASデプロイ
npm run deploy
```

---

## 16) 個人開発向け自動化ワークフロー

個人開発でも効率的で安全な開発を実現する段階的アプローチ。

### 基本開発フロー（推奨）

#### 🚀 新機能開発時の自動化ワークフロー

```bash
# 1. テスト監視モード開始（別ターミナルで常時実行）
npm run test:watch

# 2. Claude Codeでの開発
# - TDDでテストケースを先に作成
# - Claudeに実装を依頼
# - リアルタイムでテスト結果を確認
# - 失敗時はClaudeに修正を依頼

# 3. 完成後の品質チェック＆デプロイ
npm run deploy  # テスト→デプロイの一括実行
```

#### 📝 実際の操作例

```bash
# Terminal 1: テスト監視開始
npm run test:watch

# Terminal 2: 開発作業
# Claude Code: "新しい関数XXXのテストケースを作成してください"
# Claude Code: "テストに合格する実装を作成してください" 
# Claude Code: "エラーが出ています。修正してください"

# 開発完了後
npm run deploy
```

### ブランチ戦略（個人開発向け）

#### 🎯 最小構成：mainブランチ運用（現在）
**メリット**：
- シンプルで迷わない
- オーバーヘッドなし
- 小規模変更に最適

**デメリット**：
- 実験的な変更でmainが壊れるリスク
- ロールバックが困難

#### 🌲 推奨：feature/mainブランチ戦略

```bash
# 新機能開発時
git checkout -b feature/新機能名
npm run test:watch  # 開発開始

# 開発→テスト→マージの自動化
git add .
git commit -m "feat: 新機能の実装

🤖 Generated with [Claude Code](https://claude.ai/code)

Co-Authored-By: Claude <noreply@anthropic.com>"

# mainへのマージ（安全確認付き）
git checkout main
git pull  # 念のため最新取得
npm run check  # 品質チェック
git merge feature/新機能名
npm run deploy  # GASデプロイ

# クリーンアップ
git branch -d feature/新機能名
```

#### 🔄 自動化スクリプト案

`.scripts/new-feature.sh` (作成推奨):
```bash
#!/bin/bash
# 新機能開発開始の自動化

if [ -z "$1" ]; then
  echo "使用法: ./scripts/new-feature.sh <機能名>"
  exit 1
fi

FEATURE_NAME="feature/$1"

# ブランチ作成・切り替え
git checkout -b "$FEATURE_NAME"

# テスト監視開始（バックグラウンド）
echo "テスト監視モードを開始します..."
npm run test:watch &
TEST_PID=$!

echo "🚀 新機能 '$1' の開発環境が準備完了！"
echo "📝 Claude Codeでテスト→実装→修正のサイクルを開始してください"
echo "✅ 完了後は ./scripts/merge-feature.sh $1 を実行"
echo "🛑 テスト監視停止: kill $TEST_PID"
```

`.scripts/merge-feature.sh` (作成推奨):
```bash
#!/bin/bash
# 機能完成後のマージ・デプロイ自動化

if [ -z "$1" ]; then
  echo "使用法: ./scripts/merge-feature.sh <機能名>"
  exit 1
fi

FEATURE_NAME="feature/$1"
CURRENT_BRANCH=$(git branch --show-current)

# 現在のブランチ確認
if [ "$CURRENT_BRANCH" != "$FEATURE_NAME" ]; then
  echo "❌ エラー: $FEATURE_NAME ブランチに切り替えてください"
  exit 1
fi

# 最終チェック
echo "🔍 最終品質チェック中..."
npm run check
if [ $? -ne 0 ]; then
  echo "❌ テストが失敗しました。修正してから再実行してください"
  exit 1
fi

# コミット（未コミットがあれば）
if ! git diff-index --quiet HEAD --; then
  echo "📝 変更をコミットしています..."
  git add .
  git commit -m "feat: $1 の実装完了

🤖 Generated with [Claude Code](https://claude.ai/code)

Co-Authored-By: Claude <noreply@anthropic.com>"
fi

# mainにマージ
echo "🔀 mainブランチにマージ中..."
git checkout main
git pull  # リモートの最新を取得
npm run check  # mainでも確認
git merge "$FEATURE_NAME"

# デプロイ
echo "🚀 GASデプロイ中..."
npm run deploy

if [ $? -eq 0 ]; then
  echo "✅ デプロイ成功！"
  echo "🧹 フィーチャーブランチをクリーンアップしますか？ (y/N)"
  read -r response
  if [[ "$response" == "y" ]]; then
    git branch -d "$FEATURE_NAME"
    echo "🗑️ $FEATURE_NAME ブランチを削除しました"
  fi
else
  echo "❌ デプロイに失敗しました"
  exit 1
fi

echo "🎉 機能 '$1' のリリース完了！"
```

### 緊急時・実験時の運用

#### 🔥 ホットフィックス（緊急修正）
```bash
# mainで直接修正（小さな修正のみ）
npm run test:watch  # 監視開始
# Claude Codeで修正
npm run deploy     # 即座にデプロイ
```

#### 🧪 実験的な機能
```bash
git checkout -b experiment/機能名
# 自由に実験
# 成功したらfeatureブランチにリネーム
# 失敗したらブランチ削除
```

### ワークフロー選択指針

| 変更規模 | 推奨ワークフロー | 理由 |
|---------|----------------|------|
| バグ修正（1-2ファイル） | main直接 | オーバーヘッド回避 |
| 新機能追加 | feature/main | 安全性確保 |
| 大規模リファクタ | feature/main | ロールバック可能性 |
| 実験的な変更 | experiment/ | main汚染防止 |

### Claude Code連携のコツ

1. **ブランチ作成を明示**：
   ```
   "feature/ユーザー認証 ブランチで新機能を開発します"
   ```

2. **テスト先行を指示**：
   ```
   "まずテストケースを作成してからユーザー認証機能を実装してください"
   ```

3. **段階的な確認**：
   ```
   "テストが通ったらコミットして、次の機能に進んでください"
   ```

4. **自動化スクリプトの活用**：
   ```
   "./scripts/new-feature.sh ユーザー認証 を実行してください"
   ```

---
