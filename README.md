# Everyone's Answer Board (みんなの回答ボード) - StudyQuest

## 概要

**Everyone's Answer Board** は、Google Apps Script (GAS) を利用して構築された、リアルタイムでインタラクティブな回答共有ウェブアプリケーションです。教育現場やワークショップなどで、参加者からの意見や回答を即座に集約し、全員で閲覧・共有することを目的としています。

Googleスプレッドシートをデータベースとして活用し、Googleフォームと連携することで、簡単に質問を投げかけ、回答をボード形式で表示することができます。

## 📊 システム統計情報

### コードベース規模
- **総ファイル数**: 47ファイル
- **総コード行数**: 18,000行以上
- **総文字数**: 680,000文字以上
- **関数数**: 300個以上

### ファイル構成分析
| ファイル | サイズ | 行数 | 関数数 | 分割計画 |
|---------|-------|------|--------|----------|
| Core.gs | 238,795字 | 6,238行 | 110個 | 12モジュール |
| config.gs | 142,162字 | 3,725行 | 62個 | 7モジュール |
| database.gs | 112,973字 | 3,164行 | 44個 | 6モジュール |
| main.gs | 74,090字 | 2,156行 | 49個 | 4モジュール |
| cache.gs | 37,235字 | 1,111行 | 16個 | 2モジュール |

## 主な機能

*   **Googleアカウント認証:** 安全なGoogleアカウントでのログイン。
*   **ボード作成:**
    *   新規Googleフォームと連携したボードの自動生成。
    *   既存のGoogleスプレッドシートを回答データベースとして利用。
*   **リアルタイム表示:** 投稿された回答がリアルタイムでボードに反映されます。
*   **インタラクション:**
    *   各回答カードに対するリアクション機能。
    *   管理者による特定の回答のハイライト機能。
*   **柔軟なカスタマイズ:**
    *   ボードに表示する列（項目）を自由に選択・マッピング。
    *   記名/匿名表示の切り替え。
    *   リアクション数の表示/非表示設定。

## 🏗️ アーキテクチャ

### 技術スタック
*   **バックエンド:** Google Apps Script (V8 Runtime)
*   **フロントエンド:** HTML, CSS, JavaScript (TailwindCSS)
*   **データベース:** Google Sheets (with advanced caching layer)
*   **認証:** Google OAuth2
*   **開発ツール:**
    *   `clasp`: Google Apps Scriptプロジェクトのローカル開発用CLIツール
    *   `jest`: JavaScriptコードの単体テスト用フレームワーク

### システム構成図
```
┌─────────────────┐    ┌──────────────────┐    ┌─────────────────┐
│   ユーザー       │◄──►│    Web App       │◄──►│  Google Sheets  │
│  (ブラウザ)      │    │   (GAS HTML)     │    │   (データベース)  │
└─────────────────┘    └──────────────────┘    └─────────────────┘
                              │
                              ▼
                       ┌──────────────────┐
                       │  Google Forms    │
                       │  (回答収集)       │
                       └──────────────────┘
```

## 📁 ファイル構成詳細

### メインGASファイル (.gs)

#### 【大容量ファイル - 分割対象】
1. **Core.gs** (238,795文字, 110関数)
   - 主要な業務ロジックとAPIエンドポイント
   - 🎯 **分割計画**: 12モジュール (coreUtilities, setupValidation, autoStopManager等)

2. **config.gs** (142,162文字, 62関数)
   - 設定管理とヘッダーマッピング
   - 🎯 **分割計画**: 7モジュール (configValidation, configNormalization等)

3. **database.gs** (112,973文字, 44関数)
   - データベース操作とユーザー管理
   - 🎯 **分割計画**: 6モジュール (databaseCore, databaseOperations等)

4. **main.gs** (74,090文字, 49関数)
   - リクエストルーティングとページレンダリング
   - 🎯 **分割計画**: 4モジュール (requestRouting, pageRendering等)

5. **cache.gs** (37,235文字, 16関数)
   - キャッシュ管理システム
   - 🎯 **分割計画**: 2モジュール (cacheManager, cacheOptimization)

#### 【適正サイズファイル】
6. **auth.gs** (15,249文字) - 認証とアクセス制御
7. **errorHandler.gs** (9,634文字) - 統一エラーハンドリング
8. **unifiedUtilities.gs** (14,094文字) - 統一ユーティリティ
9. **url.gs** (14,056文字) - URL管理とルーティング
10. **session-utils.gs** (12,710文字) - セッション管理
11. **monitoring.gs** (11,763文字) - システム監視
12. **commonUtilities.gs** (8,199文字) - 共通ユーティリティ
13. **spreadsheetCache.gs** (7,660文字) - スプレッドシートキャッシュ
14. **unifiedExecutionCache.gs** (5,532文字) - 実行キャッシュ
15. **debugConfig.gs** (5,139文字) - デバッグ設定
16. **lockManager.gs** (2,388文字) - ロック管理
17. **keyUtils.gs** (515文字) - キーユーティリティ

### フロントエンドファイル
- **Page.html** - メイン回答ボード
- **AdminPanel.html** - 管理者パネル
- **LoginPage.html** - ログインページ
- **SharedUtilities.html** - 共通JavaScript
- **UnifiedStyles.html** - 統一スタイルシート

## 🔧 主要機能モジュール

### 1. ユーザー管理システム
```javascript
// 重要関数
- registerNewUser(adminEmail)           // 新規ユーザー登録
- findUserById(userId)                  // ユーザー検索
- verifyUserAccess(requestUserId)       // アクセス検証
- getCurrentUserStatus(requestUserId)   // ユーザー状態取得
```

### 2. 設定管理システム
```javascript
// 重要関数
- getAppConfig(requestUserId)           // アプリ設定取得
- saveSheetConfig(userId, ...)          // シート設定保存
- autoMapHeaders(headers, sheetName)    // ヘッダー自動マッピング
- normalizeConfigJson(configJson)       // 設定正規化
```

### 3. データ管理システム
```javascript
// 重要関数
- getPublishedSheetData(...)            // 公開データ取得
- addReaction(requestUserId, ...)       // リアクション追加
- toggleHighlight(requestUserId, ...)   // ハイライト切替
- refreshBoardData(requestUserId)       // ボードデータ更新
```

### 4. キャッシュシステム
```javascript
// 重要関数
- cacheManager.get(key, valueFn, options)     // キャッシュ取得
- invalidateUserCache(userId, ...)            // ユーザーキャッシュ無効化
- clearAllExecutionCache()                    // 実行キャッシュクリア
```

## ⚠️ 既知の問題と修正状況

### ✅ 修正済み (2025年8月)
1. **🚨 重大バグ修正**: main.gsのdebugLog関数の無限再帰を修正
2. **🔒 セキュリティ強化**: デバッグログから機密情報(ユーザーID、メール)を除去
3. **⚡ パフォーマンス改善**: 不要なキャッシュクリア操作を最適化
4. **🎨 スタイル統一**: インラインスタイルをCSS変数に変換
5. **🔧 関数統合完了**: ログ関数をdebugConfig.gsに統一（4ファイルの重複を解消）

### 🔄 現在対応中
1. **関数重複の統合**:
   - ✅ ログ関数の統合完了 (debugLog, errorLog, warnLog, infoLog → debugConfig.gs)
   - ❌ ユーザー情報取得関数が6個重複
   - ❌ キャッシュ管理関数が散在

2. **大容量ファイルの分割** (優先度: 高):
   - 🎯 Core.gs → 12モジュール (238KB → 各20KB以下)
   - 🎯 config.gs → 7モジュール (142KB → 各20KB以下)
   - 🎯 database.gs → 6モジュール (113KB → 各18KB以下)

### 📋 今後の課題
1. **モジュール分割の完了** - Claude処理効率の向上
2. **型安全性の向上** - JSDocの充実
3. **テストカバレッジの拡充**
4. **API仕様書の整備**

## 🔄 関数重複の整理状況

### 1. ログ関数 (統一完了)
| 関数名 | 出現ファイル | 状態 | 統一先 |
|--------|-------------|------|--------|
| debugLog | debugConfig.gs のみ | ✅ 統一完了 | debugConfig.gs |
| errorLog | debugConfig.gs のみ | ✅ 統一完了 | debugConfig.gs |
| warnLog | debugConfig.gs のみ | ✅ 統一完了 | debugConfig.gs |
| infoLog | debugConfig.gs のみ | ✅ 統一完了 | debugConfig.gs |

### 2. ユーザー情報取得 (段階的統一)
| 関数名 | ファイル | 状態 | 移行計画 |
|--------|---------|------|----------|
| getUserInfoCached | config.gs | 🔄 移行中 | → UnifiedUserManager |
| getUserInfoCachedUnified | unifiedUtilities.gs | ✅ 推奨 | メイン実装 |
| getOrFetchUserInfo | main.gs | ❌ 重複 | 段階的廃止 |
| findUserById | database.gs | ✅ 保持 | DB特化として保持 |

### 3. キャッシュ管理 (中央集約中)
| 関数名 | 現在の状態 | 統一先 |
|--------|-----------|--------|
| clearExecutionUserInfoCache | ❌ 4ファイルで重複 | unifiedExecutionCache.gs |
| invalidateUserCache | ⚠️ 散在 | CacheManager |
| clearAllExecutionCache | ❌ 複数実装 | unifiedExecutionCache.gs |

## セットアップ

### 1. 前提条件

*   [Node.js](https://nodejs.org/) と `npm` がインストールされていること。
*   Googleアカウントを持っていること。
*   `clasp` がインストールされ、Googleアカウントでログイン済みであること。
    ```bash
    npm install -g @google/clasp
    clasp login
    ```

### 2. プロジェクトのクローンと初期設定

```bash
# リポジトリをクローン
git clone https://github.com/atariryuma/Everyone-s-Answer-Board.git

# プロジェクトディレクトリに移動
cd Everyone-s-Answer-Board

# 開発依存関係をインストール
npm install
```

### 3. Google Apps Script プロジェクトの作成

`clasp` を使用して、新規または既存のGASプロジェクトに接続します。

```bash
# 新規プロジェクトを作成する場合
clasp create --title "みんなの回答ボード" --rootDir ./src

# 既存のプロジェクトに接続する場合
clasp clone <scriptId> --rootDir ./src
```

### 4. 環境設定

本アプリケーションを動作させるには、いくつかの認証情報と設定をスクリプトプロパティに保存する必要があります。

#### Script Properties 設定
```javascript
// 必須設定項目
SERVICE_ACCOUNT_CREDS: "サービスアカウント認証情報のJSON"
DATABASE_SPREADSHEET_ID: "データベーススプレッドシートのID"
ADMIN_EMAIL: "管理者メールアドレス"
```

#### OAuth Scopes 設定
```
https://www.googleapis.com/auth/script.external_request
https://www.googleapis.com/auth/forms
https://www.googleapis.com/auth/spreadsheets
https://www.googleapis.com/auth/drive
https://www.googleapis.com/auth/userinfo.email
```

#### Web App 設定
- **Execute as**: User accessing the web app
- **Access**: Anyone with Google account

詳細は以下のガイドを参照してください。
**➡️ [SETUP_GUIDE.md](./SETUP_GUIDE.md)**

## 🛠️ 開発ガイドライン

### コーディング規約
- **関数名**: camelCase (例: getUserInfo)
- **定数**: UPPER_SNAKE_CASE (例: ERROR_SEVERITY)
- **プロパティ名**: camelCase (例: spreadsheetId)
- **ファイル名**: camelCase.gs (例: userManager.gs)

### 新規関数作成時の注意点
1. **重複チェック**: 既存の類似関数がないか確認
2. **統一クラス優先**: UnifiedUserManager, CacheManager等を優先使用
3. **エラーハンドリング**: errorHandler.gsの関数を使用
4. **機密情報保護**: ログ出力で機密情報をマスク

### パフォーマンス考慮事項
- 大量データ処理時はバッチAPIを使用
- キャッシュを適切に活用
- 不要な実行レベルキャッシュクリアを避ける

### コードのデプロイ

```bash
npm run push
# または
clasp push
```

### テスト

`jest` を使用した単体テストが用意されています。

```bash
npm test
```

## 📊 監視・メンテナンス

### ログレベル
- **DEBUG**: 開発環境のみ、詳細な処理追跡
- **INFO**: 重要な処理の開始/終了
- **WARN**: 警告レベルの問題
- **ERROR**: エラーハンドリング必須

### パフォーマンス監視
- 関数実行時間の測定
- キャッシュヒット率の監視
- エラー発生率の追跡
- システムリソース使用量

## 🔮 今後の開発計画

### Phase 1: コードベースの最適化 (進行中)
- ✅ 重大バグの修正完了
- 🔄 関数重複の統合
- 🔄 大容量ファイルの分割

### Phase 2: 機能拡張
- 📱 モバイル対応の強化
- 🎨 UI/UXの改善
- 📊 アナリティクス機能の追加

### Phase 3: スケーラビリティ向上
- ⚡ パフォーマンス最適化
- 🔧 TypeScript対応検討
- 🧪 E2Eテストの導入

## 📄 ライセンス

このプロジェクトは [ISC License](./LICENSE) の下で公開されています。

---

**最終更新**: 2025年8月 | **バージョン**: 2.1 | **メンテナー**: StudyQuest開発チーム
