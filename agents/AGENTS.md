# Gemini 指示書 for "StudyQuest：みんなの回答ボード"

## 1. プロジェクト概要 (Project Overview)

このプロジェクトは、教育現場における双方向学習を促進するためのWebアプリケーション「StudyQuest：みんなの回答ボード」です。Google Apps Script (GAS)をバックエンド、Googleスプレッドシートをデータベースとして利用し、サーバーレスアーキテクチャで構築されています。

**コアコンセプト:**
* **サービスアカウントモデル:** 従来の権限問題を解決し、導入を容易にします。
* **教師中心の管理:** 教師ユーザーがボード（回答データシート）を自由に作成・管理できます。
* **リアルタイム性と安全性:** 生徒はリアルタイムで回答を共有・閲覧でき、管理者はドメイン内のアクセス制御や匿名モードで安全な環境を維持できます。

## 2. アーキテクチャと主要ファイル (Architecture & Key Files)

* **バックエンド (`/src/*.gs`):**
    * Google Apps Script (V8ランタイム) でビジネスロジックを実装します。
    * `main.gs`: `doGet(e)` を起点とするリクエストのルーティングを行います。
    * `database.gs`: Googleスプレッドシート（中央DB、回答データシート）とのデータ連携を担当します。
    * `AuthorizationService.gs`: ユーザー認証・認可ロジックを管理します。
    * `Core.gs`: ユーザー登録やボード作成などのコア機能を提供します。
* **フロントエンド (`/src/*.html`):**
    * HTML、CSS (Tailwind CSS)、JavaScriptでUIを構築します。
    * `AdminPanel.html`: 教師向けの管理画面です。
    * `Page.html`: 生徒向けの回答ボード閲覧画面です。
    * `Registration.html`: 新規ユーザー登録画面です。
    * `SharedUtilities.html` / `UnifiedStyles.html`: 共通のJSユーティリティとCSSスタイルを定義しています。
    * `ClientOptimizer.html`: `google.script.run` の呼び出しを最適化するクライアントサイドのライブラリです。
* **テスト (`/tests/*.test.js`):**
    * Jestフレームワークを使用した単体テストおよび結合テストです。
* **設定ファイル:**
    * `appsscript.json`: OAuthスコープやデプロイ設定を定義します。
    * `package.json`: claspやJestなどの開発用ライブラリを管理します。

## 3. コーディング規約 (Coding Conventions)

`agents/AGENTS.md` の規約に厳密に従ってください。

* **命名規則:**
    * 変数・関数: `camelCase`
    * 定数: `UPPER_SNAKE_CASE`
    * クラス・コンストラクタ: `PascalCase`
* **JSDoc:** 全てのサーバーサイド関数には、引数 (`@param`) と戻り値 (`@returns`) を含む詳細なJSDocを記述してください。
    ```javascript
    /**
     * 指定されたIDのユーザー情報を取得します。
     * @param {string} userId - ユーザーの一意なID。
     * @returns {object | null} ユーザー情報オブジェクト。見つからない場合はnull。
     */
    function findUserById(userId) { /* ... */ }
    ```
* **クライアント/サーバー分離:**
    * **サーバーサイド(.gs):** ビジネスロジックのみを実装します。DOM操作は一切含めません。
    * **クライアントサイド(.html):** UIの構築と、`google.script.run` を介した非同期なデータ操作に専念します。

## 4. 主要なロジックとデータモデル (Core Logic & Data Models)

### 4.1. データモデル

* **中央データベース (`Users`シート):**
    * `userId`: 一意なユーザーID (UUID)
    * `adminEmail`: 管理者のメールアドレス
    * `spreadsheetId`: 回答データシートのID
    * `configJson`: 各ボードの設定を格納するJSON文字列。この中には `publishedSheet` (公開中のシート名) や列マッピング情報が含まれます。
* **回答データシート:**
    * `タイムスタンプ`, `メールアドレス`は必須列です。
    * `回答`, `理由`, `名前`, `クラス` の列は `configJson` で動的にマッピングされます。
    * `なるほど！`, `いいね！`, `もっと知りたい！`: リアクションしたユーザーのメールアドレスがカンマ区切りで入ります。
    * `ハイライト`: `TRUE`/`FALSE`でハイライト状態を管理します。

### 4.2. ショートカットプロンプト (Shortcut Prompts)

**一般的なタスク:**
* `test:unit @path/to/file.gs`: 指定されたGASファイルの関数に対するJestの単体テストコードを `tests/` ディレクトリに生成してください。
* `doc:jsdoc @path/to/file.gs`: 指定されたファイルの全ての関数に、規約に沿ったJSDocコメントを追加してください。
* `refactor:clean @path/to/file.gs`: 指定されたファイルを、コーディング規約（特に `guard clauses` の使用）に従ってリファクタリングしてください。

**フロントエンド開発:**
* `ui:component 新しい設定項目用のトグルスイッチ`: `AdminPanel.html` に、`UnifiedStyles.html` のデザインシステムを使って、指定されたUIコンポーネントを追加してください。
* `gas:call "getPublishedSheetData"`: `Page.html` 内で、指定されたサーバーサイド関数を呼び出すための `google.script.run` を使ったJavaScriptコードを記述してください。`ClientOptimizer.html` の `runGas` ラッパー関数を使用してください。

**仕様に基づく実装:**
* `impl:feature ADM-005`: `README.md` の機能要件ID `ADM-005`（列マッピング設定）を実装するためのサーバーサイド関数 `saveSheetConfig()` と、それを呼び出すクライアントサイドのコードを記述してください。
