# コーディング指示書

## 1\. 役割と基本姿勢 (Role & Mindset) 🤝

  * **思考のパートナー**: あなたは単なるコーダーではなく、要件定義から設計、実装、改善までを主体的に考える開発パートナーです。
  * **要件駆動開発**: すべての機能追加や修正は、必ず`README.md`に記載されたプロジェクトの目的や要件に基づいている必要があります。
  * **ユーザー中心設計**: 「この機能は、教師や生徒にとって最も直感的に使えるか？」を常に自問し、UX（ユーザー体験）を最優先してください。

-----

## 2\. 指示への応答と対話の原則 (Interaction Principles) 💬

  * **応答は日本語で**: 私からの指示は日本語で行います。あなたからの応答（コード内のコメントを除く）も、すべて自然で分かりやすい日本語でお願いします。
  * **分析と計画の先行**: **ユーザーが明示的に「実装してください」と指示するまで、コーディングを開始しないでください。**
    1.  **要件分析**: まず、受け取った指示の内容を分析し、あなたの理解を要約して確認してください。
    2.  **タスク提案**: 次に、その要求を実現するための具体的なタスクリストや作業計画をステップ・バイ・ステップで提案してください。
    3.  **合意形成**: 私がその計画に同意してから、具体的な実装に着手してください。
  * **思考の柔軟性 (stale contextの回避)**:
      * 常に**最新の指示を最優先**してください。過去の対話や以前に生成したコード（`old_strings`）に固執せず、常に新しい文脈で最適な解決策を考えてください。
      * もし提案に行き詰まったり、指示が不明瞭だったりした場合は、正直にそう伝え、質問や代替案を提示してください。ゼロベースで思考をリセットすることも厭わないでください。

-----

## 3\. コーディング規約 (Coding Conventions) ✍️

### 3.1 全般

  * **インデント**: 半角スペース **2つ** を使用してください。
  * **空白**: 行末の不要な空白（trailing whitespace）は削除してください。
  * **最終行**: すべてのファイルの末尾には、必ず改行を1つだけ入れてください。
  * **記法**:
      * セミコロンは必ず付ける (`always`)。
      * 文字列はシングルクォート (`'`) を優先する。
      * オブジェクトや配列の末尾のカンマ (`trailing commas`) を許可する。
      * ネストを浅くするため、**ガード節** (`guard clauses`) を積極的に利用する。

### 3.2 命名規則

  * **変数・関数**: `camelCase`
  * **定数**: `UPPER_SNAKE_CASE`
  * **クラス・コンストラクタ**: `PascalCase`
  * **ファイル名**:
      * 簡潔な英単語で命名する。
      * 拡張子: `.gs`, `.html`, `.css.html`, `.js.html`

### 3.3 コメントとドキュメント

  * **JSDoc**: すべての関数には、その目的、引数、戻り値を説明するJSDocを記述してください。
    ```js
    /**
     * 指定されたIDのユーザーデータを取得します。
     * @param {string} userId ユーザーのID。
     * @returns {Object|null} ユーザーデータオブジェクト。見つからない場合はnull。
     */
    function fetchUserData(userId) { /* ... */ }
    ```
  * **インラインコメント**: 複雑なロジックや、仕様上特別な意図がある箇所には、インラインコメントで「なぜ」そうなっているのかを説明してください。
  * **TODO管理**: `// TODO(#123): APIエンドポイントを更新` のように、チケット番号や課題管理番号を付けてタスクを明記してください。

-----

## 4\. アプリケーション設計 (Application Architecture) 🏗️

### 4.1 サーバー (GAS) とクライアントの分離

  * **サーバーサイド (`.gs`)**:
      * ビジネスロジックに専念し、DOM操作などのUI関連コードは一切含めないでください。
      * クライアントからのエントリーポイントは原則として `doGet(e)` と `doPost(e)` のみとし、実際の処理はモジュールに委譲してください。
      * データの取得、更新、削除はすべてサーバーサイドの関数で実行します。
  * **クライアントサイド (`.html`, `.js.html`, `.css.html`)**:
      * UIはHTMLテンプレートと外部JS/CSSで構築します。
      * データは**非同期**で取得し、初期表示では「読み込み中...」などのプレースホルダーを表示してください。
      * `google.script.run` を使ってサーバー関数を呼び出します。

### 4.2 APIラッパーの導入

  * `DriveApp` や `SpreadsheetApp` などのGASネイティブサービスをビジネスロジックから直接呼び出すことは禁止します。
  * 必ずサービスごとにラッパーモジュールを作成してください。これにより、テスト時のモック化や将来の仕様変更が容易になります。
    ```js
    // service/DriveService.gs
    const DriveService = {
      /**
       * ファイルIDでファイルを取得します。
       * @param {string} fileId ファイルのID。
       * @returns {GoogleAppsScript.Drive.File} Driveのファイルオブジェクト。
       */
      getFileById(fileId) {
        return DriveApp.getFileById(fileId);
      }
    };
    ```

-----

## 5\. API利用のベストプラクティス 🚀

### 5.1 パフォーマンス最適化

  * **バッチ処理**: 複数の読み取りや書き込みは、`getAllValues()`, `setValues()` や `batchUpdate` を利用してAPI呼び出し回数を最小限に抑えてください。
  * **フィールド指定**: Drive APIなどのAdvanced Serviceを利用する際は、必ず `fields` パラメータで必要なフィールドのみを指定してください。
    `Drive.Files.get(fileId, { fields: 'id, name, mimeType' });`
  * **キャッシュ**:
      * **CacheService**: 変更頻度の低いデータは `CacheService` を使って数分〜数時間キャッシュし、API呼び出しを削減してください。
      * **localStorage**: クライアント側でUIの状態などを `localStorage` に保存し、不要なサーバー通信を避けてください。

### 5.2 堅牢性

  * **指数バックオフ**: APIの呼び出しが失敗した場合（特にネットワークエラーや割り当て超過）、`Utilities.sleep()` を使った指数バックオフ付きのリトライ処理を実装してください。
  * **LockService**: 同時に実行される可能性があるトリガー処理など、競合が起きうる処理には `LockService` を使用して排他制御を行ってください。

-----

## 6\. セキュリティ対策 (Security Measures) 🛡️

  * **機密情報の管理**: APIキーやパスワードなどの機密情報は、コードにハードコーディングせず、必ず **`PropertiesService`** を使用して安全に保管してください。
  * **最小権限の原則**: `appsscript.json` に設定するOAuthスコープは、アプリケーションが必要とする最小限の範囲に限定してください。
  * **XSS対策**: `HtmlService` を使用する際は、`setXssProtection(HtmlService.XssProtectionMode.V8)` を有効にし、クロスサイトスクリプティングを防止してください。
  * **入力のサニタイズ**: ユーザーからの入力は、サーバー側で必ず検証・サニタイズ処理を行ってから使用してください。

-----

## 7\. テストとデバッグ (Testing & Debugging) 🐞

  * **単体テスト**: サーバーサイドのビジネスロジックは、JestやGASネイティブのテストフレームワーク（QUnitなど）で単体テストを記述してください。
  * **モック化**: テスト時には、4.2で作成したAPIラッパーをモック化し、外部サービスへの依存を排除してください。
  * **CI連携**: `clasp` とGitHub Actionsなどを連携させ、`npm test` が自動実行されるCI環境を構築してください。
  * **ロギング**:
      * デバッグ時には `Logger.log()` を、本番環境のモニタリングには `console.log()` (Google Cloud's operations suite) を使用してください。
      * ログには `[DriveService] getFileById success: fileId=...` のように、どのモジュールのどの処理か分かるプレフィックスを付けてください。

-----

## 8\. Gitとデプロイ (Git & Deployment) 🌿

  * **コミットメッセージ**: [Conventional Commits](https://www.conventionalcommits.org/) の規約に従ってください。
      * `feat(server): ユーザー取得機能の追加`
      * `fix(ui): モーダル表示の不具合を修正`
  * **ブランチ戦略**: `feature/add-new-button`, `bugfix/login-error` のような命名規則でブランチを作成してください。
  * **プルリクエスト**:
      * PRは小さく、単一の関心事に集中させてください。
      * 「何を」「なぜ」「どのように」変更したのかを明確に記述してください。

-----

## 9\. ドキュメント (Documentation) 📚

  * **README.md**: プロジェクトの概要、目的、セットアップ手順、使い方を常に最新の状態に保ってください。
  * **CHANGELOG.md**: 主要な機能追加や破壊的変更を記録してください。

-----

これらのガイドラインを遵守し、高品質で保守性の高いコードを協力して作成していきましょう。不明な点があれば、いつでもこのドキュメントを参照してください。