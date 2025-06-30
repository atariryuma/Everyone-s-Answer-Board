# StudyQuest - みんなの回答ボード: APIリファレンス (Admin Logger API)

このドキュメントでは、「StudyQuest - みんなの回答ボード」の「管理者向けログ記録API」のエンドポイント、リクエスト形式、およびレスポンス形式について説明します。このAPIは、メインアプリケーションからのユーザー情報管理と監査ログの記録を目的としています。

## 1. エンドポイント

「管理者向けログ記録API」は、Google Apps Scriptのウェブアプリとしてデプロイされます。デプロイ時に発行されるウェブアプリURLがAPIのエンドポイントとなります。

*   **HTTPメソッド**: `POST` (主要なAPI呼び出し用), `GET` (接続テスト用)
*   **URL**: `https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec`
    *   `YOUR_DEPLOYMENT_ID` は、APIデプロイ時に取得した実際のデプロイIDに置き換えてください。

## 2. 認証

このAPIは、Google Apps Scriptのウェブアプリのセキュリティモデルに依存します。デプロイ時に「アクセスできるユーザー」を「全員（匿名ユーザーを含む）」に設定することで、メインアプリケーションからのアクセスを許可します。

**セキュリティに関する注意点**:
*   APIのURLは秘匿情報として扱い、メインアプリケーションの設定のみに保存してください。
*   すべてのAPI呼び出しは、APIが管理するスプレッドシートにログとして記録されます。

## 3. APIリクエスト形式

すべての `POST` リクエストは `application/json` の `Content-Type` ヘッダーを持ち、以下のJSON構造に従う必要があります。

```json
{
  "action": "[アクション名]",
  "data": { /* アクションに応じたデータ */ },
  "timestamp": "[ISO 8601形式のタイムスタンプ]",
  "requestUser": "[リクエスト元のユーザーメールアドレス]",
  "effectiveUser": "[GASのSession.getEffectiveUser().getEmail()の値]"
}
```

*   `action` (string, 必須): 実行するAPIアクションの名前。
*   `data` (object, 必須): 各アクションに固有のペイロードデータ。
*   `timestamp` (string, 任意): リクエストが生成された時刻のISO 8601形式の文字列。監査ログ用。
*   `requestUser` (string, 任意): リクエストを送信したユーザーのメールアドレス。監査ログ用。
*   `effectiveUser` (string, 任意): GASの `Session.getEffectiveUser().getEmail()` の値。監査ログ用。

## 4. APIレスポンス形式

すべてのAPIレスポンスは `application/json` 形式で返され、以下の構造に従います。

```json
{
  "success": true, // または false
  "message": "[成功またはエラーメッセージ]",
  "data": { /* アクションに応じたデータ */ },
  "error": "[エラー詳細 (success: falseの場合)]"
}
```

*   `success` (boolean): リクエストが成功したかどうかを示します。
*   `message` (string, 任意): 処理結果に関する追加情報。
*   `data` (object, 任意): 成功した場合に返されるデータ。
*   `error` (string, 任意): `success` が `false` の場合に、エラーの詳細な説明。

## 5. 利用可能なAPIアクション

### 5.1. `ping`

APIが正常に動作しているかを確認するためのテストエンドポイントです。

*   **アクション**: `ping`
*   **`data` ペイロード**: 
    ```json
    {
      "test": true
    }
    ```
*   **レスポンス `data`**: 
    ```json
    {
      "pong": true,
      "requestUser": "[リクエスト元のユーザーメールアドレス]",
      "effectiveUser": "[GASのSession.getEffectiveUser().getEmail()の値]"
    }
    ```

### 5.2. `getUserInfo`

指定された `userId` に基づいてユーザー情報を取得します。

*   **アクション**: `getUserInfo`
*   **`data` ペイロード**: 
    ```json
    {
      "userId": "[取得するユーザーのID]"
    }
    ```
*   **レスポンス `data`**: 
    ```json
    {
      "userId": "string",
      "adminEmail": "string",
      "spreadsheetId": "string",
      "spreadsheetUrl": "string",
      "createdAt": "string (ISO 8601)",
      "accessToken": "string",
      "configJson": "string (JSON形式の文字列)",
      "lastAccessedAt": "string (ISO 8601)",
      "isActive": "boolean"
    }
    ```
    *   ユーザーが見つからない場合、`success: false` と `error: "ユーザーが見つかりません"` が返されます。

### 5.3. `createUser`

新しいユーザーを作成し、データベースに保存します。

*   **アクション**: `createUser`
*   **`data` ペイロード**: 
    ```json
    {
      "userId": "[新しいユーザーのID]",
      "adminEmail": "[管理者のメールアドレス]",
      "spreadsheetId": "[関連するスプレッドシートID]",
      "spreadsheetUrl": "[関連するスプレッドシートURL]",
      "accessToken": "[アクセストークン (任意)]",
      "configJson": "[ユーザー固有の設定 (JSON形式の文字列, 任意)]",
      "isActive": "[アクティブ状態 (boolean, 任意, デフォルトはtrue)]"
    }
    ```
*   **レスポンス `data`**: 
    ```json
    {
      "userId": "string",
      "adminEmail": "string",
      "createdAt": "string (ISO 8601)"
    }
    ```

### 5.4. `updateUser`

既存のユーザー情報を更新します。

*   **アクション**: `updateUser`
*   **`data` ペイロード**: 
    ```json
    {
      "userId": "[更新するユーザーのID]",
      "spreadsheetId": "[新しいスプレッドシートID (任意)]",
      "spreadsheetUrl": "[新しいスプレッドシートURL (任意)]",
      "accessToken": "[新しいアクセストークン (任意)]",
      "configJson": "[新しいユーザー固有の設定 (JSON形式の文字列, 任意)]",
      "isActive": "[新しいアクティブ状態 (boolean, 任意)]"
    }
    ```
*   **レスポンス `data`**: 
    ```json
    {
      "userId": "string",
      "updatedAt": "string (ISO 8601)"
    }
    ```
    *   ユーザーが見つからない場合、`success: false` と `error: "ユーザーが見つかりません"` が返されます。

### 5.5. `getExistingBoard`

指定された `adminEmail` に基づいて既存のボード情報（ユーザー情報）を取得します。

*   **アクション**: `getExistingBoard`
*   **`data` ペイロード**: 
    ```json
    {
      "adminEmail": "[検索する管理者のメールアドレス]"
    }
    ```
*   **レスポンス `data`**: `getUserInfo` と同じ構造のユーザー情報オブジェクト。
    *   ボードが見つからない場合、`success: false` と `message: "既存ボードが見つかりません"` が返されます。

### 5.6. `checkExistingUser`

指定された `adminEmail` を持つユーザーが既に存在するかどうかを確認します。

*   **アクション**: `checkExistingUser`
*   **`data` ペイロード**: 
    ```json
    {
      "adminEmail": "[チェックする管理者のメールアドレス]"
    }
    ```
*   **レスポンス `data`**: 
    ```json
    {
      "exists": true, // または false
      "data": { /* ユーザーが存在する場合、getUserInfo と同じ構造のユーザー情報オブジェクト */ }
    }
    ```

