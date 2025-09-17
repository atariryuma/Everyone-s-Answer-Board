# ConfigJson 最適化計画

## 🎯 最適化されたconfigJson構造

```javascript
{
  // === 基本情報 ===
  "configVersion": "3.1",
  "etag": "72a440a2-0fef-41b1-ad5a-9e971f71840a-1757571949929",
  "claudeMdCompliant": true,

  // === データソース（統合） ===
  "dataSource": {
    "spreadsheetId": "1bMfeh98hAUpG9adstAIh5qMtdO4xnZf49CJ4a1Tb0ME",
    "sheetName": "フォームの回答 1",
    "sourceKey": "1bMfeh98hAUpG9adstAIh5qMtdO4xnZf49CJ4a1Tb0ME::フォームの回答 1",
    "formId": "1vDy5XeDoJXkRo7w4sit7OhNndt3bEtiUkV2CbRCJO9Q",
    "formTitle": "月の形"
  },

  // === 列構造（統合） ===
  "schema": {
    "headers": ["タイムスタンプ","メールアドレス","クラス","名前","月の形がどのように変わるのか、予想しよう。","予想の理由を書きましょう。","なるほど！","いいね！","もっと知りたい！","ハイライト"],
    "headersHash": "d61bec3029a96dad83af6bb60ea8a54f9e12c5ad07b3da985cd4fd6e54dbea05",
    "systemMetadata": {
      "timestamp": {"header": "タイムスタンプ", "index": 0},
      "email": {"header": "メールアドレス", "index": 1}
    },
    "columnMapping": {
      "confidence": {"reason": 95, "name": 98, "class": 98, "answer": 100},
      "mapping": {"reason": 5, "name": 3, "answer": 4, "class": 2}
    },
    "reactionMapping": {
      "UNDERSTAND": {"header": "なるほど！", "index": 6},
      "CURIOUS": {"index": 8, "header": "もっと知りたい！"},
      "HIGHLIGHT": {"index": 9, "header": "ハイライト"},
      "LIKE": {"header": "いいね！", "index": 7}
    }
  },

  // === アプリケーション状態（統合） ===
  "appState": {
    "setupStatus": "completed",           // setupComplete削除
    "published": true,                    // appPublished/isDraft統合
    "displayMode": "anonymous",
    "displaySettings": {
      "showNames": false,
      "showReactions": false
    }
  },

  // === URLs（動的生成推奨） ===
  "urls": {
    "appUrl": "https://script.google.com/a/naha-okinawa.ed.jp/macros/s/AKfycbxPSgTPmCacJBE1LyJNy-IEanq6ASJqy_2uBrfXi_mM-umtxk85WlKWLVFhOp3exVvR/exec?mode=view&userId=adb24b94-8244-4d3a-a1c3-e409f81e40a0"
  },

  // === タイムスタンプ（統合・明確化） ===
  "timestamps": {
    "createdAt": "2025-09-06T23:46:41.002Z",
    "publishedAt": "2025-09-16T22:36:39.624Z",
    "lastModified": "2025-09-16T22:36:39.624Z",
    "verifiedAt": "2025-09-16T22:36:26.998Z",
    "lastAccessedAt": "2025-09-16T22:36:39.624Z"  // 修正
  },

  // === 新規追加フィールド ===
  "security": {
    "permissions": ["read", "write"],
    "accessLevel": "admin"
  },

  "performance": {
    "dataSize": {
      "rowCount": 0,
      "columnCount": 10
    },
    "lastSync": "2025-09-16T22:36:39.624Z"
  },

  "errorHandling": {
    "lastError": null,
    "errorCount": 0,
    "retryCount": 0
  }
}
```

## 📊 削減効果

### 削除対象フィールド (7個)
1. `setupComplete` → `appState.setupStatus` に統合
2. `appPublished` → `appState.published` に統合
3. `isDraft` → `appState.published` の逆として削除
4. `spreadsheetUrl` → 動的生成に変更
5. `formUrl` → 動的生成に変更
6. `questionText` → `schema.columnMapping` から取得
7. `lastAccessedAt` → 修正が必要（9日前は古すぎ）

### 構造化効果
- **27フィールド** → **6つのカテゴリ** に整理
- データの関連性が明確化
- 重複を削減してデータ整合性を向上

### メンテナンス性向上
- 関連フィールドのグループ化
- 命名規則の統一
- 将来の拡張性を考慮した設計