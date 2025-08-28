# 未定義エラー修正プラン

## エラー分析結果

テストで検出された未定義エラー：
- **未定義関数**: 242個
- **未定義クラス**: 48個  
- **未定義変数**: 410個
- **合計**: 700個

## カテゴリー別修正戦略

### 1. 高優先度: 実際の未定義エラー (即座に修正が必要)

#### 1.1 欠落クラス定義
```
- ULog (複数ファイルで参照: Core.gs, lockManager.gs, unifiedValidationSystem.gs等)
- UError (main.gs, workflowValidation.gs)
- UnifiedErrorHandler (Core.gs)
```

#### 1.2 欠落関数定義
```
- clearTimeout (config.gs, resilientExecutor.gs) - JavaScript標準のため要修正
- fromEntries (systemIntegrationManager.gs) - Object.fromEntriesの誤参照
- URLSearchParams (unifiedBatchProcessor.gs) - GASでは非対応
```

### 2. 中優先度: 不完全な実装参照

#### 2.1 クラスコンストラクタ参照エラー
```
- ResilientExecutor (resilientExecutor.gs) - クラス内でnewしようとしている
- UnifiedBatchProcessor (config.gs, unifiedBatchProcessor.gs)
- SystemIntegrationManager (systemIntegrationManager.gs)
- UnifiedSecretManager (secretManager.gs)
```

#### 2.2 メソッドチェーン参照エラー
```
FormApp関連:
- setEmailCollectionType, setCollectEmail, getPublishedUrl, getEditUrl
- addListItem, setChoiceValues, setRequired, addTextItem
- addParagraphTextItem, setHelpText, createParagraphTextValidation
```

### 3. 低優先度: 偽陽性 (フィルター改善で対応)

#### 3.1 CSS/HTML関連 (main.gs内)
```
- gradient, rgba, translateY, scale, blur, bezier, media
- preventDefault, querySelector, reload, confirm, alert
```

#### 3.2 プロパティ名・引数名の誤検出
```
- data, status, method, column, setup, route, start, end
- success, failed, fallback, pending, completed
```

## 修正手順

### Phase 1: 重要な欠落クラス・関数を実装

1. **ULogクラスの実装**
   - Core.gsまたは専用ファイルで統合ログクラスを作成
   - 既存のLogger/console.logをラップ

2. **UErrorクラスの実装**  
   - 統合エラーハンドリングクラスを作成
   - Errorクラスを拡張

3. **JavaScript標準関数の修正**
   - clearTimeout → Utilities.sleep代替実装
   - Object.fromEntries → 手動実装またはポリフィル
   - URLSearchParams → GAS互換実装

### Phase 2: 不完全なクラス参照を修正

1. **自己参照エラーの修正**
   - クラス内でのnew呼び出しを静的メソッドに変更
   - ファクトリーパターンの適用

2. **FormApp APIメソッドの修正**
   - 正しいFormApp APIチェーンを確認
   - 未対応メソッドの代替実装

### Phase 3: テストフィルターの改善

1. **偽陽性の除外**
   - CSS/HTML関連キーワードをフィルターに追加
   - プロパティ名の誤検出を除外

2. **プロジェクト固有関数の定義**
   - 正当な参照をbuiltin扱いに追加

## 修正優先順位リスト

### 🚨 緊急 (今すぐ修正)
1. ULog クラス実装 → 5ファイルで参照
2. UError クラス実装 → 2ファイルで参照  
3. clearTimeout代替実装 → 2ファイルで参照
4. Object.fromEntries実装 → 1ファイルで参照

### ⚡ 高 (次に修正)
1. UnifiedErrorHandler実装
2. 自己参照クラス修正 (ResilientExecutor等)
3. FormApp APIメソッド修正

### 📋 中 (段階的に修正)
1. 各unified*クラスの内部参照修正
2. メソッドチェーン参照の確認・修正

### 🔧 低 (最後に修正)
1. テストフィルター改善
2. 偽陽性の除外

## 期待される効果

修正後の予想エラー数：
- **Phase 1完了後**: ~600個 (100個減少)
- **Phase 2完了後**: ~300個 (300個減少)  
- **Phase 3完了後**: ~50個 (250個減少)
- **最終目標**: 0-10個の真の未定義エラーのみ