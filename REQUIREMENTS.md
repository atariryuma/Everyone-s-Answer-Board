# Everyone's Answer Board - 実装ベース要件定義書

## 1. システム概要

### 1.1 プロジェクト名
Everyone's Answer Board（全員回答ボードシステム）

### 1.2 現在の実装状況
Google Apps Script (GAS) を基盤とした **運用中の** Webアプリケーション。Google Forms/Sheetsと連携して、教育現場での質問・回答管理を行う実稼働システム。

### 1.3 技術的特徴
- **Single Page Application**: HTML Template Engine による動的画面生成
- **サーバーレスアーキテクチャ**: Google Cloud Platform 上で完全稼働
- **リアルタイム同期**: ポーリングベースのデータ更新（3-5秒間隔）
- **多層キャッシュ**: メモリ・スクリプト・ユーザーレベルでの高速化

## 2. 現在実装済み機能

### 2.1 認証・アクセス制御 ✅
```javascript
// 実装済み認証フロー
- Google OAuth 2.0 統合認証
- ドメインベースアクセス制限
- 管理者/一般ユーザーの役割分離
- セッション管理（Google Apps Script Session API）
```

### 2.2 コア機能 ✅
```javascript
// メイン機能
- 質問・回答の投稿・表示
- リアクション機能（いいね/なるほど/もっと知りたい）
- リアルタイムポーリング更新
- レスポンシブWebデザイン（モバイル対応）
```

### 2.3 管理者機能 ✅
```javascript
// AdminPanel.html で実装済み
- Google Forms/Sheets の設定・選択
- カラムマッピング設定
- ユーザーアクセス管理
- システム状態監視
- 履歴管理（50件制限）
```

### 2.4 システム機能 ✅
```javascript
// 技術基盤
- 統合キャッシュシステム
- エラーハンドリング・ログ管理
- URL管理（開発/本番環境対応）
- 自動停止・再開機能
```

## 3. 現在の技術仕様

### 3.1 実装技術スタック
```yaml
Backend:
  - Google Apps Script (V8 Runtime)
  - Google Sheets API (Database)
  - Google Forms API (Data Input)

Frontend:
  - HTML5 Template Engine
  - CSS3 + Custom Design System
  - JavaScript ES5+ (GAS compatible)
  - TailwindCSS (CDN fallback)

Infrastructure:
  - Google Cloud Platform
  - Google Workspace Integration
  - Google OAuth 2.0
```

### 3.2 データベース設計（実装済み）
```javascript
// Users Sheet - メインテーブル
const USER_SCHEMA = {
  userId: 'String (Primary Key)',
  adminEmail: 'String (Admin Email)',
  spreadsheetId: 'String (Google Sheets ID)',
  spreadsheetUrl: 'String (Sheets URL)',
  createdAt: 'Date',
  configJson: 'JSON String (Configuration)',
  lastAccessedAt: 'Date',
  isActive: 'Boolean'
};

// DeleteLogs Sheet - 監査ログ
const DELETE_LOG_SCHEMA = {
  timestamp: 'Date',
  executorEmail: 'String',
  targetUserId: 'String',
  targetEmail: 'String',
  reason: 'String',
  deleteType: 'String'
};

// Logs Sheet - システムログ
const LOG_SCHEMA = {
  timestamp: 'Date',
  userId: 'String',
  action: 'String',
  details: 'JSON String'
};
```

### 3.3 API設計（実装済み）
```javascript
// doGet - メインエントリーポイント
function doGet(e) {
  // ルーティング: mode=admin|view|setup
  // 認証処理
  // Template rendering
}

// データAPI群
const API_ENDPOINTS = {
  getUserData: '(userId) → UserInfo + ConfigJson',
  saveAndPublish: '(config) → PublishResult',
  getSystemStatus: '() → SystemHealth',
  authenticateUser: '(email) → AuthResult',
  verifyUserAccess: '(userId) → AccessResult'
};
```

## 4. 制約事項と制限

### 4.1 Google Apps Script 制限事項
```javascript
const GAS_LIMITS = {
  実行時間: '6分/実行',
  同時実行: '30並行実行',
  メモリ: '制限あり（具体値非公開）',
  API呼び出し: '6分間100,000回',
  ファイルサイズ: '50MB/スクリプト'
};
```

### 4.2 現在の運用制限
```javascript
const OPERATIONAL_LIMITS = {
  同時接続ユーザー数: '最大100ユーザー',
  Sheetsサイズ: '200万行（Google制限）',
  キャッシュTTL: '300-3600秒（用途別）',
  ポーリング間隔: '3-5秒',
  履歴保持: '50件（パフォーマンス考慮）'
};
```

### 4.3 技術的負債
```javascript
const TECHNICAL_DEBT = {
  キャッシュクリア: '30箇所以上で重複処理',
  関数依存関係: '30ファイル間の複雑な依存',
  エラーハンドリング: '統一されていない処理パターン',
  設定管理: 'ハードコードされた値が散在'
};
```

## 5. パフォーマンス仕様

### 5.1 現在の性能指標
```javascript
const PERFORMANCE_METRICS = {
  初期読み込み: '1-3秒（キャッシュ有効時）',
  ポーリング応答: '平均500ms',
  リアクション更新: '87.5%成功率',
  管理画面操作: '2-5秒',
  データ取得API: '平均1.2秒'
};
```

### 5.2 キャッシュ戦略（実装済み）
```javascript
// 3層キャッシュアーキテクチャ
const CACHE_STRATEGY = {
  L1_メモリキャッシュ: {
    対象: '頻繁アクセスデータ',
    TTL: '実行中のみ',
    サイズ: 'GAS制限内'
  },
  L2_スクリプトキャッシュ: {
    対象: 'ユーザー共通データ',
    TTL: '300-3600秒',
    サイズ: '100KB'
  },
  L3_ユーザーキャッシュ: {
    対象: 'ユーザー固有データ',
    TTL: '1800秒',
    サイズ: '100KB'
  }
};
```

## 6. セキュリティ要件（実装済み）

### 6.1 認証・認可
```javascript
const SECURITY_IMPLEMENTATION = {
  認証: 'Google OAuth 2.0',
  認可: 'ドメインベース + ロールベース',
  セッション: 'Google Apps Script Session API',
  暗号化: 'HTTPS強制（Google強制）'
};
```

### 6.2 入力検証・出力エスケープ
```javascript
// 実装済みセキュリティ機能
function sanitizeInput(input) {
  return String(input)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#x27;');
}
```

### 6.3 監査ログ
```javascript
// DeleteLogs Sheet による完全な監査証跡
const AUDIT_LOGGING = {
  削除操作: 'DeleteLogs Sheetに記録',
  アクセス: 'Logs Sheetに記録',
  管理操作: '詳細ログとタイムスタンプ',
  エラー: 'エラーレベル別分類・記録'
};
```

## 7. 今後の開発方針

### 7.1 高優先度（技術的負債解消）
```javascript
const HIGH_PRIORITY = {
  キャッシュ統一化: {
    現状: '30箇所以上の重複処理',
    目標: '5箇所への集約',
    効果: 'パフォーマンス向上・メンテナンス性向上'
  },
  エラーハンドリング統一: {
    現状: '異なるパターンが混在',
    目標: '統一されたエラー処理クラス',
    効果: 'デバッグ効率・ユーザー体験向上'
  },
  関数依存関係整理: {
    現状: '複雑な相互依存',
    目標: '明確な依存関係ツリー',
    効果: '初期化エラーの根絶'
  }
};
```

### 7.2 中優先度（機能拡張）
```javascript
const MEDIUM_PRIORITY = {
  検索・フィルタリング強化: {
    現状: '基本的な表示のみ',
    目標: '高度な検索・ソート機能',
    技術: 'クライアントサイド実装'
  },
  アナリティクス機能: {
    現状: '基本統計のみ',
    目標: 'リアルタイム分析ダッシュボード',
    技術: 'Google Analytics連携'
  },
  通知システム: {
    現状: '未実装',
    目標: 'リアルタイム通知',
    技術: 'WebSocket or ポーリング拡張'
  }
};
```

### 7.3 低優先度（体験向上）
```javascript
const LOW_PRIORITY = {
  UI_UX改善: {
    アニメーション: 'マイクロインタラクション',
    レスポンシブ: 'タブレット最適化',
    アクセシビリティ: 'WCAG 2.1 AA準拠'
  },
  国際化: {
    多言語対応: '日本語・英語',
    ロケール: 'タイムゾーン・日付形式'
  },
  PWA化: {
    オフライン対応: 'Service Worker',
    インストール: 'Add to Home Screen'
  }
};
```

## 8. 運用・保守要件

### 8.1 現在の運用体制
```javascript
const OPERATIONS = {
  監視: 'GAS実行ログ + カスタムログ',
  バックアップ: 'Google Drive自動バックアップ',
  復旧: 'バージョン管理による巻き戻し',
  更新: 'clasp deploy による自動デプロイ'
};
```

### 8.2 保守性要件
```javascript
const MAINTAINABILITY = {
  コード品質: {
    現状: '関数分離・モジュール化済み',
    改善: 'ESLint導入・自動フォーマット'
  },
  ドキュメント: {
    現状: 'インラインコメント充実',
    改善: 'API仕様書自動生成'
  },
  テスト: {
    現状: '手動テスト中心',
    改善: '自動テストスイート構築'
  }
};
```

## 9. スケーラビリティ戦略

### 9.1 短期的対応（現在のGAS制限内）
```javascript
const SHORT_TERM_SCALING = {
  データベース最適化: {
    インデックス戦略: 'userId基準の効率的検索',
    バッチ処理: '大量データの分割処理',
    キャッシュ最適化: 'より長いTTLでの運用'
  },
  パフォーマンス改善: {
    ポーリング最適化: '差分更新による負荷軽減',
    画像最適化: 'WebP対応・遅延読み込み',
    コード分割: '必要な機能のみ読み込み'
  }
};
```

### 9.2 長期的対応（アーキテクチャ進化）
```javascript
const LONG_TERM_SCALING = {
  マイクロサービス化: {
    認証サービス: '独立したGASプロジェクト',
    データサービス: 'CRUD操作専用サービス',
    通知サービス: 'リアルタイム通信専用'
  },
  外部サービス連携: {
    データベース: 'Cloud SQL or Firebase',
    ストレージ: 'Cloud Storage',
    キャッシュ: 'Cloud Memorystore'
  }
};
```

## 10. 品質保証戦略

### 10.1 現在の品質指標
```javascript
const QUALITY_METRICS = {
  機能動作率: '87.5%（リアクション機能）',
  エラー率: '5%未満（ログ分析結果）',
  レスポンス時間: '平均1.2秒',
  可用性: '99.5%（Google SLA依存）'
};
```

### 10.2 品質向上計画
```javascript
const QUALITY_IMPROVEMENT = {
  自動テスト: {
    単体テスト: 'Jest + GAS Testing Framework',
    統合テスト: 'E2E シナリオテスト',
    負荷テスト: '100ユーザー同時接続テスト'
  },
  監視強化: {
    リアルタイム監視: 'パフォーマンスメトリクス',
    アラート: 'エラー率・レスポンス時間閾値',
    ダッシュボード: 'システム健康状態可視化'
  }
};
```

## 11. 結論

Everyone's Answer Board は **教育現場での実用に耐える高品質なシステム** として実装・運用されています。Google Apps Script の制限を巧みに回避しながら、以下の特徴を実現しています：

### 11.1 システムの強み
- **高い安定性**: Google Cloud Platform の信頼性
- **優れたUX**: レスポンシブデザインと直感的操作
- **運用の簡単さ**: Google Workspace統合による管理負荷軽減
- **セキュリティ**: Google のセキュリティ基盤活用

### 11.2 今後の発展
技術的負債の解消を優先しつつ、段階的な機能拡張により、より多くの教育現場での活用を目指します。現在の堅実な基盤の上に、スケーラブルで保守性の高いシステムへと進化させていきます。

この要件定義書は、実装済みの機能を正確に反映し、今後の開発方向性を明確にすることで、継続的なシステム改善の指針となります。