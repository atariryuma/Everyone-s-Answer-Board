# コーディング方針・ガイドライン

## 1. 開発方針の概要

### 1.1 基本原則
**Everyone's Answer Board** の今後の開発は、以下の原則に基づいて行います：

1. **既存システムの安定性を最優先** - 運用中システムへの影響を最小化
2. **段階的改善** - 大規模なリファクタリングより小さな改善の積み重ね
3. **技術的負債の計画的解消** - パフォーマンスとメンテナンス性の向上
4. **Google Apps Script制限の遵守** - プラットフォーム制約内での最適化

### 1.2 優先順位
```javascript
const DEVELOPMENT_PRIORITY = {
  P0_緊急: '本番環境の障害・セキュリティ問題',
  P1_高: '技術的負債解消・パフォーマンス改善',
  P2_中: '機能拡張・UX改善',
  P3_低: '新機能追加・実験的機能'
};
```

## 2. 技術的負債解消計画

### 2.1 キャッシュシステム統一化 (P1)
```javascript
// 現状の問題
const CACHE_ISSUES = {
  重複処理: '30箇所以上でclearExecutionUserInfoCache()呼び出し',
  不整合: '異なるTTL設定（300秒、1800秒、3600秒）',
  効率性: 'DB再クエリの頻発'
};

// 解決策
const CACHE_SOLUTION = {
  統一クラス: 'UnifiedCacheManager の全面採用',
  集約化: '5箇所への集約（Core, API, UI, Events, Forms）',
  標準化: '用途別TTL設定の標準化'
};

// 実装手順
const CACHE_REFACTORING_STEPS = [
  '1. 既存キャッシュ呼び出し箇所の調査・マッピング',
  '2. UnifiedCacheManager の機能拡張',
  '3. 段階的な置き換え（1ファイルずつ）',
  '4. テスト・検証・性能測定',
  '5. 不要なキャッシュ関数の削除'
];
```

### 2.2 エラーハンドリング統一化 (P1)
```javascript
// 現状の問題
const ERROR_HANDLING_ISSUES = {
  パターン不統一: '各モジュールで異なるエラー処理',
  ログ分散: '複数箇所での独自ログ実装',
  ユーザー体験: '一貫性のないエラーメッセージ'
};

// 標準エラーハンドリングクラス
class StandardErrorHandler {
  static handle(error, context, severity = 'ERROR') {
    // 1. エラー分類
    const category = this.categorizeError(error);
    
    // 2. ログ出力
    this.logError(error, context, category, severity);
    
    // 3. ユーザー通知
    this.notifyUser(error, category);
    
    // 4. 復旧処理
    this.attemptRecovery(error, context);
  }
  
  static categorizeError(error) {
    const patterns = {
      AUTHENTICATION: /authentication|login|oauth/i,
      AUTHORIZATION: /permission|access|forbidden/i,
      DATA: /sheet|database|not found/i,
      NETWORK: /timeout|network|connection/i,
      SYSTEM: /memory|execution|limit/i
    };
    
    for (const [category, pattern] of Object.entries(patterns)) {
      if (pattern.test(error.message)) return category;
    }
    return 'UNKNOWN';
  }
}
```

### 2.3 関数依存関係の整理 (P1)
```javascript
// 安全な関数呼び出しパターンの全面採用
const SAFE_CALLING_PATTERN = {
  原則: 'すべての関数間呼び出しでsafeCall()を使用',
  実装: '既存の直接呼び出しを段階的に置き換え',
  監視: '未定義関数エラーの完全撲滅'
};

// 推奨パターン
function recommendedPattern() {
  // ❌ 従来のパターン
  // updateUIWithNewStatus(status);
  
  // ✅ 推奨パターン
  safeCall('updateUIWithNewStatus', status);
  
  // ✅ 非同期待機パターン
  waitForFunction('updateUIWithNewStatus')
    .then(fn => fn(status))
    .catch(error => console.error('Function not available:', error));
}
```

## 3. パフォーマンス最適化方針

### 3.1 データベースアクセス最適化
```javascript
const DB_OPTIMIZATION = {
  バッチ化: {
    現状: '個別クエリの頻発',
    改善: 'バッチ操作での効率化',
    実装: 'BatchOperationManager クラス'
  },
  インデックス戦略: {
    主キー: 'userId での高速検索',
    複合: 'userId + lastAccessedAt での履歴検索',
    キャッシュ: '検索結果の適切なキャッシュ'
  },
  データ削減: {
    レスポンス: '必要なフィールドのみ取得',
    履歴: '50件制限の効率的実装',
    アーカイブ: '古いデータの自動アーカイブ'
  }
};

// BatchOperationManager の実装例
class BatchOperationManager {
  constructor() {
    this.pendingOperations = [];
    this.batchTimer = null;
  }
  
  addOperation(operation) {
    this.pendingOperations.push(operation);
    this.scheduleBatch();
  }
  
  scheduleBatch() {
    if (this.batchTimer) return;
    
    this.batchTimer = setTimeout(() => {
      this.executeBatch();
      this.batchTimer = null;
    }, 100); // 100ms後にバッチ実行
  }
  
  executeBatch() {
    if (this.pendingOperations.length === 0) return;
    
    const operations = this.pendingOperations.splice(0);
    this.processBatchOperations(operations);
  }
}
```

### 3.2 フロントエンド最適化
```javascript
const FRONTEND_OPTIMIZATION = {
  レンダリング最適化: {
    仮想スクロール: '大量データの効率的表示',
    遅延読み込み: '画面外要素の遅延レンダリング',
    メモ化: '重い計算処理のキャッシュ'
  },
  ポーリング最適化: {
    差分更新: '変更分のみの取得・反映',
    adaptive間隔: 'アクティビティに応じた間隔調整',
    バックグラウンド制御: 'タブ非表示時の制御'
  },
  リソース最適化: {
    CSS最小化: '不要なスタイルの削除',
    JavaScript分割: '機能別ファイル分割',
    画像最適化: 'WebP対応・適切なサイズ'
  }
};
```

## 4. 新機能開発ガイドライン

### 4.1 機能追加の基本方針
```javascript
const FEATURE_DEVELOPMENT = {
  段階的実装: {
    フェーズ1: 'MVP（最小限の機能）',
    フェーズ2: '基本機能の完成',
    フェーズ3: '高度な機能・最適化'
  },
  後方互換性: {
    原則: '既存機能への影響を最小化',
    テスト: '既存テストケースの継続実行',
    フォールバック: '新機能無効化での動作保証'
  },
  設定可能性: {
    フィーチャーフラグ: '機能のON/OFF制御',
    グラデーション: '段階的なロールアウト',
    ロールバック: '問題発生時の即座な無効化'
  }
};
```

### 4.2 検索・フィルタリング機能の実装方針
```javascript
// 中優先度機能の実装例
const SEARCH_IMPLEMENTATION = {
  技術選択: {
    実装場所: 'クライアントサイド（GAS制限回避）',
    ライブラリ: 'Fuse.js（軽量ファジー検索）',
    インデックス: 'ローカルインデックス構築'
  },
  段階的実装: {
    Phase1: '基本的なテキスト検索',
    Phase2: 'タグ・カテゴリフィルタ',
    Phase3: '高度な検索（日付範囲、複合条件）'
  },
  パフォーマンス: {
    キャッシュ: '検索結果のローカルキャッシュ',
    デバウンス: '検索クエリの最適化',
    仮想化: '大量結果の効率的表示'
  }
};

// 実装例
class SearchManager {
  constructor() {
    this.searchIndex = null;
    this.cache = new Map();
  }
  
  async buildIndex(data) {
    // Fuse.js を使用したインデックス構築
    const options = {
      keys: ['title', 'content', 'author'],
      threshold: 0.3
    };
    this.searchIndex = new Fuse(data, options);
  }
  
  search(query) {
    if (this.cache.has(query)) {
      return this.cache.get(query);
    }
    
    const results = this.searchIndex.search(query);
    this.cache.set(query, results);
    return results;
  }
}
```

## 5. 品質保証・テスト戦略

### 5.1 テスト実装方針
```javascript
const TESTING_STRATEGY = {
  単体テスト: {
    フレームワーク: 'Jest + GAS Testing Framework',
    カバレッジ: '新規コード90%以上',
    重点: 'ビジネスロジック・ユーティリティ関数'
  },
  統合テスト: {
    スコープ: 'API エンドポイント・データフロー',
    自動化: 'CI/CD パイプラインでの実行',
    環境: '専用テスト環境での実行'
  },
  E2Eテスト: {
    ツール: 'Playwright + GAS環境',
    シナリオ: '主要ユーザーフローの検証',
    頻度: 'リリース前・週次実行'
  }
};

// テスト実装例
describe('CacheManager', () => {
  beforeEach(() => {
    // テスト環境セットアップ
    global.CacheService = mockCacheService;
  });
  
  test('should cache and retrieve data correctly', () => {
    const manager = new UnifiedCacheManager();
    const testData = { test: 'value' };
    
    manager.put('test-key', testData, 300);
    const retrieved = manager.get('test-key');
    
    expect(retrieved).toEqual(testData);
  });
  
  test('should handle cache expiration', async () => {
    const manager = new UnifiedCacheManager();
    manager.put('test-key', 'value', 0.1); // 0.1秒TTL
    
    await new Promise(resolve => setTimeout(resolve, 200));
    const retrieved = manager.get('test-key');
    
    expect(retrieved).toBeNull();
  });
});
```

### 5.2 コード品質管理
```javascript
const CODE_QUALITY = {
  静的解析: {
    ESLint: 'Google JavaScript Style Guide準拠',
    JSDoc: '関数・クラスの完全なドキュメント',
    TypeScript: '段階的な型導入（.d.tsファイル）'
  },
  コードレビュー: {
    必須項目: 'パフォーマンス・セキュリティ・保守性',
    チェックリスト: '技術的負債の増加防止',
    自動化: 'プルリクエストでの自動チェック'
  },
  リファクタリング: {
    継続的改善: '月次リファクタリング計画',
    メトリクス監視: '複雑度・重複率の追跡',
    技術的負債: 'Technical Debt Ratioの管理'
  }
};
```

## 6. セキュリティ・コンプライアンス

### 6.1 セキュリティ開発方針
```javascript
const SECURITY_GUIDELINES = {
  入力検証: {
    原則: 'すべての外部入力の検証・サニタイズ',
    実装: 'SecurityValidator クラスの使用',
    チェック: 'XSS・インジェクション対策'
  },
  認証・認可: {
    強化: '多要素認証の検討',
    監査: 'アクセスログの詳細記録',
    権限: '最小権限の原則'
  },
  データ保護: {
    暗号化: '機密データの暗号化（該当時）',
    匿名化: '個人情報の適切な取り扱い',
    削除: 'データ保持期間の遵守'
  }
};

// セキュリティ検証クラス
class SecurityValidator {
  static validateInput(input, type) {
    switch (type) {
      case 'email':
        return this.isValidEmail(input);
      case 'text':
        return this.sanitizeText(input);
      case 'json':
        return this.validateJSON(input);
      default:
        return this.genericSanitize(input);
    }
  }
  
  static sanitizeText(text) {
    return String(text)
      .replace(/[<>\"'&]/g, char => {
        const entities = {
          '<': '&lt;',
          '>': '&gt;',
          '"': '&quot;',
          "'": '&#x27;',
          '&': '&amp;'
        };
        return entities[char];
      });
  }
}
```

## 7. 運用・監視強化

### 7.1 モニタリング実装
```javascript
const MONITORING_ENHANCEMENT = {
  メトリクス収集: {
    パフォーマンス: 'レスポンス時間・スループット',
    エラー率: 'エラー種別・頻度・影響範囲',
    利用状況: 'アクティブユーザー・機能利用率'
  },
  アラート設定: {
    閾値: 'レスポンス時間3秒超過・エラー率5%超過',
    通知: 'メール・Slack・管理画面通知',
    自動対応: '可能な範囲での自動復旧'
  },
  ダッシュボード: {
    リアルタイム: 'システム健康状態の可視化',
    履歴: 'トレンド分析・パフォーマンス推移',
    予測: '容量計画・問題予測'
  }
};

// モニタリングクラス
class SystemMonitor {
  static collectMetrics() {
    return {
      timestamp: new Date().toISOString(),
      performance: {
        responseTime: this.measureResponseTime(),
        memoryUsage: this.getMemoryUsage(),
        cacheHitRate: this.getCacheHitRate()
      },
      errors: {
        rate: this.getErrorRate(),
        types: this.getErrorTypes(),
        criticalCount: this.getCriticalErrorCount()
      },
      usage: {
        activeUsers: this.getActiveUserCount(),
        requestsPerMinute: this.getRequestRate(),
        popularFeatures: this.getFeatureUsage()
      }
    };
  }
}
```

## 8. 開発プロセス・ワークフロー

### 8.1 ブランチ戦略
```
Main Branch Strategy:
main
├── develop (開発ブランチ)
├── feature/cache-optimization (機能ブランチ)
├── hotfix/critical-bug-fix (緊急修正)
└── release/v2.1.0 (リリースブランチ)
```

### 8.2 リリースプロセス
```javascript
const RELEASE_PROCESS = {
  段階: [
    '1. 機能実装・単体テスト',
    '2. コードレビュー・統合テスト',
    '3. ステージング環境デプロイ・E2Eテスト',
    '4. 本番環境デプロイ・監視',
    '5. リリース後検証・ロールバック準備'
  ],
  自動化: {
    CI: 'GitHub Actions でのテスト自動実行',
    CD: 'clasp deploy での自動デプロイ',
    通知: 'リリース状況の自動通知'
  },
  品質ゲート: {
    テストカバレッジ: '90%以上',
    パフォーマンス: 'ベースライン以下',
    セキュリティ: '脆弱性ゼロ'
  }
};
```

## 9. 技術的意思決定フレームワーク

### 9.1 技術選択の基準
```javascript
const DECISION_CRITERIA = {
  技術的適合性: {
    GAS互換性: 'Google Apps Script での動作保証',
    制約遵守: 'プラットフォーム制限の考慮',
    パフォーマンス: '現在の性能基準維持・向上'
  },
  運用面: {
    保守性: 'メンテナンスコストの考慮',
    スケーラビリティ: '将来の拡張性',
    セキュリティ: 'セキュリティリスクの評価'
  },
  ビジネス面: {
    開発効率: '実装・テスト・デプロイ効率',
    学習コスト: 'チームの技術習得コスト',
    コミュニティ: 'サポート・ドキュメントの充実'
  }
};
```

### 9.2 アーキテクチャ決定記録 (ADR)
```markdown
# ADR-001: キャッシュシステム統一化

## 状況
現在、30箇所以上でキャッシュクリア処理が重複し、パフォーマンス問題を引き起こしている。

## 決定
UnifiedCacheManager を中心とした統一キャッシュシステムを採用する。

## 根拠
- パフォーマンス向上: DB再クエリの削減
- 保守性向上: 管理箇所の集約
- 一貫性保証: 統一されたキャッシュ戦略

## 結果
5箇所への集約により、メンテナンス性とパフォーマンスが向上。
```

## 10. 結論

このコーディングガイドラインは、Everyone's Answer Board の継続的な改善と発展を支える技術的基盤となります。

### 重要な原則
1. **安定性第一**: 運用中システムの安定性を最優先
2. **段階的改善**: 小さな改善の積み重ねによる進歩
3. **技術的負債の計画的解消**: パフォーマンスとメンテナンス性の向上
4. **品質の継続的向上**: テスト・監視・セキュリティの強化

このガイドラインに従うことで、堅実で持続可能なシステム発展を実現し、教育現場でのより良いツールとして成長させていきます。