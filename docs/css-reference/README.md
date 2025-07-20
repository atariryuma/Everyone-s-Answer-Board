# CSS設計ドキュメント

## 📁 ファイル構成

### 本番環境（GAS実際読み込み）
- `src/unified-styles.css.html` - 統一CSSシステム（25KB軽量版）

### 開発リファレンス（このフォルダ）
- `core-variables.css` - デザイントークン・CSS変数定義
- `button-components.css` - ボタンコンポーネント詳細仕様
- `glass-panels.css` - グラスモーフィズム効果仕様
- `responsive-layout.css` - レスポンシブレイアウトシステム
- `styles.css` - 統合エントリーポイント（開発版）
- `loader.html` - CSSローダーシステム

## 🎯 設計方針

### ダークカラー・フラットデザイン
- **ベース色**: `#0f0f0f` (深いダーク)
- **アクセント**: `#3b82f6` (ブルー), `#10b981` (グリーン)
- **軽微なグラスモーフィズム**: `blur(4px)`, `rgba(255,255,255,0.03)`

### 統一ボタンシステム
```css
.btn--primary    /* メインアクション */
.btn--secondary  /* サブアクション */
.btn--ghost      /* 透明ボタン */
.reaction-btn    /* Page.html専用透明リアクションボタン */
```

### レスポンシブブレークポイント
- **Mobile**: 768px以下
- **Tablet**: 768px - 1024px
- **Desktop**: 1024px以上

## 🚀 パフォーマンス最適化

### GASベストプラクティス
1. **単一ファイル読み込み**: `<?!= include('unified-styles.css'); ?>`
2. **クリティカルCSS**: 各ページで3KB以内をインライン
3. **TailwindCSS除去**: CDN依存完全削除
4. **軽量化**: 80KB → 25KB (70%削減)

### 実装結果
- **Page.html**: リアクションボタン完全透明化 ✅
- **AdminPanel.html**: 統一ダークテーマ ✅
- **LoginPage.html**: エレガントなログインフォーム ✅

## 📝 今後の拡張

新しいコンポーネント追加時は：
1. このフォルダで詳細設計・テスト
2. `src/unified-styles.css.html` に軽量版を追加
3. 全ページで動作確認

---
*Generated: 2025-07-20*
*Design System: ダークカラー・フラットデザイン・軽微なグラスモーフィズム*