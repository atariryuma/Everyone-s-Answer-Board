# バックアップ - 分離前の統合ファイル

このディレクトリには、CSS/JSファイルが分離される前の統合されたHTMLファイルが保存されています。

## バックアップファイルの概要

これらのファイルは、コミット `8187f5b` から抽出されたものです。

### ファイル一覧

| ファイル名 | 説明 | 現在の分離先 |
|-----------|------|-------------|
| `AdminPanel.html` | 管理パネルの統合ファイル | `src/client/views/AdminPanel.html` + `src/client/styles/AdminPanel.css.html` + `src/client/scripts/AdminPanel.js.html` |
| `Page.html` | メインページの統合ファイル | `src/client/views/Page.html` + `src/client/styles/Page.css.html` + `src/client/scripts/Page.js.html` |
| `Registration.html` | 登録ページの統合ファイル | `src/client/views/Registration.html` + `src/client/styles/Registration.css.html` + `src/client/scripts/Registration.js.html` |
| `SetupPage.html` | セットアップページの統合ファイル | `src/client/views/SetupPage.html` + `src/client/styles/SetupPage.css.html` + `src/client/scripts/SetupPage.js.html` |
| `Unpublished.html` | 未公開ページの統合ファイル | `src/client/views/Unpublished.html` + `src/client/styles/Unpublished.css.html` + `src/client/scripts/Unpublished.js.html` |
| `ClientOptimizer.html` | クライアント最適化ファイル | （現在は使用されていません） |

## 分離の履歴

### コミット履歴
1. `2752fbf` - AdminPanelの分離
2. `d115033` - 他のHTMLファイルの分離

### 分離の理由
- コードの保守性向上
- CSS/JSの再利用性向上
- ファイル構造の整理
- 開発効率の向上

## 復元方法

もし分離前の統合ファイルに戻したい場合は、このバックアップファイルを参考にしてください。

```bash
# 例：Page.htmlを復元する場合
cp backup-integrated-files/Page.html src/client/views/Page.html
```

## 注意事項

- これらのファイルは参考用です
- 現在のシステムは分離後の構造で動作しています
- 復元する際は、include関数の呼び出しを適切に修正してください