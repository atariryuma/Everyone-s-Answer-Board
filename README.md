# Everyone's Answer Board

> Google Apps Script (GAS) ベースの協働回答ボードシステム

## 🚀 Quick Start

### 1. セットアップ

```bash
# リポジトリをクローン
git clone https://github.com/atariryuma/Everyone-s-Answer-Board.git
cd Everyone-s-Answer-Board

# 依存関係をインストール
npm install

# .clasp.json を作成（テンプレートからコピー）
cp .clasp.json.template .clasp.json
# .clasp.json を編集して scriptId を設定

# clasp にログイン
npx clasp login
```

### 2. 開発ワークフロー

```bash
# GAS からコードを取得
npm run pull

# GAS にコードをプッシュ
npm run push

# GAS エディタを開く
npm run open

# 実行ログを確認
npm run logs
```

## 📁 プロジェクト構成

```
Everyone-s-Answer-Board/
├── src/                        # GAS ソースコード（.js + .html）
│   ├── main.js                # エントリーポイント
│   ├── helpers.js             # ユーティリティ関数
│   ├── DatabaseCore.js        # データベース操作
│   └── *.html                 # フロントエンドテンプレート
├── .clasp.json.template       # clasp 設定テンプレート
├── .claspignore               # clasp push 除外設定
├── .gitignore                 # Git 除外設定
├── package.json               # npm 設定（clasp のみ）
├── CLAUDE.md                  # 開発ガイド（Claude Code 用）
└── README.md                  # このファイル
```

## 🔧 技術スタック

- **Google Apps Script (GAS)**: サーバーサイド
- **clasp**: GAS CLI ツール
- **Git/GitHub**: バージョン管理
- **Zero Dependencies**: 外部ライブラリ不使用

## 📝 開発ルール

### ファイル拡張子

- ✅ `.js` を使用（clasp のデフォルト）
- GAS エディタでは `.gs` として表示されます

### Git 管理

**コミットするファイル:**
- `src/**/*.js` - ソースコード
- `src/**/*.html` - HTMLテンプレート
- `.clasp.json.template` - チーム共有用設定テンプレート
- `.claspignore`, `.gitignore` - 除外設定
- `package.json` - npm 設定
- `CLAUDE.md`, `README.md` - ドキュメント

**コミットしないファイル (.gitignore):**
- `.clasp.json` - 個人の scriptId を含む
- `node_modules/` - npm パッケージ
- `.DS_Store` - OS 生成ファイル

## 🏗️ アーキテクチャ

詳細は [CLAUDE.md](CLAUDE.md) を参照してください。

- **Zero-Dependency**: 外部ライブラリ不使用
- **Direct GAS API**: SpreadsheetApp, PropertiesService など直接使用
- **Batch Operations**: 70倍の性能改善
- **TTL Caching**: PropertiesService API 呼び出しを 80-90% 削減

## 📚 ドキュメント

- [CLAUDE.md](CLAUDE.md) - 開発ガイド（Claude Code 用）
- [clasp 公式ドキュメント](https://github.com/google/clasp)

## 📄 ライセンス

ISC License

## 👤 Author

Ryuma Atari
