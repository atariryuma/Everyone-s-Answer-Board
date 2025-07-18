# Everyone's Answer Board (みんなの回答ボード) - StudyQuest

## 概要

**Everyone's Answer Board** は、Google Apps Script (GAS) を利用して構築された、リアルタイムでインタラクティブな回答共有ウェブアプリケーションです。教育現場やワークショップなどで、参加者からの意見や回答を即座に集約し、全員で閲覧・共有することを目的としています。

Googleスプレッドシートをデータベースとして活用し、Googleフォームと連携することで、簡単に質問を投げかけ、回答をボード形式で表示することができます。

## 主な機能

*   **Googleアカウント認証:** 安全なGoogleアカウントでのログイン。
*   **ボード作成:**
    *   新規Googleフォームと連携したボードの自動生成。
    *   既存のGoogleスプレッドシートを回答データベースとして利用。
*   **リアルタイム表示:** 投稿された回答がリアルタイムでボードに反映されます。
*   **インタラクション:**
    *   各回答カードに対するリアクション機能。
    *   管理者による特定の回答のハイライト機能。
*   **柔軟なカスタマイズ:**
    *   ボードに表示する列（項目）を自由に選択・マッピング。
    *   記名/匿名表示の切り替え。
    *   リアクション数の表示/非表示設定。

## アーキテクチャ

*   **バックエンド:** Google Apps Script (V8 Runtime)
*   **フロントエンド:** HTML, CSS, JavaScript (外部ライブラリ不使用)
*   **データベース:** Google Sheets
*   **認証:** Google OAuth2
*   **開発ツール:**
    *   `clasp`: Google Apps Scriptプロジェクトのローカル開発用CLIツール。
    *   `jest`: JavaScriptコードの単体テスト用フレームワーク。

## セットアップ

### 1. 前提条件

*   [Node.js](https://nodejs.org/) と `npm` がインストールされていること。
*   Googleアカウントを持っていること。
*   `clasp` がインストールされ、Googleアカウントでログイン済みであること。
    ```bash
    npm install -g @google/clasp
    clasp login
    ```

### 2. プロジェクトのクローンと初期設定

```bash
# リポジトリをクローン
git clone https://github.com/atariryuma/Everyone-s-Answer-Board.git

# プロジェクトディレクトリに移動
cd Everyone-s-Answer-Board

# 開発依存関係をインストール
npm install
```

### 3. Google Apps Script プロジェクトの作成

`clasp` を使用して、新規または既存のGASプロジェクトに接続します。

```bash
# 新規プロジェクトを作成する場合
clasp create --title "みんなの回答ボード" --rootDir ./src

# 既存のプロジェクトに接続する場合
clasp clone <scriptId> --rootDir ./src
```

### 4. 環境設定

本アプリケーションを動作させるには、いくつかの認証情報と設定をスクリプトプロパティに保存する必要があります。
詳細は以下のガイドを参照してください。

**➡️ [SETUP_GUIDE.md](./SETUP_GUIDE.md)**

## 開発

### コードのデプロイ

ローカルでの変更をGoogle Apps Scriptプロジェクトに反映させるには、以下のコマンドを実行します。

```bash
npm run push
# または
clasp push
```

### テスト

`jest` を使用した単体テストが用意されています。

```bash
npm test
```

## ライセンス

このプロジェクトは [ISC License](./LICENSE) の下で公開されています。
