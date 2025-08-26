# GAS V8/ES6 コーディング方針（AI向け）

> 目的：Google Apps Script（GAS）の **V8 ランタイム** 前提で、AIが安定したコードを生成できるようにする実務ガイド。  
> 対象：コード生成AI（Claude/ChatGPT等）、レビュー用 LLM、社内自動化ボット。

---

## 0) ランタイム前提

- GAS は **V8**（Chrome/Node と同系エンジン）で動作し、**モダンな ECMAScript 構文**を利用可能。  
- `let/const`、アロー関数、テンプレートリテラル、分割代入、クラス、`Map/Set` などが使える。  
- 旧 Rhino も選べるが、**V8 を強く推奨**。

---

## 1) まず守るコーディング規範

- **`var` を禁止**。**`const` 優先**、必要時のみ `let`。  
- **関数は原則アロー関数**（`this` を必要とするクラスメソッドは通常のメソッド記法）。  
- **テンプレートリテラル**で文字列結合を可読化。  
- **分割代入 / スプレッド**で引数・配列・オブジェクト操作を簡潔に。  
- **不変データ**志向：オブジェクトの直接破壊より新オブジェクトを返す。

---

## 2) GAS ならではの設計ルール

- **エントリーポイントはグローバル関数**（トリガーやメニュー登録はトップレベル）。  
- **Apps Script のサービス API は同期的**。`UrlFetchApp` などはブロッキング呼び出し。  
- **バッチ処理**：Spreadsheet/Drive などはまとめて取得・更新。  
- **状態は PropertiesService/CacheService** に格納。  
- **ログ**は `console.log` または `Logger.log`。  
- **例外設計**：`throw new Error()` を用い、`try/catch` でハンドリング。

---

## 3) ファイル構成とモジュール化

- GAS は **ES Modules 非対応**。`import/export` はそのままでは使えない。  
- 多ファイルに分け、**グローバル名前空間を汚さない命名**で整理。  
- npm を利用する場合はローカルでバンドル（Webpack など）、`clasp` でアップロード。

```

/src               # ES6/TS 源コード
/dist/code.js      # バンドル出力（GAS 用）
/appsscript.json   # マニフェスト

````

---

## 4) 主要 ES6+ 機能の使い分け

- **クラス**：責務分離に利用。  
- **`Map/Set`**：高速検索や非文字列キー管理に有効。  
- **イテレータ/ジェネレータ**：大規模データの段階処理に。  
- **Promise/`async`**：構文は使えるが GAS API は同期。基本は不要。

---

## 5) I/O とパフォーマンス

- **Spreadsheet**：`getValues()`→配列処理→`setValues()` 一括。  
- **Drive/UrlFetch**：必要最小限。レスポンスは即 `JSON.parse`。  
- **キャッシュ**：`CacheService`/`PropertiesService` を積極利用。  
- **Utilities.sleep** は最小限。

---

## 6) 例外・リトライ・検証

- **入力検証**：Public 関数は JSDoc で型・必須性を明記。  
- **外部呼び出し**は指数バックオフを実装。  
- **失敗時の痕跡**：`console.error`＋簡潔な要約を残す。

---

## 7) ローカル開発・デプロイ

- **`clasp`** でローカル開発／デプロイ。  
- **ライブラリ**は Script ID 指定でバージョン固定。  
- **npm** は「バンドル → 単一ファイル」の原則。

---

## 8) HTML Service/フロント連携

- クライアント JS はブラウザの ES Modules 利用可。  
- サーバ側とは別。`google.script.run` で呼び出す。

---

## 9) 生成AI向けプロンプト指示（サンプル）

1. **`const`優先、`let`のみ許可、`var`禁止**。  
2. **エントリーポイントはトップレベル関数**。  
3. **バッチI/O・最小呼び出し**を強制。  
4. **JSDoc**を必須。  
5. **例外処理・ログ方針**を明記。  
6. **`async/await`は不要**。  
7. npm が必要な場合のみバンドル指示。  
8. 依存がなければ **単一ファイルに収束**。

---

## 10) 最小テンプレート（雛形）

```js
/** @OnlyCurrentDoc */

/**
 * シートのA列を集計してログ出力する
 * @return {void}
 */
function main() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const values = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
  const nonEmpty = values.flat().filter(v => String(v).trim() !== "");
  const counts = countBy(nonEmpty);
  console.log(summaryToLines(counts).join("\n"));
}

/** @param {string[]} arr @return {Map<string, number>} */
const countBy = (arr) => {
  const m = new Map();
  for (const v of arr) m.set(v, (m.get(v) ?? 0) + 1);
  return m;
};

/** @param {Map<string, number>} m @return {string[]} */
const summaryToLines = (m) =>
  Array.from(m.entries()).map(([k, n]) => `${k}: ${n}`);
````

---

## AI にありがちな誤りと対策

* **誤**「`import/export` を .gs に書く」
  → **正**「バンドルして単一 JS にしてからデプロイ」。
* **誤**「`await UrlFetchApp.fetch()`」
  → **正**「同期API。`await` は意味がない」。
* **誤**「セルを1件ずつ `setValue`」
  → **正**「`getValues`/`setValues` で一括」。
* **誤**「`var` 多用」
  → **正**「`const`/`let`」。

---
