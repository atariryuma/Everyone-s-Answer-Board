---
paths:
  - "scripts/deploy-single.js"
  - "scripts/admin-api.js"
  - "scripts/migrate-to-single.js"
  - "scripts/verify/**/*"
  - "scripts/lib/**/*"
  - "src/appsscript.json"
  - "src/Deployment.js"
  - "src/main.js"
  - "src/AdminApis.js"
  - "src/SystemController.js"
  - "src/OpsReport.js"
  - "src/Migration.js"
  - "tests/deploySingle.*"
  - "tests/deployment.*"
  - "tests/opsReport.*"
  - "tests/main.*"
  - "tests/admin.dispatch.*"
---

# 配置・Admin API・入口

## 配置の形

先生ごとに 1 つ、先生の Drive のスプレッドシートに紐づいたアプリを「先生として実行」で配置する
(`src/appsscript.json` の `webapp.executeAs = USER_DEPLOYING`)。データ (users / lessons / 授業の回答シート) は
そのスプレッドシート 1 つに入る。配置ごとに違うのは `src/Deployment.js` (先生のメールと DB の SS ID) だけで、
`scripts/deploy-single.js` が焼き込む。先生ごとの値は `scripts/single-tenants.json` (git に入れない)。
Why: 2026-10-29 から児童アカウントは「開いた人として実行」の Apps Script を一切開けない (管理側の設定)。

**エラーの報告 (ops-monitor) は 2026-11-16 にやめた。配置はいつも `--no-ops-report` を付ける。**
学校のアカウントの GAS を外のサービスにつながない、という判断 (市の決まりと、担当が替わっても回る形にするため)。
受け口の登録も消してあるので、前の配置が送っても受け取られない。`scripts/ops-report.local.json` は置かない
(フラグを付け忘れると、ファイルが無いので配置が止まる。それで気づける)。別の PC から配置するときは `single-tenants.json` だけを写す。

**検証用の配置** (名前の例 `<先生>-verify`): `extraFiles: ["scripts/verify/DevInvoke.js"]` 付き。admin API の `dev.invoke` で
画面が `google.script.run` で呼ぶ処理を名前で呼べる (児童のふりは `as`)。本番の先生には付けない。

**ほかの先生に配る (ひな形のコピー)**: `deploy-single.js template` が、`Deployment.js` を空にしたひな形を push する。
コピー先では、スプレッドシートの持ち主が管理者に、そのスプレッドシートが DB になる (`ensureCopiedDeploymentBootstrap_`)。
焼き込みが無いので、管理者と DB は CacheService に置いて読む (`copiedDeploymentProperty_`。実行ごとにプロパティを読まない)。
ひな形に鍵・報告先・先生のメールを入れない。手順と、まだ確かめていないことは [docs/DISTRIBUTION.md](../../docs/DISTRIBUTION.md)。

## 配置してよい日

- **入口 (`doGet` / `DO_GET_HANDLERS`)・許可の判定 (`checkSpreadsheetAccess_` / `teacherOnlyError_`)・
  `appsscript.json` を変える配置は、授業のある日に入れない**。入れる前に、児童のアカウントと同じ条件で開けることを
  確かめる (先生の権限の admin API や `verify:e2e` が通っただけでは足りない)。許可は開く前に判定せず GAS に任せ、
  実際に権限のエラーで落ちたときだけ知らせる。
  Why: 2026-10-22、授業の直前に入口を変えて児童 9 人が止まり、緊急に戻した。
- 児童に見える名前・ブランド・主な表示の言葉は学年の途中で変えない (変えるなら 4 月の始め)。

## Admin API (配置の状態取得・操作)

doPost の `adminApi` アクションに API キー認証でアクセス。ディスパッチャーは `src/AdminApis.js:dispatchAdminOperation()`。
`npm run api` の送り先は `single-tenants.json` の配置 (`--tenant <名前>`、省略時は検証用を除く最初の配置)。

```bash
npm run api -- systemDiagnosis                # DB接続・設定・デプロイ状態の健康診断
npm run api -- getLogs                        # 記録は無い (案内だけを返す。ops-monitor への報告は 2026-11-16 にやめた)
npm run api -- getUserConfig --userId X       # ボードの設定
npm run api -- setUserConfig --userId X --patch '{...}'  # config パッチ適用
npm run api -- previewBoard --userId X        # viewer 視点の公開ボードプレビュー
```

**いつ使うか:**

- 配置の後 → `systemDiagnosis` で壊れていないか確認
- エラー調査時 → Apps Script の「実行数」(先生が見る。ops-monitor への報告は 2026-11-16 にやめた)、`perfMetrics` でボトルネック特定 (`getLogs` は記録が無いので案内だけを返す)
- 設定変更後 → `listProperties` で反映を確認

**実行ログ**: 先生ごとの配置は既定の GCP project で動くので、Cloud Logging を CLI から読めない。
Apps Script エディタの「実行数」で見る。フロントエンドのエラーも `reportClientError` で同じ場所に出る。

**全コマンド一覧**: [docs/DEVELOPMENT.md](../../docs/DEVELOPMENT.md) §1-2。

## clasp

- 配置は `.single-stage/<名前>/` で組み立てて clasp に渡す (`-P . -I .` 必須。付けないと親の `.clasp.json` を拾う)
- File extension: `.js` locally (clasp convention), shows as `.gs` in GAS editor

### Required Script Properties

配置の初回に `ensureDeploymentBootstrap_` が `Deployment.js` の値から入れる (手で入れる必要はない):

| Property | Description |
| -------- | ----------- |
| `ADMIN_EMAIL` | この配置の先生 (管理者) |
| `DATABASE_SPREADSHEET_ID` | 先生のスプレッドシート (DB) |
| `ADMIN_API_KEY` | Admin API の鍵 (初回に `setupApiKey` で決める。`single-tenants.json` に控える) |
