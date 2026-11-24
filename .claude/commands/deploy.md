# 配置する

手順の正は `CLAUDE.md` の Dev workflow と `.claude/rules/deploy.md`。ここには写さない。順に進め、どこかで落ちたら止めて報告する。

1. `.claude/rules/deploy.md` を Read で開く。
2. **授業日の確かめ** — 差分が入口 (`doGet` / `DO_GET_HANDLERS`)・許可の判定 (`checkSpreadsheetAccess_` / `teacherOnlyError_`)・
   `src/appsscript.json` に触れていたら、本番に出す前にオーナーに「今日は授業がありますか」と聞く。授業日なら本番には出さない。
3. `npm test` と `npm run lint:errors` を通す。
4. 検証用の配置に出す: `npm run deploy:single -- deploy --name <検証用の名前> --no-ops-report`
5. `npm run verify:e2e -- <検証用の名前>` を通す。
6. 本番に出す: `npm run deploy -- --no-ops-report`
7. `npm run api -- systemDiagnosis` で壊れていないか確かめる。
8. 結果を報告する。コミットと push は、オーナーに聞いてから。

検証用の名前と本番の名前は `scripts/single-tenants.json` の `name` (中身の鍵・メールは画面に出さない)。
