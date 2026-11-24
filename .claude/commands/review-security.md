# 回答ボードのセキュリティの確かめ

一般的な監査は組み込みの `/security-review` に任せる。ここでは**この回答ボードで実際に事故になった・なりかけた所だけ**を、
今の差分 (指定が無ければ `git diff main...HEAD` と作業ツリー) について確かめる。

前提: 先生として実行 (`USER_DEPLOYING`) なので、**児童の画面から呼べる処理は、すべて先生の権限で動く**。

## 確かめること

1. **公開関数** — 名前が `_` で終わらない関数を足した・名前を変えたら、`scripts/public-functions.json` に載せてよいものか。
   画面から呼ぶ入口か、児童が任意の引数で呼んでも書き込み・秘密の読み出し・重い処理が起きないこと。`npm run check:gas` (項目 15) を通す。
2. **先生だけの処理** — 管理パネルから呼ぶ処理は、入口で `teacherOnlyError_()` を確かめているか。
3. **見せてよいシートか** — スプレッドシートを開く所は `checkSpreadsheetAccess_` を通っているか。公開の状態を変えたら
   `invalidateAccessDecisionCache_` で判定を古くしているか。
4. **シートに書く児童の文** — `__plainTextCell_` を通しているか (`=` で始まる文が数式になり、他の児童のメールを外へ送れた)。
5. **授業モードで他の児童を見せない** — 「考える」では `__maskOthersDuringInputPhase_` が **cache を取った後に** viewer ごとにかかるか。
   児童同士に名前を出していないか。記録する名前は `__resolveLessonSubmitName_` が決め、端末の `name` をそのまま書いていないか。
6. **利用者の確かめ** — `Session.getActiveUser()` だけを使う。`getEffectiveUser()` を使わない。
7. **入力と HTML** — doPost の action を足したら同時に入力の検証を足したか (`validators.js`)。児童の文は `textContent`、
   `innerHTML` なら `escapeHtml()`。
8. **秘密** — `scripts/single-tenants.json` / `.clasp*.json` / API キー・先生のメールをコミット・ログ・応答に出していないか。

## 出し方

項目ごとに「問題なし / 要修正 (ファイル:行と、児童の画面から何ができてしまうか)」を短く書く。点数は付けない。
直すときは、テストで「児童として呼ぶと断られる」ことを確かめる。
