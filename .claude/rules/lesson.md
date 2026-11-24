---
paths:
  - "src/LessonService.js"
  - "src/LessonWorkspace.html"
  - "src/DataApis.js"
  - "tests/lesson.*"
  - "tests/page.lesson*"
  - "tests/admin.lesson*"
  - "tests/data.apis.lessonMask*"
---

# 授業モード (native 入力) — 掲示板モードとは独立した縦の経路

テンプレ `dialogue-reconsider-5phase`（考える→出会う→議論する→もう一度考える→ふりかえる）は
`lessonJson.inputMode='native'` で動く。**Form も列推定も通らない**。掲示板モード
（Form + リアクション + StudyQuest 連携）は無変更で並走する。混ぜないこと。

- **フェーズが児童画面の権能を決める**。`phase.screenRole` = `input`/`browse`/`discuss`/`reinput`/`reflect`。
  「議論中は投稿できない」は UI ではなく `submitLessonAnswer` のサーバ検証で保証する。
- **「考える」では他者の行を返さない**（`__maskOthersDuringInputPhase_`）。viewer cache は
  全件を保持し、mask は cache 取得後に viewer ごとに適用する（順序を逆にすると他人の分が cache される）。
- **児童同士に名前を見せない**（native phase は `showNames:false` が既定）。教師は `isOwnBoard` で見える。
- **記録する名前はサーバが決める**（`__resolveLessonSubmitName_`）。名簿は無いので、その授業の回答シートにある
  「メール → 名前」を使う。自分のメールに名前があればそれ（端末の値は使わない）、無ければ端末の名前。ただし
  ほかのメールの名前と同じなら `NAME_TAKEN` で断る。端末が送る `name` をそのままシートに書かない。
- **`answer` と `reason` を同じ列にしない**。`validateMapping` は `numericX/Y` だけ重複を免除する。
  違反すると config 保存が落ち、教師が「開始」を押した瞬間に失敗する（`__nativeColumnMapping_`）。
- 投稿を受け付けないフェーズは**直前の入力フェーズのシート**を見る（自分のシートは空）。
  軸ラベルは授業を通して 1 つの座標系に揃える（● → ★ の比較が成立する条件）。
