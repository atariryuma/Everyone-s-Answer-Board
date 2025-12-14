---
description: clasp pushとgit pushを全自動で実行するデプロイワークフロー
---

# デプロイワークフロー

このワークフローは、GAS と GitHub の両方にコードをプッシュします。

## 手順

// turbo-all

1. 変更をステージングしてコミット（未コミットの変更がある場合）
```bash
git add -A && git diff --cached --quiet || git commit -m "chore: auto-deploy"
```

2. clasp push でGASにプッシュ
```bash
clasp push
```

3. git push でGitHubにプッシュ
```bash
git push origin main
```

4. 完了メッセージを表示
```bash
echo "✅ デプロイ完了: clasp + git push 成功"
```
