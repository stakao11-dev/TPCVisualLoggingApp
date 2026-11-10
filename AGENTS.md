# AGENTS.md

このリポジトリで作業する AI エージェント向けの指示です。

## 権限設定の変更

- `.claude/settings.json` の権限ルール（allow / ask / deny）を追加・変更・削除したときは、同じ作業の中で [.claude/README.md](.claude/README.md) の該当する表も更新すること。
  - ルールごとに「意味・理由」を書く。
  - Bash と PowerShell の両方に同じ内容のルールを入れる場合は、README では 1 行にまとめてよい。
- settings.json にはコメントや独自のキー（`_comment` など）を書かないこと。Claude Code の設定チェックで拒否される。
