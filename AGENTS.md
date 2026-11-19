# AGENTS.md

このリポジトリで作業する AI エージェント向けの指示です。

## 仕様の扱い

- 確定した仕様は `specs/NNN-機能名/`（Spec Kit が作るもの）と `.specify/memory/constitution.md` にある。実装（`/speckit-plan` 以降）はこれだけを根拠にすること。
- `docs/02_design/` は検討中の資料。背景を知るために読むのはよいが、仕様として扱って実装に使わないこと。そこにしか書かれていないことが実装に必要なら、実装を止めてユーザーに確認すること。
- `spec.md` の `Status` を `Approved` に変えるのは、ユーザーが確定を指示したときだけにすること。
- 確定した `spec.md` を変える依頼を受けたら、先に `docs/02_design/decisions/` に ADR があるか確かめ、なければ書くことを提案すること。
- `specs/` 配下の機能フォルダを移動・改名しないこと（Spec Kit が `.specify/feature.json` で場所を覚えているため）。

## 権限設定の変更

- `.claude/settings.json` の権限ルール（allow / ask / deny）を追加・変更・削除したときは、同じ作業の中で [.claude/README.md](.claude/README.md) の該当する表も更新すること。
  - ルールごとに「意味・理由」を書く。
  - Bash と PowerShell の両方に同じ内容のルールを入れる場合は、README では 1 行にまとめてよい。
- settings.json にはコメントや独自のキー（`_comment` など）を書かないこと。Claude Code の設定チェックで拒否される。
