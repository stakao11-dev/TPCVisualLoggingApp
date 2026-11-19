# specs: 確定した仕様

このフォルダには、Spec Kit で作った機能ごとの仕様を置きます。**実装はここに書かれた仕様に従います。**

検討中の資料は [docs/02_design/](../docs/02_design/) にあります。

## 構成

機能ごとに `NNN-機能名/` フォルダがあります（`/speckit-specify` が作ります）。

| ファイル | 中身 | 作るコマンド |
|---|---|---|
| `spec.md` | 何を作るか（要件、受け入れ条件） | `/speckit-specify`、`/speckit-clarify` |
| `plan.md` ほか（`research.md`、`data-model.md`、`contracts/` など） | どう作るか | `/speckit-plan` |
| `tasks.md` | 作業の一覧 | `/speckit-tasks` |

プロジェクト全体の原則は [constitution.md](../.specify/memory/constitution.md) にあります。

## 確定の扱い

- `spec.md` の `Status` が `Approved` で、`develop` にマージされているものが確定した仕様です。
- `Draft` のものは作成中です。機能ブランチ上で `/speckit-clarify` を使って詰めます。
- 確定した仕様を変えるときは、先に [docs/02_design/](../docs/02_design/) で検討して ADR を書き、プルリクエストで直します。
- 機能フォルダを別の場所へ移したり名前を変えたりしないでください。Spec Kit は `.specify/feature.json` でフォルダの場所を覚えているため、移すとコマンドが仕様を見つけられなくなります。
