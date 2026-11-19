# 02_design: 仕様の検討

このフォルダは **検討中** の資料を置く場所です。ここに書いてあることは、まだ確定した仕様ではありません。

実装で参照する確定した仕様は、リポジトリ直下の [specs/](../../specs/) にあります。

## 中身

| フォルダ | 置くもの | 書き方 |
|---|---|---|
| [discussions/](discussions/) | 論点ごとの検討メモ。案の比較、未決事項、調べたこと | 自由。[テンプレート](discussions/_template.md)あり |
| [decisions/](decisions/) | 決めたことの記録（ADR）。背景、決定、却下した案と理由 | [テンプレート](decisions/_template.md)に沿って短く |
| [backlog.md](backlog.md) | 今は入れないが、将来入れたい機能の候補 | 表に 1 行ずつ足す。書き方はファイルの中にあります |

ファイル名は次のようにします。

- 検討メモ: `YYYY-MM-DD_論点.md`（例: `2026-11-19_log-format.md`）
- 決定の記録: `ADR-NNN_論点.md`（例: `ADR-001_log-format.md`）。番号は 001 から順に振ります

## 仕様が確定するまでの流れ

1. **検討する** — `discussions/` に検討メモを書きます。
2. **決めたことを残す** — 結論が出たら `decisions/` に ADR を書き、元の検討メモのステータスを「決定済み」にして ADR へリンクします。
3. **仕様にする** — `/speckit-specify` を実行して `specs/NNN-機能名/spec.md` を作ります。入力には ADR を渡します。曖昧な点は `/speckit-clarify` で解消します。
4. **確定させる** — `spec.md` の `Status` を `Draft` から `Approved` に変え、プルリクエストで `develop` にマージします。**`develop` に入った時点で確定** です。
5. **実装する** — `/speckit-plan` → `/speckit-tasks` → `/speckit-analyze` → `/speckit-implement` の順に進めます。

確定した仕様を変えたいときも、まずここで検討し、ADR を書いてから `spec.md` を直します。

## 注意

- ここにある資料を、仕様として実装に使わないでください。実装の根拠は `specs/` と [constitution.md](../../.specify/memory/constitution.md) だけです。
- 決まらなかった案や却下した案も消さずに残します。「なぜそうしなかったか」を後から確かめるためです。
