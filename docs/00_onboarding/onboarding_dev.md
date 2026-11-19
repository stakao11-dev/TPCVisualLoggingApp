# TPCVisualLoggingApp オンボーディング

- 最終更新日: 2026-11-19
- この資料は Claude Code の `/onboarding-dev` skill（[SKILL.md](../../.claude/skills/onboarding-dev/SKILL.md)）で生成・更新しています。手で書き足した部分は、次に更新するときも残ります。
- 「未整備」と書いた章は、リポジトリにまだ情報がない部分です。分かる方は追記してください。

新しく参加した方は、まずこの資料を上から読んでください。詳しい内容は、それぞれの資料へのリンクから読めます。

---

## 1. このプロジェクトについて

TPCVisualLoggingApp は、C# で作る Windows アプリケーションです。
株式会社テクノ製の **PC ベースファインモーションコントローラ RTMC64**（以下 RTMC64）と、**通信ライブラリ（DLL）** 経由でやり取りします。

テクノ製のコントローラには、軸の位置や入出力の状態を制御周期に同期して記録する **TPC ロギング** という機能があります。
このアプリは、TPC ロギングのデータを取得・表示・解析します。これまで使ってきた Excel のツール **TPC-Excel** を拡張した進化版で、Excel に頼らないテクノ独自のアプリにします。
最初は TPC-Excel の 3 つの解析（サーボ応答精度、真円精度、直線精度）を引き継ぎ、その後、社内のアイデア出しやユーザーへのヒアリングで見つけたユースケースを追加していく計画です。
目的・背景・進め方は [企画書](../01_project/proposal.md)（ドラフト）を読んでください。

第 1 段階（最初のリリース）の範囲は決まっています（[ADR-001](../02_design/decisions/ADR-001_phase1-scope.md)）。要点は次のとおりです。

- 対象の機種は **RTMC64-EC だけ** です。開発機も RTMC64-EC です。
- TPC-Excel の 3 つの解析に加えて、コントローラから直接データを取ること、解析の進化点（数値の指標、中心や基準直線の自動推定など）、CSV への保存、長い時間の記録を入れます。
- 動作を段階的に確かめられるよう、仕様を 3 つ（spec 001〜003）に分けて作ります。

| spec | 内容 |
|---|---|
| 001 | 画面の基本、CSV の読み込み、TPC-Excel と同じ 3 つの解析、グラフ、印刷と画像保存 |
| 002 | コントローラからの取得（2 軸 × 3 種類の位置 = 6ch）、読み続けての長時間記録、CSV への保存 |
| 003 | 解析の進化点（3 種類の位置の比較を含む） |

ロギング中のリアルタイム表示は、第 1 段階には入れません（将来の候補として [バックログ](../02_design/backlog.md) に載せています）。

| 用語 | 意味 |
|---|---|
| RTMC64 | 64 軸を制御できるモーションコントローラです。PC の中で動きます。次の 2 機種があります |
| RTMC64-M3 | MECHATROLINK-Ⅲ（サーボアンプとつなぐ通信方式）を使う機種です。マニュアルでは [M3] と書かれます |
| RTMC64-EC | EtherCAT（同上）を使う機種です。マニュアルでは [EC] と書かれます |
| 通信ライブラリ | Windows のアプリと RTMC64 の間で、共有メモリを使ってデータをやり取りする DLL です。型式は RT64M3COM-NT（M3 用）と RT64ECCOM-NT（EC 用）です |
| INtime | RTMC64 が動いているリアルタイム OS です。Windows と同じ PC で動きます |
| TPC ロギング | TPC はツールパスチェッカーの略です。コントローラの中のデータ（位置、入出力、ステータスなど）を、制御周期に同期して記録する機能です |
| TPC-Excel | TPC ロギングのデータを解析する、テクノの Excel ツールです（マクロ付きのテンプレート）。このアプリの前身にあたります |
| セッティング PC | テクノ標準の運転・設定用ソフトです。TPC-Excel を使うときは、このソフトでロギングして CSV に保存します |

今後は、同じくテクノ製の次の製品にも対応する予定です。マニュアルはすでに集めてあります（[8 章](#8-資料の読み方)）。

| 製品 | 概要（各マニュアルの表紙から） |
|---|---|
| PLMC-M3EX | 横河電機の PLC「FA-M3R」専用のオープンモーションコントローラです。MECHATROLINK-Ⅲ を使います |
| SLM8000 | 標準ワンボードモーションコントローラです。AS710 コントローラボードを使い、PC とはイーサネットか RS-232C でつなぎます |

このうち **SLM8000 は、設計のときに常に 2 番目に置く機種** です（[バックログ B-003](../02_design/backlog.md)）。第 1 段階で作るのは RTMC64-EC 用だけですが、インターフェース・構造体・クラスなどを決めるときは、あとで SLM8000 にも対応できるかを確かめてください。RTMC64-EC との違いは [slm8000/sdk.md](../03_reference/slm8000/sdk.md)「RTMC64-EC との違い」にまとめています。

関係者（体制）は **未整備** です（企画書でも TBD です）。

## 2. 最初に知っておくべき注意（必読）

> **開発機は実機のコントローラ（RTMC64）に、DLL 経由でつながっています。**
> ビルドした exe やテストを実行すると、実機に指令が届き、モータや装置が動くおそれがあります。

- ビルドするだけなら安全です。[scripts/build.ps1](../../scripts/build.ps1) はビルドだけを行い、exe は実行しません。
- exe（`bin/Debug`、`bin/Release` の中のもの）や `dotnet run`、`dotnet test`、`vstest.console` を実行する前に、装置の状態と周りの安全を確認してください。
- Claude Code は、これらの操作を実行する前に毎回確認を求めます（[7 章](#7-claude-code-を使うとき)）。
- プロジェクトの原則（[constitution](../../.specify/memory/constitution.md) 原則 I）で、次のことを決めています。
  - 第 1 段階のアプリは、軸を動かす指令や、入出力・パラメータを変えるコマンドを送りません。送るのは TPC ロギングのコマンドだけです。
  - 自動テストでは通信ライブラリの DLL を読み込まず、モックに差し替えます。
- RTMC64 のソフトウェアパッケージに入っているドライバ、設定ソフト、サンプルなどのプログラムも、ビルド・実行しないでください（[sdk.md](../03_reference/rtmc64/sdk.md)「注意」）。

## 3. 開発環境のセットアップ

### 必要なツール

| ツール | 用途 | 根拠 |
|---|---|---|
| Visual Studio 2019 | C# のビルド。MSBuild（16.x）を使います | [build.ps1](../../scripts/build.ps1) が VS2019 の MSBuild を探します |
| Git | バージョン管理 | — |
| Python 3.12 | PDF マニュアルを Markdown に変換するツールと、TPC-Excel から VBA を取り出すツール用です。アプリの開発だけなら不要です | 開発者の手元にある `.venv` が Python 3.12 です |
| RTMC64 の通信ライブラリ | アプリから RTMC64 と通信します | インストールはしません。レジストリや環境変数を変えないポリシーのパッケージで、DLL を exe と同じフォルダに置いて使います（[sdk.md](../03_reference/rtmc64/sdk.md)「DLL の置き方」） |

Visual Studio 2019 のエディションやワークロード、.NET Framework のバージョンは **未整備** です（ソリューション `.sln` がまだリポジトリにありません）。

### クローン後にやること

1. リポジトリをクローンし、`develop` ブランチに切り替えます。
   ```powershell
   git clone <リポジトリの URL>
   cd TPCVisualLoggingApp
   git switch develop
   ```
2. （Python のツールを使う場合だけ）Python の仮想環境を作ります。
   ```powershell
   py -3.12 -m venv .venv
   .\.venv\Scripts\pip install docling pypdfium2 PyYAML olefile
   ```
   `requirements.txt` はまだありません。上のパッケージは [convert_pdf.py](../../scripts/convert_pdf.py) と [extract_vba.py](../../scripts/extract_vba.py) の import から調べたものです。
3. RTMC64 のソフトウェアパッケージ（通信ライブラリの本体、ヘッダーなど）を、`docs/03_reference/rtmc64/sdk/` の下に置きます。購入した人だけが入手できるものなので、Git には入っていません。置き方は [sdk.md](../03_reference/rtmc64/sdk.md) を読んでください。
   アプリが対応するのは、パッケージ V310A（通信ライブラリの DLL 1.9.0.0）以降です（[ADR-003](../02_design/decisions/ADR-003_supported-versions.md)）。
   SLM8000 のパッケージも同じように `docs/03_reference/slm8000/sdk/` の下に置きます（[slm8000/sdk.md](../03_reference/slm8000/sdk.md)）。PLMC-M3EX のパッケージの置き方の説明は **未整備** です。

## 4. リポジトリの構成

| パス | 役割 |
|---|---|
| [AGENTS.md](../../AGENTS.md) / [CLAUDE.md](../../CLAUDE.md) | AI エージェント（Claude Code など）向けの作業ルールです。CLAUDE.md は AGENTS.md を読み込むだけです |
| [.claude/settings.json](../../.claude/settings.json) | Claude Code の権限設定です（チームで共有します） |
| [.claude/README.md](../../.claude/README.md) | 権限設定の意味と理由の説明です |
| [.claude/skills/](../../.claude/skills/) | Claude Code の skill です（この資料を作る `/onboarding-dev`、営業向け資料を作る `/onboarding-sales`、PDF マニュアルを変換する `/convert-manual`） |
| [scripts/build.ps1](../../scripts/build.ps1) | VS2019 の MSBuild でビルドするスクリプトです |
| [scripts/convert_pdf.py](../../scripts/convert_pdf.py) | PDF マニュアルを Markdown に変換するスクリプトです |
| [scripts/extract_vba.py](../../scripts/extract_vba.py) | Excel ファイル（TPC-Excel など）から VBA のソースコードを取り出すスクリプトです |
| [docs/](../) | ドキュメントです（[8 章](#8-資料の読み方)） |
| [specs/](../../specs/) | 確定した仕様です（Spec Kit で機能ごとに作ります） |
| [.specify/](../../.specify/) | Spec Kit の設定・テンプレートと、プロジェクトの原則（[constitution.md](../../.specify/memory/constitution.md)）です |
| [.gitignore](../../.gitignore) / [.gitattributes](../../.gitattributes) | Git の除外設定と改行コードの設定です |

C# のソースコードは **未整備** です（まだリポジトリにありません）。

### 改行コードと文字コード

- `.cs`、`.csproj`、`.sln`、`.ps1`、`.bas`（VBA）などの Windows 系ファイルは、**CRLF** に固定しています（[.gitattributes](../../.gitattributes)）。
- Excel ファイル（`.xls`、`.xlt`、`.xlsx`、`.xlsm`、`.xltm`）は、バイナリとして扱います。
- `build.ps1` は **BOM 付き UTF-8** で保存しています。Windows PowerShell 5.1 で日本語が文字化けしないようにするためです。編集するときもこの形式を保ってください。

## 5. ビルドと実行

### ビルド

```powershell
.\scripts\build.ps1                          # Debug ビルド（.sln を自動で探します）
.\scripts\build.ps1 -Configuration Release
.\scripts\build.ps1 -Solution src\Foo.sln -Platform x64
```

- MSBuild は PATH に入っていません。スクリプトがまず vswhere で探し、見つからなければ VS2019 の既定のインストール先を探します。
- `.sln` が 0 個または複数あるときは、エラーになります。そのときは `-Solution` で指定してください。
- 詳しくは [.claude/README.md の「ビルドスクリプト」](../../.claude/README.md#ビルドスクリプトscriptsbuildps1) を読んでください。

### 実行

exe を実行すると実機と通信します。必ず [2 章](#2-最初に知っておくべき注意必読) を読んでから実行してください。
デバッグ実行の手順や、実機につながない動かし方（モックなど）は **未整備** です。

### PDF マニュアルの変換

```powershell
.\.venv\Scripts\python scripts\convert_pdf.py docs\03_reference\rtmc64\source\TB00-1044.pdf
.\.venv\Scripts\python scripts\convert_pdf.py <pdf> --reconvert   # キャッシュを使わず解析し直す
```

結果は、PDF と同じ製品フォルダの `converted/` に出力されます（例: `docs/03_reference/rtmc64/converted/`）。図の書き起こし（`figures.yaml`）と崩れた表の手修正（`tables.yaml`）は、スクリプトでは作れません。

Claude Code で `/convert-manual <pdf>` を実行すると、変換から `figures.yaml` と `tables.yaml` の作成、変換し直しての確認までを Claude が行います（[SKILL.md](../../.claude/skills/convert-manual/SKILL.md)）。書き起こした内容は、最後に人が確認してください。

### TPC-Excel の VBA の取り出し

TPC-Excel のマクロにはパスワードがかかっていますが、ソースコードはファイルの中にあり、取り出せます。取り出したものは [tpcexcel/converted/](../03_reference/tpcexcel/converted/) に置いてあります。原本が更新されたときだけ、取り出し直してください。

```powershell
.\.venv\Scripts\python scripts\extract_vba.py <Excel ファイル> <出力フォルダ>
```

具体的なコマンドは [tpcexcel/converted/README.md](../03_reference/tpcexcel/converted/README.md) にあります。

## 6. ブランチとコミットの運用

| ブランチ | 使い方 |
|---|---|
| `master` | 既定のブランチです（プルリクエストの向き先） |
| `develop` | 日々の開発に使うブランチです |

- コミットメッセージは **日本語で 1 行**、「何を、どうしたか」を書きます（例: 「.gitignore を C# 開発向けに見直し、.gitattributes を追加」）。
- リポジトリは今、個人の非公開 GitHub にあります。将来は社内 GitLab へ移す予定です。
- 機能ブランチの切り方、レビューの手順、master へのマージ条件は **未整備** です。

## 7. Claude Code を使うとき

権限設定の要点だけ書きます。全体は [.claude/README.md](../../.claude/README.md) を読んでください。

| 区分 | 主な操作 |
|---|---|
| 確認なしで実行 | 読むだけの git コマンド（status、diff、log など）、`build.ps1`、`convert_pdf.py` |
| 毎回確認 | git commit、push、merge、checkout など。exe やテストの実行（実機が動くおそれがあるため）。パッケージの追加。ファイルの削除 |
| 常に拒否 | `git push --force`、`git reset --hard`、`git clean`、`rm -rf`。原本の資料（`docs/03_reference/*/source/` の PDF や TPC-Excel）と、製品のソフトウェアパッケージ（`docs/03_reference/*/sdk/`）の書き換え。`.env` や署名鍵の読み取り |

[AGENTS.md](../../AGENTS.md) には、AI エージェントに守らせるルールを書いています。人が作業するときにも同じ考え方で進めてください。

- 実装の根拠にするのは、確定した仕様（`specs/`）と [constitution](../../.specify/memory/constitution.md) だけです。`docs/02_design/` は検討中の資料なので、実装には使いません。
- `spec.md` を `Approved` にするのは、ユーザー（プロジェクトの担当者）が確定を指示したときだけです。
- 製品のソフトウェアは要求に応じて随時進化していて、マニュアルの更新が追いついていないことがあります。マニュアルとソフトウェアパッケージ（ヘッダー、ラッパー、サンプルなど）が食い違っていたら、どちらが正しいかを自分で決めず、担当者に確認します。確認した結果は、確認した日と人を添えて資料（[sdk.md](../03_reference/rtmc64/sdk.md) など）に書きます。

- 自分だけで使う許可は `.claude/settings.local.json` に書きます（Git では共有しません）。
- `.claude/settings.json` のルールを変えたときは、[.claude/README.md](../../.claude/README.md) の表も同じ作業の中で直します（[AGENTS.md](../../AGENTS.md) のルール）。
- settings.json にはコメントや独自のキーを書けません。

## 8. 資料の読み方

`docs/` は次のように分けています。上から順に読むことをおすすめします。

| フォルダ | 中身 | 状態 |
|---|---|---|
| `00_onboarding/` | この資料 | — |
| [01_project/](../01_project/) | プロジェクトの背景・会議資料 | 企画書（ドラフト）あり。会議資料はまだありません |
| [specs/](../../specs/)（`docs/` の外） | 確定した仕様 | 置き場だけ用意。仕様はまだありません |
| [02_design/](../02_design/) | 仕様の検討中の資料（検討メモと ADR）。**確定した仕様ではありません** | 検討メモ 1 件、ADR 3 件、バックログあり |
| [03_reference/](../03_reference/) | 製品ごとのマニュアル | 3 製品の原本の PDF と、TPC-Excel 一式あり。Markdown 版は RTMC64 の 9 冊、SLM8000 の 6 冊、TPC-Excel 説明書 |
| `04_dev/` | コーディング規約などの開発ガイド | **未整備**（作る予定） |

### 01_project: プロジェクトの背景

| 文書 | 内容 |
|---|---|
| ★ [企画書](../01_project/proposal.md) | アプリの目的と背景（TPC ロギング、TPC-Excel の仕組みと制約）、ユースケース（TPC-Excel の 3 つの解析と進化点、新しい候補）、ユースケースを見つけて増やす進め方、開発の範囲（第 1 段階は RTMC64-EC）、リスク、決まったこと（ADR の一覧）。体制・スケジュールは TBD です |

### specs: 確定した仕様とプロジェクトの原則

確定した仕様は、Spec Kit で機能ごとに `specs/NNN-機能名/` に作ります。置き方と確定の扱いは [specs/README.md](../../specs/README.md) にあります。

| 文書 | 内容 | 状態 |
|---|---|---|
| ★ [constitution.md](../../.specify/memory/constitution.md) | プロジェクト全体の原則です。I 実機の安全を最優先する、II TPC-Excel と計算結果をそろえる、III 取得・表示・解析を分け、追加しやすく作る、IV 記録したデータを欠けさせず、欠けたら隠さない、V 確定した仕様と確かめたユースケースだけを作る | v1.0.0（2026-11-19 制定） |
| 機能ごとの仕様（`spec.md`） | — | まだありません。`Approved` の仕様も `Draft` の仕様も 0 件です。次は spec 001 を作る予定です |

### 02_design: 検討中の資料

**ここにあるのは確定した仕様ではありません。** 背景を知るために読み、実装の根拠には `specs/` を使ってください。書き方は [02_design/README.md](../02_design/README.md) にあります。

| 文書 | 内容 |
|---|---|
| ★ [検討メモ: 第 1 段階の範囲](../02_design/discussions/2026-11-19_phase1-scope.md) | 第 1 段階の範囲を決めた経緯と、spec 002 などで決める残りの論点（短い周期での読み出し、時間軸の求め方など） |
| ★ [ADR-001](../02_design/decisions/ADR-001_phase1-scope.md) | 第 1 段階の機能範囲（RTMC64-EC だけ、spec を 001〜003 に分ける、リアルタイム表示は入れない） |
| [ADR-002](../02_design/decisions/ADR-002_tpcexcel-maintenance.md) | TPC-Excel の保守（Office 2024 で動かす修正）は別のリポジトリで行い、計算は変えない |
| [ADR-003](../02_design/decisions/ADR-003_supported-versions.md) | 対応する製品ソフトウェアの版（開発時の最新版以降。RTMC64-EC は V310A、DLL 1.9.0.0 以降） |
| [backlog.md](../02_design/backlog.md) | 今は入れないが、将来入れたい機能の候補（リアルタイム表示、RTMC64-M3 への対応、SLM8000 への対応） |

### 03_reference: 製品ごとのマニュアル

マニュアルは製品ごとのフォルダに分けています。どの製品も、原本を `source/` に、変換したものを `converted/` に置きます。

| フォルダ | 製品 | 状態 |
|---|---|---|
| [rtmc64/](../03_reference/rtmc64/) | RTMC64-M3 / RTMC64-EC | 原本 PDF と Markdown 版（9 冊）、ソフトウェアパッケージの説明（[sdk.md](../03_reference/rtmc64/sdk.md)） |
| [plmcm3ex/](../03_reference/plmcm3ex/) | PLMC-M3EX | 原本 PDF のみ |
| [slm8000/](../03_reference/slm8000/) | SLM8000 | 原本 PDF と Markdown 版（TPC ロギングに関わる 6 冊）、ソフトウェアパッケージの説明（[sdk.md](../03_reference/slm8000/sdk.md)） |
| [tpcexcel/](../03_reference/tpcexcel/) | TPC-Excel（解析ツール） | 原本（説明書 PDF、Excel 本体、サンプル CSV）と、取り出した VBA ソース |

3 製品とも、マニュアルの種類がそろっています。「ユーザーズマニュアル」「セッティング PC」「ROMSW 設定ソフト」「通信ライブラリ」「送受信データ」「テクノコード変換ライブラリ」「G コード変換ライブラリ」は、どの製品にもあります。製品ごとの違いを調べるときは、同じ種類のマニュアルどうしを見比べてください。

#### rtmc64: RTMC64 のマニュアル

原本の PDF は [source/](../03_reference/rtmc64/source/) にあります。Markdown に変換したものは [converted/](../03_reference/rtmc64/converted/) にあります。★ はアプリを作るために最初に読むものです。

| 番号 | 名前 | 内容 | Markdown 版 |
|---|---|---|---|
| ★ TB00-1043 | 通信ライブラリ リファレンスマニュアル | アプリから RTMC64 と通信する DLL の関数仕様です（InitCommProc、SendData、ReceiveData など） | [あり](../03_reference/rtmc64/converted/TB00-1043.md) |
| ★ TB00-1044 | 送受信データ マニュアル | 通信でやり取りするデータ（パラメータ、ステータス、入出力状態、TPC ロギングのコマンドなど）の形式です | [あり](../03_reference/rtmc64/converted/TB00-1044.md) |
| ★ TB00-1041 | セッティング PC マニュアル | 設定用 PC ソフトの使い方です。TPC ロギングの画面（5-2-8-2）があります | [あり](../03_reference/rtmc64/converted/TB00-1041.md) |
| TB00-1040 | ユーザーズマニュアル | RTMC64 全体の説明です。下の表のように、編や章ごとに PDF が分かれています | 一部あり（下の表） |
| TB00-1042 | ROMSW 設定ソフト マニュアル | ユーザー設定パラメータ（ROMSW）を設定するソフトです。RTMC64-EC では、EtherCAT の通信周期は ROMSW ではなく EtherCAT のマスタースタックの設定で変えます（6-1-1-5） | [あり](../03_reference/rtmc64/converted/TB00-1042.md) |
| TB00-1045 | テクノコード変換ライブラリ リファレンスマニュアル | 運転プログラム（テクノコード）を変換するライブラリです。第 1 段階では使いません | [あり](../03_reference/rtmc64/converted/TB00-1045.md) |
| TB00-1046 | G コード変換ライブラリ リファレンスマニュアル | G コードの運転プログラムを変換するライブラリです。第 1 段階では使いません | [あり](../03_reference/rtmc64/converted/TB00-1046.md) |

ユーザーズマニュアル（TB00-1040）の内訳です。原本は [source/TB00-1040-V13/](../03_reference/rtmc64/source/TB00-1040-V13/) にあります（Ver.1.3）。

| ファイル | 内容 | Markdown 版 |
|---|---|---|
| TB00-1040 | 表紙と全体の目次 | なし |
| TB00-1040A | Ⅰ 導入編。インストールから、モータ単体での運転・調整までの手順です。初めての方はここで全体像をつかめます | なし |
| TB00-1040D | Ⅱ 機能編 1～3 章（機能概要、入出力機能、動作モード） | [あり](../03_reference/rtmc64/converted/TB00-1040D.md) |
| TB00-1040E | Ⅱ 機能編 4～5 章（INtime ユーザ制御ソフトとのインターフェース、サーボ制御）。5-17-5 にモーションアナライザでの真円精度解析の例があります | [あり](../03_reference/rtmc64/converted/TB00-1040E.md) |
| TB00-1040F | Ⅱ 機能編 6～7 章（運転プログラムの処理と各種運転方法、運転プログラム詳細。ステータスのビットの割り当てもここにあります） | [あり](../03_reference/rtmc64/converted/TB00-1040F.md) |
| TB00-1040H | Ⅲ 試運転・調整編 | なし |
| ★ TB00-1040I | Ⅲ 試運転・調整編 7-11～7-14（サーボ系の調整、**TPC による精度チェック**、現地でのサーボ立ち上げ手順例） | [あり](../03_reference/rtmc64/converted/TB00-1040I.md) |
| TB00-1040J | メンテナンス編 | なし |
| TB00-1040K | 改版履歴 | なし |

ソフトウェアパッケージ（通信ライブラリの本体、C# の宣言、サンプル）を調べて分かったこと（TPC ロギングのコマンドの値、データの大きさ、読み出しの流れ、DLL の置き方、通信設定ファイル `rt64ecdrv.ini` など）は、★ [sdk.md](../03_reference/rtmc64/sdk.md) にまとめています。パッケージの中身そのものは Git に入っていません（[3 章](#3-開発環境のセットアップ)）。

#### plmcm3ex: PLMC-M3EX のマニュアル

原本の PDF は [source/](../03_reference/plmcm3ex/source/) にあります。Markdown 版はまだありません。

| 番号 | 名前 | 内容 |
|---|---|---|
| ★ TB00-0983 | 通信ライブラリ リファレンスマニュアル（PLM3EX-COMNT） | アプリから PLMC-M3EX と通信する DLL の関数仕様です |
| ★ TB00-0984 | 送受信データ説明書 | 通信でやり取りするデータの形式です |
| TB00-0980 | ユーザーズマニュアル（導入編・ハードウェア編・機能編・試運転・調整編・メンテナンス編など） | PLMC-M3EX 全体の説明です。初めての方は「導入編」（TB00-0980A）で全体像をつかめます |
| TB00-0981 | セッティング PC マニュアル | 設定用 PC ソフトの使い方です |
| TB00-0982 | ROMSW 設定ソフト ユーザーズマニュアル | ユーザー設定パラメータ（ROMSW）を設定するソフトです |
| TB00-0985 | テクノコード仕様テキストプログラム変換ライブラリ（PLM3EX-TCNV） | テクノコードの運転プログラムを変換するライブラリです |
| TB00-0986 | G コード仕様テキストプログラム変換ライブラリ（PLM3EX-GCNV） | G コードの運転プログラムを変換するライブラリです |

#### slm8000: SLM8000 のマニュアル

原本の PDF は [source/](../03_reference/slm8000/source/) にあります。TPC ロギングに関わるものだけ、Markdown に変換して [converted/](../03_reference/slm8000/converted/) に置いています。★ は、RTMC64-EC との違いを確かめるために最初に読むものです。

| 番号 | 名前 | 内容 | Markdown 版 |
|---|---|---|---|
| ★ TB00-1133 | 通信ライブラリ リファレンスマニュアル | アプリから SLM8000 と通信する DLL の関数仕様です。接続の設定（イーサネット／RS-232C）もここにあります | [あり](../03_reference/slm8000/converted/TB00-1133.md) |
| ★ TB00-1134 | 送受信データ マニュアル | 通信でやり取りするデータの形式です。TPC ロギングのコマンドは 4-1-24～4-1-27 と 4-2-58～4-2-60 にあります | [あり](../03_reference/slm8000/converted/TB00-1134.md) |
| TB00-1131 | セッティング PC マニュアル | 設定用 PC ソフトの使い方です。TPC ロギング画面（4-2-8-2）と制御周期モニタ（4-2-8-8）があります | [あり](../03_reference/slm8000/converted/TB00-1131.md) |
| TB00-1132 | ROMSW 設定ソフト マニュアル | ユーザー設定パラメータ（ROMSW）を設定するソフトです。サーボ制御周期の設定（6-1-1）があります | [あり](../03_reference/slm8000/converted/TB00-1132.md) |
| TB00-1130 | ユーザーズマニュアル（導入編・ハードウェア編・機能編・実機試運転・調整編・メンテナンス編など） | SLM8000 全体の説明です。初めての方は「導入編」（TB00-1130A）で全体像をつかめます | 一部あり: [TB00-1130D](../03_reference/slm8000/converted/TB00-1130D.md)（機能編 1～2 章）、[TB00-1130H](../03_reference/slm8000/converted/TB00-1130H.md)（実機試運転・調整編。TPC による精度チェックを含む） |
| TB00-1135 | テクノコード変換ライブラリ リファレンスマニュアル | テクノコードの運転プログラムを変換するライブラリです | なし |
| TB00-1136 | G コード変換ライブラリ リファレンスマニュアル | G コードの運転プログラムを変換するライブラリです | なし |
| TB00-1189 | テキストコマンド通信仕様 | テキストのコマンドで SLM8000 と通信する方法です（SLM8000 だけにあります） | なし |
| TB00-1216 | イーサネット通信仕様 | イーサネットで SLM8000 と通信する方法です（SLM8000 だけにあります） | なし |

ソフトウェアパッケージ（V2.8.0A）を調べて分かったことは、★ [sdk.md](../03_reference/slm8000/sdk.md) にまとめています。TPC ロギングの値、制御周期の読み方（設定値は `slm8000drv.ini`、`DAT_CYCLETIME` は実際の周期）、RTMC64-EC との違いの表などです。

#### tpcexcel: TPC-Excel（このアプリの前身）

原本は [source/](../03_reference/tpcexcel/source/) に、取り出した VBA ソースと説明書の Markdown 版は [converted/](../03_reference/tpcexcel/converted/) にあります。解析の計算を作り直すときは、TPC-Excel と結果が一致するよう、VBA の計算式とサンプルデータで確かめてください（constitution 原則 II）。計算の根拠は、説明書ではなく VBA のソースです。

TPC-Excel を新しい Excel（Office 2024）で動かすための修正は、別のリポジトリで行います（[ADR-002](../02_design/decisions/ADR-002_tpcexcel-maintenance.md)）。修正が終わった版は、あとでこのフォルダに追加します。

| ファイル | 内容 |
|---|---|
| ★ TB04-1478-V11.pdf（[Markdown 版](../03_reference/tpcexcel/converted/TB04-1478-V11.md)） | TPC-Excel 説明書 Ver1.1（2011 年）。TPC ロギングの原理、CSV の形式、3 つの解析の内容と操作方法です。Markdown 版は本文の語順が崩れている箇所があるので、画面の吹き出しの文は図の書き起こし（「図の内容」）を読んでください |
| Tpc-Excel_V16A/TPC-Excel.xlt | TPC-Excel 本体（Excel 2000/2003 向け） |
| Tpc-Excel2010_V003/TPC-Excel2010_V003.xltm | TPC-Excel 本体（Excel 2010 向け）。解析の計算は V16A と同じです |
| Tpc-Excel_V16A/sampl-data/ | サンプルデータです。サーボ応答精度用 3 つ（servosmpl1〜3.csv）、真円精度用 2 つ（cirsmpl1〜2.csv） |
| ★ [converted/README.md](../03_reference/tpcexcel/converted/README.md) | 取り出した VBA ソースの説明です。どの解析の計算がどのモジュールにあるかの一覧があります |

## 9. 困ったときは

問い合わせ先（プロジェクトの担当者、装置の担当者、メーカーの窓口）は **未整備** です。

<!-- ここより下は手で書き足してかまいません -->
