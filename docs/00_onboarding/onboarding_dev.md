# TPCVisualLoggingApp オンボーディング

- 最終更新日: 2026-11-12
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

関係者（体制）は **未整備** です（企画書でも TBD です）。

## 2. 最初に知っておくべき注意（必読）

> **開発機は実機のコントローラ（RTMC64）に、DLL 経由でつながっています。**
> ビルドした exe やテストを実行すると、実機に指令が届き、モータや装置が動くおそれがあります。

- ビルドするだけなら安全です。[scripts/build.ps1](../../scripts/build.ps1) はビルドだけを行い、exe は実行しません。
- exe（`bin/Debug`、`bin/Release` の中のもの）や `dotnet run`、`dotnet test`、`vstest.console` を実行する前に、装置の状態と周りの安全を確認してください。
- Claude Code は、これらの操作を実行する前に毎回確認を求めます（[7 章](#7-claude-code-を使うとき)）。

## 3. 開発環境のセットアップ

### 必要なツール

| ツール | 用途 | 根拠 |
|---|---|---|
| Visual Studio 2019 | C# のビルド。MSBuild（16.x）を使います | [build.ps1](../../scripts/build.ps1) が VS2019 の MSBuild を探します |
| Git | バージョン管理 | — |
| Python 3.12 | PDF マニュアルを Markdown に変換するツールと、TPC-Excel から VBA を取り出すツール用です。アプリの開発だけなら不要です | 開発者の手元にある `.venv` が Python 3.12 です |
| RTMC64 の通信ライブラリ | アプリから RTMC64 と通信します | インストール方法は [通信ライブラリリファレンス](../03_reference/rtmc64/converted/TB00-1043.md)「6. インストール」 |

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

ベンダー DLL の置き場所は **未整備** です。`.gitignore` では `lib/x64/` と `lib/x86/` を候補にしていますが、フォルダ名は仮のものです。

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
| 常に拒否 | `git push --force`、`git reset --hard`、`git clean`、`rm -rf`。原本の資料（`docs/03_reference/*/source/` の PDF や TPC-Excel）の書き換え。`.env` や署名鍵の読み取り |

- 自分だけで使う許可は `.claude/settings.local.json` に書きます（Git では共有しません）。
- `.claude/settings.json` のルールを変えたときは、[.claude/README.md](../../.claude/README.md) の表も同じ作業の中で直します（[AGENTS.md](../../AGENTS.md) のルール）。
- settings.json にはコメントや独自のキーを書けません。

## 8. 資料の読み方

`docs/` は次のように分けています。上から順に読むことをおすすめします。

| フォルダ | 中身 | 状態 |
|---|---|---|
| `00_onboarding/` | この資料 | — |
| [01_project/](../01_project/) | プロジェクトの背景・会議資料 | 企画書（ドラフト）あり。会議資料はまだありません |
| [02_design/](../02_design/) | 仕様の検討中の資料（検討メモと ADR）。**確定した仕様ではありません**。確定した仕様は [specs/](../../specs/) にあります | 置き場と書き方だけ用意。資料はまだありません |
| [03_reference/](../03_reference/) | 製品ごとのマニュアル | 3 製品の原本の PDF と、TPC-Excel 一式あり。Markdown 版は RTMC64 の 2 冊だけ |
| `04_dev/` | コーディング規約などの開発ガイド | **未整備**（作る予定） |

### 01_project: プロジェクトの背景

| 文書 | 内容 |
|---|---|
| ★ [企画書](../01_project/proposal.md) | アプリの目的と背景（TPC ロギング、TPC-Excel の仕組みと制約）、ユースケース（TPC-Excel の 3 つの解析と、新しい候補）、ユースケースを見つけて増やす進め方、開発の範囲（第 1 段階は RTMC64）、リスク。体制・スケジュールは TBD です |

### 03_reference: 製品ごとのマニュアル

マニュアルは製品ごとのフォルダに分けています。どの製品も、原本を `source/` に、変換したものを `converted/` に置きます。

| フォルダ | 製品 | 状態 |
|---|---|---|
| [rtmc64/](../03_reference/rtmc64/) | RTMC64-M3 / RTMC64-EC | 原本 PDF と Markdown 版（2 冊） |
| [plmcm3ex/](../03_reference/plmcm3ex/) | PLMC-M3EX | 原本 PDF のみ |
| [slm8000/](../03_reference/slm8000/) | SLM8000 | 原本 PDF のみ |
| [tpcexcel/](../03_reference/tpcexcel/) | TPC-Excel（解析ツール） | 原本（説明書 PDF、Excel 本体、サンプル CSV）と、取り出した VBA ソース |

3 製品とも、マニュアルの種類がそろっています。「ユーザーズマニュアル」「セッティング PC」「ROMSW 設定ソフト」「通信ライブラリ」「送受信データ」「テクノコード変換ライブラリ」「G コード変換ライブラリ」は、どの製品にもあります。製品ごとの違いを調べるときは、同じ種類のマニュアルどうしを見比べてください。

#### rtmc64: RTMC64 のマニュアル

原本の PDF は [source/](../03_reference/rtmc64/source/) にあります。Markdown に変換したものは [converted/](../03_reference/rtmc64/converted/) にあります。★ はアプリを作るために最初に読むものです。

| 番号 | 名前 | 内容 | Markdown 版 |
|---|---|---|---|
| ★ TB00-1043 | 通信ライブラリ リファレンスマニュアル | アプリから RTMC64 と通信する DLL の関数仕様です（InitCommProc、SendData、ReceiveData など） | [あり](../03_reference/rtmc64/converted/TB00-1043.md) |
| ★ TB00-1044 | 送受信データ マニュアル | 通信でやり取りするデータ（パラメータ、ステータス、入出力状態など）の形式です | [あり](../03_reference/rtmc64/converted/TB00-1044.md) |
| TB00-1040 | ユーザーズマニュアル（導入編・機能編・試運転・調整編・メンテナンス編など） | RTMC64 全体の説明です。初めての方は「導入編」（TB00-1040A）で全体像をつかめます | なし |
| TB00-1041 | セッティング PC マニュアル | 設定用 PC ソフトの使い方です | なし |
| TB00-1042 | ROMSW 設定ソフト マニュアル | ユーザー設定パラメータ（ROMSW）を設定するソフトです | なし |
| TB00-1045 | テクノコード変換ライブラリ リファレンスマニュアル | 運転プログラム（テクノコード）を変換するライブラリです | なし |
| TB00-1046 | G コード変換ライブラリ リファレンスマニュアル | G コードの運転プログラムを変換するライブラリです | なし |

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

原本の PDF は [source/](../03_reference/slm8000/source/) にあります。Markdown 版はまだありません。

| 番号 | 名前 | 内容 |
|---|---|---|
| ★ TB00-1133 | 通信ライブラリ リファレンスマニュアル | アプリから SLM8000 と通信する DLL の関数仕様です |
| ★ TB00-1134 | 送受信データ マニュアル | 通信でやり取りするデータの形式です |
| TB00-1130 | ユーザーズマニュアル（導入編・ハードウェア編・機能編・実機試運転・調整編・メンテナンス編など） | SLM8000 全体の説明です。初めての方は「導入編」（TB00-1130A）で全体像をつかめます |
| TB00-1131 | セッティング PC マニュアル | 設定用 PC ソフトの使い方です |
| TB00-1132 | ROMSW 設定ソフト マニュアル | ユーザー設定パラメータ（ROMSW）を設定するソフトです |
| TB00-1135 | テクノコード変換ライブラリ リファレンスマニュアル | テクノコードの運転プログラムを変換するライブラリです |
| TB00-1136 | G コード変換ライブラリ リファレンスマニュアル | G コードの運転プログラムを変換するライブラリです |
| TB00-1189 | テキストコマンド通信仕様 | テキストのコマンドで SLM8000 と通信する方法です（SLM8000 だけにあります） |
| TB00-1216 | イーサネット通信仕様 | イーサネットで SLM8000 と通信する方法です（SLM8000 だけにあります） |

#### tpcexcel: TPC-Excel（このアプリの前身）

原本は [source/](../03_reference/tpcexcel/source/) に、取り出した VBA ソースは [converted/](../03_reference/tpcexcel/converted/) にあります。解析の計算を作り直すときは、TPC-Excel と結果が一致するよう、VBA の計算式とサンプルデータで確かめてください。

| ファイル | 内容 |
|---|---|
| ★ TB04-1478-V11.pdf | TPC-Excel 説明書 Ver1.1（2011 年）。TPC ロギングの原理、CSV の形式、3 つの解析の内容と操作方法です |
| Tpc-Excel_V16A/TPC-Excel.xlt | TPC-Excel 本体（Excel 2000/2003 向け） |
| Tpc-Excel2010_V003/TPC-Excel2010_V003.xltm | TPC-Excel 本体（Excel 2010 向け）。解析の計算は V16A と同じです |
| Tpc-Excel_V16A/sampl-data/ | サンプルデータです。サーボ応答精度用 3 つ（servosmpl1〜3.csv）、真円精度用 2 つ（cirsmpl1〜2.csv） |
| ★ [converted/README.md](../03_reference/tpcexcel/converted/README.md) | 取り出した VBA ソースの説明です。どの解析の計算がどのモジュールにあるかの一覧があります |
## 9. 困ったときは

問い合わせ先（プロジェクトの担当者、装置の担当者、メーカーの窓口）は **未整備** です。

<!-- ここより下は手で書き足してかまいません -->
