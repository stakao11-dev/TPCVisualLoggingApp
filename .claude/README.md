# Claude Code 権限設定の説明

[settings.json](settings.json) に書いた権限ルールの意味と、そう決めた理由をまとめています。
settings.json は JSON なのでコメントを書けません。ルールを変えたときは、このファイルも一緒に直してください。

## 前提

- アプリは C# で書き、Visual Studio 2019 でビルドします。
- 開発中も **実機のコントローラ (RTMC64) に DLL 経由で接続しています**。ビルドした exe やテストを動かすと、実機に指令が届くおそれがあります。
- Claude Code は Bash と PowerShell の両方のツールを使えます。そのため、同じ内容のルールを `Bash(...)` と `PowerShell(...)` の 2 通りで書いています。

## 3 つの区分

| 区分 | 動き |
|---|---|
| `allow` | 確認なしで実行します |
| `ask` | 実行する前に、毎回ユーザーに確認します |
| `deny` | 常に拒否します |

どこにも書いていない操作は、既定の動きに従って確認が求められます。
1 つの操作に複数のルールが当てはまるときは、`deny` → `ask` → `allow` の順に優先されます。

## allow（確認なしで実行）

| ルール | 意味・理由 |
|---|---|
| `git status` / `diff` / `log` / `show` / `branch` | 読むだけの git コマンドです。何も変更しないので許可しています |
| `scripts/build.ps1` | ビルド用のスクリプトです（詳しくは下の「ビルドスクリプト」）。ビルドはするが exe は実行しないので、実機には影響しません |
| `.venv/Scripts/python scripts/convert_pdf.py` | PDF マニュアルを Markdown に変換するスクリプトです。作業フォルダの中で完結します |
| `WebFetch(domain:learn.microsoft.com)` | C# や .NET の公式ドキュメントを読めるようにします |

## ask（毎回確認）

### Git の変更系
| ルール | 意味・理由 |
|---|---|
| `git commit` / `push` / `merge` / `rebase` | 履歴を変えたり、リモートに公開したりする操作です。コミットは毎回確認すると決めています |
| `git checkout` / `switch` / `stash` | 作業中の変更が切り替わったり退避されたりして、気づかないうちに作業状態が変わるおそれがあります |

### 実機に影響しうる操作
| ルール | 意味・理由 |
|---|---|
| `*bin/Debug/*`、`*bin/Release/*`（`\` 区切りも同様） | ビルドした exe を実行すると、DLL 経由で実機のコントローラと通信します |
| `*vstest.console*`、`dotnet test` | テストから DLL を呼ぶと、実機が動くおそれがあります。DLL をモックに差し替えた単体テストだけにする場合は、allow への移動を検討してください |
| `dotnet run` | exe の実行と同じ理由です |
| `PowerShell(Start-Process:*)` | どのプログラムでも起動できてしまうため確認します |

### パッケージ・依存関係
| ルール | 意味・理由 |
|---|---|
| `dotnet add` / `nuget install` / `Install-Package` | NuGet パッケージを追加する操作です。依存関係を増やすときは人が判断します |
| `packages.config` の Edit / Write | ファイルを直接書き換えてパッケージを追加する経路もふさぎます |
| `.venv/Scripts/pip install` | Python の仮想環境にパッケージを追加する操作です |

### その他
| ルール | 意味・理由 |
|---|---|
| `scripts/build.ps1` の Edit / Write | build.ps1 は確認なしで実行できます。中身を書き換えられると何でも実行できてしまうので、編集だけは確認します |
| `PowerShell(Remove-Item:*)` | ファイルを削除する操作です |
| `PowerShell(Stop-Process:*)` | 動いているプロセスを止める操作です。Visual Studio や通信中のプロセスを巻き込むおそれがあります |

## deny（常に拒否）

| ルール | 意味・理由 |
|---|---|
| `docs/03_reference/*/source/**` の Edit / Write | 原本の PDF マニュアルです（`*` は製品ごとのフォルダ）。変換の元になるので書き換えを禁止しています |
| `git push --force` / `-f` | リモートの履歴が壊れます |
| `git reset --hard` / `git clean` | まだコミットしていない作業が消え、元に戻せません |
| `rm -rf` | まとめて削除する操作で、元に戻せません |
| `.env`、`.env.*`、`*.pfx`、`*.snk` の Read | パスワードや署名用の鍵など、機密情報が入るファイルです |

> **注意:** シェルコマンドの deny は前方一致で判定するので、完全ではありません。
> 例えば `git push origin main --force` のように順番を変えると通ってしまいます。最後の守りは、人が確認することです。

## ビルドスクリプト（scripts/build.ps1）

MSBuild は PATH に入っていないので、このスクリプトで VS2019 の `MSBuild.exe` を探してビルドします。

```powershell
.\scripts\build.ps1                          # Debug ビルド（.sln を自動で探す）
.\scripts\build.ps1 -Configuration Release
.\scripts\build.ps1 -Solution src\Foo.sln -Platform x64
```

- MSBuild はまず vswhere で探します。見つからなければ、VS2019 の既定のインストール先を順に探します。
- ビルドだけを行い、exe は実行しません。
- Windows PowerShell 5.1 で日本語が化けないよう、**BOM 付きの UTF-8** で保存しています。編集するときもこの形式を保ってください。

## 個人用の設定

自分だけで使う許可は `.claude/settings.local.json` に書きます（Git では共有しません）。
チーム全員に関係するルールは settings.json に書き、このファイルも直してください。

## 設定が効いているか確かめる

Claude Code で `/permissions` を実行すると、読み込まれているルールの一覧が表示されます。
