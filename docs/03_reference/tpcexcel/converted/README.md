# TPC-Excel の VBA ソース

[source/](../source/) にある TPC-Excel 本体から、マクロ（VBA）のソースコードを取り出したものです。
TPCVisualLoggingApp で解析の計算を作り直すときに、元の計算式を確かめるために使います。

| フォルダ | 元のファイル | 対応 Excel |
|---|---|---|
| [Tpc-Excel_V16A/vba/](Tpc-Excel_V16A/vba/) | `source/Tpc-Excel_V16A/TPC-Excel.xlt` | Excel 2000/2003 |
| [Tpc-Excel2010_V003/vba/](Tpc-Excel2010_V003/vba/) | `source/Tpc-Excel2010_V003/TPC-Excel2010_V003.xltm` | Excel 2010 |

2 つの版で解析の計算は同じです。違うのは、Excel 2010 に合わせたグラフ表示と印刷の処理です。

## 解析の計算がある場所

| 解析 | モジュール | 手続き |
|---|---|---|
| サーボ応答精度 | `Module2` | `speedckeck`、`speedcalc` |
| 真円精度 | `Module21` | `circleckeck`、`AnalyzeErr` |
| 直線精度 | `Module11` | `linearity_chk` |
| CSV の読み込み | `Module11` | `read_csv` |
| 定数（シート名、セル番地、最大点数など） | `Module1` | 先頭の `Public Const` |

## 取り出し方

VBA プロジェクトにはパスワードがかかっていますが、これは VBA エディタで開くのを止めるだけで、ソースコードはファイルの中に圧縮して保存されています。
[scripts/extract_vba.py](../../../../scripts/extract_vba.py) がそれを展開します。

```powershell
.\.venv\Scripts\python scripts\extract_vba.py docs\03_reference\tpcexcel\source\Tpc-Excel_V16A\TPC-Excel.xlt docs\03_reference\tpcexcel\converted\Tpc-Excel_V16A\vba
.\.venv\Scripts\python scripts\extract_vba.py docs\03_reference\tpcexcel\source\Tpc-Excel2010_V003\TPC-Excel2010_V003.xltm docs\03_reference\tpcexcel\converted\Tpc-Excel2010_V003\vba
```

- 文字コードは UTF-8、改行は CRLF で出力します。
- ユーザーフォームの見た目（`.frx`）は取り出しません。フォームのモジュールには、イベント処理のコードだけが入っています。
- 取り出したファイルは手で直さないでください。直したいときは、取り出し直してください。
