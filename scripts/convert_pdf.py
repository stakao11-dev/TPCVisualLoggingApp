"""PDF マニュアルを docling で Markdown に変換する。

usage:
    .venv/Scripts/python scripts/convert_pdf.py <pdf> [-o <out.md>] [--reconvert]

処理は 2 段階:
  1. docling でレイアウト解析し、結果を .cache/docling/<名前>.json に保存する (時間がかかる)
  2. キャッシュを読み込み、後処理をかけて Markdown を書き出す

2 回目以降はキャッシュを再利用するので、後処理だけを素早くやり直せる。
解析からやり直す場合は --reconvert を付ける。

マニュアルは製品ごとに docs/03_reference/<製品>/source/ に置く (例: rtmc64)。
出力先は同じ製品の converted/ になる。-o を付けると任意の場所に出力できる。

出力 (例: docs/03_reference/rtmc64/source/TB00-1044.pdf):
    docs/03_reference/rtmc64/converted/
        TB00-1044.md        変換結果
        TB00-1044/
            images/         図の PNG (変換のたびに作り直す)
            figures.yaml    図の内容の書き起こし (手で書く)
            tables.yaml     崩れた表の手修正 (手で書く)
"""
from __future__ import annotations

import argparse
import re
import time
import unicodedata
from pathlib import Path

import pypdfium2 as pdfium
import yaml
from docling_core.types.doc import DocItemLabel, DoclingDocument, ImageRefMode, TableCell

ROOT = Path(__file__).resolve().parent.parent
CACHE_DIR = ROOT / ".cache" / "docling"
REFERENCE_DIR = ROOT / "docs" / "03_reference"


def default_out(pdf: Path) -> Path:
    """docs/03_reference/<製品>/source/ 配下の PDF なら、同じ製品の converted/<名前>.md を返す。"""
    for parent in pdf.resolve().parents:
        if parent.name == "source" and parent.parent.parent == REFERENCE_DIR:
            return parent.parent / "converted" / f"{pdf.stem}.md"
    raise SystemExit(
        f"{pdf} は docs/03_reference/<製品>/source/ の下にありません。-o で出力先を指定してください。"
    )


# ---------------------------------------------------------------------------
# 1. docling による解析
# ---------------------------------------------------------------------------
def analyze(pdf: Path) -> DoclingDocument:
    from docling.datamodel.base_models import InputFormat
    from docling.datamodel.pipeline_options import PdfPipelineOptions, TableFormerMode
    from docling.document_converter import DocumentConverter, PdfFormatOption

    opts = PdfPipelineOptions()
    opts.do_ocr = False  # テキスト埋め込み PDF なので OCR 不要
    opts.do_table_structure = True
    opts.table_structure_options.mode = TableFormerMode.ACCURATE
    opts.table_structure_options.do_cell_matching = True
    opts.generate_picture_images = True  # 図を画像として書き出す
    opts.images_scale = 2.0

    converter = DocumentConverter(
        format_options={InputFormat.PDF: PdfFormatOption(pipeline_options=opts)}
    )
    t = time.time()
    doc = converter.convert(pdf).document
    print(f"analyzed {pdf.name}: {time.time() - t:.0f}s")
    return doc


def load_or_analyze(pdf: Path, reconvert: bool) -> DoclingDocument:
    cache = CACHE_DIR / f"{pdf.stem}.json"
    if cache.exists() and not reconvert:
        print(f"use cache: {cache}")
        return DoclingDocument.load_from_json(cache)
    doc = analyze(pdf)
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    doc.save_as_json(cache, image_mode=ImageRefMode.EMBEDDED)
    return doc


# ---------------------------------------------------------------------------
# 2. ドキュメント要素レベルの補正
#    要素のテキストをマーカーに置き換えておき、Markdown 出力後に中身を差し込む
# ---------------------------------------------------------------------------
DROP = "@@DROP@@"

# C 言語構造体の行らしいテキスト
CODE_RE = re.compile(
    r"^\s*("
    r"typedef\b|struct\b|union\b|enum\b|#define\b|#include\b"
    r"|/\*|\*/|//|[{}]"
    r"|\(?#(ifdef|ifndef|if|else|endif|include)\b"
    r"|\([^)\n]*#(include|define)\b"  # (または #include ... のような注記付きの行
    r"|[A-Za-z_][\w*]*\s+[\w\s*]*\w\s*\(.*\)\s*;\s*$"  # 関数のプロトタイプ宣言 (BOOL WINAPI Func(...);)
    r"|\*\w+.*;"  # ポインタ型のメンバーで、型名と分かれた断片 (*pLogFile;)
    r"|\"[\w.]+\.h\"\s*$"  # #include と分かれたヘッダファイル名
    r"|(unsigned\s+|signed\s+)?(char|short|long|int|float|double|void)\b.*;"
    r"|(unsigned\s+|signed\s+)?(char|short|long|int|float|double|void)\s*$"  # 型名だけの断片
    r"|[A-Z][A-Z0-9_]*\s+[\w\[\]*]+\s*;"
    r"|[A-Za-z_][\w\[\]]*\s*;?\s*$"  # 型名とメンバー名が別要素に分かれた断片 (TPCSEL / sel;)
    r")"
)


def _bbox(item):
    prov = item.prov[0]
    b = prov.bbox.to_top_left_origin(page_height=_page_heights[prov.page_no])
    return prov.page_no, b.l, b.t, b.r, b.b


_page_heights: dict[int, float] = {}


def patch_document(doc: DoclingDocument, pdf: pdfium.PdfDocument, toc: list[str]) -> dict[str, str]:
    """要素を補正し、マーカー → 置換文字列の対応表を返す。"""
    global _toc_keys
    _toc_keys = {_norm(t) for t in toc}
    for no, page in doc.pages.items():
        _page_heights[no] = page.size.height

    replacements: dict[str, str] = {}

    # (0) 図と誤認識されたコードや白抜き文字のラベル (戻り値 など) は、同じ位置のテキスト要素に置き換える
    #     (コードは後の (2) でコードブロックになる)
    for it, _ in list(doc.iterate_items()):
        if it.label.value != "picture" or not it.prov:
            continue
        b = it.prov[0].bbox
        text = pdf[it.prov[0].page_no - 1].get_textpage().get_text_bounded(
            left=b.l, bottom=b.b, right=b.r, top=b.t
        ).strip()
        is_code = "typedef" in text or len(re.findall(r";\s*(/\*|$)", text, flags=re.M)) >= 2
        if is_code or text in LABELS:
            doc.insert_text(sibling=it, label=DocItemLabel.TEXT, text=text,
                            prov=it.prov[0], after=False)
            doc.delete_items(node_items=[it])

    # (0') 見出しがページの途中に並べられていたら、そのページで見出しより下にある最初の要素の前へ移す
    #      (ページ上端の関数名などで起きる。目次にある題名だけを対象にする)
    items = [it for it, _ in doc.iterate_items() if getattr(it, "prov", None)]
    for it in list(items):
        if _norm(getattr(it, "text", "") or "") not in _toc_keys:
            continue
        page_no, top = it.prov[0].page_no, it.prov[0].bbox.t
        before = next((x for x in items if x is not it and x.prov[0].page_no == page_no
                       and x.prov[0].bbox.t < top - 1), None)  # 左下原点なので t が小さい = 下
        if before is None or items.index(before) > items.index(it):
            continue
        doc.insert_text(sibling=before, label=DocItemLabel.TEXT, text=it.text,
                        prov=it.prov[0], after=False)
        doc.delete_items(node_items=[it])
        items = [x for x, _ in doc.iterate_items() if getattr(x, "prov", None)]

    items = [it for it, _ in doc.iterate_items() if getattr(it, "prov", None)]

    # (1) 数式と誤判定されて文字が消えた要素を、元の文字列に戻す
    for i, it in enumerate(items):
        if it.label.value == "formula" and not it.text and it.orig:
            key = f"@@TEXT{i}@@"
            replacements[key] = it.orig
            it.text = key

    # (2) C 言語コードは段落・表・行単位にばらばらに認識されるので、
    #     まとまりごとに PDF から文字を座標付きで取り直してコードブロックにする
    by_page: dict[int, list] = {}
    for it in items:
        by_page.setdefault(it.prov[0].page_no, []).append(it)

    block_no = 0
    for page_no, page_items in by_page.items():
        code = [it for it in page_items if _is_code(it)]
        if not code:
            continue
        others = [_bbox(it) for it in page_items if not _is_code(it)]

        # 上から順に並べ、コード以外の要素が間に挟まる所でブロックを分ける
        code.sort(key=lambda it: _bbox(it)[2])
        blocks: list[list] = []
        for it in code:
            if blocks:
                prev_b = max(_bbox(x)[4] for x in blocks[-1])
                cur_t = _bbox(it)[2]
                if not any(prev_b <= (o[2] + o[4]) / 2 <= cur_t for o in others):
                    blocks[-1].append(it)
                    continue
            blocks.append([it])

        for blk in blocks:
            text_lines = _extract_code(pdf[page_no - 1], blk)
            # 1 行だけのコメントなどは誤検出の可能性が高いので触らない
            # (typedef と関数のプロトタイプ宣言だけは対象にする)
            single_ok = text_lines and re.match(r"^(typedef\b|.*\w\s*\(.*\)\s*;)", text_lines[0])
            if len(text_lines) < 2 and not single_ok:
                continue
            text_lines, trailing, headers = _split_trailing_text(text_lines, page_items, blk)
            key = f"@@CODE{block_no}@@"
            block_no += 1
            replacements[key] = "```c\n" + "\n".join(text_lines) + "\n```"
            # 読み順で最初の要素をブロックの位置にする
            order = sorted(blk, key=items.index)
            for x in blk:
                _set_text(x, key if x is order[0] else DROP)
            # 切り離した本文は、読み順で最後の要素の直後に別の要素として置く (間にある見出しより後ろになる)
            if trailing:
                tkey = f"@@TEXT{len(items) + block_no}@@"
                replacements[tkey] = "\n\n".join(trailing)
                doc.insert_text(sibling=order[-1], label=DocItemLabel.TEXT, text=tkey,
                                prov=order[-1].prov[0], after=True)
                # 本文の見出し (《データ詳細》) がコードより前に並べられていたら、本文の直前へ移す
                for h in headers:
                    if items.index(h) < items.index(order[0]):
                        doc.insert_text(sibling=order[-1], label=DocItemLabel.TEXT, text=h.text,
                                        prov=h.prov[0], after=True)
                        h.text = DROP
    return replacements


def _split_trailing_text(lines: list[str], page_items, blk) -> tuple[list[str], list[str], list]:
    """コード要素の範囲が後ろの本文まで覆っていた場合に、構造体の閉じ括弧より後ろの本文を切り離す。

    切り離した行のうち、ページ内の他の要素 (《データ詳細》 などの見出し) の一部である行は本文から除き、
    その要素を 3 つ目の戻り値で返す。
    """
    last_close = max((i for i, l in enumerate(lines) if re.match(r"^\}\s*[\w\s,*]*;", l)), default=None)
    if last_close is None:
        return lines, [], []
    rest = lines[last_close + 1:]
    if not rest or any(CODE_RE.match(l) for l in rest):
        return lines, [], []
    others = [x for x in page_items if x not in blk and getattr(x, "text", "")]
    trailing, headers = [], []
    for l in rest:
        hit = next((x for x in others if _norm(l) in _norm(x.text)), None)
        if hit is None:
            trailing.append(l.strip())
        elif hit not in headers:
            headers.append(hit)
    return lines[: last_close + 1], trailing, headers


_toc_keys: set[str] = set()


def _is_code(item) -> bool:
    label = item.label.value
    if label == "code":
        return True
    if label == "table":
        cells = [c.text for c in item.data.table_cells]
        if any("typedef" in c for c in cells):
            return True
        # 構造体のメンバー宣言 (xxx; や /* */) を含むセルが多い表は、表に誤認識されたコード
        return sum(bool(re.search(r";|/\*|\*/", c)) for c in cells) >= 3
    if label not in ("text", "section_header", "list_item", "paragraph"):
        return False
    if _norm(item.text or "") in _toc_keys:
        return False  # 目次にある見出し (関数名など) はコードの断片に見えても対象外
    return bool(CODE_RE.match(item.text or ""))


def _set_text(item, text: str) -> None:
    if item.label.value == "table":
        item.data.table_cells = [
            TableCell(text=text, start_row_offset_idx=0, end_row_offset_idx=1,
                      start_col_offset_idx=0, end_col_offset_idx=1)
        ]
        item.data.num_rows = item.data.num_cols = 1
    else:
        item.text = text


def _width(s: str) -> int:
    """等幅フォントでの表示幅 (全角は 2)。"""
    return sum(2 if unicodedata.east_asian_width(ch) in "WF" else 1 for ch in s)


def _extract_code(page, blk) -> list[str]:
    """要素群を囲む範囲のテキストを PDF から取り直し、インデントとコメント位置を揃える。

    PDF によっては 1 文字ずつ (濁点や下線まで別々に) 配置されていて座標から組み立てると崩れるので、
    行の中身は pdfium の範囲指定テキスト取得に任せる。
    """
    boxes = [x.prov[0].bbox for x in blk]  # 左下原点
    l, r = min(b.l for b in boxes) - 1, max(b.r for b in boxes) + 1
    bottom, top = min(b.b for b in boxes) - 1, max(b.t for b in boxes) + 1
    text = page.get_textpage().get_text_bounded(left=l, bottom=bottom, right=r, top=top)

    rows = []
    for raw in text.splitlines():
        raw = re.sub(r"\s+", " ", raw).strip()
        if not raw:
            continue
        # 罫線コメント (/*****/) 以外は、コード部分とコメント部分に分ける
        pos = raw.find("/*")
        if pos > 0 and not raw.startswith("/*"):
            rows.append((raw[:pos].rstrip(), raw[pos:]))
        else:
            rows.append((raw, ""))

    def indent(code: str) -> str:
        top_level = re.match(r"^(typedef|struct|union|enum|#|}|/\*)", code)
        return code if top_level else "    " + code

    width = max((_width(indent(c)) for c, cm in rows if cm), default=0)
    width = min(width, 40)
    lines = []
    for code, comment in rows:
        code = indent(code)
        if comment:
            code += " " * max(width - _width(code), 0) + "  " + comment
        lines.append(code)
    return lines


# ---------------------------------------------------------------------------
# 3. Markdown テキストの後処理
# ---------------------------------------------------------------------------
JP = r"\u3000-\u303f\u3040-\u30ff\u4e00-\u9fff\uff00-\uffef"
SPACED_JP_RE = re.compile(rf"(?<=[{JP}]) (?=[{JP}])")
HEADING_RE = re.compile(r"^#{1,6}\s+(.*)$")
NUMBERED_RE = re.compile(r"^(\d+(?:-\d+)*)[.．]\s*\S")
BRACKET_RE = re.compile(r"^(《[^》]+》)\s*(.*)$")


LABELS = {"機能", "プロトタイプ", "引数", "戻り値", "解説", "備考", "使用例"}


def parse_pdf_toc(pdf: pdfium.PdfDocument, max_pages: int = 6) -> list[str]:
    """PDF の目次ページ (点線 + ページ番号の行) から、見出しの題名を順に取り出す。"""
    titles: list[str] = []
    pending = ""
    for i in range(min(max_pages, len(pdf))):
        for line in pdf[i].get_textpage().get_text_range().splitlines():
            line = line.strip()
            m = re.match(r"^(.*?)\s*[…・.]{3,}\s*\d+$", line)
            if m:
                # 題名が長くて 2 行に折り返した項目 (2 行目が [REQ_... で始まる) は前の行とつなげる
                title = m.group(1)
                if pending and title.startswith("["):
                    title = pending + " " + title
                titles.append(title)
                pending = ""
            else:
                pending = line
    return titles


def _norm(s: str) -> str:
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", s))


def promote_toc_headings(lines: list[str], toc: list[str]) -> list[str]:
    """目次にあるのに見出しとして認識されなかった行 (9-1. 通信 / 関数名など) を見出しにする。

    目次の順に本文を前から照合する。目次と本文が食い違っている項目は読み飛ばす。
    番号のない項目は、直前の番号付き見出しの 1 段下にする。
    """
    if not toc:
        return lines
    lines = "\n".join(lines).split("\n")  # 複数行の要素 (コードブロック) を 1 行ずつに分ける
    keys = [_norm(t) for t in toc]
    j = 0
    level = 2
    in_code = False
    start = next((i for i, l in enumerate(lines) if l.startswith("## ")), len(lines))
    for i in range(start, len(lines)):
        line = lines[i]
        if line.startswith("```"):
            in_code = not in_code
        if in_code or j >= len(keys):
            continue
        h = HEADING_RE.match(line)
        text = h.group(1) if h else line
        key = _norm(text)
        k = next((k for k in range(j, min(j + 4, len(keys))) if keys[k] == key), None)
        if k is None:
            if h:
                level = len(line) - len(line.lstrip("#"))
            continue
        j = k + 1
        n = NUMBERED_RE.match(text.strip())
        if n:
            level = min(n.group(1).count("-") + 2, 6)
            lines[i] = "#" * level + " " + text.strip()
        elif not h:
            lines[i] = "#" * min(level + 1, 6) + " " + text.strip()
    return lines


def postprocess(md: str, replacements: dict[str, str], toc: list[str]) -> str:
    out: list[str] = []
    first_heading_done = False
    after_table_marker = False
    # docling がコードと認識した要素に入れた本文マーカーは ``` で囲まれているので外す
    md = re.sub(r"^```\n(@@TEXT\d+@@)\n```$", r"\1", md, flags=re.M)
    for line in md.split("\n"):
        # 置き換えた表の区切り行 (|---|) を捨てる
        if after_table_marker and re.match(r"^\|[-:| ]+\|$", line):
            after_table_marker = False
            continue
        after_table_marker = False
        if DROP in line:
            after_table_marker = line.startswith("|")
            continue
        m = re.search(r"@@(TEXT|CODE)\d+@@", line)
        if m:
            after_table_marker = line.startswith("|")
            line = replacements[m.group(0)]
            if m.group(1) == "CODE":
                out.append(line)
                continue

        # 単語の間に入った字間スペース (デ ー タ 名 称) を詰める
        if not line.startswith("```"):
            line = SPACED_JP_RE.sub("", line)
        # 箇条書き記号の重複 (- ・) を外す
        line = re.sub(r"^(\s*)- ・\s*", r"\1- ", line)

        h = HEADING_RE.match(line)
        text = h.group(1).strip() if h else line.strip()

        # 《説明》などの小見出しは太字にし、後ろに続く文字は本文にする
        b = BRACKET_RE.match(text)
        if b and (h or text == line.strip()):
            out.append(f"**{b.group(1)}**")
            if b.group(2):
                out.append("")
                out.append(b.group(2))
            continue

        # 関数仕様の項目名 (機能 / プロトタイプ / 戻り値 など) は太字にする
        if text in LABELS:
            out.append(f"**{text}**")
            continue

        if h:
            n = NUMBERED_RE.match(text)
            if n:
                level = min(n.group(1).count("-") + 2, 6)  # 1.→## 4-1.→### 4-1-6.→####
                out.append("#" * level + " " + text)
            elif not first_heading_done:
                out.append("# " + text)  # 表紙のタイトル
            elif text.startswith("○"):
                out.append(f"**{text}**")
            else:
                out.append(text)  # 番号のない見出しは誤検出とみなして本文に戻す
            first_heading_done = True
            continue
        out.append(line)

    out = promote_toc_headings(out, toc)
    out = rebuild_toc(out)
    text = "\n".join(out)
    # docling がコードと認識した要素は ``` で囲まれて出力されるので、二重の囲みを外す
    text = re.sub(r"^```\n(```c\n.*?\n```)\n```$", r"\1", text, flags=re.S | re.M)
    # 中身をコードブロックへ吸収した後に残った、空の ``` の囲みを消す
    text = re.sub(r"^```\n```\n", "", text, flags=re.M)
    # ページをまたいで分かれたコードブロックを 1 つにつなげる
    text = re.sub(r"^```\n+```c\n", "", text, flags=re.M)
    return re.sub(r"\n{3,}", "\n\n", text).strip() + "\n"


def rebuild_toc(lines: list[str]) -> list[str]:
    """点線とページ番号で組まれた目次は表として崩れるので、本文の見出しから作り直す。"""
    start = next((i for i, l in enumerate(lines) if "目次" in l and len(l) < 20), None)
    if start is None:
        return lines
    end = next((i for i in range(start + 1, len(lines)) if lines[i].startswith("## ")), None)
    if end is None:
        return lines
    toc = []
    for l in lines[end:]:
        m = re.match(r"^(#{2,4}) (.*)$", l)
        if m:
            toc.append("  " * (len(m.group(1)) - 2) + "- " + m.group(2))
    return lines[: start + 1] + [""] + toc + [""] + lines[end:]


# ---------------------------------------------------------------------------
# 4. 図の説明文の差し込み
# ---------------------------------------------------------------------------
IMAGE_RE = re.compile(r"^!\[Image\]\((?P<path>[^)]*image_\d+_(?P<hash>[0-9a-f]{8})[0-9a-f]*\.png)\)$", re.M)


def apply_figures(md: str, figures_path: Path) -> str:
    """図の内容を書き起こした説明文 (<出力名>/figures.yaml) を、各図の直下に差し込む。

    画像だけでは後から読む AI が内容を理解できないため。キーは画像ファイル名に含まれる
    ハッシュ値の先頭 8 桁で、PDF を変換し直しても同じ図には同じキーが付く。
      <hash>:
        alt: 図の短い題名
        text: |            # 省略可 (装飾的な図など)
          図の内容 (Markdown)
    """
    if not figures_path.exists():
        print(f"no figure descriptions: {figures_path}")
        return md
    # キーのハッシュ値が数字だけ (00256120 など) の時に数値として読まれないよう、全て文字列で読む
    figures = yaml.load(figures_path.read_text(encoding="utf-8"), Loader=yaml.BaseLoader) or {}
    missing = []

    def repl(m: re.Match) -> str:
        fig = figures.get(m.group("hash"))
        if not fig:
            missing.append(m.group("hash"))
            return m.group(0)
        s = f"![{fig['alt']}]({m.group('path')})"
        if fig.get("text"):
            body = "\n".join(f"> {l}" if l else ">" for l in fig["text"].strip().split("\n"))
            s += f"\n\n> **図の内容**\n>\n{body}"
        return s

    md = IMAGE_RE.sub(repl, md)
    if missing:
        print(f"figures without description ({len(missing)}): {', '.join(missing)}")
    return md


# ---------------------------------------------------------------------------
# 5. 表の手修正の差し込み
# ---------------------------------------------------------------------------
def apply_table_fixes(md: str, fixes_path: Path) -> str:
    """結合セルなどで崩れた表を、手で直した表 (<出力名>/tables.yaml) に差し替える。

    after に書いた文字列を含む行の後ろにある、最初の表を table の内容に置き換える。
      - after: 表の直前にある本文の一部 (一意になるもの)
        table: |
          | 見出し | ... |
          |---|---|
          ...
    """
    if not fixes_path.exists():
        return md
    lines = md.split("\n")
    for fix in yaml.load(fixes_path.read_text(encoding="utf-8"), Loader=yaml.BaseLoader) or []:
        anchor = next((i for i, l in enumerate(lines) if fix["after"] in l), None)
        if anchor is None:
            print(f"table fix anchor not found: {fix['after']}")
            continue
        start = next((i for i in range(anchor + 1, len(lines)) if lines[i].startswith("|")), None)
        if start is None:
            print(f"table not found after: {fix['after']}")
            continue
        end = start
        while end < len(lines) and lines[end].startswith("|"):
            end += 1
        lines[start:end] = fix["table"].strip().split("\n")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawTextHelpFormatter)
    ap.add_argument("pdf", type=Path)
    ap.add_argument("-o", "--out", type=Path, help="出力先 (省略時は docs/03_reference/<製品>/converted/<名前>.md)")
    ap.add_argument("--reconvert", action="store_true", help="キャッシュを使わず解析し直す")
    args = ap.parse_args()

    out = args.out or default_out(args.pdf)
    doc = load_or_analyze(args.pdf, args.reconvert)
    pdf = pdfium.PdfDocument(args.pdf)
    toc = parse_pdf_toc(pdf)
    replacements = patch_document(doc, pdf, toc)

    # 付属ファイルは <出力名>/ にまとめる (図の PNG は images/、手修正は figures.yaml / tables.yaml)
    asset_dir = out.parent / out.stem
    image_dir = asset_dir / "images"
    image_dir.mkdir(parents=True, exist_ok=True)
    for old in image_dir.glob("*.png"):
        old.unlink()  # 古い図が残らないように消す (フォルダごと消すと開いている時に失敗するため中身だけ)
    doc = doc._with_pictures_refs(image_dir=image_dir, page_no=None, reference_path=out.parent)
    md = doc.export_to_markdown(
        image_mode=ImageRefMode.REFERENCED, escape_underscores=False, escape_html=False
    )

    md = postprocess(md, replacements, toc)
    md = apply_figures(md, asset_dir / "figures.yaml")
    md = apply_table_fixes(md, asset_dir / "tables.yaml")
    out.write_text(md, encoding="utf-8")
    print(f"wrote {out}")


if __name__ == "__main__":
    main()
