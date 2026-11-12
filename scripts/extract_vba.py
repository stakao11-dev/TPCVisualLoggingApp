"""Excel ファイル（.xls/.xlt/.xlsm/.xltm）から VBA のソースコードを取り出す。

VBA プロジェクトのパスワードは VBA エディタで開くのを止めるだけで、
ソースコード自体は圧縮して保存されている。それを MS-OVBA の仕様どおりに展開する。

使い方:
    .venv/Scripts/python scripts/extract_vba.py <Excel ファイル> <出力フォルダ>

モジュールごとに <モジュール名>.bas を UTF-8 で出力する。
フォームの見た目（.frx）は取り出さない。
"""
import io
import os
import struct
import sys
import zipfile

import olefile


def decompress(data):
    """MS-OVBA 2.4.1 の圧縮コンテナを展開する。"""
    if not data or data[0] != 1:
        raise ValueError("圧縮コンテナの先頭が 0x01 ではありません")
    out = bytearray()
    pos = 1
    while pos < len(data):
        header = struct.unpack_from("<H", data, pos)[0]
        size = (header & 0x0FFF) + 3
        compressed = header & 0x8000
        chunk = data[pos + 2:pos + size]
        pos += size
        if not compressed:
            out += chunk[:4096]
            continue
        start = len(out)
        i = 0
        while i < len(chunk):
            flags = chunk[i]
            i += 1
            for bit in range(8):
                if i >= len(chunk):
                    break
                if not (flags >> bit) & 1:
                    out.append(chunk[i])
                    i += 1
                    continue
                token = struct.unpack_from("<H", chunk, i)[0]
                i += 2
                bit_count = max((len(out) - start - 1).bit_length(), 4)
                length = (token & (0xFFFF >> bit_count)) + 3
                offset = (token >> (16 - bit_count)) + 1
                for _ in range(length):
                    out.append(out[-offset])
    return bytes(out)


def parse_dir(data):
    """dir ストリームから (モジュール名, ストリーム名, ソースの開始位置) の一覧と文字コードを得る。"""
    modules = []
    codepage = 932
    name = stream = None
    offset = 0
    i = 0
    while i + 6 <= len(data):
        record_id, size = struct.unpack_from("<HI", data, i)
        if record_id == 0x0009:  # PROJECTVERSION はサイズ欄と実際の長さが違う
            i += 12
            continue
        body = data[i + 6:i + 6 + size]
        i += 6 + size
        if record_id == 0x0003:
            codepage = struct.unpack_from("<H", body)[0]
        elif record_id == 0x0019:
            name = body.decode("cp%d" % codepage, "replace")
        elif record_id == 0x001A:
            stream = body.decode("cp%d" % codepage, "replace")
        elif record_id == 0x0031:
            offset = struct.unpack_from("<I", body)[0]
        elif record_id == 0x002B:  # モジュールの終わり
            modules.append((name, stream, offset))
    return modules, codepage


def open_vba_project(path):
    """VBA プロジェクトの OLE と、その中の VBA フォルダまでのパスを返す。"""
    if zipfile.is_zipfile(path):
        with zipfile.ZipFile(path) as z:
            names = [n for n in z.namelist() if n.endswith("vbaProject.bin")]
            if not names:
                raise SystemExit("VBA プロジェクトがありません: %s" % path)
            return olefile.OleFileIO(io.BytesIO(z.read(names[0]))), []
    ole = olefile.OleFileIO(path)
    for prefix in (["_VBA_PROJECT_CUR"], ["Macros"]):
        if ole.exists("/".join(prefix + ["VBA", "dir"])):
            return ole, prefix
    raise SystemExit("VBA プロジェクトがありません: %s" % path)


def main():
    if len(sys.argv) != 3:
        raise SystemExit(__doc__)
    path, outdir = sys.argv[1], sys.argv[2]
    ole, prefix = open_vba_project(path)
    modules, codepage = parse_dir(decompress(ole.openstream(prefix + ["VBA", "dir"]).read()))
    os.makedirs(outdir, exist_ok=True)
    for name, stream, offset in modules:
        raw = ole.openstream(prefix + ["VBA", stream]).read()
        source = decompress(raw[offset:]).decode("cp%d" % codepage, "replace")
        with open(os.path.join(outdir, name + ".bas"), "w", encoding="utf-8", newline="") as f:
            f.write(source)
        print("%-20s %5d 行" % (name, source.count("\n")))


if __name__ == "__main__":
    main()
