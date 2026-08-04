"""担当者の本番 TXT と、こちらで出力した TXT を突き合わせる。

  python TXTを比較.py

3段階で照合する。

  1. バイト単位  … "rb" で読んだ生のバイト列を比較。改行コード（CRLF）・タブ・
                    末尾の空白まで含めて differences を検出する
  2. 文字単位    … cp932（ANSI）でデコードし、zip() で1文字ずつ突き合わせる。
                    長さも併せて見る（zip は短い方で止まるため）
  3. 列単位      … タブで分割し、C〜G列それぞれを列ごとに比較する

最後に全ファイルを連結した単一文字列の SHA256 も照合し、ファイル単位の
照合漏れが無いことを担保する。
"""
import os, re, glob, hashlib, collections
from _共通 import TXT_TANTOUSHA, TXT_KENSHOU, COLUMNS, read_txt

A, B = TXT_TANTOUSHA, TXT_KENSHOU


def main():
    if not os.path.isdir(B):
        raise SystemExit(f"比較対象が無い。先に TXT出力を実行.py を動かすこと: {B}")

    names_a = {os.path.basename(p) for p in glob.glob(os.path.join(A, "*.txt"))}
    names_b = {os.path.basename(p) for p in glob.glob(os.path.join(B, "*.txt"))}
    common = sorted(names_a & names_b)

    # ---- 1・2. バイト単位 / 文字単位 ----
    print("=== バイト単位・文字単位 ===")
    print(f"{'ファイル':<32}{'バイト':>9}{'文字':>8}  バイト  文字")
    ng = []
    tot_b = tot_c = 0
    for n in common:
        ba = open(os.path.join(A, n), "rb").read()
        bb = open(os.path.join(B, n), "rb").read()
        sa, sb = ba.decode("cp932"), bb.decode("cp932")

        byte_ok = ba == bb
        char_ok = len(sa) == len(sb) and all(x == y for x, y in zip(sa, sb))
        if not (byte_ok and char_ok):
            ng.append(n)
            for i, (x, y) in enumerate(zip(sa, sb)):
                if x != y:
                    print(f"  !! {n} 位置 {i}: 担当者={x!r} / 検証={y!r}")
                    print(f"     担当者 {sa[max(0,i-40):i+40]!r}")
                    print(f"     検証   {sb[max(0,i-40):i+40]!r}")
                    break
            else:
                print(f"  !! {n} 長さ違い: 担当者 {len(sa)} / 検証 {len(sb)} 文字")

        tot_b += len(bb)
        tot_c += len(sb)
        print(f"{n:<32}{len(bb):>9}{len(sb):>8}  "
              f"{'一致' if byte_ok else '相違'}  {'一致' if char_ok else '相違'}")

    print(f"\n{len(common)} ファイル / {tot_b:,} バイト / {tot_c:,} 文字 を照合")
    print(f"相違のあったファイル: {len(ng)} {ng if ng else ''}")
    print(f"担当者側のみ: {sorted(names_a - names_b)}")
    print(f"検証側のみ  : {sorted(names_b - names_a)}")

    # ---- 3. 列単位 ----
    def load(d, key):
        rows = []
        for p in sorted(glob.glob(os.path.join(d, key + "_*.txt"))):
            rows += [r.split("\t") for r in read_txt(p)]
        return rows

    keys = sorted({re.match(r"(.+)_\d+\.txt$", n).group(1)
                   for n in names_b if re.match(r"(.+)_\d+\.txt$", n)})

    print("\n=== 列単位 ===")
    for key in keys:
        ra, rb = load(A, key), load(B, key)
        marks = []
        for c in range(max(len(r) for r in rb)):
            va = [r[c] if c < len(r) else None for r in ra]
            vb = [r[c] if c < len(r) else None for r in rb]
            marks.append(f"{COLUMNS[c]}={'一致' if va == vb else '★相違'}")
        print(f"{key:<22}{len(rb):>6}行  " + "  ".join(marks))

    # ---- 全結合 ----
    cat_a = "".join(open(os.path.join(A, n), "rb").read().decode("cp932") for n in common)
    cat_b = "".join(open(os.path.join(B, n), "rb").read().decode("cp932") for n in common)
    print("\n=== 全ファイル連結 ===")
    print(f"担当者 {len(cat_a):,} 文字 / 検証 {len(cat_b):,} 文字  完全一致: {cat_a == cat_b}")
    print(f"SHA256 担当者: {hashlib.sha256(cat_a.encode('utf-8')).hexdigest()}")
    print(f"SHA256 検証  : {hashlib.sha256(cat_b.encode('utf-8')).hexdigest()}")


if __name__ == "__main__":
    main()
