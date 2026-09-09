# -*- coding: utf-8 -*-
"""商品マスタ(k/51) の CS別品番 と、在庫系(S在庫) の CS品番 の表記不一致を全件洗い出す。

出力機能は、商品マスタの CS別品番 で在庫・在庫分析を絞り込んでから集計する
（main.js 49行目 → 63行目）。そのため両者の文字列が一致しない SKU は、
在庫系に正しい値があっても集計から丸ごと落ちる。

  ・一部の SKU だけ一致しない → 行は出るが在庫・販売数などが小さくなる
  ・全部の SKU が一致しない   → 在庫合計0 で除外され、行ごと消える（気づけない）

このスクリプトは在庫系を基準に「落ちている SKU」を数え、なぜ落ちたのかを分類する。

  python CS品番の表記不一致を洗い出す.py
  python CS品番の表記不一致を洗い出す.py --在庫 "...\0803\S在庫_S20260803.csv"

kintone を読むので Windows 側で実行する（Cowork のシェルは cybozu.com に出られない）。
前提: pip install requests openpyxl
"""

import argparse
import csv
import os
import sys
import time
import unicodedata

import openpyxl
import requests
from openpyxl.styles import Alignment, Font, PatternFill

HERE   = os.path.dirname(os.path.abspath(__file__))
REPO   = os.path.dirname(HERE)
BASE   = os.path.dirname(REPO)
SERVER = os.path.dirname(BASE)

CONFIG_PATH = os.path.join(
    SERVER, "260318ー登録エクセルkintone化", "納品書作成", "config", "kintone_config.xlsx")

_期間 = os.path.join(BASE, "260904-kintone動作確認", "セール情報設定リスト 二期間分",
                     "セール情報設定リスト202608089～202608219")
在庫_既定 = os.path.join(_期間, "0803", "S在庫_S20260803.csv")
出力_既定 = os.path.join(_期間, "kintone出力", "0908", "CS品番_表記不一致.xlsx")

CELL_SUBDOMAIN = "C3"
EXPECT_LABEL   = "商品マスタ"
EXPECT_APP_ID  = 51
FIELDS = ["ブランド品番", "CS別品番"]
SLEEP_SEC = 0.05

FILL_NG   = PatternFill("solid", fgColor="FFC7CE")
FILL_WARN = PatternFill("solid", fgColor="FFEB9C")


def load_config(path, label=EXPECT_LABEL):
    """設定ブックの A列から label を探し、同じ行の C列＝トークン / 次の行の C列＝アプリID を読む。

    キントーン添付DL.py と同じ並び（A65「セール返答表」/ C65 トークン / C66 アプリID）を
    前提に、行がずれても拾えるよう固定セルではなく検索で解決する。
    """
    if not os.path.exists(path):
        sys.exit("設定ファイルが見つかりません: %s\n--config でパスを指定してください。" % path)
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["設定"] if "設定" in wb.sheetnames else wb.active

    行 = None
    for r in range(1, ws.max_row + 1):
        if str(ws.cell(r, 1).value or "").strip() == label:
            行 = r
            break
    if 行 is None:
        sys.exit("設定ファイルの A列に「%s」が見つかりません: %s" % (label, path))

    subdomain = str(ws[CELL_SUBDOMAIN].value or "").strip()
    token     = str(ws.cell(行, 3).value or "").strip()
    app_id    = ws.cell(行 + 1, 3).value
    if not subdomain or not token or not app_id:
        sys.exit("「%s」(A%d) のサブドメイン / トークン(C%d) / アプリID(C%d) のいずれかが空です。"
                 % (label, 行, 行, 行 + 1))
    app_id = int(app_id)
    print("設定: %s / A%d「%s」 アプリ%d" % (subdomain, 行, label, app_id))
    if app_id != EXPECT_APP_ID:
        print("  ※ アプリIDが想定(%d)と違います。設定ブックの並びを確認してください。" % EXPECT_APP_ID)
    return subdomain, token, app_id


def get_records(base_url, headers, app_id):
    """$id を進めながら全件取る（offset を使わない = 5,000件の壁が無い）。"""
    records, last_id, t0 = [], 0, time.time()
    while True:
        q = "$id > %d order by $id asc limit 500" % last_id
        res = requests.get(base_url + "/k/v1/records.json", headers=headers,
                           params={"app": app_id, "query": q, "fields": FIELDS})
        if res.status_code != 200:
            sys.exit("レコード取得に失敗しました [%d]\n%s" % (res.status_code, res.text[:500]))
        batch = res.json().get("records", [])
        if not batch:
            break
        records.extend(batch)
        last_id = int(batch[-1]["$id"]["value"])
        if len(records) % 10000 < 500:
            print("  %d件 (%.0f秒)" % (len(records), time.time() - t0))
        time.sleep(SLEEP_SEC)
    print("商品マスタ %d件 取得 (%.0f秒)" % (len(records), time.time() - t0))
    return records


def 角括弧を外す(s):
    """【入荷禁止】のような運用マーカーを取り除く。"""
    out, 深さ = [], 0
    for ch in s:
        if ch in "【［[":
            深さ += 1
        elif ch in "】］]":
            深さ = max(0, 深さ - 1)
        elif 深さ == 0:
            out.append(ch)
    return "".join(out)


def 正規化(s):
    """比較用のキー。マーカー・改行・空白・全角半角の違いを吸収する。"""
    s = 角括弧を外す(str(s or ""))
    s = unicodedata.normalize("NFKC", s)
    s = "".join(s.split())
    return s.upper()


def 分類(元):
    """正規化で一致したとき、何が原因だったかを短く言う。"""
    理由 = []
    if any(c in 元 for c in "【［["):
        理由.append("角括弧マーカー")
    if "\n" in 元 or "\r" in 元:
        理由.append("改行")
    if 元 != 元.strip() or "　" in 元 or " " in 元:
        理由.append("空白")
    if unicodedata.normalize("NFKC", 元) != 元:
        理由.append("全角文字")
    return " / ".join(理由) or "その他の表記差"


def 前方一致で探す(cs, 候補):
    """正規化しても一致しないとき、前半が共通する相手を探す（サイズ表記のズレなど）。"""
    seg = 正規化(cs).split("-")
    best, best_n = None, 0
    for k in 候補:
        ks = 正規化(k).split("-")
        n = 0
        for a, b in zip(seg, ks):
            if a != b:
                break
            n += 1
        if n >= 3 and n > best_n:
            best, best_n = k, n
    return best


def read_在庫(path):
    if not os.path.exists(path):
        sys.exit("在庫ファイルが見つかりません: %s" % path)
    ブランド別 = {}
    with open(path, encoding="cp932", newline="", errors="replace") as f:
        r = csv.reader(f)
        h = [x.strip() for x in next(r)]
        bi, ci, qi = h.index("ブランド品番"), h.index("CS品番"), h.index("在庫数")
        ni = h.index("商品名") if "商品名" in h else None
        pi = h.index("親カテゴリー") if "親カテゴリー" in h else None
        for row in r:
            b = row[bi].strip()
            if not b:
                continue
            d = ブランド別.setdefault(b, {"sku": {}, "商品名": "", "親カテゴリー": ""})
            d["sku"][row[ci].strip()] = int(float(row[qi] or 0))
            if ni is not None and not d["商品名"]:
                d["商品名"] = row[ni].strip()
            if pi is not None and not d["親カテゴリー"]:
                d["親カテゴリー"] = row[pi].strip()
    print("S在庫 %dブランド品番 / %dSKU" % (
        len(ブランド別), sum(len(v["sku"]) for v in ブランド別.values())))
    return ブランド別


def 見出し(ws, 行, 列名):
    for i, v in enumerate(列名, 1):
        c = ws.cell(行, i, v)
        c.font = Font(bold=True)
        c.alignment = Alignment(vertical="center")
    ws.freeze_panes = ws.cell(行 + 1, 1)


def 書き出す(path, 明細, ブランド集計, 余り, 在庫元):
    wb = openpyxl.Workbook()

    ws = wb.active
    ws.title = "1_サマリ"
    ws["A1"] = "商品マスタ(k/51) と 在庫系 の CS品番 表記不一致"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "在庫元: %s" % 在庫元
    ws["A3"] = "作成: %s" % time.strftime("%Y/%m/%d %H:%M")
    全滅 = [b for b in ブランド集計 if b["一致"] == 0]
    一部 = [b for b in ブランド集計 if b["一致"] > 0]
    rows = [
        ("■ 影響のあるブランド品番", len(ブランド集計)),
        ("　うち 全滅（在庫合計0で行ごと消える）", len(全滅)),
        ("　うち 一部欠け（値が小さくなる）", len(一部)),
        ("", ""),
        ("落ちているSKU数", len(明細)),
        ("落ちている在庫数の合計", sum(m["在庫数"] for m in 明細)),
    ]
    r = 5
    for k, v in rows:
        ws.cell(r, 1, k)
        ws.cell(r, 2, v)
        r += 1
    ws.cell(r + 1, 1, "分類の内訳").font = Font(bold=True)
    r += 2
    内訳 = {}
    for m in 明細:
        内訳[m["分類"]] = 内訳.get(m["分類"], 0) + 1
    for k, v in sorted(内訳.items(), key=lambda x: -x[1]):
        ws.cell(r, 1, "　" + k)
        ws.cell(r, 2, v)
        r += 1
    ws.column_dimensions["A"].width = 46
    ws.column_dimensions["B"].width = 14

    ws = wb.create_sheet("2_落ちているSKU")
    ws["A1"] = "在庫系にあるのに、商品マスタの CS別品番 と一致しないSKU（＝集計から落ちている）"
    見出し(ws, 3, ["ブランド品番", "親カテゴリー", "商品名", "在庫系のCS品番", "在庫数",
                   "分類", "商品マスタ側の該当文字列"])
    for i, m in enumerate(明細, 4):
        for j, v in enumerate([m["ブランド品番"], m["親カテゴリー"], m["商品名"], m["CS品番"],
                               m["在庫数"], m["分類"], m["相手"]], 1):
            ws.cell(i, j, v)
        if m["全滅"]:
            for j in range(1, 8):
                ws.cell(i, j).fill = FILL_NG
    for col, w in zip("ABCDEFG", (34, 20, 34, 34, 8, 22, 34)):
        ws.column_dimensions[col].width = w
    ws.auto_filter.ref = "A3:G%d" % max(4, len(明細) + 3)

    ws = wb.create_sheet("3_ブランド品番別")
    ws["A1"] = "赤=全滅（行ごと消える） 黄=一部欠け（値が小さくなる）"
    見出し(ws, 3, ["ブランド品番", "親カテゴリー", "商品名", "S在庫SKU数", "一致SKU数",
                   "欠けSKU数", "S在庫の在庫合計", "一致分の在庫", "落ちた在庫", "状態"])
    for i, b in enumerate(ブランド集計, 4):
        for j, v in enumerate([b["ブランド品番"], b["親カテゴリー"], b["商品名"], b["SKU数"],
                               b["一致"], b["欠け"], b["在庫合計"], b["一致在庫"],
                               b["在庫合計"] - b["一致在庫"], b["状態"]], 1):
            ws.cell(i, j, v)
        fill = FILL_NG if b["一致"] == 0 else FILL_WARN
        for j in range(1, 11):
            ws.cell(i, j).fill = fill
    for col, w in zip("ABCDEFGHIJ", (34, 20, 34, 11, 11, 11, 14, 12, 12, 18)):
        ws.column_dimensions[col].width = w
    ws.auto_filter.ref = "A3:J%d" % max(4, len(ブランド集計) + 3)

    ws = wb.create_sheet("4_商品マスタ側のみ")
    ws["A1"] = "商品マスタにあるが S在庫に同じ文字列が無い CS別品番（参考。在庫が無いだけの場合も含む）"
    見出し(ws, 3, ["ブランド品番", "商品マスタのCS別品番"])
    for i, (b, cs) in enumerate(余り, 4):
        ws.cell(i, 1, b)
        ws.cell(i, 2, cs)
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 40
    ws.auto_filter.ref = "A3:B%d" % max(4, len(余り) + 3)

    wb.save(path)
    print("書き出しました: %s" % path)


def main():
    ap = argparse.ArgumentParser(description="商品マスタと在庫系のCS品番の表記不一致を洗い出す")
    ap.add_argument("--在庫", default=在庫_既定, help="S在庫のCSV")
    ap.add_argument("--出力", default=出力_既定, help="出力するxlsx")
    ap.add_argument("--config", default=CONFIG_PATH, help="kintone_config.xlsx")
    ap.add_argument("--アプリ", type=int, help="商品マスタのアプリID（設定ブックを使わないとき）")
    ap.add_argument("--トークン", help="APIトークン（設定ブックを使わないとき）")
    ap.add_argument("--サブドメイン", help="サブドメイン（設定ブックを使わないとき）")
    args = ap.parse_args()

    if args.トークン and args.アプリ and args.サブドメイン:
        subdomain, token, app_id = args.サブドメイン, args.トークン, args.アプリ
    else:
        subdomain, token, app_id = load_config(args.config)
        app_id = args.アプリ or app_id

    在庫 = read_在庫(args.在庫)

    base_url = "https://%s.cybozu.com" % subdomain
    headers  = {"X-Cybozu-API-Token": token}
    print("商品マスタ(k/%d) を取得します..." % app_id)
    recs = get_records(base_url, headers, app_id)

    マスタ = {}
    for r in recs:
        b  = str(r.get("ブランド品番", {}).get("value") or "").strip()
        c  = str(r.get("CS別品番", {}).get("value") or "").strip()
        if b and c:
            マスタ.setdefault(b, set()).add(c)

    明細, ブランド集計, 余り = [], [], []
    for b, d in sorted(在庫.items()):
        K = マスタ.get(b, set())
        Kn = {正規化(k): k for k in K}
        sku = d["sku"]
        一致 = [c for c in sku if c in K]
        欠け = [c for c in sku if c not in K]
        if 欠け:
            全滅 = len(一致) == 0
            for c in sorted(欠け):
                n = 正規化(c)
                if n in Kn:
                    相手, ぶんるい = Kn[n], 分類(Kn[n])
                else:
                    相手 = 前方一致で探す(c, K)
                    ぶんるい = "後半の表記違い（サイズ等）" if 相手 else "商品マスタに該当なし"
                明細.append({
                    "ブランド品番": b, "親カテゴリー": d["親カテゴリー"], "商品名": d["商品名"],
                    "CS品番": c, "在庫数": sku[c], "分類": ぶんるい,
                    "相手": 相手 or "", "全滅": 全滅})
            ブランド集計.append({
                "ブランド品番": b, "親カテゴリー": d["親カテゴリー"], "商品名": d["商品名"],
                "SKU数": len(sku), "一致": len(一致), "欠け": len(欠け),
                "在庫合計": sum(sku.values()),
                "一致在庫": sum(sku[c] for c in 一致),
                "状態": "全滅（行が消える）" if 全滅 else "一部欠け（値が小さい）"})
        for k in sorted(K):
            if k not in sku:
                余り.append((b, k))

    print("影響のあるブランド品番 %d（うち全滅 %d） / 落ちているSKU %d / 落ちている在庫 %d"
          % (len(ブランド集計), sum(1 for x in ブランド集計 if x["一致"] == 0),
             len(明細), sum(m["在庫数"] for m in 明細)))
    os.makedirs(os.path.dirname(args.出力), exist_ok=True)
    書き出す(args.出力, 明細, ブランド集計, 余り, args.在庫)


if __name__ == "__main__":
    main()
