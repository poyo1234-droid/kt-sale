"""検証スクリプト共通の設定とユーティリティ。

すべてのパスはこのリポジトリ（エクセル開発フォルダ）からの相対で解決する。
"""
import os

# このファイル = <BASE>\エクセル開発\検証スクリプト\_共通.py
REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))   # エクセル開発
BASE = os.path.dirname(REPO)                                        # 251016-SALEデータ一式-今野様

# 出力マクロが生成する TXT のシート名（＝ファイル名の接頭辞）
OUTPUT_SHEETS = [
    "前回と同じ",
    "記入以外前回と同じ",
    "返信あり",
    "追加セール",
    "通年-前回と同じ",
    "通年-記入以外前回と同じ",
    "通年-返信あり",
]

# TXT の列（Module3 は C〜G列を出力する）
COLUMNS = ["ブランド品番", "変更後セール価格", "開始日", "終了日", "プロパー価格"]

# 担当者が実際に ZOZO へアップロードした TXT（260731 回）
TXT_TANTOUSHA = os.path.join(
    BASE, "260731-org",
    "セール情報設定リスト202608019～202608079",
    "セール情報設定リスト202608019～202608079",
    "TXT_output今野様")

# 検証用にこちらで出力した TXT
TXT_KENSHOU = os.path.join(REPO, "_テスト用", "TXT_output")

# 出力マクロを実行する対象ブック（コピー元）
# 旧版（v50〜v52）は OLD\ に移してあるため、そちらも探す。
def book(version="v53"):
    for d in (REPO, os.path.join(REPO, "OLD")):
        p = os.path.join(d, f"sale処理-{version}.xlsm")
        if os.path.exists(p):
            return p
    return os.path.join(REPO, f"sale処理-{version}.xlsm")   # 見つからない場合は既定の場所を返す


def read_txt(path):
    """TXT を cp932 でデコードし、空行を除いた行のリストを返す。"""
    text = open(path, "rb").read().decode("cp932")
    return [r for r in text.split("\r\n") if r != ""]


# --- キントーン突合（2026-09-01 追加）------------------------------------
# 旧システム（sale処理-vNN.xlsm）が出力したセール価格設定リストと、
# キントーンが出力した同種のリストを突き合わせるためのパス。
# 回ごとに変わるため、キントーン突合.py の --kintone / --旧 で上書きできる。

_260220 = os.path.join(BASE, "260220-セール情報設定リスト202602219～202603069")

# キントーンが出力したセール価格設定リスト（Drive からローカルへ DL したもの）
KINTONE_LIST = os.path.join(_260220, "260901kintone突合")

# 旧システムが出力し、実際にメーカーへ送信したセール価格設定リスト
KYUU_LIST = os.path.join(_260220, "送信")
