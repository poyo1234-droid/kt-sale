"""ブックのコピーを開き、出力マクロを実行して TXT を生成する。

  python TXT出力を実行.py [バージョン]      例: python TXT出力を実行.py v53

本体ブックは開かない。`_テスト用\\sale処理-<版>_TXT出力.xlsm` にコピーしてから実行し、
保存せずに閉じるため、本体・git ともに無変更のまま。

出力先は `_テスト用\\TXT_output\\`（実行のたびに作り直す）。

前提:
  - pywin32 が必要（`pip install pywin32`）
  - Excel の「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」が有効であること
    （HKCU\\Software\\Microsoft\\Office\\16.0\\Excel\\Security\\AccessVBOM = 1）
  - クエリの自動更新は全て OFF のため、開いても再計算は走らない

実行するのは Module3 の `ExportColumnsAEInTXTChunks_ANSI`（C〜G列 / 500行分割 / ANSI）。
自動実行のため MsgBox 行だけコメントアウトする。出力ロジックには手を加えない。
"""
import os, sys, shutil, time
import win32com.client as win32
from _共通 import REPO, OUTPUT_SHEETS, book

MACRO = "ExportColumnsAEInTXTChunks_ANSI"


def main(version="v53"):
    src = book(version)
    if not os.path.exists(src):
        raise SystemExit(f"ブックが無い: {src}")

    workdir = os.path.join(REPO, "_テスト用")
    dst = os.path.join(workdir, f"sale処理-{version}_TXT出力.xlsm")
    out = os.path.join(workdir, "TXT_output")

    os.makedirs(workdir, exist_ok=True)
    if os.path.isdir(out):
        shutil.rmtree(out)
    shutil.copy2(src, dst)
    print(f"コピー: {dst}")

    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    xl.AskToUpdateLinks = False
    try:
        wb = xl.Workbooks.Open(dst, UpdateLinks=0)

        # 自動実行のため MsgBox のみ無効化（出力ロジックは変更しない）
        cm = wb.VBProject.VBComponents("Module3").CodeModule
        n = 0
        for i in range(cm.CountOfLines, 0, -1):
            s = cm.Lines(i, 1).strip()
            if s.startswith("MsgBox"):
                cm.ReplaceLine(i, "    ' [自動実行のため無効化] " + s)
                n += 1
        print(f"MsgBox {n} 行を無効化")

        for name in OUTPUT_SHEETS:
            wb.Worksheets(name).Activate()
            t0 = time.time()
            xl.Run(MACRO)
            print(f"  出力: {name}  ({time.time() - t0:.1f}s)")

        wb.Close(SaveChanges=False)
    finally:
        xl.Quit()

    files = sorted(os.listdir(out))
    print(f"\n{len(files)} ファイルを生成 -> {out}")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "v53")
