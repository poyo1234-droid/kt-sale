# CLAUDE.md

このフォルダで作業する際の前提情報。

## プロジェクト概要

ZOZOTOWN のセール処理を行う Excel マクロブック `sale処理-v50.xlsm` の保守・改修。
メーカーから提出されたセール価格ファイルを Power Query で集約し、ZOZO へアップロードする TSV/TXT を VBA で出力する。

- **VBA**: 12モジュール（出力マクロ、集計、クエリ更新、重複チェック）
- **Power Query**: 約90クエリ（`Section1.m`）
- ブック本体は NAS 上（`T:` = `\\dop\ドキュメント`）

## 現在の状況（2026-08-02 時点）

**「通年-記入以外前回と同じ」シートの親カテゴリ欠落バグ**を調査し、修正仕様まで確定。**実装は未着手**（xlsm には一切変更を加えていない）。

詳細は `仕様書_通年-記入以外前回と同じ_親カテゴリ欠落対応.md` を参照。再開時はまずこれを読むこと。

### 次にやること

1. GitHub リポジトリを作成し、リモート登録（`git remote add origin <URL>` → `git push -u origin main`）
2. 仕様書 5章の M コードを実装（v51 として別ファイル保存してから作業）
3. 仕様書 7章の検証項目を実施

## ファイル構成

| ファイル | 内容 |
|---|---|
| `sale処理-v50.xlsm` | 本体。VBA + Power Query |
| `saleファイルチェックーv9.xlsx` | ファイルチェック用 |
| `仕様書_通年-記入以外前回と同じ_親カテゴリ欠落対応.md` | 調査結果と修正仕様 |

## ブックの中身を調べる方法

xlsm を直接開かなくても、Python + oletools で中身を読める。**調査は必ずこの方法で行い、ブックを書き換えないこと。**

```bash
# 1. xlsm を zip として展開
Copy-Item sale処理-v50.xlsm sale.zip
Expand-Archive sale.zip -DestinationPath unz
```

**VBA の抽出**

```python
from oletools.olevba import VBA_Parser
p = VBA_Parser(r"unz\xl\vbaProject.bin")
for (fn, sn, vn, code) in p.extract_macros():
    print(vn, code)
```

**Power Query（M コード）の抽出**

```python
import re, base64, struct, zipfile, io
raw = open(r"unz\customXml\item1.xml", "rb").read().decode("utf-16")
m = re.search(r">([A-Za-z0-9+/=\s]+)</DataMashup>", raw)
b = base64.b64decode(m.group(1))
ver, plen = struct.unpack("<II", b[0:8])   # 先頭8バイトはヘッダー
z = zipfile.ZipFile(io.BytesIO(b[8:8+plen]))
z.extractall("pq")                          # pq/Formulas/Section1.m が全クエリのソース
```

**シートデータの確認**

`openpyxl.load_workbook(path, data_only=True, read_only=True)` で値を読む。

## 処理の全体像

```
メーカー提出ファイル（フォルダ別）      外部マスタ
  ├ 【記入以外前回と同じ】              ├ salegoods.csv  … 前回のセール価格
  ├ 【返信あり】                        ├ goods.csv      … 商品マスタ（子カテゴリを持つ）
  ├ 【通年セール】                      ├ メーカーマスタ … 親カテゴリ→メーカー名
  └ 【追加セール】                      └ ◎セール返答表  … メーカーごとのセール方針
        │                                      │
        └──→ fxファイル前処理 ←───────────────┘
                    │
              各シート（前処理／出力）
                    │
              VBA で TSV/TXT 出力 → ZOZO へアップロード
```

### 重要な構造

**セール返答表は メーカー / 親カテ / 子カテ の3階層**で指示を出す。各階層に抽出クエリがある（`セール返答表-記入以外-メーカー`、`セール返答表-記入以外-親カテ` など）。

- 非通年側はメーカー階層・親カテ階層を実装済み
- **通年側は階層を一切見ておらず、提出ファイルに出てきた親カテゴリで代用している** ← バグの温床
- 通年／非通年の切り分けは設定シート A17:B42（名前定義 `通年セールメーカー`）

**外部ファイルのパスは設定シートに集約**されている（`Tconfig` テーブル、`Psalegoods_cs` などのパラメータクエリが参照）。パスは別PC（`C:\Users\user\Dropbox\...`）を指しているため、**この環境からは外部ファイルを参照できない**。

## 作業上の注意

- **仕様が固まるまで実装しない。** ユーザーが自分で仕様を検討・決定する
- **回答は簡潔に。** 長い分析や表の羅列は避ける
- 選択肢を出すときは選択UIではなく文章で2〜3個に絞る
- xlsm を変更する場合は必ず**バージョンを上げて別ファイル保存**（v50 → v51）。既存ファイルは上書きしない
- git は NAS 上のため `safe.directory` の例外登録済み（グローバル設定）

## 既知の未対応課題

仕様書 8章に記載。要点のみ。

- 非通年側（`ファイル前処理-編集用-親カテ抽出`）にも同種の欠落がある
- Module3 と Module4 が完全に同一コードで、`tuunen` フラグ切り替えが無意味
- Module3/4 は最終行を A列基準で取得しているが出力は C〜G列。フィルタ非表示行も出力される
- 出力マクロがセルを1個ずつ読んでいて遅い（配列一括読み込みで大幅改善可能）
