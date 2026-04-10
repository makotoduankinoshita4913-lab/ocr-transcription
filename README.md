# APS → 宇治在庫表 変換

APS店舗在庫表から、宇治在庫表へそのまま貼り付けやすいTSVを作る `Streamlit` アプリです。

## できること

- `APS店舗在庫表.xlsx` をアップロードして変換
- 宇治在庫表向けの列順でプレビュー表示
- Googleスプレッドシートへ貼り付けるためのTSVを表示
- `O列` は空欄のままにして、既存の数式を残す

## 出力列

宇治在庫表の `B列` から貼り付ける前提で、次の列を出力します。

- `B` 車名
- `C` 仕入先
- `D` 仕入日
- `E` 金額
- `F` 車番
- `G` 在庫日数
- `H` 陸送等
- `I` ガリバー
- `J` UNO/KEP
- `K` 中野鈑金
- `L` 苗村/松崎
- `M` ウメモト
- `N` その他
- `O` は空欄

## 車名の作り方

- `車名 = 車種名 + グレード名`

## 起動方法

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## CLIで使う場合

```bash
python3 convert_aps_to_uji_tsv.py
```

## 備考

- 現状の参照用 `xlsm` / `pdf` アップロードは確認用で、変換ロジックにはまだ使っていません。
- 変換元の想定は `APS店舗在庫表.xlsx` の1シート目 (`sheet1.xml`) です。
