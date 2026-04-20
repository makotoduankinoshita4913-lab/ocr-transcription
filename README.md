# 書類転記アプリ

PDF化した書類をOCRで読み取り、スプレッドシートへ貼り付けやすいTSVを作るStreamlitアプリです。

社内共有用のメインアプリは `aps_converter_app.py` です。Bonsai/local LLMは使いません。

## 対応している書類

- APS在庫表
- オークション精算書
- ウメモト納品書
- MK石油請求書

## できること

- PDFまたは画像ファイルをアップロード
- 複数ページPDFを全ページOCR
- 車名欄付近を切り出して再OCR
- 画像を拡大、コントラスト強化、白黒化して読み取り補助
- 抽出後の表を画面上で確認・修正
- スプレッドシート貼り付け用TSVを出力
- OCRテキストを確認

## APS在庫表の使い方

`APS在庫表` タブでは、APS店舗在庫表のExcelまたはPDFをアップロードできます。

- Excel: `APS店舗在庫表.xlsx`
- PDF: `APS店舗在庫表` をPDF出力したもの

出力は、在庫表の `B列` 開始セルへ貼り付ける前提です。`O列` は出力せず、既存の数式を残します。

PDF版は、PDF内のテキストを読み取って変換します。OCRではないため速いですが、PDF出力時点で文字が欠けている場合は画面上の表で修正してください。

## 起動方法

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements-share.txt
streamlit run aps_converter_app.py
```

## OCRの事前準備

このアプリはTesseract OCRを使います。MacではHomebrewで入れます。

```bash
brew install tesseract tesseract-lang
```

## Streamlit Community Cloudに追加する場合

Streamlit Cloudの `New app` で次を指定します。

- Repository: `makotoduankinoshita4913-lab/aps-excel-convert`
- Branch: `main`
- Main file path: `aps_converter_app.py`

Cloud上では `packages.txt` によりTesseract OCRもインストールされます。

## スキャン設定の目安

- 通常は300dpiでも可
- 車名や金額が抜ける場合は600dpi推奨
- 文書モード、グレースケール、傾き補正ONがおすすめ
- 圧縮を強くしすぎるとOCR精度が落ちることがあります

## 運用時の注意

OCR結果は100%ではありません。貼り付け前に、画面上の表で日付・金額・車名を確認してください。

特に確認したい項目:

- 金額の桁
- 日付
- 車名
- 客注、店舗経費などの転記列

## 開発用メモ

- `aps_converter_app.py`: 社内共有用の起動ファイル
- `streamlit_app.py`: 社内共有用のBonsaiなし版UI本体
- `app.py`: これまでの実験版。領収書手入力とlocal LLM接続設定を含みます
- `requirements-share.txt`: 社内共有用の最小依存
- `requirements.txt`: 実験版も含めた依存

## APS変換CLI

APS在庫表変換CLIも残しています。

```bash
python3 convert_aps_to_uji_tsv.py
```
