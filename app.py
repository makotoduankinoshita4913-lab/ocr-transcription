from pathlib import Path

import pandas as pd
import streamlit as st

from convert_aps_to_uji_tsv import OUTPUT_HEADERS, rows_from_bytes, rows_to_tsv


st.set_page_config(
    page_title="APS → 宇治在庫表 変換",
    page_icon="🚗",
    layout="wide",
)


def build_dataframe(rows: list[list[str]]) -> pd.DataFrame:
    return pd.DataFrame(rows, columns=OUTPUT_HEADERS)


st.title("APS → 宇治在庫表 変換")
st.caption("APS店舗在庫表.xlsx を読み込んで、宇治在庫表へ B列から貼り付けるTSVを作成します。O列は空欄のまま残します。")
st.info("上の必須欄には `APS店舗在庫表.xlsx` を入れてください。`宇治在庫表` は完成形の参照ファイルなので、入れる場合は下の任意欄で大丈夫です。")

with st.sidebar:
    st.subheader("使い方")
    st.markdown(
        "\n".join(
            [
                "1. `APS店舗在庫表.xlsx` をアップロード",
                "2. 必要なら参照用の `xlsm` / `pdf` もアップロード",
                "3. 画面下のTSVをコピー",
                "4. 宇治在庫表の `B列` 開始セルに貼り付け",
            ]
        )
    )
    st.info("`O列` には何も出力しません。既存の数式を残す前提です。")


aps_file = st.file_uploader(
    "変換元: APS店舗在庫表.xlsx（必須）",
    type=["xlsx"],
    help="ここには宇治在庫表ではなく、APS店舗在庫表を選んでください。",
)

reference_files = st.file_uploader(
    "参照用: 宇治在庫表 / PDF / xlsm など（任意）",
    type=["xlsm", "xlsx", "pdf"],
    accept_multiple_files=True,
    help="参照確認用です。現状の変換処理では直接使いません。",
)

if reference_files:
    st.write("参照ファイル")
    for ref in reference_files:
        st.write(f"- {ref.name}")

if not aps_file:
    sample_path = Path("APS店舗在庫表.xlsx")
    if sample_path.exists():
        st.info("サンプル確認用に、ローカルの `APS店舗在庫表.xlsx` でも試せます。")
        if st.button("サンプルで読み込む"):
            aps_bytes = sample_path.read_bytes()
            rows = rows_from_bytes(aps_bytes)
            st.session_state["rows"] = rows
            st.session_state["source_name"] = sample_path.name
    else:
        st.stop()
else:
    try:
        rows = rows_from_bytes(aps_file.getvalue())
        st.session_state["rows"] = rows
        st.session_state["source_name"] = aps_file.name
    except Exception as exc:
        st.error(f"ファイルの読み取りに失敗しました: {exc}")
        st.stop()

rows = st.session_state.get("rows")
source_name = st.session_state.get("source_name", "")

if not rows:
    if "宇治在庫表" in source_name:
        st.error("これは完成形の `宇治在庫表` のようです。上の必須欄には `APS店舗在庫表.xlsx` を入れてください。")
        st.info("宇治在庫表を一緒に置いておきたい場合は、下の `参照用` 欄にアップロードしてください。")
    else:
        st.warning("変換できる行が見つかりませんでした。APS店舗在庫表のファイルか確認してください。")
    st.stop()

df = build_dataframe(rows)
tsv_text = rows_to_tsv(rows)

metric_col1, metric_col2 = st.columns(2)
metric_col1.metric("変換行数", len(rows))
metric_col2.metric("元ファイル", source_name)

st.subheader("プレビュー")
st.dataframe(df, use_container_width=True, hide_index=True)

st.subheader("貼り付け用TSV")
st.code(tsv_text, language="tsv")
st.download_button(
    label="TSVをダウンロード",
    data=tsv_text,
    file_name="uji_inventory_paste.tsv",
    mime="text/tab-separated-values",
    use_container_width=True,
)

st.text_area(
    "コピー用テキスト",
    value=tsv_text,
    height=360,
    help="ここを丸ごとコピーして、宇治在庫表の B列開始セルに貼り付けてください。",
)
