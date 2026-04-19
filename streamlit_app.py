import pandas as pd
import streamlit as st

from app import (
    ACCOUNTING_FIELD_ORDER,
    AUCTION_FIELD_ORDER,
    TAB_AUCTION,
    TAB_MK,
    TAB_UMEMOTO,
    accounting_rows_to_tsv,
    auction_to_tsv,
    extract_auction_settlement,
    extract_mk_invoice,
    extract_umemoto_delivery,
)


st.set_page_config(
    page_title="書類転記アプリ",
    page_icon="🧾",
    layout="wide",
)


def render_sidebar() -> dict[str, bool]:
    with st.sidebar:
        st.subheader("社内共有版")
        st.caption("Bonsai/local LLMは使わず、PDF画像化・OCR・抽出ルールで転記候補を作成します。")
        debug = st.toggle("デバッグ表示", value=False)

        st.subheader("スキャンの目安")
        st.markdown(
            "\n".join(
                [
                    "- 通常は300dpiでも可",
                    "- 車名や金額の読み落としがある時は600dpi推奨",
                    "- 文書モード / グレースケール / 傾き補正ONがおすすめ",
                ]
            )
        )

        st.subheader("使い方")
        st.markdown(
            "\n".join(
                [
                    "1. 書類タイプのタブを選ぶ",
                    "2. PDFまたは画像をアップロード",
                    "3. 抽出結果を表で確認・修正",
                    "4. TSVをスプレッドシートへ貼り付け",
                ]
            )
        )

    return {"debug": debug}


def render_auction_tab(settings: dict[str, bool]) -> None:
    st.subheader(TAB_AUCTION)
    st.caption("オークション精算書PDFから明細行をOCRで読み取り、転記用TSVを作成します。")

    source_file = st.file_uploader(
        "オークション精算書PDF / 画像",
        type=["pdf", "png", "jpg", "jpeg", "webp"],
        accept_multiple_files=False,
        key="auction_source",
    )
    reference_file = st.file_uploader(
        "転記先サンプルPDF（任意）",
        type=["pdf", "png", "jpg", "jpeg", "webp"],
        accept_multiple_files=False,
        key="auction_reference",
        help="確認用です。抽出処理には使いません。",
    )
    if reference_file:
        st.caption(f"転記先サンプル: {reference_file.name}")

    if not source_file:
        st.info("精算書PDFをアップロードしてください。")
        return

    with st.spinner("精算書をOCRで読み取っています..."):
        extraction = extract_auction_settlement(source_file)

    if extraction.get("error"):
        st.warning(extraction["error"])

    rows = extraction.get("rows") or []
    if not rows:
        st.warning("明細候補を抽出できませんでした。PDFの向きや解像度を確認してください。")
        render_ocr_text(extraction.get("ocr_text", ""), "auction_empty_ocr")
        return

    st.success(f"{len(rows)}件の明細候補を抽出しました。貼り付け前に表で確認・修正してください。")
    edited_df = st.data_editor(
        pd.DataFrame(rows, columns=AUCTION_FIELD_ORDER),
        width="stretch",
        hide_index=True,
        num_rows="dynamic",
        key="auction_rows_editor",
    )
    edited_rows = edited_df.fillna("").to_dict("records")
    render_tsv(auction_to_tsv(edited_rows), "auction_copy", height=220)
    render_ocr_text(extraction.get("ocr_text", ""), "auction_ocr")

    if settings["debug"]:
        st.subheader("デバッグ")
        st.write("OCR単語数", len(extraction.get("words") or []))
        st.json(extraction.get("words", [])[:120])


def render_accounting_tab(
    title: str,
    uploader_label: str,
    extractor,
    settings: dict[str, bool],
) -> None:
    st.subheader(title)
    st.caption("納品書.転記先サンプルの列順に合わせて、日付・金額・車名などをTSV化します。")

    uploaded_file = st.file_uploader(
        uploader_label,
        type=["pdf", "png", "jpg", "jpeg", "webp"],
        accept_multiple_files=False,
        key=f"{title}_source",
    )
    reference_file = st.file_uploader(
        "転記先サンプルPDF（任意）",
        type=["pdf", "png", "jpg", "jpeg", "webp"],
        accept_multiple_files=False,
        key=f"{title}_reference",
        help="確認用です。抽出処理には使いません。",
    )
    if reference_file:
        st.caption(f"転記先サンプル: {reference_file.name}")

    if not uploaded_file:
        st.info(f"{title}のPDFをアップロードしてください。")
        return

    with st.spinner(f"{title}をOCRで読み取っています..."):
        extraction = extractor(uploaded_file)

    if extraction.get("error"):
        st.warning(extraction["error"])

    rows = extraction.get("rows") or []
    if not rows:
        st.warning("転記候補を抽出できませんでした。PDFの向きや解像度を確認してください。")
        render_ocr_text(extraction.get("ocr_text", ""), f"{title}_empty_ocr")
        return

    st.success("転記候補を作成しました。貼り付け前に表で確認・修正してください。")
    edited_df = st.data_editor(
        pd.DataFrame(rows, columns=ACCOUNTING_FIELD_ORDER),
        width="stretch",
        hide_index=True,
        num_rows="dynamic",
        key=f"{title}_editor",
    )
    edited_rows = edited_df.fillna("").to_dict("records")
    render_tsv(accounting_rows_to_tsv(edited_rows), f"{title}_copy", height=160)
    render_ocr_text(extraction.get("ocr_text", ""), f"{title}_ocr")

    if settings["debug"]:
        st.subheader("デバッグ")
        st.json(extraction)


def render_tsv(tsv_text: str, key: str, height: int) -> None:
    st.subheader("スプレッドシート貼り付け用TSV")
    st.code(tsv_text, language="tsv")
    st.text_area(
        "コピー用テキスト",
        value=tsv_text,
        height=height,
        help="表を確認・修正したあと、ここをコピーしてスプレッドシートへ貼り付けてください。",
        key=key,
    )


def render_ocr_text(ocr_text: str, key: str) -> None:
    with st.expander("OCRで読み取ったテキスト", expanded=False):
        st.text_area("OCRテキスト", value=ocr_text, height=300, key=key)


def render_ui() -> None:
    settings = render_sidebar()
    st.title("書類転記アプリ")
    st.caption("Bonsaiなし版。PDF/画像をOCRで読み取り、確認・修正してからスプレッドシートへ貼り付けるためのTSVを作成します。")

    auction_tab, umemoto_tab, mk_tab = st.tabs([TAB_AUCTION, TAB_UMEMOTO, TAB_MK])
    with auction_tab:
        render_auction_tab(settings)
    with umemoto_tab:
        render_accounting_tab(
            title=TAB_UMEMOTO,
            uploader_label="ウメモト納品書PDF / 画像",
            extractor=extract_umemoto_delivery,
            settings=settings,
        )
    with mk_tab:
        render_accounting_tab(
            title=TAB_MK,
            uploader_label="MK石油請求書PDF / 画像",
            extractor=extract_mk_invoice,
            settings=settings,
        )


if __name__ == "__main__":
    render_ui()
