#!/usr/bin/env python3

import argparse
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
from io import BytesIO
from pathlib import Path


NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
OUTPUT_HEADERS = [
    "車名",
    "仕入先",
    "仕入日",
    "金額",
    "車番",
    "在庫日数",
    "陸送等",
    "ガリバー",
    "UNO/KEP",
    "中野鈑金",
    "苗村/松崎",
    "ウメモト",
    "その他",
]


def load_shared_strings(book: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in book.namelist():
        return []

    root = ET.fromstring(book.read("xl/sharedStrings.xml"))
    return [
        "".join(text.text or "" for text in item.iterfind(".//a:t", NS))
        for item in root.findall("a:si", NS)
    ]


def cell_value(row: ET.Element, col: str, shared_strings: list[str]) -> str:
    ref = f"{col}{row.attrib['r']}"
    cell = row.find(f"a:c[@r='{ref}']", NS)
    if cell is None:
        return ""

    value = cell.find("a:v", NS)
    if value is None:
        return ""

    if cell.attrib.get("t") == "s":
        return shared_strings[int(value.text)]

    return value.text or ""


def excel_serial_to_md(serial: str) -> str:
    if not serial:
        return ""

    try:
        number = float(serial)
    except ValueError:
        return serial

    date = datetime(1899, 12, 30) + timedelta(days=number)
    return f"{date.month}/{date.day}"


def ymd_to_md(value: str) -> str:
    match = re.search(r"(\d{4})/(\d{1,2})/(\d{1,2})", value)
    if not match:
        return value
    return f"{int(match.group(2))}/{int(match.group(3))}"


def normalize_money(value: str) -> str:
    return re.sub(r"[\\¥￥,，円\s]", "", value or "")


def vehicle_number(chassis_number: str) -> str:
    digits = "".join(ch for ch in chassis_number if ch.isdigit())
    return digits[-3:] if digits else ""


def build_car_name(model: str, grade: str) -> str:
    return " ".join(part for part in [model.strip(), grade.strip()] if part)


def iter_output_rows_from_book(book: zipfile.ZipFile):
    shared_strings = load_shared_strings(book)
    sheet = ET.fromstring(book.read("xl/worksheets/sheet1.xml"))

    for row in sheet.findall("a:sheetData/a:row", NS):
        row_number = int(row.attrib["r"])
        if row_number < 6 or row_number > 29:
            continue

        purchase_amount = cell_value(row, "S", shared_strings)
        if not purchase_amount:
            continue

        yield [
            build_car_name(
                cell_value(row, "I", shared_strings),
                cell_value(row, "J", shared_strings),
            ),
            cell_value(row, "O", shared_strings),
            excel_serial_to_md(cell_value(row, "G", shared_strings)),
            purchase_amount,
            vehicle_number(cell_value(row, "K", shared_strings)),
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
        ]


def iter_output_rows(xlsx_path: Path):
    with zipfile.ZipFile(xlsx_path) as book:
        yield from iter_output_rows_from_book(book)


def rows_from_bytes(content: bytes) -> list[list[str]]:
    with zipfile.ZipFile(BytesIO(content)) as book:
        return list(iter_output_rows_from_book(book))


def rows_from_pdf_bytes(content: bytes) -> list[list[str]]:
    try:
        import fitz
    except ImportError as exc:
        raise RuntimeError("PDF読み取りには PyMuPDF が必要です。") from exc

    with fitz.open(stream=content, filetype="pdf") as document:
        if document.page_count < 3:
            raise ValueError("APS在庫表PDFは少なくとも3ページ必要です。")

        vehicle_rows = _extract_pdf_vehicle_rows(document[0].get_text("words"))
        amounts = _extract_pdf_purchase_amounts(document[2].get_text("words"))

    if not vehicle_rows:
        return []

    output_rows = []
    for index, vehicle in enumerate(vehicle_rows):
        purchase_amount = amounts[index] if index < len(amounts) else ""
        output_rows.append(
            [
                vehicle["car_name"],
                vehicle["supplier"],
                vehicle["purchase_date"],
                purchase_amount,
                vehicle_number(vehicle["chassis_number"]),
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
            ]
        )

    return output_rows


def _group_words_by_y(words: list[tuple], y_min: float, y_max: float, tolerance: float = 3.0) -> list[list[tuple]]:
    rows: list[list[tuple]] = []
    row_ys: list[float] = []

    for word in sorted(words, key=lambda item: ((item[1] + item[3]) / 2, item[0])):
        y_center = (word[1] + word[3]) / 2
        if y_center < y_min or y_center > y_max:
            continue

        for index, row_y in enumerate(row_ys):
            if abs(y_center - row_y) <= tolerance:
                rows[index].append(word)
                row_ys[index] = (row_y + y_center) / 2
                break
        else:
            rows.append([word])
            row_ys.append(y_center)

    return [sorted(row, key=lambda item: item[0]) for row in rows]


def _word_texts_in_range(row: list[tuple], min_x: float, max_x: float) -> list[str]:
    values = []
    for word in row:
        x_center = (word[0] + word[2]) / 2
        if min_x <= x_center < max_x:
            values.append(str(word[4]))
    return values


def _extract_pdf_vehicle_rows(words: list[tuple]) -> list[dict[str, str]]:
    rows = []
    for row in _group_words_by_y(words, y_min=68, y_max=340):
        first_text = str(row[0][4]) if row else ""
        first_x = row[0][0] if row else 9999
        if first_x > 60 or not first_text.isdigit():
            continue

        purchase_date = ""
        for word in row:
            text = str(word[4])
            if re.fullmatch(r"\d{4}/\d{2}/\d{2}", text):
                purchase_date = ymd_to_md(text)
                break

        model = " ".join(_word_texts_in_range(row, 292, 370))
        grade_parts = _word_texts_in_range(row, 370, 450)
        chassis_parts = _word_texts_in_range(row, 450, 575)
        supplier = " ".join(_word_texts_in_range(row, 660, 760))

        grade_text, chassis_number = _split_grade_and_chassis(" ".join(grade_parts + chassis_parts))
        rows.append(
            {
                "car_name": build_car_name(model, grade_text),
                "supplier": supplier,
                "purchase_date": purchase_date,
                "chassis_number": chassis_number,
            }
        )

    return rows


def _split_grade_and_chassis(value: str) -> tuple[str, str]:
    text = re.sub(r"\s+", " ", value).strip()
    hyphen_match = re.search(r"[A-Z0-9]{2,10}-[A-Z0-9]+", text)
    vin_match = re.search(r"[A-Z]{2,}[A-Z0-9]{10,}", text)
    match = hyphen_match or vin_match
    if not match:
        return text, ""

    grade = text[: match.start()].strip()
    chassis = match.group(0)
    return grade, chassis


def _extract_pdf_purchase_amounts(words: list[tuple]) -> list[str]:
    amounts = []
    for row in _group_words_by_y(words, y_min=68, y_max=340):
        row_amounts = _word_texts_in_range(row, 174, 214)
        amount = next((normalize_money(value) for value in row_amounts if "\\" in value or "¥" in value), "")
        if amount:
            amounts.append(amount)
    return amounts


def rows_to_tsv(rows: list[list[str]]) -> str:
    return "\n".join("\t".join(row) for row in rows)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Convert APS inventory rows into tab-separated rows for the Uji sheet."
    )
    parser.add_argument(
        "xlsx",
        nargs="?",
        default="APS店舗在庫表.xlsx",
        help="Path to the APS workbook (.xlsx). Defaults to APS店舗在庫表.xlsx.",
    )
    args = parser.parse_args()

    rows = list(iter_output_rows(Path(args.xlsx)))
    if rows:
        print(rows_to_tsv(rows))

    return 0


if __name__ == "__main__":
    sys.exit(main())
