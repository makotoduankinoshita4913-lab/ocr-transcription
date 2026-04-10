#!/usr/bin/env python3

import argparse
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
            "",
        ]


def iter_output_rows(xlsx_path: Path):
    with zipfile.ZipFile(xlsx_path) as book:
        yield from iter_output_rows_from_book(book)


def rows_from_bytes(content: bytes) -> list[list[str]]:
    with zipfile.ZipFile(BytesIO(content)) as book:
        return list(iter_output_rows_from_book(book))


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
