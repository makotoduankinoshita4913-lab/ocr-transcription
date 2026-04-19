import csv
import io
import json
import re
import shutil
import subprocess
import tempfile
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import Any
from urllib.parse import urljoin

import pandas as pd
import streamlit as st

try:
    import requests
except ImportError:
    requests = None


DEFAULT_BASE_URL = "http://localhost:8080"
DEFAULT_MODEL = "Bonsai-8B-Q1_0.gguf"
APP_MODE_RECEIPT = "領収書入力"
APP_MODE_AUCTION = "オークション精算書抽出"
TAB_AUCTION = "オークション精算書"
TAB_UMEMOTO = "ウメモト納品書"
TAB_MK = "MK石油請求書"
MODE_MANUAL = "手入力のみ"
MODE_DUMMY = "手入力 + ダミー整形"
MODE_LOCAL_LLM = "手入力 + local LLM整形"
FIELD_ORDER = ["date", "payee", "amount", "description"]
FIELD_LABELS = {
    "date": "日付",
    "payee": "支払先",
    "amount": "金額",
    "description": "内容",
}
UPLOAD_FILE_TYPES = ["pdf", "png", "jpg", "jpeg", "webp", "heic", "txt", "csv"]
OCR_IMAGE_FILE_TYPES = {"png", "jpg", "jpeg", "webp", "heic"}
OCR_LANG = "jpn+eng"
ACCOUNTING_TARGET_REGIONS = [
    ("detail_mid", (0.00, 0.28, 1.00, 0.55)),
    ("detail_left", (0.00, 0.30, 0.72, 0.50)),
]
AUCTION_FIELD_ORDER = [
    "発生日",
    "出品番号",
    "車名",
    "型式",
    "車台番号",
    "年式",
    "出品料",
    "出品料税",
    "出品料税込",
    "成約料",
    "成約料税",
    "成約料税込",
    "車両金額",
    "車両金額税",
    "車両金額税込",
    "R預託金",
    "合計",
    "OCR元行",
]
AUCTION_BASE_WIDTH = 2600
AUCTION_BASE_HEIGHT = 1838
ACCOUNTING_FIELD_ORDER = [
    "日付",
    "金額",
    "車名",
    "お客様名",
    "販売",
    "買取",
    "保障",
    "クレーム",
    "客注",
    "店舗経費",
]


def build_receipt_data(date: str, payee: str, amount: str, description: str) -> dict[str, str]:
    """手入力された領収書データを内部形式にまとめる。"""
    return {
        "date": date,
        "payee": payee,
        "amount": amount,
        "description": description,
    }


def receipt_to_tsv(receipt: dict[str, Any]) -> str:
    """スプレッドシートへ貼り付けやすいTSV 1行を作る。"""
    return "\t".join(str(receipt.get(field, "")) for field in FIELD_ORDER)


def normalize_spaces(value: Any) -> str:
    text = "" if value is None else str(value)
    text = unicodedata.normalize("NFKC", text)
    text = text.replace("\u3000", " ")
    return re.sub(r"\s+", " ", text).strip()


def normalize_amount(value: Any) -> str:
    text = normalize_spaces(value)
    text = re.sub(r"[¥￥,，円\s]", "", text)
    return text.strip()


def normalize_date(value: Any) -> str:
    text = normalize_spaces(value)
    if not text:
        return text

    text = text.replace("年", "/").replace("月", "/").replace("日", "")
    text = re.sub(r"\s*/\s*", "/", text)
    text = re.sub(r"[-.]", "/", text)
    text = re.sub(r"/+", "/", text).strip("/")

    for fmt in ("%Y/%m/%d", "%y/%m/%d", "%Y/%m", "%m/%d"):
        try:
            parsed = datetime.strptime(text, fmt)
        except ValueError:
            continue

        if fmt == "%m/%d":
            parsed = parsed.replace(year=datetime.now().year)
        if fmt == "%Y/%m":
            return parsed.strftime("%Y/%m")
        return parsed.strftime("%Y/%m/%d")

    return text


def normalize_ocr_text(text: Any) -> str:
    raw_text = "" if text is None else str(text)
    raw_text = unicodedata.normalize("NFKC", raw_text).replace("\u3000", " ")
    lines = [re.sub(r"[ \t]+", " ", line).strip() for line in raw_text.splitlines()]
    lines = [line for line in lines if line]
    return "\n".join(lines)


def dummy_format_receipt(receipt: dict[str, Any]) -> dict[str, str]:
    """LLMなしでも使える簡易整形。"""
    return {
        "date": normalize_date(receipt.get("date", "")),
        "payee": normalize_spaces(receipt.get("payee", "")),
        "amount": normalize_amount(receipt.get("amount", "")),
        "description": normalize_spaces(receipt.get("description", "")),
    }


def build_llm_prompt(receipt: dict[str, Any]) -> list[dict[str, str]]:
    system_prompt = """
あなたは領収書データをスプレッドシート貼り付け用に正規化する補助ツールです。
推測は禁止です。わからない値は元の値を維持してください。
余計な説明、Markdown、コードブロックは出さず、JSONのみを返してください。
対象フィールドは date, payee, amount, description の4つだけです。
amount は可能なら数値のみ、date は可能なら YYYY/MM/DD にしてください。
空白や全角空白の整理を優先し、支払先や内容は軽い正規化だけにしてください。
日本語で処理して構いません。
""".strip()
    user_prompt = {
        "instruction": "次の手入力領収書データを正規化し、JSONのみで返してください。",
        "input": {field: str(receipt.get(field, "")) for field in FIELD_ORDER},
        "output_schema": {
            "date": "string",
            "payee": "string",
            "amount": "string",
            "description": "string",
        },
    }
    return [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": json.dumps(user_prompt, ensure_ascii=False)},
    ]


def build_llm_plain_prompt(receipt: dict[str, Any]) -> str:
    messages = build_llm_prompt(receipt)
    return "\n\n".join(f"{message['role']}:\n{message['content']}" for message in messages)


def build_chat_payload(model: str, messages: list[dict[str, str]]) -> dict[str, Any]:
    return {
        "model": model,
        "messages": messages,
        "temperature": 0,
        "max_tokens": 512,
        "stream": False,
    }


def build_completion_payload(model: str, receipt: dict[str, Any]) -> dict[str, Any]:
    return {
        "model": model,
        "prompt": build_llm_plain_prompt(receipt),
        "temperature": 0,
        "max_tokens": 512,
        "stream": False,
    }


def call_local_llm(
    base_url: str,
    model: str,
    receipt: dict[str, Any],
    timeout_seconds: float,
) -> dict[str, Any]:
    """OpenAI互換のchatを優先し、失敗時はcompletion系も試す。"""
    if requests is None:
        return {
            "endpoint": None,
            "payload": build_chat_payload(model, build_llm_prompt(receipt)),
            "raw_response": None,
            "content": "",
            "parsed": None,
            "error": "`requests` がインストールされていないため、local LLM整形をスキップしました。",
            "attempts": [],
        }

    messages = build_llm_prompt(receipt)
    chat_payload = build_chat_payload(model, messages)
    completion_payload = build_completion_payload(model, receipt)
    attempts = [
        ("v1/chat/completions", chat_payload, "chat"),
        ("v1/completions", completion_payload, "completion"),
        ("completion", completion_payload, "completion"),
    ]

    result: dict[str, Any] = {
        "endpoint": None,
        "payload": chat_payload,
        "raw_response": None,
        "content": "",
        "parsed": None,
        "error": None,
        "attempts": [],
    }

    errors = []
    for path, payload, response_kind in attempts:
        endpoint = urljoin(base_url.rstrip("/") + "/", path)
        result["endpoint"] = endpoint
        result["payload"] = payload

        try:
            response = requests.post(endpoint, json=payload, timeout=timeout_seconds)
            raw_response = response.text
            result["raw_response"] = raw_response
            attempt_info = {
                "endpoint": endpoint,
                "status_code": response.status_code,
                "raw_response": raw_response[:2000],
            }
            result["attempts"].append(attempt_info)
            response.raise_for_status()
            response_json = response.json()
            content = extract_llm_content(response_json, response_kind)
            result["content"] = content
            result["parsed"] = extract_json_from_llm_response(content)
            if result["parsed"]:
                result["error"] = None
                return result
            errors.append(f"{endpoint}: JSONを抽出できませんでした")
        except requests.exceptions.RequestException as exc:
            errors.append(f"{endpoint}: {exc}")
        except (KeyError, IndexError, TypeError, ValueError, json.JSONDecodeError) as exc:
            errors.append(f"{endpoint}: レスポンス処理失敗: {exc}")

    result["error"] = "local LLM整形に失敗しました。 " + " / ".join(errors)

    return result


def extract_llm_content(response_json: dict[str, Any], response_kind: str) -> str:
    if response_kind == "chat":
        return str(response_json["choices"][0]["message"]["content"])

    choice = response_json["choices"][0]
    if "text" in choice:
        return str(choice["text"])
    if "content" in choice:
        return str(choice["content"])
    if "message" in choice and "content" in choice["message"]:
        return str(choice["message"]["content"])
    return json.dumps(response_json, ensure_ascii=False)


def extract_json_from_llm_response(text: Any) -> dict[str, Any] | None:
    """JSONのみでない返答でも、可能なら最初のJSONオブジェクトを抽出する。"""
    if isinstance(text, dict):
        return text
    if not isinstance(text, str) or not text.strip():
        return None

    candidates = [text.strip()]
    fenced = re.findall(r"```(?:json)?\s*(\{.*?\})\s*```", text, flags=re.DOTALL | re.IGNORECASE)
    candidates.extend(fenced)

    balanced = find_first_json_object(text)
    if balanced:
        candidates.append(balanced)

    for candidate in candidates:
        try:
            parsed = json.loads(candidate)
        except json.JSONDecodeError:
            continue
        if isinstance(parsed, dict):
            return parsed

    return None


def find_first_json_object(text: str) -> str | None:
    start = text.find("{")
    if start == -1:
        return None

    depth = 0
    in_string = False
    escape = False
    for index, char in enumerate(text[start:], start=start):
        if in_string:
            if escape:
                escape = False
            elif char == "\\":
                escape = True
            elif char == '"':
                in_string = False
            continue

        if char == '"':
            in_string = True
        elif char == "{":
            depth += 1
        elif char == "}":
            depth -= 1
            if depth == 0:
                return text[start : index + 1]

    return None


def coerce_receipt_fields(parsed: dict[str, Any], fallback: dict[str, Any]) -> dict[str, str]:
    """LLMのJSONから必要な4項目だけを採用し、欠損時は元データを維持する。"""
    merged = {}
    for field in FIELD_ORDER:
        value = parsed.get(field, fallback.get(field, ""))
        merged[field] = str(value) if value is not None else str(fallback.get(field, ""))
    return merged


def build_uploaded_file_summaries(uploaded_files: list[Any] | None) -> list[dict[str, Any]]:
    """アップロードされた領収書ファイルの表示用情報を作る。"""
    summaries = []
    for uploaded_file in uploaded_files or []:
        summaries.append(
            {
                "ファイル名": uploaded_file.name,
                "種類": uploaded_file.type or "不明",
                "サイズKB": round(uploaded_file.size / 1024, 1),
            }
        )
    return summaries


def extract_text_from_plain_file(file_bytes: bytes) -> str:
    for encoding in ("utf-8-sig", "utf-8", "cp932", "shift_jis"):
        try:
            return file_bytes.decode(encoding)
        except UnicodeDecodeError:
            continue
    return file_bytes.decode("utf-8", errors="replace")


def run_tesseract(image_path: Path, work_dir: Path, psm: str = "6", timeout: int = 60) -> tuple[str, str | None]:
    tesseract_path = shutil.which("tesseract") or "/opt/homebrew/bin/tesseract"
    if not Path(tesseract_path).exists():
        return "", "Tesseractが見つかりません。`brew install tesseract tesseract-lang` が必要です。"

    try:
        result = subprocess.run(
            [tesseract_path, image_path.name, "stdout", "-l", OCR_LANG, "--psm", psm],
            cwd=work_dir,
            text=True,
            capture_output=True,
            timeout=timeout,
            errors="replace",
            check=False,
        )
    except subprocess.TimeoutExpired:
        return "", "OCRがタイムアウトしました。"

    if result.returncode != 0:
        return "", result.stderr.strip() or "OCRに失敗しました。"
    return result.stdout, None


def crop_relative_region(
    image_path: Path,
    work_dir: Path,
    region_name: str,
    box: tuple[float, float, float, float],
) -> Path | None:
    try:
        from PIL import Image

        with Image.open(image_path) as image:
            width, height = image.size
            left, top, right, bottom = box
            crop_box = (
                max(0, round(width * left)),
                max(0, round(height * top)),
                min(width, round(width * right)),
                min(height, round(height * bottom)),
            )
            if crop_box[2] <= crop_box[0] or crop_box[3] <= crop_box[1]:
                return None
            crop = image.crop(crop_box)
            output_path = work_dir / f"{image_path.stem}_{region_name}.png"
            crop.save(output_path)
            return output_path
    except Exception:
        return None


def build_enhanced_ocr_image(image_path: Path, work_dir: Path, suffix: str) -> Path | None:
    try:
        from PIL import Image, ImageEnhance, ImageFilter, ImageOps

        with Image.open(image_path) as image:
            gray = ImageOps.grayscale(image)
            gray = ImageEnhance.Contrast(gray).enhance(2.4)
            gray = gray.resize((gray.width * 3, gray.height * 3), Image.Resampling.LANCZOS)
            gray = gray.filter(ImageFilter.SHARPEN)
            enhanced = gray.point(lambda pixel: 255 if pixel > 170 else 0, mode="1").convert("L")
            output_path = work_dir / f"{image_path.stem}_{suffix}.png"
            enhanced.save(output_path)
            return output_path
    except Exception:
        return None


def extract_targeted_accounting_text(image_path: Path, work_dir: Path) -> tuple[str, list[str]]:
    """車名欄付近を切り出して、通常OCRで落ちた文字を拾い直す。"""
    texts = []
    errors = []
    for region_name, box in ACCOUNTING_TARGET_REGIONS:
        crop_path = crop_relative_region(image_path, work_dir, region_name, box)
        if not crop_path:
            continue

        candidates = [(crop_path, "6")]
        enhanced_path = build_enhanced_ocr_image(crop_path, work_dir, "enhanced")
        if enhanced_path:
            candidates.extend([(enhanced_path, "6"), (enhanced_path, "11")])

        for candidate_path, psm in candidates:
            text, error = run_tesseract(candidate_path, work_dir, psm=psm, timeout=45)
            normalized = normalize_ocr_text(text)
            if normalized:
                texts.append(f"[{region_name} psm{psm}]\n{normalized}")
            if error:
                errors.append(f"{region_name} psm{psm}: {error}")

    seen = set()
    unique_texts = []
    for text in texts:
        if text in seen:
            continue
        seen.add(text)
        unique_texts.append(text)
    return "\n\n".join(unique_texts), errors


def render_pdf_pages(
    pdf_path: Path,
    work_dir: Path,
    size: int = 2400,
    max_pages: int | None = None,
) -> tuple[list[Path], str | None]:
    try:
        import fitz
    except ImportError:
        first_page, error = render_pdf_first_page_with_quicklook(pdf_path, work_dir, size=size)
        if error:
            return [], error
        return [first_page] if first_page else [], None

    image_paths = []
    try:
        document = fitz.open(pdf_path)
        page_count = len(document) if max_pages is None else min(len(document), max_pages)
        for page_index in range(page_count):
            page = document[page_index]
            zoom = size / max(page.rect.width, page.rect.height)
            matrix = fitz.Matrix(zoom, zoom)
            pixmap = page.get_pixmap(matrix=matrix, alpha=False)
            image_path = work_dir / f"page_{page_index + 1:03}.png"
            pixmap.save(image_path)
            image_paths.append(image_path)
        document.close()
    except Exception as exc:
        return [], f"PDFの全ページ画像化に失敗しました: {exc}"

    if not image_paths:
        return [], "PDFの画像化結果が見つかりませんでした。"
    return image_paths, None


def render_pdf_first_page(pdf_path: Path, work_dir: Path, size: int = 2400) -> tuple[Path | None, str | None]:
    image_paths, error = render_pdf_pages(pdf_path, work_dir, size=size, max_pages=1)
    if error:
        return None, error
    return image_paths[0] if image_paths else None, None


def render_pdf_first_page_with_quicklook(pdf_path: Path, work_dir: Path, size: int = 2400) -> tuple[Path | None, str | None]:
    qlmanage_path = shutil.which("qlmanage") or "/usr/bin/qlmanage"
    if not Path(qlmanage_path).exists():
        return None, "PDF画像化に必要な qlmanage が見つかりません。"

    result = subprocess.run(
        [qlmanage_path, "-t", "-s", str(size), "-o", str(work_dir), str(pdf_path)],
        text=True,
        capture_output=True,
        timeout=60,
        check=False,
    )
    if result.returncode != 0:
        return None, result.stderr.strip() or result.stdout.strip() or "PDFの画像化に失敗しました。"

    png_files = sorted(work_dir.glob("*.png"), key=lambda path: path.stat().st_mtime, reverse=True)
    if not png_files:
        return None, "PDFの画像化結果が見つかりませんでした。"

    ocr_image_path = work_dir / "ocr_page.png"
    shutil.copyfile(png_files[0], ocr_image_path)
    return ocr_image_path, None


def extract_text_from_uploaded_file(uploaded_file: Any) -> dict[str, Any]:
    file_name = uploaded_file.name
    file_suffix = Path(file_name).suffix.lower().lstrip(".")
    file_bytes = uploaded_file.getvalue()

    result = {
        "file_name": file_name,
        "method": None,
        "text": "",
        "error": None,
    }

    if file_suffix in {"txt", "csv"}:
        result["method"] = "text"
        result["text"] = normalize_ocr_text(extract_text_from_plain_file(file_bytes))
        return result

    with tempfile.TemporaryDirectory(prefix="receipt_ocr_", dir=Path.cwd()) as temp_dir_name:
        work_dir = Path(temp_dir_name)
        source_path = work_dir / f"upload.{file_suffix or 'bin'}"
        source_path.write_bytes(file_bytes)

        if file_suffix == "pdf":
            ocr_image_paths, error = render_pdf_pages(source_path, work_dir, size=2400)
            if error:
                result["method"] = "pdf-render"
                result["error"] = error
                return result
        elif file_suffix in OCR_IMAGE_FILE_TYPES:
            ocr_image_path = work_dir / f"ocr_image.{file_suffix}"
            shutil.copyfile(source_path, ocr_image_path)
            ocr_image_paths = [ocr_image_path]
        else:
            result["error"] = "このファイル形式はまだ読み取り対象外です。"
            return result

        result["method"] = "tesseract-ocr"
        page_texts = []
        errors = []
        for page_index, ocr_image_path in enumerate(ocr_image_paths, start=1):
            ocr_text, error = run_tesseract(ocr_image_path, work_dir)
            normalized_text = normalize_ocr_text(ocr_text)
            if normalized_text:
                page_texts.append(f"--- page {page_index} ---\n{normalized_text}")
            if error:
                errors.append(f"{page_index}ページ目: {error}")

        result["text"] = "\n\n".join(page_texts)
        result["error"] = " / ".join(errors) if errors else None
        if not result["text"] and not result["error"]:
            result["error"] = "OCR結果が空でした。画像が小さいか、文字が読み取れない可能性があります。"
        return result


def extract_receipt_candidates_from_text(text: str) -> dict[str, str]:
    normalized_text = normalize_ocr_text(text)
    lines = normalized_text.splitlines()
    joined = "\n".join(lines)

    date = extract_candidate_date(joined)
    amount = extract_candidate_amount(joined)
    payee = extract_candidate_payee(lines)
    description = extract_candidate_description(lines)

    return dummy_format_receipt(
        {
            "date": date,
            "payee": payee,
            "amount": amount,
            "description": description,
        }
    )


def extract_candidate_date(text: str) -> str:
    patterns = [
        r"(20\d{2})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日?",
        r"(20\d{2})[/-](\d{1,2})[/-](\d{1,2})",
        r"(\d{2})[/-](\d{1,2})[/-](\d{1,2})",
    ]
    for pattern in patterns:
        matches = re.findall(pattern, text)
        if matches:
            year, month, day = matches[-1]
            if len(year) == 2:
                year = f"20{year}"
            return f"{year}/{int(month):02d}/{int(day):02d}"
    return ""


def extract_candidate_amount(text: str) -> str:
    compact_text = re.sub(r"(?<=\d)\s*,\s*(?=\d)", ",", text)
    compact_text = re.sub(r"(?<=\d)\s+(?=\d{3}(?:\D|$))", "", compact_text)
    money_matches = list(
        re.finditer(r"(?:[¥￥]\s*)?(\d{1,3}(?:,\d{3})+|\d{4,})(?:\s*円)?", compact_text)
    )
    if not money_matches:
        return ""

    amounts: list[tuple[int, int, bool]] = []
    for match in money_matches:
        cleaned = normalize_amount(match.group(1))
        if cleaned.isdigit():
            amount = int(cleaned)
            has_comma = "," in match.group(1)
            if 100 <= amount <= 50_000_000:
                amounts.append((amount, match.start(), has_comma))
    if not amounts:
        return ""

    preferred = [item for item in amounts if item[2]]
    if preferred:
        return str(max(preferred, key=lambda item: item[1])[0])
    return str(max(amounts, key=lambda item: item[1])[0])


def extract_candidate_payee(lines: list[str]) -> str:
    for line in lines:
        if "御中" in line:
            before = line.split("御中", 1)[0]
            chunks = re.split(r"\s{2,}|〒|TEL|FAX|会員|No\.", before)
            candidates = [chunk.strip(" :：,，") for chunk in chunks if chunk.strip(" :：,，")]
            if candidates:
                return candidates[-1]

    company_keywords = ("株式会社", "(株)", "㈱", "有限会社", "合同会社")
    for line in lines:
        if any(keyword in line for keyword in company_keywords):
            return line.strip(" :：,，")
    return ""


def extract_candidate_description(lines: list[str]) -> str:
    for line in lines:
        if "精算書" in line:
            return line
        if "領収" in line:
            return line
    return lines[0] if lines else ""


def extract_receipt_from_uploaded_files(uploaded_files: list[Any]) -> dict[str, Any]:
    extraction_results = []
    merged_text = []

    for uploaded_file in uploaded_files:
        result = extract_text_from_uploaded_file(uploaded_file)
        extraction_results.append(result)
        if result.get("text"):
            merged_text.append(result["text"])

    ocr_text = "\n\n".join(merged_text)
    return {
        "results": extraction_results,
        "text": ocr_text,
        "candidates": extract_receipt_candidates_from_text(ocr_text) if ocr_text else build_receipt_data("", "", "", ""),
    }


def run_tesseract_tsv(image_path: Path, work_dir: Path) -> tuple[list[dict[str, Any]], str | None]:
    tesseract_path = shutil.which("tesseract") or "/opt/homebrew/bin/tesseract"
    if not Path(tesseract_path).exists():
        return [], "Tesseractが見つかりません。`brew install tesseract tesseract-lang` が必要です。"

    try:
        result = subprocess.run(
            [tesseract_path, image_path.name, "stdout", "-l", OCR_LANG, "--psm", "11", "tsv"],
            cwd=work_dir,
            text=True,
            capture_output=True,
            timeout=90,
            errors="replace",
            check=False,
        )
    except subprocess.TimeoutExpired:
        return [], "座標付きOCRがタイムアウトしました。"

    if result.returncode != 0:
        return [], result.stderr.strip() or "座標付きOCRに失敗しました。"

    rows = csv.DictReader(io.StringIO(result.stdout), delimiter="\t")
    words = []
    for row in rows:
        if row.get("level") != "5":
            continue
        text = normalize_spaces(row.get("text", ""))
        if not text:
            continue
        try:
            words.append(
                {
                    "text": text,
                    "left": int(row["left"]),
                    "top": int(row["top"]),
                    "width": int(row["width"]),
                    "height": int(row["height"]),
                    "conf": float(row.get("conf") or 0),
                }
            )
        except (TypeError, ValueError, KeyError):
            continue
    return words, None


def auction_to_tsv(rows: list[dict[str, Any]]) -> str:
    lines = ["\t".join(AUCTION_FIELD_ORDER)]
    for row in rows:
        lines.append("\t".join(str(row.get(field, "")) for field in AUCTION_FIELD_ORDER))
    return "\n".join(lines)


def normalize_auction_date(text: str) -> str:
    match = re.search(r"(\d{2})/(\d{2})/(\d{2})", text)
    if not match:
        return ""
    year, month, day = match.groups()
    return f"20{year}/{month}/{day}"


def clean_auction_number(text: str) -> str:
    cleaned = normalize_spaces(text)
    cleaned = cleaned.replace("O", "0").replace("o", "0")
    match = re.search(r"\d{4,5}", cleaned)
    if match:
        return match.group(0)
    return re.sub(r"[^0-9]", "", cleaned)


def clean_auction_money(text: str) -> str:
    cleaned = normalize_spaces(text)
    cleaned = cleaned.replace("O", "0").replace("o", "0")
    cleaned = cleaned.replace("B", "8").replace("S", "5")
    cleaned = re.sub(r"^[Iil|]+(?=\d{4,})", "", cleaned)
    cleaned = re.sub(r"^1[.,](?=\d{4,})", "", cleaned)
    cleaned = re.sub(r"[^0-9]", "", cleaned)
    if cleaned == "110000":
        cleaned = "10000"
    return cleaned


def money_to_int(value: str) -> int | None:
    cleaned = clean_auction_money(value)
    if not cleaned:
        return None
    return int(cleaned)


def add_money_strings(*values: str) -> str:
    numbers = [money_to_int(value) for value in values]
    numbers = [number for number in numbers if number is not None]
    if not numbers:
        return ""
    return str(sum(numbers))


def scaled_range(start: int, end: int, actual: int, base: int) -> tuple[int, int]:
    return round(start * actual / base), round(end * actual / base)


def words_in_region(
    words: list[dict[str, Any]],
    x_range: tuple[int, int],
    y_range: tuple[int, int],
) -> list[dict[str, Any]]:
    x1, x2 = x_range
    y1, y2 = y_range
    found = []
    for word in words:
        center_x = word["left"] + word["width"] / 2
        center_y = word["top"] + word["height"] / 2
        if x1 <= center_x <= x2 and y1 <= center_y <= y2:
            found.append(word)
    return sorted(found, key=lambda item: (item["top"], item["left"]))


def join_region_text(
    words: list[dict[str, Any]],
    x_range: tuple[int, int],
    y_range: tuple[int, int],
) -> str:
    return normalize_spaces(" ".join(word["text"] for word in words_in_region(words, x_range, y_range)))


def best_money_in_region(
    words: list[dict[str, Any]],
    x_range: tuple[int, int],
    y_range: tuple[int, int],
    common_values: set[int] | None = None,
) -> str:
    candidates = []
    for word in words_in_region(words, x_range, y_range):
        value = money_to_int(word["text"])
        if value is not None:
            candidates.append((value, word["text"], word["left"]))
    if not candidates:
        return ""

    if common_values:
        for value, _, _ in candidates:
            if value in common_values:
                return str(value)

    plausible = [candidate for candidate in candidates if candidate[0] >= 100]
    if plausible:
        return str(max(plausible, key=lambda item: (item[2], item[0]))[0])
    return str(candidates[-1][0])


def correct_auction_fee(value: str, common_values: set[int]) -> str:
    number = money_to_int(value)
    if number is None:
        return ""
    if number in common_values:
        return str(number)
    if number <= 10 and 8000 in common_values:
        return "8000"
    if number in {1000, 10000, 110000} and 10000 in common_values:
        return "10000"
    if 10000 in common_values and abs(number - 10000) <= 10:
        return "10000"
    return str(number)


def clean_auction_car_name(text: str) -> str:
    text = normalize_spaces(text)
    text = re.sub(r"\b(Heide|Mage|Ps|aia|Sane|ae|ee|in)\b", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\b\d{4,7}\b", "", text)
    text = re.sub(r"\b[A-Z]{1,4}\d{1,3}[A-Z]?\b", "", text)
    text = text.replace('"', "").replace("'", "")
    replacements = {
        "yhA-T": "ゼットカーゴ",
        "9トカーゴ": "ゼットカーゴ",
        "デラ92ス": "デラックス",
        "デラ OD 92 ス": "デラックス",
        "デラ 92 ス": "デラックス",
        "ラリクス": "デラックス",
        "これがッ": "キャブバン",
        "これ が M ッ": "キャブバン M",
        "ネキャブ": "キャブ",
        "r7b399": "エルフトラック",
        "r7ト399": "エルフトラック",
        "ェルカト?9ク": "エルフトラック",
    }
    for before, after in replacements.items():
        text = text.replace(before, after)
    text = re.sub(r"[|()（）]", " ", text)
    text = re.sub(r"\b(Me|OD|MENT|SD|T)\b", "", text, flags=re.IGNORECASE)
    text = normalize_spaces(text)
    text = re.sub(r"ミニ\s*ミ", "ミニ", text)
    if "ゼットカーゴ" in text and "ハイ" not in text:
        text = "ハイ" + text
    if "ハイゼットカーゴ" in text and "5D" not in text:
        text = text.replace("ハイゼットカーゴ", "ハイゼットカーゴ5D")
    if "キャブバン" in text and "ミニ" not in text:
        text = "ミニ" + text
    if "エルフトラック" in text and "4D" not in text:
        text = f"{text} 4D"
    text = re.sub(r"ミニ\s*ミ", "ミニ", text)
    return text


def computed_tax(value: str) -> str:
    number = money_to_int(value)
    if number is None:
        return ""
    return str(round(number * 0.1))


def computed_auction_total(vehicle_amount: str, recycle_deposit: str, exhibit_fee: str, contract_fee: str) -> str:
    values = [money_to_int(value) for value in (vehicle_amount, recycle_deposit, exhibit_fee, contract_fee)]
    if any(value is None for value in values):
        return ""
    vehicle, recycle, exhibit, contract = values
    return str(vehicle + recycle - exhibit - contract)


def extract_auction_rows_from_words(words: list[dict[str, Any]], image_width: int, image_height: int) -> list[dict[str, str]]:
    def xr(start: int, end: int) -> tuple[int, int]:
        return scaled_range(start, end, image_width, AUCTION_BASE_WIDTH)

    def yr(start: int, end: int) -> tuple[int, int]:
        return scaled_range(start, end, image_height, AUCTION_BASE_HEIGHT)

    date_words = []
    for word in words:
        if word["left"] > image_width * 0.12:
            continue
        if not re.search(r"\d{2}/\d{2}/\d{2}", word["text"]):
            continue
        if word["top"] < image_height * 0.38 or word["top"] > image_height * 0.7:
            continue
        date_words.append(word)

    rows = []
    seen_y = []
    for date_word in sorted(date_words, key=lambda item: item["top"]):
        row_top = date_word["top"]
        if any(abs(row_top - existing) < image_height * 0.025 for existing in seen_y):
            continue

        lot_text = join_region_text(words, xr(205, 430), (row_top - 15, row_top + 35))
        lot_no = clean_auction_number(lot_text)
        if not lot_no:
            continue

        seen_y.append(row_top)
        main_band = (row_top - 18, row_top + 34)
        tax_band = (row_top + 28, row_top + 72)
        model_band = (row_top + 26, row_top + 75)
        car_text = join_region_text(words, xr(292, 730), main_band)
        model_text = join_region_text(words, xr(292, 730), model_band)
        year_text = join_region_text(words, xr(724, 780), model_band)

        model_match = re.search(r"[A-Z]{1,4}\d{1,3}[A-Z]{0,2}", model_text)
        frame_match = re.search(r"\b\d{5,8}\b", model_text)
        exhibit_fee = correct_auction_fee(best_money_in_region(words, xr(1225, 1355), main_band, {3000, 8000}), {3000, 8000})
        contract_fee = correct_auction_fee(best_money_in_region(words, xr(1355, 1485), main_band, {8000, 10000}), {8000, 10000})
        vehicle_amount = best_money_in_region(words, xr(1880, 2050), main_band)
        recycle_deposit = best_money_in_region(words, xr(2045, 2180), main_band)
        if money_to_int(vehicle_amount) and money_to_int(vehicle_amount) >= 500000 and contract_fee == "8000":
            contract_fee = "10000"
        exhibit_tax = computed_tax(exhibit_fee)
        contract_tax = computed_tax(contract_fee)
        vehicle_tax = computed_tax(vehicle_amount)
        total = computed_auction_total(vehicle_amount, recycle_deposit, exhibit_fee, contract_fee)
        if not total:
            total = best_money_in_region(words, xr(2335, 2510), main_band)
        ocr_line = join_region_text(words, xr(75, 2510), (row_top - 20, row_top + 75))

        rows.append(
            {
                "発生日": normalize_auction_date(date_word["text"]),
                "出品番号": lot_no,
                "車名": clean_auction_car_name(car_text),
                "型式": model_match.group(0) if model_match else "",
                "車台番号": frame_match.group(0) if frame_match else "",
                "年式": clean_auction_number(year_text),
                "出品料": exhibit_fee,
                "出品料税": exhibit_tax,
                "出品料税込": add_money_strings(exhibit_fee, exhibit_tax),
                "成約料": contract_fee,
                "成約料税": contract_tax,
                "成約料税込": add_money_strings(contract_fee, contract_tax),
                "車両金額": vehicle_amount,
                "車両金額税": vehicle_tax,
                "車両金額税込": add_money_strings(vehicle_amount, vehicle_tax),
                "R預託金": recycle_deposit,
                "合計": total,
                "OCR元行": ocr_line,
            }
        )

    return rows


def extract_auction_settlement(uploaded_file: Any) -> dict[str, Any]:
    file_suffix = Path(uploaded_file.name).suffix.lower().lstrip(".")
    result = {
        "rows": [],
        "words": [],
        "ocr_text": "",
        "error": None,
    }

    with tempfile.TemporaryDirectory(prefix="auction_ocr_", dir=Path.cwd()) as temp_dir_name:
        work_dir = Path(temp_dir_name)
        source_path = work_dir / f"auction_source.{file_suffix or 'bin'}"
        source_path.write_bytes(uploaded_file.getvalue())

        if file_suffix == "pdf":
            image_path, error = render_pdf_first_page(source_path, work_dir, size=2600)
            if error:
                result["error"] = error
                return result
        elif file_suffix in OCR_IMAGE_FILE_TYPES:
            image_path = work_dir / f"auction_image.{file_suffix}"
            shutil.copyfile(source_path, image_path)
        else:
            result["error"] = "PDFまたは画像ファイルを指定してください。"
            return result

        words, error = run_tesseract_tsv(image_path, work_dir)
        plain_text, plain_error = run_tesseract(image_path, work_dir)
        result["words"] = words
        result["ocr_text"] = normalize_ocr_text(plain_text)
        if error:
            result["error"] = error
            return result
        if plain_error and not result["ocr_text"]:
            result["error"] = plain_error

        try:
            from PIL import Image

            with Image.open(image_path) as image:
                width, height = image.size
        except Exception:
            width, height = AUCTION_BASE_WIDTH, AUCTION_BASE_HEIGHT

        result["rows"] = extract_auction_rows_from_words(words, width, height)
        if not result["rows"] and not result["error"]:
            result["error"] = "明細行を抽出できませんでした。PDFの向きや解像度を確認してください。"
        return result


def accounting_rows_to_tsv(rows: list[dict[str, Any]]) -> str:
    lines = ["\t".join(ACCOUNTING_FIELD_ORDER)]
    for row in rows:
        lines.append("\t".join(str(row.get(field, "")) for field in ACCOUNTING_FIELD_ORDER))
    return "\n".join(lines)


def format_accounting_date(date_text: str) -> str:
    normalized = normalize_spaces(date_text)
    match = re.search(r"(20\d{2})年\s*(\d{1,2})月\s*(\d{1,2})日", normalized)
    if match:
        _, month, day = match.groups()
        return f"{int(month)}/{int(day)}"

    def ocr_digit(value: str) -> int | None:
        cleaned = value.replace("Ｏ", "0").replace("O", "0").replace("o", "0")
        cleaned = cleaned.replace("Ｂ", "8").replace("B", "8").replace("b", "8")
        cleaned = re.sub(r"\D", "", cleaned)
        return int(cleaned) if cleaned else None

    date_normalized = normalized.replace("Ｒ", "R").replace("ー", "-")
    candidates = []
    for pattern in (
        r"[R]\s*[O0]?\s*([0-9OBob]{1,2})[./年/-]\s*([0-9OBob]{1,2})[./月/-]\s*([0-9OBob]{1,2})",
        r"(?<!\d)([0-9OBob]{1,2})[./年/-]\s*([0-9OBob]{1,2})[./月/-]\s*([0-9OBob]{1,2})(?!\d)",
    ):
        for match in re.finditer(pattern, date_normalized, flags=re.IGNORECASE):
            _year = ocr_digit(match.group(1))
            month = ocr_digit(match.group(2))
            day = ocr_digit(match.group(3))
            if month is None or day is None:
                continue
            if 1 <= month <= 12 and 1 <= day <= 31:
                candidates.append((match.start(), month, day))

    if candidates:
        _, month, day = sorted(candidates, key=lambda item: item[0])[-1]
        return f"{month}/{day}"

    match = re.search(r"(?<!\d)(\d{1,2})/(\d{1,2})(?!\d)", normalized)
    if match:
        month, day = match.groups()
        return f"{int(month)}/{int(day)}"
    return ""


def extract_best_amount(text: str, labels: list[str] | None = None) -> str:
    normalized = normalize_ocr_text(text)
    if labels:
        for label in labels:
            pattern = rf"{re.escape(label)}[^\d¥￥]*(?:[¥￥]\s*)?([\d,]{{3,}})"
            match = re.search(pattern, normalized)
            if match:
                return clean_auction_money(match.group(1))

    amounts = []
    for match in re.finditer(r"(?:[¥￥]\s*)?(\d{1,3}(?:,\d{3})+|\d{4,})(?:\s*円)?", normalized):
        amount = money_to_int(match.group(1))
        if amount is not None and 100 <= amount <= 10_000_000:
            amounts.append((amount, match.start(), "," in match.group(1)))
    if not amounts:
        return ""

    comma_amounts = [item for item in amounts if item[2]]
    if comma_amounts:
        return str(max(comma_amounts, key=lambda item: item[1])[0])
    return str(max(amounts, key=lambda item: item[1])[0])


def extract_known_car_name(text: str) -> str:
    normalized = normalize_spaces(text)
    normalized = normalized.replace("ヴウォクシー", "ヴォクシー")
    normalized = normalized.replace("ヴオクシー", "ヴォクシー")
    normalized = normalized.replace("ウオクシー", "ヴォクシー")
    normalized = normalized.replace("ウォクシー", "ヴォクシー")
    normalized = normalized.replace("ミニクーバー", "ミニクーパー")
    known_names = [
        "ヴォクシー",
        "ボクシー",
        "ポルテ",
        "アルファード",
        "ミニクーパー",
        "ハイゼットカーゴ",
        "ミニキャブ",
        "エルフトラック",
    ]
    for name in known_names:
        if name in normalized:
            return "ヴォクシー" if name == "ボクシー" else name
    return extract_car_name_from_context(normalized)


def is_plate_like_line(line: str) -> bool:
    normalized = normalize_spaces(line)
    has_area_number = re.search(r"[一-龥]{1,4}\s*\d{2,3}", normalized)
    has_kana_number = re.search(r"[ぁ-ん]\s*\d{2,4}", normalized)
    return bool(has_area_number and has_kana_number)


def clean_accounting_car_candidate(line: str) -> str:
    text = normalize_spaces(line)
    text = re.sub(r"[|_<>【】\[\]{}]", " ", text)
    text = re.sub(r"\b\d{1,6}\s*km\b", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"[一-龥]{1,4}\s*\d{2,3}\s*[ぁ-ん]\s*\d{1,4}", " ", text)
    makers = (
        "トヨタ",
        "ニッサン",
        "日産",
        "ホンダ",
        "ダイハツ",
        "スズキ",
        "マツダ",
        "ミツビシ",
        "三菱",
        "スバル",
        "レクサス",
        "BMW",
        "MINI",
    )
    for maker in makers:
        text = re.sub(rf"\b{re.escape(maker)}\b", " ", text, flags=re.IGNORECASE)
        text = text.replace(maker, " ")
    text = re.sub(r"^[^ァ-ヶA-Za-z0-9一-龥]+|[^ァ-ヶA-Za-z0-9一-龥ー]+$", "", text)
    return normalize_spaces(text)


def is_plausible_car_candidate(line: str) -> bool:
    text = clean_accounting_car_candidate(line)
    if not text or len(text) > 28:
        return False
    if not re.search(r"[ァ-ヶーA-Za-z]", text):
        return False
    excluded_words = (
        "請求書",
        "納品書",
        "御中",
        "株式会社",
        "ガリバー",
        "エムケイ",
        "タイヤ",
        "オイル",
        "ブレーキ",
        "点検",
        "技術料",
        "部品",
        "消費税",
        "合計",
        "登録番号",
        "TEL",
        "FAX",
    )
    return not any(word in text for word in excluded_words)


def extract_car_name_from_context(text: str) -> str:
    lines = normalize_ocr_text(text).splitlines()
    for index, line in enumerate(lines):
        if not is_plate_like_line(line):
            continue
        for nearby_line in lines[index + 1 : index + 4]:
            if "km" in nearby_line.lower():
                break
            if is_plausible_car_candidate(nearby_line):
                return clean_accounting_car_candidate(nearby_line)

    for index, line in enumerate(lines):
        if "km" not in line.lower():
            continue
        for nearby_line in reversed(lines[max(0, index - 3) : index]):
            if is_plausible_car_candidate(nearby_line):
                return clean_accounting_car_candidate(nearby_line)
    return ""


def build_accounting_row(
    date: str,
    amount: str,
    car_name: str,
    customer_name: str = "",
    sales: str = "",
    purchase: str = "",
    warranty: str = "",
    claim: str = "",
    customer_order: str = "",
    store_expense: str = "",
) -> dict[str, str]:
    return {
        "日付": date,
        "金額": amount,
        "車名": car_name,
        "お客様名": customer_name,
        "販売": sales,
        "買取": purchase,
        "保障": warranty,
        "クレーム": claim,
        "客注": customer_order,
        "店舗経費": store_expense,
    }


def money_values_in_text(text: str) -> list[int]:
    values = []
    normalized = normalize_spaces(text)
    for match in re.finditer(r"\d[\d,.\s]{1,18}\d", normalized):
        value = money_to_int(match.group(0))
        if value is not None and 100 <= value <= 10_000_000:
            values.append(value)
    return values


def last_money_value(text: str) -> int | None:
    values = money_values_in_text(text)
    return values[-1] if values else None


def extract_umemoto_amount(text: str) -> str:
    lines = normalize_ocr_text(text).splitlines()
    item_total = extract_umemoto_item_total(lines)
    if item_total:
        return str(item_total + round(item_total * 0.1))

    final_total = None
    subtotal = None
    tax = None

    for line in lines:
        if "消費税" in line:
            tax = last_money_value(line)
        if "小計" in line or ("摘要" in line and "計" in line):
            subtotal = last_money_value(line)
        if "担当" in line and "計" in line:
            final_total = last_money_value(line)
        elif "合計" in line or "合 計" in line:
            final_total = last_money_value(line)

    if subtotal and tax:
        return str(subtotal + tax)
    if final_total is not None:
        return str(final_total)
    return extract_best_amount(text, labels=["合計", "小計"])


def extract_umemoto_item_total(lines: list[str]) -> int:
    total = 0
    for line in lines:
        if "消費税" in line or "合計" in line or "担当" in line or "毎度" in line:
            break
        amount = extract_umemoto_item_amount(line)
        if amount is not None:
            total += amount
    return total


def extract_umemoto_item_amount(line: str) -> int | None:
    normalized = normalize_spaces(line)
    excluded = ("TEL", "FAX", "〒", "登録番号", "ガリバー", "ミスター", "タイヤマン", "御中")
    if any(word in normalized for word in excluded):
        return None
    item_keywords = (
        "脱着",
        "オイル",
        "フィルター",
        "アライメント",
        "工賃",
        "タイヤ",
        "交換",
        "バランス",
        "廃タイヤ",
        "バルブ",
        "パンク",
        "ローテーション",
    )
    if not any(keyword in normalized for keyword in item_keywords):
        return None
    if re.search(r"20\d{2}年", normalized):
        return None
    if re.fullmatch(r"\(?\d{6,9}\)?", normalized):
        return None

    comma_numbers = re.findall(r"\d{1,3},\d{3}", normalized)
    if comma_numbers:
        return int(comma_numbers[-1].replace(",", ""))

    tokens = re.findall(r"\d+", normalized)
    if len(tokens) >= 3:
        amount = int(tokens[-1])
        if 0 <= amount <= 100_000:
            return amount
    return None


def extract_mk_amount(text: str) -> str:
    lines = normalize_ocr_text(text).splitlines()
    for line in reversed(lines):
        if "御請求" in line or "請求額" in line:
            value = last_money_value(line)
            if value is not None:
                return str(value)

    taxable = None
    tax = None
    non_tax = None
    final_total = None
    tax_line_index = None
    for index, line in enumerate(lines):
        if "課税対象" in line:
            taxable = last_money_value(line)
        if "10%" in line or "10%)" in line:
            tax = last_money_value(line)
            tax_line_index = index
        if "非課税計" in line:
            values = money_values_in_text(line)
            if values:
                non_tax = values[0]
                if len(values) >= 2:
                    final_total = values[-1]

    if taxable is None and tax_line_index is not None:
        for line in reversed(lines[:tax_line_index]):
            values = [value for value in money_values_in_text(line) if value >= 1000]
            if values:
                taxable = values[-1]
                break

    if taxable is not None and tax is not None and non_tax is not None:
        return str(taxable + tax + non_tax)
    if final_total is not None:
        return str(final_total)

    excluded_patterns = ("km", "TEL", "携帯", "銀行", "登録番号", "T9", "〒", "住所")
    values = []
    for line in lines:
        if any(pattern in line for pattern in excluded_patterns):
            continue
        values.extend(money_values_in_text(line))
    plausible = [value for value in values if 1000 <= value <= 1_000_000]
    if plausible:
        return str(max(plausible))
    return extract_best_amount(text, labels=["総 計", "総計"])


def extract_accounting_text_from_upload(uploaded_file: Any) -> dict[str, Any]:
    extraction = extract_text_from_uploaded_file(uploaded_file)
    return {
        "text": extraction.get("text", ""),
        "error": extraction.get("error"),
        "source": extraction,
    }


def extract_accounting_document_from_upload(uploaded_file: Any, row_builder: Any) -> dict[str, Any]:
    file_suffix = Path(uploaded_file.name).suffix.lower().lstrip(".")
    file_bytes = uploaded_file.getvalue()
    rows = []
    page_texts = []
    errors = []

    if file_suffix in {"txt", "csv"}:
        text = normalize_ocr_text(extract_text_from_plain_file(file_bytes))
        return {
            "rows": keep_non_empty_accounting_rows([row_builder(page_text) for page_text in split_ocr_pages(text)]),
            "ocr_text": text,
            "error": None,
        }

    with tempfile.TemporaryDirectory(prefix="accounting_ocr_", dir=Path.cwd()) as temp_dir_name:
        work_dir = Path(temp_dir_name)
        source_path = work_dir / f"accounting_source.{file_suffix or 'bin'}"
        source_path.write_bytes(file_bytes)

        if file_suffix == "pdf":
            image_paths, error = render_pdf_pages(source_path, work_dir, size=3000)
            if error:
                return {"rows": [], "ocr_text": "", "error": error}
        elif file_suffix in OCR_IMAGE_FILE_TYPES:
            image_path = work_dir / f"accounting_image.{file_suffix}"
            shutil.copyfile(source_path, image_path)
            image_paths = [image_path]
        else:
            return {"rows": [], "ocr_text": "", "error": "PDFまたは画像ファイルを指定してください。"}

        for page_index, image_path in enumerate(image_paths, start=1):
            ocr_text, error = run_tesseract(image_path, work_dir)
            page_text = normalize_ocr_text(ocr_text)
            if error:
                errors.append(f"{page_index}ページ目: {error}")

            row = row_builder(page_text)
            targeted_text = ""
            if not row.get("車名"):
                targeted_text, _targeted_errors = extract_targeted_accounting_text(image_path, work_dir)
                if targeted_text:
                    page_text = normalize_ocr_text(f"{page_text}\n\n--- targeted car area ---\n{targeted_text}")
                    row = row_builder(page_text)

            rows.append(row)
            if page_text:
                page_texts.append(f"--- page {page_index} ---\n{page_text}")

    return {
        "rows": keep_non_empty_accounting_rows(rows),
        "ocr_text": "\n\n".join(page_texts),
        "error": " / ".join(errors) if errors else None,
    }


def split_ocr_pages(text: str) -> list[str]:
    normalized = normalize_ocr_text(text)
    if not normalized:
        return []

    parts = re.split(r"--- page \d+ ---", normalized)
    pages = [part.strip() for part in parts if part.strip()]
    return pages or [normalized]


def build_umemoto_row_from_text(text: str) -> dict[str, str]:
    return build_accounting_row(
        date=format_accounting_date(text),
        amount=extract_umemoto_amount(text),
        car_name=extract_known_car_name(text),
    )


def build_mk_row_from_text(text: str) -> dict[str, str]:
    customer_order = "車検" if "車検" in text or "Car" in text or "Cagr" in text else ""
    return build_accounting_row(
        date=format_accounting_date(text),
        amount=extract_mk_amount(text),
        car_name=extract_known_car_name(text),
        customer_order=customer_order,
    )


def keep_non_empty_accounting_rows(rows: list[dict[str, str]]) -> list[dict[str, str]]:
    return [row for row in rows if any(row.values())]


def extract_umemoto_delivery(uploaded_file: Any) -> dict[str, Any]:
    return extract_accounting_document_from_upload(uploaded_file, build_umemoto_row_from_text)


def extract_mk_invoice(uploaded_file: Any) -> dict[str, Any]:
    return extract_accounting_document_from_upload(uploaded_file, build_mk_row_from_text)


def render_receipt_file_uploader() -> tuple[list[Any], dict[str, Any]]:
    st.subheader("領収書ファイル")
    uploaded_files = st.file_uploader(
        "PDF / 画像 / テキストなど（任意）",
        type=UPLOAD_FILE_TYPES,
        accept_multiple_files=True,
        help="PDFや画像はTesseract OCRで読み取り、入力欄の候補にします。",
    )

    summaries = build_uploaded_file_summaries(uploaded_files)
    extraction = {
        "results": [],
        "text": "",
        "candidates": build_receipt_data("", "", "", ""),
    }
    if summaries:
        st.caption("アップロード済みファイル")
        st.dataframe(summaries, width="stretch", hide_index=True)

        with st.spinner("PDF/画像から文字を読み取っています..."):
            extraction = extract_receipt_from_uploaded_files(uploaded_files or [])

        for result in extraction["results"]:
            if result.get("error"):
                st.warning(f"{result['file_name']}: {result['error']}")

        if extraction["text"]:
            st.success("OCR結果から入力候補を作成しました。必要に応じて下の手入力欄で直してください。")
            with st.expander("OCRで読み取ったテキスト", expanded=False):
                st.text_area("OCRテキスト", value=extraction["text"], height=260)
        else:
            st.warning("文字を読み取れませんでした。スキャン画像の品質やOCR環境を確認してください。")

    return uploaded_files or [], extraction


def render_sidebar() -> dict[str, Any]:
    with st.sidebar:
        st.subheader("領収書モード")
        mode = st.radio(
            "モード",
            [MODE_MANUAL, MODE_DUMMY, MODE_LOCAL_LLM],
            index=0,
        )

        st.subheader("local LLM 接続")
        use_llm = st.toggle("LLM使用", value=True)
        base_url = st.text_input("Base URL", value=DEFAULT_BASE_URL)
        model = st.text_input("Model名", value=DEFAULT_MODEL)
        timeout_seconds = st.number_input(
            "タイムアウト秒数",
            min_value=1.0,
            max_value=120.0,
            value=30.0,
            step=1.0,
        )
        debug = st.toggle("デバッグ表示", value=False)

        st.caption("OpenAI互換の `/v1/chat/completions` へ送信します。")

    return {
        "mode": mode,
        "use_llm": use_llm,
        "base_url": base_url,
        "model": model,
        "timeout_seconds": timeout_seconds,
        "debug": debug,
    }


def format_receipt_by_mode(receipt: dict[str, str], settings: dict[str, Any]) -> tuple[dict[str, str], dict[str, Any]]:
    mode = settings["mode"]
    debug_info: dict[str, Any] = {
        "mode": mode,
        "base_url": settings["base_url"],
        "model": settings["model"],
        "payload": None,
        "llm_attempts": [],
        "raw_response": None,
        "parsed": None,
        "error": None,
        "fallback": None,
    }

    if mode == MODE_MANUAL:
        return receipt, debug_info

    if mode == MODE_DUMMY:
        return dummy_format_receipt(receipt), debug_info

    if not settings["use_llm"]:
        debug_info["fallback"] = "LLM使用がOFFのため、ダミー整形を使用しました。"
        return dummy_format_receipt(receipt), debug_info

    llm_result = call_local_llm(
        base_url=settings["base_url"],
        model=settings["model"],
        receipt=receipt,
        timeout_seconds=settings["timeout_seconds"],
    )
    debug_info["payload"] = llm_result.get("payload")
    debug_info["llm_attempts"] = llm_result.get("attempts")
    debug_info["raw_response"] = llm_result.get("raw_response")
    debug_info["parsed"] = llm_result.get("parsed")
    debug_info["error"] = llm_result.get("error")

    if llm_result.get("parsed"):
        return coerce_receipt_fields(llm_result["parsed"], receipt), debug_info

    debug_info["fallback"] = "LLM整形に失敗したため、ダミー整形を使用しました。"
    return dummy_format_receipt(receipt), debug_info


def render_output(receipt: dict[str, str], tsv_text: str, debug_info: dict[str, Any], debug: bool) -> None:
    st.subheader("整形後の項目")
    for field in FIELD_ORDER:
        st.write(f"**{FIELD_LABELS[field]}**: {receipt.get(field, '')}")

    st.subheader("スプレッドシート貼り付け用 TSV 1行")
    st.code(tsv_text, language="tsv")

    st.text_area(
        "コピー用テキスト",
        value=tsv_text,
        height=120,
        help="ここを丸ごとコピーして、スプレッドシートへ貼り付けてください。",
    )

    if debug_info.get("error"):
        st.warning(debug_info["error"])
    if debug_info.get("fallback"):
        st.info(debug_info["fallback"])

    if debug:
        st.subheader("デバッグ")
        st.write("使用モード", debug_info.get("mode"))
        st.write("Base URL", debug_info.get("base_url"))
        st.write("model名", debug_info.get("model"))
        st.write("アップロードファイル")
        st.json(debug_info.get("uploaded_files") or [])
        st.write("OCR結果")
        st.json(debug_info.get("ocr_results") or [])
        st.write("OCRテキスト")
        st.code(debug_info.get("ocr_text") or "")
        st.write("送信payload")
        st.json(debug_info.get("payload") or {})
        st.write("LLM試行履歴")
        st.json(debug_info.get("llm_attempts") or [])
        st.write("生レスポンス")
        st.code(debug_info.get("raw_response") or "")
        st.write("パース結果")
        st.json(debug_info.get("parsed") or {})
        st.write("エラー内容")
        st.code(debug_info.get("error") or "")


def render_receipt_ui(settings: dict[str, Any]) -> None:
    st.title("領収書入力")
    st.caption("手入力した領収書データを整形し、スプレッドシートへ貼り付けやすいTSV 1行を作成します。")

    uploaded_files, extraction = render_receipt_file_uploader()
    candidates = extraction["candidates"]

    with st.form("receipt_form"):
        date = st.text_input("日付", value=candidates.get("date", ""), placeholder="例: 2026/04/18")
        payee = st.text_input("支払先", value=candidates.get("payee", ""), placeholder="例: 株式会社サンプル")
        amount = st.text_input("金額", value=candidates.get("amount", ""), placeholder="例: ¥1,200")
        description = st.text_input("内容", value=candidates.get("description", ""), placeholder="例: 文房具")
        submitted = st.form_submit_button("貼り付け用データを作成", width="stretch")

    if not submitted:
        st.info("入力後に「貼り付け用データを作成」を押してください。")
        return

    receipt = build_receipt_data(date, payee, amount, description)
    formatted_receipt, debug_info = format_receipt_by_mode(receipt, settings)
    debug_info["uploaded_files"] = build_uploaded_file_summaries(uploaded_files)
    debug_info["ocr_results"] = extraction["results"]
    debug_info["ocr_text"] = extraction["text"]
    tsv_text = receipt_to_tsv(formatted_receipt)
    render_output(formatted_receipt, tsv_text, debug_info, settings["debug"])


def render_auction_ui(settings: dict[str, Any]) -> None:
    st.caption("精算書PDFから明細行をOCRで読み取り、スプレッドシートへ貼り付けやすいTSVを作成します。手書きメモは抽出対象外です。")

    source_file = st.file_uploader(
        "オークション精算書PDF / 画像",
        type=["pdf", "png", "jpg", "jpeg", "webp"],
        accept_multiple_files=False,
        help="USSなどの精算書をアップロードしてください。",
    )
    destination_file = st.file_uploader(
        "転記先サンプルPDF（任意）",
        type=["pdf", "png", "jpg", "jpeg", "webp"],
        accept_multiple_files=False,
        help="今は列構成の参考として受け付けます。抽出処理には使いません。",
    )
    if destination_file:
        st.caption(f"転記先サンプル: {destination_file.name}")

    if not source_file:
        st.info("精算書PDFをアップロードしてください。")
        return

    with st.spinner("精算書をOCRで読み取っています..."):
        extraction = extract_auction_settlement(source_file)

    if extraction.get("error"):
        st.warning(extraction["error"])

    rows = extraction.get("rows") or []
    if not rows:
        with st.expander("OCRテキスト", expanded=False):
            st.text_area("OCRテキスト", value=extraction.get("ocr_text", ""), height=300)
        return

    st.success(f"{len(rows)}件の明細候補を抽出しました。OCR誤読があるので、貼り付け前に表で直してください。")
    st.caption("例: 車名の誤読、型式、金額の桁を確認してください。手書きメモはOCR元行からも除外しきれない場合があります。")

    edited_df = st.data_editor(
        pd.DataFrame(rows, columns=AUCTION_FIELD_ORDER),
        width="stretch",
        hide_index=True,
        num_rows="dynamic",
        key="auction_rows_editor",
    )
    edited_rows = edited_df.fillna("").to_dict("records")
    tsv_text = auction_to_tsv(edited_rows)

    st.subheader("スプレッドシート貼り付け用TSV")
    st.code(tsv_text, language="tsv")
    st.text_area(
        "コピー用テキスト",
        value=tsv_text,
        height=220,
        help="表を確認・修正したあと、ここをコピーしてスプレッドシートへ貼り付けてください。",
    )

    with st.expander("OCRで読み取ったテキスト", expanded=False):
        st.text_area("OCRテキスト", value=extraction.get("ocr_text", ""), height=300)

    if settings.get("debug"):
        st.subheader("デバッグ")
        st.write("OCR単語数", len(extraction.get("words") or []))
        st.json(extraction.get("words", [])[:120])


def render_accounting_document_ui(
    title: str,
    uploader_label: str,
    extractor: Any,
    settings: dict[str, Any],
) -> None:
    st.caption("納品書.転記先サンプルの列順に合わせて、日付・金額・車名などをTSV化します。手書きメモは基本的に抽出対象外です。")

    uploaded_file = st.file_uploader(
        uploader_label,
        type=["pdf", "png", "jpg", "jpeg", "webp"],
        accept_multiple_files=False,
        key=f"{title}_uploader",
    )

    reference_file = st.file_uploader(
        "転記先サンプルPDF（任意）",
        type=["pdf", "png", "jpg", "jpeg", "webp"],
        accept_multiple_files=False,
        key=f"{title}_reference",
        help="今は列構成の参考として受け付けます。抽出処理には使いません。",
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
        st.warning("転記候補を抽出できませんでした。")
        with st.expander("OCRテキスト", expanded=False):
            st.text_area("OCRテキスト", value=extraction.get("ocr_text", ""), height=300, key=f"{title}_empty_ocr")
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
    tsv_text = accounting_rows_to_tsv(edited_rows)

    st.subheader("スプレッドシート貼り付け用TSV")
    st.code(tsv_text, language="tsv")
    st.text_area(
        "コピー用テキスト",
        value=tsv_text,
        height=160,
        help="表を確認・修正したあと、ここをコピーしてスプレッドシートへ貼り付けてください。",
        key=f"{title}_copy",
    )

    with st.expander("OCRで読み取ったテキスト", expanded=False):
        st.text_area("OCRテキスト", value=extraction.get("ocr_text", ""), height=300, key=f"{title}_ocr")

    if settings.get("debug"):
        st.subheader("デバッグ")
        st.json(extraction)


def render_ui() -> None:
    st.set_page_config(
        page_title="領収書・精算書入力",
        page_icon="🧾",
        layout="wide",
    )
    settings = render_sidebar()
    st.title("書類転記アプリ")

    auction_tab, umemoto_tab, mk_tab = st.tabs([TAB_AUCTION, TAB_UMEMOTO, TAB_MK])
    with auction_tab:
        st.subheader(TAB_AUCTION)
        render_auction_ui(settings)
    with umemoto_tab:
        st.subheader(TAB_UMEMOTO)
        render_accounting_document_ui(
            title=TAB_UMEMOTO,
            uploader_label="ウメモト納品書PDF / 画像",
            extractor=extract_umemoto_delivery,
            settings=settings,
        )
    with mk_tab:
        st.subheader(TAB_MK)
        render_accounting_document_ui(
            title=TAB_MK,
            uploader_label="MK石油請求書PDF / 画像",
            extractor=extract_mk_invoice,
            settings=settings,
        )


if __name__ == "__main__":
    render_ui()
