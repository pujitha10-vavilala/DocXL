from __future__ import annotations

import re
from pathlib import Path
from uuid import uuid4
from typing import Any

import pymupdf
import pdfplumber
import statistics
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = FastAPI(
    title="DocXL",
    description="Convert PDF documents into structured Excel workbooks.",
    version="2.3.0",
)

from fastapi.middleware.cors import CORSMiddleware
app.add_middleware(
    CORSMiddleware,
    allow_origins= ["https://doc-xl.vercel.app"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE_DIR = Path(__file__).resolve().parent.parent
UPLOAD_DIR = BASE_DIR / "uploads"
GENERATED_DIR = BASE_DIR / "generated"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
GENERATED_DIR.mkdir(parents=True, exist_ok=True)

MAX_FILE_SIZE = 25 * 1024 * 1024
MAX_PAGES = 50

# Canonicalize common variations. Unknown headers are preserved rather than guessed.
HEADER_ALIASES = {
    "qty": "Quantity", "qty.": "Quantity", "quantity": "Quantity",
    "desc": "Description", "item description": "Description",
    "unit price": "Unit Price", "unit cost": "Unit Cost",
    "item no": "Item No", "item number": "Item No", "item #": "Item No",
    "serial no": "Serial No", "serial number": "Serial No",
    "s no": "S.No", "s.no": "S.No", "s.no.": "S.No",
    "date of transaction": "Date", "transaction date": "Date",
    "customer name": "Customer", "client name": "Customer",
    "product name": "Product", "item name": "Item Name",
    "phone number": "Phone", "telephone": "Phone",
    "e-mail": "Email", "e mail": "Email",
    "opening balance": "Opening Balance", "closing balance": "Closing Balance",
}

def clean_cell(value: Any) -> str:
    if value is None:
        return ""
    value = str(value).replace("\u00a0", " ").replace("\u200b", "")
    value = value.replace("\r", " ").replace("\n", " ")
    return re.sub(r"\s+", " ", value).strip()

def header_key(value: Any) -> str:
    value = clean_cell(value).casefold()
    value = re.sub(r"[\s\-_./:#]+", " ", value).strip()
    return value

def normalize_header(row: list[str]) -> list[str]:
    result = []
    for cell in row:
        value = clean_cell(cell)
        key = header_key(value)
        result.append(HEADER_ALIASES.get(key, value))
    # Repair a few fused cells only when the fragments are unambiguous.
    # Do not force a fixed schema or column count.
    repaired = []
    for value in result:
        key = header_key(value)
        if key == "opening stockreceived":
            repaired.extend(["Opening Stock", "Received"])
        elif key == "closing stockunit cost":
            repaired.extend(["Closing Stock", "Unit Cost"])
        elif key == "damaged closing stockunit cost":
            repaired.extend(["Damaged", "Closing Stock", "Unit Cost"])
        else:
            repaired.append(value)
    return repaired

def clean_rows(raw_table: list[list[Any]]) -> list[list[str]]:
    rows = [[clean_cell(cell) for cell in (row or [])] for row in raw_table]
    rows = [row for row in rows if any(row)]
    if not rows:
        return []
    width = max(len(row) for row in rows)
    rows = [row + [""] * (width - len(row)) for row in rows]
    # Drop empty outer columns only; internal blank cells are meaningful.
    while rows and rows[0] and all(not row[0] for row in rows):
        rows = [row[1:] for row in rows]
    while rows and rows[0] and all(not row[-1] for row in rows):
        rows = [row[:-1] for row in rows]
    return rows

HEADER_WORDS = {
    "id", "no", "number", "name", "date", "address", "amount", "price",
    "quantity", "qty", "description", "status", "total", "email", "phone",
    "sku", "product", "item", "category", "country", "city", "stock",
    "revenue", "department", "employee", "customer", "order", "code",
    "value", "type", "rate", "balance", "unit", "location", "result",
    "owner", "priority", "notes", "check", "invoice", "account", "region",
}

def is_header_row(row: list[str]) -> bool:
    values = [header_key(x) for x in row if clean_cell(x)]
    if len(values) < 2:
        return False
    # Headers are usually text-heavy and contain at least one field-like word.
    word_hits = sum(
        1 for value in values
        if any(re.search(rf"\b{re.escape(word)}\b", value) for word in HEADER_WORDS)
    )
    numeric = sum(bool(re.fullmatch(r"[-+]?\d+(?:[.,]\d+)?%?", v)) for v in values)
    return word_hits >= 1 and numeric <= len(values) * 0.4

def row_signature(row: list[str]) -> tuple[str, ...]:
    return tuple(re.sub(r"[^a-z0-9]+", "", header_key(x)) for x in row)

def header_similarity(a: list[str], b: list[str]) -> float:
    if not a or not b or len(a) != len(b):
        return 0.0
    aa, bb = row_signature(a), row_signature(b)
    if aa == bb:
        return 1.0
    comparable = [(x, y) for x, y in zip(aa, bb) if x or y]
    if not comparable:
        return 0.0
    return sum(x == y for x, y in comparable) / len(comparable)

def headers_match(a: list[str], b: list[str]) -> bool:
    # Exact match is preferred; allow small extraction noise only for wide tables.
    score = header_similarity(a, b)
    return score >= (0.75 if len(a) >= 5 else 0.85)


# OCR is loaded only when needed, so digital-PDF extraction remains lightweight.
_OCR_ENGINE = None
_OCR_INIT_ERROR: str | None = None


def get_ocr_engine():
    """Initialize PaddleOCR lazily and tolerate common PaddleOCR API versions."""
    global _OCR_ENGINE, _OCR_INIT_ERROR
    if _OCR_ENGINE is not None:
        return _OCR_ENGINE
    if _OCR_INIT_ERROR is not None:
        raise RuntimeError(_OCR_INIT_ERROR)

    try:
        from paddleocr import PaddleOCR

        # PaddleOCR 2.x and 3.x use different constructor options.
        attempts = [
            {"use_angle_cls": True, "lang": "en", "show_log": False},
            {"use_textline_orientation": True, "lang": "en"},
            {"lang": "en"},
        ]
        last_error = None
        for kwargs in attempts:
            try:
                _OCR_ENGINE = PaddleOCR(**kwargs)
                return _OCR_ENGINE
            except (TypeError, ValueError) as exc:
                last_error = exc
        raise RuntimeError(f"Could not initialize PaddleOCR: {last_error}")
    except Exception as exc:
        _OCR_INIT_ERROR = str(exc)
        raise RuntimeError(f"PaddleOCR is installed but could not start: {exc}") from exc


def _parse_ocr_result(result) -> list[tuple[list[float], str, float]]:
    """Normalize PaddleOCR 2.x/3.x outputs to (box, text, confidence)."""
    found: list[tuple[list[float], str, float]] = []

    def add(box, text, confidence=0.0):
        if not text or not str(text).strip():
            return
        try:
            points = [[float(p[0]), float(p[1])] for p in box]
            xs = [p[0] for p in points]
            ys = [p[1] for p in points]
            rect = [min(xs), min(ys), max(xs), max(ys)]
            found.append((rect, clean_cell(text), float(confidence or 0.0)))
        except Exception:
            return

    def walk(obj):
        if obj is None:
            return
        if isinstance(obj, dict):
            texts = obj.get("rec_texts") or obj.get("texts") or []
            scores = obj.get("rec_scores") or obj.get("scores") or []
            boxes = obj.get("rec_polys") or obj.get("dt_polys") or obj.get("boxes") or []
            for i, text in enumerate(texts):
                if i < len(boxes):
                    add(boxes[i], text, scores[i] if i < len(scores) else 0.0)
            return
        if isinstance(obj, (list, tuple)):
            # PaddleOCR 2.x line: [box, (text, confidence)]
            if len(obj) >= 2 and isinstance(obj[1], (list, tuple)) and len(obj[1]) >= 1:
                box, rec = obj[0], obj[1]
                if isinstance(rec[0], str):
                    add(box, rec[0], rec[1] if len(rec) > 1 else 0.0)
                    return
            for child in obj:
                walk(child)

    walk(result)
    return found


def _group_ocr_words_into_rows(words: list[tuple[list[float], str, float]]) -> list[list[tuple[list[float], str, float]]]:
    if not words:
        return []
    words = sorted(words, key=lambda item: ((item[0][1] + item[0][3]) / 2, item[0][0]))
    heights = [max(1.0, box[3] - box[1]) for box, _, _ in words]
    y_tolerance = max(5.0, statistics.median(heights) * 0.65)
    rows: list[list[tuple[list[float], str, float]]] = []
    row_centers: list[float] = []
    for word in words:
        box = word[0]
        center_y = (box[1] + box[3]) / 2
        best = min(range(len(row_centers)), key=lambda i: abs(row_centers[i] - center_y), default=None)
        if best is not None and abs(row_centers[best] - center_y) <= y_tolerance:
            rows[best].append(word)
            centers = [(w[0][1] + w[0][3]) / 2 for w in rows[best]]
            row_centers[best] = sum(centers) / len(centers)
        else:
            rows.append([word])
            row_centers.append(center_y)
    for row in rows:
        row.sort(key=lambda item: item[0][0])
    return rows


def ocr_page_as_table(page, page_number: int) -> list[list[str]]:
    """OCR a page and recover table cells using detected grid lines when possible.

    The grid-based path is layout-driven (not tied to a fixed number of columns).
    If grid detection is unavailable, it falls back to the previous gap-based method.
    """
    import numpy as np

    engine = get_ocr_engine()
    pix = page.get_pixmap(matrix=pymupdf.Matrix(2, 2), alpha=False)
    image = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
    if pix.n == 4:
        image = image[:, :, :3]
    else:
        image = image[:, :, :3].copy()

    try:
        result = engine.ocr(image, cls=True)
    except TypeError:
        result = engine.ocr(image)
    words = _parse_ocr_result(result)
    if len(words) < 6:
        return []

    grouped = _group_ocr_words_into_rows(words)
    if len(grouped) < 3:
        return []

    # Try to find table grid lines in the rendered page. This supports different
    # table widths because the column count is inferred from the image itself.
    try:
        import cv2

        gray = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)
        binary = cv2.threshold(
            gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU
        )[1]
        height, width = binary.shape

        horizontal_kernel = cv2.getStructuringElement(
            cv2.MORPH_RECT, (max(25, width // 35), 1)
        )
        vertical_kernel = cv2.getStructuringElement(
            cv2.MORPH_RECT, (1, max(20, height // 45))
        )
        horizontal = cv2.morphologyEx(binary, cv2.MORPH_OPEN, horizontal_kernel)
        vertical = cv2.morphologyEx(binary, cv2.MORPH_OPEN, vertical_kernel)

        # Project line pixels; merge nearby detections into one boundary.
        x_projection = (vertical > 0).sum(axis=0)
        y_projection = (horizontal > 0).sum(axis=1)
        x_threshold = max(8, int(height * 0.012))
        y_threshold = max(8, int(width * 0.012))

        def runs(values, threshold, min_gap=3):
            indexes = np.flatnonzero(values >= threshold)
            if len(indexes) == 0:
                return []
            groups = [[int(indexes[0])]]
            for value in indexes[1:]:
                value = int(value)
                if value - groups[-1][-1] <= min_gap:
                    groups[-1].append(value)
                else:
                    groups.append([value])
            return [int(round(sum(group) / len(group))) for group in groups]

        x_lines = runs(x_projection, x_threshold, 4)
        y_lines = runs(y_projection, y_threshold, 4)

        # Require a plausible grid: at least two vertical boundaries and
        # multiple horizontal boundaries overlapping the same table area.
        if len(x_lines) >= 3 and len(y_lines) >= 4:
            # Choose the densest rectangular grid region based on OCR coverage.
            # Remove lines that are far outside the OCR content bounds.
            min_word_x = min(box[0] for box, _, _ in words)
            max_word_x = max(box[2] for box, _, _ in words)
            min_word_y = min(box[1] for box, _, _ in words)
            max_word_y = max(box[3] for box, _, _ in words)
            xs = [x for x in x_lines if min_word_x - 30 <= x <= max_word_x + 30]
            ys = [y for y in y_lines if min_word_y - 35 <= y <= max_word_y + 35]

            # Lines are only trusted if enough of the OCR content falls inside
            # the detected grid. The fallback below handles borderless tables.
            if len(xs) >= 3 and len(ys) >= 4:
                cells_by_row = []
                for top, bottom in zip(ys, ys[1:]):
                    if bottom - top < 8:
                        continue
                    row_cells = [""] * (len(xs) - 1)
                    row_has_text = False
                    for box, text, confidence in words:
                        cx = (box[0] + box[2]) / 2
                        cy = (box[1] + box[3]) / 2
                        if not (top - 2 <= cy < bottom + 2):
                            continue
                        col = next(
                            (i for i in range(len(xs) - 1)
                             if xs[i] - 2 <= cx < xs[i + 1] + 2),
                            None,
                        )
                        if col is None:
                            continue
                        if row_cells[col]:
                            row_cells[col] += " " + text
                        else:
                            row_cells[col] = text
                        row_has_text = True
                    if row_has_text:
                        cells_by_row.append([clean_cell(v) for v in row_cells])

                if len(cells_by_row) >= 3:
                    # Trim only fully empty outer columns, preserving interior blanks.
                    while cells_by_row and all(not row[0] for row in cells_by_row):
                        cells_by_row = [row[1:] for row in cells_by_row]
                    while cells_by_row and all(not row[-1] for row in cells_by_row):
                        cells_by_row = [row[:-1] for row in cells_by_row]
                    if cells_by_row and len(cells_by_row[0]) >= 2:
                        return cells_by_row
    except Exception:
        # Grid detection is an enhancement; OCR should still work without OpenCV.
        pass

    # Fallback for tables without visible grid lines: group OCR items by horizontal
    # gaps, retaining the existing best-effort behavior.
    heights = [max(1.0, w[0][3] - w[0][1]) for row in grouped for w in row]
    median_height = statistics.median(heights) if heights else 10.0
    gap_threshold = max(18.0, median_height * 2.0)

    recovered: list[list[str]] = []
    for row in grouped:
        cells: list[str] = []
        current = row[0][1]
        previous_right = row[0][0][2]
        for box, text, _ in row[1:]:
            gap = box[0] - previous_right
            if gap >= gap_threshold:
                cells.append(current.strip())
                current = text
            else:
                current = f"{current} {text}".strip()
            previous_right = max(previous_right, box[2])
        cells.append(current.strip())
        recovered.append(cells)

    multi_column = [row for row in recovered if len(row) >= 2]
    if len(multi_column) < 3:
        return []
    widths = [len(row) for row in multi_column]
    common_width = max(set(widths), key=widths.count)
    consistent = sum(width == common_width for width in widths) / len(widths)
    if common_width < 2 or consistent < 0.6:
        return []
    return [
        row + [""] * (common_width - len(row)) if len(row) <= common_width
        else row[:common_width]
        for row in multi_column
    ]


def extract_tables_from_pdf(pdf_path: Path):
    tables: list[dict[str, Any]] = []
    pages: list[dict[str, Any]] = []
    warnings: list[str] = []

    # PyMuPDF renders only pages that need OCR; pdfplumber remains the primary path.
    with pymupdf.open(pdf_path) as fitz_doc, pdfplumber.open(pdf_path) as pdf:
        if len(pdf.pages) > MAX_PAGES:
            raise HTTPException(
                status_code=400,
                detail=f"PDFs can contain at most {MAX_PAGES} pages.",
            )
        for page_index, page in enumerate(pdf.pages):
            page_number = page_index + 1
            try:
                page_text = page.extract_text() or ""
            except Exception as exc:
                page_text = ""
                warnings.append(f"Text extraction issue on page {page_number}: {exc}")

            page_tables = []
            try:
                raw_tables = page.extract_tables() or []
            except Exception as exc:
                warnings.append(f"Table extraction issue on page {page_number}: {exc}")
                raw_tables = []

            for table_number, raw_table in enumerate(raw_tables, start=1):
                rows = clean_rows(raw_table or [])
                if not rows:
                    continue
                rows[0] = normalize_header(rows[0])
                width = max(len(row) for row in rows)
                rows = [row[:width] + [""] * max(0, width - len(row)) for row in rows]
                page_tables.append({
                    "page_number": page_number,
                    "table_number": table_number,
                    "rows": rows,
                    "column_count": width,
                    "has_header": is_header_row(rows[0]),
                    "source": "pdfplumber",
                })

            ocr_used = False
            # OCR only when native extraction found no table or no selectable text.
            # This keeps digital PDFs on the existing, more reliable path.
            if not page_tables or len(page_text.strip()) < 10:
                try:
                    ocr_rows = ocr_page_as_table(fitz_doc[page_index], page_number)
                    if ocr_rows and not page_tables:
                        ocr_rows = clean_rows(ocr_rows)
                        if ocr_rows:
                            ocr_rows[0] = normalize_header(ocr_rows[0])
                            width = max(len(row) for row in ocr_rows)
                            ocr_rows = [row[:width] + [""] * max(0, width - len(row)) for row in ocr_rows]
                            page_tables.append({
                                "page_number": page_number,
                                "table_number": 1,
                                "rows": ocr_rows,
                                "column_count": width,
                                "has_header": is_header_row(ocr_rows[0]),
                                "source": "paddleocr",
                            })
                            ocr_used = True
                    elif not page_text.strip():
                        # Even when OCR cannot confidently infer a table, preserve OCR text
                        # in the page-level text export rather than silently losing it.
                        if not ocr_rows:
                            # Recover plain OCR text with a less restrictive parser.
                            pix = fitz_doc[page_index].get_pixmap(matrix=pymupdf.Matrix(2, 2), alpha=False)
                            import numpy as np
                            image = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
                            if pix.n == 4:
                                image = image[:, :, :3]
                            engine = get_ocr_engine()
                            try:
                                result = engine.ocr(image, cls=True)
                            except TypeError:
                                result = engine.ocr(image)
                            words = _parse_ocr_result(result)
                            ocr_rows = [[word[1] for word in row] for row in _group_ocr_words_into_rows(words)]
                        if ocr_rows:
                            page_text = "\n".join(" ".join(row) for row in ocr_rows)
                            ocr_used = True
                except Exception as exc:
                    warnings.append(f"OCR issue on page {page_number}: {exc}")

            pages.append({
                "page_number": page_number,
                "text": page_text,
                "ocr_used": ocr_used,
            })
            tables.extend(page_tables)

    return tables, pages, warnings

def merge_continuation_tables(tables: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Merge tables when normalized headers identify the same logical table.

    This intentionally does not merge tables merely because they have equal widths.
    It supports page gaps and intervening unrelated tables, but requires a strong
    header match to avoid combining unrelated data.
    """
    merged: list[dict[str, Any]] = []

    for item in tables:
        rows = [row[:] for row in item["rows"]]
        if not rows:
            continue
        header = rows[0]
        match_index = None

        # Search earlier tables from newest to oldest for a matching header.
        # Do not stop at a different table: a continued table may be separated
        # by an unrelated table or by a page on which it was not detected.
        for idx in range(len(merged) - 1, -1, -1):
            candidate = merged[idx]
            if candidate["column_count"] != len(header):
                continue
            if headers_match(candidate["rows"][0], header):
                match_index = idx
                break

        if match_index is not None:
            target = merged[match_index]
            # The matching first row is a repeated header, so remove it.
            rows = rows[1:]
            if rows:
                target["rows"].extend(rows)
            target["last_page"] = max(target["last_page"], item["page_number"])
            target["source_pages"].append(item["page_number"])
        else:
            merged.append({
                "rows": rows,
                "column_count": len(rows[0]),
                "first_page": item["page_number"],
                "last_page": item["page_number"],
                "source_pages": [item["page_number"]],
                "has_header": item["has_header"],
            })

    return merged

def safe_sheet_name(name: str, used: set[str]) -> str:
    base = re.sub(r"[\\/*?:\[\]]", " ", name).strip() or "Table"
    base = base[:31]
    candidate = base
    counter = 2
    while candidate.casefold() in used:
        suffix = f"_{counter}"
        candidate = base[:31 - len(suffix)] + suffix
        counter += 1
    used.add(candidate.casefold())
    return candidate

def suggested_sheet_name(rows: list[list[str]], index: int, used: set[str]) -> str:
    if rows and rows[0]:
        first = header_key(rows[0][0])
        # Keep a useful name for common inventory data without depending on it for merging.
        if first == "sku":
            return safe_sheet_name("Inventory", used)
        # Use a short descriptive name only when the first header is a clear label.
        label = clean_cell(rows[0][0])
        if label and len(label) <= 22 and re.search(r"[A-Za-z]", label):
            candidate = re.sub(r"[^A-Za-z0-9 _-]", "", label).strip()
            if candidate and candidate.casefold() not in used:
                return safe_sheet_name(candidate, used)
    return safe_sheet_name(f"Table_{index}", used)

def excel_value(value: Any, header: Any) -> Any:
    """Convert only clearly typed columns; keep identifiers and ambiguous values as text."""
    text = clean_cell(value)
    if not text:
        return None

    key = header_key(header)
    compact = re.sub(r"[^a-z0-9]+", "", key)

    # Identifiers/codes must remain text (e.g. 00123, invoice IDs, SKUs).
    identifier_tokens = (
        "id", "identifier", "sku", "code", "serial", "phone", "telephone",
        "zip", "postal", "account", "case", "record", "transaction", "check",
        "invoice", "order", "reference", "ref", "tracking", "isbn", "ssn",
    )
    if any(token in compact for token in identifier_tokens):
        return text

    # Convert numeric columns only when the column heading explicitly indicates a number.
    numeric_tokens = (
        "amount", "price", "cost", "total", "quantity", "qty", "count",
        "revenue", "balance", "rate", "percent", "percentage", "discount",
        "tax", "subtotal", "stock", "received", "damaged", "weight", "score",
    )
    if any(token in compact for token in numeric_tokens):
        candidate = text.replace(",", "").replace("$", "").replace("₹", "").replace("€", "").replace("£", "").strip()
        is_percent = candidate.endswith("%") or "percent" in compact or "percentage" in compact
        if candidate.endswith("%"):
            candidate = candidate[:-1].strip()
        if re.fullmatch(r"[-+]?(?:\d+(?:\.\d*)?|\.\d+)", candidate):
            try:
                number = float(candidate)
                if number.is_integer() and not any(ch in candidate for ch in "."):
                    number = int(number)
                return number / 100 if is_percent else number
            except (ValueError, OverflowError):
                pass

    # Parse only unambiguous ISO dates. Leave locale-dependent dates as text.
    if "date" in compact and re.fullmatch(r"\d{4}-\d{2}-\d{2}", text):
        from datetime import date
        try:
            return date.fromisoformat(text)
        except ValueError:
            return text

    return text


def create_workbook(tables: list[dict[str, Any]], pages: list[dict[str, Any]], output_path: Path) -> int:
    wb = Workbook()
    wb.remove(wb.active)
    used_names: set[str] = set()
    navy = "17365D"
    thin_gray = Side(style="thin", color="D9E2F3")

    for index, table in enumerate(tables, start=1):
        rows = table["rows"]
        if not rows:
            continue
        sheet_name = suggested_sheet_name(rows, index, used_names)
        ws = wb.create_sheet(sheet_name)
        # Keep the extracted header row as text and conservatively type data cells.
        headers = rows[0]
        ws.append([clean_cell(value) for value in headers])
        for row in rows[1:]:
            ws.append([
                excel_value(value, headers[col_index] if col_index < len(headers) else "")
                for col_index, value in enumerate(row)
            ])

        if ws.max_row:
            ws.row_dimensions[1].height = 30
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill(fill_type="solid", fgColor=navy)
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = Border(bottom=thin_gray)

            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    cell.alignment = Alignment(vertical="top", wrap_text=True)
                    if isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool):
                        cell.number_format = "#,##0.##"
                    elif hasattr(cell.value, "year") and hasattr(cell.value, "month") and hasattr(cell.value, "day"):
                        cell.number_format = "yyyy-mm-dd"

            for col_cells in ws.columns:
                max_length = max((len(str(c.value or "")) for c in col_cells), default=0)
                letter = get_column_letter(col_cells[0].column)
                ws.column_dimensions[letter].width = min(max(max_length + 2, 10), 40)

            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

    if not tables:
        ws = wb.create_sheet("Text Extraction")
        ws.append(["Page", "Extracted Text"])
        for page in pages:
            text = page["text"] or "No selectable text detected; OCR may be required for scanned pages."
            ws.append([page["page_number"], text])
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(fill_type="solid", fgColor=navy)
        ws.column_dimensions["A"].width = 12
        ws.column_dimensions["B"].width = 100
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

    wb.save(output_path)
    return sum(1 for table in tables if table["rows"])


# -------------------- DocXL authentication --------------------
# Uses only Python's standard library; no extra auth package is required.
import base64
import hashlib
import hmac
import json
import os
import secrets
import sqlite3
import time
from datetime import datetime, timezone
from email.utils import parseaddr
from fastapi import Depends, Header
from pydantic import BaseModel, Field

AUTH_DB = BASE_DIR / "docxl_users.db"
AUTH_SECRET = os.getenv("DOCXL_AUTH_SECRET", "")
if not AUTH_SECRET:
    # Development fallback. Set DOCXL_AUTH_SECRET in production.
    AUTH_SECRET = "CHANGE-ME-DocXL-dev-secret-" + str((BASE_DIR / "app").resolve())
TOKEN_TTL_SECONDS = 60 * 60


def auth_db():
    conn = sqlite3.connect(AUTH_DB)
    conn.row_factory = sqlite3.Row
    return conn


def init_auth_db():
    with auth_db() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                email TEXT NOT NULL UNIQUE COLLATE NOCASE,
                password_hash TEXT,
                google_sub TEXT UNIQUE,
                created_at TEXT NOT NULL
            )
        """)


init_auth_db()


class RegisterRequest(BaseModel):
    name: str = Field(min_length=2, max_length=80)
    email: str = Field(min_length=5, max_length=254)
    password: str = Field(min_length=8, max_length=128)


class LoginRequest(BaseModel):
    email: str
    password: str




def normalize_email(value: str) -> str:
    email = value.strip().lower()
    if len(email) > 254 or parseaddr(email)[1] != email or "@" not in email:
        raise HTTPException(status_code=422, detail="Enter a valid email address.")
    local, domain = email.rsplit("@", 1)
    if not local or "." not in domain or domain.startswith(".") or domain.endswith("."):
        raise HTTPException(status_code=422, detail="Enter a valid email address.")
    return email


def hash_password(password: str) -> str:
    salt = secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, 310_000)
    return "pbkdf2_sha256$310000$" + base64.urlsafe_b64encode(salt).decode() + "$" + base64.urlsafe_b64encode(digest).decode()


def verify_password(password: str, stored: str | None) -> bool:
    if not stored:
        return False
    try:
        scheme, rounds, salt_b64, digest_b64 = stored.split("$", 3)
        if scheme != "pbkdf2_sha256":
            return False
        salt = base64.urlsafe_b64decode(salt_b64.encode())
        expected = base64.urlsafe_b64decode(digest_b64.encode())
        actual = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, int(rounds))
        return hmac.compare_digest(actual, expected)
    except (ValueError, TypeError):
        return False


def b64url(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).rstrip(b"=").decode("ascii")


def create_access_token(user_id: int, email: str) -> str:
    now = int(time.time())
    header = b64url(json.dumps({"alg": "HS256", "typ": "JWT"}, separators=(",", ":")).encode())
    payload = b64url(json.dumps({"sub": str(user_id), "email": email, "iat": now, "exp": now + TOKEN_TTL_SECONDS}, separators=(",", ":")).encode())
    signing_input = f"{header}.{payload}".encode()
    signature = hmac.new(AUTH_SECRET.encode(), signing_input, hashlib.sha256).digest()
    return f"{header}.{payload}.{b64url(signature)}"


def decode_access_token(token: str) -> dict:
    try:
        header, payload, signature = token.split(".")
        expected = b64url(hmac.new(AUTH_SECRET.encode(), f"{header}.{payload}".encode(), hashlib.sha256).digest())
        if not hmac.compare_digest(signature, expected):
            raise ValueError("signature")
        claims = json.loads(base64.urlsafe_b64decode(payload + "=" * (-len(payload) % 4)))
        if int(claims.get("exp", 0)) <= int(time.time()):
            raise ValueError("expired")
        return claims
    except Exception:
        raise HTTPException(status_code=401, detail="Invalid or expired access token. Please sign in again.")


def current_user(authorization: str | None = Header(default=None)) -> dict:
    if not authorization or not authorization.lower().startswith("bearer "):
        raise HTTPException(status_code=401, detail="Sign in to continue.")
    claims = decode_access_token(authorization.split(" ", 1)[1].strip())
    with auth_db() as conn:
        user = conn.execute("SELECT id, name, email, created_at FROM users WHERE id = ?", (claims["sub"],)).fetchone()
    if not user:
        raise HTTPException(status_code=401, detail="User account was not found.")
    return dict(user)


def public_user(row) -> dict:
    return {"id": row["id"], "name": row["name"], "email": row["email"], "created_at": row["created_at"]}


@app.post("/api/auth/register", tags=["Authentication"])
def register_user(payload: RegisterRequest):
    email = normalize_email(payload.email)
    name = payload.name.strip()
    if len(name) < 2:
        raise HTTPException(status_code=422, detail="Name must contain at least 2 characters.")
    with auth_db() as conn:
        exists = conn.execute("SELECT id FROM users WHERE email = ?", (email,)).fetchone()
        if exists:
            raise HTTPException(status_code=409, detail="An account with this email already exists. Please sign in.")
        now = datetime.now(timezone.utc).isoformat()
        password_hash = hash_password(payload.password)
        cur = conn.execute("INSERT INTO users(name, email, password_hash, created_at) VALUES (?, ?, ?, ?)", (name, email, password_hash, now))
        user_id = cur.lastrowid
        row = conn.execute("SELECT id, name, email, created_at FROM users WHERE id = ?", (user_id,)).fetchone()
    return {"success": True, "access_token": create_access_token(user_id, email), "token_type": "bearer", "expires_in": TOKEN_TTL_SECONDS, "user": public_user(row)}


@app.post("/api/auth/login", tags=["Authentication"])
def login_user(payload: LoginRequest):
    email = normalize_email(payload.email)
    with auth_db() as conn:
        row = conn.execute("SELECT id, name, email, password_hash, created_at FROM users WHERE email = ?", (email,)).fetchone()
    if not row or not verify_password(payload.password, row["password_hash"]):
        raise HTTPException(status_code=401, detail="Incorrect email or password.")
    return {"success": True, "access_token": create_access_token(row["id"], row["email"]), "token_type": "bearer", "expires_in": TOKEN_TTL_SECONDS, "user": public_user(row)}


@app.get("/api/auth/me", tags=["Authentication"])
def get_current_user(user: dict = Depends(current_user)):
    return {"success": True, "user": user}


@app.get("/")
def home():
    return {"project": "DocXL", "status": "online", "message": "Welcome to DocXL!"}

@app.get("/health")
def health_check():
    return {"status": "healthy"}

@app.post("/api/upload", tags=["PDF Upload"])
async def upload_pdf(file: UploadFile = File(...)):
    if not file.filename:
        raise HTTPException(status_code=400, detail="Please provide a filename.")
    if Path(file.filename).suffix.lower() != ".pdf":
        raise HTTPException(status_code=400, detail="Only PDF files are allowed.")

    contents = await file.read(MAX_FILE_SIZE + 1)
    if not contents:
        raise HTTPException(status_code=400, detail="The uploaded file is empty.")
    if len(contents) > MAX_FILE_SIZE:
        raise HTTPException(status_code=413, detail="File size exceeds the 25 MB limit.")

    try:
        with pymupdf.open(stream=contents, filetype="pdf") as doc:
            if doc.is_encrypted:
                raise HTTPException(status_code=400, detail="Password-protected PDFs are not supported.")
            page_count = doc.page_count
            if page_count < 1:
                raise HTTPException(status_code=400, detail="The PDF contains no pages.")
            if page_count > MAX_PAGES:
                raise HTTPException(status_code=400, detail=f"PDFs can contain at most {MAX_PAGES} pages.")
            title = (doc.metadata or {}).get("title") or ""
    except HTTPException:
        raise
    except Exception:
        raise HTTPException(status_code=400, detail="The uploaded file is not a valid, readable PDF.")

    document_id = str(uuid4())
    (UPLOAD_DIR / f"{document_id}.pdf").write_bytes(contents)
    return {
        "success": True,
        "message": "PDF uploaded successfully.",
        "document_id": document_id,
        "original_filename": Path(file.filename).name,
        "file_size_bytes": len(contents),
        "page_count": page_count,
        "document_title": title,
        "status": "uploaded",
    }

@app.post("/api/extract/{document_id}", tags=["PDF Extraction"])
def extract_pdf(document_id: str):
    pdf_path = UPLOAD_DIR / f"{document_id}.pdf"
    if not pdf_path.exists():
        raise HTTPException(status_code=404, detail="PDF not found. Upload the PDF first.")
    try:
        tables, pages, warnings = extract_tables_from_pdf(pdf_path)
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"PDF extraction failed: {exc}")

    return {
        "success": True,
        "document_id": document_id,
        "total_pages": len(pages),
        "ocr_pages": [page["page_number"] for page in pages if page.get("ocr_used")],
        "tables_found": len(tables),
        "pages": pages,
        "tables": [
            {
                "page_number": table["page_number"],
                "table_number": table["table_number"],
                "row_count": len(table["rows"]),
                "column_count": table["column_count"],
                "source": table.get("source", "pdfplumber"),
                "data": table["rows"],
            }
            for table in tables
        ],
        "warnings": warnings,
    }

@app.post("/api/export/{document_id}", tags=["Excel Export"])
def export_to_excel(document_id: str):
    pdf_path = UPLOAD_DIR / f"{document_id}.pdf"
    if not pdf_path.exists():
        raise HTTPException(status_code=404, detail="PDF not found. Upload the PDF first.")

    excel_path = GENERATED_DIR / f"{document_id}.xlsx"
    try:
        raw_tables, pages, warnings = extract_tables_from_pdf(pdf_path)
        merged_tables = merge_continuation_tables(raw_tables)
        table_count = create_workbook(merged_tables, pages, excel_path)
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Excel generation failed: {exc}")

    return {
        "success": True,
        "message": "Excel file generated successfully.",
        "document_id": document_id,
        "tables_detected": len(raw_tables),
        "tables_exported": table_count,
        "ocr_pages": [page["page_number"] for page in pages if page.get("ocr_used")],
        "filename": excel_path.name,
        "download_endpoint": f"/api/download/{document_id}",
        "warnings": warnings,
    }

@app.get("/api/download/{document_id}", tags=["Excel Export"])
def download_excel(document_id: str):
    excel_path = GENERATED_DIR / f"{document_id}.xlsx"
    if not excel_path.exists():
        raise HTTPException(status_code=404, detail="Excel file not found. Generate the export first.")
    return FileResponse(
        path=excel_path,
        filename=f"DocXL_{document_id}.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

