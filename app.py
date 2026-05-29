from __future__ import annotations

import html
import io
import json
import os
import re
import sys
from typing import Iterable

BASE_DIR = os.path.dirname(__file__)
sys.path.insert(0, os.path.join(BASE_DIR, ".venv_deps"))
sys.path.insert(0, os.path.join(os.path.dirname(BASE_DIR), ".venv_deps"))

import requests
from flask import Flask, Response, jsonify, request, send_from_directory
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from python_calamine import CalamineWorkbook
import xlrd


try:
    from dotenv import load_dotenv
except ImportError:
    load_dotenv = None

if load_dotenv:
    load_dotenv()

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB max request size

GEMINI_API_URL = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent"


EXPECTED_HEADERS = [
    "No.",
    "Seller Name",
    "Item Name",
    "Category",
    "Brand",
    "Type",
    "ASIN Number",
    "SKU Number",
    "HSN Number",
    "MRP Price",
    "Item Color",
    "Weight",
    "Weight Unit",
    "Length",
    "Length Unit",
    "Width",
    "Width Unit",
    "Height",
    "Height Unit",
    "Channel Price",
    "Purchase Margin(%)",
    "Tax",
    "Purchase Cost",
    "Purchase Tax",
    "Final Purchase Cost",
    "JIO CODE",
    "COMPANY PROFIT MARGIN %",
    "SALE TP WITH OUT TAX",
    "GST%",
    "SALE TP WITH TAX",
    "MRP ",
    "SALE DISCOUNT AMT",
    "ASP (GROSS)",
    "GST%",
    "GST amount",
    "NET SALE",
    "AJIO MARGIN 34%",
    "PURCHASE",
    "GST%2",
    "GST2 amount",
    "BANK SETTLEMENT",
]


def normalize_cell(value: object) -> object:
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def find_header_index(rows: list[list[object]]) -> int:
    expected = {header.lower() for header in EXPECTED_HEADERS}
    best_index = 0
    best_score = -1

    for index, row in enumerate(rows[:15]):
        lowered = {
            str(cell).lower()
            for cell in row
            if cell is not None and str(cell).strip()
        }
        score = len(expected.intersection(lowered))
        if score > best_score:
            best_score = score
            best_index = index

    return best_index


def parse_excel(file_bytes: bytes) -> tuple[list[str], list[list[str]]]:
    workbook = load_workbook(filename=io.BytesIO(file_bytes), data_only=True)
    sheet = workbook.active
    raw_rows = [[normalize_cell(cell) for cell in row] for row in sheet.iter_rows(values_only=True)]
    return finalize_rows(raw_rows)


def parse_xls(file_bytes: bytes) -> tuple[list[str], list[list[str]]]:
    workbook = xlrd.open_workbook(file_contents=file_bytes)
    sheet = workbook.sheet_by_index(0)
    raw_rows = []

    for row_index in range(sheet.nrows):
        row = [normalize_cell(sheet.cell_value(row_index, col_index)) for col_index in range(sheet.ncols)]
        raw_rows.append(row)

    return finalize_rows(raw_rows)


def parse_with_calamine(file_bytes: bytes) -> tuple[list[str], list[list[str]]]:
    workbook = CalamineWorkbook.from_filelike(io.BytesIO(file_bytes))
    sheet = workbook.get_sheet_by_index(0)
    raw_rows = []

    for row in sheet.to_python(skip_empty_area=False):
        raw_rows.append([normalize_cell(cell) for cell in row])

    return finalize_rows(raw_rows)


def finalize_rows(raw_rows: list[list[str]]) -> tuple[list[str], list[list[str]]]:
    while raw_rows and not any(raw_rows[0]):
        raw_rows.pop(0)

    if not raw_rows:
        return EXPECTED_HEADERS, []

    header_index = find_header_index(raw_rows)
    header = raw_rows[header_index]
    data_rows = [row for row in raw_rows[header_index + 1 :] if any(cell != "" for cell in row)]
    column_count = max(len(header), *(len(row) for row in data_rows), len(EXPECTED_HEADERS))
    header = header + [""] * (column_count - len(header))
    for i, expected in enumerate(EXPECTED_HEADERS):
        if i < len(header) and not header[i]:
            header[i] = expected
    data_rows = [row + [""] * (column_count - len(row)) for row in data_rows]
    return header, data_rows


def extract_uploaded_file(content_type: str, body: bytes) -> tuple[str, bytes]:
    if "boundary=" not in content_type:
        raise ValueError("Invalid upload request.")

    boundary = content_type.split("boundary=", 1)[1].encode("utf-8")
    delimiter = b"--" + boundary

    for part in body.split(delimiter):
        if b"filename=" not in part:
            continue

        try:
            header_block, file_block = part.split(b"\r\n\r\n", 1)
        except ValueError as exc:
            raise ValueError("Upload format samajh nahi aaya.") from exc

        headers = header_block.decode("utf-8", errors="ignore")
        filename = "uploaded.xlsx"
        for line in headers.split("\r\n"):
            if "filename=" in line:
                filename = line.split("filename=", 1)[1].strip().strip('"')
                break

        file_bytes = file_block.rsplit(b"\r\n", 1)[0]
        return filename, file_bytes

    raise ValueError("Koi file receive nahi hui.")


def get_col_letter(col_idx: int) -> str:
    result = ""
    while col_idx >= 0:
        result = chr(col_idx % 26 + 65) + result
        col_idx = col_idx // 26 - 1
    return result


def render_table(headers: Iterable[str], rows: Iterable[Iterable[str]]) -> str:
    headers_list = list(headers)
    display_headers = [f"{get_col_letter(i)} - {cell}" for i, cell in enumerate(headers_list)]
    header_html = "".join(f"<th>{html.escape(cell)}</th>" for cell in display_headers)
    
    rows_list = list(rows)
    rows_json = json.dumps(rows_list).replace("</", "<\\/")
    headers_json = json.dumps(headers_list).replace("</", "<\\/")
    colspan = max(1, len(headers_list))
    
    empty_msg = f"<tr><td colspan='{colspan}' class='empty'>Upload ke baad rows yahan dikhengi.</td></tr>" if not rows_list else ""

    return (
        "<div class='table-wrap'>"
        "<table>"
        f"<thead><tr>{header_html}</tr></thead>"
        f"<tbody id='table-body'>{empty_msg}</tbody>"
        "</table>"
        "</div>"
        f"<script>const globalTableData = {rows_json};</script>"
        f"<script>const globalHeaders = {headers_json};</script>"
    )


def sanitize_filename(filename: str) -> str:
    cleaned = "".join(char for char in filename if char not in '<>:"/\\|?*').strip().strip(".")
    return cleaned or "ajio_export"


def coerce_export_cell(value: object) -> object:
    if value is None:
        return ""
    if isinstance(value, str):
        text = ILLEGAL_CHARACTERS_RE.sub("", value).strip()
        if not text:
            return ""
        cleaned = text.replace(",", "")
        is_percent = cleaned.endswith("%")
        if is_percent:
            cleaned = cleaned[:-1].strip()
        if re.fullmatch(r"[+-]?\d+", cleaned):
            try:
                num = int(cleaned)
                return float(num) / 100.0 if is_percent else num
            except ValueError:
                return text
        if re.fullmatch(r"[+-]?(?:\d*\.\d+|\d+\.\d*)", cleaned):
            try:
                num = float(cleaned)
                return num / 100.0 if is_percent else num
            except ValueError:
                return text
        return text
    return value


def build_export_workbook(headers: list[str], rows: list[list[object]]) -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "AJIO Export"

    sheet.append(headers)
    for row_index, row in enumerate(rows, start=2):
        padded = row + [""] * (len(headers) - len(row))
        coerced_row = [coerce_export_cell(cell) for cell in padded[: len(headers)]]
        for col_idx, cell_value in enumerate(coerced_row, start=1):
            cell = sheet.cell(row=row_index, column=col_idx, value=cell_value)
            if (
                isinstance(cell_value, (int, float))
                and "%" in headers[col_idx - 1]
                and abs(cell_value) <= 1
            ):
                cell.number_format = "0%"

    sheet.freeze_panes = "A2"

    red_fill = PatternFill(fill_type="solid", fgColor="FDE2E2")
    yellow_fill = PatternFill(fill_type="solid", fgColor="FFF4BF")
    green_fill = PatternFill(fill_type="solid", fgColor="DDF6E4")
    border = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000"),
    )

    # Apply styles to header row only (including borders)
    for col_idx in range(1, len(headers) + 1):
        cell = sheet.cell(row=1, column=col_idx)
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)
        if col_idx <= 25:
            cell.fill = red_fill
        elif col_idx == 26:
            cell.fill = yellow_fill
        else:
            cell.fill = green_fill

    # NO borders for data rows (as requested)
    if sheet.max_row > 1:
        for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=1, max_col=len(headers)):
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=False)

    # Calculate column widths efficiently
    column_widths = {}
    
    # Get header widths
    for col_idx in range(1, len(headers) + 1):
        header_val = str(headers[col_idx - 1] or "")
        column_widths[col_idx] = max(len(header_val) + 3, 12)

    # Sample data rows for width - strategic sampling for large files
    if len(rows) > 100:
        sample_size = min(500, max(100, len(rows) // 10))
        sample_indices = [int(i * len(rows) / sample_size) for i in range(sample_size)]
    else:
        sample_indices = range(len(rows))
    
    for row_idx in sample_indices:
        if row_idx < len(rows):
            for col_idx in range(1, len(headers) + 1):
                if col_idx - 1 < len(rows[row_idx]):
                    cell_val = str(rows[row_idx][col_idx - 1] or "")
                    if col_idx == 3:
                        # Item Name column - special handling
                        width = min(max(len(cell_val) + 3, 32), 80)
                    else:
                        width = min(max(len(cell_val) + 3, 12), 28)
                    column_widths[col_idx] = max(column_widths.get(col_idx, 12), width)

    # Apply column widths
    for col_idx, width in column_widths.items():
        column_letter = sheet.cell(row=1, column=col_idx).column_letter
        sheet.column_dimensions[column_letter].width = width

    # Set row heights only for header
    sheet.row_dimensions[1].height = 20

    output = io.BytesIO()
    workbook.save(output)
    return output.getvalue()


def render_page(table_html: str = "", message: str = "", default_tab: str = "auto") -> bytes:
    message_html = f"<p class='message'>{html.escape(message)}</p>" if message else "<p class='message'></p>"
    if not table_html:
        table_html = render_table(EXPECTED_HEADERS, [])

    page = f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>AJIO Cost calculator</title>
  <link rel="icon" type="image/png" href="/logo.png">
  <style>
    :root {{
      --bg-gradient: linear-gradient(135deg, #f0f4ff 0%, #fff0f5 40%, #f0fff4 70%, #fdfcf0 100%);
      --panel: rgba(255, 255, 255, 0.45);
      --panel-strong: rgba(255, 255, 255, 0.65);
      --line: rgba(255, 255, 255, 0.7);
      --line-soft: rgba(110, 140, 255, 0.15);
      --text: #2d3748;
      --muted: #718096;
      --primary: #5a67d8;
      --primary-soft: rgba(90, 103, 216, 0.1);
      --shadow: 0 25px 60px rgba(112, 135, 168, 0.15);
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: "Calibri", "Segoe UI", sans-serif;
      font-size: 16px;
      background: var(--bg-gradient);
      background-size: 400% 400%;
      animation: iris-bg 20s ease infinite;
      color: var(--text);
      min-height: 100vh;
      position: relative;
      overflow-x: hidden;
    }}
    @keyframes iris-bg {{
      0% {{ background-position: 0% 50%; }}
      50% {{ background-position: 100% 50%; }}
      100% {{ background-position: 0% 50%; }}
    }}
    body::before {{
      content: "";
      position: fixed;
      inset: 0;
      pointer-events: none;
      background: radial-gradient(circle at 50% 50%, rgba(255,255,255,0.2) 0%, transparent 80%);
      z-index: -1;
    }}
    .shell {{
      width: 100%;
      min-height: 100vh;
      margin: 0;
      background: rgba(255, 255, 255, 0.15);
      backdrop-filter: blur(30px) saturate(180%);
      -webkit-backdrop-filter: blur(30px) saturate(180%);
      overflow: hidden;
      position: relative;
    }}
    .shell::before {{
      content: "";
      position: absolute;
      inset: 0;
      background: linear-gradient(135deg, rgba(255,255,255,0.3), transparent 60%);
      pointer-events: none;
    }}
    .hero {{
      padding: 34px 40px 20px;
      border-bottom: 1px solid var(--line-soft);
      position: relative;
      display: flex;
      align-items: center;
      gap: 16px;
    }}
    .brand-logo {{
      width: 68px;
      height: 68px;
      object-fit: contain;
      flex-shrink: 0;
      filter: drop-shadow(0 10px 24px rgba(35, 48, 68, 0.12));
    }}
    h1 {{
      margin: 0;
      font-size: clamp(32px, 4.5vw, 56px);
      letter-spacing: -0.05em;
      line-height: 1.1;
      font-weight: 400;
      background: linear-gradient(135deg, #1e293b, #4762b4);
      -webkit-background-clip: text;
      -webkit-text-fill-color: transparent;
      filter: drop-shadow(0 10px 20px rgba(71, 98, 180, 0.15));
    }}
    form {{
      display: flex;
      flex-wrap: wrap;
      gap: 14px;
      padding: 24px 40px 12px;
      align-items: center;
      position: relative;
    }}
    input[type="file"],
    .margin-input-group,
    .margin-input-group input,
    .page-btn,
    .tab-btn,
    .result-card {{
      backdrop-filter: blur(18px);
      -webkit-backdrop-filter: blur(18px);
    }}
    .upload-zone {{
      margin: 10px 40px 20px;
      padding: 8px 20px;
      background: linear-gradient(135deg, 
        rgba(255, 182, 193, 0.15), 
        rgba(173, 216, 230, 0.15), 
        rgba(221, 160, 221, 0.15), 
        rgba(240, 253, 244, 0.15)
      );
      background-size: 400% 400%;
      animation: rainbow-glow 12s ease infinite;
      border: 1px solid rgba(255, 255, 255, 0.7);
      border-radius: 24px;
      display: flex;
      flex-direction: row;
      align-items: center;
      justify-content: space-between;
      gap: 16px;
      transition: all 0.3s ease;
      backdrop-filter: blur(25px);
      -webkit-backdrop-filter: blur(25px);
      box-shadow: 0 10px 30px rgba(112, 135, 168, 0.08);
      position: relative;
      overflow: hidden;
    }}
    #bulk-upload-form {{
      display: flex;
      flex-direction: row;
      align-items: center;
      gap: 16px;
      flex: 1;
      margin: 0;
      padding: 0;
    }}
    .action-group {{
      display: flex;
      align-items: center;
      gap: 10px;
    }}
    .file-info {{
      margin: -10px 40px 15px;
      font-size: 13px;
      font-weight: 400;
      color: #4762b4;
      background: rgba(255, 255, 255, 0.5);
      padding: 5px 15px;
      border-radius: 12px;
      display: inline-block;
      backdrop-filter: blur(10px);
      border: 1px solid rgba(255, 255, 255, 0.8);
    }}
    .file-input-wrapper {{
      position: relative;
      flex: 1;
      max-width: none;
    }}
    #bulk-file-input {{
      position: absolute;
      width: 0.1px;
      height: 0.1px;
      opacity: 0;
      overflow: hidden;
      z-index: -1;
    }}
    .file-label {{
      display: flex;
      flex-direction: row;
      align-items: center;
      justify-content: flex-start;
      padding: 10px 18px;
      background: rgba(255, 255, 255, 0.6);
      border: 1px solid rgba(255, 255, 255, 0.8);
      border-radius: 14px;
      cursor: pointer;
      transition: all 0.3s ease;
      gap: 12px;
      box-shadow: 0 4px 15px rgba(112, 135, 168, 0.05);
      width: 100%;
    }}
    .file-label:hover {{
      background: white;
      box-shadow: 0 8px 25px rgba(110, 140, 255, 0.12);
    }}
    .file-label .icon {{
      font-size: 20px;
      margin-bottom: 0;
      filter: drop-shadow(0 4px 8px rgba(110, 140, 255, 0.2));
    }}
    .file-label .text {{
      font-weight: 400;
      color: #4762b4;
      font-size: 14px;
      white-space: nowrap;
    }}
    .file-label .file-name {{
      margin-top: 0;
      font-size: 13px;
      color: var(--muted);
      font-weight: 400;
      background: var(--primary-soft);
      padding: 4px 10px;
      border-radius: 8px;
      max-width: 300px;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }}
    .primary-btn {{
      background: linear-gradient(135deg, #6e8cff, #8ba3ff) !important;
      color: white !important;
      padding: 10px 24px !important;
      font-size: 16px !important;
      font-weight: 400 !important;
      border: none !important;
      border-radius: 14px !important;
      box-shadow: 0 8px 20px rgba(110, 140, 255, 0.3) !important;
      cursor: pointer;
      transition: all 0.3s ease !important;
      white-space: nowrap;
    }}
    .primary-btn:hover {{
      transform: translateY(-2px);
      box-shadow: 0 12px 30px rgba(110, 140, 255, 0.4) !important;
    }}
    .btn-export {{
      background: linear-gradient(135deg, #4ade80, #22c55e) !important;
      box-shadow: 0 8px 20px rgba(34, 197, 94, 0.3) !important;
    }}
    .btn-export:hover {{
      box-shadow: 0 12px 30px rgba(34, 197, 94, 0.4) !important;
    }}
    .btn-clear {{
      background: linear-gradient(135deg, #fb7185, #e11d48) !important;
      box-shadow: 0 8px 20px rgba(225, 29, 72, 0.3) !important;
    }}
    .btn-clear:hover {{
      box-shadow: 0 12px 30px rgba(225, 29, 72, 0.4) !important;
    }}
    .btn-back {{
      background: linear-gradient(135deg, #94a3b8, #64748b) !important;
      box-shadow: 0 8px 20px rgba(100, 116, 139, 0.2) !important;
    }}
    .btn-back:hover {{
      box-shadow: 0 12px 30px rgba(100, 116, 139, 0.3) !important;
    }}
    .primary-btn:active {{
      transform: translateY(-1px);
    }}
    button {{
      border: 1px solid rgba(255,255,255,0.82);
      border-radius: 14px;
      background:
        linear-gradient(135deg, rgba(255,255,255,0.48), rgba(255,255,255,0.18)),
        linear-gradient(135deg, rgba(110,140,255,0.24), rgba(255,186,218,0.16));
      color: var(--text);
      padding: 11px 16px;
      font-size: 13px;
      font-weight: 400;
      cursor: pointer;
      box-shadow: inset 0 1px 0 rgba(255,255,255,0.78), 0 16px 36px rgba(110, 140, 255, 0.16);
      transition: transform 180ms ease, box-shadow 180ms ease, border-color 180ms ease;
    }}
    button:hover {{
      transform: translateY(-2px);
      background:
        linear-gradient(135deg, rgba(255,255,255,0.54), rgba(255,255,255,0.24)),
        linear-gradient(135deg, rgba(110,140,255,0.3), rgba(255,186,218,0.2));
      box-shadow: inset 0 1px 0 rgba(255,255,255,0.84), 0 20px 44px rgba(110, 140, 255, 0.22);
      border-color: rgba(110, 140, 255, 0.42);
    }}
    .controls-bar {{
      display: flex;
      flex-wrap: wrap;
      justify-content: space-between;
      align-items: center;
      padding: 0 40px 20px;
      gap: 16px;
    }}
    .message {{
      margin: 0;
      color: #425f81;
      font-weight: 400;
      flex: 1;
      padding: 14px 16px;
      min-height: 20px;
      background: rgba(255,255,255,0.38);
      border: 1px solid rgba(255,255,255,0.82);
      border-radius: 18px;
      box-shadow: 0 14px 34px rgba(112, 135, 168, 0.1);
      backdrop-filter: blur(16px);
      -webkit-backdrop-filter: blur(16px);
    }}
    .margin-input-group {{
      display: flex;
      align-items: center;
      gap: 10px;
      background: rgba(255, 255, 255, 0.45);
      padding: 10px 16px;
      border-radius: 18px;
      border: 1px solid rgba(255, 255, 255, 0.8);
      box-shadow: 0 15px 35px rgba(112, 135, 168, 0.08);
      transition: all 0.3s ease;
    }}
    .sample-link {{
      font-size: 13px;
      color: #6e8cff;
      text-decoration: none;
      border-bottom: 1px dashed #6e8cff;
      margin-left: 5px;
      transition: all 0.3s ease;
      display: inline-block;
    }}
    .sample-link:hover {{
      color: #5a67d8;
      border-bottom-style: solid;
      transform: translateY(-1px);
    }}

    .margin-input-group:hover {{
      transform: translateY(-2px);
      background: rgba(255, 255, 255, 0.65);
      box-shadow: 0 20px 45px rgba(112, 135, 168, 0.12);
    }}
    .margin-input-group label {{
      font-size: 14px;
      font-weight: 400;
      color: #4762b4;
      letter-spacing: 0.02em;
    }}
    .margin-input-group input {{
      width: 80px;
      padding: 8px 12px;
      border: 2px solid #e2e8f0;
      border-radius: 10px;
      font-size: 14px;
      font-weight: 400;
      background: white;
      color: #1e293b;
      transition: all 0.3s ease;
    }}
    .margin-input-group input:focus {{
      border-color: #6e8cff;
      outline: none;
      box-shadow: 0 0 0 4px rgba(110, 140, 255, 0.1);
    }}
      text-align: right;
      background: linear-gradient(135deg, rgba(110,140,255,0.14), rgba(255,255,255,0.72));
      color: #4762b4;
      font-weight: 400;
      box-shadow: inset 0 1px 0 rgba(255,255,255,0.78), 0 0 0 1px rgba(110,140,255,0.08);
      transition: border-color 180ms ease, box-shadow 180ms ease, background 180ms ease;
    }}
    .margin-input-group input:hover,
    .margin-input-group input:focus {{
      border-color: rgba(110, 140, 255, 0.48);
      background: linear-gradient(135deg, rgba(110,140,255,0.18), rgba(255,255,255,0.85));
      box-shadow: inset 0 1px 0 rgba(255,255,255,0.88), 0 0 0 3px rgba(110,140,255,0.12);
      outline: none;
    }}
    .margin-input-group input::placeholder {{
      color: #6d7fcc;
      opacity: 1;
    }}
    .page-btn {{
      border-radius: 14px;
      background: rgba(255,255,255,0.4);
      color: var(--text);
      font-weight: 400;
      padding: 8px 14px;
      font-size: 13px;
    }}
    .page-btn:disabled {{
      opacity: 0.45;
      cursor: not-allowed;
    }}
    .export-btn {{
      white-space: nowrap;
    }}
    .toggle-container {{
      display: flex;
      align-items: center;
      gap: 10px;
      padding: 8px 16px;
      background: rgba(255, 255, 255, 0.45);
      border-radius: 20px;
      border: 1px solid rgba(255, 255, 255, 0.8);
      cursor: pointer;
      user-select: none;
      transition: all 0.3s ease;
      box-shadow: 0 4px 15px rgba(112, 135, 168, 0.05);
    }}
    .toggle-container:hover {{
      background: rgba(255, 255, 255, 0.65);
      transform: translateY(-1px);
    }}
    .toggle-label {{
      font-size: 13px;
      font-weight: 400;
      color: #4762b4;
    }}
    .switch {{
      position: relative;
      display: inline-block;
      width: 44px;
      height: 22px;
    }}
    .switch input {{
      opacity: 0;
      width: 0;
      height: 0;
    }}
    .slider {{
      position: absolute;
      cursor: pointer;
      inset: 0;
      background-color: #cbd5e1;
      transition: .4s;
      border-radius: 34px;
    }}
    .slider:before {{
      position: absolute;
      content: "";
      height: 16px;
      width: 16px;
      left: 3px;
      bottom: 3px;
      background-color: white;
      transition: .4s;
      border-radius: 50%;
      box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }}
    input:checked + .slider {{
      background: linear-gradient(135deg, #6e8cff, #8ba3ff);
    }}
    input:checked + .slider:before {{
      transform: translateX(22px);
    }}
    .table-wrap {{
      overflow: auto;
      margin: 0;
      width: 100%;
      background: transparent;
      border: none;
      border-radius: 0;
      box-shadow: none;
      backdrop-filter: none;
      -webkit-backdrop-filter: none;
    }}
    table {{
      width: 100%;
      min-width: 1200px;
      border-collapse: collapse;
      background: transparent;
      border-radius: 18px;
      overflow: hidden;
    }}
    th, td {{
      padding: 12px 14px;
      border-bottom: 1px solid rgba(124, 145, 180, 0.12);
      border-right: 1px solid rgba(124, 145, 180, 0.08);
      text-align: left;
      vertical-align: top;
      white-space: nowrap;
      font-size: 15px;
    }}
    th {{
      position: sticky;
      top: 0;
      background: rgba(240, 247, 255, 0.8) !important;
      z-index: 1;
      color: #4762b4;
      font-weight: 400;
      font-size: 13px;
      letter-spacing: 0.05em;
      text-transform: uppercase;
      border-bottom: 2px solid rgba(110, 140, 255, 0.15);
    }}
    tr:nth-child(even) td {{
      background: rgba(255, 255, 255, 0.1);
    }}
    tbody tr:hover td {{
      background: rgba(255, 255, 255, 0.3);
    }}
    .empty {{
      color: var(--muted);
      text-align: center;
      padding: 28px;
    }}
    .tabs-container {{
      display: flex;
      gap: 12px;
      padding: 0 40px 20px;
      margin-bottom: 20px;
      border-bottom: 1px solid var(--line-soft);
    }}
    .tab-btn {{
      padding: 10px 16px;
      color: var(--muted);
      font-weight: 400;
      font-size: 15px;
      cursor: pointer;
      border-radius: 14px;
      box-shadow: none;
    }}
    .tab-btn.active {{
      color: white;
      background: linear-gradient(135deg, #6e8cff, #8ba3ff);
      box-shadow: 0 10px 25px rgba(110, 140, 255, 0.3);
      border: none;
    }}
    .global-inputs {{
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      padding: 0 40px 20px;
    }}
    .single-results-grid {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(210px, 1fr));
      gap: 12px;
      padding: 12px 40px 40px;
    }}
    .result-card {{
      background: rgba(255, 255, 255, 0.4);
      border: 1px solid rgba(255, 255, 255, 0.8);
      border-radius: 20px;
      padding: 14px 20px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 15px;
      box-shadow: 0 10px 30px rgba(112, 135, 168, 0.05);
      transition: all 0.3s ease;
    }}
    .result-card:hover {{
      transform: translateY(-3px);
      background: rgba(255, 255, 255, 0.6);
      box-shadow: 0 15px 40px rgba(112, 135, 168, 0.1);
    }}
    .result-card h4 {{
      margin: 0;
      font-size: 12px;
      color: var(--muted);
      text-transform: uppercase;
      letter-spacing: 0.09em;
      line-height: 1.4;
      text-align: left;
      max-width: 70%;
      font-weight: 400;
    }}
    .result-card p {{
      margin: 0;
      font-size: 18px;
      font-weight: 400;
      color: var(--text);
      flex-shrink: 0;
      text-align: right;
    }}
    .highlight-card {{
      background: linear-gradient(135deg, rgba(255,255,255,0.42), rgba(110,140,255,0.08));
      border: 1px solid rgba(110, 140, 255, 0.24);
    }}
    .popup-backdrop {{
      position: fixed;
      inset: 0;
      background: rgba(245, 249, 255, 0.52);
      backdrop-filter: blur(10px);
      -webkit-backdrop-filter: blur(10px);
      display: none;
      align-items: center;
      justify-content: center;
      padding: 16px;
      z-index: 50;
    }}
    .popup-backdrop.open {{
      display: flex;
    }}
    .popup-card {{
      width: min(480px, 95%);
      background: linear-gradient(135deg, #f1f6ff 0%, #ffffff 100%);
      border: 1px solid rgba(110, 140, 255, 0.25);
      border-radius: 32px;
      box-shadow: 0 30px 70px rgba(15, 23, 42, 0.2);
      padding: 40px;
      position: relative;
      overflow: hidden;
    }}
    .popup-card::before {{
      content: '';
      position: absolute;
      top: 0;
      left: 0;
      right: 0;
      height: 6px;
      background: linear-gradient(90deg, #6e8cff, #8ba3ff, #4ade80, #fb7185);
    }}
    .popup-card h3 {{
      font-size: 26px !important;
      font-weight: 400 !important;
      margin-bottom: 14px !important;
      color: #1e293b !important;
      letter-spacing: -0.03em !important;
      margin-top: 0;
    }}
    .popup-card p {{
      color: #64748b !important;
      font-size: 16px !important;
      line-height: 1.6 !important;
      margin-bottom: 28px !important;
    }}
    #export-file-name {{
      width: 100%;
      padding: 16px 22px !important;
      border-radius: 18px !important;
      border: 2px solid #cbd5e1 !important;
      background: #ffffff !important;
      font-size: 16px !important;
      font-weight: 400 !important;
      color: #1e293b !important;
      margin-bottom: 30px !important;
      transition: all 0.3s ease !important;
      box-shadow: inset 0 2px 4px rgba(0,0,0,0.02) !important;
      box-sizing: border-box;
    }}
    #export-file-name:focus {{
      border-color: #6e8cff !important;
      background: white !important;
      outline: none !important;
      box-shadow: 0 0 0 4px rgba(110, 140, 255, 0.15) !important;
    }}
    .popup-actions {{
      display: flex;
      justify-content: flex-end;
      gap: 12px;
      margin-top: 18px;
    }}

    /* Chatbot Styles */
    .bot-trigger {{
      position: fixed;
      bottom: 30px;
      right: 30px;
      width: 70px;
      height: 70px;
      z-index: 2000;
      cursor: pointer;
      animation: bot-float 3s ease-in-out infinite;
      filter: drop-shadow(0 10px 20px rgba(71, 98, 180, 0.25));
    }}
    @keyframes bot-float {{
      0%, 100% {{ transform: translateY(0) rotate(0deg); }}
      50% {{ transform: translateY(-10px) rotate(5deg); }}
    }}
    .bot-trigger img {{
      width: 100%;
      height: 100%;
      object-fit: contain;
    }}
    .bot-bubble {{
      position: fixed;
      bottom: 110px;
      right: 35px;
      background: rgba(255, 255, 255, 0.95);
      border: 2px solid rgba(110, 140, 255, 0.3);
      border-radius: 18px 18px 0 18px;
      padding: 12px 18px;
      max-width: 200px;
      font-size: 13px;
      font-weight: 400;
      color: #4762b4;
      box-shadow: 0 10px 25px rgba(0,0,0,0.1);
      z-index: 2000;
      opacity: 0;
      transform: translateY(20px);
      transition: all 0.5s cubic-bezier(0.175, 0.885, 0.32, 1.275);
      pointer-events: none;
      backdrop-filter: blur(10px);
    }}
    .bot-bubble.show {{
      opacity: 1;
      transform: translateY(0);
    }}
    .bot-window {{
      position: fixed;
      bottom: 110px;
      right: 30px;
      width: 380px;
      height: 550px;
      background: linear-gradient(135deg, rgba(255, 255, 255, 0.96), rgba(240, 253, 244, 0.96));
      border: 1px solid rgba(255, 255, 255, 0.8);
      border-radius: 30px;
      box-shadow: 0 25px 60px rgba(15, 23, 42, 0.25);
      z-index: 1999;
      display: flex;
      flex-direction: column;
      overflow: hidden;
      transform: scale(0.9) translateY(40px);
      opacity: 0;
      visibility: hidden;
      transition: all 0.4s cubic-bezier(0.175, 0.885, 0.32, 1.275);
      backdrop-filter: blur(25px);
    }}
    .bot-window.open {{
      opacity: 1;
      visibility: visible;
      transform: scale(1) translateY(0);
    }}
    .bot-header {{
      padding: 20px;
      background: linear-gradient(90deg, #4ade80, #6e8cff, #fb7185);
      color: white;
      display: flex;
      align-items: center;
      justify-content: space-between;
    }}
    .bot-header h3 {{
      margin: 0;
      font-size: 16px;
      font-weight: 400;
      letter-spacing: -0.02em;
    }}
    .bot-messages {{
      flex: 1;
      padding: 20px;
      overflow-y: auto;
      display: flex;
      flex-direction: column;
      gap: 12px;
      background: rgba(255, 255, 255, 0.2);
    }}
    .msg {{
      max-width: 80%;
      padding: 10px 14px;
      border-radius: 16px;
      font-size: 16px;
      line-height: 1.5;
    }}
    .msg.bot {{
      align-self: flex-start;
      background: white;
      color: #1e293b;
      border-bottom-left-radius: 4px;
      box-shadow: 0 4px 10px rgba(112, 135, 168, 0.08);
      border: 1px solid rgba(110, 140, 255, 0.1);
    }}
    .msg.user {{
      align-self: flex-end;
      background: linear-gradient(135deg, #6e8cff, #8ba3ff);
      color: white;
      border-bottom-right-radius: 4px;
      box-shadow: 0 4px 12px rgba(110, 140, 255, 0.2);
    }}
    .bot-input-area {{
      padding: 16px;
      display: flex;
      gap: 8px;
      background: white;
      border-top: 1px solid rgba(124, 145, 180, 0.1);
    }}
    .bot-input-area input {{
      flex: 1;
      padding: 10px 16px;
      border: 2px solid #f1f5f9;
      border-radius: 12px;
      font-size: 14px;
      outline: none;
      transition: all 0.3s ease;
    }}
    .bot-input-area input:focus {{
      border-color: #6e8cff;
    }}
    .send-btn {{
      background: linear-gradient(135deg, #4ade80, #22c55e);
      color: white;
      border: none;
      width: 40px;
      height: 40px;
      border-radius: 10px;
      cursor: pointer;
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 18px;
    }}
    @media (max-width: 720px) {{
      .hero, form, .global-inputs, .controls-bar, .tabs-container, .single-results-grid {{
        padding-left: 20px;
        padding-right: 20px;
      }}
      .table-wrap {{
        margin: 0 12px 24px;
      }}
    }}
  </style>
  <script src="https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js"></script>
</head>
<body>
  <main class="shell">
    <section class="hero">
      <img src="/logo.png" class="brand-logo" alt="AJIO Logo">
      <h1>AJIO Cost calculator</h1>
    </section>
    <div class="global-inputs">
      <div class="margin-input-group">
        <label for="company_margin">COMPANY PROFIT MARGIN %</label>
        <input type="number" id="company_margin" name="company_margin" placeholder="0" step="0.01">
      </div>
      <div class="margin-input-group">
        <label for="gst_margin">SALE GST %</label>
        <input type="number" id="gst_margin" name="gst_margin" placeholder="0" step="0.01">
      </div>
      <div class="margin-input-group">
        <label for="discount_margin">SALE DISCOUNT AMT%</label>
        <input type="number" id="discount_margin" name="discount_margin" placeholder="0" step="0.01">
      </div>
      <div class="margin-input-group">
        <label for="ajio_margin">AJIO MARGIN %</label>
        <input type="number" id="ajio_margin" name="ajio_margin" placeholder="0" step="0.01">
      </div>
    </div>

    <div class="tabs-container">
      <button class="tab-btn" id="tab-single">Single Item Calculator</button>
      <button class="tab-btn" id="tab-bulk">Bulk Excel Processor</button>
    </div>

    <div id="single-container" style="display: none;">
      <div style="display: flex; flex-wrap: wrap; gap: 10px; padding: 0 34px 4px;">
        <div class="margin-input-group" style="background: var(--surface);">
          <label for="single_y_val">FINAL PURCHASE COST (Y)</label>
          <input type="number" id="single_y_val" placeholder="0" step="0.01">
        </div>
        <div class="margin-input-group" style="background: var(--surface);">
          <label for="single_purchase_gst">PURCHASE GST %</label>
          <input type="number" id="single_purchase_gst" placeholder="0" step="0.01">
        </div>
      </div>
      
      <div class="single-results-grid">
          <div class="result-card"><h4>Purchase GST Amt</h4><p id="res_purchase_gst">0.00</p></div>
          <div class="result-card"><h4>Purchase TP W/O Tax</h4><p id="res_purchase_tp">0.00</p></div>
          <div class="result-card"><h4>COMPANY PROFIT MARGIN % (AA)</h4><p id="res_AA">0.00</p></div>
          <div class="result-card"><h4>SALE TP WITH TAX (AD)</h4><p id="res_AD">0.00</p></div>
          <div class="result-card highlight-card"><h4>MRP (AE)</h4><p id="res_AE">0.00</p></div>
          <div class="result-card"><h4>SALE DISCOUNT AMT (AF)</h4><p id="res_AF">0.00</p></div>
          <div class="result-card"><h4>ASP (GROSS) (AG)</h4><p id="res_AG">0.00</p></div>
          <div class="result-card"><h4>GST% (AH)</h4><p id="res_AH_pct">0%</p></div>
          <div class="result-card"><h4>GST amount (AI)</h4><p id="res_AI">0.00</p></div>
          <div class="result-card"><h4>NET SALE (AJ)</h4><p id="res_AJ">0.00</p></div>
          <div class="result-card"><h4>AJIO MARGIN 34% (AK)</h4><p id="res_AK">0.00</p></div>
          <div class="result-card"><h4>PURCHASE (AL)</h4><p id="res_AL">0.00</p></div>
          <div class="result-card"><h4>GST%2 (AM)</h4><p id="res_AM_pct">0%</p></div>
          <div class="result-card"><h4>GST2 amount (AN)</h4><p id="res_AN">0.00</p></div>
          <div class="result-card highlight-card"><h4>BANK SETTLEMENT (AO)</h4><p id="res_AO">0.00</p></div>
      </div>
    </div>

    <div id="bulk-container" style="display: none;">
      <div class="upload-zone">
        <form method="post" enctype="multipart/form-data" id="bulk-upload-form">
          <div class="file-input-wrapper" style="max-width: 250px;">
            <input type="file" id="bulk-file-input" name="file" accept=".xlsx, .xls" required>
            <label for="bulk-file-input" class="file-label" style="padding: 8px 15px;">
              <span class="icon">📁</span>
              <span class="text" id="file-name-display">Choose File</span>
            </label>
          </div>
          
          <label class="toggle-container">
            <span class="toggle-label">Purchase GST</span>
            <div class="switch">
              <input type="checkbox" name="purchase_gst_toggle">
              <span class="slider"></span>
            </div>
          </label>
          
          <button type="submit" class="primary-btn">Upload</button>
          <a href="/sample-excel" class="sample-link">Download Sample Excel</a>
        </form>

        <div class="action-group">
          <button onclick="exportToExcel()" class="primary-btn btn-export">Export</button>
          <button onclick="openClearPopup()" class="primary-btn btn-clear">Clear Data</button>
          <button onclick="history.back()" class="primary-btn btn-back">Back</button>
        </div>
      </div>
      
      {f'<div style="padding: 0 40px;"><p class="file-info">{html.escape(message)}</p></div>' if message else ''}
      
      <div style="padding: 0;">
        <div class="table-wrap">
          {table_html}
        </div>
        
        <div style="padding: 20px 0; display: flex; justify-content: center; margin-bottom: 40px;">
          <div id="pagination-controls" style="display: none; align-items: center; gap: 20px; background: rgba(255,255,255,0.5); padding: 10px 25px; border-radius: 20px; border: 1px solid rgba(255,255,255,0.8); backdrop-filter: blur(10px);">
              <div style="display:flex; align-items:center; gap:8px; margin-right:8px;">
                <label for="page-size-select" style="font-weight:700; color:#4762b4;">Show</label>
                <select id="page-size-select" style="padding:6px 8px; border-radius:8px; border:1px solid rgba(0,0,0,0.06); background:white; font-weight:700; color:#4762b4;">
                  <option value="10">10</option>
                  <option value="25">25</option>
                  <option value="50" selected>50</option>
                  <option value="100">100</option>
                  <option value="all">All</option>
                </select>
                <span style="font-weight:700; color:#4762b4;">entries</span>
              </div>
              <button id="prev-btn" class="page-btn" type="button" style="padding: 8px 20px; border-radius: 12px; font-weight: 700;">Prev</button>
              <span id="page-info" style="font-size: 15px; font-weight: 800; color: #4762b4;"></span>
              <button id="next-btn" class="page-btn" type="button" style="padding: 8px 20px; border-radius: 12px; font-weight: 700;">Next</button>
            </div>
        </div>
      </div>
    </div>
  </main>
  <div class="bot-bubble" id="bot-bubble">Hii I'am AJ your friend how can i help you</div>
  <div class="bot-trigger" onclick="toggleChat()">
    <img src="/chatbot.png" alt="AJ">
  </div>
  <div class="bot-window" id="bot-window">
    <div class="bot-header">
      <h3>AJ - Dashboard Friend</h3>
      <span style="cursor:pointer; font-size: 20px;" onclick="toggleChat()">&times;</span>
    </div>
    <div class="bot-messages" id="bot-messages">
      <div class="msg bot">Hello! I'm AJ, your dashboard friend. ✨ Before we start, may I know your name? 😊</div>
    </div>
    <div class="bot-input-area">
      <input type="text" id="bot-input" placeholder="Type your message..." onkeydown="if(event.key==='Enter') sendMessage()">
      <button class="send-btn" onclick="sendMessage()">➤</button>
    </div>
  </div>

  <div class="popup-backdrop" id="export-popup">
    <div class="popup-card">
      <h3>Export Excel</h3>
      <p>Enter the name for the exported file. File will be downloaded in .xlsx format.</p>
      <input type="text" id="export-file-name" placeholder="ajio_bulk_export">
      <div class="popup-actions">
        <button type="button" id="cancel-export" class="primary-btn btn-back">Cancel</button>
        <button type="button" id="confirm-export" class="primary-btn btn-export">Download File</button>
      </div>
    </div>
  </div>
  <div class="popup-backdrop" id="clear-popup">
    <div class="popup-card">
      <h3>Clear Data</h3>
      <p>Do you want to delete the existing data and upload new data?</p>
      <div class="popup-actions">
        <button type="button" id="cancel-clear" class="primary-btn btn-back">Cancel</button>
        <button type="button" id="confirm-clear" class="primary-btn btn-clear">Yes, Clear Data</button>
      </div>
    </div>
  </div>
  <script>
    document.addEventListener('DOMContentLoaded', () => {{
      const marginInput = document.getElementById('company_margin');
      const gstInput = document.getElementById('gst_margin');
      const discountInput = document.getElementById('discount_margin');
      const ajioInput = document.getElementById('ajio_margin');
      const tableBody = document.getElementById('table-body');
      
      const fileInput = document.getElementById('bulk-file-input');
      const fileNameDisplay = document.getElementById('file-name-display');
      const uploadForm = document.getElementById('bulk-upload-form');

      if (fileInput && fileNameDisplay) {{
        fileInput.addEventListener('change', (e) => {{
          const fileName = e.target.files[0]?.name || 'Choose File';
          fileNameDisplay.textContent = fileName;
        }});
      }}

      if (uploadForm) {{
        uploadForm.addEventListener('submit', () => {{
          const btn = uploadForm.querySelector('button[type="submit"]');
          if (btn) {{
            btn.disabled = true;
            btn.textContent = 'Uploading...';
          }}
        }});
      }}
      const prevBtn = document.getElementById('prev-btn');
      const nextBtn = document.getElementById('next-btn');
      const pageInfo = document.getElementById('page-info');
      const paginationControls = document.getElementById('pagination-controls');
      const exportPopup = document.getElementById('export-popup');
      const openExportPopupBtn = document.getElementById('open-export-popup');
      const clearDataBtn = document.getElementById('clear-data-btn');
      const cancelExportBtn = document.getElementById('cancel-export');
      const confirmExportBtn = document.getElementById('confirm-export');
      const exportFileNameInput = document.getElementById('export-file-name');
      
      const clearPopup = document.getElementById('clear-popup');
      const cancelClearBtn = document.getElementById('cancel-clear');
      const confirmClearBtn = document.getElementById('confirm-clear');
      
      let currentPage = 1;
      let pageSize = 50;
      const tableData = typeof globalTableData !== 'undefined' ? globalTableData : [];
      const exportHeaders = typeof globalHeaders !== 'undefined' ? globalHeaders : [];
      const hasUploadedData = tableData.length > 0;

      // Initialize pageSize from localStorage (supports 'all')
      const savedPageSize = localStorage.getItem('ajio_page_size');
      if (savedPageSize === 'all') {{
        pageSize = tableData.length || 50;
      }} else if (savedPageSize) {{
        const parsed = parseInt(savedPageSize, 10);
        if (!isNaN(parsed) && parsed > 0) pageSize = parsed;
      }}

      if (tableData.length > 0) {{
        paginationControls.style.display = 'flex';
        const pageSizeSelect = document.getElementById('page-size-select');
        if (pageSizeSelect) {{
          // Set select to match saved value (or default)
          if (savedPageSize === 'all') {{
            pageSizeSelect.value = 'all';
          }} else {{
            pageSizeSelect.value = String(pageSize);
          }}

          pageSizeSelect.addEventListener('change', (e) => {{
            const v = e.target.value;
            if (v === 'all') {{
              pageSize = tableData.length || 1;
            }} else {{
              pageSize = parseInt(v, 10) || 50;
            }}
            localStorage.setItem('ajio_page_size', v);
            currentPage = 1;
            renderPage();
          }});
        }}
      }}

      function escapeHtml(text) {{
        const map = {{
          '&': '&amp;',
          '<': '&lt;',
          '>': '&gt;',
          '"': '&quot;',
          "'": '&#039;'
        }};
        return String(text).replace(/[&<>"']/g, function(m) {{ return map[m]; }});
      }}

      function sanitizeFilename(name) {{
        const cleaned = String(name || '')
          .replace(/[<>:"/\\\\|?*]+/g, '')
          .trim()
          .replace(/\\.+$/, '');
        return cleaned || 'ajio_bulk_export';
      }}

      function calculateAE_raw(targetAD) {{
        if (targetAD <= 0) return 0;
        if (targetAD <= 1606) return targetAD / 0.22505;
        if (targetAD <= 2620) return targetAD / 0.186490678;
        return targetAD / 0.209579661;
      }}

      const tabSingle = document.getElementById('tab-single');
      const tabBulk = document.getElementById('tab-bulk');
      const containerSingle = document.getElementById('single-container');
      const containerBulk = document.getElementById('bulk-container');
      function activateTab(mode) {{
        if (!tabSingle || !tabBulk) return;
        const showBulk = mode === 'bulk';
        tabBulk.classList.toggle('active', showBulk);
        tabSingle.classList.toggle('active', !showBulk);
        containerBulk.style.display = showBulk ? 'block' : 'none';
        containerSingle.style.display = showBulk ? 'none' : 'block';
        localStorage.setItem('ajio_active_tab', mode);
      }}

      if (tabSingle && tabBulk) {{
        tabSingle.addEventListener('click', () => {{
          activateTab('single');
        }});

        tabBulk.addEventListener('click', () => {{
          activateTab('bulk');
        }});
      }}

      const savedTab = localStorage.getItem('ajio_active_tab');
      const serverDefaultTab = "{default_tab}";
      let initialTab = 'single';
      
      if (serverDefaultTab === 'bulk') {{
        initialTab = 'bulk';
      }} else if (hasUploadedData) {{
        initialTab = 'bulk';
      }} else if (savedTab) {{
        initialTab = savedTab;
      }}
      
      activateTab(initialTab);

      const singleYInput = document.getElementById('single_y_val');
      const singlePurchGstInput = document.getElementById('single_purchase_gst');

      function updateSingleMode() {{
        if (!singleYInput) return;
        
        const yValue = parseFloat(singleYInput.value.trim()) || 0;
        const purchGst = parseFloat(singlePurchGstInput.value.trim()) || 0;
        const margin = parseFloat(marginInput.value.trim()) || 0;
        const discount = parseFloat(discountInput.value.trim()) || 0;
        const ajio = parseFloat(ajioInput.value.trim()) || 0;
        
        // Purchase Side
        const purchGstAmt = Math.round(((yValue * purchGst) / (100 + purchGst)) * 100) / 100;
        const purchTp = Math.round((yValue - purchGstAmt) * 100) / 100;
        
        // Sales Side (Target calculations)
        const valAA = Math.round(((yValue * margin) / 100) * 100) / 100;
        const valAD = Math.round((yValue + valAA) * 100) / 100;
        
        // New direct formula for MRP (Raw decimal for precision)
        const valAE_raw = calculateAE_raw(valAD);
        
        // Final calculations based on RAW MRP for precision
        const valAF_raw = (valAE_raw * discount) / 100;
        const valAG_raw = valAE_raw - valAF_raw;
        const valAH_num = valAG_raw > 2499 ? 0.18 : 0.05;
        const valAH_pct = (valAH_num * 100) + '%';
        const valAI_raw = valAG_raw - (valAG_raw / (1 + valAH_num));
        const valAJ_raw = valAG_raw - valAI_raw;
        const valAK_raw = (valAG_raw * ajio) / 100;
        const valAL_raw = valAG_raw - valAI_raw - valAK_raw;
        const valAM_num = valAL_raw > 2499 ? 0.18 : 0.05;
        const valAM_pct = (valAM_num * 100) + '%';
        const valAN_raw = valAL_raw * valAM_num;
        const valAO_raw = valAL_raw + valAN_raw;

        document.getElementById('res_purchase_gst').innerText = purchGstAmt.toFixed(2);
        document.getElementById('res_purchase_tp').innerText = purchTp.toFixed(2);
        document.getElementById('res_AA').innerText = valAA.toFixed(2);
        document.getElementById('res_AD').innerText = valAD.toFixed(2);
        
        // Display Rounded Values as requested
        document.getElementById('res_AE').innerText = Math.round(valAE_raw).toFixed(0);
        document.getElementById('res_AF').innerText = Math.round(valAF_raw).toFixed(0);
        document.getElementById('res_AG').innerText = Math.round(valAG_raw).toFixed(0);
        
        document.getElementById('res_AH_pct').innerText = valAH_pct;
        document.getElementById('res_AI').innerText = valAI_raw.toFixed(2);
        document.getElementById('res_AJ').innerText = valAJ_raw.toFixed(2);
        document.getElementById('res_AK').innerText = valAK_raw.toFixed(2);
        document.getElementById('res_AL').innerText = valAL_raw.toFixed(2);
        document.getElementById('res_AM_pct').innerText = valAM_pct;
        document.getElementById('res_AN').innerText = valAN_raw.toFixed(2);
        document.getElementById('res_AO').innerText = valAO_raw.toFixed(2);
      }}

      if (singleYInput) singleYInput.addEventListener('input', updateSingleMode);
      if (singlePurchGstInput) singlePurchGstInput.addEventListener('input', updateSingleMode);

      function renderPage() {{
        if (tableData.length === 0) return;
        
        const totalPages = Math.ceil(tableData.length / pageSize);
        if (currentPage < 1) currentPage = 1;
        if (currentPage > totalPages) currentPage = totalPages;
        
        pageInfo.innerText = `Page ${{currentPage}} of ${{totalPages}}`;
        prevBtn.disabled = currentPage === 1;
        nextBtn.disabled = currentPage === totalPages;
        
        const start = (currentPage - 1) * pageSize;
        const end = start + pageSize;
        const pageData = tableData.slice(start, end);
        
        let html = '';
        pageData.forEach(row => {{
          let tr = '<tr>';
          row.forEach(cell => {{
            tr += `<td>${{escapeHtml(cell)}}</td>`;
          }});
          tr += '</tr>';
          html += tr;
        }});
        
        tableBody.innerHTML = html;
        updateCompanyProfit();
      }}

      function updateCompanyProfit() {{
        if (!tableData.length) return;
        
        const marginStr = marginInput ? marginInput.value.trim() : '';
        const margin = parseFloat(marginStr) || 0;
        const isMarginValid = !isNaN(parseFloat(marginStr));

        const gstStr = gstInput ? gstInput.value.trim() : '';
        const gst = parseFloat(gstStr) || 0;
        const isGstValid = !isNaN(parseFloat(gstStr));

        const discountStr = discountInput ? discountInput.value.trim() : '';
        const discount = parseFloat(discountStr) || 0;
        const isDiscountValid = !isNaN(parseFloat(discountStr));

        const ajioStr = ajioInput ? ajioInput.value.trim() : '';
        const ajio = parseFloat(ajioStr) || 0;
        const isAjioValid = !isNaN(parseFloat(ajioStr));

        // Update the entire data array so pagination works correctly
        tableData.forEach((row) => {{
          const yValueStr = row[24];
          const yValue = parseFloat(String(yValueStr).replace(/,/g, '')) || 0;

          const profit = isMarginValid ? Math.round(((yValue * margin) / 100) * 100) / 100 : 0;
          const acValue = yValue + profit;
          const abValue = isGstValid ? Math.round(((acValue * gst) / (100 + gst)) * 100) / 100 : 0;
          const aaValue = acValue - abValue;

          let aeRaw = 0;
          if (acValue > 0) {{
            aeRaw = calculateAE_raw(acValue);
          }}
          
          const afRaw = (aeRaw * discount) / 100;
          const agRaw = aeRaw - afRaw;
          const ahPctVal = agRaw > 2499 ? 0.18 : 0.05;
          const ahValue = agRaw - (agRaw / (1 + ahPctVal));
          const aiValue = agRaw - ahValue;
          const ajValue = (agRaw * ajio) / 100;
          const akValue = agRaw - ahValue - ajValue;
          const alPctVal = akValue > 2499 ? 0.18 : 0.05;
          const amValue = akValue * alPctVal;
          const anValue = akValue + amValue;

          const anyValid = isMarginValid || isGstValid || aeRaw > 0;
          
          // Store updated values back in the data source
          row[25] = '';
          row[26] = isMarginValid ? profit.toFixed(2) : '';
          row[27] = anyValid ? aaValue.toFixed(2) : '';
          row[28] = isGstValid ? abValue.toFixed(2) : '';
          row[29] = anyValid ? acValue.toFixed(2) : '';
          
          // Display Rounded Values as requested
          row[30] = aeRaw > 0 ? Math.round(aeRaw).toFixed(0) : '';
          row[31] = aeRaw > 0 ? Math.round(afRaw).toFixed(0) : '';
          row[32] = aeRaw > 0 ? Math.round(agRaw).toFixed(0) : '';
          
          row[33] = aeRaw > 0 ? (ahPctVal * 100) + '%' : '';
          row[34] = aeRaw > 0 ? ahValue.toFixed(2) : '';
          row[35] = aeRaw > 0 ? aiValue.toFixed(2) : '';
          row[36] = aeRaw > 0 ? ajValue.toFixed(2) : '';
          row[37] = aeRaw > 0 ? akValue.toFixed(2) : '';
          row[38] = aeRaw > 0 ? (alPctVal * 100) + '%' : '';
          row[39] = aeRaw > 0 ? amValue.toFixed(2) : '';
          row[40] = aeRaw > 0 ? anValue.toFixed(2) : '';
        }});

        // Now sync only the visible table rows in the DOM
        const tableRows = tableBody.querySelectorAll('tr');
        tableRows.forEach((tr, index) => {{
          const dataRowIndex = (currentPage - 1) * pageSize + index;
          if (tableData[dataRowIndex] && tr.cells.length >= 41) {{
            for (let i = 25; i <= 40; i++) {{
              tr.cells[i].innerText = tableData[dataRowIndex][i];
            }}
          }}
        }});
      }}

      function buildExportRows() {{
        return tableData.map((originalRow) => {{
          const row = [...originalRow];
          while (row.length < exportHeaders.length) {{
            row.push('');
          }}

          const yValue = parseFloat(String(row[24] || '').replace(/,/g, '')) || 0;
          const marginStr = marginInput ? marginInput.value.trim() : '';
          const margin = parseFloat(marginStr) || 0;
          const isMarginValid = !isNaN(parseFloat(marginStr));

          const gstStr = gstInput ? gstInput.value.trim() : '';
          const gst = parseFloat(gstStr) || 0;
          const isGstValid = !isNaN(parseFloat(gstStr));

          const discountStr = discountInput ? discountInput.value.trim() : '';
          const discount = parseFloat(discountStr) || 0;
          const isDiscountValid = !isNaN(parseFloat(discountStr));

          const ajioStr = ajioInput ? ajioInput.value.trim() : '';
          const ajio = parseFloat(ajioStr) || 0;
          const isAjioValid = !isNaN(parseFloat(ajioStr));

          const profit = isMarginValid ? Math.round(((yValue * margin) / 100) * 100) / 100 : 0;
          const acValue = yValue + profit;
          const abValue = isGstValid ? Math.round(((acValue * gst) / (100 + gst)) * 100) / 100 : 0;
          const aaValue = acValue - abValue;

          let aeRaw = 0;
          if (acValue > 0) {{
            aeRaw = calculateAE_raw(acValue);
          }}

          const afRaw = (aeRaw * discount) / 100;
          const agRaw = aeRaw - afRaw;
          const ahPctVal = agRaw > 2499 ? 0.18 : 0.05;
          const ahValue = agRaw - (agRaw / (1 + ahPctVal));
          const aiValue = agRaw - ahValue;
          const ajValue = (agRaw * ajio) / 100;
          const akValue = agRaw - ahValue - ajValue;
          const alPctVal = akValue > 2499 ? 0.18 : 0.05;
          const amValue = akValue * alPctVal;
          const anValue = akValue + amValue;

          row[25] = '';
          row[26] = isMarginValid ? profit.toFixed(2) : '';
          row[27] = (isMarginValid || isGstValid || aeRaw > 0) ? aaValue.toFixed(2) : '';
          row[28] = isGstValid ? abValue.toFixed(2) : '';
          row[29] = (isMarginValid || isGstValid || aeRaw > 0) ? acValue.toFixed(2) : '';
          
          row[30] = aeRaw > 0 ? Math.round(aeRaw).toFixed(0) : '';
          row[31] = aeRaw > 0 ? Math.round(afRaw).toFixed(0) : '';
          row[32] = aeRaw > 0 ? Math.round(agRaw).toFixed(0) : '';
          
          row[33] = aeRaw > 0 ? (ahPctVal * 100) + '%' : '';
          row[34] = aeRaw > 0 ? ahValue.toFixed(2) : '';
          row[35] = aeRaw > 0 ? aiValue.toFixed(2) : '';
          row[36] = aeRaw > 0 ? ajValue.toFixed(2) : '';
          row[37] = aeRaw > 0 ? akValue.toFixed(2) : '';
          row[38] = aeRaw > 0 ? (alPctVal * 100) + '%' : '';
          row[39] = aeRaw > 0 ? amValue.toFixed(2) : '';
          row[40] = aeRaw > 0 ? anValue.toFixed(2) : '';
          return row;
        }});
      }}

      function openExportPopup() {{
        if (!tableData.length) {{
          alert('Please upload Excel file first to export data.');
          return;
        }}
        exportFileNameInput.value = 'ajio_bulk_export';
        exportPopup.classList.add('open');
        exportFileNameInput.focus();
        exportFileNameInput.select();
      }}

      function closeExportPopup() {{
        exportPopup.classList.remove('open');
      }}

      function exportExcel() {{
        if (!tableData.length) {{
          closeExportPopup();
          return;
        }}

        if (typeof XLSX === 'undefined') {{
          alert('Excel library load nahi hui. Please page refresh karke try karein.');
          return;
        }}

        const safeName = sanitizeFilename(exportFileNameInput.value);
        const rows = buildExportRows();

        confirmExportBtn.disabled = true;
        confirmExportBtn.textContent = 'Preparing...';

        try {{
          // Build worksheet data: header row + data rows
          const wsData = [exportHeaders, ...rows];

          // Convert values: numbers as numbers, % strings as percentages
          const aoa = wsData.map((row, rIdx) => {{
            return row.map((cell, cIdx) => {{
              if (rIdx === 0) return cell; // header row - keep as string
              if (cell === null || cell === undefined || cell === '') return '';
              const s = String(cell).trim();
              // Handle percentage strings like "5%" or "18%"
              if (s.endsWith('%')) {{
                const num = parseFloat(s);
                if (!isNaN(num)) return num / 100;
              }}
              // Handle numeric strings
              const cleaned = s.replace(/,/g, '');
              if (cleaned !== '' && !isNaN(Number(cleaned))) return Number(cleaned);
              return s;
            }});
          }});

          const ws = XLSX.utils.aoa_to_sheet(aoa);

          const numCols = exportHeaders.length;
          const numRows = aoa.length;

          // Set column widths
          const colWidths = exportHeaders.map((h, i) => {{
            let max = String(h || '').length + 3;
            rows.forEach(r => {{
              const v = String(r[i] || '');
              const w = i === 2 ? Math.min(Math.max(v.length + 3, 32), 80) : Math.min(Math.max(v.length + 3, 12), 28);
              if (w > max) max = w;
            }});
            return {{ wch: Math.max(max, 12) }};
          }});
          ws['!cols'] = colWidths;

          // Freeze row 1
          ws['!freeze'] = {{ xSplit: 0, ySplit: 1, topLeftCell: 'A2', activePane: 'bottomLeft', state: 'frozen' }};

          // Apply styles to cells (header colors + number formats)
          const redFill   = {{ patternType: 'solid', fgColor: {{ rgb: 'FDE2E2' }} }};
          const yellowFill = {{ patternType: 'solid', fgColor: {{ rgb: 'FFF4BF' }} }};
          const greenFill  = {{ patternType: 'solid', fgColor: {{ rgb: 'DDF6E4' }} }};
          const thinBorder = {{
            top:    {{ style: 'thin', color: {{ rgb: '000000' }} }},
            bottom: {{ style: 'thin', color: {{ rgb: '000000' }} }},
            left:   {{ style: 'thin', color: {{ rgb: '000000' }} }},
            right:  {{ style: 'thin', color: {{ rgb: '000000' }} }},
          }};
          const headerFont = {{ bold: true }};
          const centerAlign = {{ horizontal: 'center', vertical: 'center' }};

          for (let C = 0; C < numCols; C++) {{
            const cellAddr = XLSX.utils.encode_cell({{ r: 0, c: C }});
            if (!ws[cellAddr]) ws[cellAddr] = {{ t: 's', v: exportHeaders[C] || '' }};
            const fill = C < 25 ? redFill : C === 25 ? yellowFill : greenFill;
            ws[cellAddr].s = {{ fill, font: headerFont, alignment: centerAlign, border: thinBorder }};
          }}

          // Apply number format for percentage columns and numeric data rows
          for (let R = 1; R < numRows; R++) {{
            for (let C = 0; C < numCols; C++) {{
              const cellAddr = XLSX.utils.encode_cell({{ r: R, c: C }});
              if (!ws[cellAddr]) continue;
              const header = exportHeaders[C] || '';
              const cell = ws[cellAddr];
              if (cell.t === 'n') {{
                // Check if original value was a percentage
                const origVal = rows[R - 1][C];
                const origStr = String(origVal || '').trim();
                if (origStr.endsWith('%') && header.toLowerCase().includes('%')) {{
                  cell.z = '0%';
                }}
              }}
            }}
          }}

          // Write workbook
          const wb = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(wb, ws, 'AJIO Export');

          XLSX.writeFile(wb, safeName + '.xlsx', {{ bookType: 'xlsx', type: 'binary', cellStyles: true }});
          closeExportPopup();
        }} catch (err) {{
          alert('Export failed: ' + (err.message || 'Unknown error'));
        }} finally {{
          confirmExportBtn.disabled = false;
          confirmExportBtn.textContent = 'Download File';
        }}
      }}

      window.exportToExcel = openExportPopup;
      window.confirmExport = exportExcel;
      window.cancelExport = closeExportPopup;
      
      window.openClearPopup = () => clearPopup.classList.add('open');
      window.closeClearPopup = () => clearPopup.classList.remove('open');

      if (prevBtn && nextBtn) {{
        prevBtn.addEventListener('click', () => {{
          if (currentPage > 1) {{
            currentPage--;
            renderPage();
          }}
        }});
        nextBtn.addEventListener('click', () => {{
          const totalPages = Math.ceil(tableData.length / pageSize);
          if (currentPage < totalPages) {{
            currentPage++;
            renderPage();
          }}
        }});
      }}


      if (marginInput) {{
        marginInput.addEventListener('input', updateCompanyProfit);
        marginInput.addEventListener('input', updateSingleMode);
      }}
      if (gstInput) {{
        gstInput.addEventListener('input', updateCompanyProfit);
        gstInput.addEventListener('input', updateSingleMode);
      }}
      if (discountInput) {{
        discountInput.addEventListener('input', updateCompanyProfit);
        discountInput.addEventListener('input', updateSingleMode);
      }}
      if (ajioInput) {{
        ajioInput.addEventListener('input', updateCompanyProfit);
        ajioInput.addEventListener('input', updateSingleMode);
      }}

      if (openExportPopupBtn) {{
        openExportPopupBtn.addEventListener('click', openExportPopup);
      }}
      if (clearDataBtn) {{
        clearDataBtn.addEventListener('click', () => {{
          if (clearPopup) clearPopup.classList.add('open');
        }});
      }}
      if (cancelClearBtn) {{
        cancelClearBtn.addEventListener('click', () => {{
          if (clearPopup) clearPopup.classList.remove('open');
        }});
      }}
      if (confirmClearBtn) {{
        confirmClearBtn.addEventListener('click', () => {{
          window.location.href = '/';
        }});
      }}
      if (clearPopup) {{
        clearPopup.addEventListener('click', (event) => {{
          if (event.target === clearPopup) {{
            clearPopup.classList.remove('open');
          }}
        }});
      }}
      if (cancelExportBtn) {{
        cancelExportBtn.addEventListener('click', closeExportPopup);
      }}
      if (confirmExportBtn) {{
        confirmExportBtn.addEventListener('click', exportExcel);
      }}
      if (exportPopup) {{
        exportPopup.addEventListener('click', (event) => {{
          if (event.target === exportPopup) {{
            closeExportPopup();
          }}
        }});
      }}
      if (exportFileNameInput) {{
        exportFileNameInput.addEventListener('keydown', (event) => {{
          if (event.key === 'Enter') {{
            event.preventDefault();
            exportExcel();
          }}
        }});
      }}

      // Chatbot Logic
      const botBubble = document.getElementById('bot-bubble');
      const botWindow = document.getElementById('bot-window');
      const botMessages = document.getElementById('bot-messages');
      const botInput = document.getElementById('bot-input');
      
      let chatState = {{
        userName: null,
        step: 0,
        data: {{}}
      }};

      const steps = [
        {{ key: 'y', label: 'PURCHASE TP (Y)', prompt: (name) => `Alright ${{name}}! First, please tell me your PURCHASE TP (Y): 💰` }},
        {{ key: 'pgst', label: 'PURCHASE GST %', prompt: (name) => `Thank you ${{name}}! Now, what is the PURCHASE GST %? 💸` }},
        {{ key: 'margin', label: 'COMPANY PROFIT MARGIN %', prompt: (name) => `Got it ${{name}}! What is the COMPANY PROFIT MARGIN %? 📈` }},
        {{ key: 'sgst', label: 'SALE GST %', prompt: (name) => `Great ${{name}}! Now tell me the SALE GST %: 🏛️` }},
        {{ key: 'discount', label: 'SALE DISCOUNT AMT %', prompt: (name) => `Almost there ${{name}}! What is the SALE DISCOUNT AMT %? 🏷️` }},
        {{ key: 'ajio', label: 'AJIO MARGIN %', prompt: (name) => `Last step ${{name}}! What is the AJIO MARGIN %? 🤝` }}
      ];

      const randomGreetings = [
        "Hii I'am AJ your friend how can i help you",
        "Need help with calculations? ⚡",
        "Type 'calculate' to start a step-by-step guide! 📖",
        "I can help you find the Bank Settlement! 💰",
        "Don't forget to export your results! ✅"
      ];

      function typeWriter(text, i, cb) {{
        if (i < text.length) {{
          botBubble.textContent += text.charAt(i);
          setTimeout(() => typeWriter(text, i + 1, cb), 50);
        }} else if (cb) {{
          cb();
        }}
      }}

      function rotateGreetings() {{
        if (botWindow.classList.contains('open')) {{
            botBubble.classList.remove('show');
            return;
        }}
        const randomMsg = randomGreetings[Math.floor(Math.random() * randomGreetings.length)];
        
        botBubble.classList.remove('show');
        setTimeout(() => {{
            botBubble.textContent = '';
            botBubble.classList.add('show');
            typeWriter(randomMsg, 0, () => {{
                setTimeout(() => botBubble.classList.remove('show'), 6000);
            }});
        }}, 500);
      }}

      setInterval(rotateGreetings, 12000);
      setTimeout(rotateGreetings, 2000);

      window.toggleChat = () => {{
        botWindow.classList.toggle('open');
        if (botWindow.classList.contains('open')) botBubble.classList.remove('show');
      }};

      window.sendMessage = async () => {{
        const text = botInput.value.trim();
        if (!text) return;

        appendMessage('user', text);
        botInput.value = '';

        // Handle User Name first
        if (!chatState.userName) {{
            chatState.userName = text;
            appendMessage('bot', `Nice to meet you, ${{chatState.userName}}! 👋 How can I help you today? You can type "calculate" to start a guided cost check! ✨`);
            return;
        }}

        // Check if user wants to start or continue calculation
        const numVal = parseFloat(text);
        
        if (text.toLowerCase().includes('calculate') || text.toLowerCase().includes('tp')) {{
            chatState.step = 0;
            chatState.data = {{}};
            appendMessage('bot', steps[0].prompt(chatState.userName));
            chatState.step = 1;
            return;
        }}

        if (chatState.step > 0 && chatState.step <= steps.length) {{
            if (isNaN(numVal)) {{
                appendMessage('bot', `Kripya ek sahi number batayein ${{steps[chatState.step-1].label}} ke liye.`);
                return;
            }}
            
            chatState.data[steps[chatState.step-1].key] = numVal;
            
            if (chatState.step < steps.length) {{
                appendMessage('bot', steps[chatState.step].prompt(chatState.userName));
                chatState.step++;
            }} else {{
                // Final Calculation
                const result = calculateFinal(chatState.data);
                appendMessage('bot', `Alright ${{chatState.userName}}, your calculation is ready! ✨\n\n` + 
                    `🛒 SALE TP WITH TAX (AD): ₹${{result.ad}}\n` +
                    `🏷️ MRP (AE): ₹${{result.ae}}\n` +
                    `📦 ASP (GROSS) (AG): ₹${{result.ag}}\n` +
                    `✅ BANK SETTLEMENT (AO): ₹${{result.ao}}\n\n` +
                    `Is there anything else I can help you with?`);
                chatState.step = 0;
            }}
            return;
        }}

        // Normal AI Chat
        try {{
          const response = await fetch('/api/chat', {{
            method: 'POST',
            headers: {{ 'Content-Type': 'application/json' }},
            body: JSON.stringify({{
              message: text,
              userName: chatState.userName
            }})
          }});

          if (!response.ok) throw new Error('API Error');
          const data = await response.json();
          const botText = data.reply;
          appendMessage('bot', botText);
        }} catch (e) {{
          appendMessage('bot', "I'am sorry, I'am having trouble connecting right now. 😔");
        }}
      }};

      function calculateFinal(d) {{
        const valAA = Math.round(((d.y * d.margin) / 100) * 100) / 100;
        const valAD = Math.round((d.y + valAA) * 100) / 100;
        
        const aeRaw = (valAD <= 1606) ? (valAD / 0.22505) : 
                      (valAD <= 2620) ? (valAD / 0.186490678) : 
                      (valAD / 0.209579661);

        const afRaw = (aeRaw * d.discount) / 100;
        const agRaw = aeRaw - afRaw;
        const ahPctVal = agRaw > 2499 ? 0.18 : 0.05;
        const ahValue = agRaw - (agRaw / (1 + ahPctVal));
        const aiValue = agRaw - ahValue;
        const ajValue = (agRaw * d.ajio) / 100;
        const akValue = agRaw - ahValue - ajValue;
        const alPctVal = akValue > 2499 ? 0.18 : 0.05;
        const amValue = akValue * alPctVal;
        const anValue = akValue + amValue;
        
        return {{ 
          ad: valAD.toFixed(2),
          ae: Math.round(aeRaw).toFixed(0),
          ag: Math.round(agRaw).toFixed(0),
          ao: anValue.toFixed(2) 
        }};
      }}

      function appendMessage(type, text) {{
        const div = document.createElement('div');
        div.className = `msg ${{type}}`;
        div.style.whiteSpace = 'pre-wrap';
        div.textContent = text;
        botMessages.appendChild(div);
        botMessages.scrollTop = botMessages.scrollHeight;
      }}

      renderPage();
    }});
  </script>
</body>
</html>
"""
    return page.encode("utf-8")


@app.get("/sample-excel")
def download_sample_excel() -> Response:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sample Template"
    
    # A1: Item Details
    ws["A1"] = "Item Details"
    ws["A1"].font = Font(size=14)
    
    # Row 2: Headers
    headers = [
        "No.", "Seller Name", "Item Name", "Category", "Brand", "Type", 
        "ASIN Number", "SKU Number", "HSN Number", "MRP Price", "Item Color", 
        "Weight", "Weight Unit", "Length", "Length Unit", "Width", 
        "Width Unit", "Height", "Height Unit", "Channel Price", 
        "Purchase Margin(%)", "Tax", "Purchase Cost", "Purchase Tax", 
        "Final Purchase Cost"
    ]
    ws.append(headers)
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    response = Response(
        output.getvalue(),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    response.headers["Content-Disposition"] = "attachment; filename=ajio_sample_template.xlsx"
    return response


@app.get("/")
def index() -> Response:
    return Response(render_page(), mimetype="text/html")


@app.get("/logo.png")
def serve_logo() -> Response:
    return send_from_directory(BASE_DIR, "logo.png")


@app.get("/chatbot.png")
def serve_chatbot() -> Response:
    return send_from_directory(BASE_DIR, "chatbot.png")


@app.post("/api/chat")
def chat_route() -> Response:
    api_key = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
    if not api_key:
        return jsonify({"error": "AI chat is not configured."}), 503

    payload = request.get_json(force=True, silent=True) or {}
    message = str(payload.get("message", "")).strip()
    user_name = str(payload.get("userName", "friend")).strip() or "friend"

    if not message:
        return jsonify({"error": "Message is required."}), 400

    gemini_payload = {
        "system_instruction": {
            "parts": [
                {
                    "text": f"""You are AJ, the user's "Dashboard Friend" and business assistant for this AJIO Cost Sheet.
Your tone is extremely friendly, professional, and helpful. Always address the user as "{user_name}".

CORE CAPABILITIES:
1. Simple Math: Perform basic arithmetic (+, -, *, /).
2. Percentage Logic: Handle queries like "X% of Y", "percentage of X", etc.
3. Advanced Business Math:
   - Reverse GST: Extract base price from a tax-inclusive price (Formula: Price / (1 + TaxRate)).
   - Profit Margin vs Markup: Calculate required selling price for a target margin.
   - Discount Stacking: Calculate final price after sequential discounts (e.g., 20% + 10%).
   - Break-even: Calculate minimum price to cover costs.
4. General Support: Answer questions about the AJIO dashboard or cost calculations.

GUIDELINES:
- Respond in "Hinglish" (a natural mix of Hindi and English).
- For math queries, clearly state the result clearly and briefly show the formula or steps used.
- If the user seems confused, offer to start the guided calculation by telling them to type "calculate".
- Keep responses concise but warm."""
                }
            ]
        },
        "contents": [{"parts": [{"text": message}]}],
    }

    try:
        gemini_response = requests.post(
            GEMINI_API_URL,
            params={"key": api_key},
            json=gemini_payload,
            timeout=30,
        )
        gemini_response.raise_for_status()
        data = gemini_response.json()
        reply = data["candidates"][0]["content"]["parts"][0]["text"]
        return jsonify({"reply": reply})
    except (KeyError, IndexError, requests.RequestException) as exc:
        app.logger.exception("Gemini chat request failed: %s", exc)
        return jsonify({"error": "AI chat request failed."}), 502


@app.post("/export")
def export_excel_route() -> Response:
    try:
        payload = request.get_json(force=True, silent=True)
        if not payload or not isinstance(payload, dict):
            app.logger.error("Export called with invalid or empty JSON payload")
            return jsonify({"error": "Invalid request payload."}), 400
        filename = sanitize_filename(str(payload.get("filename", "ajio_export")))
        headers = [ILLEGAL_CHARACTERS_RE.sub("", str(cell)) for cell in payload.get("headers", [])]
        rows = payload.get("rows", []) or []
        if not headers:
            return jsonify({"error": "No headers provided."}), 400
        workbook_bytes = build_export_workbook(headers, rows)
    except Exception as exc:  # noqa: BLE001
        app.logger.exception("Excel export failed: %s", exc)
        return jsonify({"error": f"Excel export failed: {exc}"}), 500
    response = Response(
        workbook_bytes,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response.headers["Content-Disposition"] = f'attachment; filename="{filename}.xlsx"'
    return response


@app.post("/")
def upload_excel_route() -> Response:
    try:
        upload = request.files.get("file")
        if upload is None or not upload.filename:
            raise ValueError("Koi file receive nahi hui.")

        filename = upload.filename
        file_bytes = upload.read()
        extension = os.path.splitext(filename)[1].lower()

        try:
            if extension == ".xls":
                headers, rows = parse_xls(file_bytes)
            else:
                headers, rows = parse_excel(file_bytes)
        except Exception:
            headers, rows = parse_with_calamine(file_bytes)

        # Apply Purchase GST logic if toggle is on
        if request.form.get("purchase_gst_toggle") == "on":
            for row in rows:
                if len(row) > 24: # Column Y is index 24
                    try:
                        # Ensure Y is a valid number
                        y_str = str(row[24]).replace(",", "").strip()
                        if not y_str: continue
                        val_y = float(y_str)
                        
                        val_v = 5
                        val_x = round(val_y * 5 / 105, 2)
                        val_w = round(val_y - val_x, 2)
                        
                        row[21] = str(val_v)  # V
                        row[23] = f"{val_x:.2f}" # X
                        row[22] = f"{val_w:.2f}" # W
                    except (ValueError, TypeError, IndexError):
                        continue

        page = render_page(render_table(headers, rows), f"{filename} successfully loaded.", default_tab="bulk")
        return Response(page, mimetype="text/html")
    except Exception as exc:  # noqa: BLE001
        error_page = render_page(message=f"File process nahi hui: {exc}")
        return Response(error_page, mimetype="text/html", status=400)


def run() -> None:
    port = int(os.environ.get("PORT", "8000"))
    app.run(host="0.0.0.0", port=port, debug=False)


if __name__ == "__main__":
    run()
