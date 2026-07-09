"""
Convert supplier product sheets into the Odoo import template format.
Runs in the browser via Pyodide.
"""

from __future__ import annotations

import csv
import io
import json
import re
from typing import Any

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

REQUIRED_SUPPLIER_KEYS = [
    "name",
    "barcode",
    "cost/case",
    "selling price",
    "plu",
    "markup",
    "vat",
]


def _normalize(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"\s+", " ", text)
    return text


def _looks_like_header(cells: list[Any]) -> bool:
    normalized = {_normalize(c) for c in cells if _normalize(c)}
    return "name" in normalized and (
        "barcode" in normalized or "plu" in normalized or "selling price" in normalized
    )


def _resolve_required_aliases(headers: list[str]) -> dict[str, str]:
    """Map logical required keys to actual header strings found in the file."""
    found: dict[str, str] = {}
    for header in headers:
        key = _normalize(header)
        if not key:
            continue
        if key == "name":
            found["name"] = header
        elif key == "barcode":
            found["barcode"] = header
        elif key in ("cost/case", "cost case"):
            found["cost/case"] = header
        elif key == "selling price":
            found["selling price"] = header
        elif key == "plu":
            found["plu"] = header
        elif key in ("markup", "markup(%)", "markup %"):
            found["markup"] = header
        elif key in ("vat (%)", "vat(%)", "vat %", "vat"):
            found["vat"] = header
    return found


def detect_header_row(rows: list[list[Any]]) -> int:
    for idx, row in enumerate(rows[:50]):
        if _looks_like_header(row):
            return idx
    raise ValueError(
        "Could not find a header row containing product columns "
        "(expected at least Name plus Barcode, PLU, or Selling Price)."
    )


def parse_supplier_rows(rows: list[list[Any]]) -> tuple[list[str], list[dict[str, Any]]]:
    if not rows:
        raise ValueError("The uploaded file is empty.")

    header_idx = detect_header_row(rows)
    raw_headers = rows[header_idx]
    headers: list[str] = []
    for i, h in enumerate(raw_headers):
        label = str(h).strip() if h is not None else ""
        headers.append(label or f"Column_{i + 1}")

    resolved = _resolve_required_aliases(headers)
    missing = [k for k in REQUIRED_SUPPLIER_KEYS if k not in resolved]
    if missing:
        label_map = {
            "name": "Name",
            "barcode": "Barcode",
            "cost/case": "Cost/Case",
            "selling price": "Selling Price",
            "plu": "PLU",
            "markup": "Markup / Markup(%)",
            "vat": "VAT (%) / VAT(%)",
        }
        pretty = ", ".join(label_map[m] for m in missing)
        raise ValueError(
            f"Missing required supplier column(s): {pretty}. "
            f"Found headers: {', '.join(h for h in headers if h and not h.startswith('Column_'))}"
        )

    records: list[dict[str, Any]] = []
    for row in rows[header_idx + 1 :]:
        if not row or all(c is None or str(c).strip() == "" for c in row):
            continue
        # Skip footer / non-product lines
        first = str(row[0]).strip().lower() if row and row[0] is not None else ""
        if first.startswith("page ") or first == "products":
            continue

        record: dict[str, Any] = {}
        for i, header in enumerate(headers):
            val = row[i] if i < len(row) else None
            if val is None:
                record[header] = ""
            elif isinstance(val, float) and val == int(val):
                record[header] = int(val)
            else:
                record[header] = val

        name_val = record.get(resolved["name"], "")
        if str(name_val).strip() == "":
            continue
        records.append(record)

    if not records:
        raise ValueError("No product rows found under the detected header row.")

    return headers, records


def _pick_source_header(headers: list[str], preferred_keys: list[str]) -> str | None:
    """
    Choose the best matching supplier header for a target Odoo field.
    preferred_keys are normalized names ordered from most to least specific
    (e.g. prefer 'vat (%)' over a bare 'vat' label column).
    """
    by_norm = {_normalize(h): h for h in headers if _normalize(h)}
    for key in preferred_keys:
        if key in by_norm:
            return by_norm[key]
    return None


def map_to_odoo(headers: list[str], records: list[dict[str, Any]]) -> tuple[list[str], list[dict[str, Any]]]:
    odoo_headers = [
        "Name",
        "Internal Reference",
        "Cost",
        "Sales Price",
        "Point of Sale Category / External ID",
        "Track Inventory",
        "Product Category",
        "Barcode",
        "VAT %",
        "Markup %",
        "Brand Name",
        "Gross Margin %",
        "U Cost Price",
        "Exclude from API Sync",
    ]

    # Prefer specific % / price columns when ambiguous headers exist (e.g. VAT(%) vs VAT).
    field_preferences: dict[str, list[str]] = {
        "Name": ["name"],
        "Barcode": ["barcode"],
        "Cost": ["cost/case", "cost case", "cost"],
        "Sales Price": ["selling price"],
        "Internal Reference": ["plu"],
        "Markup %": ["markup(%)", "markup %", "markup"],
        "VAT %": ["vat (%)", "vat(%)", "vat %", "vat"],
    }

    source_for_odoo: dict[str, str] = {}
    for odoo_field, keys in field_preferences.items():
        chosen = _pick_source_header(headers, keys)
        if chosen:
            source_for_odoo[odoo_field] = chosen

    mapped: list[dict[str, Any]] = []
    for record in records:
        out: dict[str, Any] = {h: "" for h in odoo_headers}
        for dst, src in source_for_odoo.items():
            val = record.get(src, "")
            if val is None:
                val = ""
            out[dst] = val
        # Static Odoo defaults (not sourced from supplier sheet)
        out["Track Inventory"] = 1
        out["Product Category"] = "All"
        out["Exclude from API Sync"] = "TRUE"
        mapped.append(out)
    return odoo_headers, mapped


def _cell_value(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, str):
        if value.strip() == "":
            return None
        # Write numeric types for pure numeric strings (keep original text otherwise)
        text = value.strip()
        if re.fullmatch(r"-?\d+", text):
            return int(text)
        if re.fullmatch(r"-?\d+\.\d+", text):
            return float(text)
        return value
    return value


def _emit_progress(pct: int, progress_cb: Any = None) -> None:
    value = int(pct)
    if progress_cb is not None:
        progress_cb(value)
        return
    try:
        from js import window  # type: ignore

        window.__setConversionProgress(value)
    except Exception:
        pass


def build_xlsx_bytes(
    template_bytes: bytes,
    mapped_rows: list[dict[str, Any]],
    progress_cb: Any = None,
    progress_start: int = 40,
    progress_end: int = 85,
) -> bytes:
    wb = load_workbook(io.BytesIO(template_bytes))
    ws = wb.active

    # Find header row (contains "Name")
    header_row_idx = None
    header_to_col: dict[str, int] = {}
    for row in ws.iter_rows(min_row=1, max_row=10):
        values = [c.value for c in row]
        if any(_normalize(v) == "name" for v in values):
            header_row_idx = row[0].row
            for cell in row:
                if cell.value is not None and str(cell.value).strip():
                    header_to_col[str(cell.value).strip()] = cell.column
            break

    if header_row_idx is None:
        raise ValueError("Odoo template is missing a header row with a Name column.")

    # Apply consistent header styling across all template columns
    header_fill = PatternFill(fill_type="solid", fgColor="D3D3D3")
    header_font = Font(bold=True, color="000000")
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="B0B0B0")
    header_border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=header_row_idx, column=col)
        if cell.value is None or str(cell.value).strip() == "":
            continue
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment
        cell.border = header_border

    # Delete existing data rows below the header, preserving column structure/styles on header
    if ws.max_row > header_row_idx:
        ws.delete_rows(header_row_idx + 1, ws.max_row - header_row_idx)

    mapped_fields = {
        "Name",
        "Barcode",
        "Cost",
        "Sales Price",
        "Internal Reference",
        "Markup %",
        "VAT %",
        "Track Inventory",
        "Product Category",
        "Exclude from API Sync",
    }

    total = max(len(mapped_rows), 1)
    for i, record in enumerate(mapped_rows):
        excel_row = header_row_idx + 1 + i
        for field in mapped_fields:
            col = header_to_col.get(field)
            if col is None:
                continue
            ws.cell(row=excel_row, column=col, value=_cell_value(record.get(field)))

        if i == 0 or (i + 1) % 25 == 0 or i + 1 == total:
            pct = progress_start + int((i + 1) / total * (progress_end - progress_start))
            _emit_progress(pct, progress_cb)

    buf = io.BytesIO()
    wb.save(buf)
    _emit_progress(progress_end, progress_cb)
    return buf.getvalue()


def build_csv_text(odoo_headers: list[str], mapped_rows: list[dict[str, Any]]) -> str:
    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=odoo_headers, extrasaction="ignore", lineterminator="\n")
    writer.writeheader()
    for row in mapped_rows:
        writer.writerow({h: "" if row.get(h) is None else row.get(h) for h in odoo_headers})
    return buf.getvalue()


def convert(
    supplier_rows_json: str,
    template_bytes: bytes,
    detected_format: str,
    progress_cb: Any = None,
) -> str:
    """
    Main entry point for the browser UI.
    supplier_rows_json: JSON list of rows (list of lists)
    Returns JSON string with preview + base64 xlsx + csv text.
    progress_cb: optional callable(int) with 0–100 progress.
                 If omitted, falls back to window.__setConversionProgress in browser.
    """
    import base64

    _emit_progress(5, progress_cb)
    rows = json.loads(supplier_rows_json)
    _emit_progress(12, progress_cb)
    headers, records = parse_supplier_rows(rows)
    _emit_progress(28, progress_cb)
    odoo_headers, mapped = map_to_odoo(headers, records)
    _emit_progress(40, progress_cb)
    xlsx_bytes = build_xlsx_bytes(template_bytes, mapped, progress_cb=progress_cb)
    _emit_progress(90, progress_cb)
    csv_text = build_csv_text(odoo_headers, mapped)
    _emit_progress(96, progress_cb)

    preview_limit = 100
    result = {
        "ok": True,
        "detected_format": detected_format,
        "supplier_headers": headers,
        "row_count": len(mapped),
        "odoo_headers": odoo_headers,
        "preview": mapped[:preview_limit],
        "xlsx_base64": base64.b64encode(xlsx_bytes).decode("ascii"),
        "csv_text": csv_text,
    }
    _emit_progress(100, progress_cb)
    return json.dumps(result)
