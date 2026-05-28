"""openpyxl Excel export: summary (12 cols A-L) and full detail sheet."""

from __future__ import annotations

import io
import re
from datetime import datetime
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

ACCOUNTING_FMT = r'_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)'
FULL_DETAIL_HEADERS: list[str] = [
    "Row Type",
    "Invoice Number",
    "Account Number",
    "Amount Due",
    "Invoice Date",
    "Due Date",
    "Invoice Status",
    "Payment Status",
    "Subtotal",
    "Tax",
    "Government Charges",
    "Billed Amount",
    "Type",
    "Shipping Adjustments",
    "Invoice Adjustments (total)",
    "Explanation",
    "Incentive Savings",
    "Pickup Date",
    "Order Date",
    "Tracking Number",
    "Service Type",
    "Postal Code",
    "Zone",
    "Weight (lbs)",
    "Published Charge",
    "Incentive Credit",
    "Billed Charge",
    "Fuel Surcharge",
    "Declared Value",
    "Surge Fee",
    "Add'l Handling",
    "Demand Surcharge",
    "Residential Surcharge",
    "Delivery Area Surcharge",
    "Shipment SO#",
    "Shipment PO#",
    "Sender Name",
    "Sender Company",
    "Receiver Name",
    "Receiver Company",
    "Receiver City/Province",
    "UserID",
    "Adjustment Pickup Date",
    "Adjustment Tracking Number",
    "Adjustment Type",
    "Adjustment Amount",
    "Adjustment SO#",
    "Adjustment PO#",
    "Adjustment Description",
]

FULL_DETAIL_WIDTHS: list[float] = [
    12,
    16,
    10,
    12,
    21,
    12,
    16,
    12,
    12,
    10,
    20,
    15,
    17,
    22,
    29,
    40,
    19,
    13,
    12,
    20,
    14,
    13,
    12,
    14,
    18,
    12,
    15,
    16,
    12,
    12,
    16,
    18,
    23,
    25,
    14,
    17,
    19,
    16,
    22,
    56,
    24,
    12,
    24,
    28,
    12,
    19,
    16,
    17,
    40,
]

FULL_DETAIL_CURRENCY_COLS = {4, 9, 10, 11, 12, 14, 15, 17, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 46}
FULL_DETAIL_DATE_COLS = {5, 6}
FULL_DETAIL_CURRENCY_FMT = r'\$#,##0.00'
FULL_DETAIL_DATE_FMT = r'mmm\ dd\,\ yyyy'
FULL_DETAIL_HEADER_FILL = PatternFill(fill_type="solid", fgColor="F2F2F2")
FULL_DETAIL_HEADER_FONT = Font(bold=True, color="FF000000")
FULL_DETAIL_HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
FULL_DETAIL_BLACK_FONT = Font(color="FF000000")

HEADER_FILL = PatternFill(
    start_color="BDD7EE",
    end_color="BDD7EE",
    fill_type="solid",
)
HEADER_FONT = Font(name="Arial", size=12, bold=True)
DATA_FONT = Font(name="Arial", size=12)
ALT_FILL = PatternFill(start_color="F5F9FC", end_color="F5F9FC", fill_type="solid")
CURRENCY_FMT = "$#,##0.00"
DATE_FMT = "mmm dd, yyyy"

# (header, dict key, is_money, is_date, width)
ColSpec = tuple[str, str, bool, bool, int]

# Matches UI summary table columns A-L
SUMMARY_COLS: list[ColSpec] = [
    ("Invoice Number", "Invoice Number", False, False, 20),
    ("Account Number", "Account Number", False, False, 14),
    ("Amount Due", "Amount Due", True, False, 14),
    ("Invoice Date", "Invoice Date", False, True, 18),
    ("Invoice Status", "Invoice Status", False, False, 12),
    ("Payment Status", "Payment Status", False, False, 14),
    ("Subtotal", "Subtotal", True, False, 12),
    ("Tax", "Tax", True, False, 12),
    ("Government Charges", "Government Charges", True, False, 16),
    ("Billed Amount", "Billed Amount", True, False, 14),
    ("Due Date", "Due Date", False, True, 18),
    ("Type", "Type", False, False, 16),
]

INVOICE_DETAIL_COLS: list[ColSpec] = [
    ("Invoice Number", "Invoice Number", False, False, 18),
    ("Account Number", "Account Number", False, False, 14),
    ("Amount Due", "Amount Due", True, False, 14),
    ("Invoice Date", "Invoice Date", False, True, 18),
    ("Due Date", "Due Date", False, True, 18),
    ("Invoice Status", "Invoice Status", False, False, 12),
    ("Payment Status", "Payment Status", False, False, 14),
    ("Subtotal", "Subtotal", True, False, 12),
    ("Tax", "Tax", True, False, 12),
    ("Government Charges", "Government Charges", True, False, 16),
    ("Billed Amount", "Billed Amount", True, False, 14),
    ("Type", "Type", False, False, 16),
    ("Shipping Adjustments", "Shipping Adjustments", True, False, 18),
    ("Invoice Adjustments (total)", "Adjustments", True, False, 22),
    ("Explanation", "Explanation", False, False, 40),
    ("Incentive Savings", "Incentive Savings", True, False, 16),
]

SHIPMENT_DETAIL_COLS: list[ColSpec] = [
    ("Invoice Number", "Invoice Number", False, False, 18),
    ("Invoice Date", "Invoice Date", False, True, 18),
    ("Pickup Date", "Pickup Date", False, False, 12),
    ("Order Date", "Order Date", False, False, 12),
    ("Tracking Number", "Tracking Number", False, False, 22),
    ("Service Type", "Service Type", False, False, 18),
    ("Postal Code", "Destination Postal Code", False, False, 12),
    ("Zone", "Zone", False, False, 8),
    ("Weight (lbs)", "Weight (lbs)", False, False, 12),
    ("Published Charge", "Published Charge", True, False, 14),
    ("Incentive Credit", "Incentive Credit", True, False, 14),
    ("Billed Charge", "Billed Charge", True, False, 14),
    ("Fuel Surcharge", "Fuel Surcharge", True, False, 14),
    ("Declared Value", "Declared Value", True, False, 14),
    ("Surge Fee", "Surge Fee", True, False, 12),
    ("Add'l Handling", "Add'l Handling", True, False, 14),
    ("Demand Surcharge", "Demand Surcharge", True, False, 14),
    ("Residential Surcharge", "Residential Surcharge", True, False, 16),
    ("Delivery Area Surcharge", "Delivery Area Surcharge", True, False, 18),
    ("SO#", "SO#", False, False, 14),
    ("PO#", "PO#", False, False, 14),
    ("Sender Name", "Sender Name", False, False, 20),
    ("Sender Company", "Sender Company", False, False, 20),
    ("Receiver Name", "Receiver Name", False, False, 20),
    ("Receiver Company", "Receiver Company", False, False, 22),
    ("Receiver City/Province", "Receiver City/Province", False, False, 22),
    ("UserID", "UserID", False, False, 14),
]


def _parse_date_for_sort(s: str) -> datetime:
    if not s or not str(s).strip():
        return datetime.min
    raw = str(s).strip()
    for fmt in ("%B %d, %Y", "%B %d %Y", "%b %d, %Y", "%b %d %Y"):
        try:
            return datetime.strptime(raw, fmt)
        except ValueError:
            continue
    collapsed = re.sub(r",\s*", " ", raw)
    for fmt in ("%B %d %Y", "%b %d %Y"):
        try:
            return datetime.strptime(collapsed, fmt)
        except ValueError:
            continue
    for fmt in ("%m/%d/%Y", "%m-%d-%Y", "%d/%m/%Y", "%Y-%m-%d", "%m/%d/%y"):
        try:
            return datetime.strptime(raw.replace("-", "/"), fmt.replace("-", "/"))
        except ValueError:
            continue
    try:
        return datetime.fromisoformat(raw[:10])
    except ValueError:
        return datetime.min


def _fmt_date_str(val: Any) -> str:
    """Return 'Month DD, YYYY' plain string (zero-padded day, e.g. 'April 07, 2026')."""
    if val is None or val == "":
        return ""
    if isinstance(val, datetime):
        return val.strftime("%B %d, %Y")
    raw = str(val).strip()
    if not raw:
        return ""
    dt = _parse_date_for_sort(raw)
    return dt.strftime("%B %d, %Y") if dt != datetime.min else raw


def _to_date(val: Any) -> datetime | str | None:
    if val is None or val == "":
        return None
    if isinstance(val, datetime):
        return val
    raw = str(val).strip()
    dt = _parse_date_for_sort(raw)
    if dt != datetime.min:
        return dt
    if re.match(r"^\d{2}/\d{2}$", raw):
        return raw
    return raw or None


def _money_fmt(cell) -> None:
    if cell.value is not None and isinstance(cell.value, (int, float)):
        cell.number_format = CURRENCY_FMT


def _date_fmt(cell) -> None:
    if isinstance(cell.value, datetime):
        cell.number_format = DATE_FMT


def _style_data(cell, *, money: bool, date: bool, alt: bool) -> None:
    cell.font = DATA_FONT
    if alt:
        cell.fill = ALT_FILL
    if money:
        _money_fmt(cell)
    if date:
        _date_fmt(cell)


def _write_headers(ws: Worksheet, cols: list[ColSpec]) -> None:
    for c, (label, _, _, _, width) in enumerate(cols, start=1):
        cell = ws.cell(row=1, column=c, value=label)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(c)].width = float(width)


def _cell_value(row: dict[str, Any], key: str) -> Any:
    if key == "Destination Postal Code":
        return row.get("Destination Postal Code") or row.get("Postal Code") or ""
    if key == "Amount Outstanding":
        v = row.get("Amount Outstanding")
        if v is None:
            v = row.get("Total Amount Outstanding")
        return v
    if key == "Subtotal" and row.get("Subtotal") is None:
        billed = row.get("Billed Amount")
        tax = row.get("Tax")
        if billed is not None and tax is not None:
            try:
                return round(float(billed) - float(tax), 2)
            except (TypeError, ValueError):
                pass
    return row.get(key)


def _write_row(
    ws: Worksheet,
    row_idx: int,
    values: list[Any],
    cols: list[ColSpec],
    *,
    alt: bool = False,
) -> None:
    for c, (label, key, money, date, _) in enumerate(cols, start=1):
        raw = values[c - 1] if c - 1 < len(values) else None
        if money and raw is not None and raw != "":
            try:
                raw = float(raw)
            except (TypeError, ValueError):
                pass
        elif date and raw is not None and raw != "":
            raw = _to_date(raw)
        cell = ws.cell(row=row_idx, column=c, value=raw if raw != "" else None)
        _style_data(cell, money=money, date=date, alt=alt)
        if label == "Explanation":
            cell.alignment = Alignment(wrap_text=True, vertical="top")


def _write_dict_row(
    ws: Worksheet,
    row_idx: int,
    data: dict[str, Any],
    cols: list[ColSpec],
    *,
    alt: bool = False,
) -> None:
    values = [_cell_value(data, key) for _, key, _, _, _ in cols]
    _write_row(ws, row_idx, values, cols, alt=alt)


def _sort_invoices(invoices: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return sorted(
        invoices,
        key=lambda r: _parse_date_for_sort(str(r.get("Invoice Date") or "")),
        reverse=True,
    )


def build_workbook_bytes(
    invoices: list[dict[str, Any]],
    shipments: list[dict[str, Any]],
    adjustments: list[dict[str, Any]],
) -> bytes:
    """Full detail workbook: one sheet named Invoices with invoice, shipment, and adjustment rows."""

    def _as_float(value: Any) -> float | None:
        if value is None or value == "":
            return None
        if isinstance(value, (int, float)):
            return float(value)
        try:
            return float(str(value).replace(",", "").strip())
        except ValueError:
            return None

    def _as_str(value: Any) -> str | None:
        if value is None:
            return None
        text = str(value)
        return text if text != "" else None

    def _build_invoice_row(invoice: dict[str, Any]) -> list[Any]:
        row = [None] * len(FULL_DETAIL_HEADERS)
        row[0] = "Invoice"
        row[1] = str(invoice.get("Invoice Number") or "").strip() or None
        row[2] = _as_str(invoice.get("Account Number"))
        row[3] = "$0.00"
        row[4] = _fmt_date_str(invoice.get("Invoice Date"))
        row[5] = _fmt_date_str(invoice.get("Due Date"))
        payment_status = _as_str(invoice.get("Payment Status")) or "Accepted"
        row[6] = "Closed "
        row[7] = payment_status
        row[8] = _as_float(invoice.get("Subtotal"))
        row[9] = _as_float(invoice.get("Tax")) or 0.0
        row[10] = _as_float(invoice.get("Government Charges"))
        row[11] = _as_float(invoice.get("Billed Amount")) or 0.0
        row[12] = _as_str(invoice.get("Type"))
        row[13] = _as_float(invoice.get("Shipping Adjustments"))
        row[14] = _as_float(invoice.get("Adjustments"))
        row[15] = _as_str(invoice.get("Explanation"))
        row[16] = _as_float(invoice.get("Incentive Savings"))
        return row

    def _build_shipment_row(invoice: dict[str, Any], shipment: dict[str, Any]) -> list[Any]:
        row = _build_invoice_row(invoice)
        row[0] = "Shipment"
        row[17] = _as_str(shipment.get("Pickup Date"))
        row[18] = _as_str(shipment.get("Order Date"))
        row[19] = _as_str(shipment.get("Tracking Number"))
        row[20] = _as_str(shipment.get("Service Type"))
        row[21] = _as_str(shipment.get("Postal Code") or shipment.get("Destination Postal Code"))
        row[22] = _as_str(shipment.get("Zone"))
        row[23] = _as_float(shipment.get("Weight (lbs)"))
        row[24] = _as_float(shipment.get("Published Charge"))
        row[25] = _as_float(shipment.get("Incentive Credit"))
        row[26] = _as_float(shipment.get("Billed Charge"))
        row[27] = _as_float(shipment.get("Fuel Surcharge"))
        row[28] = _as_float(shipment.get("Declared Value"))
        row[29] = _as_float(shipment.get("Surge Fee"))
        row[30] = _as_float(shipment.get("Add'l Handling"))
        row[31] = _as_float(shipment.get("Demand Surcharge"))
        row[32] = _as_float(shipment.get("Residential Surcharge"))
        row[33] = _as_float(shipment.get("Delivery Area Surcharge"))
        row[34] = _as_str(shipment.get("SO#"))
        row[35] = _as_str(shipment.get("PO#"))
        row[36] = _as_str(shipment.get("Sender Name"))
        row[37] = _as_str(shipment.get("Sender Company"))
        row[38] = _as_str(shipment.get("Receiver Name"))
        row[39] = _as_str(shipment.get("Receiver Company"))
        row[40] = _as_str(shipment.get("Receiver City/Province"))
        row[41] = _as_str(shipment.get("UserID"))
        return row

    def _build_adjustment_row(invoice: dict[str, Any], adjustment: dict[str, Any]) -> list[Any]:
        row = _build_invoice_row(invoice)
        row[0] = "Adjustment"
        row[42] = _as_str(adjustment.get("Pickup Date"))
        row[43] = _as_str(adjustment.get("Tracking Number"))
        row[44] = _as_str(adjustment.get("Adjustment Type"))
        row[45] = _as_float(adjustment.get("Adjustment Amount"))
        row[46] = _as_str(adjustment.get("SO#"))
        row[47] = _as_str(adjustment.get("PO#"))
        row[48] = _as_str(adjustment.get("Description / Reason") or adjustment.get("Adjustment Description"))
        return row

    def _group_by_invoice(items: list[dict[str, Any]]) -> dict[str, list[dict[str, Any]]]:
        grouped: dict[str, list[dict[str, Any]]] = {}
        for item in items:
            key = str(item.get("Invoice Number") or "").strip()
            grouped.setdefault(key, []).append(item)
        return grouped

    sorted_invoices = _sort_invoices(invoices)
    shipments_by_invoice = _group_by_invoice(shipments)
    adjustments_by_invoice = _group_by_invoice(adjustments)

    wb = Workbook()
    ws = wb.active
    ws.title = "Invoices"

    for col_idx, header in enumerate(FULL_DETAIL_HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = FULL_DETAIL_HEADER_FONT
        cell.fill = FULL_DETAIL_HEADER_FILL
        cell.alignment = FULL_DETAIL_HEADER_ALIGN
        ws.column_dimensions[get_column_letter(col_idx)].width = FULL_DETAIL_WIDTHS[col_idx - 1]

    row_idx = 2
    for invoice in sorted_invoices:
        invoice_row = _build_invoice_row(invoice)
        for col_idx, value in enumerate(invoice_row, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            if col_idx == 1:
                cell.font = FULL_DETAIL_BLACK_FONT
            if col_idx == 2:
                cell.number_format = '@'
            elif col_idx in FULL_DETAIL_CURRENCY_COLS and isinstance(value, (int, float)):
                cell.number_format = FULL_DETAIL_CURRENCY_FMT
        row_idx += 1

        inv_no = str(invoice.get("Invoice Number") or "").strip()
        for shipment in shipments_by_invoice.get(inv_no, []):
            shipment_row = _build_shipment_row(invoice, shipment)
            for col_idx, value in enumerate(shipment_row, start=1):
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                if col_idx == 1:
                    cell.font = FULL_DETAIL_BLACK_FONT
                if col_idx == 2:
                    cell.number_format = '@'
                elif col_idx in FULL_DETAIL_CURRENCY_COLS and isinstance(value, (int, float)):
                    cell.number_format = FULL_DETAIL_CURRENCY_FMT
            row_idx += 1

        for adjustment in adjustments_by_invoice.get(inv_no, []):
            adjustment_row = _build_adjustment_row(invoice, adjustment)
            for col_idx, value in enumerate(adjustment_row, start=1):
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                if col_idx == 1:
                    cell.font = FULL_DETAIL_BLACK_FONT
                if col_idx == 2:
                    cell.number_format = '@'
                elif col_idx in FULL_DETAIL_CURRENCY_COLS and isinstance(value, (int, float)):
                    cell.number_format = FULL_DETAIL_CURRENCY_FMT
            row_idx += 1

    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


def build_summary_workbook_bytes(
    invoices: list[dict[str, Any]],
) -> bytes:
    """Summary Excel — 12 columns A–L per spec."""
    # Format: _-"$"* #,##0.00_-;\-"$"* #,##0.00_-;_-"$"* "-"??_-;_-@_-
    ACCT_FMT = r'_-"$"* #,##0.00_-;\-"$"* #,##0.00_-;_-"$"* "-"??_-;_-@_-'
    CURRENCY_COLS = {7, 8, 9, 10}  # G H I J

    wb = Workbook()
    wb.calculation.fullCalcOnLoad = True
    ws = wb.active
    ws.title = "Sheet1"

    headers = [
        "Invoice Number",      # A=1
        "Account Number",      # B=2
        "Amount Due",          # C=3
        "Invoice Date",        # D=4
        "Invoice Status",      # E=5
        "Payment Status",      # F=6
        "Subtotal",            # G=7
        "Tax",                 # H=8
        "Government Charges",  # I=9
        "Billed Amount",       # J=10
        "Due Date",            # K=11
        "Type",                # L=12
    ]
    widths = [16.83, 14.83, 10.83, 17.83, 12.66, 14.16, 14.83, 11.66, 16.0, 14.83, 17.83, 15.83]

    for col_idx, label in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=label)
        cell.font = Font(bold=True, size=12)
        ws.column_dimensions[get_column_letter(col_idx)].width = widths[col_idx - 1]

    def _safe_float(val: Any) -> float:
        if val in (None, ""):
            return 0.0
        try:
            return round(float(str(val).replace(",", "").strip()), 2)
        except (TypeError, ValueError):
            return 0.0

    sorted_inv = sorted(
        invoices,
        key=lambda r: _parse_date_for_sort(str(r.get("Invoice Date") or "")),
        reverse=True,
    )

    # rows_by_account: account -> list of row indices (for explicit per-account formulas)
    rows_by_account: dict[str, list[int]] = {}

    for row_idx, inv in enumerate(sorted_inv, start=2):
        invoice_number = str(inv.get("Invoice Number") or "").strip()
        account_number = str(inv.get("Account Number") or "")
        inv_type = str(inv.get("Type") or "")
        is_import = "import" in inv_type.lower()

        tax = _safe_float(inv.get("Tax"))
        govt_charges = _safe_float(inv.get("Government Charges")) if is_import else 0.0
        billed = _safe_float(inv.get("Billed Amount"))

        rows_by_account.setdefault(account_number, []).append(row_idx)

        row_values = [
            invoice_number,                         # A=1  plain string, preserves leading zeros
            account_number,                         # B=2
            "$0.00",                                # C=3  plain string
            _fmt_date_str(inv.get("Invoice Date")),     # D=4  plain string "Month DD, YYYY"
            "Closed ",                              # E=5  always (trailing space per spec)
            "Accepted",                             # F=6  always
            f"=J{row_idx}-I{row_idx}-H{row_idx}",  # G=7  Billed - GovtChg - Tax
            tax,                                    # H=8
            govt_charges,                           # I=9
            billed,                                 # J=10
            _fmt_date_str(inv.get("Due Date")),         # K=11 plain string "Month DD, YYYY"
            inv_type,                               # L=12
        ]

        for c, value in enumerate(row_values, start=1):
            cell = ws.cell(row=row_idx, column=c, value=value)
            if c == 1:
                cell.number_format = '@'
            elif c in CURRENCY_COLS:
                cell.number_format = ACCT_FMT

    # N = index of last data row (header=1, data rows 2..N)
    N = len(sorted_inv) + 1

    bold12 = Font(bold=True, size=12)

    # Row N+2: grand total — SUM over entire data range
    grand_row = N + 2
    for c in CURRENCY_COLS:
        col_ltr = get_column_letter(c)
        cell = ws.cell(row=grand_row, column=c, value=f"=SUM({col_ltr}2:{col_ltr}{N})")
        cell.number_format = ACCT_FMT
        cell.font = bold12

    # Row N+4: column header labels for the per-account section
    hdr_row = N + 4
    for c, label in [(7, "Subtotal"), (8, "Tax"), (9, "Government charges"), (10, "Total")]:
        ws.cell(row=hdr_row, column=c, value=label).font = bold12

    # Rows N+5 onward: one row per unique account with explicit cell-reference formulas
    for i, (acct_no, row_indices) in enumerate(rows_by_account.items()):
        tr = N + 5 + i
        ws.cell(row=tr, column=6, value=acct_no).font = bold12
        for c in CURRENCY_COLS:
            col_ltr = get_column_letter(c)
            formula = "=SUM(" + "+".join(f"{col_ltr}{r}" for r in row_indices) + ")"
            cell = ws.cell(row=tr, column=c, value=formula)
            cell.number_format = ACCT_FMT
            cell.font = bold12

    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()
