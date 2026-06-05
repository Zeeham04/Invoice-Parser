"""
parser/ups_parser.py

Universal UPS invoice parser. Handles all three UPS Canada invoice formats:
  - Delivery Service Invoice (domestic/export)   → 'delivery'
  - Customs Brokerage Invoice, account 4172AV    → 'brokerage_4172av'
  - Customs Brokerage Invoice, account Y8A864    → 'brokerage_y8a864'

Usage:
    from parser.ups_parser import parse_invoice
    result = parse_invoice(pdf_bytes_or_path)
"""

from __future__ import annotations

import io
import re
from pathlib import Path
from typing import Union

import pdfplumber


class UPSParseError(Exception):
    pass


# ── Month constants ───────────────────────────────────────────────────────────

_MONTHS_FULL = (
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
)
_MONTHS_ABBR = tuple(m[:3] for m in _MONTHS_FULL)
_MONTH_PATTERN = "(?:" + "|".join(_MONTHS_FULL) + "|" + "|".join(_MONTHS_ABBR) + ")"

# BUG-05 fix: _DATE_LONG matches "January 10, 2026" or "January 10 2026" (with/without comma)
_DATE_LONG  = rf"{_MONTH_PATTERN}\.?[ \t]+\d{{1,2}},?[ \t]*\d{{4}}"
_DATE_SHORT = r"\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4}"
_DATE_ANY   = rf"(?:{_DATE_LONG}|{_DATE_SHORT})"


# ── Pre-compiled patterns ─────────────────────────────────────────────────────

_RE_INV_NUM_DELIVERY  = re.compile(r"Invoice\s+Number\s+([A-Z0-9]{4,})", re.I)
_RE_INV_NUM_BROKERAGE = re.compile(r"Invoice\s+No\.?\s*:?\s*([0-9]{4,})", re.I)

# BUG-06 fix: use [A-Z0-9]{4,} (not \d{4,}) to match alphanumeric accounts like Y8A864, 4172AV
# BUG-05/06 fix: \n?\s* handles values on the next line after the label
_RE_ACCT_DELIVERY  = re.compile(r"Account\s+Number\s*:?\s*\n?\s*([A-Z0-9]{4,})", re.I)
_RE_ACCT_BROKERAGE = re.compile(r"Account\s+No\.?\s*:?\s*\n?\s*([A-Z0-9]{4,})", re.I)

# BUG-05 fix: add \s*:?\s* so the optional colon after "Date" doesn't break the match.
# [ \t]*\n?[ \t]* lets the value sit on the same line OR the next line.
_RE_DATE_INVOICE_D = re.compile(
    rf"Invoice\s+Date\s*:?[ \t]*\n?[ \t]*({_DATE_ANY})", re.I
)
_RE_DATE_INVOICE_B = re.compile(
    rf"Invoice\s+Date\s*:?[ \t]*\n?[ \t]*({_DATE_ANY})", re.I
)
_RE_DATE_DUE_D = re.compile(
    rf"Invoice\s+Due\s+Date\s*:?[ \t]*\n?[ \t]*({_DATE_ANY})", re.I
)
_RE_DATE_DUE_B = re.compile(
    rf"Date\s+Due\s*:?[ \t]*\n?[ \t]*({_DATE_ANY})", re.I
)

_RE_AMOUNT_DUE  = re.compile(r"Amount\s+due\s+this\s+period\s+CAD\s+([\d,]+\.\d{2})", re.I)
_RE_NET_PAYABLE = re.compile(r"Net\s+Payable\s+CAD\s+([\d,]+\.\d{2})", re.I)
_RE_INCENTIVE   = re.compile(
    r"Total\s+incentive\s+savings\s+this\s+period\s+\$?\s*([\d,]+\.\d{2})", re.I
)

# BUG-01 fix: use [ \t]+ instead of \s+ so the amount MUST be on the same line.
# \s+ would match newlines, allowing a GST sub-line like "Total Taxes GST\n21.32"
# to be split across lines and matched. [ \t]+ prevents this.
# Negative lookahead (?!GST|HST|PST|QST|R\d) rejects lines like "Total Taxes GST …"
# even if whitespace handling changed.
_RE_TAX_TOTAL = re.compile(
    r"^Total[ \t]+Taxes[ \t]+(?!GST|HST|PST|QST|R\d)([\d,]+\.\d{2})[ \t]*$",
    re.I | re.M,
)
_RE_TAX_GST = re.compile(r"Total\s+Taxes\s+GST\s+\S+\s+([\d,]+\.\d{2})", re.I)
_RE_TAX_HST = re.compile(r"Total\s+Taxes\s+HST\s+\S+\s+([\d,]+\.\d{2})", re.I)
_RE_DISCOUNTS = re.compile(r"^Discounts[ \t]+([\d,]+\.\d{2})", re.I | re.M)

# BUG-09 fix: make the $ sign optional (\$?) so the pattern still matches if
# pdfplumber omits the dollar sign when extracting the summary line.
_RE_WORLDWIDE  = re.compile(r"Worldwide\s+Service\s+\$?\s*([\d,]+\.\d{2})", re.I)
_RE_CAMPUSSHIP = re.compile(r"UPS\s+CampusShip\s+\$?\s*([\d,]+\.\d{2})", re.I)
_RE_RETURNS_D  = re.compile(r"UPS\s+Returns\s+\$?\s*([\d,]+\.\d{2})", re.I)
_RE_ADJUSTMENTS = re.compile(
    r"Adjustments\s+(?:&|and|&amp;)\s+Other\s+Charges\s+\$?\s*([\d,]+\.\d{2})", re.I
)

# 4172AV brokerage charge lines (extracted from Summary section only — BUG-08)
_RE_IMPORT_FREIGHT  = re.compile(r"^Import[ \t]+Freight[ \t]+([\d,]+\.\d{2})", re.I | re.M)
_RE_FUEL_SURCHARGE  = re.compile(r"^Fuel[ \t]+Surcharge[ \t]+([\d,]+\.\d{2})", re.I | re.M)
_RE_PRINT_LABEL     = re.compile(r"^Print[ \t]+Label[ \t]+([\d,]+\.\d{2})", re.I | re.M)
_RE_SURGE_FEE       = re.compile(r"Surge\s+Fee\s*[-–]\s*Com\s+([\d,]+\.\d{2})", re.I)
_RE_GOVT_AGENCY_FEE = re.compile(r"Government\s+Agency\s+Fee\s+([\d,]+\.\d{2})", re.I)
_RE_TARIFF_LINE_FEE = re.compile(r"Additional\s+Tariff\s+Line\s+Fee\s+([\d,]+\.\d{2})", re.I)
_RE_BROKERAGE_GST   = re.compile(r"Brokerage\s+GST[/\\]?HST\s+([\d,]+\.\d{2})", re.I)

# Y8A864 brokerage charge lines (extracted from UPS charge table only — BUG-11)
_RE_PGA_DISCLAIM = re.compile(r"PGA\s+Disclaim\s+Fee\s+([\d,]+\.\d{2})", re.I)
_RE_ENTRY_PREP   = re.compile(r"Entry\s+Prep\s+Fee\s+([\d,]+\.\d{2})", re.I)
_RE_DISBURSEMENT = re.compile(r"Disbursement\s+Fee\s+([\d,]+\.\d{2})", re.I)
_RE_DUTY         = re.compile(r"^Duty[ \t]+([\d,]+\.\d{2})", re.I | re.M)
_RE_MERCH_PROC   = re.compile(r"Merchandise\s+Processing\s+Fee\s+([\d,]+\.\d{2})", re.I)


# ── Section isolation helpers (BUG-08, BUG-11) ───────────────────────────────

def _summary_section(text: str) -> str:
    """
    BUG-08: For 4172AV invoices, return only the Summary of Charges section
    (page 1).  Page 2 contains the per-shipment detail table which repeats
    the same charge labels with different per-package values; matching against
    the full text would pick up those page-2 values instead of the totals.
    """
    lo = text.lower()
    start = lo.find("summary of charges")
    if start == -1:
        return text  # fallback: full text
    for end_marker in (
        "ups returns shipment detail",
        "import shipment detail",
        "discount summary",
        "symbol explanation",
        "page 2",
    ):
        idx = lo.find(end_marker, start + 20)
        if idx != -1:
            return text[start:idx]
    return text[start:]


def _strip_customs_cad(text: str) -> str:
    """
    BUG-11: Remove the 'Customs CAD Calculations' block from Y8A864 invoices.
    That block lists CBSA tariff-class breakdown rows (e.g. a $2.00 CBSA charge
    on invoice 5732658528) which are NOT UPS fees and must not land in any column.
    We keep everything before the block and resume at the UPS charge table.
    """
    lo = text.lower()
    idx = lo.find("customs cad calculations")
    if idx == -1:
        return text
    # Resume at the UPS per-shipment charge table heading
    for end_marker in (
        "total charges for shipment",
        "entry prep fee",
        "disbursement fee",
        "pga disclaim",
        "brokerage charges",
    ):
        end_idx = lo.find(end_marker, idx)
        if end_idx != -1:
            return text[:idx] + "\n" + text[end_idx:]
    return text[:idx]


# ── Low-level helpers ─────────────────────────────────────────────────────────

def _extract_text(source: Union[str, bytes, Path]) -> str:
    """Extract all text from all PDF pages. Accepts file path, bytes, or Path."""
    if isinstance(source, (bytes, bytearray)):
        source = io.BytesIO(source)
    pages = []
    with pdfplumber.open(source) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                pages.append(t)
    if not pages:
        raise UPSParseError("No text could be extracted from the PDF")
    return "\n".join(pages)


def _find(pattern: re.Pattern, text: str, default: str = "") -> str:
    m = pattern.search(text)
    return m.group(1).strip() if m else default


def _find_float(pattern: re.Pattern, text: str, default: float = 0.0) -> float:
    m = pattern.search(text)
    if not m:
        return default
    try:
        return float(m.group(1).replace(",", ""))
    except (ValueError, IndexError):
        return default


def _find_multiline_float(label: str, text: str, default: float = 0.0) -> float:
    """
    Extract a float appearing on the same line as label OR on the next line.
    pdfplumber sometimes wraps charge amounts onto the next line.
    The ^ anchor ensures we don't match 'Total Government Charges' (has 'Total' prefix).
    """
    escaped = re.escape(label)
    m = re.search(rf"^{escaped}[ \t]*\n[ \t]*([\d,]+\.\d{{2}})", text, re.I | re.M)
    if not m:
        m = re.search(rf"^{escaped}[ \t]+([\d,]+\.\d{{2}})", text, re.I | re.M)
    if not m:
        return default
    try:
        return float(m.group(1).replace(",", ""))
    except (ValueError, IndexError):
        return default


def _normalize_date(raw: str) -> str:
    """Normalize any UPS date to 'Month D, YYYY'. Returns raw if unparseable."""
    if not raw:
        return ""
    raw = raw.strip()
    # Long form: "January 10, 2026" or "January 10 2026"
    m = re.match(rf"({_MONTH_PATTERN})\.?[ \t]+(\d{{1,2}}),?[ \t]*(\d{{4}})", raw, re.I)
    if m:
        month_raw, day, year = m.group(1), int(m.group(2)), m.group(3)
        for full in _MONTHS_FULL:
            if full.lower().startswith(month_raw.lower()[:3]):
                month_raw = full
                break
        return f"{month_raw} {day}, {year}"
    # Numeric form: 01/10/2026
    m2 = re.match(r"(\d{1,2})[/\-](\d{1,2})[/\-](\d{2,4})", raw)
    if m2:
        month_num, day, year = int(m2.group(1)), int(m2.group(2)), m2.group(3)
        if len(year) == 2:
            year = "20" + year
        if 1 <= month_num <= 12:
            return f"{_MONTHS_FULL[month_num - 1]} {day}, {year}"
    return raw


def _extract_tax(text: str) -> float:
    """
    BUG-01: Extract the grand-total tax from a delivery invoice.

    The tax section has three kinds of lines:
      Total Taxes GST R105453328    21.32   ← sub-entry, must be skipped
      Total Taxes HST R105453328     9.06   ← sub-entry, must be skipped
      Total Taxes                   30.38   ← grand-total summary line, take THIS

    _RE_TAX_TOTAL uses [ \t]+ (not \s+) so it cannot match across a newline,
    and a negative lookahead rejects lines that have GST/HST/etc. after 'Taxes'.
    Taking the last match handles multi-account invoices where the summary line
    appears once per account before a final grand total.

    Falls back to summing GST + HST individual components only when no summary
    line exists at all (some invoice layouts omit it).
    """
    summary_matches = _RE_TAX_TOTAL.findall(text)
    if summary_matches:
        try:
            return float(summary_matches[-1].replace(",", ""))
        except ValueError:
            pass
    # Fallback: sum sub-entries (only used when no summary line exists)
    gst = _find_float(_RE_TAX_GST, text)
    hst = _find_float(_RE_TAX_HST, text)
    return round(gst + hst, 2)


def _detect_format(text: str) -> str:
    """
    Detect invoice format from PDF text content only.
    Returns: 'delivery', 'brokerage_4172av', or 'brokerage_y8a864'
    """
    upper = text.upper()

    if "DELIVERY SERVICE INVOICE" in upper:
        return "delivery"

    if "CUSTOMS BROKERAGE INVOICE" in upper:
        # BUG-06 fix: use the corrected _RE_ACCT_BROKERAGE which allows [A-Z0-9]
        acct_m = _RE_ACCT_BROKERAGE.search(text)
        acct = acct_m.group(1).upper() if acct_m else ""
        if acct == "4172AV":
            return "brokerage_4172av"
        if _RE_DUTY.search(text):
            return "brokerage_y8a864"
        if acct:
            return "brokerage_y8a864"
        return "brokerage_4172av"

    if "UPS CAMPUSSHIP" in upper:
        return "delivery"

    return "delivery"


# ── Main parse function ───────────────────────────────────────────────────────

def parse_invoice(source: Union[str, bytes, Path], filename: str = "") -> dict:
    """
    Parse any UPS invoice PDF and return a structured dict.

    source  — file path (str / Path), raw bytes, or BytesIO.
    filename — used in error messages when source is bytes.

    Raises UPSParseError if the PDF cannot be parsed.
    """
    src_label = filename or (str(source) if isinstance(source, (str, Path)) else "PDF")
    text = _extract_text(source)
    fmt  = _detect_format(text)

    result: dict = {
        "invoice_number":     "",
        "account_number":     "",
        "invoice_date":       "",
        "due_date":           "",
        "invoice_type":       fmt,
        "amount_due":         "$0.00",
        "invoice_status":     "Closed",
        "payment_status":     "Accepted",
        "tax":                0.0,
        "government_charges": 0.0,
        "billed_amount":      0.0,
        "invoice_type_label": "",
        "import_freight":     0.0,
        "fuel_surcharge":     0.0,
        "print_label":        0.0,
        "surge_fees":         0.0,
        "discounts_applied":  0.0,
        "worldwide_service":  0.0,
        "campus_ship":        0.0,
        "ups_returns":        0.0,
        "adjustments_other":  0.0,
        "service_charges":    0.0,
        "govt_agency_fee":    0.0,
        "tariff_line_fee":    0.0,
        "brokerage_gst":      0.0,
        "duty_us":            0.0,
        "merch_proc_fee":     0.0,
        "disbursement_fee":   0.0,
        "entry_prep_fee":     0.0,
        "pga_disclaim_fee":   0.0,
    }

    if fmt == "delivery":
        result["invoice_number"]     = _find(_RE_INV_NUM_DELIVERY, text)
        result["account_number"]     = _find(_RE_ACCT_DELIVERY, text)
        result["invoice_date"]       = _normalize_date(_find(_RE_DATE_INVOICE_D, text))
        result["due_date"]           = _normalize_date(_find(_RE_DATE_DUE_D, text))
        result["billed_amount"]      = _find_float(_RE_AMOUNT_DUE, text)
        result["tax"]                = _extract_tax(text)
        result["government_charges"] = 0.0  # delivery invoices never carry customs duties
        result["invoice_type_label"] = "Domestic/Export"

        incentive = _find_float(_RE_INCENTIVE, text)
        # BUG-07: store discounts as negative
        result["discounts_applied"]  = -incentive if incentive else 0.0
        result["worldwide_service"]  = _find_float(_RE_WORLDWIDE, text)
        result["campus_ship"]        = _find_float(_RE_CAMPUSSHIP, text)
        result["ups_returns"]        = _find_float(_RE_RETURNS_D, text)
        result["adjustments_other"]  = _find_float(_RE_ADJUSTMENTS, text)
        result["service_charges"]    = 0.0

    elif fmt == "brokerage_4172av":
        result["invoice_number"]     = _find(_RE_INV_NUM_BROKERAGE, text)
        result["account_number"]     = _find(_RE_ACCT_BROKERAGE, text)
        result["invoice_date"]       = _normalize_date(_find(_RE_DATE_INVOICE_B, text))
        result["due_date"]           = _normalize_date(_find(_RE_DATE_DUE_B, text))
        result["billed_amount"]      = _find_float(_RE_NET_PAYABLE, text)
        result["tax"]                = 0.0
        # BUG-02: government_charges is ALWAYS 0 for 4172AV.
        # The PDF shows "Total Government Charges X.XX" but this is GST on
        # brokerage fees, not import duties. It belongs in brokerage_gst only.
        result["government_charges"] = 0.0
        result["invoice_type_label"] = "Import"

        # BUG-08 / BUG-11: extract charges from Summary of Charges section only.
        # Page 2 repeats the same labels for per-shipment detail; matching the
        # full text picks up those values instead of the page-1 totals.
        summary = _summary_section(text)
        discounts = _find_float(_RE_DISCOUNTS, summary)
        # BUG-07: discounts are positive in the PDF, store as negative
        result["discounts_applied"]  = -discounts if discounts else 0.0
        result["import_freight"]     = _find_float(_RE_IMPORT_FREIGHT, summary)
        result["fuel_surcharge"]     = _find_float(_RE_FUEL_SURCHARGE, summary)
        result["print_label"]        = _find_float(_RE_PRINT_LABEL, summary)
        result["surge_fees"]         = _find_float(_RE_SURGE_FEE, summary)
        result["govt_agency_fee"]    = _find_float(_RE_GOVT_AGENCY_FEE, summary)
        result["tariff_line_fee"]    = _find_float(_RE_TARIFF_LINE_FEE, summary)
        result["brokerage_gst"]      = _find_float(_RE_BROKERAGE_GST, summary)

    elif fmt == "brokerage_y8a864":
        result["invoice_number"]     = _find(_RE_INV_NUM_BROKERAGE, text)
        result["account_number"]     = _find(_RE_ACCT_BROKERAGE, text)
        result["invoice_date"]       = _normalize_date(_find(_RE_DATE_INVOICE_B, text))
        result["due_date"]           = _normalize_date(_find(_RE_DATE_DUE_B, text))
        result["billed_amount"]      = _find_float(_RE_NET_PAYABLE, text)
        result["tax"]                = 0.0
        result["invoice_type_label"] = "Import"
        result["discounts_applied"]  = 0.0

        # Government Charges: real customs duty, value may wrap to next line (BUG from prev sessions)
        result["government_charges"] = _find_multiline_float("Government Charges", text)

        # BUG-11: strip the Customs CAD Calculations block before extracting UPS fees.
        # That block contains CBSA tariff breakdown rows (e.g. $2.00 CBSA charge on
        # invoice 5732658528) which are not UPS charges and must not land in any column.
        ups_text = _strip_customs_cad(text)
        result["tariff_line_fee"]    = _find_float(_RE_TARIFF_LINE_FEE, ups_text)
        result["duty_us"]            = _find_float(_RE_DUTY, ups_text)
        result["merch_proc_fee"]     = _find_float(_RE_MERCH_PROC, ups_text)
        result["disbursement_fee"]   = _find_float(_RE_DISBURSEMENT, ups_text)
        result["entry_prep_fee"]     = _find_float(_RE_ENTRY_PREP, ups_text)
        result["pga_disclaim_fee"]   = _find_float(_RE_PGA_DISCLAIM, ups_text)

    # BUG-03: invoice_number must always be a plain string — never cast to int.
    result["invoice_number"] = str(result["invoice_number"])

    # Validate required fields
    if not result["invoice_number"]:
        raise UPSParseError(f"Could not extract invoice number from {src_label}")
    if not result["account_number"]:
        raise UPSParseError(f"Could not extract account number from {src_label}")
    if result["billed_amount"] == 0.0:
        raise UPSParseError(f"Could not extract billed amount from {src_label}")

    print(
        f"Parsed [{fmt}] {result['invoice_number']} | "
        f"Billed={result['billed_amount']} "
        f"Tax={result['tax']} "
        f"GovtCharge={result['government_charges']}",
        flush=True,
    )

    return result
