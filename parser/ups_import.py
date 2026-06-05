"""UPS Customs Brokerage Invoice parsing — TYPE_B (Y8A864) and TYPE_C (4172AV)."""

from __future__ import annotations

import re
from typing import Any

from .base_parser import BaseInvoiceParser, ParseResult
from .ups_domestic import _f, _find_so_po, _norm_date


# ── Detection helper ───────────────────────────────────────────────────────────

def is_type_b(text: str) -> bool:
    """
    TYPE_B: CUSTOMS BROKERAGE with bare 'Government Charges' AND 'Brokerage Charges'
    summary lines (no 'Total' prefix).  These are standalone line items with a dollar
    amount — as opposed to 'Total Government Charges' which is a rolled-up footer.
    """
    has_govt = bool(
        re.search(r"^Government Charges\s+[\d,]+\.\d{2}", text, re.I | re.M)
    )
    has_brok = bool(
        re.search(r"^Brokerage Charges\s+[\d,]+\.\d{2}", text, re.I | re.M)
    )
    return has_govt and has_brok


# ── Shared header extraction ───────────────────────────────────────────────────

def _extract_import_header(text: str) -> dict[str, str]:
    """Extract invoice-level header fields common to TYPE_B and TYPE_C."""
    flat = re.sub(r"\s+", " ", text)

    # Invoice number — require a numeric value so the columnar layout (where the
    # label "Invoice No.:" is immediately followed by the next label "Control ID:")
    # is skipped and we lock onto the clean remittance line "Invoice No. 5728196240".
    inv_no = ""
    m = re.search(r"Invoice\s+No\.?\s*:?\s*\n?\s*([0-9]{6,})", text, re.I)
    if not m:
        m = re.search(r"Invoice\s+No\.?\s*:?\s*([0-9]{6,})", flat, re.I)
    if m:
        inv_no = str(m.group(1)).strip()

    # Account number — require a token that contains a digit (e.g. "4172AV"),
    # which excludes the next label word "Invoice" in the columnar layout and the
    # "Bank Account Number" remittance field.
    account = ""
    m = re.search(r"Account\s+No\.?\s*:?\s*\n?\s*([A-Z0-9]*\d[A-Z0-9]*)", text, re.I)
    if not m:
        m = re.search(r"Account\s+No\.?\s*:?\s*([A-Z0-9]*\d[A-Z0-9]*)", flat, re.I)
    if m:
        account = m.group(1).strip()

    # Invoice date
    inv_date = ""
    m = re.search(
        r"Invoice\s+Date\s*:?\s*\n?\s*([A-Za-z]+ \d{1,2},?\s*\d{4})",
        text, re.I,
    )
    if m:
        inv_date = _norm_date(m.group(1).strip())

    # Due date — prefer the clean "Invoice Due Date <date>" remittance line.
    # The bare "Date Due:" label sits directly above the value column, so a
    # leftmost combined search would wrongly grab the invoice-date value; try the
    # unambiguous "Invoice Due Date" label first, then fall back to "Date Due".
    due_date = ""
    m = re.search(
        r"Invoice\s+Due\s+Date\s*:?\s*\n?\s*([A-Za-z]+ \d{1,2},?\s*\d{4})",
        text, re.I,
    )
    if not m:
        m = re.search(
            r"Date\s+Due\s*:?\s*\n?\s*([A-Za-z]+ \d{1,2},?\s*\d{4})",
            text, re.I,
        )
    if m:
        due_date = _norm_date(m.group(1).strip())

    # Net Payable (billed amount)
    billed_str = ""
    m = re.search(r"Net\s+Payable\s+CAD\s+([\d,]+\.\d{2})", flat, re.I)
    if m:
        billed_str = m.group(1)
    if not billed_str:
        m = re.search(r"Net\s+Payable\s*\$?\s*([\d,]+\.\d{2})", flat, re.I)
        if m:
            billed_str = m.group(1)

    return {
        "inv_no": inv_no,
        "account": account,
        "inv_date": inv_date,
        "due_date": due_date,
        "billed_str": billed_str,
    }


# ── Common shipment extraction ─────────────────────────────────────────────────

def _extract_import_shipments(
    text: str,
    invoice: dict[str, Any],
    section_keyword: str = "",
) -> list[dict[str, Any]]:
    """
    Scan for lines containing UPS tracking numbers and collect per-shipment fields
    from the following lines.  Works for both TYPE_B (Import Shipment Detail) and
    TYPE_C (UPS Returns Shipment Detail).
    """
    rows: list[dict[str, Any]] = []
    lines = text.splitlines()
    seen: set[str] = set()

    inv_no = invoice.get("Invoice Number", "")
    inv_date = invoice.get("Invoice Date", "")

    for i, line in enumerate(lines):
        tn_m = re.search(r"\b(1Z[A-Z0-9]{16})\b", line, re.I)
        if not tn_m:
            continue
        tracking = tn_m.group(1)
        if tracking in seen:
            continue
        seen.add(tracking)

        date_m = re.search(r"(\d{2}/\d{2}/\d{2,4})", line)
        pickup = date_m.group(1) if date_m else ""

        zone_m = re.search(r"PKG\s+(\d{3})", line)
        zone = zone_m.group(1) if zone_m else ""

        wt_m = re.search(r"([\d.]+)\s+lbs", line)
        weight = _f(wt_m.group(1)) if wt_m else None

        svc_m = re.search(
            r"(WW\s+Expedited|Worldwide\s+Expedited|Ground|Expedited|Standard)",
            line, re.I,
        )
        service = svc_m.group(1) if svc_m else "WW Expedited"

        row: dict[str, Any] = {
            "Invoice Number": inv_no,
            "Invoice Date": inv_date,
            "Pickup Date": pickup,
            "Order Date": pickup,
            "Tracking Number": tracking,
            "Service Type": service,
            "Destination Postal Code": "",
            "Zone": zone,
            "Weight (lbs)": weight,
            "Published Charge": None,
            "Incentive Credit": None,
            "Billed Charge": None,
            "Fuel Surcharge": None,
            "Declared Value": None,
            "Surge Fee": None,
            "Add'l Handling": None,
            "Demand Surcharge": None,
            "Residential Surcharge": None,
            "Delivery Area Surcharge": None,
            "SO#": "",
            "PO#": "",
            "Sender Name": "",
            "Sender Company": "",
            "Receiver Name": "",
            "Receiver Company": "",
            "Receiver City/Province": "",
            "UserID": "",
        }

        for j in range(i + 1, min(i + 40, len(lines))):
            lo = lines[j].lower().strip()
            nums = re.findall(r"([\d,]+\.\d{2})", lines[j])

            if "import freight" in lo and nums:
                row["Published Charge"] = _f(nums[0])
            elif "fuel surcharge" in lo and nums:
                row["Fuel Surcharge"] = _f(nums[-1])
            elif "surge fee" in lo and nums:
                row["Surge Fee"] = _f(nums[-1])
            elif re.search(r"total charges for shipment", lo):
                if nums:
                    row["Billed Charge"] = _f(nums[-1])
                break
            elif re.search(r"shipped from", lo):
                if j + 1 < len(lines):
                    row["Sender Company"] = lines[j + 1].strip()[:80]
            elif re.search(r"returned to", lo):
                for k in range(j + 1, min(j + 5, len(lines))):
                    candidate = lines[k].strip()
                    if candidate and not re.search(
                        r"(Sold To|Requested By|CANADA|^\s*$)", lines[k], re.I
                    ):
                        row["Receiver Company"] = candidate[:80]
                        break
            elif re.search(r"reference no\.?\s*1", lo):
                ref_m = re.search(r"(SO#\S+|PO#\S+)", lines[j], re.I)
                if ref_m:
                    ref = ref_m.group(1).upper()
                    if ref.startswith("SO#") and not row["SO#"]:
                        row["SO#"] = ref
                    elif ref.startswith("PO#") and not row["PO#"]:
                        row["PO#"] = ref
            elif re.search(r"reference no\.?\s*2", lo):
                ref_m = re.search(r"(SO#\S+|PO#\S+)", lines[j], re.I)
                if ref_m:
                    ref = ref_m.group(1).upper()
                    if ref.startswith("SO#") and not row["SO#"]:
                        row["SO#"] = ref
                    elif ref.startswith("PO#") and not row["PO#"]:
                        row["PO#"] = ref

        rows.append(row)

    return rows


# ── TYPE_B Parser: Y8A864 format with Government Charges + Brokerage Charges ───

class UPSImportTypeBParser(BaseInvoiceParser):
    """
    TYPE_B — Customs Brokerage Y8A864.

    Distinguished by bare 'Government Charges <amount>' and
    'Brokerage Charges <amount>' summary lines (no 'Total' prefix).
    Government Charges = actual customs duties.
    Tax = 0 (duties are not GST).
    """
    carrier_name = "UPS Import / Customs (TYPE_B)"

    def parse(self, text: str, filename: str) -> ParseResult:
        warnings: list[str] = []

        hdr = _extract_import_header(text)
        inv_no = hdr["inv_no"]
        # Pad to reference format: all-digit → 15 chars; letter-containing → prepend '000000'
        if inv_no:
            if inv_no.isdigit() and len(inv_no) < 15:
                inv_no = inv_no.zfill(15)
            elif not inv_no.isdigit() and len(inv_no) < 15:
                inv_no = "000000" + inv_no
        account = hdr["account"]
        inv_date = hdr["inv_date"]
        due_date = hdr["due_date"]
        billed = _f(hdr["billed_str"])

        # Government Charges: bare line "^Government Charges <amount>".
        # pdftotext sometimes wraps the value to the next line, so try that first.
        # The ^ anchor excludes "Total Government Charges" (has "Total" before it).
        govt: float | None = None
        m = re.search(r"^Government Charges\s*\n\s*([\d,]+\.\d{2})", text, re.I | re.M)
        if not m:
            m = re.search(r"^Government Charges\s+([\d,]+\.\d{2})", text, re.I | re.M)
        if m:
            v = _f(m.group(1))
            if v is not None and v > 0:
                govt = v

        # Brokerage Charges: same two-variant match (next-line then same-line).
        brokerage: float | None = None
        m = re.search(r"^Brokerage Charges\s*\n\s*([\d,]+\.\d{2})", text, re.I | re.M)
        if not m:
            m = re.search(r"^Brokerage Charges\s+([\d,]+\.\d{2})", text, re.I | re.M)
        if m:
            v = _f(m.group(1))
            if v is not None and v > 0:
                brokerage = v

        # TYPE_B has no GST/HST; government charges are duties
        tax = 0.0

        # Subtotal = brokerage (net of duties); fallback to billed - govt
        if brokerage is not None:
            subtotal = brokerage
        elif billed is not None:
            subtotal = round(billed - (govt or 0.0), 2)
        else:
            subtotal = None

        if not inv_no:
            warnings.append(f"{filename}: Could not extract Invoice Number.")
        if not account:
            warnings.append(f"{filename}: Could not extract Account Number.")
        if billed is None:
            warnings.append(f"{filename}: Could not find Net Payable.")

        print(
            f"Parsed [Import/TYPE_B] {inv_no} | "
            f"Billed={billed} Tax={tax} GovtCharge={govt}",
            flush=True,
        )

        invoice: dict[str, Any] = {
            "Invoice Number": inv_no,
            "Account Number": account,
            "Amount Due": billed,
            "Invoice Date": inv_date,
            "Invoice Status": "",
            "Payment Status": "",
            "Subtotal": subtotal,
            "Tax": tax,
            "Government Charges": govt,
            "Billed Amount": billed,
            "Due Date": due_date,
            "Type": "Import",
            "Shipping Adjustments": None,
            "Adjustments": None,
            "Explanation": "",
            "Incentive Savings": None,
        }

        shipments = _extract_import_shipments(text, invoice, "Import Shipment Detail")
        return ParseResult(
            invoice=invoice,
            shipments=shipments,
            adjustments=[],
            warnings=warnings,
            raw_text_snippet=text[:500],
        )


# ── TYPE_C Parser: 4172AV format, shipping charges only, no duties ─────────────

class UPSImportTypeCParser(BaseInvoiceParser):
    """
    TYPE_C — Customs Brokerage 4172AV.

    No actual duties.  Government Charges is always 0.
    Tax = GST/HST brokerage charge extracted from the invoice.
    Even if 'Total Government Charges' appears, it represents GST — ignore it.
    """
    carrier_name = "UPS Import / Customs (TYPE_C)"

    def parse(self, text: str, filename: str) -> ParseResult:
        warnings: list[str] = []
        flat = re.sub(r"\s+", " ", text)

        hdr = _extract_import_header(text)
        inv_no = hdr["inv_no"]
        # Pad numeric customs invoice numbers to 15 digits to match reference format
        # (e.g. 5728196240 -> 000005728196240).
        if inv_no.isdigit() and len(inv_no) < 15:
            inv_no = inv_no.zfill(15)
        account = hdr["account"]
        inv_date = hdr["inv_date"]
        due_date = hdr["due_date"]
        billed = _f(hdr["billed_str"])

        # TYPE_C: government charges always 0
        govt = 0.0

        # Tax = GST/HST — first match of "Brokerage GST/HST <amount>" or "GST/HST <amount>"
        tax = 0.0
        m = re.search(r"(?:Brokerage\s+)?GST/HST\s+([\d,]+\.\d{2})", flat, re.I)
        if m:
            v = _f(m.group(1))
            if v is not None:
                tax = v

        # Subtotal = billed - tax
        subtotal = round(billed - tax, 2) if billed is not None else None

        if not inv_no:
            warnings.append(f"{filename}: Could not extract Invoice Number.")
        if not account:
            warnings.append(f"{filename}: Could not extract Account Number.")
        if billed is None:
            warnings.append(f"{filename}: Could not find Net Payable.")

        print(
            f"Parsed [Import/TYPE_C] {inv_no} | "
            f"Billed={billed} Tax={tax} GovtCharge={govt}",
            flush=True,
        )

        invoice: dict[str, Any] = {
            "Invoice Number": inv_no,
            "Account Number": account,
            "Amount Due": billed,
            "Invoice Date": inv_date,
            "Invoice Status": "",
            "Payment Status": "",
            "Subtotal": subtotal,
            "Tax": tax,
            "Government Charges": govt,
            "Billed Amount": billed,
            "Due Date": due_date,
            "Type": "Import",
            "Shipping Adjustments": None,
            "Adjustments": None,
            "Explanation": "",
            "Incentive Savings": None,
        }

        shipments = _extract_import_shipments(text, invoice, "UPS Returns Shipment Detail")
        return ParseResult(
            invoice=invoice,
            shipments=shipments,
            adjustments=[],
            warnings=warnings,
            raw_text_snippet=text[:500],
        )


# Backward-compatible alias (keeps any external references intact)
UPSImportParser = UPSImportTypeCParser
