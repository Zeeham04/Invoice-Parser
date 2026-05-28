"""UPS Delivery Service Invoice (Domestic/Export) parsing — TYPE_A."""

from __future__ import annotations

import re
from datetime import datetime
from typing import Any

from .base_parser import BaseInvoiceParser, ParseResult


# ── Shared helpers (imported by ups_import.py) ─────────────────────────────────

def _f(s) -> float | None:
    """Safe float; strips commas."""
    if s is None:
        return None
    try:
        return float(str(s).replace(",", ""))
    except (ValueError, TypeError):
        return None


def _norm_date(raw: str) -> str:
    """Normalize any 'Month D(D), YYYY' string to 'Month DD, YYYY' (zero-padded day)."""
    raw = (raw or "").strip()
    if not raw:
        return ""
    for fmt in ("%B %d, %Y", "%B %d %Y"):
        try:
            return datetime.strptime(raw, fmt).strftime("%B %d, %Y")
        except ValueError:
            continue
    return raw


def _find_so_po(text: str) -> tuple[str, str]:
    """Return (SO#XXXXX, PO#XXXXX) including prefix, preferring 1st/2nd ref blocks."""
    so, po = "", ""
    for block in re.findall(r"(?:1st|2nd)\s*ref:\s*([^\n]+)", text, re.IGNORECASE):
        sm = re.search(r"\b(SO#\d+)\b", block, re.IGNORECASE)
        pm = re.search(r"\b(PO#[\w\-/\.]+)", block, re.IGNORECASE)
        if sm:
            so = sm.group(1).upper()
        if pm:
            po = pm.group(1).upper()
    if not so:
        for m in re.finditer(r"\b(SO#\d+)\b", text, re.IGNORECASE):
            so = m.group(1).upper()
    if not po:
        for m in re.finditer(r"\b(PO#[\w\-/\.]+)", text, re.IGNORECASE):
            po = m.group(1).upper()
    return so, po


# ── Tax extraction ─────────────────────────────────────────────────────────────

def _extract_total_tax(text: str) -> float | None:
    """
    Find Total Taxes using last-match strategy.

    'Total Taxes <amount>' (no GST/HST qualifier after Taxes) is the grand-total
    summary line. Taking the last match handles invoices where sub-type lines
    (Total Taxes GST ..., Total Taxes HST ...) appear first — those lines don't
    match this pattern because the qualifier word follows before the number.

    Falls back to summing individual component lines, then a bare 'Tax $' pattern.
    """
    matches = []
    for m in re.finditer(r"Total\s+Taxes?\s+([\d,]+\.\d{2})", text, re.IGNORECASE):
        matches.append(_f(m.group(1)))
    if matches:
        return matches[-1]

    # Fallback 1: sum GST + HST + PST + QST sub-lines
    components: dict[str, float] = {}
    for line in text.splitlines():
        for label, key in (
            (r"Total\s+Taxes?\s+GST\b", "gst"),
            (r"Total\s+Taxes?\s+HST\b", "hst"),
            (r"Total\s+Taxes?\s+PST\b", "pst"),
            (r"Total\s+Taxes?\s+QST\b", "qst"),
        ):
            if re.search(label, line, re.IGNORECASE):
                m = re.search(r"([\d,]+\.\d{2})\s*$", line.strip())
                if m:
                    components[key] = _f(m.group(1)) or 0.0
    if components:
        return round(sum(components.values()), 2)

    # Fallback 2: bare 'Tax $X.XX'
    m = re.search(r"\bTax\s+\$\s*([\d,]+\.\d{2})", text, re.IGNORECASE)
    if m:
        return _f(m.group(1))

    return None


# ── Shipment section helpers ────────────────────────────────────────────────────

def _shipment_section_lines(lines: list[str]) -> list[tuple[int, str]]:
    """Yield (line_index, line_text) only while inside Outbound/Inbound sections."""
    in_section = False
    result = []
    for i, line in enumerate(lines):
        s = line.strip()
        if re.match(r"^(Outbound|Inbound)\s*$", s, re.I):
            in_section = True
        elif re.match(r"^(UPS\s+CampusShip|UPS\s+Returns|Worldwide\s+Service)", s, re.I):
            in_section = True
        elif re.match(r"^Adjustments\s*&\s*Other\s*Charges", s, re.I):
            in_section = False
        elif re.match(r"^Service\s+Charges\s*$", s, re.I):
            in_section = False
        elif re.match(r"^Tax\s*$", s, re.I):
            in_section = False
        if in_section:
            result.append((i, line))
    return result


_PRIMARY_LINE = re.compile(
    r"^(?:(\d{2}/\d{2})\s+)?"
    r"(1Z[A-Z0-9]{16})\s+"
    r"((?:Standard|Express\s*(?:Saver\s*)?(?:Plus\s*)?|Expedited"
    r"|WW\s*Expedited|Worldwide\s*Expedited"
    r"|Returns\s*Standard\s*Shipment"
    r"|UPS\s*CampusShip)"
    r"(?:\s+\S+)?)\s+"
    r"([A-Z]\d[A-Z]\s?\d[A-Z]\d|\d{5}(?:-\d{4})?)\s+"
    r"(\d{3})\s+"
    r"(\d+(?:\.\d+)?)\s+lbs\s+"
    r"([\d,]+\.\d{2})\s+"
    r"([-\d,]+\.\d{2})\s+"
    r"([\d,]+\.\d{2})",
    re.IGNORECASE,
)


# ── Parser ─────────────────────────────────────────────────────────────────────

class UPSDomesticParser(BaseInvoiceParser):
    carrier_name = "UPS Domestic/Export"

    def parse(self, text: str, filename: str) -> ParseResult:
        warnings: list[str] = []
        flat = re.sub(r"\s+", " ", text)

        # Invoice number — exact string from PDF, no reformatting
        inv_no = ""
        m = re.search(r"Invoice\s+Number\s+(\S+)", flat, re.I)
        if m:
            inv_no = str(m.group(1)).strip()

        # Account number
        account = ""
        m = re.search(r"^Account\s+Number\s+([A-Z0-9]{4,12})", text, re.I | re.M)
        if m:
            account = m.group(1).strip()

        # Invoice date — handle both "Invoice Date  April 07, 2026" (inline)
        # and "Invoice Date\n   April 07, 2026" (next line)
        inv_date = ""
        m = re.search(
            r"Invoice\s+Date\s*:?\s*\n?\s*([A-Za-z]+ \d{1,2},?\s*\d{4})",
            text, re.IGNORECASE,
        )
        if m:
            inv_date = _norm_date(m.group(1).strip())

        # Due date
        due_date = ""
        m = re.search(
            r"Invoice\s+Due\s+Date\s*:?\s*\n?\s*([A-Za-z]+ \d{1,2},?\s*\d{4})",
            text, re.IGNORECASE,
        )
        if m:
            due_date = _norm_date(m.group(1).strip())

        # Billed amount
        m = re.search(r"Amount due this period\s+CAD\s+([\d,]+\.\d{2})", flat, re.I)
        billed = _f(m.group(1)) if m else None
        if billed is None:
            m = re.search(r"Amount due this period\s*\$?\s*([\d,]+\.\d{2})", flat, re.I)
            billed = _f(m.group(1)) if m else None

        # Incentive savings
        m = re.search(r"Total incentive savings this period\s+\$\s*([\d,]+\.\d{2})", flat, re.I)
        if not m:
            m = re.search(r"Total\s+Incentives\s+([\d,]+\.\d{2})", flat, re.I)
        incentive = _f(m.group(1)) if m else None

        # Tax — last match of "Total Taxes <amount>"
        tax = _extract_total_tax(text)

        # Shipping adjustments
        ship_adj = None
        m = re.search(
            r"Total\s+Shipping\s+Charge\s+Corrections\s+\d+\s+Package\(s\)\s+([\d,]+\.\d{2})",
            text, re.I,
        )
        if m:
            ship_adj = _f(m.group(1))

        # Adjustments total
        adj_total = 0.0
        m = re.search(r"Total\s+Adjustments\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})", text, re.I)
        if m:
            adj_total += _f(m.group(2)) or 0.0
        m = re.search(r"Total\s+Address\s+Corrections\s+\d+\s+([\d,]+\.\d{2})", text, re.I)
        if m:
            adj_total += _f(m.group(1)) or 0.0
        adj = round(adj_total, 2) if adj_total > 0 else None

        # Explanation
        explanations: list[str] = []
        for line in text.splitlines():
            lo = line.lower()
            if "billing adjustment" in lo and re.search(r"[\d,]+\.\d{2}", line):
                amt = re.search(r"([\d,]+\.\d{2})", line)
                explanations.append(f"Billing adj ${amt.group(1)}" if amt else "Billing adj")
            elif "total address corrections" in lo:
                amt = re.search(r"([\d,]+\.\d{2})", line)
                if amt and _f(amt.group(1)):
                    explanations.append(f"Address correction ${amt.group(1)}")
            elif "total shipping charge corrections" in lo:
                pkgs = re.search(r"(\d+)\s+Package", line, re.I)
                amt = re.search(r"([\d,]+\.\d{2})\s*$", line.strip())
                if amt:
                    pkg_str = f"{pkgs.group(1)} pkg(s)" if pkgs else ""
                    explanations.append(
                        f"Shipping corrections {pkg_str} ${amt.group(1)}".strip()
                    )
        explanation = "; ".join(explanations)

        # Subtotal (used in detail sheet; summary sheet uses the formula =J-I-H)
        subtotal = (
            round(billed - (tax or 0.0), 2) if billed is not None else None
        )

        if not inv_no:
            warnings.append(f"{filename}: Could not extract Invoice Number.")
        if not account:
            warnings.append(f"{filename}: Could not extract Account Number.")
        if billed is None:
            warnings.append(f"{filename}: Could not find Amount Due.")

        invoice: dict[str, Any] = {
            "Invoice Number": inv_no,
            "Account Number": account,
            "Amount Due": billed,
            "Invoice Date": inv_date,
            "Invoice Status": "Open",
            "Payment Status": "",
            "Subtotal": subtotal,
            "Tax": tax,
            "Government Charges": 0,
            "Billed Amount": billed,
            "Due Date": due_date,
            "Type": "Domestic/Export",
            "Shipping Adjustments": ship_adj,
            "Adjustments": adj,
            "Explanation": explanation,
            "Incentive Savings": incentive,
        }

        print(
            f"Parsed [Domestic/Export] {inv_no} | "
            f"Billed={billed} Tax={tax} GovtCharge=0",
            flush=True,
        )

        shipments = self._extract_shipments(text, invoice)
        adjustments = self._extract_adjustments(text, invoice)

        return ParseResult(
            invoice=invoice,
            shipments=shipments,
            adjustments=adjustments,
            warnings=warnings,
            raw_text_snippet=text[:500],
        )

    def _extract_shipments(
        self, text: str, invoice: dict[str, Any]
    ) -> list[dict[str, Any]]:
        lines = text.splitlines()
        section = _shipment_section_lines(lines)
        results: list[dict[str, Any]] = []
        seen: set[str] = set()
        last_date = ""

        inv_no = invoice.get("Invoice Number", "")
        inv_date = invoice.get("Invoice Date", "")

        for pos, (idx, line) in enumerate(section):
            m = _PRIMARY_LINE.match(line)
            if not m:
                continue

            tracking = m.group(2)
            if tracking in seen:
                continue
            seen.add(tracking)

            pickup = m.group(1) or last_date
            if m.group(1):
                last_date = m.group(1)

            row: dict[str, Any] = {
                "Invoice Number": inv_no,
                "Invoice Date": inv_date,
                "Pickup Date": pickup,
                "Order Date": pickup,
                "Tracking Number": tracking,
                "Service Type": m.group(3).strip(),
                "Destination Postal Code": m.group(4).replace(" ", ""),
                "Zone": m.group(5),
                "Weight (lbs)": _f(m.group(6)),
                "Published Charge": _f(m.group(7)),
                "Incentive Credit": _f(m.group(8)),
                "Billed Charge": _f(m.group(9)),
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

            got_sender = False
            got_co = False
            for fwd_pos in range(pos + 1, len(section)):
                _, fwd = section[fwd_pos]

                if _PRIMARY_LINE.match(fwd):
                    break

                lo = fwd.lower().strip()

                def last_num(s: str) -> float | None:
                    nums = re.findall(r"([\d,]+\.\d{2})", s)
                    return _f(nums[-1]) if nums else None

                if lo.startswith("fuel surcharge"):
                    row["Fuel Surcharge"] = last_num(fwd)
                elif lo.startswith("declared value"):
                    row["Declared Value"] = last_num(fwd)
                elif re.match(r"surge fee", lo):
                    row["Surge Fee"] = last_num(fwd)
                elif re.match(r"additional handling", lo):
                    row["Add'l Handling"] = last_num(fwd)
                elif re.match(r"demand surcharge", lo):
                    row["Demand Surcharge"] = last_num(fwd)
                elif re.match(r"residential surcharge", lo):
                    row["Residential Surcharge"] = last_num(fwd)
                elif re.match(r"delivery area surcharge", lo):
                    row["Delivery Area Surcharge"] = last_num(fwd)
                elif "1st ref:" in lo or "2nd ref:" in lo:
                    so_m = re.search(r"\b(SO#\d+)\b", fwd, re.I)
                    po_m = re.search(r"\b(PO#[\w\-/\.]+)", fwd, re.I)
                    if so_m and not row["SO#"]:
                        row["SO#"] = so_m.group(1).upper()
                    if po_m and not row["PO#"]:
                        row["PO#"] = po_m.group(1).upper()
                elif lo.startswith("userid"):
                    u = re.search(r"UserID\s*:?\s*(\S+)", fwd, re.I)
                    if u:
                        row["UserID"] = u.group(1)
                elif re.search(r"\bSender\b", fwd, re.I) and re.search(
                    r"\bReceiver\b", fwd, re.I
                ):
                    snd_m = re.search(
                        r"Sender\s*:?\s*(.+?)(?:\s{2,}|\bReceiver\b)", fwd, re.I
                    )
                    rcv_m = re.search(r"Receiver\s*:?\s*(.+)$", fwd, re.I)
                    if snd_m:
                        row["Sender Name"] = snd_m.group(1).strip()[:80]
                    if rcv_m:
                        row["Receiver Name"] = rcv_m.group(1).strip()[:80]
                    got_sender = True
                elif got_sender and not got_co:
                    parts = re.split(r"\s{3,}", fwd.strip())
                    if len(parts) >= 2:
                        row["Sender Company"] = parts[0].strip()[:80]
                        row["Receiver Company"] = parts[-1].strip()[:80]
                    elif len(parts) == 1 and parts[0]:
                        row["Receiver Company"] = parts[0].strip()[:80]
                    got_co = True
                elif got_co and not row["Receiver City/Province"]:
                    cp_m = re.search(
                        r"([A-Z][A-Za-zÀ-ÿ\s\-]+(?:[A-Z]{2})\s+[A-Z]\d[A-Z]\s?\d[A-Z]\d)",
                        fwd,
                    )
                    if cp_m:
                        row["Receiver City/Province"] = cp_m.group(1).strip()[:60]

            results.append(row)

        return results

    def _extract_adjustments(
        self, text: str, invoice: dict[str, Any]
    ) -> list[dict[str, Any]]:
        inv_no = invoice.get("Invoice Number", "")
        rows: list[dict[str, Any]] = []
        lines = text.splitlines()

        in_adj_section = False
        current_tracking = ""
        current_pickup = ""
        current_so = ""
        current_po = ""

        for i, line in enumerate(lines):
            lo = line.lower().strip()

            if re.match(r"adjustments\s*&\s*other\s*charges", lo, re.I):
                in_adj_section = True
                continue
            if re.match(r"(service charges|^tax)\s*$", lo, re.I):
                in_adj_section = False

            if not in_adj_section:
                continue

            if "billing adjustment" in lo and re.search(r"[\d,]+\.\d{2}", line):
                amt = re.search(r"([\d,]+\.\d{2})\s*$", line.strip())
                rows.append(
                    {
                        "Invoice Number": inv_no,
                        "Pickup Date": "",
                        "Tracking Number": "",
                        "Adjustment Type": "Billing Adjustment",
                        "Adjustment Amount": _f(amt.group(1)) if amt else None,
                        "Description / Reason": line.strip()[:500],
                        "SO#": "",
                        "PO#": "",
                    }
                )

            elif re.match(r"1z[a-z0-9]{16}", lo.split()[0] if lo.split() else ""):
                tn = line.strip().split()[0]
                if "1st ref:" not in lo and "2nd ref:" not in lo:
                    window = "\n".join(lines[i: i + 6])
                    so, po = _find_so_po(window)
                    amt = re.search(r"([\d,]+\.\d{2})\s*$", line.strip())
                    if lo.endswith("standard") or amt:
                        rows.append(
                            {
                                "Invoice Number": inv_no,
                                "Pickup Date": "",
                                "Tracking Number": tn,
                                "Adjustment Type": "Address Correction",
                                "Adjustment Amount": _f(amt.group(1)) if amt else None,
                                "Description / Reason": "Address Correction",
                                "SO#": so,
                                "PO#": po,
                            }
                        )

            elif re.match(r"\d{2}/\d{2}\s+1z", lo):
                date_m = re.match(r"(\d{2}/\d{2})", line)
                tn_m = re.search(r"(1Z[A-Z0-9]{16})", line, re.I)
                current_pickup = date_m.group(1) if date_m else ""
                current_tracking = tn_m.group(1) if tn_m else ""
                window = "\n".join(lines[i: i + 8])
                current_so, current_po = _find_so_po(window)

            elif "total shipping charge corrections" in lo:
                pkgs_m = re.search(r"(\d+)\s+Package", line, re.I)
                amt_m = re.search(r"([\d,]+\.\d{2})\s*$", line.strip())
                rows.append(
                    {
                        "Invoice Number": inv_no,
                        "Pickup Date": current_pickup,
                        "Tracking Number": current_tracking,
                        "Adjustment Type": "Shipping Charge Correction",
                        "Adjustment Amount": _f(amt_m.group(1)) if amt_m else None,
                        "Description / Reason": (
                            f"Shipping Charge Corrections "
                            f"{pkgs_m.group(1) if pkgs_m else ''} Package(s)"
                        ).strip(),
                        "SO#": current_so,
                        "PO#": current_po,
                    }
                )

        return rows
