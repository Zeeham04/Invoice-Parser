"""UPS Customs Brokerage Invoice (Import) parsing."""

from __future__ import annotations

import re
from typing import Any

from .base_parser import BaseInvoiceParser, ParseResult
from .ups_domestic import _f, _find_so_po


class UPSImportParser(BaseInvoiceParser):
    carrier_name = "UPS Import / Customs"

    _INV_NO = re.compile(r"Invoice\s+No\.?\s*:?\s*([0-9]{6,})", re.IGNORECASE)

    _ACCOUNT = re.compile(
        r"Account\s+No\.?\s*:?\s*([A-Z0-9]{4,12})",
        re.IGNORECASE,
    )

    _MONTH = (
        r"(?:January|February|March|April|May|June|July|August|"
        r"September|October|November|December|"
        r"Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Oct|Nov|Dec)"
    )
    _DATE_LONG = rf"({_MONTH}\s+\d{{1,2}},?\s*\d{{4}})"
    _DATE = re.compile(rf"Invoice\s+Date:\s*{_DATE_LONG}", re.IGNORECASE)
    _DUE = re.compile(rf"Date\s+Due:\s*{_DATE_LONG}", re.IGNORECASE)

    def parse(self, text: str, filename: str) -> ParseResult:
        warnings: list[str] = []
        flat = re.sub(r"\s+", " ", text)

        inv_no = ""
        m = self._INV_NO.search(flat)
        if m:
            inv_no = m.group(1).strip()

        account = ""
        m = self._ACCOUNT.search(flat)
        if m:
            account = m.group(1).strip()

        inv_date = ""
        m = self._DATE.search(flat)
        if m:
            inv_date = m.group(1).strip()

        due_date = ""
        m = self._DUE.search(flat)
        if m:
            due_date = m.group(1).strip()

        m_billed = re.search(r"Net Payable\s+CAD\s+([\d,]+\.\d{2})", flat, re.I)
        billed = _f(m_billed.group(1)) if m_billed else None
        if billed is None:
            m_billed = re.search(r"Net Payable\s*\$?\s*([\d,]+\.\d{2})", flat, re.I)
            billed = _f(m_billed.group(1)) if m_billed else None

        tax = None

        govt = None
        # Pattern 1: "Government Charges   480.36" or "Total Government Charges  CAD 480.36"
        for _pat in (
            r"Total\s+Government\s+Charges?\s+CAD\s+([\d,]+\.\d{2})",
            r"Total\s+Government\s+Charges?\s+([\d,]+\.\d{2})",
            r"Government\s+Charges?\s+([\d,]+\.\d{2})",
        ):
            m = re.search(_pat, flat, re.I)
            if m:
                v = _f(m.group(1))
                if v and v > 0:
                    govt = v
                    break
        # Pattern 2 (fallback): per-shipment line from raw text
        if govt is None:
            m = re.search(
                r"Total Customs Charges For Shipment\s+\S+\s+CAD\s+([\d,]+\.\d{2})",
                text,
                re.I,
            )
            if m:
                v = _f(m.group(1))
                if v and v > 0:
                    govt = v

        if not inv_no:
            warnings.append(f"{filename}: Could not extract Invoice Number.")
        if not account:
            warnings.append(f"{filename}: Could not extract Account Number.")
        if billed is None:
            warnings.append(f"{filename}: Could not find Net Payable.")

        invoice: dict[str, Any] = {
            "Invoice Number": inv_no,
            "Account Number": account,
            "Amount Due": billed,
            "Invoice Date": inv_date,
            "Invoice Status": "Open",
            "Payment Status": "",
            "Subtotal": billed,
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

        shipments = self._extract_shipments(text, invoice)
        adjustments: list[dict[str, Any]] = []

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
                r"(WW\s+Expedited|Worldwide\s+Expedited|Expedited|Standard)",
                line,
                re.I,
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

            for j in range(i + 1, min(i + 35, len(lines))):
                lo = lines[j].lower().strip()
                nums = re.findall(r"([\d,]+\.\d{2})", lines[j])

                if "import freight" in lo and nums:
                    row["Published Charge"] = _f(nums[0])
                elif "fuel surcharge" in lo and nums:
                    row["Fuel Surcharge"] = _f(nums[-1])
                elif "surge fee" in lo and nums:
                    row["Surge Fee"] = _f(nums[-1])
                elif "print label" in lo and nums:
                    pass
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

            rows.append(row)

        return rows
