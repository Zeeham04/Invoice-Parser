"""PDF text extraction and invoice routing."""

from __future__ import annotations

import io
import logging
import re
from pathlib import Path
from typing import Any

import pdfplumber

from .base_parser import ParseResult
from .ups_domestic import UPSDomesticParser
from .ups_import import UPSImportTypeBParser, UPSImportTypeCParser, is_type_b

logger = logging.getLogger(__name__)

MAX_PAGES = 50
MIN_TEXT_CHARS = 40


def extract_text_from_pdf(data: bytes, filename: str = "") -> tuple[str, str | None]:
    """
    Extract text using pdfplumber; fallback to pdfminer.six; final fallback to OCR.
    Returns (text, error_message_if_image_only).
    """
    err: str | None = None
    chunks: list[str] = []

    try:
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            n = min(len(pdf.pages), MAX_PAGES)
            for i in range(n):
                page = pdf.pages[i]
                t = page.extract_text()
                if t:
                    chunks.append(t)
    except Exception as e:
        logger.warning("pdfplumber failed for %s: %s", filename, e)
        print(f"pdfplumber failed for {filename}: {e}", flush=True)

    text = "\n".join(chunks)

    if len(text.strip()) < MIN_TEXT_CHARS:
        try:
            from pdfminer.high_level import extract_text as pdfminer_extract

            alt = pdfminer_extract(io.BytesIO(data), maxpages=MAX_PAGES) or ""
            if len(alt.strip()) > len(text.strip()):
                text = alt
        except Exception as e:
            logger.warning("pdfminer fallback failed for %s: %s", filename, e)
            print(f"pdfminer fallback failed for {filename}: {e}", flush=True)

    if len(text.strip()) < 50:
        try:
            from pdf2image import convert_from_bytes
            import pytesseract

            print(f"OCR fallback triggered for {filename}", flush=True)
            images = convert_from_bytes(data, last_page=MAX_PAGES)
            ocr_chunks: list[str] = []
            for img in images:
                page_text = pytesseract.image_to_string(img)
                if page_text:
                    ocr_chunks.append(page_text)
            ocr_text = "\n".join(ocr_chunks)
            if len(ocr_text.strip()) > len(text.strip()):
                text = ocr_text
        except Exception as e:
            logger.warning("OCR fallback failed for %s: %s", filename, e)
            print(f"OCR fallback failed for {filename}: {e}", flush=True)

    if len(text.strip()) < MIN_TEXT_CHARS:
        err = (
            f"{filename or 'PDF'} has no selectable text and OCR could not extract content. "
            "The file may be a scanned image that is not machine-readable."
        )

    return text, err


def detect_invoice_type(text: str) -> str:
    """
    Return one of:
      'type_a' — UPS Delivery Service Invoice (Domestic/Export)
      'type_b' — UPS Customs Brokerage with Government Charges + Brokerage Charges lines
      'type_c' — UPS Customs Brokerage, shipping charges only (no duties)
    """
    upper = text.upper()
    if "DELIVERY SERVICE INVOICE" in upper:
        return "type_a"
    if "CUSTOMS BROKERAGE" in upper:
        return "type_b" if is_type_b(text) else "type_c"
    # Heuristic fallback
    if "NET PAYABLE" in upper and "BROKERAGE" in upper:
        return "type_b" if is_type_b(text) else "type_c"
    return "type_a"


_SPLIT_PATTERNS = (
    re.compile(r"(?m)(^\s*DELIVERY SERVICE INVOICE\b)", re.I),
    re.compile(r"(?m)(^\s*CUSTOMS BROKERAGE INVOICE\b)", re.I),
)


def split_text_into_invoice_segments(text: str) -> list[str]:
    """
    When a single PDF contains multiple invoices, UPS repeats the invoice title line.
    Split on those markers so each segment is parsed independently.
    """
    best_matches: list[re.Match[str]] | None = None
    for pat in _SPLIT_PATTERNS:
        ms = list(pat.finditer(text))
        if best_matches is None or len(ms) > len(best_matches):
            best_matches = ms
    if not best_matches or len(best_matches) <= 1:
        return [text]
    starts = [m.start(1) for m in best_matches]
    chunks: list[str] = []
    for i in range(len(starts)):
        seg_start = 0 if i == 0 else starts[i]
        seg_end = starts[i + 1] if i + 1 < len(starts) else len(text)
        chunks.append(text[seg_start:seg_end])
    return chunks


def _make_parser(kind: str):
    if kind == "type_b":
        return UPSImportTypeBParser()
    if kind == "type_c":
        return UPSImportTypeCParser()
    return UPSDomesticParser()


def parse_pdf_bytes(data: bytes, filename: str) -> dict[str, Any]:
    """Parse a PDF from bytes; returns JSON-serializable dict."""
    text, extract_err = extract_text_from_pdf(data, filename)
    warnings: list[str] = []
    if extract_err:
        warnings.append(extract_err)
        return {
            "ok": False,
            "filename": filename,
            "invoice": {},
            "invoices": [],
            "shipments": [],
            "adjustments": [],
            "warnings": warnings,
            "detected_type": "",
        }

    segments = split_text_into_invoice_segments(text)
    invoices_out: list[dict[str, Any]] = []
    shipments_out: list[dict[str, Any]] = []
    adjustments_out: list[dict[str, Any]] = []
    kinds_seen: list[str] = []

    for seg in segments:
        kind = detect_invoice_type(seg)
        kinds_seen.append(kind)
        parser = _make_parser(kind)
        try:
            result: ParseResult = parser.parse(seg, filename)
        except Exception as e:
            logger.exception("Parse error %s", filename)
            warnings.append(f"{filename}: Parse error — {e}")
            return {
                "ok": False,
                "filename": filename,
                "invoice": {},
                "invoices": [],
                "shipments": [],
                "adjustments": [],
                "warnings": warnings,
                "detected_type": kind,
            }

        warnings.extend(result.warnings)
        invoices_out.append(result.invoice)
        shipments_out.extend(result.shipments)
        adjustments_out.extend(result.adjustments)

    unique_kinds = set(kinds_seen)
    if len(unique_kinds) > 1:
        detected_label = "Mixed"
    elif kinds_seen:
        k = kinds_seen[0]
        detected_label = "Domestic/Export" if k == "type_a" else "Import"
    else:
        detected_label = ""

    first_inv = invoices_out[0] if invoices_out else {}
    return {
        "ok": True,
        "filename": filename,
        "invoices": invoices_out,
        "invoice": first_inv,
        "shipments": shipments_out,
        "adjustments": adjustments_out,
        "warnings": warnings,
        "detected_type": detected_label,
    }


def parse_pdf_path(path: str | Path) -> dict[str, Any]:
    p = Path(path)
    data = p.read_bytes()
    return parse_pdf_bytes(data, p.name)
