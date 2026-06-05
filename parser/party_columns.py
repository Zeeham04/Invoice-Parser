"""
Coordinate-based Sender/Receiver column extraction for domestic UPS invoices.

pdfplumber's ``extract_text`` joins the two side-by-side address columns
(Sender on the left, Receiver on the right) with single spaces, e.g.::

    Sender :Zach Anderson Receiver:ZACK
    U TECHNOLOGY CORPORATION ELECTROCAN POWER SERVICES INC.
    5321 11ST NE 37 KODIAK CR
    CALGARY AB T2E8N4 NORTH YORK ON M3J3E5

That cannot be split reliably from the flattened text, so this module reopens
the PDF and uses the x-coordinate of the "Receiver" label as the column boundary
to separate the left (sender) and right (receiver) words on every line of the
block. Results are keyed by the shipment's tracking number so callers can
enrich already-parsed shipment rows.
"""

from __future__ import annotations

import io
import logging
import re
from typing import Any

import pdfplumber

from .ups_domestic import _PARTY_STOP_RE, _extract_city

logger = logging.getLogger(__name__)

MAX_PAGES = 50
_TRACK_RE = re.compile(r"\b(1Z[A-Z0-9]{16})\b")
_LINE_TOL = 3.0  # px tolerance for grouping words onto the same line


def _group_lines(words: list[dict[str, Any]]) -> list[list[dict[str, Any]]]:
    """Group extracted words into visual lines, each sorted left-to-right."""
    ordered = sorted(words, key=lambda w: (round(w["top"], 1), w["x0"]))
    lines: list[list[dict[str, Any]]] = []
    current: list[dict[str, Any]] = []
    current_top: float | None = None
    for w in ordered:
        if current_top is None or abs(w["top"] - current_top) <= _LINE_TOL:
            current.append(w)
            current_top = w["top"] if current_top is None else current_top
        else:
            lines.append(sorted(current, key=lambda x: x["x0"]))
            current = [w]
            current_top = w["top"]
    if current:
        lines.append(sorted(current, key=lambda x: x["x0"]))
    return lines


def _assemble(sender_lines: list[str], receiver_lines: list[str]) -> dict[str, str]:
    sender_lines = [s.strip() for s in sender_lines if s.strip()]
    receiver_lines = [s.strip() for s in receiver_lines if s.strip()]

    info = {
        "Sender Name": "",
        "Sender Company": "",
        "Receiver Name": "",
        "Receiver Company": "",
        "Receiver City/Province": "",
    }
    if sender_lines:
        info["Sender Name"] = sender_lines[0][:80]
        if len(sender_lines) > 1:
            info["Sender Company"] = sender_lines[1][:80]
    if receiver_lines:
        info["Receiver Name"] = receiver_lines[0][:80]
        if len(receiver_lines) > 1:
            info["Receiver Company"] = receiver_lines[1][:80]
        for ln in receiver_lines:
            city = _extract_city(ln)
            if city:
                info["Receiver City/Province"] = city[:60]
    return info


def extract_party_blocks(data: bytes) -> dict[str, dict[str, str]]:
    """
    Return ``{tracking_number: {Sender Name, Sender Company, Receiver Name,
    Receiver Company, Receiver City/Province}}`` for every Sender/Receiver block
    found in the document.
    """
    result: dict[str, dict[str, str]] = {}
    try:
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            for page in pdf.pages[:MAX_PAGES]:
                try:
                    words = page.extract_words(
                        use_text_flow=False, keep_blank_chars=False
                    )
                except Exception:  # noqa: BLE001
                    continue
                _scan_page(_group_lines(words), result)
    except Exception as e:  # noqa: BLE001
        logger.warning("party-column extraction failed: %s", e)
    return result


def _scan_page(
    lines: list[list[dict[str, Any]]], result: dict[str, dict[str, str]]
) -> None:
    current_tracking: str | None = None
    i = 0
    while i < len(lines):
        line = lines[i]
        text = " ".join(w["text"] for w in line)
        has_sender = any(w["text"].lower().startswith("sender") for w in line)
        has_receiver = any(w["text"].lower().startswith("receiver") for w in line)

        if not (has_sender and has_receiver):
            tn = _TRACK_RE.search(text)
            if tn:
                current_tracking = tn.group(1)
            i += 1
            continue

        if current_tracking is None:
            i += 1
            continue

        recv_word = next(
            (w for w in line if w["text"].lower().startswith("receiver")), None
        )
        if recv_word is None:
            i += 1
            continue
        boundary = recv_word["x0"] - 2.0

        block = [line]
        j = i + 1
        while j < len(lines):
            jtext = " ".join(w["text"] for w in lines[j]).strip()
            if not jtext:
                j += 1
                continue
            if _PARTY_STOP_RE.match(jtext) or _TRACK_RE.search(jtext):
                break
            block.append(lines[j])
            j += 1

        sender_lines: list[str] = []
        receiver_lines: list[str] = []
        for bl in block:
            left = " ".join(w["text"] for w in bl if w["x0"] < boundary)
            right = " ".join(w["text"] for w in bl if w["x0"] >= boundary)
            sender_lines.append(left)
            receiver_lines.append(right)

        if sender_lines:
            sender_lines[0] = re.sub(
                r"^sender\s*:?\s*", "", sender_lines[0], flags=re.I
            )
        if receiver_lines:
            receiver_lines[0] = re.sub(
                r"^receiver\s*:?\s*", "", receiver_lines[0], flags=re.I
            )

        if current_tracking not in result:
            result[current_tracking] = _assemble(sender_lines, receiver_lines)
        i = j
