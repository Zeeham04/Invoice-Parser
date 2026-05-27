"""Abstract base for carrier-specific invoice parsers."""

from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from typing import Any


@dataclass
class ParseResult:
    """Structured output from a single PDF parse attempt."""

    invoice: dict[str, Any]
    shipments: list[dict[str, Any]] = field(default_factory=list)
    adjustments: list[dict[str, Any]] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    raw_text_snippet: str = ""  # first ~500 chars for debugging optional


class BaseInvoiceParser(ABC):
    """Subclass per carrier / invoice layout."""

    carrier_name: str = "Unknown"

    @abstractmethod
    def parse(self, text: str, filename: str) -> ParseResult:
        """Parse full extracted PDF text into invoice rows and detail lines."""

    def safe_float(self, s: str | None) -> float | None:
        if s is None or s == "":
            return None
        cleaned = (
            s.replace(",", "")
            .replace("$", "")
            .replace("CAD", "")
            .strip()
        )
        try:
            return float(cleaned)
        except ValueError:
            return None
