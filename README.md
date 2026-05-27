# Invoice Parser

Desktop-friendly local web app: drag-and-drop UPS-style PDF invoices, preview extracted billing data, and export a formatted Excel workbook.

## Setup

```bash
pip install -r requirements.txt
```

## Run

```bash
python app.py
```

Open a browser at **http://localhost:5000**.

## Adding other carriers

1. Create a new module under `parser/` (for example `parser/fedex_domestic.py`).
2. Subclass `BaseInvoiceParser` from `parser/base_parser.py` and implement `parse(self, text: str, filename: str) -> ParseResult`.
3. Return `ParseResult` with:
   - `invoice`: summary fields (`Invoice Number`, `Account Number`, `Billed Amount`, `Tax`, …).
   - `shipments`: list of per-tracking rows matching keys expected by `exporter/excel_export.py`.
   - `adjustments`: optional correction rows.
   - `warnings`: human-readable parse warnings for the UI.
4. In `parser/pdf_parser.py`, extend `detect_invoice_type()` (or add routing logic) to instantiate your parser when the PDF text matches your carrier’s templates.

Keep extraction tolerant: missing fields should be blank or `None`, never raise.

## Stack

- **Backend:** Flask  
- **PDF:** pdfplumber (primary), pdfminer.six (fallback for weak text layers)  
- **Excel:** openpyxl  
- **Frontend:** Single-page HTML + vanilla JavaScript  

No database or cloud services — everything runs on your machine.
