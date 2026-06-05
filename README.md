# UPS Invoice Parser

A Flask web application that takes UPS Canada invoice PDFs, extracts all the
important financial data from them, and exports everything to a clean Excel
spreadsheet. You drop your PDFs in, click Parse, and get a formatted `.xlsx`
file with two tables — one summary view and one charge breakdown.

It runs on [Railway](https://railway.app/) and requires no database, no login,
and no configuration beyond deploying it.

---

## What problem does this solve?

UPS sends invoices as PDFs. There are three distinct PDF layouts depending on
which account the invoice is for and what kind of shipment it covers. Manually
keying those numbers into a spreadsheet every month takes time and introduces
errors. This tool automates the extraction entirely — you feed it any number of
PDFs at once and it produces a spreadsheet you can send directly to accounting.

---

## The three invoice formats

Understanding the formats is the single most important thing for maintaining or
debugging this tool. UPS uses three layouts, and the parser handles each one
differently.

### Format A — Delivery Service Invoice (`delivery`)

These are the domestic and export shipment invoices. The heading "Delivery
Service Invoice" appears near the top right of the first page. You'll see
service categories like "UPS CampusShip", "Worldwide Service", and "UPS Returns"
listed in the Summary of Charges section.

**Key identifiers in the PDF text:**
- `"Delivery Service Invoice"` appears as a heading
- Invoice number label: `"Invoice Number  0000Y8A864026"`
- Amount label: `"Amount due this period  CAD 562.28"`
- Tax label: `"Total Taxes  30.38"` (the summary line; individual lines like
  `"Total Taxes GST R105453328  21.32"` are sub-entries and ignored)

**What gets extracted:** invoice number, account, dates, billed amount, tax,
incentive savings (stored as a negative discount), and the dollar amount for
each service category (Worldwide Service, CampusShip, Returns, Adjustments).

**Government Charges on these invoices: always 0.** Delivery invoices don't
carry customs duties — they're domestic or export shipments.

---

### Format B — Customs Brokerage, account 4172AV (`brokerage_4172av`)

These are the import returns invoices — engineering samples being returned from
Chinese manufacturers. The heading says "CUSTOMS BROKERAGE INVOICE-UPS CARRIED
SHIPMENTS" and the account number is 4172AV.

The page structure is:
- Page 1: Summary of Charges (gross charges → discounts → Net Payable)
- Page 2: UPS Returns Shipment Detail (per-shipment customs info)
- Page 3: Discount summary

**A common source of confusion:** The PDF contains a line that says
`"Total Government Charges  1.20"`. Despite the name, this is *not* customs
duty — it's GST on the brokerage service fee, and it belongs in the
`brokerage_gst` field. For this reason, `government_charges` is hardcoded to
`0.0` for all 4172AV invoices without exception. If you see a nonzero
Government Charges value appear for a 4172AV invoice, that is a bug.

**What gets extracted:** invoice number, account, dates, Net Payable, and a
detailed charge breakdown: Import Freight, Fuel Surcharge, Print Label, Surge
Fees, Government Agency Fee, Tariff Line Fee, Brokerage GST/HST, Discounts
(stored as negative).

---

### Format C — Customs Brokerage, account Y8A864 (`brokerage_y8a864`)

These are US export invoices where actual customs duty has been imposed. The
PDF heading is the same "CUSTOMS BROKERAGE INVOICE-UPS CARRIED SHIPMENTS" but
the account is Y8A864 and the comments field says something like "CUSTOMS DUTY
AND/OR VAT IMPOSED ON YOUR SHIPMENT".

The page structure is:
- Page 1: Summary showing "Government Charges" and "Brokerage Charges" as
  two top-level line items
- Pages 2–3: Import Shipment Detail with per-tariff-class duty breakdown

**Important:** The Summary of Charges reads:
```
Government Charges     480.36
Brokerage Charges       94.08
Total Invoice           574.44
Net Payable  CAD        574.44
```

The `Government Charges` value here (480.36) *is* real customs duty and goes
into the `government_charges` field. This is different from Format B, where a
similarly-named line is actually GST.

**A known PDF text extraction quirk:** pdfplumber sometimes extracts the amount
on a *separate line* from its label, so the text looks like:
```
Government Charges
480.36
```
instead of `Government Charges  480.36`. The parser handles this with
`_find_multiline_float()`, which tries the next-line version first.

**What gets extracted:** invoice number, account, dates, Net Payable,
Government Charges (duties), and the per-shipment charge table: Tariff Line
Fee, Duty, Merchandise Processing Fee, Disbursement Fee, Entry Prep Fee,
PGA Disclaim Fee.

---

## Format detection logic

The parser **never** uses the filename to decide which format an invoice is.
It always reads the PDF text. Here's the decision tree, in order:

1. Does the text contain `"DELIVERY SERVICE INVOICE"`? → `delivery`
2. Does the text contain `"CUSTOMS BROKERAGE INVOICE"`?
   - Extract the account number (`Account No.: XXXX`)
   - If account == `4172AV` → `brokerage_4172av`
   - If a line starting with `"Duty  "` exists → `brokerage_y8a864`
   - If any other account is found → `brokerage_y8a864`
   - If no account found → `brokerage_4172av` (safe fallback)
3. Does the text contain `"UPS CAMPUSSHIP"`? → `delivery`
4. None of the above → `delivery` (safe default)

This matters because account Y8A864 can appear on *both* delivery service
invoices and customs brokerage invoices. You cannot use the account number
alone to decide — you have to read the invoice type heading.

---

## Project structure

```
invoice-parser/
│
├── app.py                  — Flask application (routes only, no business logic)
│
├── parser/
│   ├── ups_parser.py       — Universal parser (the main file; handles all 3 formats)
│   ├── ups_domestic.py     — Legacy TYPE_A parser (kept as reference, not imported)
│   ├── ups_import.py       — Legacy TYPE_B/C parser (kept as reference, not imported)
│   └── pdf_parser.py       — Legacy routing layer (not imported by new code)
│
├── exporter/
│   └── excel_export.py     — Builds the two-table Excel workbook
│
├── static/
│   └── index.html          — Single-page frontend (pure HTML/CSS/JS, no framework)
│
├── requirements.txt
└── nixpacks.toml           — Railway build config (installs tesseract, poppler)
```

---

## How data flows through the app

Here's what happens from the moment you click "Parse & Export" to when the
Excel file lands in your downloads folder:

```
Browser                       Flask (app.py)                  parser/ups_parser.py
──────                        ──────────────                  ────────────────────
1. User drops PDFs
   into the dropzone

2. Clicks "Parse & Export"
   → POST /api/parse
     multipart form,
     each PDF as a field
                              3. api_parse() receives
                                 each file as bytes
                                                              4. parse_invoice(bytes)
                                                                 _extract_text()
                                                                   pdfplumber, all pages
                                                                 _detect_format()
                                                                 extract fields by regex
                                                                 return dict

                              5. Wraps each result:
                                 {ok, filename, invoice,
                                  warnings, detected_type}
                                 Returns JSON array

6. Frontend stores
   the invoice dicts,
   shows results table

7. User clicks
   "Download Summary"
   → POST /api/export/summary
     JSON: {invoices: [...]}
                              8. api_export_summary()
                                 build_workbook_bytes()
                                 returns .xlsx bytes

9. Browser triggers
   the file download
```

---

## The Excel output

The workbook has a single sheet called `Sheet1` with two tables stacked
vertically. The sheet is frozen at row 2 so you can scroll right through the
charge breakdown while keeping the invoice numbers visible.

### Table 1 — Invoice Summary (columns A through L)

| Col | Header | Contents |
|-----|--------|----------|
| A | Invoice Number | String formatted as text (`@`) to preserve leading zeros |
| B | Account Number | e.g. `Y8A864`, `4172AV` |
| C | Amount Due | Always `$0.00` — invoices are pre-paid at time of export |
| D | Invoice Date | e.g. `January 10, 2026` — plain text, never a date serial |
| E | Invoice Status | Always `Closed` |
| F | Payment Status | Always `Accepted` |
| G | Subtotal | **Excel formula** `=J{row}-I{row}-H{row}` (Billed − Govt − Tax) |
| H | Tax | Float, accounting number format |
| I | Government Charges | Float — real customs duties only; 0 for delivery and 4172AV |
| J | Billed Amount | Float — Net Payable or Amount Due from the PDF |
| K | Due Date | Plain text date |
| L | Type | `Domestic/Export` or `Import` |

After all the data rows there are footer rows:
- **Grand total row** — `=SUM(G2:G{n})` formulas for columns G–J
- *(blank row)*
- **Per-account header** — column labels: Subtotal / Tax / Govt Charges / Billed Total
- **4172AV row** — sum of all 4172AV invoices
- **Y8A864 row** — sum of all Y8A864 invoices

**Why Column G is a formula, not a number:**
The subtotal is `Billed − GovtCharges − Tax`. Rather than computing that in
Python and writing a number, the exporter writes an Excel formula string. This
means if you manually correct a Tax or Government Charges value in Excel, the
Subtotal cell updates automatically. It also means the cell shows the formula
`=J2-I2-H2` in the formula bar, which makes auditing straightforward.

**Why dates are plain text:**
If Excel receives a Python `datetime` object it stores it as a date serial
number (something like `46026`) and applies a locale-specific display format
that can change depending on regional settings. Storing dates as plain strings
like `"January 10, 2026"` means the cell always looks exactly right without
any format codes.

### Table 2 — Charge Breakdown (columns A through U)

This table repeats the invoice number and account in columns A–B, then has one
column per charge type:

Import Freight · Fuel Surcharge · Print Label · Surge Fees · Discounts Applied ·
Worldwide Service · UPS CampusShip · UPS Returns · Adjustments & Other Charges ·
Service Charges · Govt Agency Fee · Additional Tariff Line Fee · Brokerage GST/HST ·
Duty (US Customs) · Merchandise Processing Fee · Disbursement Fee · Entry Prep Fee ·
PGA Disclaim Fee · **Total Charges**

Most of these will be zero for any given invoice — only the fields relevant to
its format will be populated. Discounts Applied is stored as a negative number.

**Row order in both tables:**
Invoices are sorted 4172AV first (by invoice date, oldest to newest), then
Y8A864 (same). This is intentional — it groups each account's invoices together
chronologically, matching the layout that accounting expects.

---

## Running locally

### Prerequisites

- Python 3.10 or later
- `pip` and a virtual environment tool

### Setup

```bash
# Clone the repo
git clone https://github.com/Zeeham04/Invoice-Parser.git
cd Invoice-Parser

# Create and activate a virtual environment
python -m venv .venv
source .venv/bin/activate        # macOS / Linux
# or
.venv\Scripts\activate           # Windows PowerShell

# Install Python dependencies
pip install -r requirements.txt
```

On macOS/Linux you also need Poppler for pdf2image:
```bash
brew install poppler tesseract   # macOS
sudo apt-get install poppler-utils tesseract-ocr   # Debian / Ubuntu
```

On Windows, install [Poppler for Windows](https://github.com/oschwartz10612/poppler-windows)
and add its `bin/` folder to your PATH.

### Running

```bash
python app.py
```

Open `http://localhost:5000`. No database, no environment variables needed for
basic local use.

---

## Deploying to Railway

The repo includes `nixpacks.toml` which tells Railway to install system
dependencies before building the Python environment:

```toml
[phases.setup]
nixPkgs = ["tesseract", "poppler_utils"]
```

Railway auto-detects the Python project, runs `pip install -r requirements.txt`,
and starts the app. Every `git push` to `main` triggers a redeploy automatically.

The deploy typically takes 60–90 seconds. Watch the Railway logs panel — it
shows startup messages and the per-invoice debug lines as PDFs are parsed.

---

## API reference

### `POST /api/parse`

Upload one or more PDF invoices. Returns structured data for each file.

**Request:** `multipart/form-data`, field name `files`, one or more PDF files.

**Response:**
```json
{
  "results": [
    {
      "ok": true,
      "filename": "invoice_jan_2026.pdf",
      "detected_type": "Domestic/Export",
      "invoice": {
        "invoice_number":     "0000Y8A864026",
        "account_number":     "Y8A864",
        "invoice_date":       "January 10, 2026",
        "due_date":           "January 26, 2026",
        "invoice_type":       "delivery",
        "invoice_type_label": "Domestic/Export",
        "amount_due":         "$0.00",
        "invoice_status":     "Closed",
        "payment_status":     "Accepted",
        "billed_amount":      562.28,
        "tax":                30.38,
        "government_charges": 0.0,
        "discounts_applied":  -709.20,
        "campus_ship":        495.58,
        "ups_returns":        36.32,
        "worldwide_service":  0.0,
        "adjustments_other":  0.0
      },
      "warnings": []
    }
  ]
}
```

If a file fails to parse, `ok` is `false`, `invoice` is `{}`, and `warnings`
contains the error message.

---

### `POST /api/export/summary`

Build and download the two-table Excel workbook.

**Request:** JSON body:
```json
{
  "invoices": [ "...array of invoice dicts from /api/parse..." ],
  "filename": "my_summary"
}
```

`filename` is optional — defaults to `invoice_summary_YYYYMMDD_HHMMSS.xlsx`.

**Response:** Binary `.xlsx` file download.

---

### `GET /api/health`

Returns `{"status": "ok"}`. Used for Railway health checks and uptime monitoring.

---

## Debugging a parse failure

When a PDF produces wrong values or fails entirely, the Railway logs are your
best starting point. Every successful parse emits:

```
Parsed [delivery] 0000Y8A864026 | Billed=562.28 Tax=30.38 GovtCharge=0.0
```

And every time the Excel exporter writes the Subtotal formula cell:

```
Writing subtotal formula to G2: =J2-I2-H2
G2 cell value after write: '=J2-I2-H2'
```

**Tax shows 0.0 but the PDF has tax:**
The PDF's tax section uses a format we haven't seen. Run `_extract_text()` on
the file locally, find the "Total Taxes" lines, and update `_extract_tax()`
in [parser/ups_parser.py](parser/ups_parser.py).

**Government Charges is 0 for a Y8A864 invoice:**
The PDF is likely wrapping the dollar amount to the next line. Check whether
`_find_multiline_float("Government Charges", text)` is matching. To inspect
the raw extracted text locally:

```python
from parser.ups_parser import _extract_text
text = _extract_text("path/to/invoice.pdf")
# Search for "Government Charges" and see what surrounds it
idx = text.find("Government Charges")
print(repr(text[idx:idx+60]))
```

**Invoice number is missing or wrong:**
Delivery invoices use the label `"Invoice Number"` while brokerage invoices
use `"Invoice No.:"`. Confirm `_detect_format()` is returning the right format
first — if it misidentifies a brokerage invoice as delivery, the wrong regex
pattern runs.

**Column G in Excel shows a number instead of a formula:**
Click the cell and look at the formula bar. If it shows a number, the formula
string was overwritten by a computed value somewhere. Check the Railway logs for
`Writing subtotal formula to G{n}` lines and verify that `G{n} cell value after write`
shows `'=J{n}-I{n}-H{n}'`. If the cell is correctly set but Excel still shows
a number, the workbook's `fullCalcOnLoad` property may not be set — confirm the
exporter sets `wb.calculation.fullCalcOnLoad = True`.

**The whole PDF fails to parse (`UPSParseError: No text could be extracted`):**
The PDF is likely a scanned image with no embedded text layer. The current
parser does not support OCR. See the known limitations section below.

---

## Adding support for a new invoice format

If UPS adds a new layout or you add a third account with different charge lines:

1. **Add a detection branch** in `_detect_format()` in
   [parser/ups_parser.py](parser/ups_parser.py). Use a string that only
   appears in that format's header, not the others.

2. **Add field extraction** — a new `elif fmt == "new_format":` block inside
   `parse_invoice()`. Use the same `_find()` / `_find_float()` helpers.
   For charges that might be on the next line after the label, use
   `_find_multiline_float("Label Text", text)`.

3. **Add any new charge field names** to the default `result` dict at the top
   of `parse_invoice()` with a default of `0.0`. This ensures the frontend
   and exporter never receive a `KeyError`.

4. **Add a column** to Table 2 in [exporter/excel_export.py](exporter/excel_export.py)
   by appending the new key to `T2_CHARGE_KEYS` and a header string to
   `T2_HEADERS`. Extend `T2_WIDTHS` with a column width.

5. **Add a sort group** in `_ACCOUNT_ORDER` in excel_export.py so the new
   account appears in the right position relative to 4172AV and Y8A864.

The frontend and app.py do not need changes — they pass the full invoice dict
through without inspecting charge fields.

---

## Known limitations

**Scanned (image-only) PDFs are not supported.**
The parser uses pdfplumber which extracts embedded text. If a PDF is a scanned
image with no text layer, pdfplumber returns nothing and `parse_invoice` raises
`UPSParseError`. An OCR fallback using pytesseract is available in the legacy
`parser/pdf_parser.py` but is not integrated into the new universal parser.
If scanned invoices become a requirement, add the OCR path from that file
into `_extract_text()`.

**Multi-invoice PDFs (one PDF, multiple invoices) are treated as a single invoice.**
UPS occasionally bundles several invoices into one PDF. The new universal parser
reads the whole document as one and extracts the first matching values for each
field. If this becomes an issue, the fix is to split the text on the invoice
heading strings (`"DELIVERY SERVICE INVOICE"`, `"CUSTOMS BROKERAGE INVOICE"`)
before calling `parse_invoice`.

**The per-account breakdown is hardcoded to 4172AV and Y8A864.**
If a third account is ever added, update `_ACCOUNT_ORDER` and add a row in
`build_workbook()` in excel_export.py.

---

## Dependencies

| Package | Why it's here |
|---------|---------------|
| `flask` | Web framework — handles HTTP routing |
| `pdfplumber` | Extracts text from PDF pages with good accuracy |
| `openpyxl` | Reads and writes `.xlsx` Excel files |
| `python-dotenv` | Loads `.env` files for local development |
| `pdfminer.six` | Secondary PDF text extractor (used by the legacy parser) |
| `pytesseract` | Python wrapper for the Tesseract OCR engine |
| `pdf2image` | Converts PDF pages to images for OCR processing |
| `Pillow` | Image library required by pdf2image |

---

## Git workflow

The `main` branch deploys automatically to Railway on every push.

```bash
git add -A
git commit -m "describe what changed and why"
git push
```

Watch the Railway dashboard for the deploy to complete, then check the logs
for startup messages or parse-time debug output.
