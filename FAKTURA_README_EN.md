# faktura.py — Native PDF Invoice Extraction by Pure Rules (No LLM)

`faktura.py` uses pdfplumber to directly parse **native PDFs with a text layer**, extracting invoice headers and line items via regex and table-row rules. **No LLM calls** — local extraction only, fast, zero token consumption, deterministic and reproducible.

> Best for: native (non-scanned) PDF invoices exported from the Serbian eFaktura system.
> Scanned/image-only PDFs have no text layer and cannot be processed here — use `racun.py` / `invoice.py` instead.

## Quick Use

```powershell
python faktura.py
```

A file dialog opens; select a PDF and it is processed automatically, producing:

```text
invoice_summary.xlsx
```

If the Excel file is already open, it is saved as `invoice_summary_FIX_HHMMSS.xlsx`.

**Language variants** (identical logic; only the UI language and output filename differ):

| Script | Language | Output File |
|--------|----------|-------------|
| `fakturaen.py` | English | `invoice_summary.xlsx` |
| `fakturars.py` | Serbian | `Zbir_Faktura.xlsx` |

## Main Features

- Tkinter file dialog **multi-select** of native PDF invoice files (process several invoices at once).
- Pure-rule extraction (based on `match_invoices.py` pdfplumber logic):
  - Header: invoice number, seller PIB, buyer PIB, `DATUM PROMETA`, official amount payable
  - Line items: description, quantity, unit price, unit (`Jedinica_Mere`), discount, net amount, PDV rate, tax status text
  - Line items support **three column layouts** (determined by whether the header has a discount column): standard mode (qty unit_price unit discount net tax_rate), special-taxation mode (qty unit_price unit discount net tax_status_text), no-discount mode (qty unit_price unit net tax_rate) — auto-detected.
  - Trailing non-numeric tokens are treated as tax-status text (e.g. `Posebni postupci oporezivanja`) and are never parsed as numbers.
- **PIB enforced as 9 digits**: regex `[0-9]{9}` + boundary `(?!\d)` so business-registry numbers etc. are never mistaken for a PIB.
- **Multi-invoice per-page split**: a PDF with several invoices (e.g. `*_sef.pdf` batches) yields one invoice per page.
- **Multi-page continuation merging**: when the same invoice number spans pages, continuation items are appended; continuation pages with no items (e.g. avans continuation pages) are auto-skipped, preventing double-counting.
- **eFaktura continuation invoice numbers**: in addition to `Broj fakture` / `račun br.`, the `Generisao sistem eFaktura pod brojem:00084/26` format is recognized.
- **Continuation footer merging**: even without line items, continuation pages have their official amount, prepayment deduction, PIBs and transaction date merged upward.
- Per-line VAT = `round(net × rate, 2)`; gross = net + VAT.
- Rate is taken from the item row's `PDV stopa`: pure digits use that rate; non-digit tax text (e.g. `Nije predmet oporezivanja PDV`) yields 0 VAT — never forced to 10%/20% from the footer.
- **Prepayment deduction** (per `avans3.py` semantics): **dynamic iteration over all tax-rate brackets**, matching `Ukupna osnovica - stopa X%` / `Ukupan iznos osnovice umanjen za osnovicu po avansu - stopa X%` / `Ukupan PDV umanjen za PDV po avansu - stopa X%` lines, summing per-rate differences to yield `Prepayment_Deduction` (gross deduction) and `Actual_Payment` (net due), included in reconciliation.
- Each invoice gets a summary row reconciled against the official amount (`OK` / `REVIEW` / `ERR`).
- **Reconciliation breakdown**: the console prints per-invoice `Net / VAT / Gross / Prepay / Due / Official`, so any mismatch can be pinpointed to net, VAT, deduction, or official extraction.
- Credit notes (storno / knjižno odobrenje / red-letter) auto-converted to negatives.
- racun.py-style structured table in PowerShell; also written to Excel.
- **Reconciliation differences → Footer comparison table**: when any invoice is not `OK`, a secondary table based purely on raw PDF footer data (Net / VAT / Total / Prepay / Actual) with a bottom SUM row is printed below the main table, for direct comparison against official figures.
- **No Chinese translation column** (no LLM means no translation; add translations separately if needed).

## Difference vs racun.py / invoice.py

| Item | faktura.py | racun.py / invoice.py |
|------|------------|------------------------|
| LLM calls | None — pure pdfplumber rules | Yes — Gemini parsing |
| PDF type | Native (text-layer) only | Native + scanned (invoice.py) |
| Chinese translation | Not output | Outputs `Description_CN` |
| Speed | Seconds, no token cost | API-dependent, rate-limited |
| Determinism | Deterministic, reproducible | LLM hallucination risk |
| Multi-invoice batches | Auto-split by page | Relies on model recognition |

## Output Structure

### Console Table

```text
Invoice_ID  Description  Seller PIB  Buyer PIB  DATUM PROMETA  Qty  Jed  Price  Net  VAT  Total  Prepay  Actual  Status
```

Each invoice's summary row shows the invoice number, seller & buyer PIB, `DATUM PROMETA`, and that invoice's quantity, unit (`Jed`), net amount, VAT, gross total, prepayment deduction (`Prepay`), actual payment (`Actual`), and reconciliation status.

### Excel Columns

| Column | Description |
|--------|-------------|
| `Invoice_ID` | Invoice number (e.g. IF26-0134) |
| `Seller_PIB` | Seller PIB; filled only on summary rows |
| `Buyer_PIB` | Buyer PIB; filled only on summary rows |
| `Datum_prometa` | Transaction/supply date; filled only on summary rows |
| `Description` | Serbian-language original description |
| `Quantity` | Quantity |
| `Jedinica_Mere` | Unit of measure (e.g. kom / t / km / paušal); filled only on detail rows |
| `Price` | Unit price |
| `Net_Amount` | Net amount (excluding VAT) |
| `PDV_Rate` | VAT rate (e.g. 20%) |
| `VAT_Amount` | VAT = round(Net × rate, 2) |
| `Total_Amount` | Detail rows = net + VAT (gross); summary/total rows = gross before deduction |
| `Prepayment_Deduction` | Prepayment deduction (gross) = Σ per rate (base deduction + VAT deduction); filled on summary/total rows |
| `Actual_Payment` | Actual amount due after deduction = official amount (`Iznos za plaćanje`); detail rows = that row's gross; summary row = invoice due |
| `Status` | Reconciliation status |

## Amount Reconciliation Rules

- Per-line VAT is independently rounded to 2 decimals, then summed (matching the Serbian invoice per-line-cent rule).
- Prepayment deduction: `Prepayment_Deduction = Σ per rate [(Ukupna osnovica − osnovica umanjena za avans) + (Ukupan PDV − PDV umanjen za avans)]`.
- Actual payment = gross total of detail rows − `Prepayment_Deduction`.
- **Semantic distinction**: `Total_Amount` = gross before deduction (recomputed from detail rows); `Actual_Payment` = amount due after deduction (official `Iznos za plaćanje`, fallback `Ukupan iznos fakture`).

### Reconciliation Status Determination

```
calc  = sum(all detail-row Total_Amount) − Prepayment_Deduction   ← actual due
off   = official amount (Iznos za plaćanje / Ukupan iznos fakture)
diff  = round(abs(calc − off), 2)                               ← difference rounded to 2 decimals
```

| Difference | Status | Meaning | Console Output |
|------------|--------|---------|----------------|
| `0.00` | **OK** | Reconciliation passed; actual due matches official exactly | `✅ Reconciliation OK` |
| `≤ 0.01` | **REVIEW** | 1-cent drift from normal rounding; needs manual review | `⚠️ 1-cent difference, review` |
| `> 0.01` | **ERR** | Mismatch; investigate net/VAT/deduction/official extraction | `❌ Mismatch` |

**Why is 0.01 not ERR?** Each line's VAT `round(Net × rate, 2)` is entered per-cent then summed, whereas the official figure sums nets first then applies the rate — two different rounding paths naturally produce ±0.01 drift. This is routine in Serbian eFaktura, not an extraction error.

- When no official amount is found, the actual due is used and `NO_OFFICIAL` is flagged.
- With no prepayment (`Prepayment_Deduction = 0`), actual payment equals the gross total, matching legacy behavior.
- Official-amount regex is confined to the current line (`[^\r\n]*?`) to avoid swallowing the next line's number.

## Footer Comparison Table (shown automatically on reconciliation differences)

When any invoice has a status other than `OK`, a **Footer data comparison table** is printed below the main console table:

```text
📋 Footer data comparison table (reconciliation-difference invoices only, all from raw PDF footer)
Broj Fakture                  Net            VAT          Total         Prepay         Actual
----------------------------------------------------------------------------------------------
IF26-0134                 12,000.00      2,400.00     14,400.00          0.00      14,400.00
IF26-0135                  8,500.00      1,700.00     10,200.00        500.00       9,700.00
----------------------------------------------------------------------------------------------
SUM                       20,500.00      4,100.00     24,600.00        500.00      24,100.00
```

| Column | Source | Description |
|--------|--------|-------------|
| `Net` | `footer_net` | PDF raw footer's `Ukupna osnovica` / `Total net` |
| `VAT` | `footer_vat` | PDF raw footer's `Ukupan PDV` / `Total VAT` |
| `Total` | `footer_total` | PDF raw footer's `Ukupan iznos fakture` |
| `Prepay` | `prep_deduction` | Prepayment deduction (gross); derived from gross-total-minus-due difference |
| `Actual` | `grand_total` | Actual due = `Iznos za plaćanje` (after deduction) |
| `SUM` | Sum of the above | Bottom total row |

> The comparison table comes **entirely from the raw PDF footer**, is not used in item-level calculation, and helps quickly locate the source of differences (net, VAT, deduction, or official extraction).

## Header Extraction Rules

| Field | Extraction Method |
|-------|-------------------|
| `Invoice_ID` | `Broj fakture` / `račun br.` / `Generisao sistem eFaktura pod brojem` / `IF\d{2}-\d{3,}` etc. |
| `Seller_PIB` | PIB inside the `Prodavac` block; `[0-9]{9}` + `(?!\d)` boundary, enforced 9 digits |
| `Buyer_PIB` | PIB after `PIB kupca`; same 9-digit enforcement |
| `Datum_prometa` | Date after `Datum prometa` |
| Official amount | `Iznos za plaćanje` preferred, fallback `Ukupan iznos fakture` / `Grand total` |
| Prepayment deduction | Derived from `Ukupan iznos fakture − Iznos za plaćanje`; only recognized when the gross total is strictly higher than the final due |

Only 9-digit numbers explicitly present on the invoice are treated as PIBs; MB, phone numbers, and other identifiers are never mistaken for PIB.

## Prepayment Deduction Constraints

The script no longer trusts individual avans fields; it first establishes a Footer amount closure:

```text
Gross total       = Ukupan iznos fakture
Final amount due  = Iznos za plaćanje
Prepayment deduction = Gross total − Final amount due
```

Constraint rules:

- When gross total equals final due, deduction is forced to `0.00`;
- Deduction must never be negative;
- Only when both a valid gross total and a valid final due exist in the footer **and** the gross is strictly higher is a deduction confirmed;
- `Ukupna osnovica ... umanjen ...` and `Ukupan PDV ... umanjen ...` lines are reference data for post-deduction base/VAT only and cannot alone establish a deduction;
- If the footer cannot form a valid closure, the suspected deduction is ignored and a warning is printed;
- A `0.01` VAT difference may be treated as rounding tail-drift; only beyond that range does it become `ERR`.

For example:

```text
Ukupan iznos fakture = 4,360,000.00
Iznos za plaćanje    = 2,180,000.00
Prepayment deduction = 2,180,000.00
```

If both are `1,286,000.00`, the prepayment deduction must be `0.00` — page appearance of avans-related text alone does not justify a second deduction.

> **Implementation detail**: The deduction is recomputed inside `build_report` from the footer closure, overriding the raw `extract_avans_deduction()` value. `extract_avans_deduction()` still **dynamically iterates over all tax-rate brackets**, computing `(Ukupna osnovica − osnovica umanjena za avans) + (Ukupan PDV − PDV umanjen za avans)` per rate then summing, and serves as the `Prepay` column value in the footer table plus a warning reference when closure fails.

## Status Levels

Each invoice keeps three main reconciliation statuses:

| Status | Condition |
|--------|-----------|
| `OK` | Actual due exactly matches the official amount |
| `REVIEW` | Only a `0.01` rounding tail-drift; needs manual review |
| `ERR` | Difference exceeds `0.01`, or prepayment/detail/official amounts cannot be closed |

`ERR` does not necessarily mean the invoice is wrong — it can also indicate a missed detail line, incomplete footer fields, or layout incompatibility. Always re-check the original PDF.

## Prerequisites

### Runtime Environment

- Python 3.x (Windows official installer includes `tkinter` — no separate install needed)
- Third-party packages to install: `pdfplumber`, `pandas`, `openpyxl`
  - `tkinter` is a standard library (`Tk`/`filedialog`); included by default with python.org installs

### One-Command Install on a New Machine

```powershell
pip install pdfplumber pandas openpyxl
```

### Verified Versions in Current Environment (reference)

| Dependency | Verified Version |
|------------|-----------------|
| Python | 3.12.10 |
| pdfplumber | 0.11.5 |
| pandas | 2.3.3 |
| openpyxl | 3.1.5 |
| tkinter | Standard library |

### PDF File Requirements

- Must be a **native PDF invoice with a text layer** (eFaktura export; Ctrl+A selectable text works).
- Scanned / pure-image PDFs have no text layer and will prompt "empty PDF / no text layer" — use `racun.py` / `invoice.py` instead.
- Invoices should follow the Serbian eFaktura layout: header contains `Broj fakture`, `Iznos za plaćanje`; detail rows contain quantity/unit-price/unit/discount/net/rate. Other languages/layouts may not be fully recognized.
- A single PDF can contain multiple invoices (e.g. `*_sef.pdf`); the script auto-splits by page. A single invoice can span multiple pages (continuations auto-merged/skipped).

## License
