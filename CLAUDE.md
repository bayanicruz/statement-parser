# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Scripts

### `statements.py` — PDF statement extractor
```bash
uv run statements.py path/to/statement.pdf   # or no arg for file picker
statement-parser                              # entry point after uv sync
```
Extracts transactions from an bank credit card statement PDF → `output/statements/<name>.xlsx`.

### `transaction.py` — CSV budget tracker
```bash
uv run transaction.py path/to/bank.csv        # or no arg for file picker
transaction-tracker                           # entry point after uv sync
```
Converts the bank transaction CSV export → `output/transactions/<name>.xlsx` with a budget summary.

bank CSV downloads to `~/Downloads/bank.csv` with no header row: `Date, Amount, Description`.

### `transaction_v2.py` — PDF budget tracker with date-range filter
```bash
uv run transaction_v2.py path/to/transactions.pdf   # or no arg for file picker
transaction-tracker-pdf                              # entry point after uv sync
```
Parses bank transaction history PDFs (saved from bank Internet Banking) → `output/transactions/<name>.xlsx`.
After selecting the PDF, two zenity calendar pickers appear for start and end date. Cancel either to skip that bound and include all transactions on that end.

---

All scripts auto-open the output after writing. If the output already exists, they prompt `[o]verwrite / [v]ersion up` — versioning appends `_v2`, `_v3`, etc.

## Architecture

### `statements.py`

Three logical stages:

1. **Extraction** (`extract`) — opens the PDF with `pdfplumber`, iterates every page, and collects transaction rows via `_page_transactions` / `_parse_row`. Table boundaries are detected automatically: starts when a row contains the word `Processed` (the column header) and stops at `Please refer` (the footer). No page range required.

2. **Parsing** (`_parse_row`) — receives a flat list of word dicts (each with `text`, `x0`, `top` from pdfplumber) for one horizontal band. Uses regex and horizontal position (`x0`) to classify each word as a date, card number (4-digit token close to the last date), money value (`$`-prefixed), or description. Handles `CR` suffix for credits by splitting into a separate `Credit ($A)` column with a negated value.

3. **Post-processing** (`_post_process`) — re-opens the written XLSX with openpyxl to insert a per-card Summary box (rows 1–N), apply AUD currency format, add AutoFilter, auto-fit column widths, normalise all fonts to Calibri, and apply tiered background highlights (+ bold) on the `Processed`–`Amount ($A)` range based on spend: $80+ yellow, $300+ orange, $600+ salmon, $1000+ red.

### Key constants

| Name | Purpose |
|---|---|
| `CARD_MAX_OFFSET` | Max horizontal distance (pts) between a 4-digit token and the last date for it to be classified as a card number rather than part of the description |
| `COL_NAMES` | Fixed output columns: `Processed, Transaction, Used, Transaction Details, Amount ($A), Balance` |
| `OUTPUT_DIR` | `output/statements/` or `output/transactions/` relative to the script; created on first run |

### `transaction.py` and `transaction_v2.py`

`transaction.py` has two stages:

1. **Loading** (`load_transactions`) — reads the headerless bank CSV, flips all amounts to positive, and classifies each row as `Expense` (originally negative) or `Credit` (originally positive). Adds an empty `Category` column for manual tagging.

2. **Post-processing** (`_post_process`) — writes a Summary box (cols G:H) showing Budget / Total Expenses / Total Credits / Net Spent / Remaining. Remaining is green when under budget, red when over. Highlights expense rows with the same spend tiers as `statements.py`; credit rows get a green highlight. Normalises all fonts to Calibri.

`transaction_v2.py` replaces the CSV loading stage with a PDF extractor (`extract`) that uses pdfplumber word positions to parse bank Internet Banking transaction history PDFs. Key details:
- Groups words by vertical band (`round(top)`), tracks month/year context from `Month YYYY` headers which can appear mid-page or carry across pages
- Detects wrapped descriptions: lines with only description text within `SUFFIX_GAP_MAX` pts below a date band are appended as suffixes; lines above a date band are held as prefixes
- Debit vs credit determined by x0 position of the `$` amount: x0 ≥ `CREDIT_X_MIN` (480) = credit
- Skips browser chrome, column headers, account info, and "Important information" footer
- After extraction, prompts for date range via `zenity --calendar`; cancelling either picker skips that bound

**Budget constant:** `BUDGET = 6250.00` at the top of `transaction.py` and `transaction_v2.py`.

## Stack

| Component | Notes |
|---|---|
| Python 3.12+ | Managed via `uv`; see `.python-version` |
| pdfplumber | Pure-Python PDF parsing — no Java dependency |
| pandas | DataFrame assembly and XLSX writing |
| openpyxl | Post-processing the XLSX (summary, formatting) |

## Compatibility

Tested against **bank Rewards Black** credit card statements. The auto-detection logic should generalise across statement periods; if extraction returns no rows, the PDF layout may have changed — check that the `Processed` column header and `Please refer` footer text still appear verbatim.

## Statement files

PDFs live in `files/` (not tracked by git). Processed XLSX outputs go to `output/`.

**Naming convention:** `files/{YY}-statement-{DD}-{Mon}.pdf` (e.g. `26-statement-24-Mar.pdf`).

## Monthly workflow

**Statements:**
1. Download the new statement PDF and rename it to `{YY}-statement-{DD}-{Mon}.pdf`, place in `files/`.
2. Run `uv run statements.py` — file picker opens, select the PDF.
3. The XLSX opens automatically. Verify the Summary box totals match the statement.

**Transactions:**
1. Export transactions from bank internet banking as CSV (`~/Downloads/bank.csv`).
2. Run `uv run transaction.py` — file picker opens, select the CSV.
3. The XLSX opens. Check the Summary box Remaining figure against your budget.
4. Fill in the Category column for each transaction.

> `zenity` is required for the file picker (`sudo apt install zenity` if missing).
