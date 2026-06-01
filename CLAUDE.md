# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Running the tool

```bash
# Install dependencies
uv sync

# Run against a statement PDF
uv run main.py path/to/statement.pdf

# Or via the installed entry point (after uv sync)
statement-parser path/to/statement.pdf
```

Output is written to `output/<statement-name>.xlsx`. The script auto-opens the file after writing and prints the path to stdout.

If the output file already exists, the script prompts: `[o]verwrite / [v]ersion up` — versioning appends `_v2`, `_v3`, etc.

## Architecture

Single-file script (`main.py`) with three logical stages:

1. **Extraction** (`extract`) — opens the PDF with `pdfplumber`, iterates every page, and collects transaction rows via `_page_transactions` / `_parse_row`. Table boundaries are detected automatically: starts when a row contains the word `Processed` (the column header) and stops at `Please refer` (the footer). No page range required.

2. **Parsing** (`_parse_row`) — receives a flat list of word dicts (each with `text`, `x0`, `top` from pdfplumber) for one horizontal band. Uses regex and horizontal position (`x0`) to classify each word as a date, card number (4-digit token close to the last date), money value (`$`-prefixed), or description. Handles `CR` suffix for credits by splitting into a separate `Credit ($A)` column with a negated value.

3. **Post-processing** (`_post_process`) — re-opens the written XLSX with openpyxl to insert a per-card Summary box (rows 1–N), apply AUD currency format, add AutoFilter, auto-fit column widths, normalise all fonts to Calibri, and apply tiered background highlights (+ bold) on the `Processed`–`Amount ($A)` range based on spend: $80+ yellow, $300+ orange, $600+ salmon, $1000+ red.

### Key constants

| Name | Purpose |
|---|---|
| `CARD_MAX_OFFSET` | Max horizontal distance (pts) between a 4-digit token and the last date for it to be classified as a card number rather than part of the description |
| `COL_NAMES` | Fixed output columns: `Processed, Transaction, Used, Transaction Details, Amount ($A), Balance` |
| `OUTPUT_DIR` | `output/` relative to the script; created on first run |

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

1. Download the new statement PDF and rename it to `{YY}-statement-{DD}-{Mon}.pdf`, place in `files/`.
2. Run `uv run main.py` — a file picker opens, select the PDF. Or pass the path directly: `uv run main.py files/<statement>.pdf`.
3. The XLSX opens automatically. Verify the Summary box totals match the statement.
4. Review and categorise transactions in the XLSX.

> `zenity` is required for the file picker (`sudo apt install zenity` if missing).
