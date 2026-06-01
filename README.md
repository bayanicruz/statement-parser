# statement-parser

A set of personal finance scripts that convert bank statement PDFs and transaction CSVs into formatted Excel workbooks with budget summaries.

No Java required — uses [pdfplumber](https://github.com/jsvine/pdfplumber) for pure-Python PDF parsing.

## Requirements

- Python 3.12+
- [uv](https://docs.astral.sh/uv/)
- `zenity` for file picker dialogs (`sudo apt install zenity`)

## Installation

```bash
git clone https://github.com/bayanicruz/statement-parser.git
cd statement-parser
uv sync
```

## Scripts

### `statements.py` — Statement PDF extractor
Extracts transactions from a credit card statement PDF → `output/statements/<name>.xlsx`.

```bash
uv run statements.py path/to/statement.pdf   # or no arg for file picker
statement-parser                              # entry point after uv sync
```

### `transaction.py` — CSV budget tracker
Converts a transaction CSV export → `output/transactions/<name>.xlsx` with a budget summary.

```bash
uv run transaction.py path/to/transactions.csv   # or no arg for file picker
transaction-tracker                               # entry point after uv sync
```

### `transaction_v2.py` — CSV budget tracker with outstanding authorisations
Same as `transaction.py` with an extra step to paste outstanding/pending authorisations and tentative planned purchases.

```bash
uv run transaction_v2.py path/to/transactions.csv   # or no arg for file picker
transaction-tracker-pdf                              # entry point after uv sync
```

## License

MIT
