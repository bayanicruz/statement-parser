# statement-parser

Extracts transactions from bank credit card statement PDFs into CSV.

No Java required — uses [pdfplumber](https://github.com/jsvine/pdfplumber) for pure-Python PDF parsing.

## Requirements

- Python 3.12+
- [uv](https://docs.astral.sh/uv/)

## Installation

```bash
git clone https://github.com/bayanicruz/statement-parser.git
cd statement-parser
uv sync
```

## Usage

```bash
uv run main.py path/to/statement.pdf
```

The CSV is written to `output/<statement-name>.csv`.

## Compatibility

Tested against **bank Rewards Black** credit card statements. The parser detects the transaction table by looking for the `Processed` column header and stops at the `Please refer to the last four digits...` footer — so it should work across statement periods without hardcoded page numbers.

May work with other bank credit card statement formats that share the same layout.

## License

MIT
