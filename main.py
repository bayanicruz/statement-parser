#!/usr/bin/env python3
"""bank credit card statement PDF → CSV extractor."""
import argparse
import re
from collections import defaultdict
from collections.abc import Iterator
from pathlib import Path

import pandas as pd
import pdfplumber

OUTPUT_DIR = Path(__file__).parent / "output"
DATE_RE = re.compile(r"^\d{2}/\d{2}/\d{4}$")
CARD_RE = re.compile(r"^\d{4}$")
CARD_MAX_OFFSET = 80  # max horizontal distance (pts) from last date to card-number word
COL_NAMES = ["Processed", "Transaction", "Used", "Transaction Details", "Amount ($A)", "Balance"]


def _parse_row(words: list[dict]) -> list[str] | None:
    """Parse one row of PDF words into a transaction record, or None if not a data row."""
    ws = sorted(words, key=lambda w: w["x0"])
    dates = [w["text"] for w in ws if DATE_RE.match(w["text"])]
    if not dates:
        return None

    money = [w for w in ws if w["text"].startswith("$")]
    first_dollar_x0 = money[0]["x0"] if money else float("inf")
    last_date_x0 = max(w["x0"] for w in ws if DATE_RE.match(w["text"]))

    def is_card(w: dict) -> bool:
        return bool(CARD_RE.match(w["text"])) and w["x0"] - last_date_x0 < CARD_MAX_OFFSET

    card = next((w["text"] for w in ws if is_card(w)), "")
    desc = " ".join(
        w["text"] for w in ws
        if not DATE_RE.match(w["text"])
        and not is_card(w)
        and not w["text"].startswith("$")
        and w["x0"] < first_dollar_x0
    )

    if len(money) >= 2:
        amt_x0, bal_x0 = money[-2]["x0"], money[-1]["x0"]
        cr = "".join(w["text"] for w in ws if not w["text"].startswith("$") and amt_x0 < w["x0"] < bal_x0)
        amount = money[-2]["text"] + cr
    else:
        amount = ""

    return [dates[0], dates[1] if len(dates) > 1 else "", card, desc, amount, money[-1]["text"] if money else ""]


def _page_transactions(page) -> Iterator[list[str]]:
    """Yield transaction rows from a single PDF page."""
    by_top: dict[int, list] = defaultdict(list)
    for w in page.extract_words():
        by_top[round(w["top"])].append(w)

    in_table = False
    for _, words in sorted(by_top.items()):
        texts = {w["text"] for w in words}
        if not in_table:
            in_table = "Processed" in texts
            continue
        if "Please" in texts and "refer" in texts:
            break
        if row := _parse_row(words):
            yield row


def extract(pdf_path: str | Path) -> pd.DataFrame:
    """Extract all transactions from a statement PDF into a DataFrame."""
    with pdfplumber.open(pdf_path) as pdf:
        rows = [row for page in pdf.pages for row in _page_transactions(page)]
    if not rows:
        raise ValueError(f"No transactions found in {pdf_path}")
    return pd.DataFrame(rows, columns=COL_NAMES)


def main() -> None:
    parser = argparse.ArgumentParser(description="Extract transactions from a bank statement PDF.")
    parser.add_argument("pdf_path", type=Path)
    args = parser.parse_args()
    OUTPUT_DIR.mkdir(exist_ok=True)
    csv_path = OUTPUT_DIR / args.pdf_path.with_suffix(".csv").name
    extract(args.pdf_path).to_csv(csv_path, index=False)
    print(csv_path)


if __name__ == "__main__":
    main()
