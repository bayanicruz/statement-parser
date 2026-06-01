#!/usr/bin/env python3
"""bank transaction history PDF → XLSX budget tracker with date-range filter."""
import argparse
import re
import subprocess
from collections import defaultdict
from datetime import date
from pathlib import Path

import pandas as pd
import pdfplumber
from openpyxl import load_workbook
from openpyxl.styles import Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BUDGET = 6250.00
OUTPUT_DIR = Path(__file__).parent / "output" / "transaction_v2"

MONTHS_FULL = {
    "January": 1, "February": 2, "March": 3, "April": 4,
    "May": 5, "June": 6, "July": 7, "August": 8,
    "September": 9, "October": 10, "November": 11, "December": 12,
}
MONTHS_ABBR = {
    "JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MAY": 5, "JUN": 6,
    "JUL": 7, "AUG": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12,
}

DATE_DAY_X_MAX = 50     # day numbers sit at x0 ≤ this
DESC_X_MIN = 90         # description words start at x0 ≥ this
CREDIT_X_MIN = 480      # amounts at x0 ≥ this are credits
SUFFIX_GAP_MAX = 15     # max vertical gap (pts) to treat a desc-only band as a suffix


def _group_by_top(words: list[dict]) -> dict[int, list[dict]]:
    bands: dict[int, list] = defaultdict(list)
    for w in words:
        bands[round(w["top"])].append(w)
    return bands


def _is_skip(words: list[dict]) -> bool:
    texts = {w["text"] for w in words}
    if "Date" in texts and "Description" in texts:
        return True
    if "Balance" in texts and "Funds" in texts:
        return True
    if "Transaction" in texts and "history" in texts:
        return True
    if "bank" in texts and "Internet" in texts and "Banking" in texts:
        return True
    if "bank" in texts and "Rewards" in texts:
        return True
    if any("xxxx" in t for t in texts):
        return True
    if any("https://" in t or t.startswith("©") or t == "ABN" for t in texts):
        return True
    return False


def _is_footer(words: list[dict]) -> bool:
    texts = {w["text"] for w in words}
    return "Important" in texts and "information" in texts


def _is_month_header(words: list[dict]) -> tuple[int, int] | None:
    texts = [w["text"] for w in sorted(words, key=lambda w: w["x0"])]
    if len(texts) == 2 and texts[0] in MONTHS_FULL and re.match(r"^\d{4}$", texts[1]):
        return MONTHS_FULL[texts[0]], int(texts[1])
    return None


def _is_date_row(words: list[dict]) -> bool:
    return any(
        w["x0"] <= DATE_DAY_X_MAX and re.match(r"^\d{1,2}$", w["text"])
        for w in words
    )


def _parse_date(words: list[dict], year: int) -> date | None:
    sorted_w = sorted(words, key=lambda w: w["x0"])
    day_w = next((w for w in sorted_w if w["x0"] <= DATE_DAY_X_MAX and re.match(r"^\d{1,2}$", w["text"])), None)
    mon_w = next((w for w in sorted_w if w["text"].upper() in MONTHS_ABBR), None)
    if not day_w or not mon_w:
        return None
    try:
        return date(year, MONTHS_ABBR[mon_w["text"].upper()], int(day_w["text"]))
    except ValueError:
        return None


def _parse_amount(words: list[dict]) -> tuple[float, str]:
    money = [w for w in words if re.match(r"^\$[\d,]+\.\d{2}$", w["text"])]
    if not money:
        return 0.0, "Expense"
    m = money[0]
    val = float(m["text"].replace("$", "").replace(",", ""))
    return val, "Credit" if m["x0"] >= CREDIT_X_MIN else "Expense"


def _desc_words(words: list[dict]) -> str:
    return " ".join(
        w["text"] for w in sorted(words, key=lambda w: w["x0"])
        if w["x0"] >= DESC_X_MIN and not re.match(r"^\$[\d,]+\.\d{2}$", w["text"])
    )


def extract(pdf_path: Path) -> pd.DataFrame:
    rows: list[dict] = []
    current_year = date.today().year

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            bands = _group_by_top(page.extract_words())
            sorted_bands = sorted(bands.items())
            pending_prefix = ""
            last_date_top = -999

            for top, words in sorted_bands:
                if _is_footer(words):
                    break
                if _is_skip(words):
                    continue

                month_header = _is_month_header(words)
                if month_header:
                    _, current_year = month_header
                    pending_prefix = ""
                    continue

                if _is_date_row(words):
                    tx_date = _parse_date(words, current_year)
                    inline_desc = _desc_words(words)
                    amount, tx_type = _parse_amount(words)
                    full_desc = (pending_prefix + " " + inline_desc).strip()
                    pending_prefix = ""
                    last_date_top = top
                    rows.append({
                        "Date": tx_date,
                        "Amount": amount,
                        "Type": tx_type,
                        "Description": full_desc,
                        "Category": "",
                        "_top": top,
                    })
                else:
                    desc_text = _desc_words(words)
                    if not desc_text:
                        continue
                    if rows and (top - last_date_top) < SUFFIX_GAP_MAX:
                        rows[-1]["Description"] = (rows[-1]["Description"] + " " + desc_text).strip()
                    else:
                        pending_prefix = (pending_prefix + " " + desc_text).strip()

    if not rows:
        raise ValueError(f"No transactions found in {pdf_path}")

    df = pd.DataFrame(rows).drop(columns=["_top"])
    df["Date"] = pd.to_datetime(df["Date"])
    return df


def _pick_date(title: str) -> str | None:
    result = subprocess.run(
        ["zenity", "--calendar", f"--title={title}", "--date-format=%Y-%m-%d"],
        capture_output=True, text=True,
    )
    return result.stdout.strip() or None


def _post_process(xlsx_path: Path, df: pd.DataFrame) -> None:
    expenses = df.loc[df["Type"] == "Expense", "Amount"].sum()
    credits = df.loc[df["Type"] == "Credit", "Amount"].sum()
    net_spent = expenses - credits
    remaining = BUDGET - net_spent

    SUMMARY_ROWS = 6
    ncols = len(df.columns)

    wb = load_workbook(xlsx_path)
    ws = wb.active
    ws.insert_rows(1, SUMMARY_ROWS)

    data_header = SUMMARY_ROWS + 1
    data_start = data_header + 1
    data_end = len(df) + SUMMARY_ROWS + 1

    ws.auto_filter.ref = f"A{data_header}:{get_column_letter(ncols)}{data_end}"

    thin = Side(style="thin")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.merge_cells("G1:H1")
    ws["G1"] = "Summary"
    ws["G1"].font = Font(name="Calibri", bold=True)

    summary_rows = [
        ("Budget",         BUDGET),
        ("Total Expenses", expenses),
        ("Total Credits",  credits),
        ("Net Spent",      net_spent),
        ("Remaining",      remaining),
    ]
    for i, (label, value) in enumerate(summary_rows, start=2):
        ws[f"G{i}"] = label
        ws[f"G{i}"].font = Font(name="Calibri")
        ws[f"H{i}"] = value
        ws[f"H{i}"].number_format = "$#,##0.00"
        ws[f"H{i}"].font = Font(name="Calibri")

    ws["H6"].fill = PatternFill("solid", fgColor="C6EFCE" if remaining >= 0 else "FFB3B3")

    for row in ws["G1:H6"]:
        for cell in row:
            cell.border = box

    amt_col_idx = df.columns.get_loc("Amount") + 1
    type_col_idx = df.columns.get_loc("Type") + 1
    highlight_to = type_col_idx

    expense_fills = [
        (1000, PatternFill("solid", fgColor="FFB3B3")),
        (600,  PatternFill("solid", fgColor="FFCCBC")),
        (300,  PatternFill("solid", fgColor="FFE0B2")),
        (80,   PatternFill("solid", fgColor="FFF2CC")),
    ]
    credit_fill = PatternFill("solid", fgColor="C6EFCE")

    for r in range(data_header, data_end + 1):
        for c in range(1, ncols + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = Font(name="Calibri", bold=cell.font.bold if cell.font else False)

    for r in range(data_start, data_end + 1):
        ws.cell(row=r, column=amt_col_idx).number_format = "$#,##0.00"
        type_val = ws.cell(row=r, column=type_col_idx).value
        amt_val = ws.cell(row=r, column=amt_col_idx).value

        if type_val == "Credit":
            for c in range(1, highlight_to + 1):
                ws.cell(row=r, column=c).fill = credit_fill
        elif isinstance(amt_val, (int, float)):
            for threshold, fill in expense_fills:
                if amt_val >= threshold:
                    for c in range(1, highlight_to + 1):
                        ws.cell(row=r, column=c).fill = fill
                        ws.cell(row=r, column=c).font = Font(name="Calibri", bold=True)
                    break

    for col in ws.iter_cols():
        max_len = max(
            (len(str(cell.value)) for cell in col
             if cell.value is not None and not str(cell.value).startswith("=")),
            default=8,
        )
        ws.column_dimensions[get_column_letter(col[0].column)].width = max_len + 2

    wb.save(xlsx_path)


def main() -> None:
    parser = argparse.ArgumentParser(description="Convert bank transaction history PDF to XLSX budget tracker.")
    parser.add_argument("pdf_path", type=Path, nargs="?")
    args = parser.parse_args()

    pdf_path = args.pdf_path
    if pdf_path is None:
        result = subprocess.run(
            ["zenity", "--file-selection", "--title=Select bank transaction PDF", "--file-filter=*.pdf"],
            capture_output=True, text=True,
        )
        chosen = result.stdout.strip()
        if not chosen:
            raise SystemExit("No file selected.")
        pdf_path = Path(chosen)

    print(f"Extracting transactions from {pdf_path.name}...")
    df = extract(pdf_path)
    print(f"Found {len(df)} transactions ({df['Date'].min().date()} → {df['Date'].max().date()})")

    start_str = _pick_date("Select start date (or Cancel for all)")
    end_str = _pick_date("Select end date (or Cancel for all)")

    if start_str:
        df = df[df["Date"] >= pd.to_datetime(start_str)]
    if end_str:
        df = df[df["Date"] <= pd.to_datetime(end_str)]

    if df.empty:
        subprocess.run(["zenity", "--error", "--text=No transactions in the selected date range."])
        raise SystemExit("No transactions in the selected date range.")

    df = df.reset_index(drop=True)
    print(f"{len(df)} transactions after date filter.")

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    xlsx_path = OUTPUT_DIR / pdf_path.with_suffix(".xlsx").name

    if xlsx_path.exists():
        result = subprocess.run(
            ["zenity", "--list", "--radiolist",
             "--title=File exists",
             f"--text={xlsx_path.name} already exists. Choose an action:",
             "--column=", "--column=Action",
             "TRUE", "Overwrite",
             "FALSE", "Version up"],
            capture_output=True, text=True,
        )
        answer = result.stdout.strip()
        if answer == "Version up":
            stem = xlsx_path.stem
            n = 2
            while xlsx_path.exists():
                xlsx_path = OUTPUT_DIR / f"{stem}_v{n}.xlsx"
                n += 1
        elif answer != "Overwrite":
            raise SystemExit("Cancelled.")

    df.to_excel(xlsx_path, index=False, engine="openpyxl")
    _post_process(xlsx_path, df)
    print(xlsx_path)
    subprocess.Popen(["xdg-open", str(xlsx_path)])


if __name__ == "__main__":
    main()
