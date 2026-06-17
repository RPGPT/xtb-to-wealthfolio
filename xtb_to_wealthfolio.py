"""Convert XTB Excel exports into Wealthfolio-friendly CSV output."""

import re
import secrets
from pathlib import Path

import pandas as pd


BUY_PATTERN = re.compile(r"OPEN BUY\s+([0-9.]+)(?:/[0-9.]+)?\s*@\s*([0-9.]+)")
SELL_PATTERN = re.compile(r"CLOSE BUY\s+([0-9.]+)(?:/[0-9.]+)?\s*@\s*([0-9.]+)")
OUTPUT_PREFIX = "xtb_wealthfolio_"
ACCOUNT_PREFIX = "XTB "


def clean_symbol(value):
    """Normalize ticker symbols for Wealthfolio output."""
    if pd.isna(value):
        return ""

    symbol = str(value).strip()
    if symbol.upper().endswith(".US"):
        return symbol[:-3]
    return symbol


def parse_trade_fields(comment, activity_type):
    """Extract quantity and unit price from trade comments."""
    if pd.isna(comment):
        return "", ""

    comment_text = str(comment)
    pattern = BUY_PATTERN if activity_type == "BUY" else SELL_PATTERN
    match = pattern.search(comment_text) if activity_type in {"BUY", "SELL"} else None

    if not match:
        return "", ""

    return match.group(1), match.group(2)


def detect_account_currency(workbook_path: Path):
    """Infer the account currency from the workbook filename."""
    name = workbook_path.name.upper()
    if "USD" in name:
        return "USD"
    if "EUR" in name:
        return "EUR"
    raise ValueError(
        f"Could not infer account currency from filename '{workbook_path.name}'. "
        "Expected filename to contain USD or EUR."
    )


def map_transfer_activity(comment, amount, account_currency):
    """Map XTB transfer rows to Wealthfolio deposit or withdrawal rows."""
    comment_text = "" if pd.isna(comment) else str(comment).lower()
    amount_value = pd.to_numeric(amount, errors="coerce")

    if account_currency == "USD":
        if "eur to usd" in comment_text:
            return "DEPOSIT"
        if "usd to eur" in comment_text:
            return "WITHDRAWAL"

    if account_currency == "EUR":
        if "usd to eur" in comment_text:
            return "DEPOSIT"
        if "eur to usd" in comment_text:
            return "WITHDRAWAL"

    if pd.notna(amount_value):
        return "DEPOSIT" if amount_value >= 0 else "WITHDRAWAL"
    return None


def map_activity_type(raw_type, amount, comment, account_currency):
    """Map raw XTB row types into Wealthfolio activity types."""
    normalized_type = str(raw_type).strip().lower()
    direct_map = {
        "stock purchase": "BUY",
        "stock sell": "SELL",
        "dividend": "DIVIDEND",
        "free funds interest": "INTEREST",
        "deposit": "DEPOSIT",
    }

    if normalized_type in direct_map:
        return direct_map[normalized_type]
    if normalized_type in {"withholding tax", "free funds interest tax"}:
        return "TAX"
    if "withdraw" in normalized_type:
        return "WITHDRAWAL"
    if normalized_type == "transfer":
        return map_transfer_activity(comment, amount, account_currency)
    return None


def choose_input_xlsx():
    """Prompt the user to choose an Excel file from the script directory."""
    current_dir = Path(__file__).resolve().parent
    xlsx_files = sorted(current_dir.glob("*.xlsx"))

    if not xlsx_files:
        raise FileNotFoundError("No .xlsx files found in the same folder as this script.")

    print("Available .xlsx files:")
    for index, file_path in enumerate(xlsx_files, start=1):
        print(f"{index}. {file_path.name}")

    while True:
        choice = input(f"Choose a file [1-{len(xlsx_files)}]: ").strip()
        try:
            selected_index = int(choice)
            if 1 <= selected_index <= len(xlsx_files):
                return xlsx_files[selected_index - 1]
        except ValueError:
            pass
        print("Invalid choice. Try again.")


def make_output_path(output_dir: Path):
    """Create a randomized output CSV path in the working directory."""
    random_hash = secrets.token_hex(6)
    return output_dir / f"{OUTPUT_PREFIX}{random_hash}.csv"


def find_col(dataframe, names):
    """Find a matching column name in a dataframe using case-insensitive lookup."""
    lowered_columns = {str(column).strip().lower(): column for column in dataframe.columns}
    for name in names:
        matched = lowered_columns.get(name.lower())
        if matched:
            return matched
    return None


def load_cash_operations(workbook_path: Path):
    """Load and normalize the XTB Cash Operations sheet."""
    dataframe = pd.read_excel(workbook_path, sheet_name="Cash Operations", header=4)
    dataframe = dataframe[dataframe["Type"].notna()].copy()
    dataframe["Time"] = pd.to_datetime(dataframe["Time"], errors="coerce")
    dataframe["Amount"] = pd.to_numeric(dataframe["Amount"], errors="coerce")
    return dataframe.sort_values("Time", kind="stable")


def build_cash_rows(cash_df, account_currency):
    """Transform cash operation rows into Wealthfolio CSV rows."""
    rows = []

    for source_index, (_, row) in enumerate(cash_df.iterrows(), start=1):
        comment = "" if pd.isna(row.get("Comment")) else str(row["Comment"])
        amount = row.get("Amount", "")
        activity_type = map_activity_type(
            row["Type"], amount, comment, account_currency
        )
        if activity_type is None:
            continue

        quantity = ""
        unit_price = ""
        if activity_type in {"BUY", "SELL"}:
            quantity, unit_price = parse_trade_fields(comment, activity_type)

        rows.append(
            {
                "date": (
                    row["Time"].strftime("%Y-%m-%d %H:%M:%S")
                    if pd.notna(row["Time"])
                    else ""
                ),
                "account": f"{ACCOUNT_PREFIX}{account_currency}",
                "symbol": clean_symbol(row.get("Ticker", "")),
                "activityType": activity_type,
                "quantity": quantity,
                "unitPrice": unit_price,
                "amount": amount,
                "currency": account_currency,
                "fee": 0,
                "comment": f"{comment} | xtb_row={source_index}",
            }
        )

    return rows


def extract_closed_position_row(row, columns, account_currency, closed_index):
    """Build a synthetic sell row from a Closed Positions correction entry."""
    origin_close = "" if pd.isna(row.get(columns["origin_close"])) else str(
        row[columns["origin_close"]]
    ).strip()
    if origin_close.lower() != "correction":
        return None

    symbol = clean_symbol(row.get(columns["symbol"], ""))
    quantity = "" if pd.isna(row.get(columns["qty"])) else str(
        row[columns["qty"]]
    ).replace(",", ".")
    close_dt = pd.to_datetime(row.get(columns["close_time"]), errors="coerce")

    if not symbol or not quantity or pd.isna(close_dt):
        return None

    close_price = ""
    close_price_col = columns.get("close_price")
    open_price_col = columns.get("open_price")
    if close_price_col and pd.notna(row.get(close_price_col)):
        close_price = str(row[close_price_col]).replace(",", ".")
    elif open_price_col and pd.notna(row.get(open_price_col)):
        close_price = str(row[open_price_col]).replace(",", ".")

    position_id = ""
    if columns.get("position_id") and pd.notna(row.get(columns["position_id"])):
        position_id = str(row[columns["position_id"]]).strip()

    instrument_name = ""
    if columns.get("instrument") and pd.notna(row.get(columns["instrument"])):
        instrument_name = str(row[columns["instrument"]]).strip()

    open_dt = ""
    if columns.get("open_time") and pd.notna(row.get(columns["open_time"])):
        open_dt = str(row[columns["open_time"]]).strip()

    comment = (
        "closed_position_correction "
        f"position={position_id} open={open_dt} close={close_dt} "
        f"name={instrument_name} | closed_row={closed_index}"
    )

    return {
        "date": close_dt.strftime("%Y-%m-%d %H:%M:%S"),
        "account": f"{ACCOUNT_PREFIX}{account_currency}",
        "symbol": symbol,
        "activityType": "SELL",
        "quantity": quantity,
        "unitPrice": close_price,
        "amount": "",
        "currency": account_currency,
        "fee": 0,
        "comment": comment,
    }


def load_closed_positions(workbook_path: Path):
    """Try to load the Closed Positions sheet and return None if missing."""
    try:
        return pd.read_excel(workbook_path, sheet_name="Closed Positions", header=4).copy()
    except (ValueError, KeyError, FileNotFoundError):
        return None


def build_closed_position_rows(closed_df, account_currency):
    """Transform Closed Positions corrections into Wealthfolio sell rows."""
    if closed_df is None:
        return []

    columns = {
        "symbol": find_col(closed_df, ["Ticker"]),
        "qty": find_col(closed_df, ["Volume"]),
        "open_price": find_col(closed_df, ["Open Price"]),
        "close_price": find_col(closed_df, ["Close Price"]),
        "open_time": find_col(closed_df, ["Open Time (UTC)"]),
        "close_time": find_col(closed_df, ["Close Time (UTC)"]),
        "origin_close": find_col(closed_df, ["Close Origin"]),
        "position_id": find_col(closed_df, ["Position ID"]),
        "instrument": find_col(closed_df, ["Instrument"]),
    }

    required = [
        columns["symbol"],
        columns["qty"],
        columns["close_time"],
        columns["origin_close"],
    ]
    if not all(required):
        return []

    rows = []
    for closed_index, (_, row) in enumerate(closed_df.iterrows(), start=1):
        extracted = extract_closed_position_row(
            row, columns, account_currency, closed_index
        )
        if extracted is not None:
            rows.append(extracted)
    return rows


def build_csv(workbook_path: Path, csv_path: Path):
    """Create a Wealthfolio-friendly CSV from an XTB workbook export."""
    account_currency = detect_account_currency(workbook_path)
    cash_df = load_cash_operations(workbook_path)
    closed_df = load_closed_positions(workbook_path)

    rows = build_cash_rows(cash_df, account_currency)
    rows.extend(build_closed_position_rows(closed_df, account_currency))

    output = pd.DataFrame(
        rows,
        columns=[
            "date",
            "account",
            "symbol",
            "activityType",
            "quantity",
            "unitPrice",
            "amount",
            "currency",
            "fee",
            "comment",
        ],
    )
    output = output.sort_values(
        ["date", "symbol", "activityType", "comment"],
        kind="stable",
    )
    output.to_csv(csv_path, index=False, encoding="utf-8")

    print(f"Detected account currency: {account_currency}")
    print(f"Saved: {csv_path.name}")
    print(f"Rows: {len(output)}")


if __name__ == "__main__":
    base_dir = Path(__file__).resolve().parent
    selected_workbook = choose_input_xlsx()
    output_path = make_output_path(base_dir)
    build_csv(selected_workbook, output_path)
