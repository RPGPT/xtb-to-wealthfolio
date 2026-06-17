import re
import secrets
from pathlib import Path

import pandas as pd


def clean_symbol(value):
    if pd.isna(value):
        return ""

    s = str(value).strip()

    if s.upper().endswith(".US"):
        return s[:-3]

    return s


def parse_trade_fields(comment, activity_type):
    if pd.isna(comment):
        return "", ""

    c = str(comment)

    if activity_type == "BUY":
        m = re.search(r"OPEN BUY\s+([0-9.]+)(?:/[0-9.]+)?\s*@\s*([0-9.]+)", c)
    elif activity_type == "SELL":
        m = re.search(r"CLOSE BUY\s+([0-9.]+)(?:/[0-9.]+)?\s*@\s*([0-9.]+)", c)
    else:
        m = None

    if not m:
        return "", ""

    return m.group(1), m.group(2)


def detect_account_currency(input_xlsx: Path):
    name = input_xlsx.name.upper()
    if "USD" in name:
        return "USD"
    if "EUR" in name:
        return "EUR"
    raise ValueError(
        f"Could not infer account currency from filename '{input_xlsx.name}'. "
        f"Expected filename to contain USD or EUR."
    )


def map_transfer_activity(comment, amount, account_currency):
    c = "" if pd.isna(comment) else str(comment).lower()
    amt = pd.to_numeric(amount, errors="coerce")

    if account_currency == "USD":
        if "eur to usd" in c:
            return "DEPOSIT"
        if "usd to eur" in c:
            return "WITHDRAWAL"

    if account_currency == "EUR":
        if "usd to eur" in c:
            return "DEPOSIT"
        if "eur to usd" in c:
            return "WITHDRAWAL"

    if pd.notna(amt):
        return "DEPOSIT" if amt >= 0 else "WITHDRAWAL"

    return None


def map_activity_type(raw_type, amount, comment, account_currency):
    t = str(raw_type).strip().lower()

    if t == "stock purchase":
        return "BUY"
    if t == "stock sell":
        return "SELL"
    if t == "dividend":
        return "DIVIDEND"
    if t in ("withholding tax", "free funds interest tax"):
        return "TAX"
    if t == "free funds interest":
        return "INTEREST"
    if t == "deposit":
        return "DEPOSIT"
    if "withdraw" in t:
        return "WITHDRAWAL"
    if t == "transfer":
        return map_transfer_activity(comment, amount, account_currency)

    return None


def choose_input_xlsx():
    script_dir = Path(__file__).resolve().parent
    xlsx_files = sorted(script_dir.glob("*.xlsx"))

    if not xlsx_files:
        raise FileNotFoundError("No .xlsx files found in the same folder as this script.")

    print("Available .xlsx files:")
    for i, file_path in enumerate(xlsx_files, start=1):
        print(f"{i}. {file_path.name}")

    while True:
        choice = input(f"Choose a file [1-{len(xlsx_files)}]: ").strip()
        try:
            idx = int(choice)
            if 1 <= idx <= len(xlsx_files):
                return xlsx_files[idx - 1]
        except ValueError:
            pass
        print("Invalid choice. Try again.")


def make_output_path(script_dir: Path):
    random_hash = secrets.token_hex(6)
    return script_dir / f"xtb_wealthfolio_{random_hash}.csv"


def find_col(df, names):
    lower = {str(c).strip().lower(): c for c in df.columns}
    for name in names:
        if name.lower() in lower:
            return lower[name.lower()]
    return None


def build_csv(input_xlsx: Path, output_csv: Path):
    account_currency = detect_account_currency(input_xlsx)

    df = pd.read_excel(input_xlsx, sheet_name="Cash Operations", header=4)
    df = df[df["Type"].notna()].copy()
    df["Time"] = pd.to_datetime(df["Time"], errors="coerce")
    df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce")
    df = df.sort_values("Time", kind="stable")

    rows = []

    for src_idx, (_, r) in enumerate(df.iterrows(), start=1):
        comment = "" if pd.isna(r.get("Comment")) else str(r["Comment"])
        amount = r.get("Amount", "")
        activity_type = map_activity_type(r["Type"], amount, comment, account_currency)
        if activity_type is None:
            continue

        symbol = clean_symbol(r.get("Ticker", ""))
        quantity = ""
        unit_price = ""

        if activity_type in ("BUY", "SELL"):
            quantity, unit_price = parse_trade_fields(comment, activity_type)

        unique_comment = f"{comment} | xtb_row={src_idx}"

        rows.append({
            "date": r["Time"].strftime("%Y-%m-%d %H:%M:%S") if pd.notna(r["Time"]) else "",
            "account": "XTB " + account_currency,
            "symbol": symbol,
            "activityType": activity_type,
            "quantity": quantity,
            "unitPrice": unit_price,
            "amount": amount,
            "currency": account_currency,
            "fee": 0,
            "comment": unique_comment
        })

    try:
        closed = pd.read_excel(input_xlsx, sheet_name="Closed Positions", header=4)
        closed = closed.copy()

        symbol_col = find_col(closed, ["Ticker"])
        qty_col = find_col(closed, ["Volume"])
        open_price_col = find_col(closed, ["Open Price"])
        close_price_col = find_col(closed, ["Close Price"])
        open_time_col = find_col(closed, ["Open Time (UTC)"])
        close_time_col = find_col(closed, ["Close Time (UTC)"])
        origin_close_col = find_col(closed, ["Close Origin"])
        position_id_col = find_col(closed, ["Position ID"])
        instrument_col = find_col(closed, ["Instrument"])

        if all([symbol_col, qty_col, close_time_col, origin_close_col]):
            for closed_idx, (_, r) in enumerate(closed.iterrows(), start=1):
                origin_close = "" if pd.isna(r.get(origin_close_col)) else str(r[origin_close_col]).strip()
                if origin_close.lower() != "correction":
                    continue

                symbol = clean_symbol(r.get(symbol_col, ""))
                quantity = "" if pd.isna(r.get(qty_col)) else str(r[qty_col]).replace(",", ".")
                close_dt = pd.to_datetime(r.get(close_time_col), errors="coerce")

                if not symbol or not quantity or pd.isna(close_dt):
                    continue

                close_price = ""
                if close_price_col and pd.notna(r.get(close_price_col)):
                    close_price = str(r[close_price_col]).replace(",", ".")
                elif open_price_col and pd.notna(r.get(open_price_col)):
                    close_price = str(r[open_price_col]).replace(",", ".")

                position_id = "" if not position_id_col or pd.isna(r.get(position_id_col)) else str(r[position_id_col]).strip()
                instrument_name = "" if not instrument_col or pd.isna(r.get(instrument_col)) else str(r[instrument_col]).strip()
                open_dt = "" if not open_time_col or pd.isna(r.get(open_time_col)) else str(r[open_time_col]).strip()

                rows.append({
                    "date": close_dt.strftime("%Y-%m-%d %H:%M:%S"),
                    "account": "XTB " + account_currency,
                    "symbol": symbol,
                    "activityType": "SELL",
                    "quantity": quantity,
                    "unitPrice": close_price,
                    "amount": "",
                    "currency": account_currency,
                    "fee": 0,
                    "comment": f"closed_position_correction position={position_id} open={open_dt} close={close_dt} name={instrument_name} | closed_row={closed_idx}"
                })

    except Exception:
        pass

    out = pd.DataFrame(rows, columns=[
        "date",
        "account",
        "symbol",
        "activityType",
        "quantity",
        "unitPrice",
        "amount",
        "currency",
        "fee",
        "comment"
    ])

    out = out.sort_values(["date", "symbol", "activityType", "comment"], kind="stable")
    out.to_csv(output_csv, index=False, encoding="utf-8")

    print(f"Detected account currency: {account_currency}")
    print(f"Saved: {output_csv.name}")
    print(f"Rows: {len(out)}")


if __name__ == "__main__":
    script_dir = Path(__file__).resolve().parent
    input_xlsx = choose_input_xlsx()
    output_csv = make_output_path(script_dir)
    build_csv(input_xlsx, output_csv)