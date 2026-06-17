# XTB to Wealthfolio

Convert XTB account exports into a format that is easier to inspect, debug, and adapt for Wealthfolio imports.

This project started as a practical tool for tracking down cash-balance mismatches between XTB exports and Wealthfolio portfolio imports. It focuses on making cash operations visible, especially around deposits, withdrawals, currency conversions, and broker-specific transaction comments.

## What it does

`XTB to Wealthfolio` helps turn messy XTB export data into something easier to reason about.

Typical use cases:
- Inspect cash operations from XTB `.xlsx` exports.
- Track running cash balances over time.
- Surface deposits, withdrawals, and transfer rows clearly.
- Debug EUR/USD conversion flows.
- Spot cases where importer logic may double-count or misclassify rows.
- Prepare data for custom Wealthfolio import workflows.

## Why this exists

XTB exports can be correct while an importer still shows the wrong cash balance. That usually happens because of mapping logic, especially around:
- currency conversion rows,
- transfer pairs,
- internal cash movements,
- embedded FX spread or rounding,
- and broker-specific comment formats.

This repo is meant to make those issues easier to diagnose.

## Input

The script is designed around XTB export files, especially account statements saved as Excel workbooks such as:

- `EUR_2376879_2006-01-01_2026-06-17.xlsx`
- `USD_52718245_2006-01-01_2026-06-17.xlsx`

It expects a `Cash Operations` sheet and reads rows from the exported workbook structure used by XTB.

## Output

Depending on the script version you use, the tool can:
- print transaction summaries to the terminal,
- calculate running cash balances,
- highlight conversion and transfer rows,
- and export cleaned or debug-friendly CSV files.

The goal is not to be a perfect universal importer. The goal is to give you a transparent starting point that you can modify for your own portfolio workflow.

## Supported transaction patterns

The debugging workflow in this repo is especially useful for:
- `Deposit`
- `Withdrawal`
- `Transfer`
- `Free funds interest`
- `Free funds interest tax`
- stock purchases and sells that indirectly affect available cash

It is particularly helpful when comments contain patterns like:
- `Currency conversion, EUR to USD`
- `Currency conversion, USD to EUR`
- `JP_MORGAN deposit`
- `PayPal deposit`
- `Withdrawal from ...`

## Requirements

- Python 3.10+
- A `requirements.txt` file is included in the repo

Install dependencies with:

```bash
pip install -r requirements.txt
```

## Usage

Place your XTB `.xlsx` exports in the same folder as the script, then run:

```bash
python xtbToWealthfolio.py
```

If the script is interactive, it will list the available `.xlsx` files and ask you to choose one.

Example:

```text
Available .xlsx files:
1. EUR_1234567_2006-01-01_2026-06-17.xlsx
2. USD_1234567_2006-01-01_2026-06-17.xlsx
Choose a file [1-2]: 1
```

## Example workflow

A common workflow looks like this:

1. Export account history from XTB using the new export as `.xlsx`.
2. Put the exported files next to the script.
3. Run the script and choose the account file.
4. Review the running cash balance and grouped transaction output.
5. Compare suspicious rows with what Wealthfolio imported.
6. Adjust mapping rules if needed.

## Notes on FX conversions

If your base account is in EUR and you buy USD stocks, the cash effect may come from a mix of:
- EUR cash movements,
- USD conversion rows,
- stock purchase rows,
- and FX spread or rounding embedded in the conversion rate.

That means the issue is often not a separate “fee” row. In many cases, the real problem is how an importer interprets transfer pairs or conversion rows.

## Customization

This project is intentionally script-first and easy to tweak.

You can adapt it to:
- produce Wealthfolio-ready CSV output,
- normalize transaction types,
- ignore certain broker comment patterns,
- add synthetic balancing rows,
- or split logic by account currency.

## Limitations

- It is tailored to XTB export structure.
- It may need changes if XTB changes column names or workbook layout.
- Broker comments are not fully standardized, so string matching may need adjustment.
- It is best treated as a practical utility, not an official accounting tool.