# excel-report-automation

[![CI](https://github.com/Aleksa5000/excel-report-automation/actions/workflows/ci.yml/badge.svg)](https://github.com/Aleksa5000/excel-report-automation/actions/workflows/ci.yml)

**Turn raw sales CSV/Excel exports into a formatted, client-ready Excel report with one command.**

## The problem it solves

Every week or month, someone exports sales data and spends hours in Excel: removing duplicates, fixing messy rows, building pivot tables, formatting columns, adding a chart, and emailing the result. It is slow, boring, and easy to get wrong.

This tool does the whole job in seconds, the same way every time. Bad rows are not silently lost: they go to a separate **Rejected rows** sheet with the reason, so nothing disappears and the numbers can be trusted.

## Features

- Reads one or more `.csv` / `.xlsx` files at once (columns: `date, region, product, units, unit_price`)
- Cleans the data: trims whitespace, drops exact duplicates, validates dates, quantities, prices, and required fields
- Invalid rows are moved to a **Rejected rows** sheet with a plain-English reason
- Builds KPIs, a monthly revenue-by-region pivot, a Top 10 products table, and a native (editable) Excel chart
- Professional formatting: bold headers, currency formats, sized columns, frozen header row, autofilter
- Optional `--email` to send the finished report via SMTP (skipped safely if not configured)

## Example

```console
$ python report.py sample_data/sales.csv -o report.xlsx
Read 510 rows from 1 file(s)
  3 duplicates removed, 6 rejected, 501 valid
  Total revenue:             569,053.23
  Total units sold:          5289
  Orders:                    501
  Average order value:       1,135.83
  Top product (by revenue):  Ergonomic Chair
  Top region (by revenue):   North
  Period:                    2025-01-01 to 2025-12-31
Report written to report.xlsx
```

### What the generated workbook contains

| Sheet | Contents |
|---|---|
| **Summary** | Total revenue, units sold, orders, average order value, top product, top region, period, rejected-row and duplicate counts |
| **Monthly revenue by region** | Pivot of revenue per month (rows) by region (columns) with totals, plus a native stacked column chart |
| **Top 10 products** | Ranked by revenue, with units sold |
| **Rejected rows** | Every row that failed validation, with a `reason` column (e.g. `invalid date`, `missing region`, `invalid units (must be a positive whole number)`) |

## Quickstart

```bash
git clone https://github.com/Aleksa5000/excel-report-automation.git
cd excel-report-automation
python -m pip install -r requirements.txt
python report.py sample_data/sales.csv -o report.xlsx
```

Pass several files to combine them: `python report.py jan.csv feb.xlsx -o q1.xlsx`.

The sample file (`sample_data/sales.csv`, ~500 rows with deliberate duplicates and bad rows) is generated deterministically by `sample_data/generate_sample.py`.

## Configuration

Emailing is optional. Copy `.env.example`, set these environment variables, then add `--email`:

| Variable | Meaning |
|---|---|
| `SMTP_HOST`, `SMTP_PORT` | Mail server (uses STARTTLS, e.g. port 587) |
| `SMTP_USER`, `SMTP_PASSWORD` | Login (also used as the sender address) |
| `REPORT_TO` | Recipient address |

If any variable is missing, the email step is skipped and the report is still written.

## Running tests

```bash
python -m pip install -r requirements.txt pytest
python -m pytest -q
```

Tests cover the cleaning rules, KPI math, Excel input, and that the output workbook has the expected sheets, formatting, and a chart. No network needed.

## Tech stack

Python 3.11+, pandas, openpyxl, standard-library `argparse` / `smtplib` / `email`, pytest, GitHub Actions.

## Need something similar?

Need a report, spreadsheet, or data-cleaning workflow automated? **Contact me on Upwork** (Aleksa Djorovic, Python automation).

## License

All rights reserved. This is a portfolio sample: you can read and run the code to evaluate my work, but you may not use it commercially, redistribute it or build on it without my written permission. See [LICENSE](LICENSE).
