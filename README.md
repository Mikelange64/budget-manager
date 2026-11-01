# Budget

A small CLI that reads a BMO bank statement CSV, categorizes each expense via a rules file, and filters/reports on where your money goes.

## Usage

```sh
python3 budget.py statement.csv
python3 budget.py statement.csv --start 2026-09-01 --end 2026-09-30 --category food --above 20
python3 budget.py statement.csv --summary
python3 budget.py statement.csv --merchants all
python3 budget.py statement.csv --report report.xlsx
```

Run `python3 budget.py --help` for all options.

## Files

- `budget.py` — the CLI (standard library only, except optional `openpyxl` for `--report`).
- `categories.json` — merchant → category rules. Edit this to grow the mapping; unmatched rows fall into `uncategorized`.

## Excel report

`--report output.xlsx` needs `openpyxl`:

```sh
pip install openpyxl
```
