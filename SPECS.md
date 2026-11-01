# Task: Bank statement categorizer CLI

There's a bank statement export in this folder: `statement.csv`. Build a Python CLI that I can point at **any** statement in this same format and that categorizes, filters and reports on my expenses.

## Before writing code
1. Open `statement.csv` and study its structure: columns, date format, how amounts and debits/credits are represented, and how the merchant/payee appears. I believe the merchant name comes right after `[PR]` in the description field, so it can likely be split from there. Confirm this, and tell me about any rows that don't follow the pattern (transfers, deposits, etc.).
2. List every unique merchant/payee you find, normalized to lowercase, with how many times each appears and the total spent. These are basically the places I go every month.
3. Propose a categorization for each one and show me before building.

## Constraints
- Core logic uses the standard library: `csv`, `argparse` (plus `datetime`, `json`, `difflib`, etc. as needed).
- The Excel report uses `openpyxl` (add it with `py add openpyxl`). Keep that import inside the report code so the rest of the CLI works without it.
- Script name: `budget.py` (or suggest something better).
- Merchant matching is case-insensitive (normalize everything to lowercase).
- Keep the merchant → category mapping in a separate, easy-to-edit file (e.g. `categories.json`) rather than hard-coded, so I can grow it as I download more statements.
- Anything that doesn't match goes into an `uncategorized` bucket so nothing silently disappears.

## Categories
| Category | Includes | Known merchants / rules |
|---|---|---|
| **groceries** | grocery stores | costco, adonis, maxi, avril |
| **bills** | rent, phone, internet | fizz (wifi); rent = transfer to Karine Benarous of $1625 **or** to Angeline of $812.50 |
| **savings & insurance** | insurance, monthly investments, savings transfers (suggest a better name if you have one) | build from the statement |
| **food** | restaurants, fast food, takeout | e.g. poulet rouge, mcdonalds; build the rest from the statement |
| **lifestyle** | personal care, entertainment, etc. | barber = transfer to Marvin Barbe (match however he appears in the file) |

### Household view
A `--household` flag that shows only shared household costs:
- groceries **excluding** avril
- rent
- home wifi only: the Fizz charge around **$44**, not the ~$30 one

## CLI
`python3 budget.py --help` should list everything. Filters must be combinable, e.g.:

python3 budget.py statement.csv --start 2026-09-01 --end 2026-09-30 --category food --above 20

Suggested arguments (improve or add where it makes sense):
- `file` (positional): path to the CSV
- `--start YYYY-MM-DD`, `--end YYYY-MM-DD`: date range (inclusive)
- `--above AMOUNT`, `--below AMOUNT`: amount filters (accept decimals)
- `--category NAME` (repeatable): filter by one or more categories
- `--household`: household view described above
- `--summary`: totals per category instead of listing transactions
- `--total`: print only the grand total of the filtered results
- `--uncategorized`: show only unmatched transactions (useful for growing the mapping)
- `--merchants [all | NAME ...]`: see the Merchants view below
- `--report OUTPUT.xlsx`: see the Excel report below

Default output is a readable aligned table in the terminal with date, merchant, category, amount, and a total at the bottom.

## Merchants view
A way to see where my money goes, by place. Three modes:

1. `--merchants` (no value): just the list of unique merchants I spent money at, alphabetical, one per line. Names only.
2. `--merchants all`: every merchant, each followed by its individual expenses, indented, with a subtotal.
3. `--merchants costco adonis` (also accept comma-separated: `--merchants costco,adonis`): same grouped output, but only for the merchants I name.

Validation: every merchant I name must exist in the filtered data (case-insensitive). For any that don't, print a clear message like `Invalid merchant: "cosco". Did you mean: costco?` (use `difflib` for suggestions). Still show results for the valid ones; exit with an error code only if none are valid.

It must respect the other filters (`--start`, `--end`, `--category`, `--above`, `--below`, `--household`), so I can see e.g. every merchant for just October.

Grouped output format:

costco                          3 transactions   $412.87
    2026-10-03    $156.20
    2026-10-14    $98.45
    2026-10-27    $158.22

adonis                          2 transactions   $87.10
    2026-10-08    $41.30
    2026-10-21    $45.80

Grand total                                      $499.97

## Excel report
`--report report.xlsx` generates a monthly spending report, respecting the same filters. It should include:
- **Summary sheet**: a table of category totals per month (months as columns, categories as rows, with totals), plus a bar chart comparing categories across months and a pie chart of the overall category breakdown.
- **One sheet per month**: that month's transactions (date, merchant, category, amount), category subtotals, and a pie chart of where the money went that month.
- **Merchants sheet**: total spent and number of transactions per merchant, sorted highest first, with a bar chart of the top 10.
- **Household sheet**: the household subset with monthly totals.
- Native Excel charts (`openpyxl.chart`), currency formatting, bold headers, frozen header rows, sensible column widths. It should look clean when opened.

## After it's built, let's discuss
Whether categorization should stay rule-based (the JSON mapping) or use my DeepSeek API key for semantic categorization of unknown merchants. My preference is to see if a well-built mapping covers my spending without needing an AI agent. I can download more statements to extend the list. Give me your honest recommendation, and if AI makes sense, suggest a hybrid (rules first, AI only for `uncategorized`).
