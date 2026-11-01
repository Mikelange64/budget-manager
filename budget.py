#!/usr/bin/env python3
"""Bank statement categorizer: categorize, filter and report on expenses.

Core logic is standard library only. openpyxl is imported lazily inside the
--report code, so everything else works without it.
"""
import argparse
import csv
import difflib
import json
import re
import sys
from collections import OrderedDict, defaultdict
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Sequence

Rule = dict[str, Any]  # one entry of "rules" in categories.json

UNCATEGORIZED = "uncategorized"
DEFAULT_MAPPING = Path(__file__).resolve().with_name("categories.json")
MERCHANT_WIDTH = 25  # fixed-width merchant field in [PR] descriptions
REGION_SUFFIXES = {"qc", "on", "bc", "ab", "ns", "nb", "mb", "sk", "pe", "nl",
                   "ca", "us", "usa", "deu", "gbr", "irl"}


# --------------------------------------------------------------------------
# Data model
# --------------------------------------------------------------------------
@dataclass
class Txn:
    date: date
    merchant: str       # display name (lowercase)
    raw: str            # merchant text as parsed from the file
    category: str
    amount: float       # positive = money spent
    household: bool = False
    label: str = ""
    desc: str = ""


def money(x: float) -> str:
    return f"${x:,.2f}"


# --------------------------------------------------------------------------
# Mapping (categories.json)
# --------------------------------------------------------------------------
def load_mapping(path: str | Path) -> tuple[list[str], list[Rule]]:
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
    except (OSError, json.JSONDecodeError) as e:
        sys.exit(f"error: cannot read mapping file {path}: {e}")
    rules = data.get("rules", [])
    for r in rules:
        m = r.get("match")
        r["_match"] = [s.lower() for s in ([m] if isinstance(m, str) else m or [])]
    return data.get("categories", []), rules


def classify(
    raw: str,
    amount: float,
    rules: list[Rule],
    day: date | None = None
) -> tuple[str, str | None, bool, str, bool]:
    """First matching rule wins. Returns (category, name, household, label, exclude)."""
    for r in rules:
        if not any(s in raw for s in r["_match"]):
            continue
        if "amount" in r and abs(amount - r["amount"]) > 0.005:
            continue
        if "date" in r and day is not None and day.isoformat() != r["date"]:
            continue
        if "min" in r and amount < r["min"] - 0.005:
            continue
        if "max" in r and amount > r["max"] + 0.005:
            continue
        return (r.get("category", "").lower(), r.get("name", "").lower() or None,
                bool(r.get("household")), r.get("label", ""), bool(r.get("exclude")))
    return UNCATEGORIZED, None, False, "", False


# --------------------------------------------------------------------------
# Statement parsing
# --------------------------------------------------------------------------
def extract_merchant(desc: str) -> str:
    """Pull the merchant/payee out of a raw description. Returns lowercase text."""
    m = re.match(r"\[(\w+)\](.*)$", desc.rstrip(), re.S)
    if not m:
        return re.sub(r"\s+", " ", desc).strip().lower()
    tag, body = m.group(1).upper(), m.group(2)

    if tag == "PR":  # card purchase: fixed-width merchant field
        text = body[:MERCHANT_WIDTH]
    elif tag == "OP":  # online / recurring: "<KIND> 11OCT2026MERCHANT  ..."
        dm = re.match(r"\s*(?:RECURRING PYMNT|ONLINE PURCHASE)\s+\d{1,2}[A-Za-z]{3}\d{4}(.*)", body)
        text = re.split(r"\s{2,}", (dm.group(1) if dm else body).strip())[0]
        words = text.split()
        if len(words) > 1 and words[-1].lower() in REGION_SUFFIXES:
            text = " ".join(words[:-1])
    elif tag == "CW":  # interac: "INTERAC ETRNSFR SENT   PAYEE   REF"
        parts = re.split(r"\s{2,}", body.strip())
        text = parts[1] if len(parts) > 1 else parts[0]
    else:  # DS / DN / TF ...: keep the words, drop long reference numbers
        parts = [p for p in re.split(r"\s{2,}", body.strip()) if not p.replace(" ", "").isdigit()]
        text = " ".join(parts)
    return re.sub(r"\s+", " ", text).strip().lower()


def clean_name(raw: str) -> str:
    name = re.sub(r"\s*#\d+", "", raw)
    return re.sub(r"\s+", " ", name).strip() or raw


def read_statement(
    path: str | Path,
    rules: list[Rule]
) -> tuple[list[Txn], list[float], list[float]]:
    """Return (expenses, ignored_credits, excluded). Credits are not expenses;
    excluded rows matched an `exclude` rule (e.g. moves between own accounts)."""
    try:
        text = Path(path).read_text(encoding="utf-8-sig")
    except OSError as e:
        sys.exit(f"error: cannot read {path}: {e}")
    lines = text.splitlines()
    hdr = next((i for i, l in enumerate(lines)
                if "transaction type" in l.lower() and "description" in l.lower()), None)
    if hdr is None:
        sys.exit(f"error: {path}: no header row (expected 'Transaction Type', 'Description', ...)")
    rows = list(csv.reader(lines[hdr:]))
    cols = [c.strip().lower() for c in rows[0]]
    try:
        i_date = cols.index("date posted")
        i_amt = cols.index("transaction amount")
        i_desc = cols.index("description")
    except ValueError as e:
        sys.exit(f"error: {path}: missing column ({e})")

    expenses, ignored, excluded = [], [], []
    for n, row in enumerate(rows[1:], start=hdr + 2):
        if not any(c.strip() for c in row):
            continue
        try:
            d = datetime.strptime(row[i_date].strip(), "%Y%m%d").date()
            amt = float(row[i_amt])
        except (ValueError, IndexError):
            print(f"warning: {path}:{n}: unreadable row skipped", file=sys.stderr)
            continue
        desc = ",".join(row[i_desc:])
        if amt >= 0:
            ignored.append(amt)
            continue
        spent = round(-amt, 2)
        raw = extract_merchant(desc)
        cat, name, hh, label, skip = classify(raw, spent, rules, d)
        if skip:
            excluded.append(spent)
            continue
        expenses.append(Txn(d, name or clean_name(raw), raw, cat, spent, hh, label or cat, desc.strip()))
    return expenses, ignored, excluded


# --------------------------------------------------------------------------
# Filtering
# --------------------------------------------------------------------------
def apply_filters(txns: list[Txn], a: argparse.Namespace) -> list[Txn]:
    cats = {c.lower() for c in (a.category or [])}
    out = []
    for t in txns:
        if a.start and t.date < a.start:
            continue
        if a.end and t.date > a.end:
            continue
        if a.above is not None and not t.amount > a.above:
            continue
        if a.below is not None and not t.amount < a.below:
            continue
        if cats and t.category not in cats:
            continue
        if a.household and not t.household:
            continue
        if a.uncategorized and t.category != UNCATEGORIZED:
            continue
        out.append(t)
    return sorted(out, key=lambda t: t.date)


# --------------------------------------------------------------------------
# Terminal output
# --------------------------------------------------------------------------
def print_table(txns: list[Txn]) -> None:
    w_m = max([len("Merchant")] + [len(t.merchant) for t in txns])
    w_c = max([len("Category")] + [len(t.category) for t in txns])
    w_a = max([len("Amount")] + [len(money(t.amount)) for t in txns] + [len(money(sum(t.amount for t in txns)))])
    head = f"{'Date':<10}  {'Merchant':<{w_m}}  {'Category':<{w_c}}  {'Amount':>{w_a}}"
    print(head)
    print("-" * len(head))
    for t in txns:
        print(f"{t.date.isoformat():<10}  {t.merchant:<{w_m}}  {t.category:<{w_c}}  {money(t.amount):>{w_a}}")
    print("-" * len(head))
    print(f"{'Total':<{10 + 2 + w_m + 2 + w_c}}  {money(sum(t.amount for t in txns)):>{w_a + 2}}")


def category_totals(
    txns: list[Txn],
    order: Sequence[str]
) -> list[tuple[str, int, float]]:
    tot: defaultdict[str, list[float]] = defaultdict(lambda: [0, 0.0])
    for t in txns:
        tot[t.category][0] += 1
        tot[t.category][1] += t.amount
    cats = [c for c in order if c in tot] + [c for c in tot if c not in order]
    return [(c, int(tot[c][0]), round(tot[c][1], 2)) for c in cats]


def print_summary(txns: list[Txn], order: Sequence[str]) -> None:
    rows = category_totals(txns, order)
    grand = sum(r[2] for r in rows)
    w = max([len("Category")] + [len(r[0]) for r in rows])
    print(f"{'Category':<{w}}  {'Count':>5}  {'Total':>12}  {'Share':>6}")
    print("-" * (w + 29))
    for c, n, s in sorted(rows, key=lambda r: -r[2]):
        print(f"{c:<{w}}  {n:>5}  {money(s):>12}  {s / grand * 100 if grand else 0:>5.1f}%")
    print("-" * (w + 29))
    print(f"{'Total':<{w}}  {sum(r[1] for r in rows):>5}  {money(grand):>12}")


def group_by_merchant(txns: list[Txn]) -> OrderedDict[str, list[Txn]]:
    g: OrderedDict[str, list[Txn]] = OrderedDict()
    for t in txns:
        g.setdefault(t.merchant, []).append(t)
    return g


def suggest(name: str, names: Sequence[str]) -> list[str]:
    subs = [n for n in names if name in n or n in name]
    close = difflib.get_close_matches(name, names, n=3, cutoff=0.6)
    return list(OrderedDict.fromkeys(subs + close))[:3]


def print_merchants(txns: list[Txn], requested: list[str]) -> int:
    groups = group_by_merchant(txns)
    if requested == []:  # names only
        for name in sorted(groups):
            print(name)
        return 0
    if requested == ["all"]:
        chosen = list(groups)
    else:
        names = list(groups)
        chosen = []
        for r in requested:
            r = r.lower()
            if r in groups:
                if r not in chosen:
                    chosen.append(r)
            else:
                sug = suggest(r, names)
                tip = f' Did you mean: {", ".join(sug)}?' if sug else ""
                print(f'Invalid merchant: "{r}".{tip}', file=sys.stderr)
        if not chosen:
            print("error: none of the requested merchants were found", file=sys.stderr)
            return 2
    chosen.sort(key=lambda n: -sum(t.amount for t in groups[n]))
    grand = 0.0
    for name in chosen:
        items = groups[name]
        sub = sum(t.amount for t in items)
        grand += sub
        label = f"{len(items)} transaction{'s' if len(items) != 1 else ''}"
        print(f"{name:<32}{label:<16}{money(sub):>10}")
        for t in items:
            print(f"    {t.date.isoformat()}".ljust(48) + f"{money(t.amount):>10}")
        print()
    print(f"{'Grand total':<48}{money(grand):>10}")
    return 0


# --------------------------------------------------------------------------
# Excel report
# --------------------------------------------------------------------------
def write_report(txns: list[Txn], order: Sequence[str], path: str | Path) -> None:
    try:
        from openpyxl import Workbook
        from openpyxl.chart import BarChart, PieChart, Reference
        from openpyxl.styles import Alignment, Font, PatternFill
        from openpyxl.utils import get_column_letter
    except ImportError:
        sys.exit("error: --report needs openpyxl (e.g. `uv pip install openpyxl`)")

    CUR = '"$"#,##0.00'
    HEAD_FILL = PatternFill("solid", fgColor="DDE7F3")
    bold = Font(bold=True)

    def head(ws: Any, row: int, col: int, labels: Sequence[str]) -> None:
        for i, text in enumerate(labels):
            c = ws.cell(row=row, column=col + i, value=text)
            c.font, c.fill = bold, HEAD_FILL
            c.alignment = Alignment(horizontal="center")

    def widths(ws: Any, spec: dict[str, float]) -> None:
        for col, w in spec.items():
            ws.column_dimensions[col].width = w

    def style_chart(ch: Any, title: str, w: float = 17, h: float = 9) -> Any:
        ch.title, ch.width, ch.height = title, w, h
        if hasattr(ch, "x_axis"):
            ch.x_axis.delete = False
            ch.y_axis.delete = False
        return ch

    def pie(
        ws: Any,
        title: str,
        cat_col: int,
        val_col: int,
        r1: int,
        r2: int,
        anchor: str
    ) -> None:
        ch = style_chart(PieChart(), title, 12, 9)
        ch.add_data(Reference(ws, min_col=val_col, min_row=r1 - 1, max_row=r2), titles_from_data=True)
        ch.set_categories(Reference(ws, min_col=cat_col, min_row=r1, max_row=r2))
        ws.add_chart(ch, anchor)

    months = sorted({t.date.strftime("%Y-%m") for t in txns})
    wb = Workbook()

    # ---- Summary: categories x months
    ws = wb.active
    ws.title = "Summary"
    cats = [c for c in order if any(t.category == c for t in txns)]
    cats += sorted({t.category for t in txns} - set(cats))
    grid = {(t.category, t.date.strftime("%Y-%m")): 0.0 for t in txns}
    for t in txns:
        grid[(t.category, t.date.strftime("%Y-%m"))] += t.amount
    head(ws, 1, 1, ["Category"] + months + ["Total"])
    for r, c in enumerate(cats, start=2):
        ws.cell(row=r, column=1, value=c)
        for j, m in enumerate(months, start=2):
            ws.cell(row=r, column=j, value=round(grid.get((c, m), 0.0), 2)).number_format = CUR
        ws.cell(row=r, column=len(months) + 2,
                value=round(sum(grid.get((c, m), 0.0) for m in months), 2)).number_format = CUR
    tr = len(cats) + 2
    ws.cell(row=tr, column=1, value="Total").font = bold
    for j in range(2, len(months) + 3):
        col = get_column_letter(j)
        cell = ws.cell(row=tr, column=j, value=f"=SUM({col}2:{col}{tr - 1})")
        cell.number_format, cell.font = CUR, bold
    ws.freeze_panes = "B2"
    widths(ws, {"A": 22, **{get_column_letter(j): 14 for j in range(2, len(months) + 3)}})
    bar = style_chart(BarChart(), "Spending by category and month", 20, 10)
    bar.type = "col"
    bar.add_data(Reference(ws, min_col=2, max_col=len(months) + 1, min_row=1, max_row=tr - 1), titles_from_data=True)
    bar.set_categories(Reference(ws, min_col=1, min_row=2, max_row=tr - 1))
    bar.y_axis.numFmt = '"$"#,##0'
    ws.add_chart(bar, f"A{tr + 3}")
    pie(ws, "Overall category breakdown", 1, len(months) + 2, 2, tr - 1, f"L{tr + 3}")

    # ---- One sheet per month
    for m in months:
        mt = [t for t in txns if t.date.strftime("%Y-%m") == m]
        ws = wb.create_sheet(m)
        head(ws, 1, 1, ["Date", "Merchant", "Category", "Amount"])
        for r, t in enumerate(mt, start=2):
            ws.cell(row=r, column=1, value=t.date).number_format = "yyyy-mm-dd"
            ws.cell(row=r, column=2, value=t.merchant)
            ws.cell(row=r, column=3, value=t.category)
            ws.cell(row=r, column=4, value=t.amount).number_format = CUR
        end = len(mt) + 1
        ws.cell(row=end + 1, column=1, value="Total").font = bold
        c = ws.cell(row=end + 1, column=4, value=f"=SUM(D2:D{end})")
        c.number_format, c.font = CUR, bold
        head(ws, 1, 6, ["Category", "Subtotal"])
        sub = category_totals(mt, order)
        for r, (cat, _n, s) in enumerate(sub, start=2):
            ws.cell(row=r, column=6, value=cat)
            ws.cell(row=r, column=7, value=s).number_format = CUR
        ws.freeze_panes = "A2"
        widths(ws, {"A": 13, "B": 28, "C": 16, "D": 13, "E": 3, "F": 18, "G": 13})
        pie(ws, f"{m}: where the money went", 6, 7, 2, len(sub) + 1, f"F{len(sub) + 4}")

    # ---- Merchants
    ws = wb.create_sheet("Merchants")
    head(ws, 1, 1, ["Merchant", "Transactions", "Total"])
    groups = sorted(group_by_merchant(txns).items(), key=lambda kv: -sum(t.amount for t in kv[1]))
    for r, (name, items) in enumerate(groups, start=2):
        ws.cell(row=r, column=1, value=name)
        ws.cell(row=r, column=2, value=len(items))
        ws.cell(row=r, column=3, value=round(sum(t.amount for t in items), 2)).number_format = CUR
    ws.freeze_panes = "A2"
    widths(ws, {"A": 30, "B": 14, "C": 14})
    top = min(len(groups), 10) + 1
    ch = style_chart(BarChart(), "Top 10 merchants", 18, 10)
    ch.type = "bar"
    ch.legend = None
    ch.add_data(Reference(ws, min_col=3, min_row=1, max_row=top), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=2, max_row=top))
    ch.x_axis.scaling.orientation = "maxMin"  # biggest on top
    ch.y_axis.numFmt = '"$"#,##0'
    ws.add_chart(ch, "E2")

    # ---- Household (built from the household flag on the filtered rows)
    ws = wb.create_sheet("Household")
    hh = [t for t in txns if t.household]
    head(ws, 1, 1, ["Date", "Merchant", "Item", "Amount"])
    for r, t in enumerate(hh, start=2):
        ws.cell(row=r, column=1, value=t.date).number_format = "yyyy-mm-dd"
        ws.cell(row=r, column=2, value=t.merchant)
        ws.cell(row=r, column=3, value=t.label)
        ws.cell(row=r, column=4, value=t.amount).number_format = CUR
    end = len(hh) + 1
    ws.cell(row=end + 1, column=1, value="Total").font = bold
    c = ws.cell(row=end + 1, column=4, value=f"=SUM(D2:D{end})" if hh else 0)
    c.number_format, c.font = CUR, bold
    head(ws, 1, 6, ["Month", "Household total"])
    hm = [(m, round(sum(t.amount for t in hh if t.date.strftime("%Y-%m") == m), 2)) for m in months]
    for r, (m, s) in enumerate(hm, start=2):
        ws.cell(row=r, column=6, value=m)
        ws.cell(row=r, column=7, value=s).number_format = CUR
    ws.freeze_panes = "A2"
    widths(ws, {"A": 13, "B": 28, "C": 16, "D": 13, "E": 3, "F": 12, "G": 17})
    if hh:
        ch = style_chart(BarChart(), "Household cost per month", 14, 8)
        ch.type = "col"
        ch.legend = None
        ch.add_data(Reference(ws, min_col=7, min_row=1, max_row=len(hm) + 1), titles_from_data=True)
        ch.set_categories(Reference(ws, min_col=6, min_row=2, max_row=len(hm) + 1))
        ch.y_axis.numFmt = '"$"#,##0'
        ws.add_chart(ch, f"F{len(hm) + 4}")

    try:
        wb.save(path)
    except OSError as e:
        sys.exit(f"error: cannot write {path}: {e}")


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------
def iso_date(s: str) -> date:
    try:
        return datetime.strptime(s, "%Y-%m-%d").date()
    except ValueError:
        raise argparse.ArgumentTypeError(f"invalid date '{s}' (use YYYY-MM-DD)")


def recategorize(
    a: argparse.Namespace,
    order: Sequence[str],
    rules: list[Rule]
) -> None:
    """Insert an override rule at the top of categories.json (it beats every other rule)."""
    merchant, cat = a.recategorize[0].lower(), a.recategorize[1].lower()
    if cat != "exclude" and cat not in order:
        sys.exit(f'error: unknown category "{cat}". Known: {", ".join(order)}, exclude '
                 f'(add new ones to "categories" in the mapping file first)')
    txns = []
    for f in a.file:
        txns += read_statement(f, rules)[0]
    hits = [t for t in txns if t.merchant == merchant or merchant in t.raw]
    if a.on:
        hits = [t for t in hits if t.date == a.on]
    if a.amount is not None:
        hits = [t for t in hits if abs(t.amount - a.amount) < 0.005]
    if not hits:
        names = sorted({t.merchant for t in txns})
        sug = suggest(merchant, names)
        sys.exit(f'error: no transaction matches "{merchant}"' + (f" (did you mean: {', '.join(sug)}?)" if sug else ""))
    rule: Rule = {"match": sorted({t.raw for t in hits}) if len({t.raw for t in hits}) > 1 else hits[0].raw}
    if a.amount is not None:
        rule["amount"] = a.amount
    if a.on:
        rule["date"] = a.on.isoformat()
    if cat == "exclude":
        rule["exclude"] = True
    else:
        rule["category"] = cat
    if len({t.merchant for t in hits}) == 1:
        rule["name"] = hits[0].merchant
    path = Path(a.mapping)
    lines = path.read_text(encoding="utf-8").splitlines(keepends=True)
    i = next(n for n, l in enumerate(lines) if l.strip().startswith('"rules"'))
    lines.insert(i + 1, "    " + json.dumps(rule) + ",\n")
    path.write_text("".join(lines), encoding="utf-8")
    print(f"Rule added to {path.name}: {json.dumps(rule)}")
    for t in sorted(hits, key=lambda t: t.date):
        print(f"  {t.date}  {t.merchant:<24} {money(t.amount):>10}  {t.category} -> {cat}")


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="budget.py",
        description="Categorize, filter and report on bank statement expenses. "
                    "Credits (income, deposits) are ignored: only money spent is counted.",
        epilog="Example: budget.py statement.csv --start 2026-09-01 --end 2026-09-30 --category food --above 20\n"
               "Note: put the file(s) before --merchants, which accepts several names.",
        formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("file", nargs="+", help="statement CSV file(s)")
    p.add_argument("--start", type=iso_date, metavar="YYYY-MM-DD", help="first date, inclusive")
    p.add_argument("--end", type=iso_date, metavar="YYYY-MM-DD", help="last date, inclusive")
    p.add_argument("--above", type=float, metavar="AMOUNT", help="only amounts strictly above this")
    p.add_argument("--below", type=float, metavar="AMOUNT", help="only amounts strictly below this")
    p.add_argument("--category", action="append", metavar="NAME",
                   help="only this category (repeatable)")
    p.add_argument("--household", action="store_true",
                   help="shared household costs only (groceries except avril, rent, home wifi)")
    p.add_argument("--summary", action="store_true", help="totals per category instead of a listing")
    p.add_argument("--total", action="store_true", help="print only the grand total")
    p.add_argument("--uncategorized", action="store_true", help="only transactions with no matching rule")
    p.add_argument("--merchants", nargs="*", metavar="all|NAME",
                   help="no value: list merchants; 'all': every merchant with its expenses; "
                        "or names (space or comma separated) to show only those")
    p.add_argument("--report", metavar="OUTPUT.xlsx", help="write an Excel report (needs openpyxl)")
    p.add_argument("--recategorize", nargs=2, metavar=("MERCHANT", "CATEGORY"),
                   help="move a merchant to another category (CATEGORY may be 'exclude'). "
                        "Narrow it to one transaction with --on and/or --amount")
    p.add_argument("--on", type=iso_date, metavar="YYYY-MM-DD", help="with --recategorize: only that date")
    p.add_argument("--amount", type=float, metavar="AMOUNT", help="with --recategorize: only that amount")
    p.add_argument("--mapping", default=str(DEFAULT_MAPPING), metavar="FILE",
                   help="merchant -> category rules (default: categories.json next to this script)")
    return p


def main(argv: Sequence[str] | None = None) -> int:
    a = build_parser().parse_args(argv)
    order, rules = load_mapping(a.mapping)
    unknown = [c for c in (a.category or []) if c.lower() not in {o.lower() for o in order} | {UNCATEGORIZED}]
    for c in unknown:
        print(f'warning: unknown category "{c}" (known: {", ".join(order + [UNCATEGORIZED])})', file=sys.stderr)

    if a.recategorize:
        recategorize(a, order, rules)
        return 0

    txns, ignored, excluded = [], [], []
    for f in a.file:
        t, i, x = read_statement(f, rules)
        txns += t
        ignored += i
        excluded += x
    shown = apply_filters(txns, a)

    if a.merchants is not None:
        req = [n.strip().lower() for item in a.merchants for n in item.split(",") if n.strip()]
        if not shown:
            print("No transactions match the filters.", file=sys.stderr)
            return 1
        mode = ["all"] if "all" in req else req
        return print_merchants(shown, mode)

    if not shown:
        print("No transactions match the filters.", file=sys.stderr)
        return 1

    if a.report:
        write_report(shown, order + [UNCATEGORIZED], a.report)
        print(f"Report written to {a.report} ({len(shown)} transactions, {money(sum(t.amount for t in shown))})")
        return 0
    if a.total:
        print(money(sum(t.amount for t in shown)))
        return 0

    (print_summary(shown, order + [UNCATEGORIZED]) if a.summary else print_table(shown))
    notes = []
    if ignored:
        notes.append(f"{len(ignored)} credit(s) totalling {money(sum(ignored))} ignored: not expenses")
    if excluded:
        notes.append(f"{len(excluded)} row(s) totalling {money(sum(excluded))} excluded by rule (own-account moves)")
    if notes:
        print("\n(" + "; ".join(notes) + ")", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
