#!/usr/bin/env python3
"""Truemed -> accountant xlsx, Lithuania time.

Builds the monthly Truemed report for the accountant from:
  1. the Truemed "payments and refunds" CSV (app.truemed.com, dates are UTC)
  2. the complete Shopify orders export for the month being closed (dates +0300)

Logic (full explanation in the handover doc "Truemed Report Pipeline"):
  - Charges are matched to Shopify orders via Shopify ID == Payment Reference
    (split payments checked against "Payment References" too). Matched rows get
    the exact Shopify "Paid at" timestamp, which is Lithuania time.
  - An unmatched charge dated on the LAST UTC day of the orders-export month is
    provably next-month in LT: the export is complete, so absence means paid
    after month end; the UTC date caps it before 03:00 LT on day 1. It is
    shifted to the 1st, highlighted yellow, and noted.
  - Any other unmatched charge dated INSIDE the orders-export month means the
    orders export is incomplete: the script prints a WARNING and marks the row
    "PATIKRINTI" instead of guessing.
  - Unmatched charges dated after the export month (next month's orders) keep
    the UTC date; the month assignment is still exact.
  - Refund rows carry the refund date (UTC, no time available); kept with note.
  - One checkout can produce two Charge rows under one Shopify ID (small
    residual + main amount). Both are real; both are kept.

Usage:
  python3 truemed_report.py --truemed <truemed.csv> --orders <orders_export.csv> [--out report.xlsx]

Requires: Python 3.9+, openpyxl. Formulas recalculate when opened in Excel.
"""
import argparse
import csv
import sys
from collections import Counter
from datetime import date, timedelta
from decimal import Decimal
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

TYPE_LT = {"Charge": "Mokėjimas", "Full Refund": "Grąžinimas (pilnas)",
           "Partial Refund": "Grąžinimas (dalinis)"}
HDR = Font(name="Arial", bold=True, size=10, color="FFFFFF")
HDR_FILL = PatternFill("solid", fgColor="1F5B33")
BODY = Font(name="Arial", size=10)
NOTE_F = Font(name="Arial", size=9, italic=True, color="7F7F7F")
YEL = PatternFill("solid", fgColor="FFF2CC")
AMT = "#,##0.00;[Red]-#,##0.00"


def load_orders(path):
    """Payment reference -> (order name, 'YYYY-MM-DD HH:MM' LT). Also the export's month."""
    sh, months = {}, Counter()
    with open(path, encoding="utf-8-sig", newline="") as f:
        for r in csv.DictReader(f):
            pa = (r.get("Paid at") or "").strip()
            if not pa:
                continue
            months[pa[:7]] += 1
            refs = [(r.get("Payment Reference") or "").strip()]
            refs += [x.strip() for x in (r.get("Payment References") or "").split("+")]
            for ref in refs:
                if ref:
                    sh[ref] = (r["Name"], pa[:16])
    if not sh:
        raise ValueError("orders export has no rows with 'Paid at' - wrong file?")
    cover = months.most_common(1)[0][0]
    return sh, cover


def month_last_day(ym):
    y, m = int(ym[:4]), int(ym[5:7])
    nxt = date(y + (m == 12), m % 12 + 1, 1)
    return (nxt - timedelta(days=1)).isoformat()


def load_truemed(path, sh, cover):
    boundary = month_last_day(cover)
    rows, warns = [], []
    with open(path, encoding="utf-8-sig", newline="") as f:
        for r in csv.DictReader(f):
            typ = r["Type"]
            ref = r["Shopify ID"]
            note, order_no = "", sh.get(ref, ("", ""))[0]
            if typ == "Charge" and ref in sh:
                dt = sh[ref][1]                       # exact LT timestamp
                eff = dt[:10]
            else:
                dt = eff = r["Charge Date"]           # UTC date only
                if typ != "Charge":
                    note = "Grąžinimo data UTC"
                elif r["Charge Date"] == boundary:
                    eff = dt = (date.fromisoformat(boundary) + timedelta(days=1)).isoformat()
                    note = (f"Truemed UTC data {boundary}; LT laiku {dt}, "
                            f"priskirta kitam mėnesiui")
                elif r["Charge Date"][:7] <= cover:
                    note = "PATIKRINTI: nerasta Shopify eksporte, data UTC"
                    warns.append(f"unmatched charge inside {cover}: "
                                 f"{r['Charge Date']} {r['Name']} {r['Order Total']}")
                else:
                    note = "UTC data (kito mėnesio užsakymas, be tikslaus LT laiko)"
            rows.append({"dt": dt, "eff": eff, "men": eff[:7], "tipas": TYPE_LT[typ],
                         "uzs": order_no, "pirk": r["Name"],
                         "suma": Decimal(r["Order Total"]),
                         "kom": Decimal(r["Fee"]) if r["Fee"] else Decimal(0),
                         "payd": r["Payout Date"], "payid": r["Payout ID"], "note": note})
    rows.sort(key=lambda x: (x["eff"], x["dt"]))
    return rows, warns


def write_workbook(rows, out_path):
    months = sorted({r["men"] for r in rows})
    wb = Workbook()

    # ---------- Operacijos ----------
    ws = wb.active
    ws.title = "Operacijos"
    ws["A1"] = "Guard Blinds — Truemed operacijų ataskaita (mokėjimai, komisiniai, grąžinimai)"
    ws["A1"].font = Font(name="Arial", bold=True, size=12)
    ws["A2"] = (f"Laikotarpis: {rows[0]['eff']} iki {rows[-1]['eff']}. Datos Lietuvos laiku, "
                "konvertuotos iš Truemed UTC pagal Shopify apmokėjimo laiką. Eilutėse be Shopify "
                "atitikmens palikta UTC data (žr. pastabą). Komisinis nurodytas prie kiekvienos "
                "operacijos; grąžinimo atveju Truemed grąžina dalį komisinio (teigiama suma).")
    ws["A2"].font = NOTE_F
    cols = [("Eil. Nr.", 8), ("Data (LT)", 17), ("Mėnuo", 9), ("Tipas", 19),
            ("Užsakymo Nr.", 13), ("Pirkėjas", 26), ("Suma (USD)", 13),
            ("Truemed komisinis (USD)", 21), ("Neto (USD)", 13), ("Išmokos data", 13),
            ("Payout ID", 30), ("Pastaba", 52)]
    hr = 4
    for i, (name, w) in enumerate(cols, 1):
        c = ws.cell(row=hr, column=i, value=name)
        c.font, c.fill = HDR, HDR_FILL
        ws.column_dimensions[get_column_letter(i)].width = w
    for n, r in enumerate(rows, 1):
        rn = hr + n
        vals = [n, r["dt"], r["men"], r["tipas"], r["uzs"], r["pirk"], float(r["suma"]),
                float(r["kom"]), f"=G{rn}+H{rn}", r["payd"], r["payid"], r["note"]]
        for col, v in enumerate(vals, 1):
            c = ws.cell(row=rn, column=col, value=v)
            c.font = BODY
            if col in (7, 8, 9):
                c.number_format = AMT
        if "priskirta kitam mėnesiui" in r["note"] or "PATIKRINTI" in r["note"]:
            for col in range(1, 13):
                ws.cell(row=rn, column=col).fill = YEL
    last = hr + len(rows)
    ws.freeze_panes = f"A{hr + 1}"
    ws.auto_filter.ref = f"A{hr}:L{last}"

    # ---------- Suvestinė ----------
    s = wb.create_sheet("Suvestinė")
    s["A1"] = "Truemed suvestinė pagal mėnesį (LT laiku)"
    s["A1"].font = Font(name="Arial", bold=True, size=12)
    s["A2"] = ("Sumos skaičiuojamos formulėmis iš lapo 'Operacijos'. Uždaryto mėnesio bruto turi "
               "sutapti su Shopify 'Payments by type' ataskaitos Truemed eilute centas į centą. "
               "Išmokos per Stripe tiesiai į banko sąskaitą per ~2 d. d.; užšaldytų lėšų nėra.")
    s["A2"].font = NOTE_F
    r0 = 4
    for i, h in enumerate(["Rodiklis"] + months, 1):
        c = s.cell(row=r0, column=i, value=h)
        c.font, c.fill = HDR, HDR_FILL
    O, L = "Operacijos", last
    metrics = [
        ("Mokėjimų skaičius",
         '=COUNTIFS({O}!$C$5:$C${L},{m},{O}!$D$5:$D${L},"Mokėjimas")', "0"),
        ("Pardavimai bruto (USD)",
         '=SUMIFS({O}!$G$5:$G${L},{O}!$C$5:$C${L},{m},{O}!$D$5:$D${L},"Mokėjimas")', AMT),
        ("Truemed komisiniai (USD)",
         '=SUMIFS({O}!$H$5:$H${L},{O}!$C$5:$C${L},{m},{O}!$D$5:$D${L},"Mokėjimas")', AMT),
        ("Grąžinimai (USD)",
         '=SUMIFS({O}!$G$5:$G${L},{O}!$C$5:$C${L},{m},{O}!$D$5:$D${L},"Grąžinimas*")', AMT),
        ("Komisinio grąžinimas (USD)",
         '=SUMIFS({O}!$H$5:$H${L},{O}!$C$5:$C${L},{m},{O}!$D$5:$D${L},"Grąžinimas*")', AMT),
        ("Neto (USD)", '=SUMIFS({O}!$I$5:$I${L},{O}!$C$5:$C${L},{m})', AMT),
    ]
    for j, (label, f, fmt) in enumerate(metrics, 1):
        rn = r0 + j
        s.cell(row=rn, column=1, value=label).font = BODY
        for k, m in enumerate(months, 2):
            c = s.cell(row=rn, column=k, value=f.format(O=O, L=L, m=f'"{m}"'))
            c.font = BODY
            c.number_format = fmt
    s.column_dimensions["A"].width = 30
    for k in range(len(months)):
        s.column_dimensions[get_column_letter(2 + k)].width = 20

    # ---------- Išmokos ----------
    po = wb.create_sheet("Išmokos")
    po["A1"] = "Truemed išmokos į banko sąskaitą"
    po["A1"].font = Font(name="Arial", bold=True, size=12)
    po["A2"] = ("Suma skaičiuojama formule iš lapo 'Operacijos' (mokėjimai + grąžinimai + "
                "komisiniai pagal Payout ID) ir atitinka banko įplauką. Paskutinės išmokos dar "
                "gali būti pakeliui.")
    po["A2"].font = NOTE_F
    payouts = {}
    for r in rows:
        if r["payid"]:
            payouts.setdefault(r["payid"], r["payd"])
    pol = sorted(payouts.items(), key=lambda kv: kv[1])
    r0 = 4
    for i, h in enumerate(["Išmokos data", "Payout ID", "Suma (USD)"], 1):
        c = po.cell(row=r0, column=i, value=h)
        c.font, c.fill = HDR, HDR_FILL
    for j, (pid, pd) in enumerate(pol, 1):
        rn = r0 + j
        po.cell(row=rn, column=1, value=pd).font = BODY
        po.cell(row=rn, column=2, value=pid).font = BODY
        c = po.cell(row=rn, column=3,
                    value=f'=SUMIFS({O}!$I$5:$I${L},{O}!$K$5:$K${L},B{rn})')
        c.font = BODY
        c.number_format = AMT
    tr = r0 + len(pol) + 1
    po.cell(row=tr, column=2, value="Iš viso:").font = Font(name="Arial", bold=True, size=10)
    c = po.cell(row=tr, column=3, value=f"=SUM(C{r0 + 1}:C{r0 + len(pol)})")
    c.font = Font(name="Arial", bold=True, size=10)
    c.number_format = AMT
    for col, w in (("A", 14), ("B", 32), ("C", 14)):
        po.column_dimensions[col].width = w

    wb.save(out_path)
    return len(pol)


def run(truemed_csv, orders_csv, out_dir=".", out_path=None):
    """Programmatic entry point used by the web UI.

    Returns (output_path, warnings). Mirrors main() but without argparse or
    console output. `warnings` is non-empty when the orders export looks
    incomplete - those rows are marked PATIKRINTI and the file must not be
    sent until Shopify is re-exported.
    """
    sh, cover = load_orders(orders_csv)
    rows, warns = load_truemed(truemed_csv, sh, cover)
    if not rows:
        raise ValueError("no rows in the Truemed CSV")
    if out_path is None:
        name = f"Truemed_ataskaita_{rows[0]['eff']}_iki_{rows[-1]['eff']}.xlsx"
        out_path = str(Path(out_dir) / name)
    write_workbook(rows, out_path)
    return out_path, warns


def main():
    ap = argparse.ArgumentParser(description="Truemed -> accountant xlsx, Lithuania time")
    ap.add_argument("--truemed", required=True, help="Truemed payments and refunds CSV")
    ap.add_argument("--orders", required=True, help="complete Shopify orders export for the month")
    ap.add_argument("--out", default=None, help="output xlsx path")
    a = ap.parse_args()

    try:
        sh, cover = load_orders(a.orders)
    except ValueError as exc:
        sys.exit(str(exc))
    rows, warns = load_truemed(a.truemed, sh, cover)
    if not rows:
        sys.exit("no rows in the Truemed CSV")
    out = a.out or f"Truemed_ataskaita_{rows[0]['eff']}_iki_{rows[-1]['eff']}.xlsx"
    n_po = write_workbook(rows, out)
    print(f"orders export covers {cover} | {len(rows)} transactions | {n_po} payouts -> {out}")
    for w in warns:
        print("WARNING", w)
    if warns:
        print("WARNING rows are marked PATIKRINTI in the file - the orders export "
              "looks incomplete, re-export before sending.")


if __name__ == "__main__":
    main()
