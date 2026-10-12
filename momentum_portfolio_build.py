import csv
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter

CAPITAL_PER_STOCK = 100000
AS_OF_DATE = dt.date(2026, 10, 9)

HEADER_FILL = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
GREEN_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
RED_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
WHITE_FILL = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
BOLD = Font(bold=True)
THIN = Side(style="thin", color="B7B7B7")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
DATE_FMT = "dd-mm-yy"

# Sr, Company, Symbol, Rating, First Signal Date, Buy Date (or None), Buy Price (or None), Latest Price, Times Appeared
ROWS = [
    (1, "Prime Securities Limited", "PRIMESECU", "4\u2605", dt.date(2026,10,1), dt.date(2026,10,5), 300.75690815518607, 299.45, 2),
    (2, "Nuvama Wealth Management Limited", "NUVAMA", "4\u2605", dt.date(2026,10,5), dt.date(2026,10,6), 1852.406285634973, 1764.4, 1),
    (3, "Haleos Labs Limited", "HALEOSLABS", "4\u2605", dt.date(2026,10,5), dt.date(2026,10,6), 1581.0657289002556, 1563.7, 1),
    (4, "Lloyds Enterprises Limited", "LLOYDSENT", "4\u2605", dt.date(2026,10,5), dt.date(2026,10,6), 79.87788224188353, 73.73, 1),
    (5, "Eternal Limited", "ETERNAL", "5\u2605", dt.date(2026,10,5), dt.date(2026,10,6), 325.7758833187933, 323.45, 1),
    (6, "Neuland Laboratories Limited", "NEULANDLAB", "5\u2605", dt.date(2026,10,5), dt.date(2026,10,6), 21634.279035154417, 20890.0, 1),
    (7, "Radico Khaitan Limited", "RADICO", "5\u2605", dt.date(2026,10,5), dt.date(2026,10,6), 4439.430190431294, 4550.0, 1),
    (8, "Jamna Auto Industries Limited", "JAMNAAUTO", "4\u2605", dt.date(2026,10,6), dt.date(2026,10,7), 150.81910329439526, 142.67, 1),
    (9, "GE Power India Limited", "GVPIL", "4\u2605", dt.date(2026,10,6), dt.date(2026,10,7), 706.3605293178542, 656.65, 1),
    (10, "Omax Autos Limited", "OMAXAUTO", "4\u2605", dt.date(2026,10,6), dt.date(2026,10,7), 209.42647785282773, 200.45, 1),
    (11, "Black Box Limited", "BBOX", "4\u2605", dt.date(2026,10,6), dt.date(2026,10,7), 867.0813048124938, 892.1, 1),
    (12, "Vishnu Chemicals Limited", "VISHNU", "5\u2605", dt.date(2026,10,6), dt.date(2026,10,7), 686.1164222672483, 654.85, 1),
    (13, "Senores Pharmaceuticals Limited", "SENORES", "5\u2605", dt.date(2026,10,6), dt.date(2026,10,7), 1400.6967089171683, 1320.0, 1),
    (14, "United Foodbrands Limited", "UFBL", "5\u2605", dt.date(2026,10,6), dt.date(2026,10,7), 693.3128545174588, 697.1, 1),
    (15, "Ather Energy Limited", "ATHERENERG", "5\u2605", dt.date(2026,10,6), dt.date(2026,10,7), 1449.375537377983, 1434.4, 2),
    (16, "Ador Welding Limited", "ADOR", "5\u2605", dt.date(2026,10,6), dt.date(2026,10,7), 1548.0966805766313, 1479.6, 1),
    (17, "Mangal Credit and Fincorp Limited", "MANCREDIT", "4\u2605", dt.date(2026,10,7), dt.date(2026,10,8), 227.64327698824815, 227.2, 1),
    (18, "Ujaas Energy Limited", "UEL", "4\u2605", dt.date(2026,10,9), None, None, 214.94, 1),
    (19, "Mayur Uniquoters Ltd", "MAYURUNIQ", "4\u2605", dt.date(2026,10,9), None, None, 788.4, 1),
]

SIGNALS_LOG = [
    (dt.date(2026,10,1), "4\u2605", "PRIMESECU", "Prime Securities Limited"),
    (dt.date(2026,10,2), "4\u2605", "PRIMESECU", "Prime Securities Limited"),
    (dt.date(2026,10,5), "4\u2605", "NUVAMA", "Nuvama Wealth Management Limited"),
    (dt.date(2026,10,5), "4\u2605", "HALEOSLABS", "Haleos Labs Limited"),
    (dt.date(2026,10,5), "4\u2605", "LLOYDSENT", "Lloyds Enterprises Limited"),
    (dt.date(2026,10,5), "5\u2605", "ETERNAL", "Eternal Limited"),
    (dt.date(2026,10,5), "5\u2605", "NEULANDLAB", "Neuland Laboratories Limited"),
    (dt.date(2026,10,5), "5\u2605", "RADICO", "Radico Khaitan Limited"),
    (dt.date(2026,10,6), "4\u2605", "JAMNAAUTO", "Jamna Auto Industries Limited"),
    (dt.date(2026,10,6), "4\u2605", "GVPIL", "GE Power India Limited"),
    (dt.date(2026,10,6), "4\u2605", "OMAXAUTO", "Omax Autos Limited"),
    (dt.date(2026,10,6), "4\u2605", "BBOX", "Black Box Limited"),
    (dt.date(2026,10,6), "5\u2605", "VISHNU", "Vishnu Chemicals Limited"),
    (dt.date(2026,10,6), "5\u2605", "SENORES", "Senores Pharmaceuticals Limited"),
    (dt.date(2026,10,6), "5\u2605", "UFBL", "United Foodbrands Limited"),
    (dt.date(2026,10,6), "5\u2605", "ATHERENERG", "Ather Energy Limited"),
    (dt.date(2026,10,6), "5\u2605", "ADOR", "Ador Welding Limited"),
    (dt.date(2026,10,7), "4\u2605", "MANCREDIT", "Mangal Credit and Fincorp Limited"),
    (dt.date(2026,10,9), "4\u2605", "UEL", "Ujaas Energy Limited"),
    (dt.date(2026,10,9), "4\u2605", "MAYURUNIQ", "Mayur Uniquoters Ltd"),
    (dt.date(2026,10,9), "5\u2605", "ATHERENERG", "Ather Energy Limited"),
]

EXCLUDED = [
    ("BHARATBOND-APR30", "BHARATBOND-APR30", "5\u2605 Buys (RSI)", dt.date(2026,10,5),
     "Not in NSE official equity list - Bharat Bond ETF (debt ETF), not a company stock"),
    ("PHARMABEES", "Nippon India ETF Nifty Pharma", "5\u2605 Buys (RSI)", dt.date(2026,10,6),
     "In NSE official ETF list; symbol also contains 'BEES'"),
    ("CNXPHARMA", "CNXPHARMA", "5\u2605 Buys (RSI)", dt.date(2026,10,6),
     "Not in NSE official equity list - this is an index value (Nifty Pharma), not a tradable security"),
]


def write_csvs(scratch_dir):
    with open(f"{scratch_dir}/portfolio.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Sr No","Company","Symbol","Rating","Signal Date(s)","First Signal Date",
                    "Buy Date","Buy Price","Latest Price","Times Appeared"])
        for r in ROWS:
            sr, company, symbol, rating, first_sig, buy_date, buy_price, latest, times = r
            w.writerow([
                sr, company, symbol, rating,
                first_sig.strftime("%d%m%y"),
                first_sig.strftime("%Y-%m-%d"),
                buy_date.strftime("%Y-%m-%d") if buy_date else "",
                f"{buy_price:.6f}" if buy_price is not None else "",
                f"{latest:.4f}",
                times,
            ])

    with open(f"{scratch_dir}/signals_log.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Date","Rating","Symbol","Company"])
        for d, rating, symbol, company in SIGNALS_LOG:
            w.writerow([d.strftime("%Y-%m-%d"), rating, symbol, company])

    with open(f"{scratch_dir}/excluded.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Name/Symbol As Seen","Full Name","Source Screener","Email Date","Reason Excluded"])
        for name, full, src, d, reason in EXCLUDED:
            w.writerow([name, full, src, d.strftime("%Y-%m-%d"), reason])


def autofit(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w


def style_header_row(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = BOLD
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER


def build_workbook(scratch_dir):
    wb = Workbook()
    ws = wb.active
    ws.title = "Portfolio"
    ws.sheet_view.showGridLines = True

    n = len(ROWS)
    first_data_row = 8
    last_data_row = first_data_row + n - 1  # 26

    # ---- Top block: capital + as-of date ----
    ws["A1"] = "Capital per stock (Rs)"
    ws["B1"] = CAPITAL_PER_STOCK
    ws["A1"].font = BOLD
    ws["A3"] = "As-of date"
    ws["B3"] = AS_OF_DATE
    ws["B3"].number_format = DATE_FMT
    ws["A3"].font = BOLD

    # ---- Summary block (D1:G6) ----
    summary_labels = [
        ("D1", "Total Invested",        "E1", f"=SUM(J{first_data_row}:J{last_data_row})"),
        ("F1", "Current Value",         "G1", f"=SUM(L{first_data_row}:L{last_data_row})"),
        ("D2", "Total P&L (Rs)",        "E2", f"=SUM(M{first_data_row}:M{last_data_row})"),
        ("F2", "Total P&L %",           "G2", f"=IF(E1=0,\"\",G1/E1-1)"),
        ("D3", "Winners",               "E3", f"=COUNTIF(M{first_data_row}:M{last_data_row},\">0\")"),
        ("F3", "Losers",                "G3", f"=COUNTIF(M{first_data_row}:M{last_data_row},\"<0\")"),
        ("D4", "Win Rate",              "E4", f"=IF((E3+G3)=0,\"\",E3/(E3+G3))"),
        ("F4", "Avg Return / Stock",    "G4", f"=AVERAGE(N{first_data_row}:N{last_data_row})"),
        ("D5", "Best Stock",            "E5", f"=INDEX(B{first_data_row}:B{last_data_row},MATCH(MAX(N{first_data_row}:N{last_data_row}),N{first_data_row}:N{last_data_row},0))"),
        ("F5", "Worst Stock",           "G5", f"=INDEX(B{first_data_row}:B{last_data_row},MATCH(MIN(N{first_data_row}:N{last_data_row}),N{first_data_row}:N{last_data_row},0))"),
        ("D6", "Avg Return 4\u2605*",   "E6", f"=AVERAGEIF(D{first_data_row}:D{last_data_row},\"4\u2605*\",N{first_data_row}:N{last_data_row})"),
        ("F6", "Avg Return 5\u2605*",   "G6", f"=AVERAGEIF(D{first_data_row}:D{last_data_row},\"5\u2605*\",N{first_data_row}:N{last_data_row})"),
    ]
    for lbl_cell, lbl_text, val_cell, formula in summary_labels:
        ws[lbl_cell] = lbl_text
        ws[lbl_cell].font = BOLD
        ws[val_cell] = formula

    for cell in ["G2", "E4"]:
        ws[cell].number_format = "0.00%"
    for cell in ["E1", "G1", "E2"]:
        ws[cell].number_format = "#,##0"
    for cell in ["G4"]:
        ws[cell].number_format = "0.00%"
    for cell in ["E6", "G6"]:
        ws[cell].number_format = "0.00%"

    # ---- Header row (row 7) ----
    headers = ["Sr No","Company","Symbol","Rating","Signal Date(s)","First Signal Date",
               "Buy Date","Buy Price","Qty","Invested","Latest Price","Current Value",
               "P&L Rs","P&L %","Days Held","Status","Times Appeared"]
    for idx, h in enumerate(headers, start=1):
        ws.cell(row=7, column=idx, value=h)
    style_header_row(ws, 7, len(headers))
    ws.row_dimensions[7].height = 32

    # ---- Data rows ----
    for i, r in enumerate(ROWS):
        sr, company, symbol, rating, first_sig, buy_date, buy_price, latest, times = r
        row = first_data_row + i
        ws.cell(row=row, column=1, value=sr)
        ws.cell(row=row, column=2, value=company)
        ws.cell(row=row, column=3, value=symbol)
        ws.cell(row=row, column=4, value=rating)
        ws.cell(row=row, column=5, value=first_sig.strftime("%d%m%y"))
        ws.cell(row=row, column=6, value=first_sig)
        ws.cell(row=row, column=6).number_format = DATE_FMT
        if buy_date:
            ws.cell(row=row, column=7, value=buy_date)
            ws.cell(row=row, column=7).number_format = DATE_FMT
        else:
            ws.cell(row=row, column=7, value="")
        if buy_price is not None:
            ws.cell(row=row, column=8, value=round(buy_price, 2))
        else:
            ws.cell(row=row, column=8, value="")
        ws.cell(row=row, column=8).number_format = "#,##0.00"

        ws.cell(row=row, column=9, value=f"=IF(H{row}=\"\",\"\",ROUNDDOWN($B$1/H{row},0))")
        ws.cell(row=row, column=10, value=f"=IF(H{row}=\"\",\"\",H{row}*I{row})")
        ws.cell(row=row, column=10).number_format = "#,##0.00"

        ws.cell(row=row, column=11, value=round(latest, 2))
        ws.cell(row=row, column=11).number_format = "#,##0.00"

        ws.cell(row=row, column=12, value=f"=IF(OR(I{row}=\"\",K{row}=\"\"),\"\",K{row}*I{row})")
        ws.cell(row=row, column=12).number_format = "#,##0.00"
        ws.cell(row=row, column=13, value=f"=IF(L{row}=\"\",\"\",L{row}-J{row})")
        ws.cell(row=row, column=13).number_format = "#,##0.00"
        ws.cell(row=row, column=14, value=f"=IF(OR(H{row}=\"\",K{row}=\"\"),\"\",K{row}/H{row}-1)")
        ws.cell(row=row, column=14).number_format = "0.00%"
        ws.cell(row=row, column=15, value=f"=IF(G{row}=\"\",\"\",$B$3-G{row})")
        ws.cell(row=row, column=16, value=f"=IF(H{row}=\"\",\"Pending buy\",IF(N{row}>=0,\"In profit\",\"In loss\"))")
        ws.cell(row=row, column=17, value=times)

        for c in range(1, 18):
            ws.cell(row=row, column=c).border = BORDER
            if c not in (2,):
                ws.cell(row=row, column=c).alignment = Alignment(horizontal="center")

    # ---- Conditional formatting (profit/loss) on N column, applied across row ----
    cf_range = f"A{first_data_row}:Q{last_data_row}"
    ws.conditional_formatting.add(
        cf_range,
        FormulaRule(formula=[f"$N{first_data_row}>0"], fill=GREEN_FILL)
    )
    ws.conditional_formatting.add(
        cf_range,
        FormulaRule(formula=[f"$N{first_data_row}<0"], fill=RED_FILL)
    )

    # ---- Freeze header ----
    ws.freeze_panes = "A8"

    # ---- Column widths ----
    widths = {1:6, 2:30, 3:12, 4:8, 5:14, 6:14, 7:12, 8:12, 9:8, 10:13,
              11:12, 12:13, 13:12, 14:10, 15:10, 16:13, 17:10}
    autofit(ws, widths)

    # ---- White background explicitly on used range ----
    for row in ws.iter_rows(min_row=1, max_row=last_data_row, min_col=1, max_col=17):
        for cell in row:
            if cell.fill.start_color.rgb in (None, "00000000"):
                pass  # leave default (white) background

    # =========================================================
    # Signals Log sheet
    # =========================================================
    ws2 = wb.create_sheet("Signals Log")
    log_headers = ["Date", "Rating", "Symbol", "Company"]
    for idx, h in enumerate(log_headers, start=1):
        ws2.cell(row=1, column=idx, value=h)
    style_header_row(ws2, 1, len(log_headers))
    for i, (d, rating, symbol, company) in enumerate(SIGNALS_LOG, start=2):
        ws2.cell(row=i, column=1, value=d)
        ws2.cell(row=i, column=1).number_format = DATE_FMT
        ws2.cell(row=i, column=2, value=rating)
        ws2.cell(row=i, column=3, value=symbol)
        ws2.cell(row=i, column=4, value=company)
        for c in range(1, 5):
            ws2.cell(row=i, column=c).border = BORDER
            ws2.cell(row=i, column=c).alignment = Alignment(horizontal="center" if c != 4 else "left")
    ws2.freeze_panes = "A2"
    autofit(ws2, {1:12, 2:8, 3:14, 4:34})

    # =========================================================
    # Excluded sheet
    # =========================================================
    ws3 = wb.create_sheet("Excluded")
    exc_headers = ["Name/Symbol As Seen", "Full Name", "Source Screener", "Email Date", "Reason Excluded"]
    for idx, h in enumerate(exc_headers, start=1):
        ws3.cell(row=1, column=idx, value=h)
    style_header_row(ws3, 1, len(exc_headers))
    for i, (name, full, src, d, reason) in enumerate(EXCLUDED, start=2):
        ws3.cell(row=i, column=1, value=name)
        ws3.cell(row=i, column=2, value=full)
        ws3.cell(row=i, column=3, value=src)
        ws3.cell(row=i, column=4, value=d)
        ws3.cell(row=i, column=4).number_format = DATE_FMT
        ws3.cell(row=i, column=5, value=reason)
        for c in range(1, 6):
            ws3.cell(row=i, column=c).border = BORDER
            ws3.cell(row=i, column=c).alignment = Alignment(horizontal="center" if c in (1,3,4) else "left", wrap_text=(c==5))
    ws3.freeze_panes = "A2"
    autofit(ws3, {1:20, 2:30, 3:18, 4:14, 5:55})

    wb.save(f"{scratch_dir}/portfolio.xlsx")


if __name__ == "__main__":
    SCRATCH = r"C:\Users\BLUEOC~1\AppData\Local\Temp\claude\C--Users-BlueOcean-Ops-Team\507250cd-54ac-4c6a-941c-eb02119ac53a\scratchpad"
    write_csvs(SCRATCH)
    build_workbook(SCRATCH)
    print("Done")
