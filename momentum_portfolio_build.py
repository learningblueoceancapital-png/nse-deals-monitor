#!/usr/bin/env python3
"""
Momentum Portfolio tracker — builds a dummy portfolio from the "4 Star Buy" and
"5 Star Buys (RSI)" chartink signal emails (chartink_4star_buy.py / chartink_screener.py
in this same repo), which log their scraped rows to signals_raw_log.csv.

Stateless-recompute design: signals_log.csv, excluded.csv and portfolio.csv are fully
rederived from signals_raw_log.csv + live NSE data on every run (previously-settled Buy
Prices are cached from the existing portfolio.csv rather than refetched, since a past
trading day's price never changes). This avoids incremental-patch bugs — re-running is
always safe and idempotent.

Modes:
  --mode silent   Mon-Thu ~7pm IST update. No email. Exits quietly outside Mon-Thu.
  --mode friday   Friday (or Thursday if Friday is an NSE holiday) bhavcopy poll +
                  weekly "Momentum Portfolio" email. Safe to invoke repeatedly every
                  10 minutes — idempotent via portfolio_state.json.

Env vars:
  GMAIL_USER / GMAIL_APP_PASSWORD   SMTP creds (same secrets as the sibling screeners)
  DRY_RUN                           "true" (default) restricts recipients to Kinjal only
"""

import os
import re
import csv
import sys
import json
import time
import ssl
import smtplib
import logging
import datetime
import zipfile
import io
import urllib.request
import urllib.error
from email.message import EmailMessage

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger("momentum_portfolio")

IST = datetime.timezone(datetime.timedelta(hours=5, minutes=30))

GMAIL_USER     = os.getenv("GMAIL_USER", "")
GMAIL_PASSWORD = os.getenv("GMAIL_APP_PASSWORD", "")
DRY_RUN        = os.getenv("DRY_RUN", "true").strip().lower() == "true"

PROD_RECIPIENTS = ["lavneesh@blueoceancapital.co.in", "kinjal.shah@blueoceancapital.co.in"]
DRY_RUN_RECIPIENTS = ["kinjal.shah@blueoceancapital.co.in"]

CAPITAL_PER_STOCK = 100000
STAR_LABEL = {"4star": "4★", "5star": "5★"}
EXCLUDE_KEYWORDS = ["ETF", "BEES", "GOLD", "SILVER", "LIQUID", "INDEX", "FUND", "FOF", "INVIT", "REIT"]

RAW_LOG_PATH      = "signals_raw_log.csv"
SIGNALS_LOG_PATH  = "signals_log.csv"
EXCLUDED_PATH     = "excluded.csv"
PORTFOLIO_PATH    = "portfolio.csv"
XLSX_PATH         = "portfolio.xlsx"
STATE_PATH        = "portfolio_state.json"
FRIDAY_LOG_PATH   = "friday_publish_log.csv"

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")


# ══════════════════════════════════════════════════════════════════════════════
# NSE data access
# ══════════════════════════════════════════════════════════════════════════════

def _http_get(url: str, timeout: int = 30) -> bytes:
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return resp.read()


def _http_head_ok(url: str, timeout: int = 20) -> bool:
    req = urllib.request.Request(url, headers={"User-Agent": UA}, method="HEAD")
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return resp.status == 200
    except urllib.error.HTTPError:
        return False
    except urllib.error.URLError:
        return False


def fetch_equity_list() -> dict:
    """symbol -> company name, from NSE's official equity list."""
    data = _http_get("https://archives.nseindia.com/content/equities/EQUITY_L.csv")
    text = data.decode("utf-8", errors="replace")
    reader = csv.DictReader(io.StringIO(text))
    out = {}
    for row in reader:
        sym = row.get("SYMBOL", "").strip()
        name = row.get("NAME OF COMPANY", "").strip()
        if sym:
            out[sym] = name
    log.info("Fetched NSE equity list: %d symbols", len(out))
    return out


def fetch_etf_set() -> set:
    data = _http_get("https://archives.nseindia.com/content/equities/eq_etfseclist.csv")
    text = data.decode("utf-8", errors="replace")
    reader = csv.DictReader(io.StringIO(text))
    out = set()
    for row in reader:
        sym = row.get("Symbol", "").strip()
        if sym:
            out.add(sym)
    log.info("Fetched NSE ETF list: %d symbols", len(out))
    return out


def fetch_nse_holidays() -> set:
    """Returns a set of datetime.date for NSE Capital Market (CM) segment trading holidays.
    NSE's holiday-master API has no year filter — it just returns whatever calendar year(s)
    it currently has published, so there is nothing to parametrize here."""
    try:
        # Warm up a session cookie — the holiday API 403s without one.
        req = urllib.request.Request("https://www.nseindia.com/all-reports",
                                      headers={"User-Agent": UA})
        opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor())
        opener.open(req, timeout=20)
        req2 = urllib.request.Request(
            "https://www.nseindia.com/api/holiday-master?type=trading",
            headers={"User-Agent": UA, "Accept": "application/json"},
        )
        data = opener.open(req2, timeout=20).read()
        parsed = json.loads(data.decode("utf-8"))
        holidays = set()
        for entry in parsed.get("CM", []):
            d = datetime.datetime.strptime(entry["tradingDate"], "%d-%b-%Y").date()
            holidays.add(d)
        log.info("Fetched %d NSE CM-segment holidays", len(holidays))
        return holidays
    except Exception as exc:
        log.warning("Could not fetch NSE holiday calendar (%s) — falling back to weekends-only", exc)
        return set()


_bhav_cache: dict = {}


def load_bhavcopy(d: datetime.date) -> dict | None:
    """Returns {symbol: row-dict} for the given trading date, or None if not published."""
    key = d.strftime("%Y%m%d")
    if key in _bhav_cache:
        return _bhav_cache[key]

    url = f"https://nsearchives.nseindia.com/content/cm/BhavCopy_NSE_CM_0_0_0_{key}_F_0000.csv.zip"
    try:
        raw = _http_get(url, timeout=30)
    except urllib.error.HTTPError as exc:
        if exc.code == 404:
            _bhav_cache[key] = None
            return None
        raise
    except urllib.error.URLError:
        _bhav_cache[key] = None
        return None

    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        csv_name = next(n for n in z.namelist() if n.endswith(".csv"))
        text = z.read(csv_name).decode("utf-8")

    out = {}
    reader = csv.DictReader(io.StringIO(text))
    for row in reader:
        if row.get("Sgmt") == "CM" and row.get("FinInstrmTp") == "STK" and row.get("SctySrs") in ("EQ", "BE"):
            sym = row["TckrSymb"]
            if sym not in out or (out[sym]["SctySrs"] != "EQ" and row["SctySrs"] == "EQ"):
                out[sym] = row
    _bhav_cache[key] = out
    log.info("Loaded bhavcopy %s: %d symbols", d.isoformat(), len(out))
    return out


def bhavcopy_exists(d: datetime.date) -> bool:
    key = d.strftime("%Y%m%d")
    if key in _bhav_cache:
        return _bhav_cache[key] is not None
    url = f"https://nsearchives.nseindia.com/content/cm/BhavCopy_NSE_CM_0_0_0_{key}_F_0000.csv.zip"
    return _http_head_ok(url)


def next_trading_day(d: datetime.date, holidays: set) -> datetime.date:
    nd = d + datetime.timedelta(days=1)
    while nd.weekday() >= 5 or nd in holidays:
        nd += datetime.timedelta(days=1)
    return nd


def latest_available_trading_date(today: datetime.date, holidays: set, max_lookback: int = 10) -> datetime.date | None:
    d = today
    for _ in range(max_lookback):
        if d.weekday() < 5 and d not in holidays and bhavcopy_exists(d):
            return d
        d -= datetime.timedelta(days=1)
    return None


def symbol_avg_price(bhav: dict, symbol: str) -> float | None:
    row = bhav.get(symbol)
    if not row:
        return None
    try:
        vol = float(row["TtlTradgVol"])
        val = float(row["TtlTrfVal"])
        return (val / vol) if vol else None
    except (KeyError, ValueError, ZeroDivisionError):
        return None


def symbol_close_price(bhav: dict, symbol: str) -> float | None:
    row = bhav.get(symbol)
    if not row:
        return None
    try:
        return float(row["ClsPric"])
    except (KeyError, ValueError):
        return None


# ══════════════════════════════════════════════════════════════════════════════
# Signal classification
# ══════════════════════════════════════════════════════════════════════════════

def classify_symbol(symbol: str, stockname: str, equity_map: dict, etf_set: set):
    reasons = []
    if symbol not in equity_map:
        reasons.append("Not in NSE official equity list")
    if symbol in etf_set:
        reasons.append("In NSE official ETF list")
    hay = f"{symbol} {stockname}".upper()
    hit_kws = [kw for kw in EXCLUDE_KEYWORDS if kw in hay]
    if hit_kws:
        reasons.append(f"Name/symbol contains: {', '.join(hit_kws)}")
    company = equity_map.get(symbol, stockname)
    return (len(reasons) > 0), reasons, company


def read_raw_log() -> list[dict]:
    if not os.path.exists(RAW_LOG_PATH):
        return []
    with open(RAW_LOG_PATH, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def build_signals_and_excluded(raw_rows: list[dict], equity_map: dict, etf_set: set):
    """Returns (signals_log_rows, excluded_rows) — both full, deduped, rederived."""
    signals = []  # (date, rating, symbol, company)
    excluded_seen = {}  # symbol -> (date, rating, stockname, reasons)

    for row in raw_rows:
        date = row["Date"]
        rating = row["Rating"]
        symbol = row["Symbol"].strip()
        stockname = row["StockName"].strip()
        if not symbol:
            continue
        is_excluded, reasons, company = classify_symbol(symbol, stockname, equity_map, etf_set)
        if is_excluded:
            if symbol not in excluded_seen:
                excluded_seen[symbol] = (date, rating, stockname, reasons)
        else:
            signals.append((date, rating, symbol, company))

    excluded_rows = [
        {"symbol": s, "name": v[2], "rating": STAR_LABEL.get(v[1], v[1]),
         "date": v[0], "reason": "; ".join(v[3])}
        for s, v in excluded_seen.items()
    ]
    return signals, excluded_rows


def aggregate_portfolio(signals: list[tuple]) -> list[dict]:
    by_symbol: dict = {}
    for date_str, rating, symbol, company in signals:
        d = datetime.date.fromisoformat(date_str)
        by_symbol.setdefault(symbol, {"company": company, "appearances": []})
        by_symbol[symbol]["appearances"].append((d, rating))

    rows = []
    for symbol, info in by_symbol.items():
        appearances = sorted(info["appearances"], key=lambda x: x[0])
        first_date = appearances[0][0]
        times = len(appearances)

        rating_first_seen: dict = {}
        for d, r in appearances:
            rating_first_seen.setdefault(r, d)
        ordered = sorted(rating_first_seen.items(), key=lambda kv: kv[1])

        rating_field = "/".join(STAR_LABEL.get(r, r) for r, _ in ordered)
        sigdates_field = "/".join(d.strftime("%d%m%y") for _, d in ordered)

        rows.append({
            "symbol": symbol,
            "company": info["company"],
            "rating": rating_field,
            "signal_dates": sigdates_field,
            "first_signal_date": first_date,
            "times_appeared": times,
        })
    return sorted(rows, key=lambda r: (r["first_signal_date"], r["symbol"]))


# ══════════════════════════════════════════════════════════════════════════════
# Price fill-in
# ══════════════════════════════════════════════════════════════════════════════

def load_existing_portfolio() -> dict:
    """symbol -> {buy_date, buy_price} cached from the previous portfolio.csv, if any."""
    cache = {}
    if not os.path.exists(PORTFOLIO_PATH):
        return cache
    with open(PORTFOLIO_PATH, newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            sym = row.get("Symbol", "")
            bp = row.get("Buy Price", "")
            bd = row.get("Buy Date", "")
            if sym and bp:
                cache[sym] = {"buy_date": bd, "buy_price": float(bp)}
    return cache


def fill_prices(rows: list[dict], today: datetime.date, holidays: set, flags: list) -> datetime.date:
    cache = load_existing_portfolio()

    for r in rows:
        symbol = r["symbol"]
        cached = cache.get(symbol)
        if cached:
            r["buy_date"] = datetime.date.fromisoformat(cached["buy_date"])
            r["buy_price"] = cached["buy_price"]
            continue

        buy_day = next_trading_day(r["first_signal_date"], holidays)
        if buy_day > today:
            r["buy_date"] = None
            r["buy_price"] = None
            continue

        bhav = load_bhavcopy(buy_day)
        if bhav is None:
            if buy_day < today:
                flags.append(f"{symbol}: expected bhavcopy missing for {buy_day.isoformat()} (buy date) — left pending, flagged")
            r["buy_date"] = None
            r["buy_price"] = None
            continue

        price = symbol_avg_price(bhav, symbol)
        if price is None:
            flags.append(f"{symbol}: no trade data in bhavcopy {buy_day.isoformat()} (illiquid?) — left pending, flagged")
            r["buy_date"] = None
            r["buy_price"] = None
        else:
            r["buy_date"] = buy_day
            r["buy_price"] = round(price, 6)

    as_of = latest_available_trading_date(today, holidays)
    if as_of is None:
        raise RuntimeError("Could not find any recent published NSE bhavcopy — aborting")
    latest_bhav = load_bhavcopy(as_of)

    for r in rows:
        price = symbol_close_price(latest_bhav, r["symbol"])
        if price is None:
            flags.append(f"{r['symbol']}: no close price in latest bhavcopy {as_of.isoformat()} — left blank, flagged")
        r["latest_price"] = price

    return as_of


# ══════════════════════════════════════════════════════════════════════════════
# File I/O — csv outputs
# ══════════════════════════════════════════════════════════════════════════════

def write_portfolio_csv(rows: list[dict]):
    with open(PORTFOLIO_PATH, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Sr No", "Company", "Symbol", "Rating", "Signal Date(s)", "First Signal Date",
                    "Buy Date", "Buy Price", "Latest Price", "Times Appeared"])
        for i, r in enumerate(rows, start=1):
            w.writerow([
                i, r["company"], r["symbol"], r["rating"], r["signal_dates"],
                r["first_signal_date"].strftime("%Y-%m-%d"),
                r["buy_date"].strftime("%Y-%m-%d") if r["buy_date"] else "",
                f"{r['buy_price']:.6f}" if r["buy_price"] is not None else "",
                f"{r['latest_price']:.4f}" if r["latest_price"] is not None else "",
                r["times_appeared"],
            ])


def write_signals_log_csv(signals: list[tuple]):
    with open(SIGNALS_LOG_PATH, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Date", "Rating", "Symbol", "Company"])
        for date_str, rating, symbol, company in sorted(signals, key=lambda s: (s[0], s[1], s[2])):
            w.writerow([date_str, STAR_LABEL.get(rating, rating), symbol, company])


def write_excluded_csv(excluded_rows: list[dict]):
    with open(EXCLUDED_PATH, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Name/Symbol As Seen", "Full Name", "Source Screener", "Email Date", "Reason Excluded"])
        for e in sorted(excluded_rows, key=lambda x: x["date"]):
            w.writerow([e["symbol"], e["name"], e["rating"], e["date"], e["reason"]])


# ══════════════════════════════════════════════════════════════════════════════
# Excel generation
# ══════════════════════════════════════════════════════════════════════════════

HEADER_FILL = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
GREEN_FILL  = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
RED_FILL    = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
BOLD   = Font(bold=True)
THIN   = Side(style="thin", color="B7B7B7")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
DATE_FMT = "dd-mm-yy"


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


def build_workbook(rows: list[dict], signals: list[tuple], excluded_rows: list[dict], as_of: datetime.date):
    wb = Workbook()
    ws = wb.active
    ws.title = "Portfolio"

    n = len(rows)
    first_data_row = 8
    last_data_row = first_data_row + n - 1

    ws["A1"] = "Capital per stock (Rs)"
    ws["B1"] = CAPITAL_PER_STOCK
    ws["A1"].font = BOLD
    ws["A3"] = "As-of date"
    ws["B3"] = as_of
    ws["B3"].number_format = DATE_FMT
    ws["A3"].font = BOLD

    summary = [
        ("D1", "Total Invested",     "E1", f"=SUM(J{first_data_row}:J{last_data_row})"),
        ("F1", "Current Value",      "G1", f"=SUM(L{first_data_row}:L{last_data_row})"),
        ("D2", "Total P&L (Rs)",     "E2", f"=SUM(M{first_data_row}:M{last_data_row})"),
        ("F2", "Total P&L %",        "G2", "=IF(E1=0,\"\",G1/E1-1)"),
        ("D3", "Winners",            "E3", f"=COUNTIF(M{first_data_row}:M{last_data_row},\">0\")"),
        ("F3", "Losers",             "G3", f"=COUNTIF(M{first_data_row}:M{last_data_row},\"<0\")"),
        ("D4", "Win Rate",           "E4", "=IF((E3+G3)=0,\"\",E3/(E3+G3))"),
        ("F4", "Avg Return / Stock", "G4", f"=AVERAGE(N{first_data_row}:N{last_data_row})"),
        ("D5", "Best Stock",         "E5", f"=INDEX(B{first_data_row}:B{last_data_row},MATCH(MAX(N{first_data_row}:N{last_data_row}),N{first_data_row}:N{last_data_row},0))"),
        ("F5", "Worst Stock",        "G5", f"=INDEX(B{first_data_row}:B{last_data_row},MATCH(MIN(N{first_data_row}:N{last_data_row}),N{first_data_row}:N{last_data_row},0))"),
        ("D6", "Avg Return 4★*", "E6", f"=AVERAGEIF(D{first_data_row}:D{last_data_row},\"4★*\",N{first_data_row}:N{last_data_row})"),
        ("F6", "Avg Return 5★*", "G6", f"=AVERAGEIF(D{first_data_row}:D{last_data_row},\"5★*\",N{first_data_row}:N{last_data_row})"),
    ]
    for lbl_cell, lbl_text, val_cell, formula in summary:
        ws[lbl_cell] = lbl_text
        ws[lbl_cell].font = BOLD
        ws[val_cell] = formula
    for cell in ["G2", "E4", "G4", "E6", "G6"]:
        ws[cell].number_format = "0.00%"
    for cell in ["E1", "G1", "E2"]:
        ws[cell].number_format = "#,##0"

    headers = ["Sr No", "Company", "Symbol", "Rating", "Signal Date(s)", "First Signal Date",
               "Buy Date", "Buy Price", "Qty", "Invested", "Latest Price", "Current Value",
               "P&L Rs", "P&L %", "Days Held", "Status", "Times Appeared"]
    for idx, h in enumerate(headers, start=1):
        ws.cell(row=7, column=idx, value=h)
    style_header_row(ws, 7, len(headers))
    ws.row_dimensions[7].height = 32

    for i, r in enumerate(rows, start=0):
        row = first_data_row + i
        ws.cell(row=row, column=1, value=i + 1)
        ws.cell(row=row, column=2, value=r["company"])
        ws.cell(row=row, column=3, value=r["symbol"])
        ws.cell(row=row, column=4, value=r["rating"])
        ws.cell(row=row, column=5, value=r["signal_dates"])
        ws.cell(row=row, column=6, value=r["first_signal_date"])
        ws.cell(row=row, column=6).number_format = DATE_FMT
        if r["buy_date"]:
            ws.cell(row=row, column=7, value=r["buy_date"])
            ws.cell(row=row, column=7).number_format = DATE_FMT
        else:
            ws.cell(row=row, column=7, value="")
        ws.cell(row=row, column=8, value=round(r["buy_price"], 2) if r["buy_price"] is not None else "")
        ws.cell(row=row, column=8).number_format = "#,##0.00"

        ws.cell(row=row, column=9, value=f"=IF(H{row}=\"\",\"\",ROUNDDOWN($B$1/H{row},0))")
        ws.cell(row=row, column=10, value=f"=IF(H{row}=\"\",\"\",H{row}*I{row})")
        ws.cell(row=row, column=10).number_format = "#,##0.00"
        ws.cell(row=row, column=11, value=round(r["latest_price"], 2) if r["latest_price"] is not None else "")
        ws.cell(row=row, column=11).number_format = "#,##0.00"
        ws.cell(row=row, column=12, value=f"=IF(OR(I{row}=\"\",K{row}=\"\"),\"\",K{row}*I{row})")
        ws.cell(row=row, column=12).number_format = "#,##0.00"
        ws.cell(row=row, column=13, value=f"=IF(L{row}=\"\",\"\",L{row}-J{row})")
        ws.cell(row=row, column=13).number_format = "#,##0.00"
        ws.cell(row=row, column=14, value=f"=IF(OR(H{row}=\"\",K{row}=\"\"),\"\",K{row}/H{row}-1)")
        ws.cell(row=row, column=14).number_format = "0.00%"
        ws.cell(row=row, column=15, value=f"=IF(G{row}=\"\",\"\",$B$3-G{row})")
        ws.cell(row=row, column=16, value=f"=IF(H{row}=\"\",\"Pending buy\",IF(N{row}>=0,\"In profit\",\"In loss\"))")
        ws.cell(row=row, column=17, value=r["times_appeared"])

        for c in range(1, 18):
            ws.cell(row=row, column=c).border = BORDER
            if c != 2:
                ws.cell(row=row, column=c).alignment = Alignment(horizontal="center")

    cf_range = f"A{first_data_row}:Q{last_data_row}"
    ws.conditional_formatting.add(cf_range, FormulaRule(formula=[f"$N{first_data_row}>0"], fill=GREEN_FILL))
    ws.conditional_formatting.add(cf_range, FormulaRule(formula=[f"$N{first_data_row}<0"], fill=RED_FILL))

    ws.freeze_panes = "A8"
    autofit(ws, {1: 6, 2: 30, 3: 12, 4: 8, 5: 14, 6: 14, 7: 12, 8: 12, 9: 8, 10: 13,
                 11: 12, 12: 13, 13: 12, 14: 10, 15: 10, 16: 13, 17: 10})

    ws2 = wb.create_sheet("Signals Log")
    for idx, h in enumerate(["Date", "Rating", "Symbol", "Company"], start=1):
        ws2.cell(row=1, column=idx, value=h)
    style_header_row(ws2, 1, 4)
    for i, (date_str, rating, symbol, company) in enumerate(
            sorted(signals, key=lambda s: (s[0], s[1], s[2])), start=2):
        ws2.cell(row=i, column=1, value=datetime.date.fromisoformat(date_str))
        ws2.cell(row=i, column=1).number_format = DATE_FMT
        ws2.cell(row=i, column=2, value=STAR_LABEL.get(rating, rating))
        ws2.cell(row=i, column=3, value=symbol)
        ws2.cell(row=i, column=4, value=company)
        for c in range(1, 5):
            ws2.cell(row=i, column=c).border = BORDER
            ws2.cell(row=i, column=c).alignment = Alignment(horizontal="center" if c != 4 else "left")
    ws2.freeze_panes = "A2"
    autofit(ws2, {1: 12, 2: 8, 3: 14, 4: 34})

    ws3 = wb.create_sheet("Excluded")
    for idx, h in enumerate(["Name/Symbol As Seen", "Full Name", "Source Screener", "Email Date", "Reason Excluded"], start=1):
        ws3.cell(row=1, column=idx, value=h)
    style_header_row(ws3, 1, 5)
    for i, e in enumerate(sorted(excluded_rows, key=lambda x: x["date"]), start=2):
        ws3.cell(row=i, column=1, value=e["symbol"])
        ws3.cell(row=i, column=2, value=e["name"])
        ws3.cell(row=i, column=3, value=e["rating"])
        ws3.cell(row=i, column=4, value=datetime.date.fromisoformat(e["date"]))
        ws3.cell(row=i, column=4).number_format = DATE_FMT
        ws3.cell(row=i, column=5, value=e["reason"])
        for c in range(1, 6):
            ws3.cell(row=i, column=c).border = BORDER
            ws3.cell(row=i, column=c).alignment = Alignment(horizontal="center" if c in (1, 3, 4) else "left", wrap_text=(c == 5))
    ws3.freeze_panes = "A2"
    autofit(ws3, {1: 20, 2: 30, 3: 18, 4: 14, 5: 55})

    inject_cached_values(wb, ws, rows, first_data_row, last_data_row, as_of)


def inject_cached_values(wb, ws, rows, first_data_row, last_data_row, as_of):
    """openpyxl never evaluates formulas (no cached <v>) — some previewers show blank
    cells without one. Compute values in Python and inject <v> via raw zip/XML patch so
    the live formula AND a correct preview value both exist. See
    feedback_large_attachment_relay.md / project_momentum_portfolio.md for why this
    matters and the self-closing-<v/> gotcha below."""
    wb.save(XLSX_PATH)  # ensure a baseline file exists on disk to patch

    cells = {}
    for i, r in enumerate(rows):
        row = first_data_row + i
        h, k = r["buy_price"], r["latest_price"]
        qty = int(CAPITAL_PER_STOCK // h) if h else None
        inv = h * qty if (h is not None and qty is not None) else None
        cv = k * qty if (qty is not None and k is not None) else None
        pnl = (cv - inv) if (cv is not None and inv is not None) else None
        pct = (k / h - 1) if (h is not None and k is not None) else None
        days = (as_of - r["buy_date"]).days if r["buy_date"] else None
        status = "Pending buy" if h is None else ("In profit" if (pct or 0) >= 0 else "In loss")

        cells[f"I{row}"] = ("n", qty) if qty is not None else ("b", "")
        cells[f"J{row}"] = ("n", round(inv, 2)) if inv is not None else ("b", "")
        cells[f"L{row}"] = ("n", round(cv, 2)) if cv is not None else ("b", "")
        cells[f"M{row}"] = ("n", round(pnl, 2)) if pnl is not None else ("b", "")
        cells[f"N{row}"] = ("n", pct) if pct is not None else ("b", "")
        cells[f"O{row}"] = ("n", days) if days is not None else ("b", "")
        cells[f"P{row}"] = ("s", status)

    invs = [v for k_, (t, v) in cells.items() if k_[0] == "J" and t == "n"]
    cvs  = [v for k_, (t, v) in cells.items() if k_[0] == "L" and t == "n"]
    pnls = [v for k_, (t, v) in cells.items() if k_[0] == "M" and t == "n"]
    pcts = [v for k_, (t, v) in cells.items() if k_[0] == "N" and t == "n"]
    winners = sum(1 for p in pnls if p > 0)
    losers  = sum(1 for p in pnls if p < 0)
    total_inv, total_cv = sum(invs), sum(cvs)

    n_rows = [k_ for k_, (t, v) in cells.items() if k_[0] == "N" and t == "n"]
    if n_rows:
        best_row = max(n_rows, key=lambda k_: cells[k_][1])
        worst_row = min(n_rows, key=lambda k_: cells[k_][1])
        best_name = rows[int(re.sub(r"\D", "", best_row)) - first_data_row]["company"]
        worst_name = rows[int(re.sub(r"\D", "", worst_row)) - first_data_row]["company"]
    else:
        best_name = worst_name = ""

    r4 = [cells[f"N{first_data_row+i}"][1] for i, r in enumerate(rows)
          if r["rating"].startswith("4") and cells[f"N{first_data_row+i}"][0] == "n"]
    r5 = [cells[f"N{first_data_row+i}"][1] for i, r in enumerate(rows)
          if r["rating"].startswith("5") and cells[f"N{first_data_row+i}"][0] == "n"]

    cells["E1"] = ("n", round(total_inv, 2))
    cells["G1"] = ("n", round(total_cv, 2))
    cells["E2"] = ("n", round(total_cv - total_inv, 2))
    cells["G2"] = ("n", (total_cv / total_inv - 1)) if total_inv else ("b", "")
    cells["E3"] = ("n", winners)
    cells["G3"] = ("n", losers)
    cells["E4"] = ("n", winners / (winners + losers)) if (winners + losers) else ("b", "")
    cells["G4"] = ("n", sum(pcts) / len(pcts)) if pcts else ("b", "")
    cells["E5"] = ("s", best_name)
    cells["G5"] = ("s", worst_name)
    cells["E6"] = ("n", sum(r4) / len(r4)) if r4 else ("b", "")
    cells["G6"] = ("n", sum(r5) / len(r5)) if r5 else ("b", "")

    def patch_cell(xml, ref, kind, value):
        pattern = re.compile(r'(<c r="' + re.escape(ref) + r'"[^>]*>)(.*?)(</c>)', re.DOTALL)
        m = pattern.search(xml)
        if not m:
            return xml
        open_tag, inner, close_tag = m.groups()
        inner_no_v = re.sub(r"<v\s*/>|<v>.*?</v>", "", inner, flags=re.DOTALL)
        if kind == "n":
            v_xml = f"<v>{value}</v>"
            new_open = re.sub(r' t="[^"]*"', "", open_tag)
        else:
            esc = str(value).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            v_xml = f"<v>{esc}</v>"
            new_open = re.sub(r' t="[^"]*"', ' t="str"', open_tag) if ' t="' in open_tag else open_tag[:-1] + ' t="str">'
        new_inner = inner_no_v + v_xml
        return xml[:m.start()] + new_open + new_inner + close_tag + xml[m.end():]

    with zipfile.ZipFile(XLSX_PATH, "r") as z:
        names = z.namelist()
        data = {n: z.read(n) for n in names}

    sheet1 = data["xl/worksheets/sheet1.xml"].decode("utf-8")
    for ref, (kind, value) in cells.items():
        sheet1 = patch_cell(sheet1, ref, kind, value)
    data["xl/worksheets/sheet1.xml"] = sheet1.encode("utf-8")

    tmp = XLSX_PATH + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for n in names:
            z.writestr(n, data[n])
    os.replace(tmp, XLSX_PATH)


# ══════════════════════════════════════════════════════════════════════════════
# Email
# ══════════════════════════════════════════════════════════════════════════════

def recipients() -> list[str]:
    return DRY_RUN_RECIPIENTS if DRY_RUN else PROD_RECIPIENTS


def send_friday_email(rows: list[dict], as_of: datetime.date, new_names: list[str],
                       data_stale: bool) -> None:
    if not GMAIL_USER or not GMAIL_PASSWORD:
        log.warning("Gmail credentials not set — skipping email")
        return

    settled = [r for r in rows if r["latest_price"] is not None and r["buy_price"] is not None]
    pnls, invs, pcts = [], [], []
    for r in settled:
        qty = int(CAPITAL_PER_STOCK // r["buy_price"])
        inv = r["buy_price"] * qty
        cv = r["latest_price"] * qty
        pnls.append(cv - inv)
        invs.append(inv)
        pcts.append(r["latest_price"] / r["buy_price"] - 1)

    # Dollar-weighted, matching the xlsx Summary block's G2 = SUM(M)/SUM(J) exactly
    total_pnl_pct = (sum(pnls) / sum(invs) * 100) if invs and sum(invs) else 0.0
    winners = sum(1 for p in pnls if p > 0)
    losers = sum(1 for p in pnls if p < 0)
    if pcts:
        best_i = max(range(len(settled)), key=lambda i: pcts[i])
        worst_i = min(range(len(settled)), key=lambda i: pcts[i])
        best = f"{settled[best_i]['symbol']} ({pcts[best_i]*100:+.2f}%)"
        worst = f"{settled[worst_i]['symbol']} ({pcts[worst_i]*100:+.2f}%)"
    else:
        best = worst = "n/a"

    lines = [
        f"Total P&L: {total_pnl_pct:+.2f}%",
        f"Winners: {winners} | Losers: {losers}",
        f"Best: {best} | Worst: {worst}",
        f"New names this week: {', '.join(new_names) if new_names else 'none'}",
    ]
    if data_stale:
        lines.append(f"(Note: today's NSE data was not published by the 8:00 PM IST cutoff — "
                      f"figures use the last available close, {as_of.strftime('%d %b %Y')}.)")
    if DRY_RUN:
        lines.append("(DRY RUN — sent only to Kinjal for review, not the full recipient list.)")

    body = "\n".join(lines)

    msg = EmailMessage()
    msg["Subject"] = "Momentum Portfolio"
    msg["From"] = GMAIL_USER
    msg["To"] = ", ".join(recipients())
    msg.set_content(body)

    with open(XLSX_PATH, "rb") as f:
        msg.add_attachment(
            f.read(), maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename="portfolio.xlsx",
        )

    ctx = ssl.create_default_context()
    with smtplib.SMTP("smtp.gmail.com", 587) as smtp:
        smtp.ehlo()
        smtp.starttls(context=ctx)
        smtp.login(GMAIL_USER, GMAIL_PASSWORD)
        smtp.send_message(msg)
    log.info("Momentum Portfolio email sent -> %s", recipients())


# ══════════════════════════════════════════════════════════════════════════════
# State
# ══════════════════════════════════════════════════════════════════════════════

def load_state() -> dict:
    if os.path.exists(STATE_PATH):
        with open(STATE_PATH, encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_state(state: dict):
    with open(STATE_PATH, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2)


def log_friday_publish(effective_date: datetime.date, available: bool, checked_at_ist: str):
    file_exists = os.path.exists(FRIDAY_LOG_PATH)
    with open(FRIDAY_LOG_PATH, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if not file_exists:
            w.writerow(["Date", "DataAvailable", "SentAtIST"])
        w.writerow([effective_date.isoformat(), "yes" if available else "no (used stale data)", checked_at_ist])


# ══════════════════════════════════════════════════════════════════════════════
# Core pipeline
# ══════════════════════════════════════════════════════════════════════════════

def run_pipeline():
    today = datetime.datetime.now(IST).date()
    equity_map = fetch_equity_list()
    etf_set = fetch_etf_set()
    holidays = fetch_nse_holidays()

    raw_rows = read_raw_log()
    signals, excluded_rows = build_signals_and_excluded(raw_rows, equity_map, etf_set)
    rows = aggregate_portfolio(signals)

    flags: list = []
    as_of = fill_prices(rows, today, holidays, flags)

    write_signals_log_csv(signals)
    write_excluded_csv(excluded_rows)
    write_portfolio_csv(rows)
    build_workbook(rows, signals, excluded_rows, as_of)

    if flags:
        for f in flags:
            log.warning("FLAG: %s", f)

    return rows, signals, excluded_rows, as_of, holidays


def mode_silent():
    today = datetime.datetime.now(IST).date()
    if today.weekday() not in (0, 1, 2, 3):  # Mon-Thu only
        log.info("Not Mon-Thu (weekday=%d) — silent update skipped today", today.weekday())
        return
    rows, signals, excluded_rows, as_of, _ = run_pipeline()
    log.info("Silent update done. %d stocks tracked, as-of %s", len(rows), as_of.isoformat())


def mode_friday():
    today = datetime.datetime.now(IST).date()
    now = datetime.datetime.now(IST)
    wd = today.weekday()  # Mon=0 .. Fri=4

    holidays = fetch_nse_holidays()

    if wd == 4:  # Friday
        if today in holidays:
            log.info("Today (Friday) is an NSE holiday — ceremony shifts to Thursday, nothing to do")
            return
    elif wd == 3:  # Thursday
        tomorrow = today + datetime.timedelta(days=1)
        if tomorrow not in holidays:
            log.info("Thursday, but tomorrow (Friday) is not a holiday — not this week's ceremony day")
            return
        log.info("Friday is an NSE holiday — running today's (Thursday) ceremony instead")
    else:
        log.info("Not Thursday/Friday (weekday=%d) — nothing to do", wd)
        return

    state = load_state()
    if state.get("last_friday_email_date") == today.isoformat():
        log.info("Already sent today's Momentum Portfolio email (%s) — exiting", today.isoformat())
        return

    CUTOFF = now.replace(hour=20, minute=0, second=0, microsecond=0)
    data_available = bhavcopy_exists(today)
    checked_at = now.strftime("%H:%M IST")

    if not data_available and now < CUTOFF:
        log.info("Today's bhavcopy not published yet (checked %s) — will retry next poll", checked_at)
        return

    if not data_available:
        log.warning("Cutoff (8:00 PM IST) reached and today's data still not published — "
                    "sending with last-available data")

    rows, signals, excluded_rows, as_of, _ = run_pipeline()

    previous_cutoff = state.get("last_friday_email_date")
    new_names = sorted({r["symbol"] for r in rows
                         if previous_cutoff and r["first_signal_date"].isoformat() > previous_cutoff}) \
        if previous_cutoff else sorted(r["symbol"] for r in rows
                                        if (today - r["first_signal_date"]).days <= 7)

    send_friday_email(rows, as_of, new_names, data_stale=not data_available)

    log_friday_publish(today, data_available, checked_at)
    state["last_friday_email_date"] = today.isoformat()
    save_state(state)
    log.info("Friday ceremony complete for %s", today.isoformat())


def main():
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--mode", choices=["silent", "friday"], default="silent")
    args = parser.parse_args()

    log.info("Momentum Portfolio build — mode=%s DRY_RUN=%s", args.mode, DRY_RUN)
    if args.mode == "friday":
        mode_friday()
    else:
        mode_silent()


if __name__ == "__main__":
    main()
