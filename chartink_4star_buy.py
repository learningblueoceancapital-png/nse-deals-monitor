#!/usr/bin/env python3
"""
Chartink Screener Email — 4 Star Buy
──────────────────────────────────────
Scrapes https://chartink.com/screener/4-star-buy using Playwright,
filters to cash segment only (excludes ETFs and Indices),
and emails results as a formatted HTML table.

Env vars (or GitHub Secrets):
  GMAIL_USER          your Gmail address
  GMAIL_APP_PASSWORD  16-char Gmail App Password
"""

import os
import re
import sys
import smtplib
import ssl
import logging
import datetime
from email.message import EmailMessage
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# ── Logging ────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger("chartink_4star")

# ── Config ─────────────────────────────────────────────────────────────────────
GMAIL_USER     = os.getenv("GMAIL_USER", "")
GMAIL_PASSWORD = os.getenv("GMAIL_APP_PASSWORD", "")

EMAIL_TO = [
    "lavneesh@blueoceancapital.co.in",
    "operations@blueoceancapital.co.in",
    "research@blueoceancapital.co.in",
]

SCREENER_URL  = "https://chartink.com/screener/4-star-buy"
SCREENER_NAME = "4 Star Buy"
IST           = datetime.timezone(datetime.timedelta(hours=5, minutes=30))

BROWSER_ARGS = [
    "--no-sandbox",
    "--disable-blink-features=AutomationControlled",
    "--disable-dev-shm-usage",
    "--window-size=1920,1080",
]

# Patterns to identify indices and ETFs (cash segment exclusion)
_INDEX_SYMBOL_RE = re.compile(
    r"^(NIFTY|SENSEX|BSE|DJIA|NASDAQ|S&P|INDIA\s*VIX|MIDCPNIFTY|BANKEX)",
    re.IGNORECASE,
)
_ETF_NAME_RE = re.compile(r"\betf\b|\bfund\b|\bindex\b|\bbees\b", re.IGNORECASE)


# ══════════════════════════════════════════════════════════════════════════════
# Scraping
# ══════════════════════════════════════════════════════════════════════════════

def fetch_screener() -> list[dict]:
    """Load Chartink screener via Playwright and return cash-segment rows only."""
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True, args=BROWSER_ARGS)
        ctx = browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            viewport={"width": 1920, "height": 1080},
            locale="en-IN",
            timezone_id="Asia/Kolkata",
        )
        ctx.add_init_script(
            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined});"
        )
        page = ctx.new_page()

        try:
            log.info("Loading screener page …")
            page.goto(SCREENER_URL, wait_until="domcontentloaded", timeout=60_000)
            page.wait_for_selector(
                "table.scan-results-table tbody tr",
                timeout=30_000,
            )
            page.wait_for_timeout(3_000)

            all_rows = _extract_table(page)
            log.info("Raw rows from screener: %d", len(all_rows))

            cash_rows = _filter_cash_segment(all_rows)
            excluded  = len(all_rows) - len(cash_rows)
            if excluded:
                log.info("Excluded %d ETF/Index row(s) — cash segment rows: %d", excluded, len(cash_rows))

            # Re-number Sr. after filtering
            for i, row in enumerate(cash_rows, start=1):
                if "Sr." in row:
                    row["Sr."] = str(i)

            return cash_rows

        except PWTimeout:
            log.warning("Timeout waiting for screener results — page may be empty or slow")
            return []
        finally:
            browser.close()


def _clean_header(text: str) -> str:
    return text.split("\n")[0].strip()


def _clean_cell(text: str) -> str:
    return text.split("\n")[0].strip()


def _extract_table(page) -> list[dict]:
    raw_headers = page.eval_on_selector_all(
        "table.scan-results-table thead th",
        "els => els.map(el => el.innerText.trim())",
    )
    headers = [_clean_header(h) for h in raw_headers if _clean_header(h)]

    raw_rows = page.query_selector_all("table.scan-results-table tbody tr")

    results = []
    for row in raw_rows:
        cells  = row.query_selector_all("td")
        values = [_clean_cell(c.inner_text()) for c in cells]
        if not any(values):
            continue
        values  = values[: len(headers)]
        padded  = (values + [""] * len(headers))[: len(headers)]
        results.append(dict(zip(headers, padded)))

    return results


def _filter_cash_segment(rows: list[dict]) -> list[dict]:
    """
    Remove ETFs and Indices. Cash segment equity stocks have plain NSE symbols
    (letters only, ≤ ~15 chars). Indices have NIFTY*/SENSEX* symbols; ETFs
    typically have 'ETF' / 'Fund' / 'BEES' in their name.
    """
    clean = []
    for row in rows:
        symbol = row.get("Symbol", "")
        name   = row.get("Stock Name", "")

        if _INDEX_SYMBOL_RE.match(symbol):
            log.debug("Excluded index: %s (%s)", name, symbol)
            continue
        if _ETF_NAME_RE.search(name):
            log.debug("Excluded ETF: %s (%s)", name, symbol)
            continue

        clean.append(row)
    return clean


# ══════════════════════════════════════════════════════════════════════════════
# Email
# ══════════════════════════════════════════════════════════════════════════════

def _build_html(rows: list[dict], run_date: str) -> str:
    if not rows:
        return f"""
        <html><body style="font-family:Calibri,Arial,sans-serif;color:#222;">
        <h2 style="color:#1F4E79;">{SCREENER_NAME} — {run_date}</h2>
        <p>No present stocks under this chartink screener.</p>
        <p>Regards,<br>Aumkar</p>
        </body></html>
        """

    headers = list(rows[0].keys())

    header_html = "".join(
        f'<th style="background:#1F4E79;color:#fff;padding:8px 12px;'
        f'text-align:left;white-space:nowrap;">{h}</th>'
        for h in headers
    )

    rows_html = ""
    for i, row in enumerate(rows):
        bg    = "#D6E4F0" if i % 2 == 0 else "#ffffff"
        cells = "".join(
            f'<td style="padding:7px 12px;border-bottom:1px solid #cce0f0;">'
            f'{row.get(h, "")}</td>'
            for h in headers
        )
        rows_html += f'<tr style="background:{bg};">{cells}</tr>\n'

    return f"""
    <html>
    <body style="font-family:Calibri,Arial,sans-serif;color:#222;margin:0;padding:20px;">
      <h2 style="color:#1F4E79;margin-bottom:4px;">{SCREENER_NAME}</h2>
      <p style="color:#555;margin-top:0;font-size:13px;">
        Run date: <strong>{run_date}</strong> &nbsp;|&nbsp;
        {len(rows)} stock{"s" if len(rows) != 1 else ""} matched
        (cash segment only)
      </p>
      <table cellspacing="0" cellpadding="0" border="0"
             style="border-collapse:collapse;border:1px solid #1F4E79;min-width:500px;">
        <thead><tr>{header_html}</tr></thead>
        <tbody>{rows_html}</tbody>
      </table>
      <p style="color:#888;font-size:11px;margin-top:16px;">
        Source: <a href="{SCREENER_URL}" style="color:#1F4E79;">{SCREENER_URL}</a>
      </p>
      <p style="margin-top:24px;">Regards,<br>Aumkar</p>
    </body>
    </html>
    """


def send_email(rows: list[dict], run_date: str) -> None:
    if not GMAIL_USER or not GMAIL_PASSWORD:
        log.warning("Gmail credentials not set — email skipped.")
        return

    count   = len(rows)
    subject = f"{SCREENER_NAME} | {run_date} | {count} stock{'s' if count != 1 else ''}"

    plain = f"{SCREENER_NAME} — {run_date}\n{count} stocks matched (cash segment only).\n\n"
    if rows:
        headers = list(rows[0].keys())
        plain  += "\t".join(headers) + "\n"
        for row in rows:
            plain += "\t".join(str(row.get(h, "")) for h in headers) + "\n"
    plain += "\nRegards,\nAumkar"

    html = _build_html(rows, run_date)

    msg            = EmailMessage()
    msg["Subject"] = subject
    msg["From"]    = GMAIL_USER
    msg["To"]      = ", ".join(EMAIL_TO)
    msg.set_content(plain)
    msg.add_alternative(html, subtype="html")

    try:
        ctx = ssl.create_default_context()
        with smtplib.SMTP("smtp.gmail.com", 587) as smtp:
            smtp.ehlo()
            smtp.starttls(context=ctx)
            smtp.login(GMAIL_USER, GMAIL_PASSWORD)
            smtp.send_message(msg)
        log.info("Email sent → %s", EMAIL_TO)
    except Exception as exc:
        log.error("Email failed: %s", exc)
        raise


# ══════════════════════════════════════════════════════════════════════════════
# Entry point
# ══════════════════════════════════════════════════════════════════════════════

def main() -> None:
    now      = datetime.datetime.now(IST)
    run_date = now.strftime("%d %b %Y")
    log.info("4 Star Buy screener run — %s IST", now.strftime("%Y-%m-%d %H:%M"))

    rows = fetch_screener()
    send_email(rows, run_date)
    log.info("Done.")


if __name__ == "__main__":
    main()
