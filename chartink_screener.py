#!/usr/bin/env python3
"""
Chartink Screener Email — 5 Star Buy
─────────────────────────────────────
Scrapes https://chartink.com/screener/5-star-buy-2 using Playwright,
extracts the results table, and emails it as a formatted HTML table.

Env vars (or GitHub Secrets):
  GMAIL_USER          your Gmail address
  GMAIL_APP_PASSWORD  16-char Gmail App Password
  EMAIL_TO            comma-separated recipients (override default list)
"""

import os
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
log = logging.getLogger("chartink")

# ── Config ─────────────────────────────────────────────────────────────────────
GMAIL_USER     = os.getenv("GMAIL_USER", "")
GMAIL_PASSWORD = os.getenv("GMAIL_APP_PASSWORD", "")
_default_to    = (
    "lavneesh@blueoceancapital.co.in,"
    "operations@blueoceancapital.co.in,"
    "research@blueoceancapital.co.in,"
    "aumkar.rasal@gmail.com"
)
EMAIL_TO = [
    e.strip()
    for e in os.getenv("EMAIL_TO", _default_to).split(",")
    if e.strip()
]

SCREENER_URL  = "https://chartink.com/screener/5-star-buy-2"
SCREENER_NAME = "5 Star Buy"
IST           = datetime.timezone(datetime.timedelta(hours=5, minutes=30))

BROWSER_ARGS = [
    "--no-sandbox",
    "--disable-blink-features=AutomationControlled",
    "--disable-dev-shm-usage",
    "--window-size=1920,1080",
]


# ══════════════════════════════════════════════════════════════════════════════
# Scraping
# ══════════════════════════════════════════════════════════════════════════════

def fetch_screener() -> list[dict]:
    """
    Launch headless Chromium, load the Chartink screener page, wait for
    the results table to populate, and return rows as list-of-dicts.
    """
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

            # Chartink uses scan-results-table (Vue.js SPA, renders async)
            page.wait_for_selector(
                "table.scan-results-table tbody tr",
                timeout=30_000,
            )
            # Small extra wait for all rows to render
            page.wait_for_timeout(3_000)

            rows = _extract_table(page)
            log.info("Extracted %d rows from screener", len(rows))
            return rows

        except PWTimeout:
            log.warning("Timeout waiting for screener results — page may be empty or slow")
            return []
        finally:
            browser.close()


def _clean_header(text: str) -> str:
    """Strip the 'Sort table by … in … order' suffix Chartink adds to th text."""
    return text.split("\n")[0].strip()


def _clean_cell(text: str) -> str:
    """Strip embedded buttons like 'Exclude ETFs' from cell text."""
    return text.split("\n")[0].strip()


def _extract_table(page) -> list[dict]:
    """Pull header + data rows from the Chartink scan-results-table."""
    raw_headers = page.eval_on_selector_all(
        "table.scan-results-table thead th",
        "els => els.map(el => el.innerText.trim())",
    )
    headers = [_clean_header(h) for h in raw_headers if _clean_header(h)]

    raw_rows = page.query_selector_all("table.scan-results-table tbody tr")

    results = []
    for row in raw_rows:
        cells = row.query_selector_all("td")
        values = [_clean_cell(c.inner_text()) for c in cells]
        if not any(values):
            continue
        # Drop trailing empty column (the '+' add-column placeholder)
        values = values[: len(headers)]
        padded = (values + [""] * len(headers))[: len(headers)]
        results.append(dict(zip(headers, padded)))

    return results


# ══════════════════════════════════════════════════════════════════════════════
# Email
# ══════════════════════════════════════════════════════════════════════════════

def _build_html(rows: list[dict], run_date: str) -> str:
    if not rows:
        return f"""
        <html><body style="font-family:Calibri,Arial,sans-serif;color:#222;">
        <h2 style="color:#1F4E79;">{SCREENER_NAME} — {run_date}</h2>
        <p>No stocks matched the screener today.</p>
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
        bg = "#D6E4F0" if i % 2 == 0 else "#ffffff"
        cells = "".join(
            f'<td style="padding:7px 12px;border-bottom:1px solid #cce0f0;">{row.get(h, "")}</td>'
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

    count  = len(rows)
    subject = f"{SCREENER_NAME} | {run_date} | {count} stock{'s' if count != 1 else ''}"

    plain = f"{SCREENER_NAME} — {run_date}\n{count} stocks matched.\n\n"
    if rows:
        headers = list(rows[0].keys())
        plain += "\t".join(headers) + "\n"
        for row in rows:
            plain += "\t".join(str(row.get(h, "")) for h in headers) + "\n"
    plain += "\nRegards,\nAumkar"

    html = _build_html(rows, run_date)

    msg             = EmailMessage()
    msg["Subject"]  = subject
    msg["From"]     = GMAIL_USER
    msg["To"]       = ", ".join(EMAIL_TO)
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

    log.info("Chartink screener run — %s IST", now.strftime("%Y-%m-%d %H:%M"))

    rows = fetch_screener()
    send_email(rows, run_date)

    log.info("Done.")


if __name__ == "__main__":
    main()
