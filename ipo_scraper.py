#!/usr/bin/env python3
"""
IPO Anchor Lock-in Tracker
──────────────────────────
Scrapes Chittorgarh (primary) for IPO data, calculates anchor investor
lock-in expiry dates, and exports to docs/data.json for the static website.

Data collected per IPO:
  - Company name, category (Mainboard / SME), fiscal year
  - Open / Close / Allotment / Listing dates
  - Issue size (₹ Cr), issue price
  - Anchor shares allotted, anchor allocation value (₹ Cr)
  - 30-day lock-in expiry (50 % of anchor shares)
  - 90-day lock-in expiry (remaining 50 % of anchor shares)

Usage:
  python ipo_scraper.py             # incremental update (new IPOs only)
  python ipo_scraper.py --backfill  # force re-scrape all IPO detail pages
  python ipo_scraper.py --export    # re-export JSON without scraping
"""

import argparse
import json
import logging
import os
import re
import sqlite3
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup

# ── Logging ────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger("ipo_tracker")

# ── Paths ──────────────────────────────────────────────────────────────────────
ROOT     = Path(__file__).parent
DB_PATH  = ROOT / "ipo_data.db"
JSON_OUT = ROOT / "docs" / "data.json"

# ── Constants ──────────────────────────────────────────────────────────────────
BASE_URL      = "https://www.chittorgarh.com"
REQUEST_DELAY  = 1.5   # seconds between requests
TIMEOUT        = 30

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-IN,en;q=0.9",
    "Referer": "https://www.chittorgarh.com/",
}

# Chittorgarh list-page URLs — Next.js/React app, requires Playwright.
# Canonical URL format: /report/ipo-in-india-list-main-board-sme/82/{type}/?year={YYYY}
# Calendar year 2024 = Apr-Dec 2024 (part of FY2025)
# Calendar year 2025 = Jan-Mar 2025 (rest of FY2025) + Apr-Dec 2025 (start of FY2026)
# Calendar year 2026 = Jan-Mar 2026 (rest of FY2026)
_BASE_LIST = "https://www.chittorgarh.com/report/ipo-in-india-list-main-board-sme/82"
LIST_PAGES = [
    (f"{_BASE_LIST}/mainboard/?year=2024", "Mainboard"),
    (f"{_BASE_LIST}/sme/?year=2024",       "SME"),
    (f"{_BASE_LIST}/mainboard/?year=2025", "Mainboard"),
    (f"{_BASE_LIST}/sme/?year=2025",       "SME"),
    (f"{_BASE_LIST}/mainboard/?year=2026", "Mainboard"),
    (f"{_BASE_LIST}/sme/?year=2026",       "SME"),
]

# IPO ID pattern on Chittorgarh: /ipo/{slug}-ipo/{id}/
IPO_URL_RE = re.compile(r'https://www\.chittorgarh\.com/ipo/[^/]+-ipo/(\d+)/')


# ── Database ───────────────────────────────────────────────────────────────────

def open_db() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("""
        CREATE TABLE IF NOT EXISTS ipos (
            source_url       TEXT PRIMARY KEY,
            company_name     TEXT NOT NULL,
            category         TEXT NOT NULL,
            fy               TEXT,
            open_date        TEXT,
            close_date       TEXT,
            allotment_date   TEXT,
            listing_date     TEXT,
            issue_size_cr    REAL,
            issue_price      TEXT,
            anchor_shares    INTEGER,
            anchor_value_cr  REAL,
            lockin_30d       TEXT,
            lockin_90d       TEXT,
            detail_scraped   INTEGER DEFAULT 0,
            last_updated     TEXT
        )
    """)
    conn.commit()
    return conn


def upsert(conn: sqlite3.Connection, row: dict) -> None:
    row.setdefault("last_updated", utcnow())
    conn.execute("""
        INSERT INTO ipos (
            source_url, company_name, category, fy,
            open_date, close_date, allotment_date, listing_date,
            issue_size_cr, issue_price,
            anchor_shares, anchor_value_cr,
            lockin_30d, lockin_90d,
            detail_scraped, last_updated
        ) VALUES (
            :source_url, :company_name, :category, :fy,
            :open_date, :close_date, :allotment_date, :listing_date,
            :issue_size_cr, :issue_price,
            :anchor_shares, :anchor_value_cr,
            :lockin_30d, :lockin_90d,
            :detail_scraped, :last_updated
        )
        ON CONFLICT(source_url) DO UPDATE SET
            company_name    = excluded.company_name,
            category        = excluded.category,
            fy              = COALESCE(excluded.fy,             fy),
            open_date       = COALESCE(excluded.open_date,      open_date),
            close_date      = COALESCE(excluded.close_date,     close_date),
            allotment_date  = COALESCE(excluded.allotment_date, allotment_date),
            listing_date    = COALESCE(excluded.listing_date,   listing_date),
            issue_size_cr   = COALESCE(excluded.issue_size_cr,  issue_size_cr),
            issue_price     = COALESCE(excluded.issue_price,    issue_price),
            anchor_shares   = COALESCE(excluded.anchor_shares,  anchor_shares),
            anchor_value_cr = COALESCE(excluded.anchor_value_cr,anchor_value_cr),
            lockin_30d      = COALESCE(excluded.lockin_30d,     lockin_30d),
            lockin_90d      = COALESCE(excluded.lockin_90d,     lockin_90d),
            detail_scraped  = MAX(excluded.detail_scraped, detail_scraped),
            last_updated    = excluded.last_updated
    """, row)
    conn.commit()


# ── Helpers ────────────────────────────────────────────────────────────────────

def utcnow() -> str:
    return datetime.now(tz=__import__("datetime").timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def parse_date(raw: str | None) -> str | None:
    if not raw:
        return None
    raw = re.sub(r"(\d+)(st|nd|rd|th)", r"\1", raw.strip())
    raw = raw.strip().rstrip("T")  # remove trailing "T" artifacts from DOM
    if raw in ("", "-", "N/A", "TBA"):
        return None
    for fmt in (
        "%d %b %Y", "%d-%b-%Y", "%b %d, %Y", "%B %d, %Y",
        "%d/%m/%Y", "%Y-%m-%d", "%d %B %Y", "%a, %b %d, %Y",
        "%a, %B %d, %Y", "%b %d %Y", "%d %b, %Y",
    ):
        try:
            return datetime.strptime(raw, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return None


def parse_crore(raw: str | None) -> float | None:
    if not raw:
        return None
    val = re.sub(r"[₹₹Rs\s]", "", raw, flags=re.IGNORECASE)
    val = re.sub(r"(?i)(crore[s]?|cr\.?)", "", val)
    val = re.sub(r"[,]", "", val).strip()
    try:
        return round(float(val), 2)
    except (ValueError, TypeError):
        return None


def parse_number(raw: str | None) -> int | None:
    if not raw:
        return None
    val = re.sub(r"[,\s]", "", raw.strip())
    try:
        return int(float(val))
    except (ValueError, TypeError):
        return None


def determine_fy(date_str: str | None) -> str | None:
    if not date_str:
        return None
    try:
        d = datetime.strptime(date_str, "%Y-%m-%d")
        fy_year = d.year if d.month >= 4 else d.year - 1
        return f"FY{fy_year + 1}"
    except ValueError:
        return None


def calc_lockin(allotment: str | None) -> tuple[str | None, str | None]:
    if not allotment:
        return None, None
    try:
        d = datetime.strptime(allotment, "%Y-%m-%d")
        return (
            (d + timedelta(days=30)).strftime("%Y-%m-%d"),
            (d + timedelta(days=90)).strftime("%Y-%m-%d"),
        )
    except ValueError:
        return None, None


# ── HTTP session ───────────────────────────────────────────────────────────────

_session: requests.Session | None = None


def get_session() -> requests.Session:
    global _session
    if _session is None:
        _session = requests.Session()
        _session.headers.update(HEADERS)
    return _session


def fetch_soup(url: str, retries: int = 3) -> BeautifulSoup | None:
    s = get_session()
    for attempt in range(1, retries + 1):
        try:
            r = s.get(url, timeout=TIMEOUT)
            if r.status_code == 404:
                return None
            r.raise_for_status()
            return BeautifulSoup(r.content, "html.parser")
        except Exception as exc:
            log.warning("Attempt %d/%d — %s — %s", attempt, retries, url, exc)
            if attempt < retries:
                time.sleep(attempt * 2)
    return None


# ── List page: Playwright (JS-rendered) ───────────────────────────────────────

def scrape_list_with_playwright(url: str, category: str) -> list[dict]:
    """
    Use Playwright to render a Chittorgarh IPO list page (Next.js/React app).

    Flow:
      1. Navigate and wait for network idle.
      2. Wait 6.5 s for the ad interstitial timer, then click "Skip Ad".
      3. Wait for the IPO table to hydrate (look for an anchor link matching
         the IPO detail URL pattern).
      4. Extract all unique IPO detail links + company names via JS.
    """
    try:
        from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
    except ImportError:
        log.error("Playwright not installed — run: pip install playwright && playwright install chromium")
        return []

    log.info("Playwright list [%s]: %s", category, url)
    items: dict[str, str] = {}  # url → company_name

    with sync_playwright() as pw:
        browser = pw.chromium.launch(
            headless=True,
            args=["--no-sandbox", "--disable-dev-shm-usage",
                  "--disable-blink-features=AutomationControlled"],
        )
        ctx = browser.new_context(
            user_agent=HEADERS["User-Agent"],
            locale="en-IN",
            timezone_id="Asia/Kolkata",
            extra_http_headers={
                "Accept-Language": "en-IN,en-US;q=0.9,en;q=0.8",
            },
        )
        ctx.add_init_script(
            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined});"
        )
        page = ctx.new_page()
        try:
            page.goto(url, wait_until="networkidle", timeout=60_000)

            # Ad interstitial has a 5-second timer; wait 6.5 s then click Skip
            page.wait_for_timeout(6_500)
            for skip_sel in ["text=Skip Ad", "text=Skip", "[id*=skip]", "[class*=skip-ad]"]:
                try:
                    btn = page.locator(skip_sel).first
                    if btn.is_visible(timeout=500):
                        btn.click()
                        log.debug("Clicked '%s'", skip_sel)
                        break
                except Exception:
                    pass

            # Wait for the React component to render IPO links
            try:
                page.wait_for_selector(
                    "a[href*='-ipo/'][href*='chittorgarh']",
                    timeout=20_000,
                )
            except PWTimeout:
                log.warning("Timeout waiting for IPO links on %s", url)

            page.wait_for_timeout(3_000)

            # Extract via JS — more reliable than BeautifulSoup on rendered HTML
            raw = page.evaluate(r"""
                () => {
                    const links = Array.from(document.querySelectorAll('a[href]'));
                    return links
                        .filter(a => /\/ipo\/[^/]+-ipo\/\d+\//.test(a.href))
                        .map(a => ({href: a.href, text: a.textContent.trim()}));
                }
            """)
            for item in (raw or []):
                href = item.get("href", "")
                text = item.get("text", "")
                if href and IPO_URL_RE.match(href) and href not in items:
                    name = re.sub(r"\s*(?:IPO|Ltd\.?|Limited)\s*$", "", text,
                                  flags=re.IGNORECASE).strip()
                    if name:
                        items[href] = name

        except Exception as exc:
            log.error("Playwright error on %s: %s", url, exc)
        finally:
            browser.close()

    ipos = [
        {"source_url": href, "company_name": name,
         "category": category, "detail_scraped": 0}
        for href, name in items.items()
    ]
    log.info("  Found %d IPOs", len(ipos))
    return ipos


# ── Detail page: requests + BeautifulSoup ─────────────────────────────────────

def _extract_li_dates(soup: BeautifulSoup) -> dict[str, str | None]:
    """
    Extract dates from Chittorgarh's <li class="d-flex ..."> key-value list items.
    Structure: <span>[keyword link]Allotment</span> <span class="text-end">date</span>
    """
    result: dict[str, str | None] = {}
    for li in soup.find_all("li", class_=re.compile(r"d-flex")):
        spans = li.find_all("span")
        if len(spans) < 2:
            continue
        key = spans[0].get_text(" ", strip=True).lower()
        val = spans[-1].get_text(strip=True)
        if not val or val in ("—", "-", ""):
            continue
        if "open" in key and "ipo" in key:
            result.setdefault("open_date", parse_date(val))
        elif "close" in key and "ipo" in key:
            result.setdefault("close_date", parse_date(val))
        elif "allotment" in key and "basis" not in key:
            result.setdefault("allotment_date", parse_date(val))
        elif "listing" in key or "listed on" in key:
            result.setdefault("listing_date", parse_date(val))
    return result


def _extract_anchor_table(soup: BeautifulSoup) -> dict:
    """
    Find and parse the anchor investor table.
    Identified by presence of 'Anchor Portion' or 'Bid Date' row with exactly 2 cols.
    """
    for table in soup.find_all("table"):
        rows = table.find_all("tr")
        kv: dict[str, str] = {}
        for row in rows:
            cells = row.find_all(["td", "th"])
            if len(cells) == 2:
                k = cells[0].get_text(" ", strip=True).lower()
                v = cells[1].get_text(" ", strip=True)
                if k and v:
                    kv[k] = v

        # Confirm this is the anchor table
        keys_str = " ".join(kv.keys())
        if "anchor portion" not in keys_str and "bid date" not in keys_str:
            continue

        anchor_shares   = parse_number(kv.get("shares offered"))
        anchor_value_cr = parse_crore(
            kv.get("anchor portion (₹ cr.)")
            or kv.get("anchor portion (₹ cr)")
            or next((v for k, v in kv.items() if "anchor portion" in k), None)
        )
        lockin_30d = parse_date(
            next((v for k, v in kv.items() if "50%" in k), None)
        )
        lockin_90d = parse_date(
            next((v for k, v in kv.items() if "remaining" in k), None)
        )
        return {
            "anchor_shares":   anchor_shares,
            "anchor_value_cr": anchor_value_cr,
            "lockin_30d":      lockin_30d,
            "lockin_90d":      lockin_90d,
        }
    return {}


def scrape_detail(url: str) -> dict:
    """
    Scrape an individual Chittorgarh IPO page (static HTML, no JS).

    The page has two distinct data regions:
      <li class="d-flex ..."> elements: dates (open, close, allotment, listing)
      <table> elements: issue size, anchor shares, lock-in dates, financials
    """
    soup = fetch_soup(url)
    if not soup:
        return {}

    # ── Company name ──
    company_name = ""
    for sel in ["h1", "h2"]:
        tag = soup.find(sel)
        if tag:
            txt = tag.get_text(strip=True)
            # Strip "IPO" and everything after it (e.g. "IPO Details", "IPO Review")
            name = re.sub(r"\s*\bIPO\b.*", "", txt, flags=re.IGNORECASE).strip()
            if name and len(name) > 3:
                company_name = name
                break
    if not company_name:
        title = soup.find("title")
        if title:
            txt = title.get_text(strip=True)
            company_name = re.sub(r"\s*[-–|].*", "", txt).strip()
            company_name = re.sub(r"\s*\bIPO\b.*", "", company_name, flags=re.IGNORECASE).strip()

    # ── Dates (from <li class="d-flex"> spans) ──
    li_dates = _extract_li_dates(soup)
    open_date      = li_dates.get("open_date")
    close_date     = li_dates.get("close_date")
    allotment_date = li_dates.get("allotment_date")
    listing_date   = li_dates.get("listing_date")

    # ── Fallback: parse "IPO Date" from Table 0 ──
    if not open_date or not close_date:
        tables = soup.find_all("table")
        for table in tables[:3]:
            for row in table.find_all("tr"):
                cells = row.find_all(["td", "th"])
                if len(cells) < 2:
                    continue
                k = cells[0].get_text(strip=True).lower()
                v = cells[1].get_text(strip=True)
                if "ipo date" in k or "issue date" in k:
                    # "12 to 16 Jul, 2024" or "31 Dec, 2024 to 2 Jan, 2025"
                    m = re.match(r"(\d+(?:\s+\w+,?\s*\d{4})?)\s+to\s+(.+)", v)
                    if m:
                        raw_open  = m.group(1).strip()
                        raw_close = m.group(2).strip()
                        # If open part has no month, inherit from close
                        if re.match(r"^\d+$", raw_open):
                            m2 = re.match(r"\d+\s+(\w+),?\s*(\d{4})", raw_close)
                            if m2:
                                raw_open = f"{raw_open} {m2.group(1)} {m2.group(2)}"
                        open_date  = open_date  or parse_date(raw_open)
                        close_date = close_date or parse_date(raw_close)
                elif "listed on" in k or "listing date" in k:
                    listing_date = listing_date or parse_date(v)

    # ── Issue size: extract ₹ Cr from "X shares (agg. up to ₹ Y Cr)" ──
    issue_size_cr = None
    tables = soup.find_all("table")
    for table in tables[:3]:
        for row in table.find_all("tr"):
            cells = row.find_all(["td", "th"])
            if len(cells) < 2:
                continue
            k = cells[0].get_text(strip=True).lower()
            v = cells[1].get_text(strip=True)
            if "total issue size" in k or ("issue size" in k and "lot" not in k):
                # Try direct parse first
                cr = parse_crore(v)
                if cr:
                    issue_size_cr = cr
                    break
                # Extract from "X shares (agg. up to ₹ Y Cr)"
                m = re.search(r"₹\s*([\d,\.]+)\s*(?:cr|crore)", v, re.IGNORECASE)
                if m:
                    issue_size_cr = parse_crore(m.group(1))
                    break
        if issue_size_cr:
            break

    # ── Issue price ──
    issue_price = None
    for table in tables[:2]:
        for row in table.find_all("tr"):
            cells = row.find_all(["td", "th"])
            if len(cells) < 2:
                continue
            k = cells[0].get_text(strip=True).lower()
            v = cells[1].get_text(strip=True)
            if "price band" in k or "issue price" in k or "offer price" in k:
                issue_price = v
                break
        if issue_price:
            break

    # ── Anchor data (from dedicated 2-column anchor table) ──
    anchor = _extract_anchor_table(soup)
    anchor_shares   = anchor.get("anchor_shares")
    anchor_value_cr = anchor.get("anchor_value_cr")
    lockin_30d      = anchor.get("lockin_30d")
    lockin_90d      = anchor.get("lockin_90d")

    # Anchor shares fallback: look in the investor category table
    if not anchor_shares:
        for table in tables:
            for row in table.find_all("tr"):
                cells = row.find_all(["td", "th"])
                if len(cells) >= 2:
                    k = cells[0].get_text(strip=True).lower()
                    v = cells[1].get_text(strip=True)
                    if "anchor investor shares offered" in k or "anchor investor shares" in k:
                        anchor_shares = parse_number(v)
                        break
            if anchor_shares:
                break

    # Fall back to calculating lock-in from allotment date
    if not (lockin_30d and lockin_90d):
        calc_30, calc_90 = calc_lockin(allotment_date)
        lockin_30d = lockin_30d or calc_30
        lockin_90d = lockin_90d or calc_90

    ref_date = open_date or allotment_date or listing_date
    return {
        "company_name":    company_name,
        "open_date":       open_date,
        "close_date":      close_date,
        "allotment_date":  allotment_date,
        "listing_date":    listing_date,
        "issue_size_cr":   issue_size_cr,
        "issue_price":     issue_price,
        "anchor_shares":   anchor_shares,
        "anchor_value_cr": anchor_value_cr,
        "lockin_30d":      lockin_30d,
        "lockin_90d":      lockin_90d,
        "fy":              determine_fy(ref_date),
        "detail_scraped":  1,
        "last_updated":    utcnow(),
    }


# ── Main pipeline ──────────────────────────────────────────────────────────────

def _blank_stub() -> dict:
    return dict(
        fy=None,
        open_date=None, close_date=None, allotment_date=None, listing_date=None,
        issue_size_cr=None, issue_price=None,
        anchor_shares=None, anchor_value_cr=None,
        lockin_30d=None, lockin_90d=None,
    )


def run_scrape(backfill: bool = False, list_only: bool = False) -> None:
    conn = open_db()

    # Step 1 — Collect IPO URLs from Chittorgarh list pages via Playwright
    seen_urls: set[str] = set()
    all_stubs: list[dict] = []

    for list_url, category in LIST_PAGES:
        stubs = scrape_list_with_playwright(list_url, category)
        for s in stubs:
            if s["source_url"] not in seen_urls:
                seen_urls.add(s["source_url"])
                all_stubs.append(s)
        time.sleep(REQUEST_DELAY)

    log.info("Total unique IPO links discovered: %d", len(all_stubs))

    # Insert stubs (COALESCE in upsert preserves existing detail data)
    for stub in all_stubs:
        stub.update(_blank_stub())
        upsert(conn, stub)

    if list_only:
        log.info("--list-only: skipping detail scraping.")
        conn.close()
        return

    # Step 2 — Scrape detail pages
    if backfill:
        rows = conn.execute(
            "SELECT source_url, company_name, category FROM ipos ORDER BY rowid"
        ).fetchall()
    else:
        rows = conn.execute("""
            SELECT source_url, company_name, category FROM ipos
            WHERE detail_scraped = 0
               OR (allotment_date IS NULL AND (open_date >= date('now','-90 days') OR open_date IS NULL))
            ORDER BY rowid
        """).fetchall()

    log.info("Detail pages to scrape: %d", len(rows))

    for i, row in enumerate(rows, 1):
        url = row["source_url"]
        log.info("[%d/%d] %s", i, len(rows), row["company_name"])
        detail = scrape_detail(url)
        if not detail:
            log.warning("  No data returned — skipping")
            time.sleep(REQUEST_DELAY)
            continue

        # Preserve existing company name if page didn't yield one
        if not detail.get("company_name"):
            detail["company_name"] = row["company_name"]

        detail["source_url"] = url
        detail["category"]   = row["category"]
        upsert(conn, detail)
        time.sleep(REQUEST_DELAY)

    conn.close()
    log.info("Scraping complete.")


def export_json() -> None:
    """Export SQLite → docs/data.json for the static website."""
    conn = open_db()
    rows = conn.execute("""
        SELECT
            source_url, company_name, category, fy,
            open_date, close_date, allotment_date, listing_date,
            issue_size_cr, issue_price,
            anchor_shares, anchor_value_cr,
            lockin_30d, lockin_90d, last_updated
        FROM ipos
        ORDER BY
            CASE WHEN open_date IS NULL THEN 1 ELSE 0 END,
            open_date DESC
    """).fetchall()
    conn.close()

    ipos = []
    today = datetime.now(tz=__import__("datetime").timezone.utc).date()

    for r in rows:
        d = dict(r)
        for field in ("lockin_30d", "lockin_90d"):
            if d.get(field):
                try:
                    expiry = datetime.strptime(d[field], "%Y-%m-%d").date()
                    delta  = (expiry - today).days
                    d[f"{field}_days"]   = delta
                    d[f"{field}_status"] = (
                        "expired"  if delta < 0 else
                        "critical" if delta <= 7 else
                        "warning"  if delta <= 30 else
                        "active"
                    )
                except ValueError:
                    d[f"{field}_days"]   = None
                    d[f"{field}_status"] = None
            else:
                d[f"{field}_days"]   = None
                d[f"{field}_status"] = None
        ipos.append(d)

    payload = {
        "generated_at":   utcnow(),
        "generated_date": today.isoformat(),
        "total":          len(ipos),
        "ipos":           ipos,
    }

    JSON_OUT.parent.mkdir(parents=True, exist_ok=True)
    JSON_OUT.write_text(json.dumps(payload, indent=2, default=str), encoding="utf-8")
    log.info("Exported %d IPOs → %s", len(ipos), JSON_OUT)


# ── Entry point ────────────────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser(description="IPO Anchor Lock-in Tracker")
    ap.add_argument("--backfill",  action="store_true", help="Re-scrape all IPO detail pages")
    ap.add_argument("--export",    action="store_true", help="Export JSON only — no scraping")
    ap.add_argument("--list-only", action="store_true", help="Collect list pages only, skip detail scraping")
    args = ap.parse_args()

    if args.export:
        export_json()
        return

    run_scrape(backfill=args.backfill, list_only=vars(args).get("list_only", False))
    export_json()
    log.info("Done.")


if __name__ == "__main__":
    main()
