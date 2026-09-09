#!/usr/bin/env python3
"""
Chartink Screener Watchdog
───────────────────────────
Runs after the 4 Star Buy, 5 Star Buy (RSI) and Stage 2 screeners' primary +
backup cron triggers should have fired. GitHub Actions occasionally drops
scheduled ("cron") triggers under platform load, so this checks the Actions
run history for a successful run today; if any screener never ran, it
re-triggers it via workflow_dispatch and emails an alert so the miss
doesn't go unnoticed again.

The Stage 2 screener lives in a separate repo (chartink-screener-alert), so
GH_TOKEN must be a token with actions:write on BOTH this repo and that one
(a fine-grained PAT) rather than the default same-repo GITHUB_TOKEN.

Env vars (or GitHub Secrets):
  GH_TOKEN             GitHub token with actions:write on this repo AND
                        learningblueoceancapital-png/chartink-screener-alert
  GMAIL_USER           your Gmail address
  GMAIL_APP_PASSWORD   16-char Gmail App Password
"""

import os
import json
import smtplib
import ssl
import subprocess
import datetime
import logging
from email.message import EmailMessage

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger("watchdog")

REPO           = os.getenv("GITHUB_REPOSITORY", "")
GMAIL_USER     = os.getenv("GMAIL_USER", "")
GMAIL_PASSWORD = os.getenv("GMAIL_APP_PASSWORD", "")

EMAIL_TO = [
    "lavneesh@blueoceancapital.co.in",
    "param.parekh@blueoceancapital.co.in",
    "kinjal.shah@blueoceancapital.co.in",
]

WORKFLOWS = [
    (REPO, "chartink_4star_buy.yml"),
    (REPO, "chartink_screener.yml"),
    ("learningblueoceancapital-png/chartink-screener-alert", "stage2_screener.yml"),
]
LOOKBACK_HOURS = 6  # watchdog fires well after all screeners' primary+backup triggers


def ran_successfully_today(repo: str, workflow: str) -> bool:
    out = subprocess.run(
        ["gh", "run", "list", "--repo", repo, "--workflow", workflow, "--status", "success",
         "--limit", "5", "--json", "createdAt"],
        capture_output=True, text=True, check=True,
    ).stdout
    now = datetime.datetime.now(datetime.timezone.utc)
    for run in json.loads(out):
        created = datetime.datetime.fromisoformat(run["createdAt"].replace("Z", "+00:00"))
        if (now - created).total_seconds() < LOOKBACK_HOURS * 3600:
            return True
    return False


def trigger(repo: str, workflow: str) -> None:
    subprocess.run(["gh", "workflow", "run", "--repo", repo, workflow], check=True)
    log.info("Re-triggered %s (%s)", workflow, repo)


def send_alert(missed: list[str]) -> None:
    if not GMAIL_USER or not GMAIL_PASSWORD:
        log.warning("GMAIL_USER/GMAIL_APP_PASSWORD not set — skipping alert email")
        return

    today = datetime.datetime.now(datetime.timezone.utc).strftime("%d %b %Y")
    lines = [f"  - {w}" for w in missed]

    msg = EmailMessage()
    msg["Subject"] = f"Chartink Watchdog Alert | {today} | {len(missed)} screener(s) missed"
    msg["From"]    = GMAIL_USER
    msg["To"]      = ", ".join(EMAIL_TO)
    msg.set_content(
        "GitHub Actions' scheduled trigger appears to have been dropped for the "
        f"following screener(s) on {today}:\n\n"
        + "\n".join(lines)
        + "\n\nThey have been re-triggered automatically just now — check your "
          "inbox shortly for the results.\n\n"
        + (f"Run history: https://github.com/{REPO}/actions\n\n" if REPO else "\n")
        + "Regards,\nChartink Watchdog"
    )

    ctx = ssl.create_default_context()
    with smtplib.SMTP("smtp.gmail.com", 587) as smtp:
        smtp.ehlo()
        smtp.starttls(context=ctx)
        smtp.login(GMAIL_USER, GMAIL_PASSWORD)
        smtp.send_message(msg)
    log.info("Alert email sent to %s", ", ".join(EMAIL_TO))


def main() -> None:
    missed = []
    for repo, wf in WORKFLOWS:
        if ran_successfully_today(repo, wf):
            log.info("%s (%s): OK — successful run within last %sh", wf, repo, LOOKBACK_HOURS)
        else:
            log.warning("%s (%s): MISSED — no successful run in last %sh", wf, repo, LOOKBACK_HOURS)
            missed.append(f"{wf} ({repo})")
            trigger(repo, wf)

    if missed:
        send_alert(missed)
    else:
        log.info("All screeners ran on schedule today. No action needed.")


if __name__ == "__main__":
    main()

