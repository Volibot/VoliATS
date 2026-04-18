"""
Local scheduler — polls every POLL_INTERVAL_MINUTES (default 5).
Use this if you want to run the extractor on your own server instead of GitHub Actions.

Usage:
    POLL_INTERVAL_MINUTES=5 python scheduler.py
"""

import os
import time
import logging
import schedule
from email_extractor import process_emails

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
log = logging.getLogger(__name__)

INTERVAL = int(os.getenv("POLL_INTERVAL_MINUTES", "5"))


def job() -> None:
    try:
        process_emails()
    except Exception as exc:
        log.error(f"Extractor run failed: {exc}", exc_info=True)


if __name__ == "__main__":
    log.info(f"Scheduler started — polling every {INTERVAL} minute(s).")
    job()  # run immediately on startup
    schedule.every(INTERVAL).minutes.do(job)
    while True:
        schedule.run_pending()
        time.sleep(30)
