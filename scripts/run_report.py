"""
PageSpeed -> Excel (via Power Automate) -> ClickUp comment automation.

Run on Mon/Wed/Fri via GitHub Actions. Reads config/urls.json, calls the
PageSpeed Insights API for each URL (mobile + desktop), pushes the rows to
an Excel table through a Power Automate HTTP-trigger flow, renders a styled
table image with Playwright, and posts it to a ClickUp task as a comment +
attachment.

Required environment variables:
  PAGESPEED_API_KEY
  CLICKUP_API_TOKEN
  CLICKUP_TASK_ID
  POWER_AUTOMATE_WEBHOOK_URL
"""

import json
import os
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone, timedelta

import requests
from playwright.sync_api import sync_playwright

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
URLS_FILE = os.path.join(ROOT, "config", "urls.json")
OUTPUT_IMAGE = os.path.join(ROOT, "report.png")

PAGESPEED_API_KEY = os.environ["PAGESPEED_API_KEY"]
CLICKUP_API_TOKEN = os.environ["CLICKUP_API_TOKEN"]
CLICKUP_TASK_ID = os.environ["CLICKUP_TASK_ID"]
POWER_AUTOMATE_WEBHOOK_URL = os.environ["POWER_AUTOMATE_WEBHOOK_URL"]

# How many PageSpeed calls to run at once. PageSpeed's own quota comfortably
# allows this; raise cautiously if you want it faster still.
MAX_WORKERS = 1

# Adjust to your local timezone for the date shown in the report/comment.
IST = timezone(timedelta(hours=5, minutes=30))

_print_lock = threading.Lock()


def safe_print(msg):
    with _print_lock:
        print(msg)


def load_urls():
    with open(URLS_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def raise_with_body(resp):
    """Like resp.raise_for_status(), but prints the response body first so
    failures are actionable in CI logs instead of a bare 'Bad Request'."""
    if not resp.ok:
        print(f"  [http {resp.status_code}] response body: {resp.text[:2000]}")
        resp.raise_for_status()


def fetch_pagespeed(url, strategy, attempts=3):
    api_url = "https://www.googleapis.com/pagespeedonline/v5/runPagespeed"
    params = {"url": url, "strategy": strategy, "key": PAGESPEED_API_KEY}
    last_error = None
    for attempt in range(1, attempts + 1):
        try:
            resp = requests.get(api_url, params=params, timeout=60)
            raise_with_body(resp)
            data = resp.json()
            audits = data["lighthouseResult"]["audits"]
            score = data["lighthouseResult"]["categories"]["performance"]["score"]
            return {
                "score": round(score * 100),
                "fcp": round(audits["first-contentful-paint"]["numericValue"] / 1000, 2),
                "lcp": round(audits["largest-contentful-paint"]["numericValue"] / 1000, 2),
                "tbt": round(audits["total-blocking-time"]["numericValue"]),
                "cls": round(audits["cumulative-layout-shift"]["numericValue"], 3),
                "si": round(audits["speed-index"]["numericValue"] / 1000, 2),
            }
        except Exception as exc:  # noqa: BLE001 - want to retry on any failure
            last_error = exc
            safe_print(f"  [warn] {strategy} attempt {attempt} failed for {url}: {exc}")
            if attempt < attempts:
                time.sleep(5 * attempt)
    raise RuntimeError(f"PageSpeed failed for {url} ({strategy}): {last_error}")


# None marks a metric that failed after all retries. Formatted per
# destination: -1 (a value that can't occur naturally) for Excel's numeric
# columns, "ERR" for the human-readable image.
ERROR_METRICS = {"score": None, "fcp": None, "lcp": None, "tbt": None, "cls": None, "si": None}


def for_excel(v):
    return -1 if v is None else v


def for_display(v):
    return "ERR" if v is None else v


def build_rows(urls, date_str):
    # Fire mobile+desktop for every URL as concurrent tasks (bounded by
    # MAX_WORKERS). Each individual call still goes through fetch_pagespeed's
    # own retry/backoff logic (the failsafe) - concurrency only affects how
    # many of those retry-protected calls are in flight at once.
    #
    # A URL that still fails after all retries no longer aborts the whole
    # run - it's recorded as an "ERR" row so the other ~49 successful calls
    # (and the Excel write / image / ClickUp post) aren't thrown away over
    # one flaky PageSpeed request.
    tasks = [(url, strategy) for url in urls for strategy in ("mobile", "desktop")]
    results = {}
    completed = 0
    failed_tasks = []

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        future_to_task = {
            executor.submit(fetch_pagespeed, url, strategy): (url, strategy)
            for url, strategy in tasks
        }
        for future in as_completed(future_to_task):
            url, strategy = future_to_task[future]
            try:
                results[(url, strategy)] = future.result()
            except Exception as exc:  # noqa: BLE001 - degrade, don't abort the run
                safe_print(f"  [error] giving up on {strategy} for {url}: {exc}")
                results[(url, strategy)] = ERROR_METRICS
                failed_tasks.append((url, strategy))
            completed += 1
            safe_print(f"[{completed}/{len(tasks)}] done: {strategy} - {url}")

    if failed_tasks:
        safe_print(f"WARNING: {len(failed_tasks)} call(s) failed after retries and are marked ERR:")
        for url, strategy in failed_tasks:
            safe_print(f"  - {strategy}: {url}")

    rows = [
        {
            "date": date_str,
            "url": url,
            "mobile": results[(url, "mobile")],
            "desktop": results[(url, "desktop")],
        }
        for url in urls
    ]
    return rows, failed_tasks


def push_to_excel(rows, date_str):
    payload = {
        "rows": [
            {
                "date": date_str,
                "url": r["url"],
                "m_score": for_excel(r["mobile"]["score"]),
                "m_fcp": for_excel(r["mobile"]["fcp"]),
                "m_lcp": for_excel(r["mobile"]["lcp"]),
                "m_tbt": for_excel(r["mobile"]["tbt"]),
                "m_cls": for_excel(r["mobile"]["cls"]),
                "m_si": for_excel(r["mobile"]["si"]),
                "d_score": for_excel(r["desktop"]["score"]),
                "d_fcp": for_excel(r["desktop"]["fcp"]),
                "d_lcp": for_excel(r["desktop"]["lcp"]),
                "d_tbt": for_excel(r["desktop"]["tbt"]),
                "d_cls": for_excel(r["desktop"]["cls"]),
                "d_si": for_excel(r["desktop"]["si"]),
            }
            for r in rows
        ]
    }
    resp = requests.post(POWER_AUTOMATE_WEBHOOK_URL, json=payload, timeout=120)
    raise_with_body(resp)
    print(f"Excel webhook: {resp.status_code}")


def render_html(rows, date_str):
    def row_html(r):
        m, d = r["mobile"], r["desktop"]
        return f"""
        <tr>
          <td>{r['date']}</td>
          <td class="url"><a href="{r['url']}">{r['url']}</a></td>
          <td>{for_display(m['score'])}</td><td>{for_display(m['fcp'])}</td><td>{for_display(m['lcp'])}</td>
          <td>{for_display(m['tbt'])}</td><td>{for_display(m['cls'])}</td><td>{for_display(m['si'])}</td>
          <td>{for_display(d['score'])}</td><td>{for_display(d['fcp'])}</td><td>{for_display(d['lcp'])}</td>
          <td>{for_display(d['tbt'])}</td><td>{for_display(d['cls'])}</td><td>{for_display(d['si'])}</td>
        </tr>"""

    rows_html = "\n".join(row_html(r) for r in rows)

    return f"""
    <html>
    <head>
    <style>
      body {{ font-family: Arial, Helvetica, sans-serif; margin: 0; padding: 12px; background: #fff; }}
      table {{ border-collapse: collapse; font-size: 12px; }}
      th, td {{ border: 1px solid #cccccc; padding: 4px 6px; text-align: center; white-space: nowrap; }}
      td.url {{ text-align: left; max-width: 420px; white-space: normal; }}
      td.url a {{ color: #1155cc; text-decoration: underline; }}
      thead th.group-mobile {{ background: #fce4b4; }}
      thead th.group-desktop {{ background: #c9daf8; }}
      thead th {{ font-weight: bold; }}
    </style>
    </head>
    <body>
    <table id="report">
      <thead>
        <tr>
          <th rowspan="2">Date</th>
          <th rowspan="2">Page URLs</th>
          <th colspan="6" class="group-mobile">Mobile</th>
          <th colspan="6" class="group-desktop">Desktop</th>
        </tr>
        <tr>
          <th class="group-mobile">Score</th><th class="group-mobile">FCP (s)</th>
          <th class="group-mobile">LCP (s)</th><th class="group-mobile">Total Blocking Time (ms)</th>
          <th class="group-mobile">CLS</th><th class="group-mobile">Speed Index (s)</th>
          <th class="group-desktop">Score</th><th class="group-desktop">FCP (s)</th>
          <th class="group-desktop">LCP (s)</th><th class="group-desktop">Total Blocking Time (ms)</th>
          <th class="group-desktop">CLS</th><th class="group-desktop">Speed Index (s)</th>
        </tr>
      </thead>
      <tbody>
        {rows_html}
      </tbody>
    </table>
    </body>
    </html>
    """


def render_image(rows, date_str):
    html = render_html(rows, date_str)
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()
        page.set_content(html)
        page.locator("#report").screenshot(path=OUTPUT_IMAGE)
        browser.close()
    print(f"Rendered image: {OUTPUT_IMAGE}")


def post_clickup_comment(date_str, failed_tasks):
    text = f"PageSpeed Report - {date_str}"
    if failed_tasks:
        text += f"\n({len(failed_tasks)} check(s) failed after retries and are marked ERR - see GitHub Actions log)"
    resp = requests.post(
        f"https://api.clickup.com/api/v2/task/{CLICKUP_TASK_ID}/comment",
        headers={"Authorization": CLICKUP_API_TOKEN, "Content-Type": "application/json"},
        json={"comment_text": text, "notify_all": False},
        timeout=30,
    )
    raise_with_body(resp)
    print(f"ClickUp comment posted: {resp.status_code}")


def upload_clickup_attachment():
    with open(OUTPUT_IMAGE, "rb") as f:
        resp = requests.post(
            f"https://api.clickup.com/api/v2/task/{CLICKUP_TASK_ID}/attachment",
            headers={"Authorization": CLICKUP_API_TOKEN},
            files={"attachment": ("report.png", f, "image/png")},
            timeout=60,
        )
    raise_with_body(resp)
    print(f"ClickUp attachment uploaded: {resp.status_code}")


def main():
    date_str = datetime.now(IST).strftime("%d-%b-%Y")
    urls = load_urls()
    print(f"Running PageSpeed report for {len(urls)} URLs on {date_str}")

    rows, failed_tasks = build_rows(urls, date_str)
    push_to_excel(rows, date_str)
    render_image(rows, date_str)
    post_clickup_comment(date_str, failed_tasks)
    upload_clickup_attachment()

    print("Done.")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:  # noqa: BLE001
        print(f"FATAL: {exc}", file=sys.stderr)
        sys.exit(1)