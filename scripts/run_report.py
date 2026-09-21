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
import time
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

# Adjust to your local timezone for the date shown in the report/comment.
IST = timezone(timedelta(hours=5, minutes=30))


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
            print(f"  [warn] {strategy} attempt {attempt} failed for {url}: {exc}")
            if attempt < attempts:
                time.sleep(5 * attempt)
    raise RuntimeError(f"PageSpeed failed for {url} ({strategy}): {last_error}")


def build_rows(urls, date_str):
    rows = []
    for i, url in enumerate(urls, start=1):
        print(f"[{i}/{len(urls)}] {url}")
        mobile = fetch_pagespeed(url, "mobile")
        time.sleep(1)
        desktop = fetch_pagespeed(url, "desktop")
        time.sleep(1)
        rows.append({"date": date_str, "url": url, "mobile": mobile, "desktop": desktop})
    return rows


def push_to_excel(rows, date_str):
    payload = {
        "rows": [
            {
                "date": date_str,
                "url": r["url"],
                "m_score": r["mobile"]["score"],
                "m_fcp": r["mobile"]["fcp"],
                "m_lcp": r["mobile"]["lcp"],
                "m_tbt": r["mobile"]["tbt"],
                "m_cls": r["mobile"]["cls"],
                "m_si": r["mobile"]["si"],
                "d_score": r["desktop"]["score"],
                "d_fcp": r["desktop"]["fcp"],
                "d_lcp": r["desktop"]["lcp"],
                "d_tbt": r["desktop"]["tbt"],
                "d_cls": r["desktop"]["cls"],
                "d_si": r["desktop"]["si"],
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
          <td>{m['score']}</td><td>{m['fcp']}</td><td>{m['lcp']}</td>
          <td>{m['tbt']}</td><td>{m['cls']}</td><td>{m['si']}</td>
          <td>{d['score']}</td><td>{d['fcp']}</td><td>{d['lcp']}</td>
          <td>{d['tbt']}</td><td>{d['cls']}</td><td>{d['si']}</td>
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


def post_clickup_comment(date_str):
    resp = requests.post(
        f"https://api.clickup.com/api/v2/task/{CLICKUP_TASK_ID}/comment",
        headers={"Authorization": CLICKUP_API_TOKEN, "Content-Type": "application/json"},
        json={"comment_text": f"PageSpeed Report - {date_str}", "notify_all": False},
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

    rows = build_rows(urls, date_str)
    push_to_excel(rows, date_str)
    render_image(rows, date_str)
    post_clickup_comment(date_str)
    upload_clickup_attachment()

    print("Done.")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:  # noqa: BLE001
        print(f"FATAL: {exc}", file=sys.stderr)
        sys.exit(1)
