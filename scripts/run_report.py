"""
PageSpeed -> Excel (via Power Automate) -> ClickUp comment automation.

Run on alternate weekdays via GitHub Actions. Reads config/urls.json, calls the
PageSpeed Insights API for each URL (mobile + desktop), pushes the rows to
an Excel table through a Power Automate HTTP-trigger flow, renders a styled
table image with Playwright, uploads it to a ClickUp task and posts a comment
containing the image URL (ClickUp shows a preview) plus audit details for every
page whose score is below SCORE_THRESHOLD (default 60).

Required environment variables:
  PAGESPEED_API_KEY
  CLICKUP_API_TOKEN
  CLICKUP_TASK_ID
  POWER_AUTOMATE_WEBHOOK_URL

Optional:
  MAX_WORKERS       concurrent PageSpeed calls (default 3)
  SCORE_THRESHOLD   report audit details for scores below this (default 60)
  MAX_AUDITS        audits listed per URL/device (default 5)
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

# How many PageSpeed calls to run at once. Lower to 1 if you hit 429s.
MAX_WORKERS = int(os.environ.get("MAX_WORKERS", "3"))

# Pages scoring below this get their failing audits listed in the comment.
SCORE_THRESHOLD = int(os.environ.get("SCORE_THRESHOLD", "60"))
MAX_AUDITS = int(os.environ.get("MAX_AUDITS", "5"))

# ClickUp comments have a size limit; stay safely under it.
MAX_COMMENT_CHARS = 9000

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
        safe_print(f"  [http {resp.status_code}] response body: {resp.text[:2000]}")
        resp.raise_for_status()


def extract_issues(data, limit=MAX_AUDITS):
    """Top failing performance audits, biggest estimated savings first."""
    audits = data["lighthouseResult"]["audits"]
    refs = data["lighthouseResult"]["categories"]["performance"]["auditRefs"]
    issues = []
    for ref in refs:
        if ref.get("group") not in ("load-opportunities", "diagnostics"):
            continue
        audit = audits.get(ref["id"], {})
        score = audit.get("score")
        if score is None or score >= 0.9:
            continue
        savings_ms = (audit.get("details") or {}).get("overallSavingsMs", 0) or 0
        issues.append(
            {
                "title": audit.get("title", ref["id"]),
                "display": audit.get("displayValue", ""),
                "savings_ms": round(savings_ms),
            }
        )
    issues.sort(key=lambda i: i["savings_ms"], reverse=True)
    return issues[:limit]


def fetch_pagespeed(url, strategy, attempts=3):
    api_url = "https://www.googleapis.com/pagespeedonline/v5/runPagespeed"
    params = {"url": url, "strategy": strategy, "key": PAGESPEED_API_KEY}
    last_error = None
    for attempt in range(1, attempts + 1):
        backoff = 5 * attempt
        try:
            resp = requests.get(api_url, params=params, timeout=90)
            if resp.status_code == 429 or resp.status_code >= 500:
                backoff = 15 * attempt  # rate limited / server error: wait longer
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
                "issues": extract_issues(data),
            }
        except Exception as exc:  # noqa: BLE001 - want to retry on any failure
            last_error = exc
            safe_print(f"  [warn] {strategy} attempt {attempt} failed for {url}: {exc}")
            if attempt < attempts:
                time.sleep(backoff)
    raise RuntimeError(f"PageSpeed failed for {url} ({strategy}): {last_error}")


# None marks a metric that failed after all retries. Formatted per
# destination: -1 (a value that can't occur naturally) for Excel's numeric
# columns, "ERR" for the human-readable image.
ERROR_METRICS = {
    "score": None, "fcp": None, "lcp": None, "tbt": None, "cls": None, "si": None,
    "issues": [],
}


def for_excel(v):
    return -1 if v is None else v


def for_display(v):
    return "ERR" if v is None else v


def build_rows(urls, date_str):
    # Fire mobile+desktop for every URL as concurrent tasks (bounded by
    # MAX_WORKERS). Each call keeps its own retry/backoff logic. A URL that
    # still fails after all retries is recorded as an "ERR" row so the rest of
    # the run (Excel write / image / ClickUp post) isn't thrown away.
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


def build_improvements(rows):
    """Audit details for every URL/device whose score is below SCORE_THRESHOLD."""
    lines = []
    for r in rows:
        for label in ("mobile", "desktop"):
            m = r[label]
            if m["score"] is None or m["score"] >= SCORE_THRESHOLD:
                continue
            lines.append(f"\n{r['url']} - {label.title()} (score {m['score']})")
            if not m["issues"]:
                lines.append("  - No audit details available")
            for i in m["issues"]:
                extras = []
                if i["savings_ms"]:
                    extras.append(f"saves ~{i['savings_ms'] / 1000:.1f}s")
                if i["display"]:
                    extras.append(i["display"])
                suffix = f" ({', '.join(extras)})" if extras else ""
                lines.append(f"  - {i['title']}{suffix}")
    if not lines:
        return ""
    return f"\n\nAudits to fix (score below {SCORE_THRESHOLD}):" + "\n".join(lines)


def upload_clickup_attachment(date_str):
    """Upload report.png to the task. Returns the file URL (or None)."""
    with open(OUTPUT_IMAGE, "rb") as f:
        resp = requests.post(
            f"https://api.clickup.com/api/v2/task/{CLICKUP_TASK_ID}/attachment",
            headers={"Authorization": CLICKUP_API_TOKEN},
            files={"attachment": (f"PageSpeed-{date_str}.png", f, "image/png")},
            timeout=60,
        )
    raise_with_body(resp)
    print(f"ClickUp attachment uploaded: {resp.status_code}")
    try:
        return resp.json().get("url")
    except ValueError:
        return None


def post_clickup_comment(date_str, failed_tasks, image_url=None, improvements=""):
    text = f"PageSpeed Report - {date_str}"
    if image_url:
        # URL on its own line so ClickUp can unfurl it into an image preview.
        text += f"\n\n{image_url}"
    if failed_tasks:
        text += f"\n\n({len(failed_tasks)} check(s) failed after retries and are marked ERR - see GitHub Actions log)"
    text += improvements
    if len(text) > MAX_COMMENT_CHARS:
        text = text[:MAX_COMMENT_CHARS] + "\n... (truncated)"
    resp = requests.post(
        f"https://api.clickup.com/api/v2/task/{CLICKUP_TASK_ID}/comment",
        headers={"Authorization": CLICKUP_API_TOKEN, "Content-Type": "application/json"},
        json={"comment_text": text, "notify_all": False},
        timeout=30,
    )
    raise_with_body(resp)
    print(f"ClickUp comment posted: {resp.status_code}")


def main():
    date_str = datetime.now(IST).strftime("%d-%b-%Y")
    urls = load_urls()
    print(
        f"Running PageSpeed report for {len(urls)} URLs on {date_str} "
        f"(workers={MAX_WORKERS}, threshold={SCORE_THRESHOLD})"
    )

    rows, failed_tasks = build_rows(urls, date_str)
    push_to_excel(rows, date_str)
    render_image(rows, date_str)

    # Upload first so the comment can include the image URL. If the upload
    # fails, still post the text comment and keep the error visible.
    image_url = None
    upload_failed = False
    try:
        image_url = upload_clickup_attachment(date_str)
    except Exception as exc:  # noqa: BLE001
        upload_failed = True
        safe_print(f"[warn] ClickUp attachment upload failed: {exc}")

    post_clickup_comment(date_str, failed_tasks, image_url, build_improvements(rows))

    if upload_failed:
        # Fail the workflow run so you get GitHub's failure notification.
        raise RuntimeError("Comment posted, but the image attachment upload failed")

    print("Done.")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:  # noqa: BLE001
        print(f"FATAL: {exc}", file=sys.stderr)
        sys.exit(1)
