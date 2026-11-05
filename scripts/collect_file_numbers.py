#!/usr/bin/env python3
"""Collect unique Federal Court file numbers for Competition Law matters.

Queries the Commonwealth Courts Portal federal e-search in date windows of at
most 90 days (the portal's limit), follows pagination, and writes the unique
file numbers to a JSON file. Handles the "scheduled maintenance" notice page
that sometimes appears on first visit by submitting its "continue" form.
"""
import argparse
import html
import http.cookiejar
import json
import re
import sys
import time
import urllib.parse
import urllib.request
from datetime import date, datetime, timedelta, timezone

BASE = "https://www.comcourts.gov.au"
QUERY_PATH = "/public/esearch/federal/query"
WINDOW_DAYS = 90  # inclusive span must be <= 90 days
USER_AGENT = "Mozilla/5.0 (compatible; competition-law-judgments collector)"

FILE_RE = re.compile(r'<a class="fileName" href="[^"]*">\s*([^<]+?)\s*</a>')
NEXT_RE = re.compile(r'<a class="next_page"[^>]*href="([^"]+)"')
WARN_FORM_RE = re.compile(r'<form action="\s*([^"]*warned[^"]*)"', re.I)
TOKEN_RE = re.compile(r'name="authenticity_token"[^>]*value="([^"]+)"')


def is_results_page(body: str) -> bool:
    return "Search results" in body or "No files found" in body


class Client:
    def __init__(self, delay=1.0):
        self.jar = http.cookiejar.CookieJar()
        self.opener = urllib.request.build_opener(
            urllib.request.HTTPCookieProcessor(self.jar)
        )
        self.opener.addheaders = [("User-Agent", USER_AGENT)]
        self.delay = delay

    def _open(self, url, data=None, retries=4):
        for attempt in range(retries + 1):
            try:
                with self.opener.open(url, data=data, timeout=60) as r:
                    body = r.read().decode("utf-8", errors="replace")
                time.sleep(self.delay)
                return body
            except Exception as e:  # noqa: BLE001
                if attempt == retries:
                    raise
                wait = 2 ** (attempt + 1)
                print(f"  retry {attempt + 1} after error: {e} (sleep {wait}s)", file=sys.stderr)
                time.sleep(wait)

    def get(self, url):
        """GET a page, clicking through the maintenance notice if shown."""
        body = self._open(url)
        for _ in range(3):
            m = WARN_FORM_RE.search(body)
            if not m or is_results_page(body):
                return body
            action = urllib.parse.urljoin(BASE, m.group(1).strip())
            tok = TOKEN_RE.search(body)
            data = urllib.parse.urlencode(
                {"authenticity_token": html.unescape(tok.group(1)) if tok else ""}
            ).encode()
            print("  maintenance notice detected; clicking continue", file=sys.stderr)
            self._open(action, data=data)
            body = self._open(url)
        return body


def query_url(start: date, end: date) -> str:
    params = {
        "last_name": "",
        "given_name": "",
        "filed_after": start.strftime("%d/%m/%Y"),
        "filed_before": end.strftime("%d/%m/%Y"),
        "court": "any",
        "file_status": "any",
        "registry": "any",
        "action_type": "Competition_Law",
        "search_by": "party_name",
    }
    return BASE + QUERY_PATH + "?" + urllib.parse.urlencode(params)


def windows(start: date, end: date, days: int = WINDOW_DAYS):
    cur = start
    while cur <= end:
        w_end = min(cur + timedelta(days=days - 1), end)
        yield cur, w_end
        cur = w_end + timedelta(days=1)


def collect_window(client: Client, start: date, end: date) -> set:
    found = set()
    url = query_url(start, end)
    page = 1
    seen_urls = set()
    while url and url not in seen_urls:
        seen_urls.add(url)
        body = client.get(url)
        if "No files found" in body:
            break
        if "Search results" not in body:
            raise RuntimeError(f"Unexpected response for {url}")
        nums = {html.unescape(n).strip() for n in FILE_RE.findall(body)}
        found |= nums
        nxt = NEXT_RE.search(body)
        url = urllib.parse.urljoin(BASE, html.unescape(nxt.group(1))) if nxt else None
        page += 1
    print(f"{start} -> {end}: {len(found)} file numbers ({page - 1} page(s))", file=sys.stderr)
    return found


def load_existing(path):
    try:
        with open(path) as f:
            return json.load(f)
    except FileNotFoundError:
        return None


def main():
    ap = argparse.ArgumentParser(
        description="By default only searches for new filings since the last run "
        "(re-checking the last --overlap-days) and merges them into --output."
    )
    ap.add_argument("--start", help="YYYY-MM-DD: search from this date instead of incrementally")
    ap.add_argument("--end", help="YYYY-MM-DD (default: today)")
    ap.add_argument("--years", type=float, help="search this many years back instead of incrementally")
    ap.add_argument("--overlap-days", type=int, default=30,
                    help="in incremental mode, re-search this many days before the last run's end date")
    ap.add_argument("--default-start", default="1990-01-01",
                    help="start date used when there is no existing output file")
    ap.add_argument("--replace", action="store_true",
                    help="overwrite the output instead of merging with existing file numbers")
    ap.add_argument("--output", default="data/competition_law_file_numbers.json")
    ap.add_argument("--delay", type=float, default=1.0)
    args = ap.parse_args()

    existing = None if args.replace else load_existing(args.output)
    end = date.fromisoformat(args.end) if args.end else date.today()
    if args.start:
        start = date.fromisoformat(args.start)
    elif args.years:
        start = end - timedelta(days=round(365.25 * args.years))
    elif existing:
        start = date.fromisoformat(existing["filed_before"]) - timedelta(days=args.overlap_days)
    else:
        start = date.fromisoformat(args.default_start)
    print(f"Searching {start} -> {end}", file=sys.stderr)

    client = Client(delay=args.delay)
    all_nums = set()
    for w_start, w_end in windows(start, end):
        all_nums |= collect_window(client, w_start, w_end)

    covered_from, covered_to = start, end
    if existing:
        before = set(existing["file_numbers"])
        new = all_nums - before
        print(f"{len(new)} new file numbers: {sorted(new)}", file=sys.stderr)
        all_nums |= before
        covered_from = min(start, date.fromisoformat(existing["filed_after"]))
        covered_to = max(end, date.fromisoformat(existing["filed_before"]))

    def sort_key(n):
        m = re.match(r"([A-Z]+)(\d+)/(\d{4})", n)
        return (int(m.group(3)), m.group(1), int(m.group(2))) if m else (0, n, 0)

    out = {
        "source": BASE + QUERY_PATH,
        "action_type": "Competition_Law",
        "filed_after": covered_from.isoformat(),
        "filed_before": covered_to.isoformat(),
        "generated_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "count": len(all_nums),
        "file_numbers": sorted(all_nums, key=sort_key),
    }
    with open(args.output, "w") as f:
        json.dump(out, f, indent=2)
        f.write("\n")
    print(f"Wrote {len(all_nums)} unique file numbers to {args.output}", file=sys.stderr)


if __name__ == "__main__":
    main()
