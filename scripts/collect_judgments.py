#!/usr/bin/env python3
"""Find and download the published judgments for each Competition Law file.

Step 1 (plain HTTP, reusing collect_file_numbers.Client): for each file
number in data/competition_law_file_numbers.json, fetch the file's public
"List of orders" page on the Commonwealth Courts Portal

    https://www.comcourts.gov.au/file/FEDERAL/P/VID71/2023/order_list

and pick out every event with a "View judgment" link,
javascript:openFCAJudgment('2023', 'FCA', '665'). The portal's own JS turns
those into judgment URLs (see judgment_url). The results go to
data/judgments.json.

Step 2 (real Chrome): judgments.fedcourt.gov.au is behind a Cloudflare
challenge, so each judgment's "Original Word Document" is fetched the same
way nwbort/tribunal-tracker does it: drive Chrome via nodriver (headful
under Xvfb) and fetch the file from inside the page (see fedcourt_browser).
Files land in data/judgments/<year>/ under the Federal Court's own file
name (e.g. 2023FCA0665.docx). Judgments already downloaded are skipped.

Usage:
    xvfb-run -a python scripts/collect_judgments.py            # everything
    python scripts/collect_judgments.py --no-download          # step 1 only
    xvfb-run -a python scripts/collect_judgments.py --only VID71/2023
"""
import argparse
import asyncio
import glob
import html
import json
import os
import re
import sys
import time
import urllib.parse
from datetime import datetime, timezone

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from collect_file_numbers import BASE, Client  # noqa: E402

ORDER_LIST_URL = BASE + "/file/FEDERAL/P/{file_number}/order_list"
FEDCOURT_JUDGMENTS = "https://www.judgments.fedcourt.gov.au/judgments/Judgments"
AUSTLII_CASES = "https://www.austlii.edu.au/au/cases/cth"

# javascript:openFCAJudgment('2023', 'FCA', '665') (args: year, court,
# number). openAustliiJudgment takes the same arguments.
JUDGMENT_LINK_RE = re.compile(
    r"open(FCA|Austlii)Judgment\(\s*'(\d{4})'\s*,\s*'([A-Za-z]+)'\s*,\s*'(\d+)'\s*\)"
)
EVENT_ROW_RE = re.compile(r'<tr class="spanned clickable-row"')
CELL_RE = re.compile(r"<td\b[^>]*>(.*?)</td>", re.S)
DATE_RE = re.compile(r"^\d{1,2}-[A-Za-z]{3}-\d{4}$")
NEXT_RE = re.compile(r'<a class="next_page"[^>]*href="([^"]+)"')
ORDER_LIST_MARKER = '<h1 class="page_title">List of orders</h1>'


def log(msg):
    print(msg, file=sys.stderr, flush=True)


def judgment_url(launcher: str, year: str, court: str, number: str) -> str:
    """Mirror the portal's ccp_application JS: openFCAJudgment sends FCA and
    FCAFC to the Federal Court's site (number zero-padded to 4) and anything
    else to AustLII, which is all openAustliiJudgment does."""
    c = court.lower()
    if launcher == "FCA" and c == "fca":
        return f"{FEDCOURT_JUDGMENTS}/fca/single/{year}/{year}fca{number.zfill(4)}"
    if launcher == "FCA" and c == "fcafc":
        return f"{FEDCOURT_JUDGMENTS}/fca/full/{year}/{year}fcafc{number.zfill(4)}"
    return f"{AUSTLII_CASES}/{court}/{year}/{number}.html"


def text_of(cell: str) -> str:
    return re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", " ", cell))).strip()


def iso_date(text: str) -> str:
    try:
        return datetime.strptime(text, "%d-%b-%Y").date().isoformat()
    except ValueError:
        return text


def parse_order_list(body: str) -> list[dict]:
    """Every judgment linked from a List of orders page, in page order."""
    found = []
    # Each event starts with a clickable row; its links (orders, judgment)
    # sit in that row's last cell, before the next event's row.
    chunks = EVENT_ROW_RE.split(body)[1:]
    for chunk in chunks:
        chunk = html.unescape(chunk)
        first_row = chunk.split("</tr>", 1)[0]
        cells = [text_of(c) for c in CELL_RE.findall(first_row)]
        event_date = event = None
        for i, c in enumerate(cells):
            if DATE_RE.match(c):
                event_date = iso_date(c)
                # date, time, event type, judge
                event = cells[i + 2] if i + 2 < len(cells) else None
                break
        for launcher, year, court, number in JUDGMENT_LINK_RE.findall(chunk):
            court = court.upper()
            found.append(
                {
                    "citation": f"[{year}] {court} {int(number)}",
                    "court": court,
                    "year": int(year),
                    "number": int(number),
                    "url": judgment_url(launcher, year, court, number),
                    "event_date": event_date,
                    "event": event,
                }
            )
    return found


def scan_file(client: Client, file_number: str) -> list[dict]:
    url = ORDER_LIST_URL.format(file_number=file_number)
    judgments, seen = [], set()
    while url and url not in seen:
        seen.add(url)
        body = client.get(url)
        if ORDER_LIST_MARKER not in body:
            raise RuntimeError(f"no List of orders at {url} (login page or error)")
        judgments += parse_order_list(body)
        nxt = NEXT_RE.search(body)
        url = urllib.parse.urljoin(url, html.unescape(nxt.group(1))) if nxt else None
    # One entry per judgment per file (a judgment can hang off two events).
    unique = {}
    for j in judgments:
        unique.setdefault(j["url"], j)
    return list(unique.values())


def load_json(path, default=None):
    try:
        with open(path) as f:
            return json.load(f)
    except FileNotFoundError:
        return default


def write_json(path, data):
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
        f.write("\n")
    os.replace(tmp, path)


def file_stem(j: dict) -> str:
    """The Federal Court's file-store name, e.g. 2023FCA0665."""
    return f"{j['year']}{j['court']}{j['number']:04d}"


def existing_download(j: dict, out_dir: str) -> str | None:
    if j.get("file") and os.path.exists(j["file"]):
        return j["file"]
    for path in sorted(glob.glob(os.path.join(out_dir, str(j["year"]), file_stem(j) + ".*"))):
        if not path.endswith(".tmp"):
            return path
    return None


async def download_all(judgments: dict, out_dir: str, delay: float, limit: int | None,
                       save) -> tuple[int, int]:
    import fedcourt_browser as fb

    pending = {}  # url -> judgment entries sharing it (across file numbers)
    for entries in judgments.values():
        for j in entries:
            path = existing_download(j, out_dir)
            if path:
                j["file"] = path
                j.pop("download_error", None)
            elif j["url"].startswith(FEDCOURT_JUDGMENTS):
                pending.setdefault(j["url"], []).append(j)
            else:
                # AustLII links (courts other than FCA/FCAFC) aren't fetched.
                j["download_error"] = "not on judgments.fedcourt.gov.au"
    urls = list(pending)
    if limit is not None:
        urls = urls[:limit]
    log(f"{len(pending)} judgments to download; trying {len(urls)}")
    if not urls:
        return 0, 0

    ok = failed = 0
    async with fb.Browser() as browser:
        for n, url in enumerate(urls, 1):
            log(f"[{n}/{len(urls)}] {pending[url][0]['citation']}")
            error = None
            try:
                result = await fb.download_judgment(browser, url)
            except fb.NotPublished:
                result, error = None, "no document link on the judgment page (unpublished?)"
            if result is None:
                failed += 1
                for j in pending[url]:
                    j["download_error"] = error or "download failed"
                    j["file"] = None
            else:
                filename, data, _ = result
                dest = os.path.join(out_dir, str(pending[url][0]["year"]), filename)
                os.makedirs(os.path.dirname(dest), exist_ok=True)
                with open(dest + ".tmp", "wb") as f:
                    f.write(data)
                os.replace(dest + ".tmp", dest)
                log(f"  saved {dest} ({len(data)} bytes)")
                ok += 1
                for j in pending[url]:
                    j["file"] = dest
                    j.pop("download_error", None)
            if n % 10 == 0:
                save()  # keep progress if the run is cut short
            if n < len(urls):
                await asyncio.sleep(delay)
    return ok, failed


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--file-numbers", default="data/competition_law_file_numbers.json")
    ap.add_argument("--output", default="data/judgments.json")
    ap.add_argument("--out-dir", default="data/judgments", help="where documents are saved")
    ap.add_argument("--only", nargs="+", metavar="FILE_NUMBER",
                    help="only these file numbers (e.g. VID71/2023)")
    ap.add_argument("--new-only", action="store_true",
                    help="only scan file numbers not already in --output")
    ap.add_argument("--no-scan", action="store_true",
                    help="skip the portal scan; download what --output already lists")
    ap.add_argument("--no-download", action="store_true", help="only scan the portal")
    ap.add_argument("--limit", type=int, help="download at most this many judgments")
    ap.add_argument("--delay", type=float, default=1.5,
                    help="seconds between portal requests")
    ap.add_argument("--download-delay", type=float, default=3.0,
                    help="seconds between judgment downloads")
    args = ap.parse_args()

    existing = load_json(args.output, {}) or {}
    judgments: dict = existing.get("judgments", {})
    numbers = args.only or ([] if args.no_scan else load_json(args.file_numbers)["file_numbers"])

    def save():
        with_j = {k: v for k, v in judgments.items() if v}
        out = {
            "source": ORDER_LIST_URL.format(file_number="{file_number}"),
            "generated_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
            "files_checked": len(judgments),
            "files_with_judgments": len(with_j),
            "judgment_count": len({j["url"] for v in with_j.values() for j in v}),
            "downloaded_count": len({j["file"] for v in with_j.values() for j in v if j.get("file")}),
            "judgments": judgments,
        }
        # Keep the previous timestamp when nothing else has changed.
        prev = load_json(args.output, {}) or {}
        if prev.get("generated_at") and {**out, "generated_at": prev["generated_at"]} == prev:
            return
        write_json(args.output, out)

    scan_errors = []
    if not args.no_scan:
        todo = [n for n in numbers if not (args.new_only and n in judgments)]
        client = Client(delay=args.delay)
        log(f"Scanning {len(todo)} file numbers")
        for i, fn in enumerate(todo, 1):
            try:
                found = scan_file(client, fn)
            except Exception as e:  # noqa: BLE001
                log(f"  {fn}: ERROR {e}; keeping previous result")
                scan_errors.append(fn)
                continue
            # Carry download state over from the previous run.
            prev = {j["url"]: j for j in judgments.get(fn, [])}
            for j in found:
                for key in ("file", "download_error"):
                    if prev.get(j["url"], {}).get(key) is not None:
                        j[key] = prev[j["url"]][key]
            judgments[fn] = found
            if found:
                log(f"  [{i}/{len(todo)}] {fn}: {', '.join(j['citation'] for j in found)}")
            if i % 50 == 0:
                log(f"  [{i}/{len(todo)}] scanned")
                save()
        save()

    failed = 0
    if not args.no_download:
        ok, failed = asyncio.run(
            download_all(judgments, args.out_dir, args.download_delay, args.limit, save)
        )
        log(f"Downloaded {ok}, failed {failed}")
    save()
    log(f"Wrote {args.output}")
    if scan_errors:
        log(f"Scan errors ({len(scan_errors)}): {' '.join(scan_errors)}")
    # Fail the run only if a whole step looks broken, not for the odd
    # unpublished or suppressed judgment.
    if scan_errors and len(scan_errors) == len(numbers):
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
