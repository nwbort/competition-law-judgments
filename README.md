# competition-law-judgments
## Competition Law file numbers

`scripts/collect_file_numbers.py` searches the Commonwealth Courts Portal federal e-search
(`action_type=Competition_Law`) in 90-day windows, follows pagination,
clicks through the maintenance notice if shown, and writes the unique file numbers to
`data/competition_law_file_numbers.json`.

By default it is incremental: it re-searches from 30 days before the previous run's end date
to today and merges any new file numbers into the existing JSON. Use `--start YYYY-MM-DD` to
backfill an earlier range (results are still merged), or `--replace` to rebuild from scratch.

The `Collect Competition Law file numbers` GitHub Actions workflow runs it weekly and on demand
(Actions → Run workflow, with an optional backfill start date), committing the JSON when it changes.

## Judgments

`scripts/collect_judgments.py` finds the published judgments for each file number and downloads them.

1. **Find.** For each file number it fetches the file's public *List of orders* page on the
   Commonwealth Courts Portal, `https://www.comcourts.gov.au/file/FEDERAL/P/VID71/2023/order_list`
   (no login needed). It reuses the collector's `Client`, so it keeps cookies, retries, clicks
   through the maintenance notice and waits between requests (`--delay`, default 1.5s).
   An event with a published judgment, usually "Publication of Reasons", links to
   `javascript:openFCAJudgment('2023', 'FCA', '665')`. The script builds the URL the same way the
   portal's own JavaScript does:
   - FCA: `https://www.judgments.fedcourt.gov.au/judgments/Judgments/fca/single/2023/2023fca0665`
   - FCAFC: `https://www.judgments.fedcourt.gov.au/judgments/Judgments/fca/full/{year}/{year}fcafc{nnnn}`
   - any other court: `https://www.austlii.edu.au/au/cases/cth/{COURT}/{year}/{n}.html`. These are
     listed but not downloaded.
2. **Download.** judgments.fedcourt.gov.au, and the www.fedcourt.gov.au file store its Word links
   redirect to, both answer plain HTTP clients with a Cloudflare "Just a moment..." challenge.
   So the documents are fetched the way
   [nwbort/tribunal-tracker](https://github.com/nwbort/tribunal-tracker) fetches the Competition
   Tribunal's decisions. It drives a real Chrome through [nodriver](https://github.com/ultrafunkamsterdam/nodriver),
   headful under Xvfb, and waits for the challenge to clear (clicking the Turnstile box if one shows).
   It then follows the judgment page's "Original Word Document" link to the file store, and
   fetches the file with `fetch()` from inside the page. See `scripts/fedcourt_browser.py`.
   Documents are saved under the Federal Court's own file name, e.g.
   `data/judgments/2023/2023FCA0665.docx`. A judgment that is already on disk is never fetched
   again. One that appears in several files is downloaded once. Downloads are spaced
   `--download-delay` seconds apart (default 3s).

Output is `data/judgments.json`, which maps each checked file number to its judgments. An empty
list means none were found:

```json
"VID71/2023": [
  {
    "citation": "[2023] FCA 665",
    "court": "FCA",
    "year": 2023,
    "number": 665,
    "url": "https://www.judgments.fedcourt.gov.au/judgments/Judgments/fca/single/2023/2023fca0665",
    "event_date": "2023-06-19",
    "event": "Publication of Reasons",
    "file": "data/judgments/2023/2023FCA0665.docx"
  }
]
```

Only judgments the portal links to are found. Older files (2008, for example) list "Judgment"
events with no link, so they come back with an empty list even though judgments exist.

If a download fails, `file` is `null` and `download_error` says why. A common reason is a
judgment that the portal links to but that has no document yet ("unpublished?"). Failed
downloads are retried on the next run. If a file's order list can't be fetched, its previous
result is kept.

```bash
pip install -r requirements.txt
xvfb-run -a python scripts/collect_judgments.py                         # scan every file, download what's missing
xvfb-run -a python scripts/collect_judgments.py --only VID71/2023       # just these files
python scripts/collect_judgments.py --no-download                       # scan only (no browser needed)
xvfb-run -a python scripts/collect_judgments.py --no-scan --limit 20    # download from the existing JSON
```

The `Collect judgments` GitHub Actions workflow runs after each successful run of the file number
workflow, and can also be started by hand with an optional download limit. It installs Chrome and
Xvfb if they're missing, then commits `data/judgments.json` and `data/judgments/`.
