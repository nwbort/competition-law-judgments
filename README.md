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
