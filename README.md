# competition-law-judgments
## Competition Law file numbers

`scripts/collect_file_numbers.py` searches the Commonwealth Courts Portal federal e-search
(`action_type=Competition_Law`) in 90-day windows from 1 Jan 2000 (configurable with `--start`), follows pagination,
clicks through the maintenance notice if shown, and writes the unique file numbers to
`data/competition_law_file_numbers.json`.

The `Collect Competition Law file numbers` GitHub Actions workflow runs it weekly and on demand
(Actions → Run workflow), committing the JSON when it changes.
