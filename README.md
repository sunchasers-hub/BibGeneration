# Race Bib Generator

Generates one PowerPoint bib slide per runner from an Excel entry list, using
your existing bib design as the template.

## Requirements
```
pip install python-pptx pandas openpyxl
```

## Files
- `generate_bibs.py` — the generator script
- `template.pptx` — your bib design (copy of the sample you sent). The first
  slide is used as the master; edit the text boxes here if you want to change
  fonts, colors, or layout — every generated bib will follow suit.
- `sample_runners.xlsx` — example input file showing the expected columns

## Excel format
| Date | First Name | Last Name | Distance |
|---|---|---|---|
| 2026-06-28 | Aditi | Sharma | 5K |
| 2026-06-28 | Rohit | Verma | 5K |
| 2026-06-28 | Karan | Mehta | 10K |

Column names are matched case-insensitively, so "date", "Date", "DATE" all work.

## Run it
```
python generate_bibs.py runners.xlsx template.pptx output_bibs.pptx
```

This produces `output_bibs.pptx` with one slide per runner, in the same order
as the Excel rows.

## How bib numbering works
Each distance category gets its own number series, starting at
`(leading digits of the distance) x 100`:
- `5K`  -> starts at 500  (500, 501, 502, ...)
- `10K` -> starts at 1000 (1000, 1001, ...)
- `21K` -> starts at 2100
- `42K` -> starts at 4200

For distance labels without a leading number (e.g. "Fun Run", "Half
Marathon"), add an explicit entry to `CATEGORY_START_OVERRIDES` near the top
of `generate_bibs.py`:
```python
CATEGORY_START_OVERRIDES = {
    "Fun Run": 100,
    "Half Marathon": 2100,
    "Marathon": 4200,
}
```

**Heads up on capacity:** with the x100 rule, each category has room for 100
runners before its numbers would reach the next category's start (e.g. 5K
maxes out at 599 before bumping into 10K's 600s... actually it's safe up to
999). If you expect more than ~100 runners in the smallest category, widen
the gaps — e.g. multiply by 1000 instead of 100, or set explicit starts in
`CATEGORY_START_OVERRIDES`.

## How the "Month" replacement works
The script looks for the literal word "Month" (as a whole word, so it won't
mangle "Monthly") anywhere in the title text box and replaces it with the
full month name taken from each runner's Date, e.g. "Monthly Run Month 2026"
-> "Monthly Run June 2026".

## Customizing the template
Open `template.pptx` in PowerPoint and edit slide 1 directly:
- Change the logo/image, colors, fonts, position — freely
- Keep the literal text `FirstName LastName` somewhere as the name placeholder
- Keep the literal text `500` or `1000` somewhere as the bib-number placeholder
- Keep the literal word `Month` in the title where you want the month to appear
The script finds these by exact/partial text match, so as long as those
placeholder tokens stay in the text boxes, formatting changes are picked up
automatically.
