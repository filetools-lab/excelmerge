# Local Excel Report Merger

A lightweight static web app for merging multiple Excel reports into a single summary workbook. It uses pure HTML, CSS, JavaScript, and SheetJS in the browser—no backend or upload server required.

## Features

- Switch between English and Chinese from the top language selector
- Preview the workflow with a left/right before-and-after workbook demo, including source files, sheet names, row counts, and exported workbook sheets
- Upload multiple `.xlsx`, `.xls`, or `.csv` files
- Add files in multiple batches to a pending merge queue before running the merge
- Review header and content-type mismatches before merging, with the option to return and check files or continue anyway
- Merge all rows into one worksheet, or keep each source worksheet as a separate output sheet in one workbook
- Add `Source File` and `Source Sheet` columns for traceability
- Add a first-page `Cover Page` with file count, row count, source worksheet count, and output sheet links when exporting multiple workbook sheets
- Generate per-column summary statistics:
  - Non-empty values
  - Blank values
  - Unique values
  - Numeric sum, average, minimum, and maximum
- Preview the first 100 merged rows
- Export `summary.xlsx` with `Merged Data` and `Summary Stats` sheets

## Run locally

Open `index.html` directly in a modern browser, or serve the folder with any static file server:

```bash
python3 -m http.server 8000
```

Then visit <http://localhost:8000>.

> Note: Excel parsing and export are powered by SheetJS loaded in the browser from the official SheetJS CDN. File contents stay in the browser and are not sent to a backend.
