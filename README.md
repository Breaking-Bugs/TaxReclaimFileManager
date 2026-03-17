# Tax Reclaim File Manager

This project processes Excel-based tax reclaim data together with PDF files.

It can:
- read one or more Excel files
- match rows to PDFs using a request ID
- create renamed PDF outputs
- sort outputs into BO folders
- merge all PDFs into one file
- merge PDFs per BO
- create a full run snapshot and manifest
- optionally clean up processed input files

## Requirements

- Python 3.9+
- `openpyxl`
- `PyPDF2`

Example install:

```bash
pip install openpyxl PyPDF2
```

## Project Structure

```text
TaxReclaimFileManager/
  input/
    excel/
    pdf_inbox/
  archive/
  settings.json
  sort_rename_merge.py
  run_sort_rename_merge.bat
```

## How It Works

1. Excel files are read from `input/excel`.
2. PDFs are read from `input/pdf_inbox`.
3. A timestamped run folder is created under `archive/`.
4. The input files are copied into a snapshot for auditability.
5. Excel rows are matched to PDFs using the configured `pdf_base` column.
6. Depending on `settings.json`, PDFs are copied, renamed, sorted, and merged.
7. A `manifest.json` and `log.txt` are written into the run folder.
8. A compact run summary is printed in the shell at the end.

The script always works from the snapshot after the initial input copy. This keeps archives deterministic and audit-ready.

## Run The Script

On Windows:

```bat
run_sort_rename_merge.bat
```

Or directly:

```bash
python sort_rename_merge.py
```

## Excel Requirements

The script validates that all required Excel headers exist before processing.

By default these columns are expected:
- `Financial Instrument`
- `Paymt. Date`
- `BO Name`
- `Clearstream Request ID`

You can change the column names in `settings.json`.

## Settings

The script is heavily controlled through `settings.json`.

### `columns`

Maps logical fields to Excel header names.

```json
"columns": {
  "isin": "Financial Instrument",
  "date": "Paymt. Date",
  "bo": "BO Name",
  "pdf_base": "Clearstream Request ID"
}
```

### `actions`

Controls what the script should produce.

```json
"actions": {
  "create_outputs": true,
  "sort_into_bo_folders": true,
  "rename_pdfs": true,
  "merge_all": true,
  "per_bo_merge": false,
  "cleanup_input": true
}
```

Meaning:
- `create_outputs`: create individual PDF files in the output area
- `sort_into_bo_folders`: create BO subfolders
- `rename_pdfs`: rename output PDFs using the configured pattern
- `merge_all`: create one merged PDF for all matched records
- `per_bo_merge`: create one merged PDF per BO
- `cleanup_input`: delete successfully processed input Excel/PDF files after the run

### `lookup`

Controls PDF matching behavior.

```json
"lookup": {
  "strategies": ["exact"],
  "case_insensitive": true
}
```

### `runtime`

Controls how aggressively the script retries file operations.

```json
"runtime": {
  "mode": "normal",
  "normal": {
    "retries": 1,
    "delay_ms": 0
  },
  "robust": {
    "retries": 5,
    "delay_ms": 250
  }
}
```

Use:
- `normal` for local or stable drives
- `robust` for slower, encrypted, or lock-prone drives

### `naming`

Controls filename and folder naming.

```json
"naming": {
  "date_format": "%Y-%m-%d",
  "output_filename_pattern": "{date}_{isin}.pdf",
  "bo_folder_pattern": "{bo}",
  "merged_all_filename": "merged_all.pdf",
  "merged_bo_filename_pattern": "merged_{bo}.pdf"
}
```

Available placeholders:
- `{isin}`
- `{bo}`
- `{date}`
- `{pdf_base}`

Examples:
- `"{isin}_{date}.pdf"`
- `"{date}_{bo}_{isin}.pdf"`
- `"BO_{bo}"`

### `paths`

Controls where input and output folders live.

```json
"paths": {
  "excel_input": "input/excel",
  "pdf_input": "input/pdf_inbox",
  "archive": "archive",
  "snapshot_excel": "input_snapshot/excel",
  "snapshot_pdf": "input_snapshot/pdf_inbox",
  "individual_output": "sorted_by_bo",
  "merged_output": "merged"
}
```

### `sanitize_filenames`

If `true`, invalid Windows filename characters are replaced automatically.

## Run Output

Each run creates a folder like:

```text
archive/2026-03-14_15-35-26/
```

Inside it you will typically find:
- `input_snapshot/excel/`
- `input_snapshot/pdf_inbox/`
- `sorted_by_bo/`
- `merged/`
- `manifest.json`
- `log.txt`
- `warnings.txt` if unprocessed inbox PDFs remain

## Manifest

The manifest contains:
- run ID
- final status
- start and end time
- duration in seconds
- input snapshots
- processed row counts
- missing PDFs
- defective PDFs
- unprocessed inbox PDFs
- created output files
- created merged files
- effective settings

## Shell Summary

At the end of each run the script prints a short summary to the shell, including:
- final status: `SUCCESS` or `SUCCESS WITH WARNINGS`
- Excel and PDF file counts
- processed, valid, and skipped record counts
- missing, defective, and unprocessed PDF counts
- created output and merge file counts
- run folder path
- total duration

## Notes

- If a required Excel column is missing, that Excel file is skipped with an error log entry.
- If a matching PDF is missing, the row is skipped and logged.
- If multiple PDFs match the same request ID, the row is skipped and logged.
- Duplicate output names are made unique automatically by appending `_1`, `_2`, etc.
- In robust mode, sensitive file operations are retried automatically.

## Typical Use Cases

### Rename + sort + merge all

```json
"actions": {
  "create_outputs": true,
  "sort_into_bo_folders": true,
  "rename_pdfs": true,
  "merge_all": true,
  "per_bo_merge": false,
  "cleanup_input": false
}
```

### Only merge, no individual outputs

```json
"actions": {
  "create_outputs": false,
  "sort_into_bo_folders": false,
  "rename_pdfs": false,
  "merge_all": true,
  "per_bo_merge": false,
  "cleanup_input": false
}
```

### Per-BO merge only

```json
"actions": {
  "create_outputs": true,
  "sort_into_bo_folders": true,
  "rename_pdfs": true,
  "merge_all": false,
  "per_bo_merge": true,
  "cleanup_input": false
}
```
