# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Single-script Python CLI tool that extracts **Package Ref No.1** and **Tracking No.** pairs from UPS WorldShip "Daily Shipment Detail Report" PDFs. It detects VOID shipments and can output results as a console table, CSV, or Excel (.xlsx).

## Running the Tool

```bash
python ups_extract.py                        # interactive mode (drag-and-drop prompt)
python ups_extract.py <pdf_file>             # CLI mode, prints table
python ups_extract.py <pdf_file> -o out.csv  # save to CSV
python ups_extract.py <pdf_file> --format xlsx  # save to Excel
```

## Dependencies

- **pdfplumber** (required) — PDF text extraction
- **openpyxl** (optional) — only needed for `.xlsx` output

Install: `pip install pdfplumber openpyxl`

## Architecture

Everything lives in `ups_extract.py`. Key functions:

- `extract_shipment_data(pdf_path)` — core parser: opens PDF with pdfplumber, walks lines with regex to pair "Package Ref No.1" values with "Tracking No." (1Z...) entries, tracks VOID status
- `print_table(records)` — formatted console output with active/voided summary
- `save_csv()` / `save_xlsx()` — file output (xlsx includes styled headers and red-highlighted VOID rows)
- `prompt_for_pdf()` / `prompt_for_output()` — interactive mode UI
- `clean_path()` — handles drag-and-drop quoted paths on Windows

## PDF Parsing Details

The parser relies on these patterns in UPS report text:
- `Package Ref No.1: <value>` — sets the current reference for subsequent tracking numbers
- `Tracking No.: 1Z<alphanumeric>` — captures tracking number, pairs with current ref
- `\bVOID\b` — marks shipments as voided (can appear before or after tracking line)

Records are returned as `list[dict]` with keys: `"Package Ref No.1"`, `"Tracking No."`, `"Status"`.
