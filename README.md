# UPS Daily Shipment Detail Report — PDF Extractor

A Python CLI tool that extracts **Package Ref No.1** and **Tracking No.** from UPS WorldShip "Daily Shipment Detail Report" PDFs. Automatically detects voided shipments and exports to CSV or Excel.

## Installation

```bash
pip install pdfplumber openpyxl
```

- **pdfplumber** — required for PDF parsing
- **openpyxl** — optional, only needed for Excel (.xlsx) output

## Usage

### Interactive Mode

```bash
python ups_extract.py
```

You'll be prompted to drag-and-drop or type the path to a PDF file, then choose an output format.

### Command Line

```bash
python ups_extract.py report.pdf                   # print table to console
python ups_extract.py report.pdf -o results.csv    # save as CSV
python ups_extract.py report.pdf -o results.xlsx   # save as Excel
python ups_extract.py report.pdf --format xlsx     # auto-name output file
```

## Output

Each record contains:

| Field | Description |
|---|---|
| Package Ref No.1 | Customer reference number from the shipment |
| Tracking No. | UPS tracking number (1Z...) |
| Status | `Active` or `VOID` |

### Console Output Example

```
------------------------------------------------------------
# | Package Ref No.1     | Tracking No.        | Status
------------------------------------------------------------
1 | INV-2024-001         | 1ZABC12345678901234 |
2 | INV-2024-002         | 1ZDEF98765432109876 | VOID
------------------------------------------------------------
Total: 2 packages (1 active, 1 voided)
```

### Excel Output

Excel files include formatted headers and red-highlighted rows for voided shipments.

## License

MIT
