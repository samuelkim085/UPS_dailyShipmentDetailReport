"""
UPS Daily Shipment Detail Report - PDF Extractor
Extracts Package Ref No.1 and Tracking No. from UPS WorldShip PDF reports.

Usage:
    python ups_extract.py                          # interactive prompt / drag-and-drop
    python ups_extract.py <pdf_file>
    python ups_extract.py <pdf_file> -o output.csv
    python ups_extract.py <pdf_file> --format xlsx
"""

import argparse
import csv
import re
import sys
from pathlib import Path

try:
    import pdfplumber
except ImportError:
    print("Error: pdfplumber is required. Install it with:")
    print("  pip install pdfplumber")
    sys.exit(1)


def extract_shipment_data(pdf_path: str) -> list[dict]:
    """Extract Package Ref No.1 and Tracking No. pairs from a UPS PDF report."""
    records = []
    current_ref = None
    is_void = False

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue

            lines = text.split("\n")
            for line in lines:
                # Check for VOID status — can appear before or after tracking line
                if re.search(r"\bVOID\b", line) and "Voided" not in line:
                    # If VOID appears after tracking was already recorded, mark last record
                    if records and "Tracking" not in line:
                        records[-1]["Status"] = "VOID"
                    else:
                        is_void = True

                # Match shipment-level Package Ref No.1 (sets current ref)
                ref_match = re.search(r"Package Ref No\.1:\s*(.+?)(?:\s{2,}|$)", line)
                if ref_match:
                    current_ref = ref_match.group(1).strip()

                # Match Tracking No.
                tracking_match = re.search(r"Tracking No\.:\s*(1Z[A-Z0-9]+)", line)
                if tracking_match:
                    tracking = tracking_match.group(1).strip()
                    ref_value = current_ref or ""
                    status = "VOID" if is_void else "Active"
                    records.append({
                        "Package Ref No.1": ref_value,
                        "Tracking No.": tracking,
                        "Status": status,
                    })
                    is_void = False

    return records


def print_table(records: list[dict]) -> None:
    """Print records as a formatted table to the console."""
    if not records:
        print("No records found.")
        return

    # Column widths
    num_w = len(str(len(records)))
    ref_w = max(len(r["Package Ref No.1"]) for r in records)
    ref_w = max(ref_w, len("Package Ref No.1"))
    trk_w = max(len(r["Tracking No."]) for r in records)
    trk_w = max(trk_w, len("Tracking No."))
    sts_w = max(len(r["Status"]) for r in records)
    sts_w = max(sts_w, len("Status"))

    header = f"{'#':>{num_w}} | {'Package Ref No.1':<{ref_w}} | {'Tracking No.':<{trk_w}} | {'Status':<{sts_w}}"
    sep = "-" * len(header)

    print(sep)
    print(header)
    print(sep)
    for i, r in enumerate(records, 1):
        status_display = r["Status"] if r["Status"] == "VOID" else ""
        print(f"{i:>{num_w}} | {r['Package Ref No.1']:<{ref_w}} | {r['Tracking No.']:<{trk_w}} | {status_display:<{sts_w}}")
    print(sep)

    active = sum(1 for r in records if r["Status"] == "Active")
    voided = sum(1 for r in records if r["Status"] == "VOID")
    print(f"Total: {len(records)} packages ({active} active, {voided} voided)")


def save_csv(records: list[dict], output_path: str) -> None:
    """Save records to a CSV file."""
    with open(output_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["Package Ref No.1", "Tracking No.", "Status"])
        writer.writeheader()
        writer.writerows(records)
    print(f"Saved {len(records)} records to {output_path}")


def save_xlsx(records: list[dict], output_path: str) -> None:
    """Save records to an Excel file."""
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    except ImportError:
        print("Error: openpyxl is required for Excel output. Install it with:")
        print("  pip install openpyxl")
        sys.exit(1)

    wb = Workbook()
    ws = wb.active
    ws.title = "UPS Shipments"

    # Header style
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    headers = ["#", "Package Ref No.1", "Tracking No.", "Status"]
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
        cell.border = thin_border

    # Data rows
    void_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    for i, r in enumerate(records, 1):
        row = i + 1
        ws.cell(row=row, column=1, value=i).border = thin_border
        ws.cell(row=row, column=2, value=r["Package Ref No.1"]).border = thin_border
        ws.cell(row=row, column=3, value=r["Tracking No."]).border = thin_border
        status_cell = ws.cell(row=row, column=4, value=r["Status"])
        status_cell.border = thin_border
        if r["Status"] == "VOID":
            for col in range(1, 5):
                ws.cell(row=row, column=col).fill = void_fill

    # Auto-fit column widths
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 24
    ws.column_dimensions["D"].width = 10

    wb.save(output_path)
    print(f"Saved {len(records)} records to {output_path}")


def clean_path(raw: str) -> str:
    """Clean a file path from user input or drag-and-drop.

    Handles:
      - Surrounding quotes (double or single)
      - Trailing whitespace / newlines
      - Drag-and-drop paths that may include quotes on Windows
    """
    path = raw.strip().strip('"').strip("'").strip()
    return path


def prompt_for_pdf() -> Path:
    """Interactively ask the user for a PDF file path."""
    print("=" * 60)
    print("  UPS Daily Shipment Detail Report - PDF Extractor")
    print("=" * 60)
    print()
    print("  Drag and drop a PDF file here, or type the file path:")
    print()

    while True:
        try:
            raw = input("  PDF file: ")
        except (EOFError, KeyboardInterrupt):
            print("\nCancelled.")
            sys.exit(0)

        pdf_path = Path(clean_path(raw))

        if not pdf_path.as_posix():
            print("  No input provided. Please try again.\n")
            continue

        if not pdf_path.exists():
            print(f"  File not found: {pdf_path}")
            print("  Please try again.\n")
            continue

        if pdf_path.suffix.lower() != ".pdf":
            print(f"  Not a PDF file: {pdf_path.name}")
            print("  Please provide a .pdf file.\n")
            continue

        return pdf_path


def prompt_for_output(pdf_path: Path) -> tuple[Path | None, str | None]:
    """Ask the user if they want to save output to a file."""
    print()
    print("  Save output to file?")
    print("    1) CSV")
    print("    2) Excel (.xlsx)")
    print("    3) No, just show the table")
    print()

    while True:
        try:
            choice = input("  Choice [1/2/3]: ").strip()
        except (EOFError, KeyboardInterrupt):
            print()
            return None, None

        if choice == "1":
            default = pdf_path.with_suffix(".csv")
            return default, "csv"
        elif choice == "2":
            default = pdf_path.with_suffix(".xlsx")
            return default, "xlsx"
        elif choice == "3" or choice == "":
            return None, None
        else:
            print("  Invalid choice. Enter 1, 2, or 3.\n")


def main():
    parser = argparse.ArgumentParser(
        description="Extract Package Ref No.1 and Tracking No. from UPS Daily Shipment Detail Report PDFs."
    )
    parser.add_argument("pdf", nargs="?", default=None, help="Path to the UPS PDF report (or drag-and-drop)")
    parser.add_argument("-o", "--output", help="Output file path (CSV or XLSX)")
    parser.add_argument(
        "--format",
        choices=["csv", "xlsx"],
        default=None,
        help="Output format (auto-detected from -o extension, defaults to csv)",
    )

    args = parser.parse_args()

    # Interactive mode: no PDF argument provided
    if args.pdf is None:
        pdf_path = prompt_for_pdf()
        print(f"\n  Reading: {pdf_path}\n")
        records = extract_shipment_data(str(pdf_path))
        print_table(records)

        # Ask about saving
        output_path, fmt = prompt_for_output(pdf_path)
        if output_path and fmt:
            if fmt == "xlsx":
                save_xlsx(records, str(output_path))
            else:
                save_csv(records, str(output_path))
        return

    # CLI mode: PDF path provided as argument
    pdf_path = Path(clean_path(args.pdf))
    if not pdf_path.exists():
        print(f"Error: File not found: {pdf_path}")
        sys.exit(1)

    print(f"Reading: {pdf_path}")
    records = extract_shipment_data(str(pdf_path))
    print_table(records)

    # Save to file if requested
    if args.output:
        output_path = Path(args.output)
        fmt = args.format
        if not fmt:
            fmt = output_path.suffix.lstrip(".").lower()
            if fmt not in ("csv", "xlsx"):
                fmt = "csv"

        if fmt == "xlsx":
            save_xlsx(records, str(output_path))
        else:
            save_csv(records, str(output_path))
    elif args.format:
        stem = pdf_path.stem
        output_path = pdf_path.parent / f"{stem}.{args.format}"
        if args.format == "xlsx":
            save_xlsx(records, str(output_path))
        else:
            save_csv(records, str(output_path))


if __name__ == "__main__":
    main()
