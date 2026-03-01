import ExcelJS from "exceljs";
import type { ShipmentRecord } from "./types";

/**
 * Build an XLSX buffer with styled headers and VOID row highlighting.
 * Matches the openpyxl output from the original Python implementation.
 */
export async function buildXlsx(
  records: ShipmentRecord[]
): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("UPS Shipments");

  // Header styling — blue background (#4472C4), white bold text
  const headerFill: ExcelJS.FillPattern = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF4472C4" },
  };
  const headerFont: Partial<ExcelJS.Font> = {
    bold: true,
    color: { argb: "FFFFFFFF" },
  };
  const thinBorder: Partial<ExcelJS.Borders> = {
    top: { style: "thin" },
    left: { style: "thin" },
    bottom: { style: "thin" },
    right: { style: "thin" },
  };

  // Column widths
  ws.columns = [
    { header: "#", key: "num", width: 6 },
    { header: "Package Ref No.1", key: "ref", width: 40 },
    { header: "Tracking No.", key: "tracking", width: 24 },
    { header: "Status", key: "status", width: 10 },
  ];

  // Style header row
  const headerRow = ws.getRow(1);
  headerRow.eachCell((cell) => {
    cell.fill = headerFill;
    cell.font = headerFont;
    cell.alignment = { horizontal: "center" };
    cell.border = thinBorder;
  });

  // VOID row fill — light red (#FFC7CE)
  const voidFill: ExcelJS.FillPattern = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFFFC7CE" },
  };

  // Data rows
  for (let i = 0; i < records.length; i++) {
    const r = records[i];
    const row = ws.addRow({
      num: i + 1,
      ref: r["Package Ref No.1"],
      tracking: r["Tracking No."],
      status: r.Status,
    });

    row.eachCell((cell) => {
      cell.border = thinBorder;
      if (r.Status === "VOID") {
        cell.fill = voidFill;
      }
    });
  }

  const buffer = await wb.xlsx.writeBuffer();
  return Buffer.from(buffer);
}
