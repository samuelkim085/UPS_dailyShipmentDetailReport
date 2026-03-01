import type { ShipmentRecord } from "./types";

/**
 * Build CSV string from shipment records.
 */
export function buildCsv(records: ShipmentRecord[]): string {
  const header = '"Package Ref No.1","Tracking No.","Status"';
  const rows = records.map((r) => {
    const ref = r["Package Ref No.1"].replace(/"/g, '""');
    const tracking = r["Tracking No."].replace(/"/g, '""');
    const status = r.Status.replace(/"/g, '""');
    return `"${ref}","${tracking}","${status}"`;
  });
  return [header, ...rows].join("\n");
}
