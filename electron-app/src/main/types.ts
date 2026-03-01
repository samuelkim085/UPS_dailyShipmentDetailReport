export interface ShipmentRecord {
  "Package Ref No.1": string;
  "Tracking No.": string;
  Status: "Active" | "VOID";
}

export interface DbExportResult {
  inserted: number;
  skipped: number;
  total: number;
  message: string;
}
