import { Pool } from "pg";
import type { ShipmentRecord, DbExportResult } from "./types";

function getPool(): Pool {
  return new Pool({
    host: process.env.DB_HOST || "localhost",
    port: parseInt(process.env.DB_PORT || "5432"),
    database: process.env.DB_NAME || "divalink",
    user: process.env.DB_USER || "postgres",
    password: process.env.DB_PASSWORD || "",
    connectionTimeoutMillis: 5000,
  });
}

/**
 * Test database connectivity. Returns true if connection succeeds.
 */
export async function testDbConnection(): Promise<{
  ok: boolean;
  error?: string;
}> {
  const pool = getPool();
  try {
    const client = await pool.connect();
    client.release();
    return { ok: true };
  } catch (err: any) {
    return { ok: false, error: err.message || "Connection failed" };
  } finally {
    await pool.end();
  }
}

/**
 * Export records to PostgreSQL shipments.ups_daily_detail table.
 * Uses ON CONFLICT (tracking_no) DO NOTHING for deduplication.
 */
export async function exportToDb(
  records: ShipmentRecord[],
  filename: string
): Promise<DbExportResult> {
  const pool = getPool();
  const client = await pool.connect();

  try {
    await client.query("BEGIN");
    let inserted = 0;
    let skipped = 0;

    for (const r of records) {
      const tracking = r["Tracking No."].trim();
      const packageRef = r["Package Ref No.1"].trim();
      const status = r.Status.trim();

      if (!tracking) {
        skipped++;
        continue;
      }

      const res = await client.query(
        `INSERT INTO shipments.ups_daily_detail
            (package_ref, tracking_no, status, pdf_filename)
         VALUES ($1, $2, $3, $4)
         ON CONFLICT (tracking_no) DO NOTHING`,
        [packageRef, tracking, status, filename]
      );

      if (res.rowCount && res.rowCount > 0) {
        inserted++;
      } else {
        skipped++;
      }
    }

    await client.query("COMMIT");
    return {
      inserted,
      skipped,
      total: records.length,
      message: `${inserted} inserted, ${skipped} skipped (duplicate)`,
    };
  } catch (err: any) {
    await client.query("ROLLBACK");
    throw new Error(
      err.message || "Database write failed. Transaction rolled back."
    );
  } finally {
    client.release();
    await pool.end();
  }
}
