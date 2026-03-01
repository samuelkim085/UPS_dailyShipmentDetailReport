import { execFile } from "child_process";
import * as path from "path";
import * as fs from "fs";
import type { ShipmentRecord } from "./types";

// Path to ups_extract.py — one level above electron-app/
const SCRIPT_PATH = path.resolve(__dirname, "../../../ups_extract.py");

// Fallback: try common Python executables in order
const PYTHON_CANDIDATES = ["python", "python3", "py"];

/**
 * Resolve the Python executable available on this system.
 * Returns the first candidate that exits successfully.
 */
function findPython(): Promise<string> {
  return new Promise((resolve, reject) => {
    const candidates = [...PYTHON_CANDIDATES];

    function tryNext() {
      const candidate = candidates.shift();
      if (!candidate) {
        reject(new Error("Python not found. Install Python 3 and ensure it is in PATH."));
        return;
      }
      execFile(candidate, ["--version"], { timeout: 5000 }, (err) => {
        if (err) {
          tryNext();
        } else {
          resolve(candidate);
        }
      });
    }

    tryNext();
  });
}

/**
 * Extract Package Ref No.1 and Tracking No. pairs from a UPS PDF report.
 * Delegates to ups_extract.py via child_process for accurate parsing.
 */
export async function extractShipmentData(
  pdfPath: string
): Promise<ShipmentRecord[]> {
  // Verify script exists
  if (!fs.existsSync(SCRIPT_PATH)) {
    throw new Error(`[ERROR] ups_extract.py not found at: ${SCRIPT_PATH}`);
  }

  // Verify PDF exists
  if (!fs.existsSync(pdfPath)) {
    throw new Error(`[ERROR] PDF file not found: ${pdfPath}`);
  }

  const python = await findPython();

  return new Promise((resolve, reject) => {
    execFile(
      python,
      [SCRIPT_PATH, pdfPath, "--json"],
      {
        timeout: 60000,        // 60s max
        maxBuffer: 10 * 1024 * 1024, // 10MB stdout buffer
        encoding: "utf8",
      },
      (err, stdout, stderr) => {
        if (err) {
          const detail = stderr ? `\n${stderr.trim()}` : "";
          reject(new Error(`[ERROR] Python extraction failed: ${err.message}${detail}`));
          return;
        }

        const output = stdout.trim();
        if (!output) {
          reject(new Error("[ERROR] Python script returned empty output."));
          return;
        }

        try {
          const records: ShipmentRecord[] = JSON.parse(output);
          resolve(records);
        } catch (parseErr) {
          reject(new Error(`[ERROR] Failed to parse JSON from Python output: ${output.substring(0, 200)}`));
        }
      }
    );
  });
}
