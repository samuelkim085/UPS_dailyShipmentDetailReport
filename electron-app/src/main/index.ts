import { app, BrowserWindow, ipcMain, dialog } from "electron";
import type { IpcMainInvokeEvent } from "electron";
import * as path from "path";
import * as fs from "fs";
import * as dotenv from "dotenv";
import { extractShipmentData } from "./pdf-parser";
import { buildXlsx } from "./xlsx-export";
import { exportToDb, testDbConnection } from "./db";

// Load .env from project root (parent of electron-app)
const envPath = path.join(__dirname, "..", "..", "..", ".env");
if (fs.existsSync(envPath)) {
  dotenv.config({ path: envPath });
}

let mainWindow: BrowserWindow | null = null;

function createWindow(): void {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    minWidth: 900,
    minHeight: 600,
    webPreferences: {
      preload: process.env.NODE_ENV === "development"
        ? path.join(__dirname, "..", "..", "out", "preload", "index.js")
        : path.join(__dirname, "..", "..", "out", "preload", "index.js"),
      contextIsolation: true,
      nodeIntegration: false,
    },
    title: "UPS Shipment Detail Report",
    backgroundColor: "#0a0a0a",
    autoHideMenuBar: true,
  });

  // dev: load Vite dev server, prod: load built file
  if (process.env.NODE_ENV === "development") {
    mainWindow!.loadURL("http://127.0.0.1:6200");
    // mainWindow!.webContents.openDevTools();
  } else {
    mainWindow!.loadFile(
      path.join(__dirname, "..", "..", "out", "renderer", "index.html")
    );
  }

  mainWindow!.on("closed", () => {
    mainWindow = null;
  });
}

app.whenReady().then(createWindow);

app.on("window-all-closed", () => {
  app.quit();
});

// --- IPC Handlers ---

ipcMain.handle("open-file-dialog", async () => {
  const result = await dialog.showOpenDialog({
    title: "Select UPS Shipment Detail Report PDF",
    filters: [{ name: "PDF Files", extensions: ["pdf"] }],
    properties: ["openFile"],
  });
  return result.canceled ? null : result.filePaths[0];
});

ipcMain.handle(
  "extract-pdf",
  async (_event: IpcMainInvokeEvent, filePath: string) => {
    return extractShipmentData(filePath);
  }
);

ipcMain.handle(
  "extract-pdf-buffer",
  async (_event: IpcMainInvokeEvent, buffer: ArrayBuffer) => {
    const tmpPath = path.join(
      app.getPath("temp"),
      `ups-extract-${Date.now()}.pdf`
    );
    fs.writeFileSync(tmpPath, Buffer.from(buffer));
    try {
      return await extractShipmentData(tmpPath);
    } finally {
      fs.unlinkSync(tmpPath);
    }
  }
);

ipcMain.handle(
  "save-csv",
  async (_event: IpcMainInvokeEvent, csvContent: string, defaultName: string) => {
    const result = await dialog.showSaveDialog({
      title: "Save CSV",
      defaultPath: defaultName,
      filters: [{ name: "CSV Files", extensions: ["csv"] }],
    });
    if (!result.canceled && result.filePath) {
      fs.writeFileSync(result.filePath, "\ufeff" + csvContent, "utf-8");
      return result.filePath;
    }
    return null;
  }
);

ipcMain.handle(
  "save-xlsx",
  async (_event: IpcMainInvokeEvent, records: any[], defaultName: string) => {
    const result = await dialog.showSaveDialog({
      title: "Save Excel",
      defaultPath: defaultName,
      filters: [
        { name: "Excel Files", extensions: ["xlsx"] },
      ],
    });
    if (!result.canceled && result.filePath) {
      const xlsxBuffer = await buildXlsx(records);
      fs.writeFileSync(result.filePath, xlsxBuffer);
      return result.filePath;
    }
    return null;
  }
);

ipcMain.handle(
  "export-db",
  async (_event: IpcMainInvokeEvent, records: any[], filename: string) => {
    return exportToDb(records, filename);
  }
);

ipcMain.handle("test-db-connection", async () => {
  return testDbConnection();
});
