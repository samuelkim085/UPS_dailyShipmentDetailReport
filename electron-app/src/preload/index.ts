import { contextBridge, ipcRenderer, webUtils } from "electron";

contextBridge.exposeInMainWorld("electronAPI", {
  openFileDialog: (): Promise<string | null> =>
    ipcRenderer.invoke("open-file-dialog"),

  extractPdf: (filePath: string): Promise<any[]> =>
    ipcRenderer.invoke("extract-pdf", filePath),

  extractPdfBuffer: (buffer: ArrayBuffer): Promise<any[]> =>
    ipcRenderer.invoke("extract-pdf-buffer", buffer),

  saveCsv: (csvContent: string, defaultName: string): Promise<string | null> =>
    ipcRenderer.invoke("save-csv", csvContent, defaultName),

  saveXlsx: (
    records: any[],
    defaultName: string
  ): Promise<string | null> =>
    ipcRenderer.invoke("save-xlsx", records, defaultName),

  exportDb: (records: any[], filename: string): Promise<any> =>
    ipcRenderer.invoke("export-db", records, filename),

  testDbConnection: (): Promise<{ ok: boolean; error?: string }> =>
    ipcRenderer.invoke("test-db-connection"),

  // Get native file path from a drag-and-dropped File object
  getPathForFile: (file: File): string =>
    webUtils.getPathForFile(file),
});
