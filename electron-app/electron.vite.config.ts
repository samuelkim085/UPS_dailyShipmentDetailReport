import { defineConfig, externalizeDepsPlugin } from "electron-vite";
import { resolve } from "path";

export default defineConfig({
  main: {
    plugins: [externalizeDepsPlugin()],
    build: {
      outDir: "out/main",
    },
  },
  preload: {
    plugins: [externalizeDepsPlugin()],
    build: {
      outDir: "out/preload",
    },
  },
  renderer: {
    root: resolve("src/renderer"),
    server: {
      port: 6200,
      host: "127.0.0.1",
      strictPort: false,
    },
    build: {
      outDir: "out/renderer",
      rollupOptions: {
        input: resolve("src/renderer/index.html"),
      },
    },
  },
});
