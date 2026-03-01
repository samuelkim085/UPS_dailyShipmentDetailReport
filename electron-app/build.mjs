import * as esbuild from "esbuild";

// Plugin: Replace `require("electron")` with `require("node:electron")`
// which bypasses node_modules resolution and hits Electron's built-in loader.
const electronPlugin = {
  name: "electron-external",
  setup(build) {
    build.onResolve({ filter: /^electron$/ }, () => ({
      path: "electron",
      external: true,
    }));
  },
};

const commonOptions = {
  bundle: true,
  platform: "node",
  target: "node20",
  format: "cjs",
  sourcemap: true,
  plugins: [electronPlugin],
};

// Main process
await esbuild.build({
  ...commonOptions,
  entryPoints: ["src/main/index.ts"],
  outfile: "dist/main/index.js",
});

// Preload script
await esbuild.build({
  ...commonOptions,
  entryPoints: ["src/preload/index.ts"],
  outfile: "dist/preload/index.js",
});

console.log("Build complete.");
