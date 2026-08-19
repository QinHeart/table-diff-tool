import react from "@vitejs/plugin-react";
import { resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { defineConfig } from "vite";

const projectRoot = fileURLToPath(new URL(".", import.meta.url));

export default defineConfig({
  root: resolve(projectRoot, "static-site"),
  publicDir: resolve(projectRoot, "public"),
  base: process.env.VITE_BASE_PATH || "/",
  plugins: [react()],
  build: {
    outDir: resolve(projectRoot, process.env.VITE_OUT_DIR || "dist"),
    emptyOutDir: true,
    rollupOptions: {
      input: {
        main: resolve(projectRoot, "static-site/index.html"),
        candidate: resolve(projectRoot, "static-site/candidate/index.html"),
      },
    },
  },
});
