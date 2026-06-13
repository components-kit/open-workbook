import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";

export default defineConfig({
  plugins: [react()],
  build: {
    emptyOutDir: false,
    outDir: "dist",
    rollupOptions: {
      input: "src/taskpane.tsx",
      output: {
        assetFileNames: "taskpane.[name][extname]",
        chunkFileNames: "taskpane.[name].js",
        entryFileNames: "taskpane.bundle.js"
      }
    }
  }
});
