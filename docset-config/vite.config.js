import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { viteSingleFile } from "vite-plugin-singlefile";

// One self-contained index.html, because a Dataverse web resource is a single
// file — external chunks would each need uploading and referencing separately.
export default defineConfig({
  plugins: [react(), viteSingleFile()],
  base: "./",
  build: {
    assetsInlineLimit: 100000000,
    chunkSizeWarningLimit: 100000000,
    cssCodeSplit: false,
    rollupOptions: { output: { inlineDynamicImports: true } },
  },
});
