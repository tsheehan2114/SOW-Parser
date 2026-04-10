import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],
  server: {
    proxy: {
      // In local dev, proxy /api/* to a local Express or use Vercel CLI
      // Run `vercel dev` instead of `npm run dev` for full local testing
    },
  },
});
