import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],
  server: {
    port: 5173,
  },
  build: {
    rollupOptions: {
      output: {
        // Separa las librerías grandes en chunks propios para que el navegador
        // las cachee independientemente del código de la app.
        manualChunks: {
          // React core — cambia muy poco, se cachea bien
          "vendor-react": ["react", "react-dom"],
          // SheetJS pesa ~1 MB; va en su propio chunk solo cuando se importa
          "vendor-xlsx": ["xlsx"],
          // html2pdf.js + jspdf + html2canvas — solo en la pantalla de remisiones
          "vendor-pdf": ["html2pdf.js"],
        },
      },
    },
    // Avisa si algún chunk supera 600 KB (valor por defecto 500)
    chunkSizeWarningLimit: 600,
  },
});
