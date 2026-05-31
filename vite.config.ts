import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// Configuración de Vite. El servidor de desarrollo corre en el puerto 5173,
// que es el origen permitido por CORS en el backend (.env -> CORS_ORIGIN).
export default defineConfig({
  plugins: [react()],
  server: {
    port: 5173,
  },
});
