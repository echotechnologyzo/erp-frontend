// ==========================================================================
// Cliente HTTP mínimo para hablar con la API del ERP.
// Adjunta automáticamente el token JWT guardado y centraliza el manejo de
// errores de respuesta.
// ==========================================================================
const API_URL = import.meta.env.VITE_API_URL ?? "http://localhost:4000/api";

const CLAVE_TOKEN = "echo_erp_token";

// Guarda / lee / borra el token en el almacenamiento local del navegador.
export const tokenStore = {
  get: () => localStorage.getItem(CLAVE_TOKEN),
  set: (token: string) => localStorage.setItem(CLAVE_TOKEN, token),
  clear: () => localStorage.removeItem(CLAVE_TOKEN),
};

// Realiza una petición a la API. Lanza un Error con el mensaje del backend
// si la respuesta no es exitosa.
export async function api<T>(
  ruta: string,
  opciones: RequestInit = {}
): Promise<T> {
  const token = tokenStore.get();

  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 30000); // 30 s máximo

  let resp: Response;
  try {
    resp = await fetch(`${API_URL}${ruta}`, {
      ...opciones,
      signal: controller.signal,
      headers: {
        "Content-Type": "application/json",
        ...(token ? { Authorization: `Bearer ${token}` } : {}),
        ...opciones.headers,
      },
    });
  } catch (e: unknown) {
    if (e instanceof Error && e.name === "AbortError") {
      throw new Error("La petición tardó demasiado. Verifica tu conexión e intenta de nuevo.");
    }
    throw e;
  } finally {
    clearTimeout(timer);
  }

  const datos = await resp.json().catch(() => ({}));

  if (!resp.ok) {
    // Si el backend mandó un array de errores detallados (validación Zod),
    // los incluimos en el mensaje para que lleguen a la pantalla.
    const errores: string[] = Array.isArray(datos.errores) ? datos.errores : [];
    const detalle =
      errores.length > 0 ? "\n" + errores.slice(0, 8).join("\n") : "";
    throw new Error((datos.mensaje ?? "Error al conectar con el servidor.") + detalle);
  }

  return datos as T;
}
