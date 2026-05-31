// ==========================================================================
// Contexto de autenticación.
// Mantiene el usuario logueado en memoria y expone login/logout a toda la app.
// ==========================================================================
import { createContext, useContext, useEffect, useState, type ReactNode } from "react";
import { api, tokenStore } from "../api/client";

export interface Usuario {
  id: string;
  nombre: string;
  email: string;
  rol: "ADMIN" | "OPERADOR";
  sedeId: string | null;
}

interface RespuestaLogin {
  token: string;
  usuario: Usuario;
}

interface AuthCtx {
  usuario: Usuario | null;
  cargando: boolean;
  login: (email: string, password: string) => Promise<void>;
  logout: () => void;
}

const Contexto = createContext<AuthCtx | null>(null);

export function AuthProvider({ children }: { children: ReactNode }) {
  const [usuario, setUsuario] = useState<Usuario | null>(null);
  const [cargando, setCargando] = useState(true);

  // Al cargar la app, si hay token guardado, recuperamos el perfil.
  useEffect(() => {
    const token = tokenStore.get();
    if (!token) {
      setCargando(false);
      return;
    }
    api<Usuario>("/auth/me")
      .then(setUsuario)
      .catch(() => tokenStore.clear())
      .finally(() => setCargando(false));
  }, []);

  async function login(email: string, password: string) {
    const data = await api<RespuestaLogin>("/auth/login", {
      method: "POST",
      body: JSON.stringify({ email, password }),
    });
    tokenStore.set(data.token);
    setUsuario(data.usuario);
  }

  function logout() {
    tokenStore.clear();
    setUsuario(null);
  }

  return (
    <Contexto.Provider value={{ usuario, cargando, login, logout }}>
      {children}
    </Contexto.Provider>
  );
}

// Hook para consumir el contexto de autenticación.
export function useAuth() {
  const ctx = useContext(Contexto);
  if (!ctx) throw new Error("useAuth debe usarse dentro de <AuthProvider>.");
  return ctx;
}
