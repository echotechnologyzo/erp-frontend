// ==========================================================================
// Componente raíz. Decide qué mostrar según el estado de autenticación.
// ==========================================================================
import { useAuth } from "./auth/AuthContext";
import { Login } from "./pages/Login";
import { Dashboard } from "./pages/Dashboard";

export function App() {
  const { usuario, cargando } = useAuth();

  // Mientras verificamos si hay sesión guardada, evitamos parpadeos.
  if (cargando) {
    return (
      <div style={{ display: "grid", placeItems: "center", minHeight: "100vh" }}>
        <p>Cargando…</p>
      </div>
    );
  }

  // Sin usuario -> login. Con usuario -> panel principal.
  return usuario ? <Dashboard /> : <Login />;
}
