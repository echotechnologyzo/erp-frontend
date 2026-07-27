// ==========================================================================
// Componente raíz. Decide qué mostrar según el estado de autenticación.
// Si la URL tiene ?reset=TOKEN muestra el formulario de nueva contraseña.
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

  // Si hay ?reset=TOKEN en la URL, mostrar el formulario de nueva contraseña
  // (aunque el usuario esté logueado, el enlace lo lleva aquí directamente).
  const params = new URLSearchParams(window.location.search);
  const resetToken = params.get("reset");
  if (resetToken) return <Login resetToken={resetToken} />;

  // Sin usuario -> login. Con usuario -> panel principal.
  return usuario ? <Dashboard /> : <Login />;
}
