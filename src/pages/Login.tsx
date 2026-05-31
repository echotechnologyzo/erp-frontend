// ==========================================================================
// Pantalla de inicio de sesión con la identidad de Echo Tecnología.
// ==========================================================================
import { useState, type FormEvent } from "react";
import { useAuth } from "../auth/AuthContext";

export function Login() {
  const { login } = useAuth();
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState<string | null>(null);
  const [enviando, setEnviando] = useState(false);

  async function onSubmit(e: FormEvent) {
    e.preventDefault();
    setError(null);
    setEnviando(true);
    try {
      await login(email, password);
      // Al loguear, AuthContext actualiza el usuario y App muestra el dashboard.
    } catch (err) {
      setError(err instanceof Error ? err.message : "Error inesperado.");
    } finally {
      setEnviando(false);
    }
  }

  return (
    <div className="login-pantalla">
      {/* Panel de marca con degradado azul→turquesa */}
      <div className="login-marca">
        <h1 className="logo">
          Echo<span> Tecnología</span>
        </h1>
        <p className="lema">
          Tecnología que protege y conecta tu hogar. Sistema de gestión
          interno: inventario, clientes y remisiones.
        </p>
      </div>

      {/* Formulario de acceso */}
      <div className="login-form-zona">
        <form className="login-card" onSubmit={onSubmit}>
          <h2>Iniciar sesión</h2>
          <p className="sub">Ingresa tus credenciales para continuar.</p>

          {error && <div className="alerta-error">{error}</div>}

          <div className="campo">
            <label htmlFor="email">Correo electrónico</label>
            <input
              id="email"
              type="email"
              value={email}
              onChange={(e) => setEmail(e.target.value)}
              placeholder="usuario@empresa.com"
              required
              autoComplete="email"
            />
          </div>

          <div className="campo">
            <label htmlFor="password">Contraseña</label>
            <input
              id="password"
              type="password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              placeholder="••••••••"
              required
              autoComplete="current-password"
            />
          </div>

          <button className="btn-primario" type="submit" disabled={enviando}>
            {enviando ? "Ingresando…" : "Ingresar"}
          </button>
        </form>
      </div>
    </div>
  );
}
