// ==========================================================================
// Pantalla de inicio de sesión con la identidad de Echo Tecnología.
// Maneja tres estados:
//   "login"  → formulario normal de acceso
//   "forgot" → ingresar email para recibir el enlace
//   "reset"  → ingresar nueva contraseña (cuando hay ?reset=TOKEN en la URL)
// ==========================================================================
import { useState, type FormEvent } from "react";
import { useAuth } from "../auth/AuthContext";
import { authApi } from "../api/recursos";

type Vista = "login" | "forgot" | "reset";

export function Login({ resetToken }: { resetToken?: string }) {
  const { login } = useAuth();
  const [vista, setVista] = useState<Vista>(resetToken ? "reset" : "login");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [nuevaPassword, setNuevaPassword] = useState("");
  const [confirmar, setConfirmar] = useState("");
  const [error, setError] = useState<string | null>(null);
  const [exito, setExito] = useState<string | null>(null);
  const [enviando, setEnviando] = useState(false);

  async function onLogin(e: FormEvent) {
    e.preventDefault();
    setError(null);
    setEnviando(true);
    try {
      await login(email, password);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Error inesperado.");
    } finally {
      setEnviando(false);
    }
  }

  async function onForgot(e: FormEvent) {
    e.preventDefault();
    setError(null);
    setEnviando(true);
    try {
      await authApi.olvidarPassword(email);
      setExito("Si ese correo está registrado, recibirás el enlace en unos minutos.");
    } catch (err) {
      setError(err instanceof Error ? err.message : "Error al enviar el correo.");
    } finally {
      setEnviando(false);
    }
  }

  async function onReset(e: FormEvent) {
    e.preventDefault();
    setError(null);
    if (nuevaPassword !== confirmar) {
      setError("Las contraseñas no coinciden.");
      return;
    }
    setEnviando(true);
    try {
      await authApi.resetPassword(resetToken!, nuevaPassword);
      setExito("Contraseña actualizada. Ya puedes iniciar sesión.");
      // Limpiar el token de la URL sin recargar la página.
      window.history.replaceState({}, "", window.location.pathname);
      setTimeout(() => setVista("login"), 2000);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Error al restablecer la contraseña.");
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

      {/* Formulario */}
      <div className="login-form-zona">

        {/* ── Iniciar sesión ── */}
        {vista === "login" && (
          <form className="login-card" onSubmit={onLogin}>
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

            <button
              type="button"
              className="btn-secundario"
              style={{ marginTop: 8, fontSize: 13 }}
              onClick={() => { setVista("forgot"); setError(null); setExito(null); }}
            >
              ¿Olvidaste tu contraseña?
            </button>
          </form>
        )}

        {/* ── Olvidé mi contraseña ── */}
        {vista === "forgot" && (
          <form className="login-card" onSubmit={onForgot}>
            <h2>Recuperar contraseña</h2>
            <p className="sub">Ingresa tu correo y te enviaremos un enlace para restablecer tu contraseña.</p>

            {error && <div className="alerta-error">{error}</div>}
            {exito && <div className="alerta-exito">{exito}</div>}

            {!exito && (
              <div className="campo">
                <label htmlFor="email-rec">Correo electrónico</label>
                <input
                  id="email-rec"
                  type="email"
                  value={email}
                  onChange={(e) => setEmail(e.target.value)}
                  placeholder="usuario@empresa.com"
                  required
                  autoComplete="email"
                />
              </div>
            )}

            {!exito && (
              <button className="btn-primario" type="submit" disabled={enviando}>
                {enviando ? "Enviando…" : "Enviar enlace"}
              </button>
            )}

            <button
              type="button"
              className="btn-secundario"
              style={{ marginTop: 8, fontSize: 13 }}
              onClick={() => { setVista("login"); setError(null); setExito(null); }}
            >
              ← Volver al inicio de sesión
            </button>
          </form>
        )}

        {/* ── Nueva contraseña (desde el enlace del correo) ── */}
        {vista === "reset" && (
          <form className="login-card" onSubmit={onReset}>
            <h2>Nueva contraseña</h2>
            <p className="sub">Ingresa tu nueva contraseña.</p>

            {error && <div className="alerta-error">{error}</div>}
            {exito && <div className="alerta-exito">{exito}</div>}

            {!exito && (
              <>
                <div className="campo">
                  <label htmlFor="nueva">Nueva contraseña</label>
                  <input
                    id="nueva"
                    type="password"
                    value={nuevaPassword}
                    onChange={(e) => setNuevaPassword(e.target.value)}
                    placeholder="Mínimo 6 caracteres"
                    required
                    minLength={6}
                    autoComplete="new-password"
                  />
                </div>

                <div className="campo">
                  <label htmlFor="confirmar">Confirmar contraseña</label>
                  <input
                    id="confirmar"
                    type="password"
                    value={confirmar}
                    onChange={(e) => setConfirmar(e.target.value)}
                    placeholder="Repite la contraseña"
                    required
                    autoComplete="new-password"
                  />
                </div>

                <button className="btn-primario" type="submit" disabled={enviando}>
                  {enviando ? "Guardando…" : "Guardar contraseña"}
                </button>
              </>
            )}
          </form>
        )}

      </div>
    </div>
  );
}
