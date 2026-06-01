// ==========================================================================
// Pantalla de Usuarios (solo ADMIN): lista el personal con acceso al sistema
// y permite crear nuevos usuarios (operador o administrador).
// La contraseña exige mínimo 8 caracteres, con al menos una letra y un número.
// ==========================================================================
import { useEffect, useState, type FormEvent } from "react";
import { useAuth } from "../auth/AuthContext";
import {
  usuariosApi,
  catalogosApi,
  type Usuario,
  type NuevoUsuario,
  type Sede,
} from "../api/recursos";

export function Usuarios() {
  const { usuario: actual } = useAuth(); // usuario logueado (para no actuar sobre sí mismo)
  const [lista, setLista] = useState<Usuario[]>([]);
  const [sedes, setSedes] = useState<Sede[]>([]);
  const [cargando, setCargando] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [aviso, setAviso] = useState<string | null>(null);
  const [modal, setModal] = useState(false);
  const [passwordDe, setPasswordDe] = useState<Usuario | null>(null); // usuario al que cambiar contraseña
  const [ocupado, setOcupado] = useState<string | null>(null); // id en proceso (deshabilita botones)

  // Mapa id→nombre de sede para mostrar la sede de cada usuario.
  const nombreSede = (id: string | null) =>
    id ? sedes.find((s) => s.id === id)?.nombre ?? "—" : "—";

  async function cargar() {
    setCargando(true);
    setError(null);
    try {
      setLista(await usuariosApi.listar());
    } catch (e) {
      setError(e instanceof Error ? e.message : "Error al cargar usuarios.");
    } finally {
      setCargando(false);
    }
  }

  useEffect(() => {
    catalogosApi.sedes().then(setSedes);
    cargar();
  }, []);

  // Activar / desactivar un usuario.
  async function alternarEstado(u: Usuario) {
    setError(null);
    setAviso(null);
    setOcupado(u.id);
    try {
      await usuariosApi.cambiarEstado(u.id, !u.activo);
      setAviso(`Usuario ${u.nombre} ${u.activo ? "desactivado" : "activado"}.`);
      cargar();
    } catch (e) {
      setError(e instanceof Error ? e.message : "Error al cambiar el estado.");
    } finally {
      setOcupado(null);
    }
  }

  // Eliminar un usuario (con confirmación).
  async function eliminar(u: Usuario) {
    if (!window.confirm(`¿Eliminar al usuario "${u.nombre}"? Esta acción no se puede deshacer.`)) {
      return;
    }
    setError(null);
    setAviso(null);
    setOcupado(u.id);
    try {
      await usuariosApi.eliminar(u.id);
      setAviso(`Usuario ${u.nombre} eliminado.`);
      cargar();
    } catch (e) {
      setError(e instanceof Error ? e.message : "Error al eliminar el usuario.");
    } finally {
      setOcupado(null);
    }
  }

  return (
    <div>
      <div className="dash-topbar">
        <h2>Usuarios</h2>
        <button className="btn-primario" style={{ width: "auto" }} onClick={() => setModal(true)}>
          + Crear usuario
        </button>
      </div>

      {error && <div className="alerta-error">{error}</div>}
      {aviso && <div className="alerta-ok">{aviso}</div>}

      {cargando ? (
        <p>Cargando…</p>
      ) : (
        <div className="tabla-wrap">
          <table className="tabla">
            <thead>
              <tr>
                <th>Nombre</th>
                <th>Email</th>
                <th>Rol</th>
                <th>Sede</th>
                <th>Estado</th>
                <th>Creado</th>
                <th>Acciones</th>
              </tr>
            </thead>
            <tbody>
              {lista.map((u) => {
                const esYo = actual?.id === u.id; // no se permite actuar sobre uno mismo
                const enProceso = ocupado === u.id;
                return (
                  <tr key={u.id}>
                    <td><strong>{u.nombre}</strong>{esYo && <span className="muted"> (tú)</span>}</td>
                    <td>{u.email}</td>
                    <td><span className="chip-rol">{u.rol}</span></td>
                    <td>{nombreSede(u.sedeId)}</td>
                    <td>
                      <span className={u.activo ? "badge-nuevo" : "badge-sede"}>
                        {u.activo ? "Activo" : "Inactivo"}
                      </span>
                    </td>
                    <td>{new Date(u.creadoEn).toLocaleDateString("es-CO")}</td>
                    <td>
                      <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                        <button
                          className="btn-secundario"
                          style={{ padding: "6px 10px" }}
                          disabled={esYo || enProceso}
                          title={esYo ? "No puedes desactivar tu propio usuario" : ""}
                          onClick={() => alternarEstado(u)}
                        >
                          {u.activo ? "Desactivar" : "Activar"}
                        </button>
                        <button
                          className="btn-secundario"
                          style={{ padding: "6px 10px" }}
                          disabled={enProceso}
                          onClick={() => setPasswordDe(u)}
                        >
                          Contraseña
                        </button>
                        <button
                          className="btn-secundario"
                          style={{ padding: "6px 10px", color: "var(--echo-coral)", borderColor: "var(--echo-coral)" }}
                          disabled={esYo || enProceso}
                          title={esYo ? "No puedes eliminar tu propio usuario" : ""}
                          onClick={() => eliminar(u)}
                        >
                          Eliminar
                        </button>
                      </div>
                    </td>
                  </tr>
                );
              })}
              {lista.length === 0 && (
                <tr>
                  <td colSpan={7} className="muted" style={{ textAlign: "center", padding: 24 }}>
                    No hay usuarios. Crea el primero.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      )}

      {modal && (
        <ModalCrearUsuario
          sedes={sedes}
          onCerrar={() => setModal(false)}
          onCreado={() => { setModal(false); cargar(); }}
        />
      )}

      {passwordDe && (
        <ModalPassword
          usuario={passwordDe}
          onCerrar={() => setPasswordDe(null)}
          onGuardado={() => {
            setPasswordDe(null);
            setAviso("Contraseña actualizada.");
          }}
        />
      )}
    </div>
  );
}

function ModalCrearUsuario({
  sedes,
  onCerrar,
  onCreado,
}: {
  sedes: Sede[];
  onCerrar: () => void;
  onCreado: () => void;
}) {
  const [form, setForm] = useState<NuevoUsuario>({
    nombre: "",
    email: "",
    password: "",
    rol: "OPERADOR",
    sedeId: "",
  });
  const [error, setError] = useState<string | null>(null);
  const [guardando, setGuardando] = useState(false);

  function set<K extends keyof NuevoUsuario>(campo: K, valor: NuevoUsuario[K]) {
    setForm((f) => ({ ...f, [campo]: valor }));
  }

  async function onSubmit(e: FormEvent) {
    e.preventDefault();
    setError(null);
    setGuardando(true);
    try {
      // sedeId vacío → no se envía (el admin puede no tener sede).
      await usuariosApi.crear({ ...form, sedeId: form.sedeId || undefined });
      onCreado();
    } catch (err) {
      setError(err instanceof Error ? err.message : "Error al crear el usuario.");
    } finally {
      setGuardando(false);
    }
  }

  return (
    <div className="modal-fondo" onClick={onCerrar}>
      <div className="modal" style={{ maxWidth: 540 }} onClick={(e) => e.stopPropagation()}>
        <h2>+ Crear usuario</h2>
        <form onSubmit={onSubmit}>
          {error && <div className="alerta-error">{error}</div>}
          <div className="grid-2">
            <div className="campo">
              <label>Nombre *</label>
              <input value={form.nombre} onChange={(e) => set("nombre", e.target.value)} required />
            </div>
            <div className="campo">
              <label>Email *</label>
              <input type="email" value={form.email} onChange={(e) => set("email", e.target.value)} required />
            </div>
          </div>
          <div className="grid-2">
            <div className="campo">
              <label>Rol *</label>
              <select value={form.rol} onChange={(e) => set("rol", e.target.value as NuevoUsuario["rol"])}>
                <option value="OPERADOR">Operador</option>
                <option value="ADMIN">Administrador</option>
              </select>
            </div>
            <div className="campo">
              <label>Sede</label>
              <select value={form.sedeId ?? ""} onChange={(e) => set("sedeId", e.target.value)}>
                <option value="">Sin asignar</option>
                {sedes.map((s) => <option key={s.id} value={s.id}>{s.nombre}</option>)}
              </select>
            </div>
          </div>
          <div className="campo">
            <label>Contraseña *</label>
            <input
              type="password"
              value={form.password}
              onChange={(e) => set("password", e.target.value)}
              required
            />
            <small className="muted">Mínimo 8 caracteres, con al menos una letra y un número.</small>
          </div>
          <div className="modal-acciones">
            <button type="button" className="btn-secundario" onClick={onCerrar}>Cancelar</button>
            <button type="submit" className="btn-primario" style={{ width: "auto" }} disabled={guardando}>
              {guardando ? "Guardando…" : "Crear usuario"}
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}

// --------------------------------------------------------------------------
// Modal para que el ADMIN cambie la contraseña de un usuario.
// Pide la nueva contraseña dos veces y aplica la misma política de seguridad.
// --------------------------------------------------------------------------
function ModalPassword({
  usuario,
  onCerrar,
  onGuardado,
}: {
  usuario: Usuario;
  onCerrar: () => void;
  onGuardado: () => void;
}) {
  const [password, setPassword] = useState("");
  const [confirmar, setConfirmar] = useState("");
  const [error, setError] = useState<string | null>(null);
  const [guardando, setGuardando] = useState(false);

  async function onSubmit(e: FormEvent) {
    e.preventDefault();
    setError(null);
    if (password !== confirmar) {
      setError("Las contraseñas no coinciden.");
      return;
    }
    setGuardando(true);
    try {
      await usuariosApi.cambiarPassword(usuario.id, password);
      onGuardado();
    } catch (err) {
      setError(err instanceof Error ? err.message : "Error al cambiar la contraseña.");
    } finally {
      setGuardando(false);
    }
  }

  return (
    <div className="modal-fondo" onClick={onCerrar}>
      <div className="modal" style={{ maxWidth: 440 }} onClick={(e) => e.stopPropagation()}>
        <h2>Cambiar contraseña</h2>
        <p className="muted" style={{ marginTop: -8 }}>{usuario.nombre} · {usuario.email}</p>
        <form onSubmit={onSubmit}>
          {error && <div className="alerta-error">{error}</div>}
          <div className="campo">
            <label>Nueva contraseña *</label>
            <input type="password" value={password} onChange={(e) => setPassword(e.target.value)} required />
            <small className="muted">Mínimo 8 caracteres, con al menos una letra y un número.</small>
          </div>
          <div className="campo">
            <label>Confirmar contraseña *</label>
            <input type="password" value={confirmar} onChange={(e) => setConfirmar(e.target.value)} required />
          </div>
          <div className="modal-acciones">
            <button type="button" className="btn-secundario" onClick={onCerrar}>Cancelar</button>
            <button type="submit" className="btn-primario" style={{ width: "auto" }} disabled={guardando}>
              {guardando ? "Guardando…" : "Cambiar contraseña"}
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}
