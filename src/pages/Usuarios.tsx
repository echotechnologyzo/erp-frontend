// ==========================================================================
// Pantalla de Usuarios (solo ADMIN): lista el personal con acceso al sistema
// y permite crear nuevos usuarios (operador o administrador).
// La contraseña exige mínimo 8 caracteres, con al menos una letra y un número.
// ==========================================================================
import { useEffect, useState, type FormEvent } from "react";
import {
  usuariosApi,
  catalogosApi,
  type Usuario,
  type NuevoUsuario,
  type Sede,
} from "../api/recursos";

export function Usuarios() {
  const [lista, setLista] = useState<Usuario[]>([]);
  const [sedes, setSedes] = useState<Sede[]>([]);
  const [cargando, setCargando] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [modal, setModal] = useState(false);

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

  return (
    <div>
      <div className="dash-topbar">
        <h2>Usuarios</h2>
        <button className="btn-primario" style={{ width: "auto" }} onClick={() => setModal(true)}>
          + Crear usuario
        </button>
      </div>

      {error && <div className="alerta-error">{error}</div>}

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
              </tr>
            </thead>
            <tbody>
              {lista.map((u) => (
                <tr key={u.id}>
                  <td><strong>{u.nombre}</strong></td>
                  <td>{u.email}</td>
                  <td><span className="chip-rol">{u.rol}</span></td>
                  <td>{nombreSede(u.sedeId)}</td>
                  <td>
                    <span className={u.activo ? "badge-nuevo" : "badge-sede"}>
                      {u.activo ? "Activo" : "Inactivo"}
                    </span>
                  </td>
                  <td>{new Date(u.creadoEn).toLocaleDateString("es-CO")}</td>
                </tr>
              ))}
              {lista.length === 0 && (
                <tr>
                  <td colSpan={6} className="muted" style={{ textAlign: "center", padding: 24 }}>
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
