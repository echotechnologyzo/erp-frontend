// ==========================================================================
// Pantalla de Empleados: directorio sencillo del personal por sede.
// ==========================================================================
import { useEffect, useState, type FormEvent } from "react";
import {
  empleadosApi,
  catalogosApi,
  type Empleado,
  type NuevoEmpleado,
  type Sede,
} from "../api/recursos";

export function Empleados() {
  const [lista, setLista] = useState<Empleado[]>([]);
  const [buscar, setBuscar] = useState("");
  const [cargando, setCargando] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [editar, setEditar] = useState<Empleado | "nuevo" | null>(null);

  async function cargar(q = buscar) {
    setCargando(true);
    setError(null);
    try {
      setLista(await empleadosApi.listar(q));
    } catch (e) {
      setError(e instanceof Error ? e.message : "Error al cargar empleados.");
    } finally {
      setCargando(false);
    }
  }

  useEffect(() => {
    cargar("");
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  return (
    <div>
      <div className="dash-topbar">
        <h2>Empleados</h2>
        <button className="btn-primario" style={{ width: "auto" }} onClick={() => setEditar("nuevo")}>
          + Crear empleado
        </button>
      </div>

      <div className="barra-busqueda">
        <input
          placeholder="Buscar por nombre, documento o cargo…"
          value={buscar}
          onChange={(e) => setBuscar(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && cargar(buscar)}
        />
        <button className="btn-secundario" onClick={() => cargar(buscar)}>Buscar</button>
      </div>

      {error && <div className="alerta-error">{error}</div>}

      {cargando ? (
        <p>Cargando…</p>
      ) : (
        <div className="tabla-wrap">
          <table className="tabla">
            <thead>
              <tr>
                <th>Documento</th>
                <th>Empleado</th>
                <th>Cargo</th>
                <th>Sede</th>
                <th>Teléfono</th>
                <th>Email</th>
                <th></th>
              </tr>
            </thead>
            <tbody>
              {lista.map((e) => (
                <tr key={e.id}>
                  <td>{e.documento}</td>
                  <td><strong>{e.nombre}</strong></td>
                  <td>{e.cargo ?? "—"}</td>
                  <td>{e.sede ? <span className="badge-sede">{e.sede.nombre}</span> : "—"}</td>
                  <td>{e.telefono ?? "—"}</td>
                  <td>{e.email ?? "—"}</td>
                  <td>
                    <button className="btn-secundario" style={{ padding: "6px 12px" }} onClick={() => setEditar(e)}>
                      Editar
                    </button>
                  </td>
                </tr>
              ))}
              {lista.length === 0 && (
                <tr>
                  <td colSpan={7} className="muted" style={{ textAlign: "center", padding: 24 }}>
                    No hay empleados. Crea el primero.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      )}

      {editar && (
        <ModalEmpleado
          empleado={editar === "nuevo" ? null : editar}
          onCerrar={() => setEditar(null)}
          onGuardado={() => {
            setEditar(null);
            cargar();
          }}
        />
      )}
    </div>
  );
}

function ModalEmpleado({
  empleado,
  onCerrar,
  onGuardado,
}: {
  empleado: Empleado | null;
  onCerrar: () => void;
  onGuardado: () => void;
}) {
  const [sedes, setSedes] = useState<Sede[]>([]);
  const [form, setForm] = useState<NuevoEmpleado>({
    documento: empleado?.documento ?? "",
    nombre: empleado?.nombre ?? "",
    cargo: empleado?.cargo ?? "",
    telefono: empleado?.telefono ?? "",
    email: empleado?.email ?? "",
    sedeId: empleado?.sedeId ?? "",
  });
  const [error, setError] = useState<string | null>(null);
  const [guardando, setGuardando] = useState(false);

  useEffect(() => {
    catalogosApi.sedes().then(setSedes);
  }, []);

  function set<K extends keyof NuevoEmpleado>(campo: K, valor: NuevoEmpleado[K]) {
    setForm((f) => ({ ...f, [campo]: valor }));
  }

  async function onSubmit(e: FormEvent) {
    e.preventDefault();
    setError(null);
    setGuardando(true);
    try {
      if (empleado) await empleadosApi.actualizar(empleado.id, form);
      else await empleadosApi.crear(form);
      onGuardado();
    } catch (err) {
      setError(err instanceof Error ? err.message : "Error al guardar el empleado.");
    } finally {
      setGuardando(false);
    }
  }

  // Eliminar el empleado en edición (con confirmación).
  async function eliminar() {
    if (!empleado) return;
    if (!window.confirm(`¿Eliminar al empleado "${empleado.nombre}"? Esta acción no se puede deshacer.`)) {
      return;
    }
    setError(null);
    setGuardando(true);
    try {
      await empleadosApi.eliminar(empleado.id);
      onGuardado();
    } catch (err) {
      setError(err instanceof Error ? err.message : "Error al eliminar el empleado.");
    } finally {
      setGuardando(false);
    }
  }

  return (
    <div className="modal-fondo">
      <div className="modal" style={{ maxWidth: 540 }} onClick={(e) => e.stopPropagation()}>
        <h2>{empleado ? "Editar empleado" : "+ Crear empleado"}</h2>
        <form onSubmit={onSubmit}>
          {error && <div className="alerta-error">{error}</div>}
          <div className="grid-2">
            <div className="campo">
              <label>Documento *</label>
              <input value={form.documento ?? ""} onChange={(e) => set("documento", e.target.value)} required />
            </div>
            <div className="campo">
              <label>Nombre *</label>
              <input value={form.nombre ?? ""} onChange={(e) => set("nombre", e.target.value)} required />
            </div>
          </div>
          <div className="grid-2">
            <div className="campo">
              <label>Cargo</label>
              <input value={form.cargo ?? ""} onChange={(e) => set("cargo", e.target.value)} />
            </div>
            <div className="campo">
              <label>Sede</label>
              <select value={form.sedeId ?? ""} onChange={(e) => set("sedeId", e.target.value)}>
                <option value="">Sin asignar</option>
                {sedes.map((s) => <option key={s.id} value={s.id}>{s.nombre}</option>)}
              </select>
            </div>
          </div>
          <div className="grid-2">
            <div className="campo">
              <label>Teléfono</label>
              <input value={form.telefono ?? ""} onChange={(e) => set("telefono", e.target.value)} />
            </div>
            <div className="campo">
              <label>Email</label>
              <input type="email" value={form.email ?? ""} onChange={(e) => set("email", e.target.value)} />
            </div>
          </div>
          <div className="modal-acciones" style={{ justifyContent: empleado ? "space-between" : "flex-end" }}>
            {empleado && (
              <button
                type="button"
                className="btn-secundario"
                style={{ color: "var(--echo-coral)", borderColor: "var(--echo-coral)" }}
                onClick={eliminar}
                disabled={guardando}
              >
                Eliminar
              </button>
            )}
            <div style={{ display: "flex", gap: 12 }}>
              <button type="button" className="btn-secundario" onClick={onCerrar}>Cancelar</button>
              <button type="submit" className="btn-primario" style={{ width: "auto" }} disabled={guardando}>
                {guardando ? "Guardando…" : empleado ? "Guardar cambios" : "Crear empleado"}
              </button>
            </div>
          </div>
        </form>
      </div>
    </div>
  );
}
