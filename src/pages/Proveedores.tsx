// ==========================================================================
// Pantalla de Proveedores: directorio sencillo (listar, buscar, crear, editar).
// ==========================================================================
import { useEffect, useState, type FormEvent } from "react";
import { proveedoresApi, type Proveedor, type NuevoProveedor } from "../api/recursos";

export function Proveedores() {
  const [lista, setLista] = useState<Proveedor[]>([]);
  const [buscar, setBuscar] = useState("");
  const [cargando, setCargando] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [editar, setEditar] = useState<Proveedor | "nuevo" | null>(null);

  async function cargar(q = buscar) {
    setCargando(true);
    setError(null);
    try {
      setLista(await proveedoresApi.listar(q));
    } catch (e) {
      setError(e instanceof Error ? e.message : "Error al cargar proveedores.");
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
        <h2>Proveedores</h2>
        <button className="btn-primario" style={{ width: "auto" }} onClick={() => setEditar("nuevo")}>
          + Crear proveedor
        </button>
      </div>

      <div className="barra-busqueda">
        <input
          placeholder="Buscar por nombre o NIT…"
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
                <th>NIT / Documento</th>
                <th>Proveedor</th>
                <th>Teléfono</th>
                <th>Email</th>
                <th>Dirección</th>
                <th></th>
              </tr>
            </thead>
            <tbody>
              {lista.map((p) => (
                <tr key={p.id}>
                  <td>{p.documento}</td>
                  <td><strong>{p.nombre}</strong></td>
                  <td>{p.telefono ?? "—"}</td>
                  <td>{p.email ?? "—"}</td>
                  <td>{p.direccion ?? "—"}</td>
                  <td>
                    <button className="btn-secundario" style={{ padding: "6px 12px" }} onClick={() => setEditar(p)}>
                      Editar
                    </button>
                  </td>
                </tr>
              ))}
              {lista.length === 0 && (
                <tr>
                  <td colSpan={6} className="muted" style={{ textAlign: "center", padding: 24 }}>
                    No hay proveedores. Crea el primero.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      )}

      {editar && (
        <ModalProveedor
          proveedor={editar === "nuevo" ? null : editar}
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

function ModalProveedor({
  proveedor,
  onCerrar,
  onGuardado,
}: {
  proveedor: Proveedor | null;
  onCerrar: () => void;
  onGuardado: () => void;
}) {
  const [form, setForm] = useState<NuevoProveedor>({
    documento: proveedor?.documento ?? "",
    nombre: proveedor?.nombre ?? "",
    telefono: proveedor?.telefono ?? "",
    email: proveedor?.email ?? "",
    direccion: proveedor?.direccion ?? "",
  });
  const [error, setError] = useState<string | null>(null);
  const [guardando, setGuardando] = useState(false);

  function set<K extends keyof NuevoProveedor>(campo: K, valor: NuevoProveedor[K]) {
    setForm((f) => ({ ...f, [campo]: valor }));
  }

  async function onSubmit(e: FormEvent) {
    e.preventDefault();
    setError(null);
    setGuardando(true);
    try {
      if (proveedor) await proveedoresApi.actualizar(proveedor.id, form);
      else await proveedoresApi.crear(form);
      onGuardado();
    } catch (err) {
      setError(err instanceof Error ? err.message : "Error al guardar el proveedor.");
    } finally {
      setGuardando(false);
    }
  }

  return (
    <div className="modal-fondo" onClick={onCerrar}>
      <div className="modal" style={{ maxWidth: 540 }} onClick={(e) => e.stopPropagation()}>
        <h2>{proveedor ? "Editar proveedor" : "+ Crear proveedor"}</h2>
        <form onSubmit={onSubmit}>
          {error && <div className="alerta-error">{error}</div>}
          <div className="grid-2">
            <div className="campo">
              <label>NIT / Documento *</label>
              <input value={form.documento ?? ""} onChange={(e) => set("documento", e.target.value)} required />
            </div>
            <div className="campo">
              <label>Nombre *</label>
              <input value={form.nombre ?? ""} onChange={(e) => set("nombre", e.target.value)} required />
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
          <div className="campo">
            <label>Dirección</label>
            <input value={form.direccion ?? ""} onChange={(e) => set("direccion", e.target.value)} />
          </div>
          <div className="modal-acciones">
            <button type="button" className="btn-secundario" onClick={onCerrar}>Cancelar</button>
            <button type="submit" className="btn-primario" style={{ width: "auto" }} disabled={guardando}>
              {guardando ? "Guardando…" : proveedor ? "Guardar cambios" : "Crear proveedor"}
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}
