// ==========================================================================
// Pantalla de Remisiones de compra.
// Lista las remisiones, permite crear una nueva (que actualiza inventario y
// costos) e importar masivamente desde un archivo Excel.
// ==========================================================================
import { useEffect, useState, type ChangeEvent, type FormEvent } from "react";
import * as XLSX from "xlsx";
import {
  comprasApi,
  catalogosApi,
  articulosApi,
  proveedoresApi,
  type Compra,
  type Sede,
  type Articulo,
  type Proveedor,
  type NuevaCompraItem,
  type FilaImportCompra,
  type EditarCompra,
} from "../api/recursos";
import { SelectorArticulo } from "../components/SelectorArticulo";
import { useAuth } from "../auth/AuthContext";

const moneda = (v: number) =>
  new Intl.NumberFormat("es-CO", { style: "currency", currency: "COP", maximumFractionDigits: 0 }).format(v);

export function Compras() {
  const [compras, setCompras] = useState<Compra[]>([]);
  const [cargando, setCargando] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [aviso, setAviso] = useState<string | null>(null);
  const [modal, setModal] = useState(false);
  const [editar, setEditar] = useState<Compra | null>(null);

  async function cargar() {
    setCargando(true);
    setError(null);
    try {
      setCompras(await comprasApi.listar());
    } catch (e) {
      setError(e instanceof Error ? e.message : "Error al cargar remisiones.");
    } finally {
      setCargando(false);
    }
  }

  useEffect(() => {
    cargar();
  }, []);

  // --- Eliminar una compra (revierte el stock) ---
  async function eliminar(c: Compra) {
    if (!window.confirm(
      `¿Eliminar la remisión de compra ${c.documento}? Se devolverá (saldrá) del inventario el stock que ingresó.`
    )) return;
    setError(null);
    setAviso(null);
    try {
      await comprasApi.eliminar(c.id);
      setAviso(`Remisión de compra ${c.documento} eliminada (stock revertido).`);
      cargar();
    } catch (e) {
      setError(e instanceof Error ? e.message : "Error al eliminar la remisión de compra.");
    }
  }

  // --- Importación desde Excel ---
  async function onArchivoExcel(e: ChangeEvent<HTMLInputElement>) {
    const archivo = e.target.files?.[0];
    if (!archivo) return;
    setError(null);
    setAviso(null);
    try {
      const buffer = await archivo.arrayBuffer();
      const libro = XLSX.read(buffer);
      const hoja = libro.Sheets[libro.SheetNames[0]];
      // Cada fila del Excel se mapea a una FilaImportCompra (los encabezados
      // de la primera fila deben coincidir con estos nombres).
      const filas = XLSX.utils.sheet_to_json<FilaImportCompra>(hoja);
      if (filas.length === 0) throw new Error("El archivo no tiene filas.");
      const r = await comprasApi.importar(filas);
      setAviso(`Importación exitosa: ${r.remisionesCreadas} remisiones (${r.filasProcesadas} filas).`);
      cargar();
    } catch (err) {
      setError(err instanceof Error ? err.message : "Error al importar el Excel.");
    } finally {
      e.target.value = ""; // permite volver a elegir el mismo archivo
    }
  }

  // Descarga una plantilla de Excel con los encabezados esperados.
  function descargarPlantilla() {
    const ejemplo: FilaImportCompra[] = [
      {
        sede: "Bogotá",
        proveedorDocumento: "912000471",
        proveedorNombre: "AMAZON.COM SALES, INC",
        proveedorDireccion: "",
        remisionProveedor: "20/26/05/2026",
        articuloCodigo: "35",
        descripcion: "Fire TV Stick 4K",
        cantidad: 10,
        costoUnitario: 110000,
        descuento: 0,
      },
    ];
    const hoja = XLSX.utils.json_to_sheet(ejemplo);
    const libro = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(libro, hoja, "Compras");
    XLSX.writeFile(libro, "plantilla_remisiones_compra.xlsx");
  }

  return (
    <div>
      <div className="dash-topbar">
        <h2>Remisiones de compra</h2>
        <div style={{ display: "flex", gap: 10 }}>
          <button className="btn-secundario" onClick={descargarPlantilla}>
            Descargar plantilla Excel
          </button>
          <label className="btn-secundario" style={{ cursor: "pointer" }}>
            Importar Excel
            <input type="file" accept=".xlsx,.xls" hidden onChange={onArchivoExcel} />
          </label>
          <button className="btn-primario" style={{ width: "auto" }} onClick={() => setModal(true)}>
            + Crear remisión
          </button>
        </div>
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
                <th>Fecha</th>
                <th>ID</th>
                <th>Sede</th>
                <th>Proveedor</th>
                <th>Remisión prov.</th>
                <th>Total</th>
                <th>Estado</th>
                <th></th>
              </tr>
            </thead>
            <tbody>
              {compras.map((c) => (
                <tr key={c.id}>
                  <td>{new Date(c.fecha).toLocaleString("es-CO")}</td>
                  <td><strong>{c.documento}</strong></td>
                  <td>{c.sede}</td>
                  <td>
                    {c.proveedor}
                    <div className="muted">NIT: {c.proveedorDocumento}</div>
                  </td>
                  <td>{c.remisionProveedor ?? "—"}</td>
                  <td>{moneda(Number(c.total))}</td>
                  <td><span className="badge-sede">Pago total</span></td>
                  <td>
                    <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                      <button className="btn-secundario" style={{ padding: "6px 10px" }} onClick={() => setEditar(c)}>
                        Editar
                      </button>
                      <button
                        className="btn-secundario"
                        style={{ padding: "6px 10px", color: "var(--echo-coral)", borderColor: "var(--echo-coral)" }}
                        onClick={() => eliminar(c)}
                      >
                        Eliminar
                      </button>
                    </div>
                  </td>
                </tr>
              ))}
              {compras.length === 0 && (
                <tr>
                  <td colSpan={8} className="muted" style={{ textAlign: "center", padding: 24 }}>
                    No hay remisiones. Crea una o importa desde Excel.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      )}

      {modal && <ModalCrearCompra onCerrar={() => setModal(false)} onCreado={() => { setModal(false); cargar(); }} />}

      {editar && (
        <ModalEditarCompra
          compra={editar}
          onCerrar={() => setEditar(null)}
          onGuardado={() => { setEditar(null); cargar(); }}
        />
      )}
    </div>
  );
}

// --------------------------------------------------------------------------
// Modal de edición de la CABECERA de una compra (proveedor, remisión del
// proveedor, observación). Las líneas/cantidades NO se editan aquí porque ya
// movieron inventario; para cambiarlas: eliminar la compra y volver a crearla.
// --------------------------------------------------------------------------
function ModalEditarCompra({
  compra,
  onCerrar,
  onGuardado,
}: {
  compra: Compra;
  onCerrar: () => void;
  onGuardado: () => void;
}) {
  const [documento, setDocumento] = useState(compra.proveedorDocumento);
  const [nombre, setNombre] = useState(compra.proveedor);
  const [remisionProveedor, setRemisionProveedor] = useState(compra.remisionProveedor ?? "");
  const [observacion, setObservacion] = useState(compra.observacion ?? "");
  const [provBuscar, setProvBuscar] = useState("");
  const [provOpc, setProvOpc] = useState<Proveedor[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [guardando, setGuardando] = useState(false);

  useEffect(() => {
    const t = setTimeout(async () => {
      if (provBuscar.trim().length < 2) { setProvOpc([]); return; }
      try { setProvOpc(await proveedoresApi.listar(provBuscar)); } catch { setProvOpc([]); }
    }, 300);
    return () => clearTimeout(t);
  }, [provBuscar]);

  async function onSubmit(e: FormEvent) {
    e.preventDefault();
    setError(null);
    setGuardando(true);
    try {
      const datos: EditarCompra = {
        proveedor: { documento, nombre },
        remisionProveedor: remisionProveedor || undefined,
        observacion: observacion || undefined,
      };
      await comprasApi.editar(compra.id, datos);
      onGuardado();
    } catch (err) {
      setError(err instanceof Error ? err.message : "Error al guardar la remisión de compra.");
    } finally {
      setGuardando(false);
    }
  }

  return (
    <div className="modal-fondo">
      <div className="modal" style={{ maxWidth: 560 }} onClick={(e) => e.stopPropagation()}>
        <h2>Editar remisión de compra {compra.documento}</h2>
        <form onSubmit={onSubmit}>
          {error && <div className="alerta-error">{error}</div>}

          {/* Buscar proveedor existente por nombre o NIT */}
          <div className="campo" style={{ position: "relative" }}>
            <label>Buscar proveedor (nombre o NIT)</label>
            <input
              value={provBuscar}
              onChange={(e) => setProvBuscar(e.target.value)}
              placeholder="Escribe para reemplazar el proveedor…"
            />
            {provOpc.length > 0 && (
              <div className="lista-opciones" style={{ position: "absolute", zIndex: 10, width: "100%", maxHeight: 220, overflowY: "auto" }}>
                {provOpc.map((p) => (
                  <button
                    type="button"
                    key={p.id}
                    className="opcion"
                    onMouseDown={() => { setDocumento(p.documento); setNombre(p.nombre); setProvBuscar(""); setProvOpc([]); }}
                  >
                    <strong>{p.nombre}</strong> · NIT {p.documento}
                  </button>
                ))}
              </div>
            )}
          </div>

          <div className="grid-2">
            <div className="campo">
              <label>NIT / Documento proveedor *</label>
              <input value={documento} onChange={(e) => setDocumento(e.target.value)} required />
            </div>
            <div className="campo">
              <label>Nombre proveedor *</label>
              <input value={nombre} onChange={(e) => setNombre(e.target.value)} required />
            </div>
          </div>

          <div className="campo">
            <label>Remisión del proveedor</label>
            <input value={remisionProveedor} onChange={(e) => setRemisionProveedor(e.target.value)} />
          </div>
          <div className="campo">
            <label>Observación</label>
            <input value={observacion} onChange={(e) => setObservacion(e.target.value)} />
          </div>

          <p className="muted">
            Para cambiar artículos o cantidades, elimina esta remisión y vuelve a crearla
            (eso ajusta el inventario correctamente).
          </p>

          <div className="modal-acciones">
            <button type="button" className="btn-secundario" onClick={onCerrar}>Cancelar</button>
            <button type="submit" className="btn-primario" style={{ width: "auto" }} disabled={guardando}>
              {guardando ? "Guardando…" : "Guardar cambios"}
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}

// --------------------------------------------------------------------------
// Modal de creación de remisión de compra.
// --------------------------------------------------------------------------
function ModalCrearCompra({ onCerrar, onCreado }: { onCerrar: () => void; onCreado: () => void }) {
  const { usuario } = useAuth(); // para prellenar la sede asignada al usuario
  const [sedes, setSedes] = useState<Sede[]>([]);
  const [articulos, setArticulos] = useState<Articulo[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [guardando, setGuardando] = useState(false);

  const [sedeId, setSedeId] = useState("");
  const [documento, setDocumento] = useState("");
  const [nombre, setNombre] = useState("");
  const [remisionProveedor, setRemisionProveedor] = useState("");
  const [items, setItems] = useState<NuevaCompraItem[]>([
    { articuloId: "", cantidad: 1, costoUnitario: 0 },
  ]);

  // Buscador de proveedor (por nombre o NIT) para reutilizar uno existente.
  const [provBuscar, setProvBuscar] = useState("");
  const [provOpc, setProvOpc] = useState<Proveedor[]>([]);

  useEffect(() => {
    catalogosApi.sedes().then((s) => {
      setSedes(s);
      const propia = usuario?.sedeId && s.some((x) => x.id === usuario.sedeId)
        ? usuario.sedeId
        : s[0]?.id;
      if (propia) setSedeId(propia);
    });
    articulosApi.listar().then(setArticulos);
  }, []);

  // Búsqueda de proveedores con debounce sencillo.
  useEffect(() => {
    const t = setTimeout(async () => {
      if (provBuscar.trim().length < 2) {
        setProvOpc([]);
        return;
      }
      try {
        setProvOpc(await proveedoresApi.listar(provBuscar));
      } catch {
        setProvOpc([]);
      }
    }, 300);
    return () => clearTimeout(t);
  }, [provBuscar]);

  function actualizarItem(idx: number, campo: keyof NuevaCompraItem, valor: string | number) {
    setItems((arr) => arr.map((it, i) => (i === idx ? { ...it, [campo]: valor } : it)));
  }
  const agregarItem = () => setItems((a) => [...a, { articuloId: "", cantidad: 1, costoUnitario: 0 }]);
  const quitarItem = (idx: number) => setItems((a) => a.filter((_, i) => i !== idx));

  const total = items.reduce((acc, it) => acc + it.cantidad * it.costoUnitario - (it.descuento ?? 0), 0);

  async function onSubmit(e: FormEvent) {
    e.preventDefault();
    setError(null);
    if (items.some((it) => !it.articuloId)) {
      setError("Selecciona el artículo en todas las líneas.");
      return;
    }
    setGuardando(true);
    try {
      await comprasApi.crear({
        sedeId,
        proveedor: { documento, nombre },
        remisionProveedor: remisionProveedor || undefined,
        items,
      });
      onCreado();
    } catch (err) {
      setError(err instanceof Error ? err.message : "Error al crear la remisión.");
    } finally {
      setGuardando(false);
    }
  }

  return (
    <div className="modal-fondo">
      <div className="modal" style={{ maxWidth: 760 }} onClick={(e) => e.stopPropagation()}>
        <h2>+ Crear remisión de compra</h2>
        <form onSubmit={onSubmit}>
          {error && <div className="alerta-error">{error}</div>}

          <div className="grid-2">
            <div className="campo">
              <label>Sede *</label>
              <select value={sedeId} onChange={(e) => setSedeId(e.target.value)} required>
                {sedes.map((s) => (
                  <option key={s.id} value={s.id}>{s.nombre}</option>
                ))}
              </select>
            </div>
            <div className="campo">
              <label>Remisión del proveedor</label>
              <input value={remisionProveedor} onChange={(e) => setRemisionProveedor(e.target.value)} />
            </div>
          </div>

          {/* Buscar un proveedor existente por nombre o NIT (rellena los campos). */}
          <div className="campo" style={{ position: "relative" }}>
            <label>Buscar proveedor (nombre o NIT)</label>
            <input
              value={provBuscar}
              onChange={(e) => setProvBuscar(e.target.value)}
              placeholder="Escribe para reutilizar un proveedor existente…"
            />
            {provOpc.length > 0 && (
              <div
                className="lista-opciones"
                style={{ position: "absolute", zIndex: 10, width: "100%", maxHeight: 220, overflowY: "auto" }}
              >
                {provOpc.map((p) => (
                  <button
                    type="button"
                    key={p.id}
                    className="opcion"
                    onMouseDown={() => {
                      setDocumento(p.documento);
                      setNombre(p.nombre);
                      setProvBuscar("");
                      setProvOpc([]);
                    }}
                  >
                    <strong>{p.nombre}</strong> · NIT {p.documento}
                  </button>
                ))}
              </div>
            )}
          </div>

          <div className="grid-2">
            <div className="campo">
              <label>NIT / Documento proveedor *</label>
              <input value={documento} onChange={(e) => setDocumento(e.target.value)} required />
            </div>
            <div className="campo">
              <label>Nombre proveedor *</label>
              <input value={nombre} onChange={(e) => setNombre(e.target.value)} required />
            </div>
          </div>

          <h3 style={{ margin: "10px 0" }}>Conceptos</h3>
          {items.map((it, idx) => (
            <div key={idx} className="linea-item">
              <SelectorArticulo
                articulos={articulos}
                valor={it.articuloId}
                onElegir={(id) => actualizarItem(idx, "articuloId", id)}
              />
              <input
                type="number" min="1" step="1" title="Cantidad" style={{ width: 80 }}
                value={it.cantidad}
                onChange={(e) => actualizarItem(idx, "cantidad", Number(e.target.value))}
              />
              <input
                type="number" min="0" title="Costo unitario" style={{ width: 130 }}
                value={it.costoUnitario}
                onChange={(e) => actualizarItem(idx, "costoUnitario", Number(e.target.value))}
              />
              <span className="muted" style={{ minWidth: 110, textAlign: "right" }}>
                {moneda(it.cantidad * it.costoUnitario)}
              </span>
              {items.length > 1 && (
                <button type="button" className="btn-quitar" onClick={() => quitarItem(idx)}>×</button>
              )}
            </div>
          ))}
          <button type="button" className="btn-secundario" onClick={agregarItem}>+ Agregar artículo</button>

          <div className="total-neto">TOTAL NETO: <strong>{moneda(total)}</strong></div>
          <p className="muted">Moneda: COP · Forma de pago: Contado</p>

          <div className="modal-acciones">
            <button type="button" className="btn-secundario" onClick={onCerrar}>Cancelar</button>
            <button type="submit" className="btn-primario" style={{ width: "auto" }} disabled={guardando}>
              {guardando ? "Guardando…" : "Crear remisión de compra"}
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}
