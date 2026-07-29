// ==========================================================================
// Pantalla de Remisiones de compra.
// Lista las remisiones, permite crear una nueva (que actualiza inventario y
// costos) e importar masivamente desde un archivo Excel.
// ==========================================================================
import { useEffect, useRef, useState, type ChangeEvent, type FormEvent } from "react";
import * as XLSX from "xlsx";
import {
  comprasApi,
  catalogosApi,
  articulosApi,
  proveedoresApi,
  type Compra,
  type CompraCompleta,
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

type DatosPreImport = {
  sede: string;
  documento: string;
  nombre: string;
  remisionProveedor: string;
  filas: FilaImportCompra[];
};

export function Compras() {
  const [compras, setCompras] = useState<Compra[]>([]);
  const [cargando, setCargando] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [aviso, setAviso] = useState<string | null>(null);
  const [modal, setModal] = useState(false);
  const [preImport, setPreImport] = useState<DatosPreImport | null>(null);
  const [editar, setEditar] = useState<Compra | null>(null);
  const [verPdf, setVerPdf] = useState<CompraCompleta | null>(null);

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

  async function abrirPdf(id: string) {
    try {
      setVerPdf(await comprasApi.obtener(id));
    } catch (e) {
      setError(e instanceof Error ? e.message : "Error al cargar el detalle.");
    }
  }

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

  // --- Importación desde Excel: solo lee el archivo y abre el modal pre-cargado ---
  async function onArchivoExcel(e: ChangeEvent<HTMLInputElement>) {
    const archivo = e.target.files?.[0];
    if (!archivo) return;
    setError(null);
    setAviso(null);
    try {
      const buffer = await archivo.arrayBuffer();
      const libro = XLSX.read(buffer);
      const hoja = libro.Sheets[libro.SheetNames[0]];
      const filas = XLSX.utils.sheet_to_json<FilaImportCompra>(hoja);
      if (filas.length === 0) throw new Error("El archivo no tiene filas.");
      const primera = filas[0];
      setPreImport({
        sede: String(primera.sede ?? ""),
        documento: String(primera.proveedorDocumento ?? ""),
        nombre: String(primera.proveedorNombre ?? ""),
        remisionProveedor: String(primera.remisionProveedor ?? ""),
        filas,
      });
      setModal(true);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Error al leer el Excel.");
    } finally {
      e.target.value = "";
    }
  }

  // Descarga una plantilla de Excel con los encabezados esperados y una hoja
  // de referencia con las sedes y códigos de artículos disponibles.
  async function descargarPlantilla() {
    const [sedes, articulos] = await Promise.all([
      catalogosApi.sedes(),
      articulosApi.listar(""),
    ]);

    const ejemplo: FilaImportCompra[] = [
      {
        sede: sedes[0]?.nombre ?? "Bogotá",
        proveedorDocumento: "912000471",
        proveedorNombre: "PROVEEDOR EJEMPLO S.A.",
        proveedorDireccion: "",
        remisionProveedor: "FAC-001",
        articuloCodigo: articulos[0]?.codigo ?? "001",
        descripcion: articulos[0]?.nombre ?? "Producto ejemplo",
        cantidad: 1,
        costoUnitario: 100000,
        descuento: 0,
      },
    ];

    const hojaCompras = XLSX.utils.json_to_sheet(ejemplo);

    // Hoja de referencia: sedes y artículos válidos
    const refSedes = sedes.map((s) => ({ Sede: s.nombre }));
    const refArticulos = articulos.map((a) => ({
      "Código (articuloCodigo)": a.codigo,
      "Nombre del artículo": a.nombre,
    }));

    const hojaSedes = XLSX.utils.json_to_sheet(refSedes);
    const hojaArticulos = XLSX.utils.json_to_sheet(refArticulos);

    const libro = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(libro, hojaCompras, "Compras");
    XLSX.utils.book_append_sheet(libro, hojaSedes, "Sedes válidas");
    XLSX.utils.book_append_sheet(libro, hojaArticulos, "Artículos válidos");

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
                      <button className="btn-secundario" style={{ padding: "6px 10px" }} onClick={() => abrirPdf(c.id)}>
                        Ver / PDF
                      </button>
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

      {modal && (
        <ModalCrearCompra
          initialData={preImport}
          onCerrar={() => { setModal(false); setPreImport(null); }}
          onCreado={() => { setModal(false); setPreImport(null); cargar(); }}
        />
      )}

      {editar && (
        <ModalEditarCompra
          compra={editar}
          onCerrar={() => setEditar(null)}
          onGuardado={() => { setEditar(null); cargar(); }}
        />
      )}

      {verPdf && <CompraImprimible compra={verPdf} onCerrar={() => setVerPdf(null)} />}
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
function ModalCrearCompra({
  onCerrar,
  onCreado,
  initialData,
}: {
  onCerrar: () => void;
  onCreado: () => void;
  initialData?: DatosPreImport | null;
}) {
  const { usuario } = useAuth();
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

  const [provBuscar, setProvBuscar] = useState("");
  const [provOpc, setProvOpc] = useState<Proveedor[]>([]);

  useEffect(() => {
    Promise.all([catalogosApi.sedes(), articulosApi.listar("")]).then(([s, arts]) => {
      setSedes(s);
      setArticulos(arts);

      if (initialData) {
        // Pre-llenado desde Excel: buscar sede por nombre
        const sedePorNombre = s.find(
          (x) => x.nombre.trim().toLowerCase() === initialData.sede.trim().toLowerCase()
        );
        const sedeInicial = sedePorNombre?.id
          ?? (usuario?.sedeId && s.some((x) => x.id === usuario.sedeId) ? usuario.sedeId : s[0]?.id);
        if (sedeInicial) setSedeId(sedeInicial);

        setDocumento(initialData.documento);
        setNombre(initialData.nombre);
        setRemisionProveedor(initialData.remisionProveedor);

        // Mapear articuloCodigo → articuloId
        const porCodigo = new Map(arts.map((a) => [String(a.codigo).trim(), a.id]));
        const itemsMapeados = initialData.filas.map((f) => ({
          articuloId: porCodigo.get(String(f.articuloCodigo).trim()) ?? "",
          cantidad: Number(f.cantidad) || 1,
          costoUnitario: Number(f.costoUnitario) || 0,
          descuento: Number(f.descuento) || 0,
        }));
        if (itemsMapeados.length > 0) setItems(itemsMapeados);
      } else {
        const propia = usuario?.sedeId && s.some((x) => x.id === usuario.sedeId)
          ? usuario.sedeId
          : s[0]?.id;
        if (propia) setSedeId(propia);
      }
    });
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
                type="text" inputMode="numeric" title="Costo unitario" style={{ width: 130 }}
                value={it.costoUnitario ? it.costoUnitario.toLocaleString("es-CO") : ""}
                onChange={(e) => actualizarItem(idx, "costoUnitario", Number(e.target.value.replace(/\D/g, "")) || 0)}
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

// --------------------------------------------------------------------------
// Vista imprimible de la nota de remisión de compra.
// --------------------------------------------------------------------------
const EMISOR = {
  nombre: "ECHO TECNOLOGÍA",
  representante: "YESICA ZULUAGA OSPINA",
  nit: "1017175943",
  regimen: "Régimen ordinario No responsable de IVA",
  direccion: "Antioquia / Medellín / Carrera 55 #12Sur 09 Torre 3 Apto 9920",
  telefonos: "3207548718",
  email: "echotecnologia@echotecnologia.co",
  web: "echotecnologia.co",
};

function CompraImprimible({ compra: c, onCerrar }: { compra: CompraCompleta; onCerrar: () => void }) {
  const ref = useRef<HTMLDivElement>(null);
  const fecha = new Date(c.fecha).toLocaleString("es-CO");
  const cop = (v: number) => "$" + Math.round(v).toLocaleString("en-US");

  function imprimir() {
    const prev = document.title;
    document.title = `Compra ${c.documento}`;
    window.addEventListener("afterprint", () => { document.title = prev; }, { once: true });
    window.print();
  }

  return (
    <div className="modal-fondo">
      <div className="modal" style={{ maxWidth: 900, maxHeight: "92vh", overflow: "auto" }} onClick={(e) => e.stopPropagation()}>
        <div className="modal-acciones no-print" style={{ justifyContent: "space-between", marginTop: 0, marginBottom: 12 }}>
          <h2 style={{ margin: 0 }}>Compra {c.documento}</h2>
          <div style={{ display: "flex", gap: 10 }}>
            <button className="btn-secundario" onClick={onCerrar}>Cerrar</button>
            <button className="btn-primario" style={{ width: "auto" }} onClick={imprimir}>
              Imprimir / Guardar PDF
            </button>
          </div>
        </div>

        <div className="remision-print" ref={ref} style={{ fontFamily: "Arial, sans-serif", fontSize: 13, color: "#000" }}>
          {/* Encabezado */}
          <table style={{ width: "100%", borderCollapse: "collapse", marginBottom: 10 }}>
            <tbody>
              <tr>
                <td style={{ width: "20%", verticalAlign: "middle" }}>
                  <img src="/logo-echo.png" alt="Echo" style={{ maxWidth: 90, maxHeight: 70 }} onError={(e) => { (e.target as HTMLImageElement).style.display = "none"; }} />
                </td>
                <td style={{ textAlign: "center", verticalAlign: "middle" }}>
                  <div style={{ fontWeight: "bold", fontSize: 16 }}>{EMISOR.nombre}</div>
                  <div style={{ fontWeight: "bold" }}>{EMISOR.representante}</div>
                  <div><strong>NIT:</strong> {EMISOR.nit} | {EMISOR.regimen}</div>
                  <div>{EMISOR.direccion}</div>
                  <div><strong>Teléfonos:</strong> {EMISOR.telefonos}</div>
                  <div><strong>Email:</strong> {EMISOR.email} &nbsp; <strong>Página web:</strong> {EMISOR.web}</div>
                </td>
                <td style={{ width: "22%", textAlign: "right", verticalAlign: "top" }}>
                  <div style={{ fontWeight: "bold", fontSize: 14 }}>NOTA DE REMISIÓN DE<br />COMPRA</div>
                  <div style={{ marginTop: 6 }}><strong>No.</strong> {c.documento}</div>
                  <div>{fecha}</div>
                </td>
              </tr>
            </tbody>
          </table>

          {/* Datos de la compra */}
          <table style={{ width: "100%", borderCollapse: "collapse", border: "1px solid #000", marginBottom: 8, fontSize: 12 }}>
            <tbody>
              <tr>
                <td style={{ border: "1px solid #000", padding: "4px 8px", width: "50%" }}>
                  <strong>Proveedor:</strong> {c.proveedor} &nbsp; NE. {c.proveedorDocumento}
                </td>
                <td style={{ border: "1px solid #000", padding: "4px 8px" }}>
                  <strong>Remisión del proveedor:</strong> {c.remisionProveedor ?? "—"}
                </td>
              </tr>
              <tr>
                <td colSpan={2} style={{ border: "1px solid #000", padding: "4px 8px" }}>
                  <strong>Dirección:</strong> {c.proveedorDireccion ?? "—"}
                </td>
              </tr>
              <tr>
                <td style={{ border: "1px solid #000", padding: "4px 8px" }}>
                  <strong>Sucursal:</strong> {c.sede}
                </td>
                <td style={{ border: "1px solid #000", padding: "4px 8px" }}>
                  <strong>Bodega:</strong> {c.sede}
                </td>
              </tr>
              {c.elaboradoPor && (
                <tr>
                  <td colSpan={2} style={{ border: "1px solid #000", padding: "4px 8px" }}>
                    <strong>Elaboró:</strong> {c.elaboradoPor}
                  </td>
                </tr>
              )}
            </tbody>
          </table>

          {/* Tabla de artículos */}
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12, marginBottom: 8 }}>
            <thead>
              <tr style={{ background: "#f0f0f0" }}>
                <th style={{ border: "1px solid #000", padding: "4px 6px", textAlign: "left" }}>REF.</th>
                <th style={{ border: "1px solid #000", padding: "4px 6px", textAlign: "left" }}>DESCRIPCIÓN</th>
                <th style={{ border: "1px solid #000", padding: "4px 6px", textAlign: "center" }}>CANTIDAD</th>
                <th style={{ border: "1px solid #000", padding: "4px 6px", textAlign: "right" }}>PRECIO UD.</th>
                <th style={{ border: "1px solid #000", padding: "4px 6px", textAlign: "right" }}>BRUTO</th>
                <th style={{ border: "1px solid #000", padding: "4px 6px", textAlign: "right" }}>IMPUESTOS</th>
              </tr>
            </thead>
            <tbody>
              {c.detalles.map((d) => (
                <tr key={d.id}>
                  <td style={{ border: "1px solid #000", padding: "4px 6px" }}>{d.codigo}</td>
                  <td style={{ border: "1px solid #000", padding: "4px 6px" }}>{d.descripcion}</td>
                  <td style={{ border: "1px solid #000", padding: "4px 6px", textAlign: "center" }}>{d.cantidad}</td>
                  <td style={{ border: "1px solid #000", padding: "4px 6px", textAlign: "right" }}>{cop(d.costoUnitario)}</td>
                  <td style={{ border: "1px solid #000", padding: "4px 6px", textAlign: "right" }}>{cop(d.cantidad * d.costoUnitario)}</td>
                  <td style={{ border: "1px solid #000", padding: "4px 6px", textAlign: "right" }}>—</td>
                </tr>
              ))}
            </tbody>
          </table>

          {/* Pie */}
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
            <tbody>
              <tr>
                <td style={{ border: "1px solid #000", padding: "6px 8px", verticalAlign: "top", width: "35%" }}>
                  <strong>FORMAS DE PAGO:</strong><br />
                  Contado - Valor: {cop(c.total)}<br />
                  <br /><strong>MEDIOS DE PAGO:</strong><br />
                  Efectivo - Valor: {cop(c.total)}
                </td>
                <td style={{ border: "1px solid #000", padding: "6px 8px", verticalAlign: "top" }}>
                  <strong>DETALLE DE IMPUESTOS Y RETENCIONES:</strong><br />
                  {c.observacion ? c.observacion : "-No registra-"}
                </td>
                <td style={{ border: "1px solid #000", padding: "6px 8px", textAlign: "right", verticalAlign: "top", width: "22%" }}>
                  <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4 }}>
                    <span>SUBTOTAL</span><span>{cop(c.subtotal)}</span>
                  </div>
                  {c.descuento > 0 && (
                    <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4 }}>
                      <span>DESCUENTO</span><span>-{cop(c.descuento)}</span>
                    </div>
                  )}
                  <div style={{ display: "flex", justifyContent: "space-between", fontWeight: "bold", borderTop: "1px solid #000", paddingTop: 4 }}>
                    <span>TOTAL NETO</span><span>{cop(c.total)}</span>
                  </div>
                </td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}
