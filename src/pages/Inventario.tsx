// ==========================================================================
// Pantalla de Inventario con tres pestañas:
//  - Existencias: stock y costo promedio por sede + ajustes manuales.
//  - Traslados:   mover artículos entre Bogotá y Medellín.
//  - Trazabilidad: kardex de movimientos (compras, ajustes, traslados…).
// ==========================================================================
import { useEffect, useState, type ChangeEvent, type FormEvent } from "react";
import * as XLSX from "xlsx";
import {
  inventarioApi,
  catalogosApi,
  articulosApi,
  type Existencia,
  type Movimiento,
  type Sede,
  type Articulo,
  type FilaImportAjuste,
} from "../api/recursos";
import { SelectorArticulo } from "../components/SelectorArticulo";
import { leerLibroExcel } from "../utils/excel";
import { useAuth } from "../auth/AuthContext";

const moneda = (v: number | string) =>
  new Intl.NumberFormat("es-CO", { style: "currency", currency: "COP", maximumFractionDigits: 0 }).format(Number(v));

type Pestana = "existencias" | "traslados" | "trazabilidad";

export function Inventario() {
  const [pestana, setPestana] = useState<Pestana>("existencias");

  return (
    <div>
      <div className="dash-topbar">
        <h2>Inventario</h2>
      </div>

      {/* Pestañas */}
      <div className="tabs">
        <button className={pestana === "existencias" ? "tab activo" : "tab"} onClick={() => setPestana("existencias")}>
          Existencias
        </button>
        <button className={pestana === "traslados" ? "tab activo" : "tab"} onClick={() => setPestana("traslados")}>
          Traslados
        </button>
        <button className={pestana === "trazabilidad" ? "tab activo" : "tab"} onClick={() => setPestana("trazabilidad")}>
          Trazabilidad
        </button>
      </div>

      {pestana === "existencias" && <Existencias />}
      {pestana === "traslados" && <Traslados />}
      {pestana === "trazabilidad" && <Trazabilidad />}
    </div>
  );
}

// --------------------------------------------------------------------------
// Pestaña EXISTENCIAS: stock y costo promedio por sede, con ajuste manual.
// --------------------------------------------------------------------------
function Existencias() {
  const { usuario } = useAuth();
  const esAdmin = usuario?.rol === "ADMIN"; // el operador no ve el costo promedio
  const [filas, setFilas] = useState<Existencia[]>([]);
  const [sedes, setSedes] = useState<Sede[]>([]);
  const [sedeId, setSedeId] = useState("");
  const [buscar, setBuscar] = useState("");
  const [cargando, setCargando] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [aviso, setAviso] = useState<string | null>(null);
  const [ajuste, setAjuste] = useState<Existencia | null>(null);
  const [importando, setImportando] = useState(false);

  async function cargar() {
    setCargando(true);
    setError(null);
    try {
      setFilas(await inventarioApi.existencias(sedeId, buscar));
    } catch (e) {
      setError(e instanceof Error ? e.message : "Error al cargar existencias.");
    } finally {
      setCargando(false);
    }
  }

  // Plantilla de Excel para el ajuste por conteo físico.
  function descargarPlantilla() {
    const datos = [
      ["Sede", "Código", "Cantidad real", "Costo unitario", "Observación"],
      ["Medellín", "000006", 12, 120000, "Conteo físico"],
      ["Bogotá", "35", 8, 110000, ""],
    ];
    const hoja = XLSX.utils.aoa_to_sheet(datos);
    const libro = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(libro, hoja, "Ajuste");
    XLSX.writeFile(libro, "plantilla_ajuste_inventario.xlsx");
  }

  // Importa el conteo físico: deja el stock de cada artículo en la cantidad real.
  async function onArchivoExcel(e: ChangeEvent<HTMLInputElement>) {
    const archivo = e.target.files?.[0];
    if (!archivo) return;
    setError(null);
    setAviso(null);
    setImportando(true);
    try {
      const libro = await leerLibroExcel(archivo);
      const hoja = libro.Sheets[libro.SheetNames[0]];
      const matriz = XLSX.utils.sheet_to_json<(string | number)[]>(hoja, { header: 1, blankrows: false });
      if (matriz.length < 2) throw new Error("El archivo no tiene filas de datos.");
      const norm = (s: unknown) =>
        String(s ?? "").toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/[^a-z0-9]/g, "");
      const enc = (matriz[0] as unknown[]).map((c) => norm(c));
      const idx = {
        sede: enc.findIndex((h) => h.includes("sede")),
        codigo: enc.findIndex((h) => h.includes("codigo") || h.includes("cod")),
        cantidad: enc.findIndex((h) => h.includes("cantidad") || h.includes("real") || h.includes("conteo")),
        costo: enc.findIndex((h) => h.includes("costo")),
        obs: enc.findIndex((h) => h.includes("observ")),
      };
      if (idx.sede < 0 || idx.codigo < 0 || idx.cantidad < 0) {
        throw new Error("No se reconocieron las columnas Sede, Código y Cantidad real.");
      }
      const filas: FilaImportAjuste[] = [];
      for (const row of matriz.slice(1) as unknown[][]) {
        const sede = String(row[idx.sede] ?? "").trim();
        const codigo = String(row[idx.codigo] ?? "").trim();
        const cantRaw = row[idx.cantidad];
        if (!sede || !codigo || cantRaw === undefined || cantRaw === null || String(cantRaw).trim() === "") continue;
        const cantidadReal = Number(cantRaw);
        if (!Number.isFinite(cantidadReal)) continue;
        const fila: FilaImportAjuste = { sede, articuloCodigo: codigo, cantidadReal };
        if (idx.costo >= 0 && row[idx.costo] !== undefined && String(row[idx.costo]).trim() !== "") {
          const c = Number(row[idx.costo]);
          if (Number.isFinite(c)) fila.costoUnitario = c;
        }
        if (idx.obs >= 0 && row[idx.obs]) fila.observacion = String(row[idx.obs]).trim();
        filas.push(fila);
      }
      if (!filas.length) throw new Error("No se encontraron filas válidas (Sede, Código y Cantidad real).");
      const r = await inventarioApi.importarAjustes(filas);
      setAviso(
        `Ajuste aplicado: ${r.ajustados} artículos actualizados` +
          (r.sinCambio ? `, ${r.sinCambio} sin cambio` : "") +
          (r.errores.length ? `, ${r.errores.length} con error` : "") +
          "."
      );
      cargar();
    } catch (err) {
      setError(err instanceof Error ? err.message : "Error al importar el ajuste.");
    } finally {
      setImportando(false);
      e.target.value = "";
    }
  }

  // Deshace el último lote de ajuste importado (restaura el stock previo).
  async function revertirAjuste() {
    if (!window.confirm(
      "Deshacer el ÚLTIMO ajuste importado: se devolverá el stock de esos artículos al valor que tenían antes. ¿Continuar?"
    )) return;
    setError(null);
    setAviso(null);
    setImportando(true);
    try {
      const r = await inventarioApi.revertirAjuste();
      setAviso(
        `Ajuste deshecho: ${r.revertidos} de ${r.enLote} artículos restaurados` +
          (r.errores.length ? `, ${r.errores.length} con error` : "") +
          ". Revisa el stock y, si hace falta, el costo de los que habían subido."
      );
      cargar();
    } catch (e) {
      setError(e instanceof Error ? e.message : "Error al deshacer el ajuste.");
    } finally {
      setImportando(false);
    }
  }

  useEffect(() => {
    catalogosApi.sedes().then(setSedes);
  }, []);
  useEffect(() => {
    cargar();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sedeId]);

  return (
    <div>
      <div className="barra-busqueda">
        <select value={sedeId} onChange={(e) => setSedeId(e.target.value)} style={{ maxWidth: 200 }}>
          <option value="">Todas las sedes</option>
          {sedes.map((s) => (
            <option key={s.id} value={s.id}>{s.nombre}</option>
          ))}
        </select>
        <input
          placeholder="Buscar por nombre o código…"
          value={buscar}
          onChange={(e) => setBuscar(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && cargar()}
        />
        <button className="btn-secundario" onClick={cargar}>Buscar</button>
        <button className="btn-secundario" onClick={descargarPlantilla}>Plantilla ajuste</button>
        <label
          className="btn-secundario"
          style={{ cursor: importando ? "wait" : "pointer", opacity: importando ? 0.6 : 1 }}
        >
          {importando ? "Importando…" : "Importar ajuste (Excel)"}
          <input type="file" accept=".xls,.xlsx" hidden disabled={importando} onChange={onArchivoExcel} />
        </label>
        {esAdmin && (
          <button
            className="btn-secundario"
            style={{ color: "var(--echo-coral)", borderColor: "var(--echo-coral)" }}
            disabled={importando}
            onClick={revertirAjuste}
          >
            Deshacer último ajuste
          </button>
        )}
      </div>
      <p className="muted" style={{ marginTop: -8, marginBottom: 12 }}>
        El ajuste por Excel es un <strong>conteo físico</strong>: escribe la cantidad real de
        cada artículo y el sistema deja el stock exactamente en ese valor (por sede).
      </p>

      {error && <div className="alerta-error">{error}</div>}
      {aviso && <div className="alerta-ok">{aviso}</div>}

      {cargando ? (
        <p>Cargando…</p>
      ) : (
        <div className="tabla-wrap">
          <table className="tabla">
            <thead>
              <tr>
                <th>Código</th>
                <th>Artículo</th>
                <th>Sede</th>
                <th>Cantidad</th>
                {esAdmin && <th>Costo promedio</th>}
                <th></th>
              </tr>
            </thead>
            <tbody>
              {filas.map((f) => (
                <tr key={f.id}>
                  <td>{f.articulo.codigo}</td>
                  <td><strong>{f.articulo.nombre}</strong></td>
                  <td><span className="badge-sede">{f.sede.nombre}</span></td>
                  <td>
                    <strong>{Number(f.cantidad)}</strong>
                    {Number(f.cantidad) <= Number(f.stockMinimo) && Number(f.stockMinimo) > 0 && (
                      <span className="muted" style={{ color: "var(--echo-coral)" }}> · bajo</span>
                    )}
                  </td>
                  {esAdmin && <td>{moneda(f.costoPromedio)}</td>}
                  <td>
                    <button className="btn-secundario" style={{ padding: "6px 12px" }} onClick={() => setAjuste(f)}>
                      Ajustar
                    </button>
                  </td>
                </tr>
              ))}
              {filas.length === 0 && (
                <tr>
                  <td colSpan={esAdmin ? 6 : 5} className="muted" style={{ textAlign: "center", padding: 24 }}>
                    No hay existencias para los filtros seleccionados.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      )}

      {ajuste && (
        <ModalAjuste
          existencia={ajuste}
          onCerrar={() => setAjuste(null)}
          onHecho={(msg) => {
            setAjuste(null);
            setAviso(msg);
            cargar();
          }}
        />
      )}
    </div>
  );
}

// --------------------------------------------------------------------------
// Modal de AJUSTE de existencias (entrada/salida manual sobre una sede).
// --------------------------------------------------------------------------
function ModalAjuste({
  existencia,
  onCerrar,
  onHecho,
}: {
  existencia: Existencia;
  onCerrar: () => void;
  onHecho: (msg: string) => void;
}) {
  const { usuario } = useAuth();
  const esAdmin = usuario?.rol === "ADMIN"; // el operador no ve/edita el costo
  const [tipo, setTipo] = useState<"entrada" | "salida">("entrada");
  const [cantidad, setCantidad] = useState(1);
  const [costoUnitario, setCostoUnitario] = useState(Number(existencia.costoPromedio));
  const [observacion, setObservacion] = useState("");
  const [error, setError] = useState<string | null>(null);
  const [guardando, setGuardando] = useState(false);

  async function onSubmit(e: FormEvent) {
    e.preventDefault();
    setError(null);
    setGuardando(true);
    try {
      // Salida = cantidad negativa. Entrada = positiva (con costo para el promedio).
      const signo = tipo === "salida" ? -1 : 1;
      await inventarioApi.ajustar({
        articuloId: existencia.articulo.id,
        sedeId: existencia.sede.id,
        cantidad: signo * cantidad,
        // El operador no fija costo: las entradas conservan el costo promedio actual.
        costoUnitario: tipo === "entrada" && esAdmin ? costoUnitario : undefined,
        observacion: observacion || undefined,
      });
      onHecho(`Ajuste aplicado a ${existencia.articulo.nombre} en ${existencia.sede.nombre}.`);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Error al aplicar el ajuste.");
    } finally {
      setGuardando(false);
    }
  }

  return (
    <div className="modal-fondo">
      <div className="modal" style={{ maxWidth: 480 }} onClick={(e) => e.stopPropagation()}>
        <h2>Ajustar existencias</h2>
        <p className="muted" style={{ marginTop: -10 }}>
          {existencia.articulo.codigo} · {existencia.articulo.nombre} — {existencia.sede.nombre}
          {" · "}stock actual: <strong>{Number(existencia.cantidad)}</strong>
        </p>
        <form onSubmit={onSubmit}>
          {error && <div className="alerta-error">{error}</div>}

          <div className="grid-2">
            <div className="campo">
              <label>Tipo de ajuste *</label>
              <select value={tipo} onChange={(e) => setTipo(e.target.value as "entrada" | "salida")}>
                <option value="entrada">Entrada (sumar)</option>
                <option value="salida">Salida (restar)</option>
              </select>
            </div>
            <div className="campo">
              <label>Cantidad *</label>
              <input
                type="number" min="1" step="1" required
                value={cantidad}
                onChange={(e) => setCantidad(Number(e.target.value))}
              />
            </div>
          </div>

          {tipo === "entrada" && esAdmin && (
            <div className="campo">
              <label>Costo unitario (recalcula el costo promedio)</label>
              <input
                type="number" min="0"
                value={costoUnitario}
                onChange={(e) => setCostoUnitario(Number(e.target.value))}
              />
            </div>
          )}

          <div className="campo">
            <label>Observación</label>
            <input value={observacion} onChange={(e) => setObservacion(e.target.value)} placeholder="Motivo del ajuste" />
          </div>

          <div className="modal-acciones">
            <button type="button" className="btn-secundario" onClick={onCerrar}>Cancelar</button>
            <button type="submit" className="btn-primario" style={{ width: "auto" }} disabled={guardando}>
              {guardando ? "Guardando…" : "Aplicar ajuste"}
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}

// --------------------------------------------------------------------------
// Pestaña TRASLADOS: mover artículos de una sede a otra.
// --------------------------------------------------------------------------
function Traslados() {
  const [sedes, setSedes] = useState<Sede[]>([]);
  const [articulos, setArticulos] = useState<Articulo[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [aviso, setAviso] = useState<string | null>(null);
  const [guardando, setGuardando] = useState(false);

  const [sedeOrigenId, setSedeOrigenId] = useState("");
  const [sedeDestinoId, setSedeDestinoId] = useState("");
  const [observacion, setObservacion] = useState("");
  const [items, setItems] = useState<{ articuloId: string; cantidad: number }[]>([
    { articuloId: "", cantidad: 1 },
  ]);

  useEffect(() => {
    catalogosApi.sedes().then((s) => {
      setSedes(s);
      if (s[0]) setSedeOrigenId(s[0].id);
      if (s[1]) setSedeDestinoId(s[1].id);
    });
    articulosApi.listar().then(setArticulos);
  }, []);

  function actualizarItem(idx: number, campo: "articuloId" | "cantidad", valor: string | number) {
    setItems((arr) => arr.map((it, i) => (i === idx ? { ...it, [campo]: valor } : it)));
  }
  const agregarItem = () => setItems((a) => [...a, { articuloId: "", cantidad: 1 }]);
  const quitarItem = (idx: number) => setItems((a) => a.filter((_, i) => i !== idx));

  async function onSubmit(e: FormEvent) {
    e.preventDefault();
    setError(null);
    setAviso(null);
    if (sedeOrigenId === sedeDestinoId) {
      setError("La sede de origen y destino deben ser distintas.");
      return;
    }
    if (items.some((it) => !it.articuloId)) {
      setError("Selecciona el artículo en todas las líneas.");
      return;
    }
    setGuardando(true);
    try {
      const r = await inventarioApi.trasladar({ sedeOrigenId, sedeDestinoId, observacion: observacion || undefined, items });
      setAviso(`Traslado #${r.numero} registrado correctamente.`);
      setItems([{ articuloId: "", cantidad: 1 }]);
      setObservacion("");
    } catch (err) {
      setError(err instanceof Error ? err.message : "Error al registrar el traslado.");
    } finally {
      setGuardando(false);
    }
  }

  return (
    <div className="tabla-wrap" style={{ padding: 24 }}>
      {error && <div className="alerta-error">{error}</div>}
      {aviso && <div className="alerta-ok">{aviso}</div>}

      <form onSubmit={onSubmit}>
        <div className="grid-2">
          <div className="campo">
            <label>Sede de origen *</label>
            <select value={sedeOrigenId} onChange={(e) => setSedeOrigenId(e.target.value)} required>
              {sedes.map((s) => (
                <option key={s.id} value={s.id}>{s.nombre}</option>
              ))}
            </select>
          </div>
          <div className="campo">
            <label>Sede de destino *</label>
            <select value={sedeDestinoId} onChange={(e) => setSedeDestinoId(e.target.value)} required>
              {sedes.map((s) => (
                <option key={s.id} value={s.id}>{s.nombre}</option>
              ))}
            </select>
          </div>
        </div>

        <h3 style={{ margin: "10px 0" }}>Artículos a trasladar</h3>
        {items.map((it, idx) => (
          <div key={idx} className="linea-item">
            <SelectorArticulo
              articulos={articulos}
              valor={it.articuloId}
              onElegir={(id) => actualizarItem(idx, "articuloId", id)}
            />
            <input
              type="number" min="1" step="1" title="Cantidad" style={{ width: 90 }}
              value={it.cantidad}
              onChange={(e) => actualizarItem(idx, "cantidad", Number(e.target.value))}
            />
            {items.length > 1 && (
              <button type="button" className="btn-quitar" onClick={() => quitarItem(idx)}>×</button>
            )}
          </div>
        ))}
        <button type="button" className="btn-secundario" onClick={agregarItem}>+ Agregar artículo</button>

        <div className="campo" style={{ marginTop: 16 }}>
          <label>Observación</label>
          <input value={observacion} onChange={(e) => setObservacion(e.target.value)} placeholder="Motivo del traslado" />
        </div>

        <div className="modal-acciones">
          <button type="submit" className="btn-primario" style={{ width: "auto" }} disabled={guardando}>
            {guardando ? "Registrando…" : "Registrar traslado"}
          </button>
        </div>
      </form>
    </div>
  );
}

// --------------------------------------------------------------------------
// Pestaña TRAZABILIDAD: kardex de movimientos con filtros por sede y tipo.
// --------------------------------------------------------------------------
const ETIQUETA_TIPO: Record<string, string> = {
  COMPRA: "Compra",
  VENTA_REMISION: "Venta / remisión",
  TRASLADO_ENTRADA: "Traslado (entrada)",
  TRASLADO_SALIDA: "Traslado (salida)",
  AJUSTE: "Ajuste",
};

function Trazabilidad() {
  const [movimientos, setMovimientos] = useState<Movimiento[]>([]);
  const [sedes, setSedes] = useState<Sede[]>([]);
  const [sedeId, setSedeId] = useState("");
  const [tipo, setTipo] = useState("");
  const [cargando, setCargando] = useState(true);
  const [error, setError] = useState<string | null>(null);

  async function cargar() {
    setCargando(true);
    setError(null);
    try {
      setMovimientos(await inventarioApi.movimientos({ sedeId: sedeId || undefined, tipo: tipo || undefined }));
    } catch (e) {
      setError(e instanceof Error ? e.message : "Error al cargar la trazabilidad.");
    } finally {
      setCargando(false);
    }
  }

  useEffect(() => {
    catalogosApi.sedes().then(setSedes);
  }, []);
  useEffect(() => {
    cargar();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sedeId, tipo]);

  return (
    <div>
      <div className="barra-busqueda">
        <select value={sedeId} onChange={(e) => setSedeId(e.target.value)} style={{ maxWidth: 200 }}>
          <option value="">Todas las sedes</option>
          {sedes.map((s) => (
            <option key={s.id} value={s.id}>{s.nombre}</option>
          ))}
        </select>
        <select value={tipo} onChange={(e) => setTipo(e.target.value)} style={{ maxWidth: 220 }}>
          <option value="">Todos los movimientos</option>
          {Object.entries(ETIQUETA_TIPO).map(([k, v]) => (
            <option key={k} value={k}>{v}</option>
          ))}
        </select>
      </div>

      {error && <div className="alerta-error">{error}</div>}

      {cargando ? (
        <p>Cargando…</p>
      ) : (
        <div className="tabla-wrap">
          <table className="tabla">
            <thead>
              <tr>
                <th>Fecha</th>
                <th>Artículo</th>
                <th>Sede</th>
                <th>Movimiento</th>
                <th>Cantidad</th>
                <th>Referencia</th>
              </tr>
            </thead>
            <tbody>
              {movimientos.map((m) => {
                const cant = Number(m.cantidad);
                return (
                  <tr key={m.id}>
                    <td>{new Date(m.fecha).toLocaleString("es-CO")}</td>
                    <td>
                      <strong>{m.articulo.nombre}</strong>
                      <div className="muted">{m.articulo.codigo}</div>
                    </td>
                    <td>{m.sede.nombre}</td>
                    <td>{ETIQUETA_TIPO[m.tipo] ?? m.tipo}</td>
                    <td style={{ color: cant < 0 ? "var(--echo-coral)" : "var(--echo-azul)", fontWeight: 500 }}>
                      {cant > 0 ? `+${cant}` : cant}
                    </td>
                    <td className="muted">{m.referencia ?? "—"}</td>
                  </tr>
                );
              })}
              {movimientos.length === 0 && (
                <tr>
                  <td colSpan={6} className="muted" style={{ textAlign: "center", padding: 24 }}>
                    No hay movimientos para los filtros seleccionados.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}
