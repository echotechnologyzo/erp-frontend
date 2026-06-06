// ==========================================================================
// Pantalla de Artículos: listado con búsqueda y stock por sede + creación.
// Replica la lógica del módulo de inventario del ERP actual con la marca Echo.
// Incluye importación masiva desde el Excel exportado de Effi (mapea los
// encabezados tolerando acentos/codificación y carga el stock inicial por sede).
// ==========================================================================
import { useEffect, useState, type ChangeEvent, type FormEvent } from "react";
import * as XLSX from "xlsx";
import { useAuth } from "../auth/AuthContext";
import { leerLibroExcel } from "../utils/excel";
import {
  articulosApi,
  uploadsApi,
  catalogosApi,
  type Articulo,
  type NuevoArticulo,
  type Marca,
  type FilaImportArticulo,
} from "../api/recursos";

// Normaliza un encabezado: minúsculas, sin acentos ni caracteres raros.
const norm = (s: unknown) =>
  String(s ?? "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "") // quita acentos
    .replace(/[^a-z0-9]/g, "");

// Mapea los encabezados del Excel de artículos de Effi a los campos internos.
// Devuelve, por cada campo reconocido, el índice de su columna.
function mapearEncabezados(encabezados: string[]): Record<string, number> {
  const idx: Record<string, number> = {};
  const set = (campo: string, i: number) => {
    if (idx[campo] === undefined) idx[campo] = i;
  };
  encabezados.forEach((h, i) => {
    const n = norm(h);
    if (!n) return;
    if (n === "id") set("id", i); // ID interno de Effi → se usa para el código
    else if (n === "nombre") set("nombre", i);
    else if (n === "marca") set("marca", i);
    else if (n.includes("descripci")) set("descripcion", i);
    else if (n.includes("url")) set("urlImagen", i);
    else if (n.includes("costo") && n.includes("manual")) set("costoManual", i);
    else if (n.includes("costo") && n.includes("promedio")) set("costoPromedio", i);
    else if (n.includes("costo") && n.includes("ltimo")) set("ultimoCosto", i); // último
    else if (n.includes("minimo") || n.includes("mnimo")) set("precioMinimo", i);
    else if (n.includes("tarifa")) set("tarifaNormal", i);
    // Stock por bodega/sede (la columna "Stock total empresa" se ignora).
    else if (n.includes("stock") && n.includes("medell")) set("stockMedellin", i);
    else if (n.includes("stock") && n.includes("bogot")) set("stockBogota", i);
  });
  return idx;
}

// Campos numéricos (se convierten a número al armar la fila).
const CAMPOS_NUM = [
  "costoManual",
  "costoPromedio",
  "ultimoCosto",
  "precioMinimo",
  "tarifaNormal",
  "stockMedellin",
  "stockBogota",
];

export function Articulos() {
  const { usuario } = useAuth();
  const esAdmin = usuario?.rol === "ADMIN";
  const [articulos, setArticulos] = useState<Articulo[]>([]);
  const [buscar, setBuscar] = useState("");
  const [cargando, setCargando] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [aviso, setAviso] = useState<string | null>(null);
  const [importando, setImportando] = useState(false);
  // null = cerrado · "nuevo" = crear · Articulo = editar ese artículo
  const [editar, setEditar] = useState<Articulo | "nuevo" | null>(null);
  const [modalMarcas, setModalMarcas] = useState(false);

  // Carga la lista (se vuelve a llamar al buscar o tras crear un artículo).
  async function cargar() {
    setCargando(true);
    setError(null);
    try {
      setArticulos(await articulosApi.listar(buscar));
    } catch (e) {
      setError(e instanceof Error ? e.message : "Error al cargar artículos.");
    } finally {
      setCargando(false);
    }
  }

  useEffect(() => {
    cargar();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const formatoMoneda = (v: number) =>
    new Intl.NumberFormat("es-CO", { style: "currency", currency: "COP", maximumFractionDigits: 0 }).format(v);

  // --- Importación desde el Excel de Effi ---
  async function onArchivoExcel(e: ChangeEvent<HTMLInputElement>) {
    const archivo = e.target.files?.[0];
    if (!archivo) return;
    setError(null);
    setAviso(null);
    setImportando(true);
    try {
      const libro = await leerLibroExcel(archivo);
      const hoja = libro.Sheets[libro.SheetNames[0]];
      // Matriz (array de arrays) para mapear por posición de columna.
      const matriz = XLSX.utils.sheet_to_json<(string | number)[]>(hoja, {
        header: 1,
        blankrows: false,
      });
      if (matriz.length < 2) throw new Error("El archivo no tiene filas de datos.");

      const encabezados = (matriz[0] as unknown[]).map((c) => String(c ?? ""));
      const idx = mapearEncabezados(encabezados);
      if (idx.nombre === undefined) {
        throw new Error(
          "No se reconoció la columna de nombre. Verifica que sea el Excel de artículos de Effi."
        );
      }

      const campos = Object.keys(idx);
      const filas: FilaImportArticulo[] = [];
      for (const row of matriz.slice(1)) {
        const fila: FilaImportArticulo = {};
        for (const campo of campos) {
          const valor = (row as unknown[])[idx[campo]];
          if (valor === undefined || valor === null || String(valor).trim() === "") continue;
          if (campo === "id") {
            // El ID de Effi se convierte en código de referencia.
            fila.codigo = "EF-" + String(valor).trim();
          } else if (CAMPOS_NUM.includes(campo)) {
            const n = Number(valor);
            if (Number.isFinite(n)) (fila as Record<string, unknown>)[campo] = n;
          } else {
            (fila as Record<string, unknown>)[campo] = String(valor).trim();
          }
        }
        if (fila.nombre) filas.push(fila);
      }
      if (filas.length === 0) throw new Error("No se encontraron filas con nombre de artículo.");

      const r = await articulosApi.importar(filas);
      setAviso(
        `Importación finalizada: ${r.creadas} artículos creados` +
          (r.conStock ? ` (${r.conStock} con stock inicial)` : "") +
          (r.omitidas ? `, ${r.omitidas} ya existían` : "") +
          (r.descartadas ? `, ${r.descartadas} sin nombre` : "") +
          (r.errores.length ? `, ${r.errores.length} con error` : "") +
          "."
      );
      cargar();
    } catch (err) {
      setError(err instanceof Error ? err.message : "Error al importar el Excel.");
    } finally {
      setImportando(false);
      e.target.value = "";
    }
  }

  // --- Activar / desactivar un artículo ---
  async function alternarEstado(a: Articulo) {
    setError(null);
    setAviso(null);
    try {
      await articulosApi.actualizar(a.id, { activo: !a.activo });
      setAviso(`Artículo ${a.nombre} ${a.activo ? "desactivado" : "activado"}.`);
      cargar();
    } catch (e) {
      setError(e instanceof Error ? e.message : "Error al cambiar el estado.");
    }
  }

  // --- Eliminar un artículo (con confirmación) ---
  async function eliminar(a: Articulo) {
    if (!window.confirm(`¿Eliminar el artículo "${a.nombre}"? Esta acción no se puede deshacer.`)) {
      return;
    }
    setError(null);
    setAviso(null);
    try {
      await articulosApi.eliminar(a.id);
      setAviso(`Artículo ${a.nombre} eliminado.`);
      cargar();
    } catch (e) {
      setError(e instanceof Error ? e.message : "Error al eliminar el artículo.");
    }
  }

  // --- Borrar TODO el catálogo (para limpiar y reimportar) ---
  async function borrarTodos() {
    if (!window.confirm(
      "Esto BORRARÁ TODOS los artículos, su inventario y el kardex (stock de apertura). " +
      "Úsalo solo para reimportar el catálogo actualizado. ¿Continuar?"
    )) return;
    // Segunda confirmación por ser una acción irreversible.
    if (!window.confirm("Confirma de nuevo: se eliminarán TODOS los artículos. Esta acción no se puede deshacer.")) {
      return;
    }
    setError(null);
    setAviso(null);
    try {
      const r = await articulosApi.eliminarTodos();
      setAviso(`Se eliminaron ${r.eliminados} artículos. Ya puedes importar el catálogo actualizado.`);
      cargar();
    } catch (e) {
      const msg = e instanceof Error ? e.message : "Error al borrar los artículos.";
      // Si está bloqueado por documentos (remisiones/compras/traslados), ofrecer
      // el borrado total que también los limpia (útil durante la migración).
      if (/documentos|remisiones|compras|traslados/i.test(msg)) {
        if (window.confirm(
          msg +
          "\n\n¿Hacer un BORRADO TOTAL que elimine también esas remisiones, compras y traslados? " +
          "Solo recomendado durante la migración. Esta acción no se puede deshacer."
        )) {
          try {
            const r = await articulosApi.eliminarTodos(true);
            setAviso(`Borrado total: ${r.eliminados} artículos y todos los documentos de prueba eliminados. Ya puedes reimportar.`);
            cargar();
          } catch (e2) {
            setError(e2 instanceof Error ? e2.message : "Error en el borrado total.");
          }
        }
      } else {
        setError(msg);
      }
    }
  }

  // --- Plantilla de Excel para importar artículos ---
  function descargarPlantilla() {
    const ejemplo = [
      {
        ID: 1,
        Nombre: "AMAZON ECHO DOT 5",
        "Sucursal principal": "Medellín",
        Marca: "AMAZON",
        Descripción: "Parlante inteligente",
        "URL imagen": "",
        "Costo manual": 120000,
        "Costo promedio": 120000,
        "Último costo": 120000,
        "Precio mínimo de venta": 150000,
        "Precio: Tarifa normal": 199000,
        "Stock total empresa": 5,
        "Stock bodega: Medellín (Sucursal: Medellín)": 3,
        "Stock bodega: Bogotá (Sucursal: Bogotá)": 2,
      },
    ];
    const hoja = XLSX.utils.json_to_sheet(ejemplo);
    const libro = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(libro, hoja, "Artículos");
    XLSX.writeFile(libro, "plantilla_articulos_echo.xlsx");
  }

  return (
    <div>
      <div className="dash-topbar">
        <h2>Artículos</h2>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
          <button className="btn-secundario" onClick={descargarPlantilla}>
            Plantilla Excel
          </button>
          <button className="btn-secundario" onClick={() => setModalMarcas(true)}>
            Marcas
          </button>
          {esAdmin && (
            <button
              className="btn-secundario"
              style={{ color: "var(--echo-coral)", borderColor: "var(--echo-coral)" }}
              onClick={borrarTodos}
            >
              Borrar todos
            </button>
          )}
          <label
            className="btn-secundario"
            style={{ cursor: importando ? "wait" : "pointer", opacity: importando ? 0.6 : 1 }}
          >
            {importando ? "Importando…" : "Importar Excel (Effi)"}
            <input type="file" accept=".xls,.xlsx" hidden disabled={importando} onChange={onArchivoExcel} />
          </label>
          <button className="btn-primario" style={{ width: "auto" }} onClick={() => setEditar("nuevo")}>
            + Nuevo artículo
          </button>
        </div>
      </div>

      <div className="barra-busqueda">
        <input
          placeholder="Buscar por nombre o código…"
          value={buscar}
          onChange={(e) => setBuscar(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && cargar()}
        />
        <button className="btn-secundario" onClick={cargar}>
          Buscar
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
                <th>Código</th>
                <th>Artículo</th>
                <th>Categoría</th>
                <th>Tarifa normal</th>
                <th>Garantía</th>
                <th>Stock por sede</th>
                <th>Stock total</th>
                <th>Estado</th>
                <th></th>
              </tr>
            </thead>
            <tbody>
              {articulos.map((a) => (
                <tr key={a.id} style={{ opacity: a.activo ? 1 : 0.55 }}>
                  <td>{a.codigo}</td>
                  <td>
                    <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                      {a.urlFoto && (
                        <img
                          src={a.urlFoto}
                          alt={a.nombre}
                          className="miniatura-articulo"
                          onError={(e) => {
                            (e.currentTarget as HTMLImageElement).style.display = "none";
                          }}
                        />
                      )}
                      <div>
                        <strong>{a.nombre}</strong>
                        {a.marca && <div className="muted">{a.marca.nombre}</div>}
                      </div>
                    </div>
                  </td>
                  <td>{a.categoria?.nombre ?? "—"}</td>
                  <td>{formatoMoneda(a.tarifaNormal)}</td>
                  <td>{a.garantiaMeses ?? 6} meses</td>
                  <td>
                    {a.inventario.map((inv) => (
                      <span key={inv.sede.id} className="badge-sede">
                        {inv.sede.nombre}: {inv.cantidad}
                      </span>
                    ))}
                  </td>
                  <td>
                    <strong>{a.stockTotal}</strong>
                  </td>
                  <td>
                    <span className={a.activo ? "badge-nuevo" : "badge-sede"}>
                      {a.activo ? "Activo" : "Inactivo"}
                    </span>
                  </td>
                  <td>
                    <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                      <button className="btn-secundario" style={{ padding: "6px 10px" }} onClick={() => setEditar(a)}>
                        Editar
                      </button>
                      <button className="btn-secundario" style={{ padding: "6px 10px" }} onClick={() => alternarEstado(a)}>
                        {a.activo ? "Desactivar" : "Activar"}
                      </button>
                      <button
                        className="btn-secundario"
                        style={{ padding: "6px 10px", color: "var(--echo-coral)", borderColor: "var(--echo-coral)" }}
                        onClick={() => eliminar(a)}
                      >
                        Eliminar
                      </button>
                    </div>
                  </td>
                </tr>
              ))}
              {articulos.length === 0 && (
                <tr>
                  <td colSpan={9} className="muted" style={{ textAlign: "center", padding: 24 }}>
                    No hay artículos. Crea el primero con “Nuevo artículo”.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      )}

      {editar && (
        <ModalArticulo
          articulo={editar === "nuevo" ? null : editar}
          onCerrar={() => setEditar(null)}
          onGuardado={() => {
            setEditar(null);
            cargar();
          }}
        />
      )}

      {modalMarcas && (
        <ModalMarcas
          onCerrar={() => {
            setModalMarcas(false);
            cargar(); // refresca por si algún artículo quedó sin marca
          }}
        />
      )}
    </div>
  );
}

// --------------------------------------------------------------------------
// Modal para gestionar (eliminar) marcas. Al borrar una marca, los artículos
// que la usaban quedan sin marca (no se eliminan).
// --------------------------------------------------------------------------
function ModalMarcas({ onCerrar }: { onCerrar: () => void }) {
  const [marcas, setMarcas] = useState<Marca[]>([]);
  const [cargando, setCargando] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [aviso, setAviso] = useState<string | null>(null);

  async function cargar() {
    setCargando(true);
    try {
      setMarcas(await catalogosApi.marcas());
    } catch (e) {
      setError(e instanceof Error ? e.message : "Error al cargar marcas.");
    } finally {
      setCargando(false);
    }
  }

  useEffect(() => {
    cargar();
  }, []);

  async function eliminar(m: Marca) {
    if (!window.confirm(`¿Eliminar la marca "${m.nombre}"? Los artículos que la usen quedarán sin marca.`)) {
      return;
    }
    setError(null);
    setAviso(null);
    try {
      const r = await catalogosApi.eliminarMarca(m.id);
      setAviso(
        `Marca "${m.nombre}" eliminada` +
          (r.articulosDesvinculados ? ` (${r.articulosDesvinculados} artículos quedaron sin marca)` : "") +
          "."
      );
      cargar();
    } catch (e) {
      setError(e instanceof Error ? e.message : "Error al eliminar la marca.");
    }
  }

  // Unifica duplicadas y pasa todo a MAYÚSCULAS.
  async function normalizar() {
    if (!window.confirm("¿Pasar todas las marcas a MAYÚSCULAS y fusionar las duplicadas?")) {
      return;
    }
    setError(null);
    setAviso(null);
    try {
      const r = await catalogosApi.normalizarMarcas();
      setAviso(`Listo: ${r.fusionadas} marcas fusionadas, ${r.renombradas} pasadas a mayúsculas.`);
      cargar();
    } catch (e) {
      setError(e instanceof Error ? e.message : "Error al normalizar las marcas.");
    }
  }

  return (
    <div className="modal-fondo" onClick={onCerrar}>
      <div className="modal" style={{ maxWidth: 480 }} onClick={(e) => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
          <h2 style={{ margin: 0 }}>Marcas</h2>
          <button type="button" className="btn-secundario" style={{ padding: "6px 12px" }} onClick={normalizar}>
            Unificar duplicadas (MAYÚSCULAS)
          </button>
        </div>
        {error && <div className="alerta-error" style={{ marginTop: 12 }}>{error}</div>}
        {aviso && <div className="alerta-ok" style={{ marginTop: 12 }}>{aviso}</div>}
        {cargando ? (
          <p>Cargando…</p>
        ) : (
          <div style={{ maxHeight: 360, overflowY: "auto" }}>
            {marcas.map((m) => (
              <div
                key={m.id}
                style={{
                  display: "flex",
                  justifyContent: "space-between",
                  alignItems: "center",
                  padding: "8px 0",
                  borderBottom: "1px solid #f0eeea",
                }}
              >
                <span><strong>{m.nombre}</strong></span>
                <button
                  className="btn-secundario"
                  style={{ padding: "4px 10px", color: "var(--echo-coral)", borderColor: "var(--echo-coral)" }}
                  onClick={() => eliminar(m)}
                >
                  Eliminar
                </button>
              </div>
            ))}
            {marcas.length === 0 && <p className="muted">No hay marcas registradas.</p>}
          </div>
        )}
        <div className="modal-acciones">
          <button type="button" className="btn-secundario" onClick={onCerrar}>Cerrar</button>
        </div>
      </div>
    </div>
  );
}

// --------------------------------------------------------------------------
// Modal de creación / edición de artículo. Permite cambiar nombre, código,
// marca, precios, garantía, descripción e imagen (por URL o subida).
// --------------------------------------------------------------------------
function ModalArticulo({
  articulo,
  onCerrar,
  onGuardado,
}: {
  articulo: Articulo | null;
  onCerrar: () => void;
  onGuardado: () => void;
}) {
  const esEdicion = !!articulo;
  const [form, setForm] = useState<NuevoArticulo>({
    codigo: articulo?.codigo ?? "",
    nombre: articulo?.nombre ?? "",
    descripcion: articulo?.descripcion ?? "",
    costo: articulo?.costo ?? 0,
    precioMinimo: articulo?.precioMinimo ?? 0,
    tarifaNormal: articulo?.tarifaNormal ?? 0,
    garantiaMeses: articulo?.garantiaMeses ?? 6,
    urlFoto: articulo?.urlFoto ?? "",
    marca: articulo?.marca?.nombre ?? "",
  });
  const [error, setError] = useState<string | null>(null);
  const [guardando, setGuardando] = useState(false);
  const [subiendo, setSubiendo] = useState(false);
  const [marcas, setMarcas] = useState<Marca[]>([]);

  // Marcas existentes para sugerir en el campo (datalist). Se puede escribir una
  // nueva: el backend la crea/reutiliza por nombre, igual que en la importación.
  useEffect(() => {
    catalogosApi.marcas().then(setMarcas).catch(() => setMarcas([]));
  }, []);

  function set<K extends keyof NuevoArticulo>(campo: K, valor: NuevoArticulo[K]) {
    setForm((f) => ({ ...f, [campo]: valor }));
  }

  // Sube una imagen desde el computador y guarda la URL devuelta.
  async function onSubirArchivo(e: ChangeEvent<HTMLInputElement>) {
    const archivo = e.target.files?.[0];
    if (!archivo) return;
    setError(null);
    setSubiendo(true);
    try {
      const { url } = await uploadsApi.subirImagen(archivo);
      set("urlFoto", url);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Error al subir la imagen.");
    } finally {
      setSubiendo(false);
      e.target.value = "";
    }
  }

  async function onSubmit(e: FormEvent) {
    e.preventDefault();
    setError(null);
    setGuardando(true);
    try {
      if (articulo) await articulosApi.actualizar(articulo.id, form);
      else await articulosApi.crear(form);
      onGuardado();
    } catch (err) {
      setError(err instanceof Error ? err.message : "Error al guardar el artículo.");
    } finally {
      setGuardando(false);
    }
  }

  return (
    <div className="modal-fondo" onClick={onCerrar}>
      <div className="modal" style={{ maxWidth: 680 }} onClick={(e) => e.stopPropagation()}>
        <h2>{esEdicion ? "Editar artículo" : "+ Crear nuevo artículo"}</h2>
        <form onSubmit={onSubmit}>
          {error && <div className="alerta-error">{error}</div>}

          <div className="grid-2">
            <div className="campo">
              <label>Nombre *</label>
              <input value={form.nombre} onChange={(e) => set("nombre", e.target.value)} required />
            </div>
            <div className="campo">
              <label>Código de referencia *</label>
              <input value={form.codigo} onChange={(e) => set("codigo", e.target.value)} required />
            </div>
          </div>

          <div className="grid-2">
            <div className="campo">
              <label>Costo</label>
              <input
                type="number"
                min="0"
                value={form.costo}
                onChange={(e) => set("costo", Number(e.target.value))}
              />
            </div>
            <div className="campo">
              <label>Precio: Tarifa normal</label>
              <input
                type="number"
                min="0"
                value={form.tarifaNormal}
                onChange={(e) => set("tarifaNormal", Number(e.target.value))}
              />
            </div>
          </div>

          <div className="grid-2">
            <div className="campo">
              <label>Precio mínimo de venta</label>
              <input
                type="number"
                min="0"
                value={form.precioMinimo}
                onChange={(e) => set("precioMinimo", Number(e.target.value))}
              />
            </div>
            <div className="campo">
              <label>Garantía (meses)</label>
              <input
                type="number"
                min="0"
                step="1"
                value={form.garantiaMeses ?? 6}
                onChange={(e) => set("garantiaMeses", Number(e.target.value))}
              />
            </div>
          </div>

          <div className="grid-2">
            <div className="campo">
              <label>Marca</label>
              <input
                list="lista-marcas"
                value={form.marca ?? ""}
                onChange={(e) => set("marca", e.target.value)}
                placeholder="Escribe o elige una marca…"
              />
              <datalist id="lista-marcas">
                {marcas.map((m) => (
                  <option key={m.id} value={m.nombre} />
                ))}
              </datalist>
            </div>
            <div className="campo">
              <label>Descripción</label>
              <input value={form.descripcion ?? ""} onChange={(e) => set("descripcion", e.target.value)} />
            </div>
          </div>

          {/* Imagen del artículo: subir un archivo del PC o pegar una URL
              (las fotos importadas de Effi vienen como URLs de S3). */}
          <div className="campo">
            <label>Imagen</label>
            <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
              <input
                type="url"
                style={{ flex: 1, minWidth: 220 }}
                placeholder="https://…/foto.jpg  o sube un archivo →"
                value={form.urlFoto ?? ""}
                onChange={(e) => set("urlFoto", e.target.value)}
              />
              <label
                className="btn-secundario"
                style={{ cursor: subiendo ? "wait" : "pointer", opacity: subiendo ? 0.6 : 1, whiteSpace: "nowrap" }}
              >
                {subiendo ? "Subiendo…" : "Subir del PC"}
                <input type="file" accept="image/*" hidden disabled={subiendo} onChange={onSubirArchivo} />
              </label>
            </div>
          </div>
          {form.urlFoto ? (
            <div style={{ marginBottom: 12, display: "flex", alignItems: "flex-start", gap: 12 }}>
              <img
                src={form.urlFoto}
                alt="Vista previa"
                className="previa-articulo"
                onError={(e) => {
                  (e.currentTarget as HTMLImageElement).style.display = "none";
                }}
              />
              <button type="button" className="btn-secundario" style={{ padding: "6px 12px" }} onClick={() => set("urlFoto", "")}>
                Quitar imagen
              </button>
            </div>
          ) : null}

          <div className="modal-acciones">
            <button type="button" className="btn-secundario" onClick={onCerrar}>
              Cancelar
            </button>
            <button type="submit" className="btn-primario" style={{ width: "auto" }} disabled={guardando}>
              {guardando ? "Guardando…" : esEdicion ? "Guardar cambios" : "Crear artículo"}
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}
