// ==========================================================================
// Pantalla de Clientes.
//  - Listado paginado con búsqueda.
//  - Crear / editar cliente.
//  - Importar desde el Excel exportado de Effi (mapea encabezados aunque vengan
//    con acentos/codificación rara; envía por lotes para no saturar la API).
//  - Exportar a Excel y a CSV de Google Contacts.
// ==========================================================================
import { useEffect, useState, type ChangeEvent, type FormEvent } from "react";
import * as XLSX from "xlsx";
import { leerLibroExcel } from "../utils/excel";
import {
  clientesApi,
  empleadosApi,
  type Cliente,
  type NuevoCliente,
  type Empleado,
  type FilaImportCliente,
  type ResumenImportClientes,
} from "../api/recursos";

const TAM_PAGINA = 50;
const TAM_LOTE = 1000; // filas por petición al importar

// Normaliza un encabezado: minúsculas, sin acentos ni caracteres raros.
const norm = (s: unknown) =>
  String(s ?? "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "") // quita acentos
    .replace(/[^a-z0-9]/g, "");

// Mapea los encabezados del Excel a los campos del cliente. Devuelve, por cada
// campo, el índice de columna donde está. Tolera la codificación de Effi.
function mapearEncabezados(encabezados: string[]): Record<string, number> {
  const idx: Record<string, number> = {};
  const set = (campo: string, i: number) => {
    if (idx[campo] === undefined) idx[campo] = i;
  };
  encabezados.forEach((h, i) => {
    const n = norm(h);
    if (!n || n.includes("ideffi")) return; // ignorar IDs internos de Effi
    if (n.includes("tipodeidentificaci")) set("tipoIdentificacion", i);
    else if (n.includes("merodeidentificaci")) set("documento", i);
    else if (n === "nombre") set("nombre", i);
    else if (n.includes("fono1")) set("telefono1", i);
    else if (n.includes("fono2")) set("telefono2", i);
    else if (n === "celular") set("celular", i);
    else if (n === "whatsapp") set("whatsapp", i);
    else if (n === "email") set("email", i);
    else if (n === "pais" || n === "pas") set("pais", i);
    else if (n === "departamento") set("departamento", i);
    else if (n === "ciudad") set("ciudad", i);
    else if (n === "direccion" || n === "direccin") set("direccion", i);
    else if (n.includes("tipodepersona")) set("tipoPersona", i);
    else if (n.includes("tipodecliente")) set("tipoCliente", i);
    else if (n.includes("tarifa")) set("tarifaPrecios", i);
    else if (n.includes("formadepago")) set("formaPago", i);
    else if (n.includes("moneda")) set("moneda", i);
    else if (n === "vendedor") set("vendedor", i);
    else if (n.includes("observaci")) set("observacion", i);
    else if (n === "vigencia") set("vigencia", i);
  });
  return idx;
}

export function Clientes() {
  const [lista, setLista] = useState<Cliente[]>([]);
  const [total, setTotal] = useState(0);
  const [pagina, setPagina] = useState(1);
  const [buscar, setBuscar] = useState("");
  const [cargando, setCargando] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [aviso, setAviso] = useState<string | null>(null);
  const [editar, setEditar] = useState<Cliente | "nuevo" | null>(null);
  const [importando, setImportando] = useState(false);

  async function cargar(p = pagina, q = buscar) {
    setCargando(true);
    setError(null);
    try {
      const r = await clientesApi.listar(q, p, TAM_PAGINA);
      setLista(r.datos);
      setTotal(r.total);
      setPagina(r.pagina);
    } catch (e) {
      setError(e instanceof Error ? e.message : "Error al cargar clientes.");
    } finally {
      setCargando(false);
    }
  }

  useEffect(() => {
    cargar(1, "");
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const totalPaginas = Math.max(1, Math.ceil(total / TAM_PAGINA));

  // --- Eliminar cliente (con confirmación) ---
  async function eliminar(c: Cliente) {
    if (!window.confirm(`¿Eliminar al cliente "${c.nombre}"? Esta acción no se puede deshacer.`)) {
      return;
    }
    setError(null);
    setAviso(null);
    try {
      await clientesApi.eliminar(c.id);
      setAviso(`Cliente ${c.nombre} eliminado.`);
      cargar();
    } catch (e) {
      setError(e instanceof Error ? e.message : "Error al eliminar el cliente.");
    }
  }

  // --- Importación desde Excel de Effi ---
  async function onArchivoExcel(e: ChangeEvent<HTMLInputElement>) {
    const archivo = e.target.files?.[0];
    if (!archivo) return;
    setError(null);
    setAviso(null);
    setImportando(true);
    try {
      const libro = await leerLibroExcel(archivo);
      const hoja = libro.Sheets[libro.SheetNames[0]];
      // Leemos como matriz (array de arrays) para mapear por posición de columna.
      const matriz = XLSX.utils.sheet_to_json<(string | number)[]>(hoja, {
        header: 1,
        blankrows: false,
      });
      if (matriz.length < 2) throw new Error("El archivo no tiene filas de datos.");

      const encabezados = (matriz[0] as unknown[]).map((c) => String(c ?? ""));
      const idx = mapearEncabezados(encabezados);
      if (idx.documento === undefined || idx.nombre === undefined) {
        throw new Error(
          "No se reconocieron las columnas de documento y nombre. Verifica que sea el Excel de clientes de Effi."
        );
      }

      const campos = Object.keys(idx);
      const filas: FilaImportCliente[] = [];
      for (const row of matriz.slice(1)) {
        const fila: FilaImportCliente = {};
        for (const campo of campos) {
          const valor = (row as unknown[])[idx[campo]];
          if (valor !== undefined && valor !== null && String(valor).trim() !== "") {
            fila[campo] = String(valor).trim();
          }
        }
        if (fila.documento && fila.nombre) filas.push(fila);
      }
      if (filas.length === 0) throw new Error("No se encontraron filas con documento y nombre.");

      // Envío por lotes para no exceder límites de tamaño/peticiones.
      const acumulado: ResumenImportClientes = {
        recibidas: 0,
        validas: 0,
        creadas: 0,
        duplicadasOmitidas: 0,
        descartadas: 0,
      };
      const totalLotes = Math.ceil(filas.length / TAM_LOTE);
      for (let i = 0; i < filas.length; i += TAM_LOTE) {
        const lote = filas.slice(i, i + TAM_LOTE);
        setAviso(`Importando… lote ${Math.floor(i / TAM_LOTE) + 1} de ${totalLotes}`);
        const r = await clientesApi.importar(lote);
        acumulado.recibidas += r.recibidas;
        acumulado.validas += r.validas;
        acumulado.creadas += r.creadas;
        acumulado.duplicadasOmitidas += r.duplicadasOmitidas;
        acumulado.descartadas += r.descartadas;
      }
      setAviso(
        `Importación finalizada: ${acumulado.creadas} clientes creados` +
          (acumulado.duplicadasOmitidas ? `, ${acumulado.duplicadasOmitidas} ya existían` : "") +
          (acumulado.descartadas ? `, ${acumulado.descartadas} filas descartadas` : "") +
          "."
      );
      cargar(1, "");
    } catch (err) {
      setError(err instanceof Error ? err.message : "Error al importar el Excel.");
    } finally {
      setImportando(false);
      e.target.value = "";
    }
  }

  // --- Exportación a Excel ---
  async function exportarExcel() {
    setError(null);
    try {
      const todos = await clientesApi.exportar();
      const filas = todos.map((c) => ({
        "Tipo de identificación": c.tipoIdentificacion ?? "",
        "Número de identificación": c.documento,
        Nombre: c.nombre,
        "Teléfono 1": c.telefono1 ?? "",
        "Teléfono 2": c.telefono2 ?? "",
        Celular: c.celular ?? "",
        WhatsApp: c.whatsapp ?? "",
        Email: c.email ?? "",
        País: c.pais ?? "",
        Departamento: c.departamento ?? "",
        Ciudad: c.ciudad ?? "",
        Dirección: c.direccion ?? "",
        "Tipo de persona": c.tipoPersona ?? "",
        "Tipo de cliente": c.tipoCliente ?? "",
        "Tarifa de precios": c.tarifaPrecios ?? "",
        "Forma de pago": c.formaPago ?? "",
        Moneda: c.moneda ?? "",
        Vendedor: c.vendedor ?? "",
        Observación: c.observacion ?? "",
        Vigencia: c.activo ? "Vigente" : "Anulado",
      }));
      const hoja = XLSX.utils.json_to_sheet(filas);
      const libro = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(libro, hoja, "Clientes");
      XLSX.writeFile(libro, "clientes_echo.xlsx");
    } catch (err) {
      setError(err instanceof Error ? err.message : "Error al exportar.");
    }
  }

  // --- Exportación a CSV de Google Contacts ---
  async function exportarGoogle() {
    setError(null);
    try {
      const todos = await clientesApi.exportar();
      const cab = [
        "Name",
        "Given Name",
        "E-mail 1 - Value",
        "Phone 1 - Type",
        "Phone 1 - Value",
        "Phone 2 - Type",
        "Phone 2 - Value",
        "Address 1 - Formatted",
        "Notes",
      ];
      const esc = (v: string) => `"${(v ?? "").replace(/"/g, '""')}"`;
      const lineas = [cab.map(esc).join(",")];
      for (const c of todos) {
        lineas.push(
          [
            c.nombre,
            c.nombre,
            c.email ?? "",
            "Mobile",
            c.celular ?? c.whatsapp ?? "",
            "Home",
            c.telefono1 ?? c.telefono2 ?? "",
            [c.direccion, c.ciudad, c.departamento, c.pais].filter(Boolean).join(", "),
            c.observacion ?? "",
          ]
            .map((x) => esc(String(x ?? "")))
            .join(",")
        );
      }
      // BOM para que Google/Excel respeten los acentos.
      const blob = new Blob(["﻿" + lineas.join("\r\n")], { type: "text/csv;charset=utf-8;" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "clientes_google_contacts.csv";
      a.click();
      URL.revokeObjectURL(url);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Error al exportar a Google Contacts.");
    }
  }

  return (
    <div>
      <div className="dash-topbar">
        <h2>Clientes <span className="muted">· {total.toLocaleString("es-CO")} registros</span></h2>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
          <button className="btn-secundario" onClick={exportarExcel}>Exportar a Excel</button>
          <button className="btn-secundario" onClick={exportarGoogle}>Exportar a Google Contacts</button>
          <label className="btn-secundario" style={{ cursor: importando ? "wait" : "pointer", opacity: importando ? 0.6 : 1 }}>
            {importando ? "Importando…" : "Importar Excel (Effi)"}
            <input type="file" accept=".xls,.xlsx" hidden disabled={importando} onChange={onArchivoExcel} />
          </label>
          <button className="btn-primario" style={{ width: "auto" }} onClick={() => setEditar("nuevo")}>
            + Crear cliente
          </button>
        </div>
      </div>

      <div className="barra-busqueda">
        <input
          placeholder="Buscar por nombre, documento, email o celular…"
          value={buscar}
          onChange={(e) => setBuscar(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && cargar(1, buscar)}
        />
        <button className="btn-secundario" onClick={() => cargar(1, buscar)}>Buscar</button>
      </div>

      {error && <div className="alerta-error">{error}</div>}
      {aviso && <div className="alerta-ok">{aviso}</div>}

      {cargando ? (
        <p>Cargando…</p>
      ) : (
        <>
          <div className="tabla-wrap">
            <table className="tabla">
              <thead>
                <tr>
                  <th>Documento</th>
                  <th>Cliente</th>
                  <th>Teléfono</th>
                  <th>Email</th>
                  <th>Ciudad</th>
                  <th>Tarifa / pago</th>
                  <th>Vigencia</th>
                  <th></th>
                </tr>
              </thead>
              <tbody>
                {lista.map((c) => (
                  <tr key={c.id}>
                    <td>
                      {c.documento}
                      {c.tipoIdentificacion && <div className="muted">{c.tipoIdentificacion}</div>}
                    </td>
                    <td><strong>{c.nombre}</strong></td>
                    <td>
                      {c.celular || c.telefono1 || "—"}
                      {c.whatsapp && <div className="muted">WhatsApp: {c.whatsapp}</div>}
                    </td>
                    <td>{c.email ?? "—"}</td>
                    <td>{c.ciudad ?? "—"}</td>
                    <td>
                      {c.tarifaPrecios ?? "—"}
                      <div className="muted">{c.formaPago ?? ""}</div>
                    </td>
                    <td>
                      <span className="badge-sede" style={{ background: c.activo ? undefined : "#f0d0d0" }}>
                        {c.activo ? "Vigente" : "Anulado"}
                      </span>
                    </td>
                    <td>
                      <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                        <button className="btn-secundario" style={{ padding: "6px 12px" }} onClick={() => setEditar(c)}>
                          Editar
                        </button>
                        <button
                          className="btn-secundario"
                          style={{ padding: "6px 12px", color: "var(--echo-coral)", borderColor: "var(--echo-coral)" }}
                          onClick={() => eliminar(c)}
                        >
                          Eliminar
                        </button>
                      </div>
                    </td>
                  </tr>
                ))}
                {lista.length === 0 && (
                  <tr>
                    <td colSpan={8} className="muted" style={{ textAlign: "center", padding: 24 }}>
                      No hay clientes. Crea uno o importa el Excel de Effi.
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>

          {/* Paginación */}
          <div style={{ display: "flex", justifyContent: "center", alignItems: "center", gap: 14, marginTop: 16 }}>
            <button className="btn-secundario" disabled={pagina <= 1} onClick={() => cargar(pagina - 1)}>
              ← Anterior
            </button>
            <span className="muted">Página {pagina} de {totalPaginas}</span>
            <button className="btn-secundario" disabled={pagina >= totalPaginas} onClick={() => cargar(pagina + 1)}>
              Siguiente →
            </button>
          </div>
        </>
      )}

      {editar && (
        <ModalCliente
          cliente={editar === "nuevo" ? null : editar}
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

// --------------------------------------------------------------------------
// Modal de creación / edición de cliente (campos conservados del formulario
// de Effi; se omiten los marcados en rojo).
// --------------------------------------------------------------------------
const TIPOS_ID = ["Cédula de ciudadanía", "NIT", "Cédula de extranjería", "Pasaporte", "Otro"];

export function ModalCliente({
  cliente,
  onCerrar,
  onGuardado,
  prefill,
}: {
  cliente: Cliente | null;
  onCerrar: () => void;
  // Devuelve el cliente guardado (creado o editado) para que el llamador lo use.
  onGuardado: (guardado: Cliente) => void;
  // Valores iniciales para una alta rápida (p. ej. desde la remisión).
  prefill?: { documento?: string; nombre?: string };
}) {
  const [form, setForm] = useState<NuevoCliente>({
    documento: cliente?.documento ?? prefill?.documento ?? "",
    tipoIdentificacion: cliente?.tipoIdentificacion ?? "Cédula de ciudadanía",
    nombre: cliente?.nombre ?? prefill?.nombre ?? "",
    email: cliente?.email ?? "",
    telefono1: cliente?.telefono1 ?? "",
    telefono2: cliente?.telefono2 ?? "",
    celular: cliente?.celular ?? "",
    whatsapp: cliente?.whatsapp ?? "",
    pais: cliente?.pais ?? "Colombia",
    departamento: cliente?.departamento ?? "",
    ciudad: cliente?.ciudad ?? "",
    direccion: cliente?.direccion ?? "",
    tipoPersona: cliente?.tipoPersona ?? "Física (natural)",
    tipoCliente: cliente?.tipoCliente ?? "Común",
    tarifaPrecios: cliente?.tarifaPrecios ?? "Tarifa normal",
    formaPago: cliente?.formaPago ?? "CONTADO",
    moneda: cliente?.moneda ?? "COP",
    vendedor: cliente?.vendedor ?? "",
    observacion: cliente?.observacion ?? "",
  });
  const [error, setError] = useState<string | null>(null);
  const [guardando, setGuardando] = useState(false);
  // Vendedores disponibles (Empleados) para el desplegable, evita errores de tipeo.
  const [empleados, setEmpleados] = useState<Empleado[]>([]);

  useEffect(() => {
    empleadosApi.listar().then(setEmpleados).catch(() => setEmpleados([]));
  }, []);

  function set<K extends keyof NuevoCliente>(campo: K, valor: NuevoCliente[K]) {
    setForm((f) => ({ ...f, [campo]: valor }));
  }

  async function onSubmit(e: FormEvent) {
    e.preventDefault();
    setError(null);
    setGuardando(true);
    try {
      const guardado = cliente
        ? await clientesApi.actualizar(cliente.id, form)
        : await clientesApi.crear(form);
      onGuardado(guardado);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Error al guardar el cliente.");
    } finally {
      setGuardando(false);
    }
  }

  return (
    <div className="modal-fondo">
      <div className="modal" style={{ maxWidth: 720 }} onClick={(e) => e.stopPropagation()}>
        <h2>{cliente ? "Editar cliente" : "+ Crear cliente"}</h2>
        <form onSubmit={onSubmit}>
          {error && <div className="alerta-error">{error}</div>}

          <div className="grid-2">
            <div className="campo">
              <label>Tipo de identificación</label>
              <select value={form.tipoIdentificacion ?? ""} onChange={(e) => set("tipoIdentificacion", e.target.value)}>
                {TIPOS_ID.map((t) => <option key={t} value={t}>{t}</option>)}
              </select>
            </div>
            <div className="campo">
              <label>Número de identificación *</label>
              <input value={form.documento} onChange={(e) => set("documento", e.target.value)} required />
            </div>
          </div>

          <div className="campo">
            <label>Nombre / Razón social *</label>
            <input value={form.nombre} onChange={(e) => set("nombre", e.target.value)} required />
          </div>

          <div className="grid-2">
            <div className="campo">
              <label>Email</label>
              <input type="email" value={form.email ?? ""} onChange={(e) => set("email", e.target.value)} />
            </div>
            <div className="campo">
              <label>Celular</label>
              <input value={form.celular ?? ""} onChange={(e) => set("celular", e.target.value)} />
            </div>
          </div>

          <div className="grid-2">
            <div className="campo">
              <label>WhatsApp</label>
              <input value={form.whatsapp ?? ""} onChange={(e) => set("whatsapp", e.target.value)} />
            </div>
            <div className="campo">
              <label>Teléfono 1</label>
              <input value={form.telefono1 ?? ""} onChange={(e) => set("telefono1", e.target.value)} />
            </div>
          </div>

          <div className="grid-2">
            <div className="campo">
              <label>Departamento</label>
              <input value={form.departamento ?? ""} onChange={(e) => set("departamento", e.target.value)} />
            </div>
            <div className="campo">
              <label>Ciudad</label>
              <input value={form.ciudad ?? ""} onChange={(e) => set("ciudad", e.target.value)} />
            </div>
          </div>

          <div className="campo">
            <label>Dirección</label>
            <input value={form.direccion ?? ""} onChange={(e) => set("direccion", e.target.value)} />
          </div>

          <div className="grid-2">
            <div className="campo">
              <label>Tipo de cliente</label>
              <input value={form.tipoCliente ?? ""} onChange={(e) => set("tipoCliente", e.target.value)} />
            </div>
            <div className="campo">
              <label>Tarifa de precios</label>
              <input value={form.tarifaPrecios ?? ""} onChange={(e) => set("tarifaPrecios", e.target.value)} />
            </div>
          </div>

          <div className="campo">
            <label>Vendedor</label>
            <select value={form.vendedor ?? ""} onChange={(e) => set("vendedor", e.target.value)}>
              <option value="">— Sin asignar —</option>
              {/* Si el valor actual no está en Empleados, lo conservamos como opción. */}
              {form.vendedor && !empleados.some((em) => em.nombre === form.vendedor) && (
                <option value={form.vendedor}>{form.vendedor} (actual)</option>
              )}
              {empleados.map((em) => (
                <option key={em.id} value={em.nombre}>{em.nombre}</option>
              ))}
            </select>
            {empleados.length === 0 && (
              <small className="muted">No hay empleados aún. Créalos en la pantalla Empleados.</small>
            )}
          </div>

          <p className="muted">Moneda: COP · Forma de pago: Contado</p>

          <div className="modal-acciones">
            <button type="button" className="btn-secundario" onClick={onCerrar}>Cancelar</button>
            <button type="submit" className="btn-primario" style={{ width: "auto" }} disabled={guardando}>
              {guardando ? "Guardando…" : cliente ? "Guardar cambios" : "Crear cliente"}
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}
