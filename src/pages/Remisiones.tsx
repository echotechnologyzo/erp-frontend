// ==========================================================================
// Pantalla de Remisiones de VENTA.
//  - Lista las remisiones emitidas (con su comisión).
//  - Permite crear una nueva (descuenta stock y calcula comisión 4%/1%).
//  - Permite importar masivamente desde Excel (Google Sheets / WhatsApp).
//  - Genera el PDF imprimible replicando el formato MED-9480 (window.print()).
// ==========================================================================
import { useEffect, useRef, useState, type ChangeEvent, type FormEvent } from "react";
import * as XLSX from "xlsx";
import { useAuth } from "../auth/AuthContext";
import {
  remisionesApi,
  clientesApi,
  catalogosApi,
  articulosApi,
  empleadosApi,
  type Remision,
  type RemisionCompleta,
  type Sede,
  type Articulo,
  type Cliente,
  type Empleado,
  type NuevaRemisionItem,
  type FilaImportRemision,
} from "../api/recursos";
import { ModalCliente } from "./Clientes";
import { SelectorArticulo } from "../components/SelectorArticulo";

// Formato de dinero igual al del PDF actual: "$205,000" (coma de miles).
const pesos = (v: number) => "$" + Math.round(Number(v)).toLocaleString("en-US");

export function Remisiones() {
  const [lista, setLista] = useState<Remision[]>([]);
  const [buscar, setBuscar] = useState("");
  const [cargando, setCargando] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [aviso, setAviso] = useState<string | null>(null);
  const [modal, setModal] = useState(false);
  const [duplicarBase, setDuplicarBase] = useState<RemisionCompleta | null>(null);
  const [imprimir, setImprimir] = useState<RemisionCompleta | null>(null);
  const [modalConsec, setModalConsec] = useState(false);
  const { usuario } = useAuth();
  const esAdmin = usuario?.rol === "ADMIN";

  async function cargar(q = buscar) {
    setCargando(true);
    setError(null);
    try {
      setLista(await remisionesApi.listar({ buscar: q }));
    } catch (e) {
      setError(e instanceof Error ? e.message : "Error al cargar remisiones.");
    } finally {
      setCargando(false);
    }
  }

  useEffect(() => {
    cargar("");
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  async function verPdf(id: string) {
    setError(null);
    try {
      setImprimir(await remisionesApi.obtener(id));
    } catch (e) {
      setError(e instanceof Error ? e.message : "Error al abrir la remisión.");
    }
  }

  // Duplicar: trae los datos de la remisión y abre el formulario precargado.
  async function duplicar(r: Remision) {
    setError(null);
    try {
      const completa = await remisionesApi.obtener(r.id);
      setDuplicarBase(completa);
    } catch (e) {
      setError(e instanceof Error ? e.message : "Error al duplicar la remisión.");
    }
  }

  // Anular una remisión: devuelve el stock al inventario y la marca ANULADA.
  async function anular(r: Remision) {
    if (!window.confirm(`¿Anular la remisión ${r.documento}? Se devolverá el stock al inventario.`)) {
      return;
    }
    setError(null);
    setAviso(null);
    try {
      await remisionesApi.anular(r.id);
      setAviso(`Remisión ${r.documento} anulada (stock devuelto al inventario).`);
      cargar();
    } catch (e) {
      setError(e instanceof Error ? e.message : "Error al anular la remisión.");
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
      const filas = XLSX.utils.sheet_to_json<FilaImportRemision>(hoja);
      if (filas.length === 0) throw new Error("El archivo no tiene filas.");
      const r = await remisionesApi.importar(filas);
      setAviso(`Importación exitosa: ${r.remisionesCreadas} remisiones (${r.filasProcesadas} filas).`);
      cargar();
    } catch (err) {
      setError(err instanceof Error ? err.message : "Error al importar el Excel.");
    } finally {
      e.target.value = "";
    }
  }

  // Plantilla de Excel con los encabezados esperados (alineada al grupo de
  // WhatsApp → Google Sheets: una fila por artículo).
  function descargarPlantilla() {
    const ejemplo: FilaImportRemision[] = [
      {
        sede: "Medellín",
        clienteDocumento: "3113816369",
        clienteNombre: "Alexis Mejía",
        clienteTelefono: "3113816369",
        clienteDireccion: "Medellín",
        vendedor: "Jonathan Velásquez",
        remision: "WHATSAPP-001",
        fecha: "2026-05-22",
        medioPago: "Transferencia bancaria - BANCOLOMBIA Ahorros #69385143982",
        observacion: "",
        articuloCodigo: "000006",
        descripcion: "AMAZON ECHO DOT 5TH GEN NEGRO",
        cantidad: 2,
        precioUnitario: 205000,
        descuento: 0,
      },
    ];
    const hoja = XLSX.utils.json_to_sheet(ejemplo);
    const libro = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(libro, hoja, "Remisiones");
    XLSX.writeFile(libro, "plantilla_remisiones_venta.xlsx");
  }

  return (
    <div>
      <div className="dash-topbar">
        <h2>Remisiones de venta</h2>
        <div style={{ display: "flex", gap: 10 }}>
          <button className="btn-secundario" onClick={descargarPlantilla}>
            Descargar plantilla Excel
          </button>
          <label className="btn-secundario" style={{ cursor: "pointer" }}>
            Importar Excel
            <input type="file" accept=".xlsx,.xls" hidden onChange={onArchivoExcel} />
          </label>
          {esAdmin && (
            <button className="btn-secundario" onClick={() => setModalConsec(true)}>
              Consecutivos
            </button>
          )}
          <button className="btn-primario" style={{ width: "auto" }} onClick={() => setModal(true)}>
            + Crear remisión
          </button>
        </div>
      </div>

      <div className="barra-busqueda">
        <input
          placeholder="Buscar por cliente o documento…"
          value={buscar}
          onChange={(e) => setBuscar(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && cargar(buscar)}
        />
        <button className="btn-secundario" onClick={() => cargar(buscar)}>Buscar</button>
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
                <th>No.</th>
                <th>Sede</th>
                <th>Cliente</th>
                <th>Vendedor</th>
                <th>Total</th>
                {esAdmin && <th>Comisión</th>}
                <th></th>
              </tr>
            </thead>
            <tbody>
              {lista.map((r) => (
                <tr key={r.id}>
                  <td>{new Date(r.fecha).toLocaleDateString("es-CO")}</td>
                  <td><strong>{r.documento}</strong></td>
                  <td>{r.sede}</td>
                  <td>
                    {r.cliente}
                    <div className="muted">{r.clienteDocumento}</div>
                  </td>
                  <td>{r.vendedor}</td>
                  <td>{pesos(r.total)}</td>
                  {esAdmin && (
                    <td>
                      <span className={r.esRecompra ? "badge-sede" : "badge-nuevo"}>
                        {r.esRecompra ? "Recompra" : "Nuevo"} {Number(r.comisionPct)}%
                      </span>
                      <div className="muted">{pesos(r.comisionValor)}</div>
                    </td>
                  )}
                  <td>
                    <div style={{ display: "flex", gap: 6, flexWrap: "wrap", alignItems: "center" }}>
                      <button className="btn-secundario" style={{ padding: "6px 12px" }} onClick={() => verPdf(r.id)}>
                        Ver / PDF
                      </button>
                      <button className="btn-secundario" style={{ padding: "6px 12px" }} onClick={() => duplicar(r)}>
                        Duplicar
                      </button>
                      {r.estado === "ANULADA" ? (
                        <span className="badge-sede">Anulada</span>
                      ) : (
                        <button
                          className="btn-secundario"
                          style={{ padding: "6px 12px", color: "var(--echo-coral)", borderColor: "var(--echo-coral)" }}
                          onClick={() => anular(r)}
                        >
                          Anular
                        </button>
                      )}
                    </div>
                  </td>
                </tr>
              ))}
              {lista.length === 0 && (
                <tr>
                  <td colSpan={esAdmin ? 8 : 7} className="muted" style={{ textAlign: "center", padding: 24 }}>
                    No hay remisiones. Crea una o importa desde Excel.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      )}

      {(modal || duplicarBase) && (
        <ModalCrearRemision
          base={duplicarBase}
          onCerrar={() => { setModal(false); setDuplicarBase(null); }}
          onCreado={(id) => {
            setModal(false);
            setDuplicarBase(null);
            cargar();
            if (id) verPdf(id); // abre el PDF inmediatamente para imprimir/guardar
          }}
        />
      )}

      {imprimir && <RemisionImprimible remision={imprimir} onCerrar={() => setImprimir(null)} />}

      {modalConsec && <ModalConsecutivos onCerrar={() => setModalConsec(false)} />}
    </div>
  );
}

// --------------------------------------------------------------------------
// Modal de creación de remisión de venta.
// --------------------------------------------------------------------------
function ModalCrearRemision({
  base,
  onCerrar,
  onCreado,
}: {
  base?: RemisionCompleta | null; // si viene, se DUPLICA esa remisión
  onCerrar: () => void;
  onCreado: (id?: string) => void;
}) {
  const { usuario } = useAuth(); // para prellenar la sede asignada al usuario
  const esAdmin = usuario?.rol === "ADMIN"; // el operador no ve la comisión
  const [sedes, setSedes] = useState<Sede[]>([]);
  const [articulos, setArticulos] = useState<Articulo[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [guardando, setGuardando] = useState(false);

  const [sedeId, setSedeId] = useState(base?.sedeId ?? "");
  const [vendedor, setVendedor] = useState(base?.vendedor ?? "");
  const [medioPago, setMedioPago] = useState(base?.medioPago ?? "");
  const [observacion, setObservacion] = useState(base?.observacion ?? "");
  const [empleados, setEmpleados] = useState<Empleado[]>([]);

  // Buscador de cliente. Si duplicamos, precargamos el cliente de la base.
  const [clienteBuscar, setClienteBuscar] = useState("");
  const [clienteOpc, setClienteOpc] = useState<Cliente[]>([]);
  const [cliente, setCliente] = useState<Cliente | null>(
    base
      ? ({ id: base.clienteId, nombre: base.cliente.nombre, documento: base.cliente.documento } as Cliente)
      : null
  );
  const [crearCliente, setCrearCliente] = useState(false); // abre el modal rápido de alta

  const [items, setItems] = useState<NuevaRemisionItem[]>(
    base && base.detalles.length
      ? base.detalles.map((d) => ({
          articuloId: d.articuloId,
          cantidad: Number(d.cantidad),
          precioUnitario: Number(d.precioUnitario),
          descuento: Number(d.descuento),
          garantiaMeses: d.garantiaMeses,
        }))
      : [{ articuloId: "", cantidad: 1, precioUnitario: 0 }]
  );

  useEffect(() => {
    catalogosApi.sedes().then((s) => {
      setSedes(s);
      // Al duplicar, conserva la sede de la base; si no, la del usuario / la primera.
      if (base?.sedeId) {
        setSedeId(base.sedeId);
        return;
      }
      const propia = usuario?.sedeId && s.some((x) => x.id === usuario.sedeId)
        ? usuario.sedeId
        : s[0]?.id;
      if (propia) setSedeId(propia);
    });
    articulosApi.listar().then(setArticulos);
    empleadosApi.listar().then(setEmpleados).catch(() => setEmpleados([]));
  }, []);

  // Buscar clientes (con debounce sencillo).
  useEffect(() => {
    if (cliente) return; // ya hay uno elegido
    const t = setTimeout(async () => {
      if (clienteBuscar.trim().length < 2) {
        setClienteOpc([]);
        return;
      }
      const r = await clientesApi.listar(clienteBuscar, 1, 8);
      setClienteOpc(r.datos);
    }, 300);
    return () => clearTimeout(t);
  }, [clienteBuscar, cliente]);

  function actualizarItem(idx: number, campo: keyof NuevaRemisionItem, valor: string | number) {
    setItems((arr) => arr.map((it, i) => (i === idx ? { ...it, [campo]: valor } : it)));
  }
  // Al elegir un artículo, precargamos su tarifa normal como precio sugerido.
  function elegirArticulo(idx: number, articuloId: string) {
    const art = articulos.find((a) => a.id === articuloId);
    setItems((arr) =>
      arr.map((it, i) =>
        i === idx ? { ...it, articuloId, precioUnitario: art ? Number(art.tarifaNormal) : 0 } : it
      )
    );
  }
  const agregarItem = () => setItems((a) => [...a, { articuloId: "", cantidad: 1, precioUnitario: 0 }]);
  const quitarItem = (idx: number) => setItems((a) => a.filter((_, i) => i !== idx));

  const total = items.reduce(
    (acc, it) => acc + it.cantidad * (it.precioUnitario ?? 0) - (it.descuento ?? 0),
    0
  );

  async function onSubmit(e: FormEvent) {
    e.preventDefault();
    setError(null);
    if (!cliente) {
      setError("Selecciona un cliente.");
      return;
    }
    if (items.some((it) => !it.articuloId)) {
      setError("Selecciona el artículo en todas las líneas.");
      return;
    }
    setGuardando(true);
    try {
      const creada = await remisionesApi.crear({
        sedeId,
        clienteId: cliente.id,
        vendedor: vendedor || undefined,
        medioPago: medioPago || undefined,
        observacion: observacion || undefined,
        items,
      });
      onCreado(creada.id);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Error al crear la remisión.");
    } finally {
      setGuardando(false);
    }
  }

  return (
    <>
    <div className="modal-fondo">
      <div className="modal" style={{ maxWidth: 820 }} onClick={(e) => e.stopPropagation()}>
        <h2>{base ? "+ Duplicar remisión de venta" : "+ Crear remisión de venta"}</h2>
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
              <label>Vendedor</label>
              <select value={vendedor} onChange={(e) => setVendedor(e.target.value)}>
                <option value="">— Selecciona —</option>
                {empleados.map((em) => (
                  <option key={em.id} value={em.nombre}>{em.nombre}</option>
                ))}
              </select>
            </div>
          </div>

          {/* Selección de cliente */}
          <div className="campo">
            <label>Cliente *</label>
            {cliente ? (
              <div className="cliente-elegido">
                <span><strong>{cliente.nombre}</strong> · {cliente.documento}</span>
                <button type="button" className="btn-quitar" onClick={() => { setCliente(null); setClienteBuscar(""); }}>×</button>
              </div>
            ) : (
              <>
                <input
                  value={clienteBuscar}
                  onChange={(e) => setClienteBuscar(e.target.value)}
                  placeholder="Buscar cliente por nombre o documento…"
                />
                {clienteOpc.length > 0 && (
                  <div className="lista-opciones">
                    {clienteOpc.map((c) => (
                      <button type="button" key={c.id} className="opcion" onClick={() => { setCliente(c); setClienteOpc([]); }}>
                        <strong>{c.nombre}</strong> · {c.documento}
                      </button>
                    ))}
                  </div>
                )}
                {/* Si no está creado, se puede dar de alta sin salir de la remisión */}
                <div style={{ marginTop: 8 }}>
                  {clienteBuscar.trim().length >= 2 && clienteOpc.length === 0 && (
                    <span className="muted" style={{ marginRight: 10 }}>
                      Sin coincidencias para “{clienteBuscar}”.
                    </span>
                  )}
                  <button
                    type="button"
                    className="btn-secundario"
                    style={{ padding: "6px 12px" }}
                    onClick={() => setCrearCliente(true)}
                  >
                    + Crear cliente nuevo
                  </button>
                </div>
              </>
            )}
          </div>

          <h3 style={{ margin: "10px 0" }}>Artículos</h3>
          {items.map((it, idx) => {
            const art = articulos.find((a) => a.id === it.articuloId);
            return (
              <div key={idx} className="linea-item">
                <SelectorArticulo
                  articulos={articulos}
                  valor={it.articuloId}
                  onElegir={(id) => elegirArticulo(idx, id)}
                />
                <input
                  type="number" min="1" step="1" title="Cantidad" style={{ width: 70 }}
                  value={it.cantidad}
                  onChange={(e) => actualizarItem(idx, "cantidad", Number(e.target.value))}
                />
                <input
                  type="number" min="0" title="Precio unitario" style={{ width: 120 }}
                  value={it.precioUnitario ?? 0}
                  onChange={(e) => actualizarItem(idx, "precioUnitario", Number(e.target.value))}
                />
                <input
                  type="number" min="0" step="1" title="Garantía (meses)" style={{ width: 70 }}
                  placeholder={art ? String(art.garantiaMeses) : "6"}
                  value={it.garantiaMeses ?? ""}
                  onChange={(e) => actualizarItem(idx, "garantiaMeses", Number(e.target.value))}
                />
                <span className="muted" style={{ minWidth: 100, textAlign: "right" }}>
                  {pesos(it.cantidad * (it.precioUnitario ?? 0))}
                </span>
                {items.length > 1 && (
                  <button type="button" className="btn-quitar" onClick={() => quitarItem(idx)}>×</button>
                )}
              </div>
            );
          })}
          <button type="button" className="btn-secundario" onClick={agregarItem}>+ Agregar artículo</button>
          <p className="muted" style={{ marginTop: 4 }}>
            Garantía: por defecto la del artículo (déjala vacía para usarla).
          </p>

          <div className="grid-2">
            <div className="campo">
              <label>Medio de pago</label>
              <input value={medioPago} onChange={(e) => setMedioPago(e.target.value)} placeholder="Transferencia bancaria - BANCOLOMBIA Ahorros #…" />
            </div>
            <div className="campo">
              <label>Observación</label>
              <input value={observacion} onChange={(e) => setObservacion(e.target.value)} />
            </div>
          </div>

          <div className="total-neto">TOTAL NETO: <strong>{pesos(total)}</strong></div>
          <p className="muted">
            Moneda: COP · Forma de pago: Contado
            {esAdmin ? " · Comisión: 4% cliente nuevo / 1% recompra (automática)." : "."}
          </p>

          <div className="modal-acciones">
            <button type="button" className="btn-secundario" onClick={onCerrar}>Cancelar</button>
            <button type="submit" className="btn-primario" style={{ width: "auto" }} disabled={guardando}>
              {guardando ? "Guardando…" : "Emitir remisión"}
            </button>
          </div>
        </form>
      </div>
    </div>

    {/* Alta de cliente con el formulario COMPLETO (el mismo de la pantalla
        Clientes), sin salir de la remisión. Así no hay que editarlo después. */}
    {crearCliente && (
      <ModalCliente
        cliente={null}
        prefill={
          /^\d+$/.test(clienteBuscar.trim())
            ? { documento: clienteBuscar.trim() }
            : { nombre: clienteBuscar.trim() }
        }
        onCerrar={() => setCrearCliente(false)}
        onGuardado={(creado) => {
          setCliente(creado);
          setClienteOpc([]);
          setClienteBuscar("");
          setCrearCliente(false);
        }}
      />
    )}
    </>
  );
}

// --------------------------------------------------------------------------
// Modal (ADMIN) para fijar el PRÓXIMO consecutivo de remisión por sede.
// Permite continuar la numeración de Effi (p. ej. Medellín = 9644, Bogotá = 3081).
// --------------------------------------------------------------------------
function ModalConsecutivos({ onCerrar }: { onCerrar: () => void }) {
  const [sedes, setSedes] = useState<Sede[]>([]);
  const [valores, setValores] = useState<Record<string, string>>({});
  const [cargando, setCargando] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [aviso, setAviso] = useState<string | null>(null);
  const [guardando, setGuardando] = useState<string | null>(null);

  useEffect(() => {
    catalogosApi
      .sedes()
      .then((s) => {
        setSedes(s);
        const v: Record<string, string> = {};
        s.forEach((se) => { v[se.id] = String(se.consecutivoRemision ?? 1); });
        setValores(v);
      })
      .catch((e) => setError(e instanceof Error ? e.message : "Error al cargar sedes."))
      .finally(() => setCargando(false));
  }, []);

  async function guardar(s: Sede) {
    const n = Number(valores[s.id]);
    if (!Number.isInteger(n) || n < 1) {
      setError("El número debe ser un entero mayor o igual a 1.");
      return;
    }
    setError(null);
    setAviso(null);
    setGuardando(s.id);
    try {
      await catalogosApi.actualizarConsecutivo(s.id, n);
      setAviso(`La próxima remisión de ${s.nombre} será ${s.nombre.substring(0, 3).toUpperCase()}-${n}.`);
    } catch (e) {
      setError(e instanceof Error ? e.message : "Error al guardar el consecutivo.");
    } finally {
      setGuardando(null);
    }
  }

  return (
    <div className="modal-fondo">
      <div className="modal" style={{ maxWidth: 520 }} onClick={(e) => e.stopPropagation()}>
        <h2>Consecutivos de remisión por sede</h2>
        <p className="muted">
          Fija el <strong>próximo número</strong> que se asignará en cada sede (para continuar la
          numeración de Effi). El documento se forma como PREFIJO-número, p. ej. MED-9644.
        </p>
        {error && <div className="alerta-error">{error}</div>}
        {aviso && <div className="alerta-ok">{aviso}</div>}
        {cargando ? (
          <p>Cargando…</p>
        ) : (
          sedes.map((s) => (
            <div key={s.id} className="grid-2" style={{ alignItems: "end", marginBottom: 10 }}>
              <div className="campo">
                <label>{s.nombre} · próximo {s.nombre.substring(0, 3).toUpperCase()}-</label>
                <input
                  type="number"
                  min="1"
                  step="1"
                  value={valores[s.id] ?? ""}
                  onChange={(e) => setValores((v) => ({ ...v, [s.id]: e.target.value }))}
                />
              </div>
              <button
                type="button"
                className="btn-secundario"
                style={{ height: 44 }}
                disabled={guardando === s.id}
                onClick={() => guardar(s)}
              >
                {guardando === s.id ? "Guardando…" : "Guardar"}
              </button>
            </div>
          ))
        )}
        <div className="modal-acciones">
          <button type="button" className="btn-secundario" onClick={onCerrar}>Cerrar</button>
        </div>
      </div>
    </div>
  );
}

// --------------------------------------------------------------------------
// Vista IMPRIMIBLE: replica el formato del PDF MED-9480.
// Se muestra como overlay; "Imprimir / Guardar PDF" llama a window.print().
// El CSS @media print (theme.css) oculta todo menos .remision-print.
// --------------------------------------------------------------------------
// Datos fijos del emisor (Manual de marca / encabezado del PDF actual).
const EMISOR = {
  nombre: "ECHO TECNOLOGÍA",
  representante: "YESICA ZULUAGA OSPINA",
  nit: "1017175943",
  regimen: "Régimen ordinario No responsable de IVA",
  telefonos: "3207548718",
  email: "echotecnologia@echotecnologia.co",
  web: "echotecnologia.co",
};

function RemisionImprimible({ remision: r, onCerrar }: { remision: RemisionCompleta; onCerrar: () => void }) {
  const ref = useRef<HTMLDivElement>(null);
  const fechaTxt = (s: string | null) => (s ? new Date(s).toLocaleString("es-CO") : "—");
  const totalUnidades = r.detalles.reduce((a, d) => a + Number(d.cantidad), 0);
  // La garantía solo aplica a productos (los valores a terceros, p. ej. flete,
  // no llevan garantía). Si todas comparten meses, mostramos un texto único.
  const conGarantia = r.detalles.filter((d) => !d.esTercero);
  const garantias = [...new Set(conGarantia.map((d) => d.garantiaMeses))];
  const dirEmisor = r.sede.direccion ?? "Antioquia / Medellín / Carrera 55 #12Sur 09 Torre 3 Apto 9920";

  // Imprime/guarda como PDF; el nombre del archivo sale del título del documento,
  // así que lo fijamos a "Remisión <No.>" y lo restauramos al terminar.
  function imprimir() {
    const tituloPrevio = document.title;
    document.title = `Remisión ${r.documento}`;
    const restaurar = () => {
      document.title = tituloPrevio;
      window.removeEventListener("afterprint", restaurar);
    };
    window.addEventListener("afterprint", restaurar);
    window.print();
  }

  return (
    // OJO: NO poner "no-print" en este contenedor: al imprimir, display:none en
    // un ancestro oculta también el documento (.remision-print) y el PDF sale en
    // blanco. La impresión se controla con visibility en @media print; aquí solo
    // marcamos como no-print la barra de botones.
    <div className="modal-fondo">
      <div className="modal" style={{ maxWidth: 900, maxHeight: "92vh", overflow: "auto" }} onClick={(e) => e.stopPropagation()}>
        <div className="modal-acciones no-print" style={{ justifyContent: "space-between", marginTop: 0, marginBottom: 12 }}>
          <h2 style={{ margin: 0 }}>Remisión {r.documento}</h2>
          <div style={{ display: "flex", gap: 10 }}>
            <button className="btn-secundario" onClick={onCerrar}>Cerrar</button>
            <button className="btn-primario" style={{ width: "auto" }} onClick={imprimir}>
              Imprimir / Guardar PDF
            </button>
          </div>
        </div>

        {/* Documento imprimible */}
        <div className="remision-print" ref={ref}>
          {r.estado === "ANULADA" && <div className="rp-marca-anulada">ANULADA</div>}
          <div className="rp-cab">
            <img
              src="/logo-echo.png"
              alt="ECHO Tecnología en Casa"
              className="rp-logo-img"
              onError={(e) => { (e.currentTarget as HTMLImageElement).style.display = "none"; }}
            />

            <div className="rp-emisor">
              <h1>{EMISOR.nombre}</h1>
              <div className="rp-rep">{EMISOR.representante}</div>
              <div><strong>NIT:</strong> {EMISOR.nit} | {EMISOR.regimen}</div>
              <div>{dirEmisor}</div>
              <div><strong>Teléfonos:</strong> {r.sede.telefono ?? EMISOR.telefonos}</div>
              <div><strong>Email:</strong> {EMISOR.email}</div>
              <div><strong>Página web:</strong> {EMISOR.web}</div>
            </div>
            <div className="rp-doc">
              <div className="rp-titulo">REMISIÓN DE VENTA</div>
              <div className="rp-num">No. {r.documento}</div>
              <div><strong>Creación:</strong> {fechaTxt(r.fecha)}</div>
              <div><strong>Vencimiento:</strong> {fechaTxt(r.vencimiento ?? r.fecha)}</div>
              <div><strong>Página:</strong> 1 de 1</div>
            </div>
          </div>

          {/* Datos del cliente */}
          <table className="rp-cliente">
            <tbody>
              <tr>
                <td><strong>Cliente:</strong> {r.cliente.nombre}</td>
                <td><strong>{r.cliente.tipoIdentificacion ?? "CC"}:</strong> {r.cliente.documento}</td>
              </tr>
              <tr><td colSpan={2}><strong>Teléfono:</strong> {r.cliente.telefono || "—"}</td></tr>
              <tr>
                <td colSpan={2}>
                  <strong>Dirección:</strong> {r.cliente.direccion ?? "."} / {r.cliente.ciudad ?? ""} / {r.cliente.departamento ?? ""} / {r.cliente.pais ?? "Colombia"}
                </td>
              </tr>
              <tr><td colSpan={2}><strong>Elaboró:</strong> {r.elaboro} ({r.sede.nombre})</td></tr>
              <tr><td colSpan={2}><strong>Vendedor:</strong> {r.vendedor ?? "—"}</td></tr>
            </tbody>
          </table>

          {/* Detalle de ítems */}
          <table className="rp-items">
            <thead>
              <tr>
                <th>ÍTEM</th>
                <th>REF.</th>
                <th style={{ textAlign: "left" }}>DESCRIPCIÓN</th>
                <th>CANT.</th>
                <th>PRECIO UD.</th>
                <th>DESCUENTO</th>
                <th>TOTAL NETO</th>
              </tr>
            </thead>
            <tbody>
              {r.detalles.map((d) => (
                <tr key={d.item}>
                  <td>{d.item}</td>
                  <td>{d.ref}</td>
                  <td style={{ textAlign: "left" }}>{d.descripcion}</td>
                  <td>{Number(d.cantidad)}</td>
                  <td>{pesos(d.precioUnitario)}</td>
                  <td>{pesos(d.descuento)}</td>
                  <td>{pesos(d.subtotal)}</td>
                </tr>
              ))}
            </tbody>
          </table>
          <div className="rp-totalitems">
            Total ítems: {r.detalles.length} - Total unidades: {totalUnidades}
          </div>

          {/* Pago + totales */}
          <table className="rp-pago">
            <tbody>
              <tr>
                <td className="rp-pago-izq">
                  <div><strong>FORMA DE PAGO:</strong></div>
                  <div>{r.formaPago === "CONTADO" ? "Contado" : r.formaPago} - Valor: {pesos(r.total)}</div>
                  <div style={{ marginTop: 6 }}><strong>MEDIOS DE PAGO:</strong></div>
                  <div>{r.medioPago ?? "—"}{r.medioPago ? ` - Valor: ${pesos(r.total)}` : ""}</div>
                </td>
                <td className="rp-pago-med">-No registran impuestos ni retenciones-</td>
                <td className="rp-pago-tot">
                  <div><strong>SUBTOTAL</strong> {pesos(r.subtotal)}</div>
                  {Number(r.descuento) > 0 && <div><strong>DESCUENTO</strong> {pesos(r.descuento)}</div>}
                  <div className="rp-tot-final"><strong>TOTAL NETO</strong> {pesos(r.total)}</div>
                </td>
              </tr>
            </tbody>
          </table>

          {/* Garantía (solo productos; los valores a terceros no llevan) */}
          <div className="rp-garantia">
            ¡Muchas gracias por tu compra!{" "}
            {conGarantia.length === 0
              ? ""
              : garantias.length === 1
                ? `Todos nuestros productos cuentan con garantía de ${garantias[0]} meses.`
                : "Garantía por producto: " +
                  conGarantia.map((d) => `${d.nombre} (${d.garantiaMeses} meses)`).join(", ") +
                  "."}
          </div>

          {/* Firmas */}
          <table className="rp-firmas">
            <tbody>
              <tr>
                <td>
                  <div className="rp-firma-linea">_____________________</div>
                  Firma y sello emisor
                </td>
                <td><strong>Observación:</strong> {r.observacion ?? "-No registra-"}</td>
                <td>
                  <div className="rp-recibi"><strong>Recibí a satisfacción:</strong></div>
                  <div className="rp-firma-linea">___________________________________</div>
                  Nombre, identificación, sello y fecha
                </td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}
