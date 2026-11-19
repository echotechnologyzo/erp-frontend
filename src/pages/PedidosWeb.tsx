// ==========================================================================
// Pedidos de la tienda web (echo-web) + recuperación de carritos abandonados.
// Lista y filtra los pedidos por estado. Los PENDIENTE son carritos que no
// completaron el pago (abandonados): se resaltan y traen botón de WhatsApp
// para hacer seguimiento. También permite cambiar el estado del pedido.
// ==========================================================================
import { Fragment, useCallback, useEffect, useState } from "react";
import { pedidosWebApi, type PedidoWeb, type GuiaPedido } from "../api/recursos";

const pesos = (v: number) => "$" + Math.round(Number(v)).toLocaleString("es-CO");
const fechaHora = (s: string) =>
  new Date(s).toLocaleString("es-CO", { day: "2-digit", month: "2-digit", hour: "2-digit", minute: "2-digit" });

const ESTADOS: { id: string; texto: string; color: string }[] = [
  { id: "PENDIENTE", texto: "Pendiente (abandonado)", color: "#f59e0b" },
  { id: "PAGADO", texto: "Pagado", color: "#22c55e" },
  { id: "CONFIRMADO", texto: "Confirmado", color: "#0ea5e9" },
  { id: "PREPARANDO", texto: "Preparando", color: "#6366f1" },
  { id: "ENVIADO", texto: "Enviado", color: "#3b82f6" },
  { id: "ENTREGADO", texto: "Entregado", color: "#16a34a" },
  { id: "CANCELADO", texto: "Cancelado", color: "#6b7280" },
];
const EST = (id: string) => ESTADOS.find((e) => e.id === id) ?? { id, texto: id, color: "#6b7280" };

const TAM = 20;

function waLink(tel: string, nombre: string, ref: string, total: number, estado: string): string {
  const soloDig = tel.replace(/\D/g, "");
  const numero = soloDig.startsWith("57") ? soloDig : `57${soloDig}`;
  const saludo = `Hola ${nombre.split(" ")[0] || ""} 👋, soy de Echo Tecnología.`;
  const cuerpo =
    estado === "PENDIENTE"
      ? `Vi que empezaste un pedido (${ref}) por ${pesos(total)} y quedó pendiente de pago. ¿Te ayudo a completarlo?`
      : `Te escribo por tu pedido ${ref} (${pesos(total)}).`;
  return `https://wa.me/${numero}?text=${encodeURIComponent(`${saludo} ${cuerpo}`)}`;
}

// Mensaje de RECONFIRMACIÓN del pedido para enviar al cliente por WhatsApp.
function mensajeReconfirmacion(p: PedidoWeb): string {
  const primer = p.cliente.nombre.split(" ")[0] || "";
  const productos = p.items
    .map((i) => `${i.cantidad}× ${i.nombre}${i.color ? ` (${i.color})` : ""}`)
    .join(", ");
  const dir = `${p.envio.direccion}, ${p.envio.ciudad}${p.envio.referencia ? ` (${p.envio.referencia})` : ""}`;
  const con =
    p.metodoPago === "WOMPI"
      ? "Pago en línea (Wompi)"
      : `Contraentrega${p.transportadora ? ` con ${p.transportadora}` : ""}`;
  return (
    `Hola ${primer} 👋 Soy de Echo Tecnología, confirmo tu pedido ${p.referencia}:\n\n` +
    `Nombre: ${p.cliente.nombre}\n` +
    `Dirección: ${dir}\n` +
    `Teléfono: ${p.cliente.telefono}\n` +
    `Producto: ${productos}\n` +
    `Total: ${pesos(p.total)}\n` +
    `Con: ${con}\n\n` +
    `¿Todo correcto para despacharlo? 🙌`
  );
}

// Bloque de guía: genera la guía de Skydropx para el pedido (recotiza con la
// dirección real + contraentrega si aplica) y muestra tracking + PDF.
function BotonGuia({ p, onDone }: { p: PedidoWeb; onDone: () => void }) {
  const [cargando, setCargando] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [guia, setGuia] = useState<GuiaPedido | null>(null);

  if (p.zonaEnvio !== "NACIONAL") {
    return (
      <div>
        <strong>Guía</strong>
        <div style={{ fontSize: 13, color: "#64748b", marginTop: 6, maxWidth: 240 }}>
          Mensajero directo (zona fija), sin guía de transportadora.
        </div>
      </div>
    );
  }

  async function generar() {
    if (!window.confirm(`¿Generar la guía de Skydropx para ${p.referencia}?\nEsto crea un envío REAL con la transportadora.`)) return;
    setCargando(true);
    setError(null);
    try {
      const res = await pedidosWebApi.generarGuia(p.referencia);
      setGuia(res.guia);
      onDone();
    } catch (e) {
      setError(e instanceof Error ? e.message : "No se pudo generar la guía.");
    } finally {
      setCargando(false);
    }
  }

  const ya = p.tieneGuia || !!guia;
  return (
    <div>
      <strong>Guía</strong>
      <div style={{ marginTop: 6 }}>
        {!ya && (
          <button type="button" className="btn-secundario" disabled={cargando} onClick={generar}>
            {cargando ? "Generando…" : "🏷️ Generar guía"}
          </button>
        )}
        {guia && (
          <div style={{ fontSize: 13, marginTop: 6, lineHeight: 1.6 }}>
            ✓ {guia.carrier} · {guia.tracking}
            {guia.precio != null ? ` · ${pesos(guia.precio)}` : ""}
            {guia.guiaUrl && (
              <>
                <br />
                <a href={guia.guiaUrl} target="_blank" rel="noopener noreferrer">
                  Ver / descargar guía (PDF)
                </a>
              </>
            )}
          </div>
        )}
        {!guia && p.tieneGuia && (
          <div style={{ fontSize: 13, color: "#16a34a", marginTop: 6 }}>✓ Guía ya generada para este pedido.</div>
        )}
        {error && <div style={{ fontSize: 13, color: "#dc2626", marginTop: 6, maxWidth: 240 }}>{error}</div>}
      </div>
    </div>
  );
}

// Bloque de reconfirmación: copiar el mensaje o abrir el WhatsApp del cliente.
function BotonReconfirmar({ p }: { p: PedidoWeb }) {
  const [copiado, setCopiado] = useState(false);
  const msg = mensajeReconfirmacion(p);
  const tel = p.cliente.telefono.replace(/\D/g, "");
  const numero = tel.startsWith("57") ? tel : `57${tel}`;

  async function copiar() {
    try {
      await navigator.clipboard.writeText(msg);
    } catch {
      // Fallback para navegadores sin API de portapapeles.
      const ta = document.createElement("textarea");
      ta.value = msg;
      document.body.appendChild(ta);
      ta.select();
      try { document.execCommand("copy"); } catch { /* noop */ }
      document.body.removeChild(ta);
    }
    setCopiado(true);
    setTimeout(() => setCopiado(false), 1600);
  }

  return (
    <div>
      <strong>Reconfirmación</strong>
      <div style={{ display: "flex", gap: 8, marginTop: 6, flexWrap: "wrap" }}>
        <button type="button" className="btn-secundario" onClick={copiar}>
          {copiado ? "✓ Copiado" : "📋 Copiar"}
        </button>
        <a
          href={`https://wa.me/${numero}?text=${encodeURIComponent(msg)}`}
          target="_blank"
          rel="noopener noreferrer"
          style={{
            display: "inline-block",
            background: "#25D366",
            color: "#fff",
            borderRadius: 8,
            padding: "6px 12px",
            fontWeight: 600,
            textDecoration: "none",
            fontSize: 13,
          }}
        >
          WhatsApp con datos
        </a>
      </div>
      <pre
        style={{
          marginTop: 8,
          whiteSpace: "pre-wrap",
          fontFamily: "inherit",
          fontSize: 12,
          color: "#64748b",
          background: "#fff",
          border: "1px solid #e2e8f0",
          borderRadius: 8,
          padding: 10,
          maxWidth: 320,
        }}
      >
        {msg}
      </pre>
    </div>
  );
}

export function PedidosWeb() {
  const [datos, setDatos] = useState<PedidoWeb[]>([]);
  const [total, setTotal] = useState(0);
  const [resumen, setResumen] = useState<Record<string, number>>({});
  const [estado, setEstado] = useState("PENDIENTE"); // arranca en abandonados
  const [buscar, setBuscar] = useState("");
  const [pagina, setPagina] = useState(1);
  const [cargando, setCargando] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [expandido, setExpandido] = useState<string | null>(null);

  const cargar = useCallback(async () => {
    setCargando(true);
    setError(null);
    try {
      const [lista, res] = await Promise.all([
        pedidosWebApi.listar(estado, buscar, pagina, TAM),
        pedidosWebApi.resumen(),
      ]);
      setDatos(lista.datos);
      setTotal(lista.total);
      setResumen(res);
    } catch (e) {
      setError(e instanceof Error ? e.message : "Error cargando pedidos.");
    } finally {
      setCargando(false);
    }
  }, [estado, buscar, pagina]);

  useEffect(() => {
    cargar();
  }, [cargar]);

  async function cambiarEstado(ref: string, nuevo: string) {
    try {
      await pedidosWebApi.actualizarEstado(ref, nuevo);
      await cargar();
    } catch (e) {
      alert(e instanceof Error ? e.message : "No se pudo cambiar el estado.");
    }
  }

  const totalPaginas = Math.max(1, Math.ceil(total / TAM));
  const totalPedidos = Object.values(resumen).reduce((a, b) => a + b, 0);

  return (
    <div>
      <div className="dash-topbar">
        <h2>Pedidos web</h2>
      </div>

      {/* Filtros por estado */}
      <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginBottom: 16 }}>
        <Chip activo={estado === ""} onClick={() => { setEstado(""); setPagina(1); }} color="#334155">
          Todos ({totalPedidos})
        </Chip>
        {ESTADOS.map((e) => (
          <Chip
            key={e.id}
            activo={estado === e.id}
            onClick={() => { setEstado(e.id); setPagina(1); }}
            color={e.color}
          >
            {e.texto} ({resumen[e.id] ?? 0})
          </Chip>
        ))}
      </div>

      {/* Búsqueda */}
      <div style={{ marginBottom: 16, maxWidth: 420 }}>
        <input
          className="campo-input"
          placeholder="Buscar por referencia, cliente, correo, teléfono o ciudad…"
          value={buscar}
          onChange={(e) => { setBuscar(e.target.value); setPagina(1); }}
        />
      </div>

      {error && <p style={{ color: "#ef4444" }}>{error}</p>}
      {cargando ? (
        <p style={{ color: "#64748b" }}>Cargando pedidos…</p>
      ) : datos.length === 0 ? (
        <p style={{ color: "#64748b" }}>No hay pedidos en este filtro.</p>
      ) : (
        <div style={{ overflowX: "auto" }}>
          <table className="tabla" style={{ width: "100%", borderCollapse: "collapse", fontSize: 14 }}>
            <thead>
              <tr style={{ textAlign: "left", borderBottom: "2px solid #e5e7eb" }}>
                <th style={th}>Fecha</th>
                <th style={th}>Referencia</th>
                <th style={th}>Cliente</th>
                <th style={th}>Ciudad</th>
                <th style={th}>Total</th>
                <th style={th}>Pago</th>
                <th style={th}>Estado</th>
                <th style={th}></th>
              </tr>
            </thead>
            <tbody>
              {datos.map((p) => {
                const abandonado = p.estado === "PENDIENTE";
                return (
                  <Fragment key={p.referencia}>
                    <tr
                      style={{
                        borderBottom: "1px solid #eef2f7",
                        background: abandonado ? "rgba(245,158,11,0.07)" : undefined,
                        cursor: "pointer",
                      }}
                      onClick={() => setExpandido(expandido === p.referencia ? null : p.referencia)}
                    >
                      <td style={td}>{fechaHora(p.creadoEn)}</td>
                      <td style={{ ...td, fontFamily: "monospace", fontSize: 12 }}>{p.referencia}</td>
                      <td style={td}>
                        <div style={{ fontWeight: 600 }}>{p.cliente.nombre}</div>
                        <div style={{ color: "#64748b", fontSize: 12 }}>{p.cliente.telefono}</div>
                      </td>
                      <td style={td}>{p.envio.ciudad}</td>
                      <td style={{ ...td, fontWeight: 700, whiteSpace: "nowrap" }}>{pesos(p.total)}</td>
                      <td style={td}>{p.metodoPago === "WOMPI" ? "En línea" : "Contraentrega"}</td>
                      <td style={td}>
                        <span style={{ ...badge, background: EST(p.estado).color }}>{EST(p.estado).texto}</span>
                      </td>
                      <td style={{ ...td, whiteSpace: "nowrap" }} onClick={(e) => e.stopPropagation()}>
                        <a
                          href={waLink(p.cliente.telefono, p.cliente.nombre, p.referencia, p.total, p.estado)}
                          target="_blank"
                          rel="noopener noreferrer"
                          style={{
                            display: "inline-block",
                            background: "#25D366",
                            color: "#fff",
                            borderRadius: 8,
                            padding: "6px 12px",
                            fontWeight: 600,
                            textDecoration: "none",
                            fontSize: 13,
                          }}
                        >
                          WhatsApp
                        </a>
                      </td>
                    </tr>
                    {expandido === p.referencia && (
                      <tr style={{ background: "#f8fafc" }}>
                        <td colSpan={8} style={{ padding: 16 }}>
                          <div style={{ display: "flex", flexWrap: "wrap", gap: 24 }}>
                            <div>
                              <strong>Productos</strong>
                              <ul style={{ margin: "6px 0 0", paddingLeft: 18 }}>
                                {p.items.map((i) => (
                                  <li key={i.slug}>
                                    {i.cantidad}× {i.nombre}
                                    {i.color ? ` (${i.color})` : ""} — {pesos(i.subtotal)}
                                  </li>
                                ))}
                              </ul>
                              <div style={{ marginTop: 6, color: "#64748b", fontSize: 13 }}>
                                Subtotal {pesos(p.subtotal)} · Envío {pesos(p.costoEnvio)}
                                {p.transportadora ? ` (${p.transportadora})` : ""}
                              </div>
                            </div>
                            <div>
                              <strong>Contacto y envío</strong>
                              <div style={{ fontSize: 13, marginTop: 6, lineHeight: 1.6 }}>
                                {p.cliente.email}
                                <br />
                                {p.cliente.documento ? `Doc: ${p.cliente.documento}` : ""}
                                <br />
                                {p.envio.direccion}, {p.envio.ciudad} ({p.envio.departamento})
                                {p.envio.referencia ? ` — ${p.envio.referencia}` : ""}
                              </div>
                            </div>
                            <div>
                              <strong>Cambiar estado</strong>
                              <div style={{ marginTop: 6 }}>
                                <select
                                  className="campo-input"
                                  value={p.estado}
                                  onChange={(e) => cambiarEstado(p.referencia, e.target.value)}
                                >
                                  {ESTADOS.map((e) => (
                                    <option key={e.id} value={e.id}>
                                      {e.texto}
                                    </option>
                                  ))}
                                </select>
                              </div>
                            </div>
                            <BotonGuia p={p} onDone={cargar} />
                            <BotonReconfirmar p={p} />
                          </div>
                        </td>
                      </tr>
                    )}
                  </Fragment>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      {/* Paginación */}
      {total > TAM && (
        <div style={{ display: "flex", alignItems: "center", gap: 12, marginTop: 16 }}>
          <button className="btn-secundario" disabled={pagina <= 1} onClick={() => setPagina((p) => p - 1)}>
            ← Anterior
          </button>
          <span style={{ color: "#64748b" }}>
            Página {pagina} de {totalPaginas} · {total} pedidos
          </span>
          <button
            className="btn-secundario"
            disabled={pagina >= totalPaginas}
            onClick={() => setPagina((p) => p + 1)}
          >
            Siguiente →
          </button>
        </div>
      )}
    </div>
  );
}

const th: React.CSSProperties = { padding: "8px 10px", color: "#475569", fontWeight: 600 };
const td: React.CSSProperties = { padding: "10px" };
const badge: React.CSSProperties = {
  color: "#fff",
  borderRadius: 999,
  padding: "3px 10px",
  fontSize: 12,
  fontWeight: 600,
  whiteSpace: "nowrap",
};

function Chip({
  activo,
  onClick,
  color,
  children,
}: {
  activo: boolean;
  onClick: () => void;
  color: string;
  children: React.ReactNode;
}) {
  return (
    <button
      onClick={onClick}
      style={{
        border: `1.5px solid ${activo ? color : "#e2e8f0"}`,
        background: activo ? color : "#fff",
        color: activo ? "#fff" : "#334155",
        borderRadius: 999,
        padding: "6px 14px",
        fontSize: 13,
        fontWeight: 600,
        cursor: "pointer",
      }}
    >
      {children}
    </button>
  );
}
