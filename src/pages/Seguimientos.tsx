// ==========================================================================
// Módulo de Seguimientos de envíos (Skydropx).
// Muestra todos los envíos activos (no entregados) con su estado y los que
// requieren atención (novedad, pendiente más de 2 días, etc.).
// ==========================================================================
import { useEffect, useState } from "react";
import { skydropxApi, type Envio } from "../api/recursos";

const ESTADO_LABEL: Record<string, { texto: string; color: string }> = {
  pendiente_recoleccion: { texto: "Pendiente recolección", color: "#f59e0b" },
  en_transito:           { texto: "En tránsito",           color: "#3b82f6" },
  entregado:             { texto: "Entregado",             color: "#22c55e" },
  novedad:               { texto: "Novedad",               color: "#ef4444" },
  cancelado:             { texto: "Cancelado",             color: "#6b7280" },
};

const COBRO_LABEL: Record<string, string> = {
  pending:   "Cobro pendiente",
  collected: "Cobrado",
  failed:    "Cobro fallido",
  cancelled: "Cobro cancelado",
};

function diasDesde(fecha: string) {
  return Math.floor((Date.now() - new Date(fecha).getTime()) / 86_400_000);
}

function requiereAtencion(e: Envio) {
  if (e.estado === "novedad") return true;
  if (e.estado === "pendiente_recoleccion" && diasDesde(e.creadoEn) >= 2) return true;
  if (e.estadoCobro === "failed") return true;
  return false;
}

const pesos = (v: number) => "$" + Math.round(Number(v)).toLocaleString("en-US");
const fechaCorta = (s: string) => new Date(s).toLocaleDateString("es-CO");

export function Seguimientos() {
  const [envios, setEnvios] = useState<Envio[]>([]);
  const [cargando, setCargando] = useState(true);
  const [actualizando, setActualizando] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [filtro, setFiltro] = useState<"todos" | "atencion">("todos");

  async function cargar() {
    setCargando(true);
    setError(null);
    try {
      setEnvios(await skydropxApi.seguimientos());
    } catch (e) {
      setError(e instanceof Error ? e.message : "Error cargando seguimientos.");
    } finally {
      setCargando(false);
    }
  }

  async function actualizarTodos() {
    setActualizando(true);
    try {
      await skydropxApi.actualizarEstados();
      await cargar();
    } catch (e) {
      setError(e instanceof Error ? e.message : "Error actualizando estados.");
    } finally {
      setActualizando(false);
    }
  }

  useEffect(() => { cargar(); }, []);

  const visibles = filtro === "atencion" ? envios.filter(requiereAtencion) : envios;
  const conAtencion = envios.filter(requiereAtencion).length;

  return (
    <div className="pagina">
      <div className="pagina-header">
        <h1>Seguimientos de envíos</h1>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
          <button className="btn-secundario" onClick={cargar} disabled={cargando}>
            {cargando ? "Cargando…" : "Refrescar"}
          </button>
          <button className="btn-secundario" onClick={actualizarTodos} disabled={actualizando}>
            {actualizando ? "Actualizando…" : "Actualizar estados desde Skydropx"}
          </button>
        </div>
      </div>

      {error && <div className="alerta-error">{error}</div>}

      {/* Resumen rápido */}
      <div style={{ display: "flex", gap: 12, marginBottom: 20, flexWrap: "wrap" }}>
        <TarjetaResumen titulo="Total activos" valor={envios.length} color="#3b82f6" />
        <TarjetaResumen titulo="Requieren atención" valor={conAtencion} color="#ef4444" />
        <TarjetaResumen titulo="En tránsito" valor={envios.filter((e) => e.estado === "en_transito").length} color="#3b82f6" />
        <TarjetaResumen titulo="Pendiente recolección" valor={envios.filter((e) => e.estado === "pendiente_recoleccion").length} color="#f59e0b" />
      </div>

      {/* Filtros */}
      <div style={{ display: "flex", gap: 8, marginBottom: 16 }}>
        {(["todos", "atencion"] as const).map((f) => (
          <button
            key={f}
            className={filtro === f ? "btn-primario" : "btn-secundario"}
            style={{ width: "auto" }}
            onClick={() => setFiltro(f)}
          >
            {f === "todos" ? "Todos" : `Requieren atención (${conAtencion})`}
          </button>
        ))}
      </div>

      {cargando ? (
        <p className="muted">Cargando…</p>
      ) : visibles.length === 0 ? (
        <p className="muted">{filtro === "atencion" ? "Ningún envío requiere atención." : "No hay envíos activos."}</p>
      ) : (
        <div style={{ overflowX: "auto" }}>
          <table className="tabla">
            <thead>
              <tr>
                <th>Alerta</th>
                <th>Remisión</th>
                <th>Cliente</th>
                <th>Ciudad</th>
                <th>Transportadora</th>
                <th>Tracking</th>
                <th>Estado</th>
                <th>Contraentrega</th>
                <th>Cobro</th>
                <th>Días</th>
                <th>Guía</th>
              </tr>
            </thead>
            <tbody>
              {visibles.map((e) => {
                const alerta = requiereAtencion(e);
                const est = ESTADO_LABEL[e.estado] ?? { texto: e.estado, color: "#6b7280" };
                const dias = diasDesde(e.creadoEn);
                return (
                  <tr key={e.id} style={{ background: alerta ? "rgba(239,68,68,0.06)" : undefined }}>
                    <td style={{ textAlign: "center", fontSize: 18 }}>
                      {alerta ? "⚠️" : ""}
                    </td>
                    <td style={{ fontWeight: 600 }}>
                      {e.remision?.documento ?? "—"}<br />
                      <span className="muted" style={{ fontSize: 12 }}>{e.remision ? fechaCorta(e.remision.fecha) : ""}</span>
                    </td>
                    <td>
                      {e.remision?.cliente.nombre ?? "—"}<br />
                      <span className="muted" style={{ fontSize: 12 }}>
                        {e.remision?.cliente.whatsapp ?? e.remision?.cliente.telefono ?? ""}
                      </span>
                    </td>
                    <td>{e.remision?.cliente.ciudad ?? "—"}</td>
                    <td style={{ textTransform: "capitalize" }}>{e.carrier}</td>
                    <td style={{ fontFamily: "monospace", fontSize: 13 }}>{e.tracking}</td>
                    <td>
                      <span style={{
                        background: est.color + "22",
                        color: est.color,
                        borderRadius: 6,
                        padding: "2px 8px",
                        fontSize: 12,
                        fontWeight: 600,
                        whiteSpace: "nowrap",
                      }}>
                        {est.texto}
                      </span>
                    </td>
                    <td style={{ textAlign: "right" }}>
                      {e.montoContraentrega ? pesos(e.montoContraentrega) : "—"}
                    </td>
                    <td>
                      {e.estadoCobro ? (
                        <span style={{ fontSize: 12, color: e.estadoCobro === "collected" ? "#22c55e" : e.estadoCobro === "failed" ? "#ef4444" : "#f59e0b" }}>
                          {COBRO_LABEL[e.estadoCobro] ?? e.estadoCobro}
                        </span>
                      ) : "—"}
                    </td>
                    <td style={{ textAlign: "center", color: dias >= 5 ? "#ef4444" : dias >= 3 ? "#f59e0b" : undefined }}>
                      {dias}d
                    </td>
                    <td>
                      {e.guiaUrl ? (
                        <a href={e.guiaUrl} target="_blank" rel="noreferrer" style={{ color: "var(--primario)", fontSize: 13 }}>
                          PDF
                        </a>
                      ) : "—"}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

function TarjetaResumen({ titulo, valor, color }: { titulo: string; valor: number; color: string }) {
  return (
    <div style={{
      background: "var(--superficie)", border: "1px solid var(--borde,#ddd)",
      borderRadius: 10, padding: "12px 20px", minWidth: 140, textAlign: "center",
    }}>
      <div style={{ fontSize: 28, fontWeight: 700, color }}>{valor}</div>
      <div style={{ fontSize: 13, color: "var(--muted,#888)" }}>{titulo}</div>
    </div>
  );
}
