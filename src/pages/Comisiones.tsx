// ==========================================================================
// Reporte de COMISIONES de las vendedoras (solo ADMIN).
// Regla del negocio: 4% si el cliente es NUEVO, 1% si es RECOMPRA.
// La comisión ya viene "congelada" en cada remisión; aquí solo se agrupa por
// vendedor y rango de fechas para liquidar los pagos.
// ==========================================================================
import { useEffect, useState } from "react";
import { remisionesApi, catalogosApi, type ReporteComisiones, type Sede } from "../api/recursos";

const pesos = (v: number) => "$" + Math.round(Number(v)).toLocaleString("en-US");

// Primer día del mes actual en formato YYYY-MM-DD (valor por defecto del filtro).
function inicioDeMes() {
  const d = new Date();
  return new Date(d.getFullYear(), d.getMonth(), 1).toISOString().slice(0, 10);
}
function hoy() {
  return new Date().toISOString().slice(0, 10);
}

export function Comisiones() {
  const [reporte, setReporte] = useState<ReporteComisiones | null>(null);
  const [sedes, setSedes] = useState<Sede[]>([]);
  const [cargando, setCargando] = useState(true);
  const [error, setError] = useState<string | null>(null);

  const [desde, setDesde] = useState(inicioDeMes());
  const [hasta, setHasta] = useState(hoy());
  const [sedeId, setSedeId] = useState("");

  async function cargar() {
    setCargando(true);
    setError(null);
    try {
      // Sumamos un día a "hasta" para incluir todas las remisiones de ese día.
      const hastaFin = `${hasta}T23:59:59`;
      setReporte(await remisionesApi.comisiones({ desde, hasta: hastaFin, sedeId: sedeId || undefined }));
    } catch (e) {
      setError(e instanceof Error ? e.message : "Error al cargar el reporte.");
    } finally {
      setCargando(false);
    }
  }

  useEffect(() => {
    catalogosApi.sedes().then(setSedes);
    cargar();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  return (
    <div>
      <div className="dash-topbar">
        <h2>Comisiones de vendedoras</h2>
      </div>

      <div className="barra-busqueda" style={{ flexWrap: "wrap", gap: 12 }}>
        <div className="campo" style={{ marginBottom: 0 }}>
          <label>Desde</label>
          <input type="date" value={desde} onChange={(e) => setDesde(e.target.value)} />
        </div>
        <div className="campo" style={{ marginBottom: 0 }}>
          <label>Hasta</label>
          <input type="date" value={hasta} onChange={(e) => setHasta(e.target.value)} />
        </div>
        <div className="campo" style={{ marginBottom: 0 }}>
          <label>Sede</label>
          <select value={sedeId} onChange={(e) => setSedeId(e.target.value)}>
            <option value="">Todas</option>
            {sedes.map((s) => <option key={s.id} value={s.id}>{s.nombre}</option>)}
          </select>
        </div>
        <button className="btn-primario" style={{ width: "auto", alignSelf: "end" }} onClick={cargar}>
          Calcular
        </button>
      </div>

      <p className="muted">Comisión: 4% por venta a cliente nuevo · 1% por recompra.</p>

      {error && <div className="alerta-error">{error}</div>}

      {cargando ? (
        <p>Cargando…</p>
      ) : reporte ? (
        <div className="tabla-wrap">
          <table className="tabla">
            <thead>
              <tr>
                <th>Vendedor</th>
                <th>Remisiones</th>
                <th>Clientes nuevos</th>
                <th>Recompras</th>
                <th>Ventas</th>
                <th>Comisión a pagar</th>
              </tr>
            </thead>
            <tbody>
              {reporte.filas.map((f) => (
                <tr key={f.vendedor}>
                  <td><strong>{f.vendedor}</strong></td>
                  <td>{f.remisiones}</td>
                  <td>{f.nuevos}</td>
                  <td>{f.recompras}</td>
                  <td>{pesos(f.ventas)}</td>
                  <td><strong style={{ color: "var(--echo-azul)" }}>{pesos(f.comision)}</strong></td>
                </tr>
              ))}
              {reporte.filas.length === 0 && (
                <tr>
                  <td colSpan={6} className="muted" style={{ textAlign: "center", padding: 24 }}>
                    No hay remisiones en el rango seleccionado.
                  </td>
                </tr>
              )}
            </tbody>
            {reporte.filas.length > 0 && (
              <tfoot>
                <tr>
                  <td><strong>Total</strong></td>
                  <td><strong>{reporte.totales.remisiones}</strong></td>
                  <td></td>
                  <td></td>
                  <td><strong>{pesos(reporte.totales.ventas)}</strong></td>
                  <td><strong style={{ color: "var(--echo-azul)" }}>{pesos(reporte.totales.comision)}</strong></td>
                </tr>
              </tfoot>
            )}
          </table>
        </div>
      ) : null}
    </div>
  );
}
