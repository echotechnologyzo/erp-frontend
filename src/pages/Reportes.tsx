// ==========================================================================
// Reportes para el ADMIN: promedio de ventas, utilidades y productos más
// vendidos. Todo sale de las remisiones de venta emitidas (precio y costo
// "congelados"), filtrable por rango de fechas y sede.
// ==========================================================================
import { useEffect, useState } from "react";
import {
  reportesApi,
  catalogosApi,
  type ResumenVentas,
  type ProductoTop,
  type Sede,
} from "../api/recursos";

const pesos = (v: number) => "$" + Math.round(Number(v)).toLocaleString("en-US");
const pct = (v: number) => `${Number(v).toFixed(1)}%`;

function inicioDeMes() {
  const d = new Date();
  return new Date(d.getFullYear(), d.getMonth(), 1).toISOString().slice(0, 10);
}
function hoy() {
  return new Date().toISOString().slice(0, 10);
}

export function Reportes() {
  const [ventas, setVentas] = useState<ResumenVentas | null>(null);
  const [top, setTop] = useState<ProductoTop[]>([]);
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
      const filtros = { desde, hasta: `${hasta}T23:59:59`, sedeId: sedeId || undefined };
      const [v, t] = await Promise.all([
        reportesApi.ventas(filtros),
        reportesApi.topProductos({ ...filtros, limite: 15 }),
      ]);
      setVentas(v);
      setTop(t);
    } catch (e) {
      setError(e instanceof Error ? e.message : "Error al cargar los reportes.");
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
        <h2>Reportes</h2>
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
          Generar
        </button>
      </div>

      {error && <div className="alerta-error">{error}</div>}

      {cargando ? (
        <p>Cargando…</p>
      ) : ventas ? (
        <>
          {/* Tarjetas resumen */}
          <div className="tarjetas">
            <div className="tarjeta">
              <p className="muted">Ventas totales</p>
              <h3 style={{ color: "var(--echo-azul)" }}>{pesos(ventas.totalVentas)}</h3>
              <p className="muted">{ventas.totalUnidades} unidades</p>
            </div>
            <div className="tarjeta">
              <p className="muted">Remisiones</p>
              <h3>{ventas.numRemisiones}</h3>
              <p className="muted">Ticket prom.: {pesos(ventas.ticketPromedio)}</p>
            </div>
            <div className="tarjeta">
              <p className="muted">Utilidad</p>
              <h3 style={{ color: "var(--echo-azul)" }}>{pesos(ventas.utilidad)}</h3>
              <p className="muted">Margen: {pct(ventas.margen)}</p>
            </div>
          </div>

          {/* Ventas por sede */}
          <h3 style={{ margin: "18px 0 8px" }}>Ventas por sede</h3>
          <div className="tabla-wrap">
            <table className="tabla">
              <thead>
                <tr>
                  <th>Sede</th>
                  <th>Remisiones</th>
                  <th>Ventas</th>
                  <th>Utilidad</th>
                </tr>
              </thead>
              <tbody>
                {ventas.porSede.map((s) => (
                  <tr key={s.sede}>
                    <td><strong>{s.sede}</strong></td>
                    <td>{s.remisiones}</td>
                    <td>{pesos(s.ventas)}</td>
                    <td>{pesos(s.utilidad)}</td>
                  </tr>
                ))}
                {ventas.porSede.length === 0 && (
                  <tr>
                    <td colSpan={4} className="muted" style={{ textAlign: "center", padding: 24 }}>
                      No hay ventas en el rango seleccionado.
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>

          {/* Productos más vendidos */}
          <h3 style={{ margin: "18px 0 8px" }}>Productos más vendidos</h3>
          <div className="tabla-wrap">
            <table className="tabla">
              <thead>
                <tr>
                  <th>#</th>
                  <th>Código</th>
                  <th>Artículo</th>
                  <th>Unidades</th>
                  <th>Ventas</th>
                  <th>Utilidad</th>
                </tr>
              </thead>
              <tbody>
                {top.map((p, i) => (
                  <tr key={p.codigo}>
                    <td>{i + 1}</td>
                    <td>{p.codigo}</td>
                    <td><strong>{p.nombre}</strong></td>
                    <td>{p.unidades}</td>
                    <td>{pesos(p.ventas)}</td>
                    <td>{pesos(p.utilidad)}</td>
                  </tr>
                ))}
                {top.length === 0 && (
                  <tr>
                    <td colSpan={6} className="muted" style={{ textAlign: "center", padding: 24 }}>
                      Sin datos en el rango seleccionado.
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </>
      ) : null}
    </div>
  );
}
