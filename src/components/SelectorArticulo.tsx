// ==========================================================================
// Selector de artículo con BÚSQUEDA (por nombre o código), reutilizable en
// remisiones, compras e inventario. Evita depender de un desplegable largo.
// Muestra el artículo elegido y, al enfocar, permite teclear para filtrar.
// Si se pasa sedeId, muestra el stock disponible junto a cada opción.
// ==========================================================================
import { useState } from "react";
import type { Articulo } from "../api/recursos";

export function SelectorArticulo({
  articulos,
  valor,
  onElegir,
  placeholder = "Buscar artículo por nombre o código…",
  sedeId,
}: {
  articulos: Articulo[];
  valor: string;
  onElegir: (id: string) => void;
  placeholder?: string;
  sedeId?: string;
}) {
  const [q, setQ] = useState("");
  const [abierto, setAbierto] = useState(false);
  const elegido = articulos.find((a) => a.id === valor);
  const filtro = q.trim().toLowerCase();
  const opciones = (filtro
    ? articulos.filter(
        (a) => a.nombre.toLowerCase().includes(filtro) || a.codigo.toLowerCase().includes(filtro)
      )
    : articulos
  ).slice(0, 15);

  const stockEnSede = (a: Articulo) => {
    if (!sedeId) return null;
    const inv = a.inventario?.find((i) => i.sede.id === sedeId);
    return inv ? Number(inv.cantidad) : 0;
  };

  return (
    <div style={{ position: "relative", flex: 1, minWidth: 220 }}>
      <input
        value={abierto ? q : elegido ? `${elegido.codigo} · ${elegido.nombre}` : q}
        placeholder={placeholder}
        onFocus={() => { setAbierto(true); setQ(""); }}
        onChange={(e) => { setQ(e.target.value); setAbierto(true); }}
        onBlur={() => setTimeout(() => setAbierto(false), 150)}
      />
      {abierto && opciones.length > 0 && (
        <div
          className="lista-opciones"
          style={{ position: "absolute", zIndex: 10, width: "100%", maxHeight: 240, overflowY: "auto" }}
        >
          {opciones.map((a) => {
            const stock = stockEnSede(a);
            const sinStock = stock !== null && stock <= 0;
            return (
              <button
                type="button"
                key={a.id}
                className="opcion"
                onMouseDown={() => { onElegir(a.id); setAbierto(false); }}
              >
                <strong>{a.codigo}</strong> · {a.nombre}
                {stock !== null && (
                  <span style={{
                    marginLeft: 8,
                    fontSize: "0.8em",
                    fontWeight: 600,
                    color: sinStock ? "#c53030" : "#276749",
                    background: sinStock ? "#fff5f5" : "#f0fff4",
                    border: `1px solid ${sinStock ? "#feb2b2" : "#9ae6b4"}`,
                    borderRadius: 4,
                    padding: "1px 6px",
                  }}>
                    {sinStock ? "Sin stock" : `${stock} uds`}
                  </span>
                )}
              </button>
            );
          })}
        </div>
      )}
    </div>
  );
}
