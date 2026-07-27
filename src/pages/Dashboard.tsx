// ==========================================================================
// Panel principal con navegación lateral.
// Cambia de vista (inicio, artículos, …) con estado local; cuando crezca se
// puede migrar a react-router sin tocar las pantallas.
// ==========================================================================
import { lazy, Suspense, useState } from "react";
import { useAuth } from "../auth/AuthContext";

// Lazy loading: cada módulo se descarga solo cuando el usuario lo abre por
// primera vez, reduciendo el JS inicial de ~1 MB a ~150 KB.
const Articulos   = lazy(() => import("./Articulos").then((m) => ({ default: m.Articulos })));
const Inventario  = lazy(() => import("./Inventario").then((m) => ({ default: m.Inventario })));
const Compras     = lazy(() => import("./Compras").then((m) => ({ default: m.Compras })));
const Clientes    = lazy(() => import("./Clientes").then((m) => ({ default: m.Clientes })));
const Proveedores = lazy(() => import("./Proveedores").then((m) => ({ default: m.Proveedores })));
const Empleados   = lazy(() => import("./Empleados").then((m) => ({ default: m.Empleados })));
const Remisiones  = lazy(() => import("./Remisiones").then((m) => ({ default: m.Remisiones })));
const Comisiones  = lazy(() => import("./Comisiones").then((m) => ({ default: m.Comisiones })));
const Reportes    = lazy(() => import("./Reportes").then((m) => ({ default: m.Reportes })));
const Usuarios    = lazy(() => import("./Usuarios").then((m) => ({ default: m.Usuarios })));

type Vista =
  | "inicio"
  | "articulos"
  | "inventario"
  | "compras"
  | "proveedores"
  | "clientes"
  | "empleados"
  | "remisiones"
  | "usuarios"
  | "comisiones"
  | "reportes";

// Logo de la marca. Usa la imagen /logo-echo.png (colócala en erp-frontend/public).
// Si la imagen no está disponible, muestra el texto "Echo·ERP" como respaldo.
function LogoEcho({ alto = 70 }: { alto?: number }) {
  const [falla, setFalla] = useState(false);
  if (falla) {
    return (
      <span className="logo">
        Echo<span>·ERP</span>
      </span>
    );
  }
  return (
    <img
      src="/logo-echo.png"
      alt="ECHO · Tecnología en Casa"
      className="logo-img"
      style={{ height: alto }}
      onError={() => setFalla(true)}
    />
  );
}

export function Dashboard() {
  const { usuario, logout } = useAuth();
  const [vista, setVista] = useState<Vista>("inicio");
  // Estado del menú lateral en móvil (en escritorio siempre está visible).
  const [menuAbierto, setMenuAbierto] = useState(false);
  if (!usuario) return null;

  const esAdmin = usuario.rol === "ADMIN";

  // Enlace de la barra lateral que marca el activo, cambia de vista y
  // cierra el menú en móvil tras elegir una opción.
  const Link = ({ id, children }: { id: Vista; children: string }) => (
    <a
      href="#"
      className={vista === id ? "activo" : ""}
      onClick={(e) => {
        e.preventDefault();
        setVista(id);
        setMenuAbierto(false);
      }}
    >
      {children}
    </a>
  );

  return (
    <div className={`dash${menuAbierto ? " menu-abierto" : ""}`}>
      {/* Barra superior solo visible en móvil: botón hamburguesa + logo */}
      <header className="dash-movil-bar">
        <button
          className="btn-menu"
          aria-label="Abrir menú"
          aria-expanded={menuAbierto}
          onClick={() => setMenuAbierto((v) => !v)}
        >
          <span></span>
          <span></span>
          <span></span>
        </button>
        <span className="dash-movil-logo">
          <LogoEcho alto={34} />
        </span>
      </header>

      {/* Capa oscura que cierra el menú al tocar fuera (solo móvil) */}
      <div className="dash-overlay" onClick={() => setMenuAbierto(false)} />

      <aside className="dash-sidebar">
        <div className="logo-sidebar">
          <LogoEcho />
        </div>
        <nav className="dash-nav">
          <Link id="inicio">Inicio</Link>
          <Link id="articulos">Artículos</Link>
          <Link id="inventario">Inventario</Link>
          {esAdmin && <Link id="compras">Compras</Link>}
          <Link id="proveedores">Proveedores</Link>
          <Link id="clientes">Clientes</Link>
          {esAdmin && <Link id="empleados">Empleados</Link>}
          <Link id="remisiones">Remisiones</Link>
          {esAdmin && <Link id="usuarios">Usuarios</Link>}
          {esAdmin && <Link id="comisiones">Comisiones</Link>}
          {esAdmin && <Link id="reportes">Reportes</Link>}
        </nav>
        {/* Cerrar sesión siempre visible, al pie de la barra lateral */}
        <button className="btn-salir-side" onClick={logout}>
          Cerrar sesión
        </button>
      </aside>

      <main className="dash-main">
        {vista === "inicio" && (
          <Inicio
            nombre={usuario.nombre}
            rol={usuario.rol}
            esAdmin={esAdmin}
            onSalir={logout}
            onIr={setVista}
          />
        )}
        <Suspense fallback={<div style={{ padding: 32 }}>Cargando módulo…</div>}>
          {vista === "articulos"   && <Articulos />}
          {vista === "inventario"  && <Inventario />}
          {vista === "compras"     && esAdmin && <Compras />}
          {vista === "proveedores" && <Proveedores />}
          {vista === "clientes"    && <Clientes />}
          {vista === "empleados"   && esAdmin && <Empleados />}
          {vista === "remisiones"  && <Remisiones />}
          {vista === "comisiones"  && esAdmin && <Comisiones />}
          {vista === "reportes"    && esAdmin && <Reportes />}
          {vista === "usuarios"    && esAdmin && <Usuarios />}
        </Suspense>
      </main>
    </div>
  );
}

// Vista de inicio con las tarjetas de módulos.
function Inicio({
  nombre,
  rol,
  esAdmin,
  onSalir,
  onIr,
}: {
  nombre: string;
  rol: string;
  esAdmin: boolean;
  onSalir: () => void;
  onIr: (v: Vista) => void;
}) {
  const modulos: { id: Vista; titulo: string; desc: string; fase: string }[] = [
    { id: "articulos", titulo: "Artículos", desc: "Crear y editar productos.", fase: "Disponible" },
    { id: "inventario", titulo: "Inventario", desc: "Existencias, ajustes y traslados.", fase: "Disponible" },
    { id: "proveedores", titulo: "Proveedores", desc: "Directorio de proveedores.", fase: "Disponible" },
    { id: "clientes", titulo: "Clientes", desc: "Directorio + importar/exportar.", fase: "Disponible" },
    { id: "remisiones", titulo: "Remisiones", desc: "Ventas, garantía e impresión PDF.", fase: "Disponible" },
  ];
  const modulosAdmin: { id: Vista; titulo: string; desc: string; fase: string }[] = [
    { id: "compras", titulo: "Compras", desc: "Remisiones de compra + Excel.", fase: "Disponible" },
    { id: "empleados", titulo: "Empleados", desc: "Personal por sede.", fase: "Disponible" },
    { id: "usuarios", titulo: "Usuarios", desc: "Crear y administrar usuarios.", fase: "Disponible" },
    { id: "comisiones", titulo: "Comisiones", desc: "Pago a vendedoras (4% / 1%).", fase: "Disponible" },
    { id: "reportes", titulo: "Reportes", desc: "Ventas, utilidades, top productos.", fase: "Disponible" },
  ];
  return (
    <>
      <div className="inicio-logo">
        <LogoEcho alto={96} />
      </div>
      <div className="dash-topbar">
        <h2>
          Hola, {nombre}
          <span className="chip-rol">{rol}</span>
        </h2>
        <button className="btn-salir" onClick={onSalir}>
          Cerrar sesión
        </button>
      </div>
      <div className="tarjetas">
        {[...modulos, ...(esAdmin ? modulosAdmin : [])].map((m) => (
          <div className="tarjeta" key={m.titulo}>
            <h3>{m.titulo}</h3>
            <p>{m.desc}</p>
            <p style={{ marginTop: 8, color: "var(--echo-azul)" }}>{m.fase}</p>
            <button
              className="btn-primario"
              style={{ width: "auto", marginTop: 14 }}
              onClick={() => onIr(m.id)}
            >
              Entrar →
            </button>
          </div>
        ))}
      </div>
    </>
  );
}
