// ==========================================================================
// Panel principal con navegación lateral.
// Cambia de vista (inicio, artículos, …) con estado local; cuando crezca se
// puede migrar a react-router sin tocar las pantallas.
// ==========================================================================
import { useState } from "react";
import { useAuth } from "../auth/AuthContext";
import { Articulos } from "./Articulos";
import { Inventario } from "./Inventario";
import { Compras } from "./Compras";
import { Clientes } from "./Clientes";
import { Proveedores } from "./Proveedores";
import { Empleados } from "./Empleados";
import { Remisiones } from "./Remisiones";
import { Comisiones } from "./Comisiones";
import { Reportes } from "./Reportes";
import { Usuarios } from "./Usuarios";

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
          Echo<span>·ERP</span>
        </span>
      </header>

      {/* Capa oscura que cierra el menú al tocar fuera (solo móvil) */}
      <div className="dash-overlay" onClick={() => setMenuAbierto(false)} />

      <aside className="dash-sidebar">
        <h1 className="logo">
          Echo<span>·ERP</span>
        </h1>
        <nav className="dash-nav">
          <Link id="inicio">Inicio</Link>
          <Link id="articulos">Artículos</Link>
          <Link id="inventario">Inventario</Link>
          <Link id="compras">Compras</Link>
          <Link id="proveedores">Proveedores</Link>
          <Link id="clientes">Clientes</Link>
          <Link id="empleados">Empleados</Link>
          <Link id="remisiones">Remisiones</Link>
          {esAdmin && <Link id="usuarios">Usuarios</Link>}
          {esAdmin && <Link id="comisiones">Comisiones</Link>}
          {esAdmin && <Link id="reportes">Reportes</Link>}
        </nav>
      </aside>

      <main className="dash-main">
        {vista === "inicio" && <Inicio nombre={usuario.nombre} rol={usuario.rol} esAdmin={esAdmin} onSalir={logout} />}
        {vista === "articulos" && <Articulos />}
        {vista === "inventario" && <Inventario />}
        {vista === "compras" && <Compras />}
        {vista === "proveedores" && <Proveedores />}
        {vista === "clientes" && <Clientes />}
        {vista === "empleados" && <Empleados />}
        {vista === "remisiones" && <Remisiones />}
        {vista === "comisiones" && esAdmin && <Comisiones />}
        {vista === "reportes" && esAdmin && <Reportes />}
        {vista === "usuarios" && esAdmin && <Usuarios />}
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
}: {
  nombre: string;
  rol: string;
  esAdmin: boolean;
  onSalir: () => void;
}) {
  const modulos = [
    { titulo: "Artículos", desc: "Crear y editar productos.", fase: "Disponible" },
    { titulo: "Inventario", desc: "Existencias, ajustes y traslados.", fase: "Disponible" },
    { titulo: "Compras", desc: "Remisiones de compra + Excel.", fase: "Disponible" },
    { titulo: "Proveedores", desc: "Directorio de proveedores.", fase: "Disponible" },
    { titulo: "Clientes", desc: "Directorio + importar/exportar.", fase: "Disponible" },
    { titulo: "Empleados", desc: "Personal por sede.", fase: "Disponible" },
    { titulo: "Remisiones", desc: "Ventas, garantía e impresión PDF.", fase: "Disponible" },
  ];
  const modulosAdmin = [
    { titulo: "Usuarios", desc: "Crear y administrar usuarios.", fase: "Disponible" },
    { titulo: "Comisiones", desc: "Pago a vendedoras (4% / 1%).", fase: "Disponible" },
    { titulo: "Reportes", desc: "Ventas, utilidades, top productos.", fase: "Disponible" },
  ];

  return (
    <>
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
          </div>
        ))}
      </div>
    </>
  );
}
