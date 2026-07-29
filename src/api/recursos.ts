// ==========================================================================
// Funciones tipadas para consumir los recursos del ERP (artículos, inventario).
// Centralizan las rutas para no repetir strings por toda la app.
// ==========================================================================
import { api } from "./client";

export interface InventarioSede {
  cantidad: number;
  costoPromedio: number;
  sede: { id: string; nombre: string };
}

export interface Articulo {
  id: string;
  codigo: string;
  nombre: string;
  descripcion: string | null;
  tipo: string;
  tarifaNormal: number;
  precioMinimo: number;
  costo: number;
  garantiaMeses: number;
  urlFoto: string | null;
  activo: boolean;
  esTercero: boolean; // flete u otro valor a terceros (no inventario/venta/utilidad)
  categoria: { nombre: string } | null;
  marca: { nombre: string } | null;
  inventario: InventarioSede[];
  stockTotal: number;
}

export interface NuevoArticulo {
  codigo: string;
  nombre: string;
  descripcion?: string;
  tipo?: string;
  costo?: number;
  precioMinimo?: number;
  tarifaNormal?: number;
  tarifaMayorista?: number;
  garantiaMeses?: number;
  urlFoto?: string;
  // Nombre de marca (el backend la crea/reutiliza). Alternativa a marcaId.
  marca?: string;
  activo?: boolean;
  // Valor a terceros (flete): no controla inventario ni cuenta como venta.
  esTercero?: boolean;
}

// Fila plana para importar artículos desde el Excel de Effi (claves ya
// normalizadas en el frontend a partir de los encabezados).
export interface FilaImportArticulo {
  codigo?: string;
  nombre?: string;
  marca?: string;
  descripcion?: string;
  urlImagen?: string;
  costoManual?: number;
  costoPromedio?: number;
  ultimoCosto?: number;
  precioMinimo?: number;
  tarifaNormal?: number;
  garantiaMeses?: number;
  stockMedellin?: number;
  stockBogota?: number;
}

export interface ResumenImportArticulos {
  recibidas: number;
  creadas: number;
  omitidas: number; // ya existían (mismo código)
  descartadas: number; // sin nombre
  conStock: number; // con stock inicial cargado
  errores: { codigo: string; nombre: string; error: string }[];
}

// --- Artículos ---
export const articulosApi = {
  listar: (buscar = "") =>
    api<Articulo[]>(`/articulos${buscar ? `?buscar=${encodeURIComponent(buscar)}` : ""}`),
  crear: (datos: NuevoArticulo) =>
    api<Articulo>("/articulos", { method: "POST", body: JSON.stringify(datos) }),
  actualizar: (id: string, datos: Partial<NuevoArticulo>) =>
    api<Articulo>(`/articulos/${id}`, { method: "PUT", body: JSON.stringify(datos) }),
  eliminar: (id: string) =>
    api<{ ok: boolean }>(`/articulos/${id}`, { method: "DELETE" }),
  // Borrado masivo del catálogo (limpiar para reimportar). Solo ADMIN.
  // forzar=true limpia también remisiones/compras/traslados (reset de migración).
  eliminarTodos: (forzar = false) =>
    api<{ ok: boolean; eliminados: number; documentosBorrados: boolean }>(
      `/articulos${forzar ? "?forzar=true" : ""}`,
      { method: "DELETE" }
    ),
  importar: (filas: FilaImportArticulo[]) =>
    api<ResumenImportArticulos>("/articulos/importar", {
      method: "POST",
      body: JSON.stringify({ filas }),
    }),
};

// --- Subida de imágenes ---
// No usa el cliente JSON porque envía FormData (multipart). El backend devuelve
// una URL absoluta que se guarda en Articulo.urlFoto.
const API_BASE = import.meta.env.VITE_API_URL ?? "http://localhost:4000/api";
export const uploadsApi = {
  subirImagen: async (archivo: File): Promise<{ url: string }> => {
    const fd = new FormData();
    fd.append("imagen", archivo);
    const token = localStorage.getItem("echo_erp_token");
    const res = await fetch(`${API_BASE}/uploads/imagen`, {
      method: "POST",
      // OJO: no fijar Content-Type, el navegador pone el boundary del multipart.
      headers: token ? { Authorization: `Bearer ${token}` } : undefined,
      body: fd,
    });
    const cuerpo = await res.json().catch(() => ({}));
    if (!res.ok) {
      throw new Error((cuerpo as { mensaje?: string }).mensaje ?? "Error al subir la imagen.");
    }
    return cuerpo as { url: string };
  },
};

// --- Inventario ---
export interface Movimiento {
  id: string;
  tipo: string;
  cantidad: number;
  fecha: string;
  referencia: string | null;
  articulo: { codigo: string; nombre: string };
  sede: { nombre: string };
}

// Existencia de un artículo en una sede concreta.
export interface Existencia {
  id: string;
  cantidad: number;
  costoPromedio: number;
  stockMinimo: number;
  articulo: { id: string; codigo: string; nombre: string; tarifaNormal: number };
  sede: { id: string; nombre: string };
}

export interface NuevoTraslado {
  sedeOrigenId: string;
  sedeDestinoId: string;
  observacion?: string;
  items: { articuloId: string; cantidad: number }[];
}

export const inventarioApi = {
  // Existencias por sede (filtros opcionales de sede y búsqueda).
  existencias: (sedeId = "", buscar = "") => {
    const qs = new URLSearchParams();
    if (sedeId) qs.set("sedeId", sedeId);
    if (buscar) qs.set("buscar", buscar);
    const s = qs.toString();
    return api<Existencia[]>(`/inventario${s ? `?${s}` : ""}`);
  },
  // Kardex / trazabilidad (filtros opcionales de sede y tipo de movimiento).
  movimientos: (filtros: { sedeId?: string; tipo?: string } = {}) => {
    const qs = new URLSearchParams();
    if (filtros.sedeId) qs.set("sedeId", filtros.sedeId);
    if (filtros.tipo) qs.set("tipo", filtros.tipo);
    const s = qs.toString();
    return api<Movimiento[]>(`/inventario/movimientos${s ? `?${s}` : ""}`);
  },
  ajustar: (datos: {
    articuloId: string;
    sedeId: string;
    cantidad: number;
    costoUnitario?: number;
    observacion?: string;
  }) => api("/inventario/ajustes", { method: "POST", body: JSON.stringify(datos) }),
  trasladar: (datos: NuevoTraslado) =>
    api<{ ok: boolean; numero: number }>("/inventario/traslados", {
      method: "POST",
      body: JSON.stringify(datos),
    }),
};

// --- Catálogos ---
export interface Sede {
  id: string;
  nombre: string;
  consecutivoRemision?: number; // próximo número de remisión de la sede
}
export interface Marca {
  id: string;
  nombre: string;
}
export const catalogosApi = {
  sedes: () => api<Sede[]>("/catalogos/sedes"),
  marcas: () => api<Marca[]>("/catalogos/marcas"),
  // Fija el próximo número de remisión de una sede (consecutivo). Solo ADMIN.
  actualizarConsecutivo: (id: string, consecutivoRemision: number) =>
    api<{ id: string; nombre: string; consecutivoRemision: number }>(
      `/catalogos/sedes/${id}/consecutivo`,
      { method: "PATCH", body: JSON.stringify({ consecutivoRemision }) }
    ),
  eliminarMarca: (id: string) =>
    api<{ ok: boolean; articulosDesvinculados: number }>(`/catalogos/marcas/${id}`, {
      method: "DELETE",
    }),
  // Pasa todas las marcas a MAYÚSCULAS y fusiona las duplicadas.
  normalizarMarcas: () =>
    api<{ ok: boolean; fusionadas: number; renombradas: number }>("/catalogos/marcas/normalizar", {
      method: "POST",
    }),
};

// --- Compras (remisiones de compra) ---
export interface Compra {
  id: string;
  documento: string; // BOG-574, MED-1274…
  fecha: string;
  sede: string;
  proveedor: string;
  proveedorDocumento: string; // NIT/documento del proveedor
  remisionProveedor: string | null;
  observacion: string | null;
  subtotal: number;
  total: number;
  estado: string;
  items: number;
}

// Edición de la cabecera de una compra (no toca las líneas).
export interface EditarCompra {
  proveedor: { documento: string; nombre: string };
  remisionProveedor?: string;
  observacion?: string;
}

export interface NuevaCompraItem {
  articuloId: string;
  descripcion?: string;
  cantidad: number;
  costoUnitario: number;
  descuento?: number;
}

export interface NuevaCompra {
  sedeId: string;
  proveedor: { documento: string; nombre: string; direccion?: string };
  remisionProveedor?: string;
  observacion?: string;
  items: NuevaCompraItem[];
}

// Fila plana para importar desde Excel.
export interface FilaImportCompra {
  sede: string;
  proveedorDocumento: string;
  proveedorNombre: string;
  proveedorDireccion?: string;
  remisionProveedor?: string;
  articuloCodigo: string;
  descripcion?: string;
  cantidad: number;
  costoUnitario: number;
  descuento?: number;
}

export interface CompraDetalle {
  id: string;
  codigo: string;
  descripcion: string;
  cantidad: number;
  costoUnitario: number;
  descuento: number;
  subtotal: number;
}

export interface CompraCompleta {
  id: string;
  documento: string;
  fecha: string;
  sede: string;
  sedeDireccion: string | null;
  proveedor: string;
  proveedorDocumento: string;
  proveedorDireccion: string | null;
  remisionProveedor: string | null;
  observacion: string | null;
  subtotal: number;
  descuento: number;
  total: number;
  elaboradoPor: string | null;
  detalles: CompraDetalle[];
}

export const comprasApi = {
  listar: () => api<Compra[]>("/compras"),
  obtener: (id: string) => api<CompraCompleta>(`/compras/${id}`),
  crear: (datos: NuevaCompra) =>
    api<{ id: string }>("/compras", { method: "POST", body: JSON.stringify(datos) }),
  importar: (filas: FilaImportCompra[]) =>
    api<{ remisionesCreadas: number; filasProcesadas: number }>("/compras/importar", {
      method: "POST",
      body: JSON.stringify({ filas }),
    }),
  editar: (id: string, datos: EditarCompra) =>
    api<{ id: string }>(`/compras/${id}`, { method: "PUT", body: JSON.stringify(datos) }),
  eliminar: (id: string) =>
    api<{ ok: boolean }>(`/compras/${id}`, { method: "DELETE" }),
};

// --- Clientes ---
export interface Cliente {
  id: string;
  documento: string;
  tipoIdentificacion: string | null;
  nombre: string;
  email: string | null;
  telefono1: string | null;
  telefono2: string | null;
  celular: string | null;
  whatsapp: string | null;
  pais: string | null;
  departamento: string | null;
  ciudad: string | null;
  direccion: string | null;
  tipoPersona: string | null;
  tipoCliente: string | null;
  tarifaPrecios: string | null;
  formaPago: string | null;
  moneda: string | null;
  vendedor: string | null;
  observacion: string | null;
  activo: boolean;
  compraPrevia: boolean; // ya tenía compras antes (recompra)
}

// Datos editables de un cliente (crear/editar).
export type NuevoCliente = Partial<Omit<Cliente, "id" | "activo">> & {
  documento: string;
  nombre: string;
};

export interface ListadoClientes {
  total: number;
  pagina: number;
  tam: number;
  datos: Cliente[];
}

export interface ResumenImportClientes {
  recibidas: number;
  validas: number;
  creadas: number;
  duplicadasOmitidas: number;
  descartadas: number;
}

// Fila flexible para importar (claves ya normalizadas desde el Excel de Effi).
export type FilaImportCliente = Record<string, string | number | undefined>;

export const clientesApi = {
  listar: (buscar = "", pagina = 1, tam = 50) => {
    const qs = new URLSearchParams({ pagina: String(pagina), tam: String(tam) });
    if (buscar) qs.set("buscar", buscar);
    return api<ListadoClientes>(`/clientes?${qs.toString()}`);
  },
  exportar: () => api<Cliente[]>("/clientes/exportar"),
  crear: (datos: NuevoCliente) =>
    api<Cliente>("/clientes", { method: "POST", body: JSON.stringify(datos) }),
  actualizar: (id: string, datos: NuevoCliente) =>
    api<Cliente>(`/clientes/${id}`, { method: "PUT", body: JSON.stringify(datos) }),
  importar: (filas: FilaImportCliente[]) =>
    api<ResumenImportClientes>("/clientes/importar", {
      method: "POST",
      body: JSON.stringify({ filas }),
    }),
  eliminar: (id: string) =>
    api<{ ok: boolean }>(`/clientes/${id}`, { method: "DELETE" }),
  // Marca TODOS los clientes existentes como "con compra previa" (recompra).
  marcarRecompra: () =>
    api<{ ok: boolean; actualizados: number }>("/clientes/marcar-recompra", { method: "POST" }),
};

// --- Proveedores ---
export interface Proveedor {
  id: string;
  documento: string;
  nombre: string;
  direccion: string | null;
  telefono: string | null;
  email: string | null;
}
// La dirección ya no se captura en el formulario de proveedor.
export type NuevoProveedor = Omit<Proveedor, "id" | "direccion">;

export const proveedoresApi = {
  listar: (buscar = "") =>
    api<Proveedor[]>(`/proveedores${buscar ? `?buscar=${encodeURIComponent(buscar)}` : ""}`),
  crear: (datos: NuevoProveedor) =>
    api<Proveedor>("/proveedores", { method: "POST", body: JSON.stringify(datos) }),
  actualizar: (id: string, datos: NuevoProveedor) =>
    api<Proveedor>(`/proveedores/${id}`, { method: "PUT", body: JSON.stringify(datos) }),
};

// --- Empleados ---
export interface Empleado {
  id: string;
  documento: string;
  nombre: string;
  cargo: string | null;
  telefono: string | null;
  email: string | null;
  sedeId: string | null;
  sede: { id: string; nombre: string } | null;
}
export interface NuevoEmpleado {
  documento: string;
  nombre: string;
  cargo?: string;
  telefono?: string;
  email?: string;
  sedeId?: string;
}

export const empleadosApi = {
  listar: (buscar = "") =>
    api<Empleado[]>(`/empleados${buscar ? `?buscar=${encodeURIComponent(buscar)}` : ""}`),
  crear: (datos: NuevoEmpleado) =>
    api<Empleado>("/empleados", { method: "POST", body: JSON.stringify(datos) }),
  actualizar: (id: string, datos: NuevoEmpleado) =>
    api<Empleado>(`/empleados/${id}`, { method: "PUT", body: JSON.stringify(datos) }),
  eliminar: (id: string) =>
    api<{ ok: boolean }>(`/empleados/${id}`, { method: "DELETE" }),
};

// --- Remisiones de VENTA ---
// Fila del listado de remisiones.
export interface Remision {
  id: string;
  documento: string; // BOG-574, MED-9480…
  numero: number;
  fecha: string;
  sede: string;
  cliente: string;
  clienteDocumento: string;
  vendedor: string;
  estado: string;
  total: number;
  esRecompra: boolean;
  comisionPct: number;
  comisionValor: number;
  items: number;
}

// Detalle completo (para imprimir el PDF).
export interface RemisionDetalleLinea {
  item: number;
  articuloId: string;
  ref: string;
  nombre: string;
  esTercero: boolean;
  descripcion: string;
  cantidad: number;
  precioUnitario: number;
  descuento: number;
  subtotal: number;
  garantiaMeses: number;
}
export interface RemisionCompleta {
  id: string;
  documento: string;
  numero: number;
  fecha: string;
  vencimiento: string | null;
  estado: string;
  sedeId: string;
  clienteId: string;
  sede: { nombre: string; direccion: string | null; telefono: string | null };
  cliente: {
    nombre: string;
    documento: string;
    tipoIdentificacion: string | null;
    telefono: string;
    whatsapp: string | null;
    email: string | null;
    direccion: string | null;
    ciudad: string | null;
    departamento: string | null;
    pais: string | null;
  };
  elaboro: string;
  vendedor: string | null;
  formaPago: string;
  medioPago: string | null;
  observacion: string | null;
  subtotal: number;
  descuento: number;
  total: number;
  esRecompra: boolean;
  comisionPct: number;
  comisionValor: number;
  detalles: RemisionDetalleLinea[];
}

export interface NuevaRemisionItem {
  articuloId: string;
  descripcion?: string;
  cantidad: number;
  precioUnitario?: number;
  descuento?: number;
  garantiaMeses?: number;
}
export interface NuevaRemision {
  sedeId: string;
  clienteId: string;
  vendedor?: string;
  observacion?: string;
  medioPago?: string;
  items: NuevaRemisionItem[];
}

// Fila plana para importar remisiones desde Excel (Google Sheets / WhatsApp).
export interface FilaImportRemision {
  sede: string;
  clienteDocumento: string | number; // número de documento (Excel lo manda como number)
  clienteTipoIdentificacion?: string;
  clienteNombre?: string;
  clienteTelefono?: string | number;
  clienteDireccion?: string;
  clienteCiudad?: string;
  clienteDepartamento?: string;
  vendedor?: string;
  remision?: string;
  fecha?: string;
  medioPago?: string;
  observacion?: string;
  articuloCodigo: string;
  descripcion?: string;
  cantidad: number;
  precioUnitario: number;
  descuento?: number;
}

// Reporte de comisiones por vendedor.
export interface FilaComision {
  vendedor: string;
  nuevos: number;
  recompras: number;
  ventas: number;
  comision: number;
  remisiones: number;
}
export interface ReporteComisiones {
  filas: FilaComision[];
  totales: { ventas: number; comision: number; remisiones: number };
}

// --------------------------------------------------------------------------
// CUENTA DE COBRO (generada bajo demanda desde una remisión)
// --------------------------------------------------------------------------
export interface CuentaCobro {
  tipo: string;
  numero: string;
  numeroConsecutivo: number;
  fecha: string;
  vencimiento: string;
  estado: string;
  emisor: {
    nombre: string;
    representante: string;
    nit: string;
    regimen: string;
    sede: string;
    direccion: string;
    telefono: string;
    email: string;
    web: string;
  };
  cliente: {
    nombre: string;
    documento: string;
    tipoIdentificacion: string;
    telefono: string;
    direccion: string | null;
    ciudad: string | null;
    departamento: string | null;
    pais: string | null;
  };
  remisionDocumento: string;
  remisionFecha: string;
  items: {
    item: number;
    codigo: string;
    descripcion: string;
    cantidad: number;
    precioUnitario: number;
    descuento: number;
    subtotal: number;
  }[];
  subtotal: number;
  descuento: number;
  total: number;
  formaPago: string;
  medioPago: string | null;
  observacion: string | null;
  elaboro: string;
  vendedor: string | null;
}

export const remisionesApi = {
  listar: (
    filtros: {
      sedeId?: string;
      vendedor?: string;
      buscar?: string;
      desde?: string;
      hasta?: string;
      pagina?: number;
      tam?: number;
    } = {}
  ) => {
    const qs = new URLSearchParams();
    Object.entries(filtros).forEach(([k, v]) => v && qs.set(k, String(v)));
    const s = qs.toString();
    return api<{ total: number; pagina: number; tam: number; datos: Remision[] }>(
      `/remisiones${s ? `?${s}` : ""}`
    );
  },
  obtener: (id: string) => api<RemisionCompleta>(`/remisiones/${id}`),
  crear: (datos: NuevaRemision) =>
    api<{ id: string; numero: number }>("/remisiones", { method: "POST", body: JSON.stringify(datos) }),
  anular: (id: string) =>
    api<{ ok: boolean }>(`/remisiones/${id}/anular`, { method: "PATCH" }),
  importar: (filas: FilaImportRemision[]) =>
    api<{ remisionesCreadas: number; filasProcesadas: number }>("/remisiones/importar", {
      method: "POST",
      body: JSON.stringify({ filas }),
    }),
  comisiones: (filtros: { desde?: string; hasta?: string; sedeId?: string } = {}) => {
    const qs = new URLSearchParams();
    Object.entries(filtros).forEach(([k, v]) => v && qs.set(k, String(v)));
    const s = qs.toString();
    return api<ReporteComisiones>(`/remisiones/comisiones${s ? `?${s}` : ""}`);
  },
  cuentaCobro: (id: string) =>
    api<CuentaCobro>(`/remisiones/${id}/cuenta-cobro`),
  enviarCorreo: (id: string, emailDestinatario?: string) =>
    api<{ ok: boolean; enviadoA: string }>(`/remisiones/${id}/enviar-correo`, {
      method: "POST",
      body: JSON.stringify({ emailDestinatario }),
    }),
  enviarCuentaCobroCorreo: (id: string, emailDestinatario?: string) =>
    api<{ ok: boolean; enviadoA: string }>(`/remisiones/${id}/enviar-cuenta-cobro-correo`, {
      method: "POST",
      body: JSON.stringify({ emailDestinatario }),
    }),
};

// --- Usuarios (solo ADMIN) ---
export interface Usuario {
  id: string;
  nombre: string;
  email: string;
  rol: "ADMIN" | "OPERADOR";
  sedeId: string | null;
  activo: boolean;
  creadoEn: string;
}
export interface NuevoUsuario {
  nombre: string;
  email: string;
  password: string;
  rol: "ADMIN" | "OPERADOR";
  sedeId?: string;
}

// Edición de un usuario (sin contraseña; esa tiene su propio endpoint).
export interface EditarUsuario {
  nombre?: string;
  rol?: "ADMIN" | "OPERADOR";
  sedeId?: string | null;
}

export const usuariosApi = {
  listar: () => api<Usuario[]>("/usuarios"),
  crear: (datos: NuevoUsuario) =>
    api<Usuario>("/usuarios", { method: "POST", body: JSON.stringify(datos) }),
  actualizar: (id: string, datos: EditarUsuario) =>
    api<Usuario>(`/usuarios/${id}`, { method: "PUT", body: JSON.stringify(datos) }),
  // Activar / desactivar un usuario.
  cambiarEstado: (id: string, activo: boolean) =>
    api<Usuario>(`/usuarios/${id}/estado`, {
      method: "PATCH",
      body: JSON.stringify({ activo }),
    }),
  // Cambiar la contraseña de un usuario.
  cambiarPassword: (id: string, password: string) =>
    api<{ ok: boolean }>(`/usuarios/${id}/password`, {
      method: "PATCH",
      body: JSON.stringify({ password }),
    }),
  // Eliminar un usuario.
  eliminar: (id: string) =>
    api<{ ok: boolean }>(`/usuarios/${id}`, { method: "DELETE" }),
};

// --- Reportes (solo ADMIN) ---
export interface ResumenVentas {
  totalVentas: number;
  numRemisiones: number;
  ticketPromedio: number;
  totalUnidades: number;
  utilidad: number;
  margen: number; // %
  porSede: { sede: string; ventas: number; utilidad: number; remisiones: number }[];
}
export interface ValorMercanciaArticulo {
  codigo: string;
  nombre: string;
  marca: string;
  sedes: { nombre: string; cantidad: number; costoPromedio: number; valor: number }[];
  totalUnidades: number;
  totalValor: number;
}
export interface ValorMercancia {
  articulos: ValorMercanciaArticulo[];
  totalEmpresa: number;
}

export interface ProductoTop {
  codigo: string;
  nombre: string;
  unidades: number;
  ventas: number;
  utilidad: number;
}

export const reportesApi = {
  ventas: (filtros: { desde?: string; hasta?: string; sedeId?: string } = {}) => {
    const qs = new URLSearchParams();
    Object.entries(filtros).forEach(([k, v]) => v && qs.set(k, String(v)));
    const s = qs.toString();
    return api<ResumenVentas>(`/reportes/ventas${s ? `?${s}` : ""}`);
  },
  topProductos: (filtros: { desde?: string; hasta?: string; sedeId?: string; limite?: number } = {}) => {
    const qs = new URLSearchParams();
    Object.entries(filtros).forEach(([k, v]) => v && qs.set(k, String(v)));
    const s = qs.toString();
    return api<ProductoTop[]>(`/reportes/top-productos${s ? `?${s}` : ""}`);
  },
  valorMercancia: () => api<ValorMercancia>("/reportes/valor-mercancia"),
};

// --- Auth público (recuperación de contraseña) ---
export const authApi = {
  olvidarPassword: (email: string) =>
    api<{ ok: boolean }>("/auth/forgot-password", { method: "POST", body: JSON.stringify({ email }) }),
  resetPassword: (token: string, password: string) =>
    api<{ ok: boolean }>("/auth/reset-password", { method: "POST", body: JSON.stringify({ token, password }) }),
};
