// =============================================================================
// Utilidades para parsear mensajes de WhatsApp con pedidos.
// Detecta automáticamente: cliente, productos, forma de pago y domicilio.
//
// Formato típico de mensaje:
//   ⭐ la estrella
//   1 onn HD : 115.000
//   Lina maria alzate cano
//   3003926693
//   Calle 82 sur 61 48...
//   Pago por transferencia
//   Cc43877391
//   Lalzatenaly@hotmail.com
//   Domicilio : 18.000
//   Total : 133.000
// =============================================================================

export interface ItemParsadoWA {
  cantidad: number;
  descripcion: string; // texto original del WhatsApp
  precio: number;      // pesos colombianos
}

export interface MensajeParsado {
  documento: string;
  nombre: string;
  telefono: string;
  email: string;
  direccion: string;
  medioPago: string;   // descripción libre del medio de pago
  items: ItemParsadoWA[];
  domicilio: number;   // 0 si no hay
}

/** Convierte precio colombiano ("115.000", "1.500.000") a número entero */
export function parsePrecioWA(s: string): number {
  // En Colombia el punto es separador de miles: "115.000" → 115000
  return parseInt(s.replace(/\./g, "").replace(/,/g, ""), 10) || 0;
}

/** Extrae los datos del pedido del mensaje de WhatsApp */
export function parsearMensajeWA(texto: string): MensajeParsado {
  const lineas = texto.split("\n").map((l) => l.trim()).filter(Boolean);

  // ── Documento CC ────────────────────────────────────────────────────────────
  // Acepta: "Cc43877391", "CC: 43877391", "Cedula 43877391"
  const docMatch = texto.match(/(?:cc|c\.c\.?|cedula|cédula)\s*:?\s*(\d{5,12})/i);
  const documento = docMatch?.[1] ?? "";

  // ── Teléfono (10 dígitos colombianos, inicia en 3) ──────────────────────────
  const telMatch = texto.match(/\b(3\d{9})\b/);
  const telefono = telMatch?.[1] ?? "";

  // ── Email ────────────────────────────────────────────────────────────────────
  const emailMatch = texto.match(/[\w.+-]+@[\w.-]+\.[a-zA-Z]{2,}/);
  const email = emailMatch?.[0] ?? "";

  // ── Medio de pago (por defecto Efectivo si no se especifica) ────────────────
  let medioPago = "Efectivo"; // default
  if (/transferencia/i.test(texto))    medioPago = "Transferencia bancaria";
  else if (/nequi/i.test(texto))       medioPago = "Nequi";
  else if (/daviplata/i.test(texto))   medioPago = "Daviplata";
  else if (/bancolombia/i.test(texto)) medioPago = "Transferencia bancaria - BANCOLOMBIA";
  // "efectivo" explícito ya es el default, pero lo dejamos explícito igual

  // ── Dirección (línea con palabras clave de vía) ─────────────────────────────
  const esDireccion = /\b(calle|carrera|cra\.?|cll\.?|cl\.?|av\.?|avenida|diagonal|transversal|km\.?)\b/i;
  const direccion = lineas.find((l) => esDireccion.test(l)) ?? "";

  // ── Líneas a ignorar en el nombre/productos ──────────────────────────────────
  const esLineaEspecial = (l: string) =>
    esDireccion.test(l) ||
    /^(total|pago\s|transferencia|nequi|bancolombia|daviplata|efectivo|cc|c\.c\.|cedula|cédula)/i.test(l) ||
    /^[⭐🌟✅📦]/.test(l) ||
    (emailMatch != null && l.includes(emailMatch[0])) ||
    (telefono !== "" && l === telefono);

  // ── Productos ("1 onn HD : 115.000") y domicilio ────────────────────────────
  const items: ItemParsadoWA[] = [];
  let domicilio = 0;

  for (const linea of lineas) {
    if (esDireccion.test(linea)) continue;
    if (esLineaEspecial(linea)) continue;

    // Domicilio: "Domicilio : 18.000"
    if (/^domicilio\s*:/i.test(linea)) {
      const m = linea.match(/:\s*([\d.]+)/);
      if (m) domicilio = parsePrecioWA(m[1]);
      continue;
    }

    // Producto: "N <nombre> : <precio>" — acepta texto adicional tras el precio
    // ("cada uno", "c/u", etc.)
    const prodMatch = linea.match(/^(\d+)\s+(.+?)\s*:\s*([\d.,]+)/);
    if (prodMatch) {
      items.push({
        cantidad: parseInt(prodMatch[1], 10),
        descripcion: prodMatch[2].trim(),
        precio: parsePrecioWA(prodMatch[3]),
      });
    }
  }

  // ── Nombre del cliente ───────────────────────────────────────────────────────
  // Primera línea de solo letras, ≥ 2 palabras, que no sea dato especial
  const nombre = lineas.find(
    (l) =>
      !esLineaEspecial(l) &&
      !items.some((i) => l.includes(i.descripcion)) &&
      /^[a-záéíóúüñA-ZÁÉÍÓÚÜÑ]/i.test(l) &&
      !/[@\d:]/.test(l) &&
      l.split(/\s+/).length >= 2
  ) ?? "";

  return { documento, nombre, telefono, email, direccion, medioPago, items, domicilio };
}

/** Busca el artículo del catálogo más parecido a la descripción del WhatsApp */
export function buscarArticuloSimilar<
  T extends { id: string; nombre: string; codigo: string }
>(descripcion: string, catalogo: T[]): T | null {
  const palabras = descripcion
    .toLowerCase()
    .split(/\s+/)
    .filter((p) => p.length > 2);

  if (palabras.length === 0) return null;

  let mejor: T | null = null;
  let mejorPuntaje = 0;

  for (const art of catalogo) {
    const nombreArt = art.nombre.toLowerCase();
    const puntaje = palabras.reduce(
      (acc, p) => acc + (nombreArt.includes(p) ? 1 : 0),
      0
    );
    if (puntaje > mejorPuntaje) {
      mejor = art;
      mejorPuntaje = puntaje;
    }
  }

  return mejorPuntaje > 0 ? mejor : null;
}
