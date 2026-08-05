// =============================================================================
// Parser de mensajes de WhatsApp con pedidos.
// Detecta: cliente, productos, forma de pago, domicilio/flete.
//
// Formatos soportados:
//   CON COLON  → "1 onn HD : 115.000"          (precio después de ":")
//   SIN COLON  → "1 echo show 5 blanca 290.000" (precio al final)
//   EMOJI      → "🧍Nombre: Daniela Espinosa"
//   GUIÓN      → "-Jairo Alexander González"
//   CC PUNTOS  → "Cc :1.093.220.049"
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
  medioPago: string;
  items: ItemParsadoWA[];
  domicilio: number; // 0 si no hay
}

/** Convierte precio colombiano a número: "115.000" → 115000, "480,000" → 480000 */
export function parsePrecioWA(s: string): number {
  return parseInt(s.replace(/\./g, "").replace(/,/g, ""), 10) || 0;
}

/** Elimina emojis, guiones y caracteres especiales del inicio de línea */
function limpiar(l: string): string {
  return l.replace(/^[\s\-*•·_⭐️⭐🌟✅📦🧍📞📍💵🪪✉️🛑#]+/, "").trim();
}

/** Extrae el valor después de una etiqueta ("Nombre: VALUE") en el texto completo */
function etiqueta(texto: string, regex: RegExp): string {
  return texto.match(regex)?.[1]?.trim() ?? "";
}

// Precio colombiano: dígitos + separador de miles (punto o coma) + 3 dígitos
// Ejemplos: "290.000", "1.390.000", "480,000", "32.600"
const PRECIO_COL = /\d{1,3}(?:[.,]\d{3})+/;

// Producto sin colon: "12 on hd 110.000 configurados"
// Captura: (cantidad) (descripcion) (precio_colombiano) [texto_extra_opcional]
const PROD_SIN_COLON = new RegExp(
  `^(\\d+)\\s+(.*?)\\s+(${PRECIO_COL.source})(?:\\s.*)?$`
);

// Líneas que siempre se ignoran (normalizado con limpiar())
const SKIP = /^(total|pago\s|transferencia|nequi|bancolombia|daviplata|efectivo|contra[\s-]*entrega|enviar\s|forma\s+de\s+pago|m[eé]todo\s+de\s+pago|m[eé]todo|nombre\s*:|tel[eé]fono\s*:|celular\s*:|correo\s*(electr[oó]nico)?\s*[:\s]|direcci[oó]n\s*:|cc|c\.c\.|c[eé]dula)/i;

// Palabras que identifican una dirección en la línea
const ES_DIR = /\b(calle|carrera|cra\.?|cll\.?|cl\.?|av\.?|avenida|diagonal|transversal|km\.?)\b/i;

export function parsearMensajeWA(texto: string): MensajeParsado {
  const lineas = texto.split("\n").map((l) => l.trim()).filter(Boolean);

  // ── Documento CC ─────────────────────────────────────────────────────────────
  // Acepta: "Cc43877391", "CC: 43877391", "Cc :1.093.220.049", "#de Cédula:1065006265"
  // Los números con puntos (1.093.220.049) son CC colombianos con separador de miles.
  const docMatch = texto.match(
    /(?:cc|c\.c\.?|c[eé]dula)\s*[:#]?\s*([\d.]{6,18})/i
  );
  const documento = docMatch ? docMatch[1].replace(/\./g, "") : "";

  // ── Teléfono (10 dígitos, inicia en 3) ───────────────────────────────────────
  const telMatch = texto.match(/\b(3\d{9})\b/);
  const telefono = telMatch?.[1] ?? "";

  // ── Email ─────────────────────────────────────────────────────────────────────
  const emailMatch = texto.match(/[\w.+-]+@[\w.-]+\.[a-zA-Z]{2,}/);
  const email = emailMatch?.[0] ?? "";

  // ── Medio de pago (default: Efectivo) ────────────────────────────────────────
  let medioPago = "Efectivo";
  if (/transferencia/i.test(texto))              medioPago = "Transferencia bancaria";
  else if (/nequi/i.test(texto))                 medioPago = "Nequi";
  else if (/daviplata/i.test(texto))             medioPago = "Daviplata";
  else if (/bancolombia/i.test(texto))           medioPago = "Transferencia bancaria - BANCOLOMBIA";
  else if (/contra[\s-]*entrega/i.test(texto))   medioPago = "Contra entrega";

  // ── Dirección ─────────────────────────────────────────────────────────────────
  // Primero busca etiqueta "Dirección: VALUE"; si no, busca línea con palabras de vía.
  const direccion =
    etiqueta(texto, /direcci[oó]n\s*:\s*([^\n]+)/i) ||
    lineas.find((l) => ES_DIR.test(limpiar(l))) ||
    "";

  // ── Función: ¿es una línea especial a omitir? ─────────────────────────────────
  function esEspecial(l: string): boolean {
    const ll = limpiar(l);
    return (
      ES_DIR.test(l) ||
      SKIP.test(ll) ||
      (emailMatch != null && l.includes(emailMatch[0])) ||
      (telefono !== "" && ll === telefono)
    );
  }

  // ── Items y domicilio ─────────────────────────────────────────────────────────
  const items: ItemParsadoWA[] = [];
  let domicilio = 0;

  for (const linea of lineas) {
    if (esEspecial(linea)) continue;

    const ll = limpiar(linea);

    // Domicilio/Envío/Flete como línea separada: "Domicilio : 18.000", "Envío: 32,600"
    if (/^(domicilio|env[íi]o|flete|transporte|despacho)\s*:/i.test(ll)) {
      const m = linea.match(/:\s*([\d.,]+)/);
      if (m) { domicilio = parsePrecioWA(m[1]); continue; }
    }

    // ── Formato CON COLON: "N nombre : precio [texto extra]" ──────────────────
    const conColon = ll.match(/^(\d+)\s+(.+?)\s*:\s*([\d.,]+)/);
    if (conColon) {
      const p = parsePrecioWA(conColon[3]);
      if (p > 0) {
        items.push({
          cantidad: parseInt(conColon[1], 10),
          descripcion: conColon[2].trim(),
          precio: p,
        });
        continue;
      }
    }

    // ── Formato SIN COLON: "N nombre precio [texto extra]" ────────────────────
    // El precio es el primer token con patrón de miles colombiano (≥ 1.000).
    const sinColon = ll.match(PROD_SIN_COLON);
    if (sinColon) {
      const p = parsePrecioWA(sinColon[3]);
      if (p >= 1000) {
        items.push({
          cantidad: parseInt(sinColon[1], 10),
          descripcion: sinColon[2].trim(),
          precio: p,
        });
      }
    }
  }

  // ── Nombre del cliente ────────────────────────────────────────────────────────
  // 1) Buscar etiqueta "Nombre: VALUE" (formato emoji/label).
  // 2) Fallback: primera línea de texto con ≥ 2 palabras, sin dígitos/símbolos,
  //    con letras minúsculas (evita encabezados de ciudad en MAYÚSCULAS).
  let nombre =
    etiqueta(texto, /(?:^|\n)[^\n]*nombre\s*:\s*([^\n]+)/im) ||
    lineas.find((l) => {
      const ll = limpiar(l);
      return (
        !esEspecial(l) &&
        !items.some((i) => ll.toLowerCase().includes(i.descripcion.toLowerCase())) &&
        /^[a-záéíóúüñA-ZÁÉÍÓÚÜÑ]/i.test(ll) &&           // empieza con letra
        !/[@\d:]/.test(ll) &&                               // sin @, dígitos, dos puntos
        /[a-záéíóúüñ]/.test(ll) &&                         // tiene al menos una minúscula
        ll.split(/\s+/).length >= 2                         // mínimo 2 palabras
      );
    }) ||
    "";

  // Si vino con prefijo especial (guión, asterisco), limpiarlo
  if (nombre && !/^[\p{L}]/u.test(nombre)) {
    nombre = limpiar(nombre);
  }

  return {
    documento,
    nombre,
    telefono,
    email,
    direccion,
    medioPago,
    items,
    domicilio,
  };
}

/**
 * Busca el artículo del catálogo más parecido a la descripción.
 * Incluye palabras de 2+ letras (hd, 4k, tv, etc.) para mayor precisión.
 */
export function buscarArticuloSimilar<
  T extends { id: string; nombre: string; codigo: string }
>(descripcion: string, catalogo: T[]): T | null {
  const palabras = descripcion
    .toLowerCase()
    .split(/\s+/)
    .filter((p) => p.length >= 2);

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
