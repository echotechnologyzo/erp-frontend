// ==========================================================================
// Lectura robusta de archivos Excel para las importaciones.
//
// Soporta dos formatos:
//   1. .xlsx / .xls binarios (se leen como ArrayBuffer).
//   2. La "tabla HTML guardada como .xls" que exporta Effi, codificada en
//      Windows-1252 SIN <meta charset>. Si se lee con la detección automática
//      de SheetJS, los acentos corrompen las etiquetas </th> y las columnas se
//      fusionan (p. ej. "Número de identificación" y "Nombre" desaparecen).
//      Por eso, cuando detectamos HTML, lo decodificamos como Windows-1252 y lo
//      pasamos como texto.
// ==========================================================================
import * as XLSX from "xlsx";

export async function leerLibroExcel(archivo: File): Promise<XLSX.WorkBook> {
  const buffer = await archivo.arrayBuffer();
  const bytes = new Uint8Array(buffer);

  // Miramos el inicio del archivo para saber si es una tabla HTML.
  const prefijo = new TextDecoder("windows-1252")
    .decode(bytes.slice(0, 1024))
    .trim()
    .toLowerCase();
  const esHtml =
    prefijo.startsWith("<") ||
    prefijo.includes("<table") ||
    prefijo.includes("<html") ||
    prefijo.includes("<tr");

  if (esHtml) {
    // Decodificar como Windows-1252 conserva acentos y las etiquetas </td></th>.
    const texto = new TextDecoder("windows-1252").decode(bytes);
    return XLSX.read(texto, { type: "string" });
  }

  // Archivos binarios reales (.xlsx / .xls).
  return XLSX.read(buffer, { type: "array" });
}
