// ==========================================================================
// Datos DIVIPOLA — Departamentos y municipios de Colombia.
// Fuente: DANE, Codificación de la División Político-Administrativa (2026).
// Generado automáticamente desde el CSV oficial; no editar a mano.
// ==========================================================================
import rawData from "./divipola.json";

export interface Municipio {
  codigo: string; // 5 dígitos, ej. "05001"
  nombre: string;
}

export interface Departamento {
  departamento: string; // nombre oficial DANE en mayúsculas
  codigo: string;       // 2 dígitos, ej. "05"
  municipios: Municipio[];
}

export const DIVIPOLA: Departamento[] = rawData as Departamento[];

// Índice rápido: nombre departamento → lista de municipios
export const municipiosPorDepto: Record<string, Municipio[]> = Object.fromEntries(
  DIVIPOLA.map((d) => [d.departamento, d.municipios])
);

// Lista de nombres de departamentos ordenados, con los más usados primero
const PRIMEROS = ["ANTIOQUIA", "BOGOTÁ, D.C."];
export const nombresDepartamentos: string[] = [
  ...PRIMEROS,
  ...DIVIPOLA.map((d) => d.departamento).filter((n) => !PRIMEROS.includes(n)),
];

// Departamento y ciudad por defecto según la sede
export const DEFAULT_DEPTO_MEDELLIN = "ANTIOQUIA";
export const DEFAULT_CIUDAD_MEDELLIN = "MEDELLÍN";
export const DEFAULT_DEPTO_BOGOTA = "BOGOTÁ, D.C.";
export const DEFAULT_CIUDAD_BOGOTA = "BOGOTÁ, D.C.";
