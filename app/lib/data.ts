import type ExcelJS from "exceljs";

export type Employee = { matricula: string; nome: string };
export type RecordRow = Employee & { id: number; etiqueta: string; setor: string; data: string; status: "Confirmado" };
export type CompletedBox = { id: string; etiqueta: string; data: string; recordCount: number; exportedAt: string; demo?: boolean };
export type ImportResult = {
  employees: Employee[]; sheetName: string; duplicates: number; conflicts: string[];
  invalidRows: number; sheetsFound: number;
};

export function normalizeHeader(value: string) {
  return value.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().replace(/[^A-Z0-9]+/g, " ").trim();
}

export function normalizeMatricula(value: string) {
  return value.trim().replace(/[^A-Za-z0-9]/g, "").toUpperCase();
}

function cellText(cell: ExcelJS.Cell, identifier = false): string {
  let value = cell.value;
  if (value === null || value === undefined) return "";
  if (typeof value === "object" && "result" in value) value = value.result ?? null;
  if (value === null || value === undefined || typeof value === "boolean" || value instanceof Date) return "";
  if (typeof value === "object") {
    if ("richText" in value) return value.richText.map((part) => part.text).join("").trim();
    if ("text" in value) return value.text.trim();
    return "";
  }
  if (identifier && typeof value === "number") {
    if (!Number.isSafeInteger(value) || value < 0) return "";
    // ExcelJS cell.text does not apply Excel's all-zero number format.
    const format = cell.numFmt?.split(";")[0];
    if (format && /^0{1,32}$/.test(format)) return String(value).padStart(format.length, "0");
  }
  return String(value).trim();
}

export function extractSocEmployees(workbook: ExcelJS.Workbook): ImportResult {
  const candidates: ImportResult[] = [];
  for (const worksheet of workbook.worksheets) {
    for (let rowNumber = 1; rowNumber <= Math.min(50, worksheet.rowCount); rowNumber++) {
      const header = worksheet.getRow(rowNumber);
      let matriculaColumn = 0;
      let nameColumn = 0;
      header.eachCell((cell, column) => {
        const text = normalizeHeader(cellText(cell));
        if (!matriculaColumn && (text.includes("MATRICULA") || ["CHAPA", "CODIGO FUNCIONARIO", "CODIGO DO FUNCIONARIO", "CODIGO EMPREGADO", "CODIGO DO EMPREGADO", "ID FUNCIONARIO"].includes(text))) matriculaColumn = column;
        if (!nameColumn && ["NOME", "NOME COMPLETO", "NOME FUNCIONARIO", "NOME DO FUNCIONARIO", "NOME EMPREGADO", "NOME DO EMPREGADO", "FUNCIONARIO", "EMPREGADO"].includes(text)) nameColumn = column;
      });
      if (!matriculaColumn || !nameColumn) continue;
      const unique = new Map<string, Employee>();
      const conflicts = new Set<string>();
      let duplicates = 0;
      let invalidRows = 0;
      worksheet.eachRow((row, index) => {
        if (index <= rowNumber) return;
        const rawId = cellText(row.getCell(matriculaColumn), true);
        const matricula = normalizeMatricula(rawId);
        const nome = cellText(row.getCell(nameColumn)).replace(/\s+/g, " ").trim();
        if (!rawId && !nome) return;
        if (!matricula || !nome) { invalidRows++; return; }
        if (conflicts.has(matricula)) return;
        const previous = unique.get(matricula);
        if (previous) {
          if (previous.nome.toLocaleUpperCase("pt-BR") === nome.toLocaleUpperCase("pt-BR")) duplicates++;
          else { conflicts.add(matricula); unique.delete(matricula); }
        } else unique.set(matricula, { matricula, nome });
      });
      if (unique.size || conflicts.size) candidates.push({ employees: [...unique.values()], sheetName: worksheet.name, duplicates, conflicts: [...conflicts], invalidRows, sheetsFound: 0 });
      break;
    }
  }
  const result = candidates.sort((a, b) => b.employees.length - a.employees.length)[0];
  if (!result) throw new Error("Não encontrei registros com matrícula e nome. Confira se as duas colunas têm cabeçalhos nas primeiras 50 linhas.");
  if (!result.employees.length) throw new Error("As matrículas encontradas têm nomes conflitantes. Confira a base antes de importar.");
  return { ...result, sheetsFound: candidates.length };
}

export function todayIso() {
  const now = new Date();
  return new Date(now.getTime() - now.getTimezoneOffset() * 60_000).toISOString().slice(0, 10);
}

export function validDate(value: unknown): value is string {
  if (typeof value !== "string" || !/^\d{4}-\d{2}-\d{2}$/.test(value)) return false;
  const date = new Date(`${value}T12:00:00Z`);
  return Number.isFinite(date.getTime()) && date.toISOString().slice(0, 10) === value;
}

export function formatDate(value: string) {
  return validDate(value) ? new Intl.DateTimeFormat("pt-BR", { timeZone: "UTC" }).format(new Date(`${value}T12:00:00Z`)) : "—";
}

export function readHistory(value: string | null): CompletedBox[] {
  try {
    const rows: unknown = JSON.parse(value ?? "[]");
    if (!Array.isArray(rows)) return [];
    return rows.filter((item) => item && typeof item.id === "string" && typeof item.etiqueta === "string" && validDate(item.data) && Number.isSafeInteger(item.recordCount) && item.recordCount > 0 && typeof item.exportedAt === "string" && Number.isFinite(Date.parse(item.exportedAt)) && !item.demo).slice(0, 1000).map(({ id, etiqueta, data, recordCount, exportedAt }) => ({ id, etiqueta, data, recordCount, exportedAt }));
  } catch { return []; }
}

export const DEMO_EMPLOYEES: Employee[] = Array.from({ length: 8 }, (_, index) => ({
  matricula: `DEMO${String(index + 1).padStart(3, "0")}`,
  nome: `Pessoa de demonstração ${String(index + 1).padStart(2, "0")}`,
}));
