import assert from "node:assert/strict";
import test from "node:test";
import ExcelJS from "exceljs";
import { extractSocEmployees, normalizeMatricula, readHistory, validDate } from "../app/lib/data.ts";
import { createBoxWorkbook, createHistoryWorkbook } from "../app/lib/workbooks.ts";

test("importação preserva zeros, consolida repetições e bloqueia nomes conflitantes", async () => {
  const book = new ExcelJS.Workbook();
  const sheet = book.addWorksheet("Headcount");
  sheet.addRow(["Relatório sintético de teste"]);
  sheet.addRow(["Matrícula", "Nome do funcionário"]);
  sheet.addRow(["000123", "Pessoa de teste A"]);
  sheet.addRow(["000123", "Pessoa de teste A"]);
  const padded = sheet.addRow([42, "Pessoa de teste B"]);
  padded.getCell(1).numFmt = "000000";
  sheet.addRow(["CONFLITO", "Pessoa de teste C"]);
  sheet.addRow(["CONFLITO", "Pessoa de teste D"]);
  sheet.addRow(["CONFLITO", "Pessoa de teste C"]);
  sheet.addRow(["SEM-NOME", ""]);
  sheet.addRow([Number.MAX_SAFE_INTEGER + 1, "Identificador numérico sem precisão"]);
  sheet.addRow([{ formula: '"000099"', result: "000099" }, "Pessoa de teste E"]);
  const loaded = new ExcelJS.Workbook();
  await loaded.xlsx.load(await book.xlsx.writeBuffer());
  const result = extractSocEmployees(loaded);
  assert.deepEqual(result.employees.map((employee) => employee.matricula), ["000123", "000042", "000099"]);
  assert.deepEqual(result.conflicts, ["CONFLITO"]);
  assert.equal(result.duplicates, 1);
  assert.equal(result.invalidRows, 2);
});

test("arquivo incompatível retorna erro legível", () => {
  const book = new ExcelJS.Workbook();
  book.addWorksheet("Sem dados").addRow(["Outro campo"]);
  assert.throws(() => extractSocEmployees(book), /matrícula e nome/);
});

test("a aba selecionada e as alternativas são informadas", () => {
  const book = new ExcelJS.Workbook();
  const a = book.addWorksheet("Resumo");
  a.addRows([["Matrícula", "Nome"], ["A", "Exemplo A"]]);
  const b = book.addWorksheet("Base completa");
  b.addRows([["Matrícula", "Nome"], ["B", "Exemplo B"], ["C", "Exemplo C"]]);
  const result = extractSocEmployees(book);
  assert.equal(result.sheetName, "Base completa");
  assert.equal(result.sheetsFound, 2);
  assert.equal(result.employees.length, 2);
});

test("Excel exportado mantém cinco campos, zeros, data e etiqueta corrigidas", async () => {
  const rows = [{ id: 1, etiqueta: "ETIQUETA-ANTIGA", nome: "=nome sintético", matricula: "000123", setor: "Saúde Ocupacional", data: "2025-01-01", status: "Confirmado" }];
  const book = createBoxWorkbook(rows, "  CAIXA-ATUAL  ", "2026-10-15");
  const loaded = new ExcelJS.Workbook();
  await loaded.xlsx.load(await book.xlsx.writeBuffer());
  const sheet = loaded.getWorksheet("Documentos");
  assert.deepEqual(sheet.getRow(5).values.slice(1), ["CAIXA-ATUAL", "=nome sintético", "000123", "Saúde Ocupacional", "15/10/2026"]);
  assert.equal(sheet.getCell("B5").type, ExcelJS.ValueType.String);
  assert.equal(sheet.getCell("C5").numFmt, "@");
  assert.equal(sheet.getTable("TabelaDocumentos").table.columns.length, 5);
  assert.throws(() => createBoxWorkbook([], "CAIXA", "2026-10-15"));
  assert.throws(() => createBoxWorkbook(rows, "CAIXA", "2026-02-30"));
});

test("histórico inválido não quebra a página e remove propriedades pessoais", () => {
  const valid = { id: "1", etiqueta: "CAIXA", data: "2026-10-15", recordCount: 3, exportedAt: "2026-10-15T12:00:00Z" };
  assert.deepEqual(readHistory("{quebrado"), []);
  assert.deepEqual(readHistory(JSON.stringify([null, {}, { ...valid, data: "inválida" }, { ...valid, demo: true }, { ...valid, nome: "não deve persistir", matricula: "123" }])), [valid]);
  const report = createHistoryWorkbook([valid]);
  assert.equal(report.getWorksheet("Resumo de caixas").getCell("C2").value, 3);
  assert.equal(validDate("2026-02-30"), false);
  assert.equal(normalizeMatricula(" 00.123-4 "), "001234");
});

test("importa 70 mil registros e encontra o último sem truncamento", async (t) => {
  const book = new ExcelJS.Workbook();
  const sheet = book.addWorksheet("Headcount sintético 70000");
  sheet.addRow(["Matricula", "Nome"]);
  for (let i = 1; i <= 70_000; i++) sheet.addRow([String(i).padStart(8, "0"), `Pessoa sintética ${i}`]);
  const buffer = await book.xlsx.writeBuffer();
  const loaded = new ExcelJS.Workbook();
  const started = performance.now();
  await loaded.xlsx.load(buffer);
  const result = extractSocEmployees(loaded);
  const index = new Map(result.employees.map((employee) => [employee.matricula, employee]));
  assert.equal(result.employees.length, 70_000);
  assert.equal(index.get("00070000").nome, "Pessoa sintética 70000");
  assert.equal(index.get("00000001").nome, "Pessoa sintética 1");
  t.diagnostic(`70.000 registros sintéticos; ${(buffer.byteLength / 1024 / 1024).toFixed(2)} MB; leitura e índice ${(performance.now() - started).toFixed(0)} ms no ambiente de teste Node.`);
});
