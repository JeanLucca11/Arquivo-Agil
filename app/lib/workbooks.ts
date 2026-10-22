import ExcelJS from "exceljs";
import { formatDate, validDate } from "./data.ts";
import type { RecordRow, CompletedBox } from "./data";
export function createBoxWorkbook(records: RecordRow[], label: string, date: string) {
  if (!records.length || !label.trim() || !validDate(date)) throw new Error("Confira a etiqueta, a data e os registros da caixa.");
    const workbook = new ExcelJS.Workbook();
    workbook.creator = "Arquivo Ágil — projeto autoral de Jean";
    workbook.title = `Documentos da caixa ${label}`;
    workbook.subject = "Planilhamento de documentos físicos";
    workbook.created = new Date();

    const worksheet = workbook.addWorksheet("Documentos", {
      views: [{ state: "frozen", ySplit: 4 }],
      pageSetup: { orientation: "landscape", fitToPage: true, fitToWidth: 1 },
    });

    worksheet.mergeCells("A1:E1");
    worksheet.getCell("A1").value = `ARQUIVO ÁGIL  •  CAIXA ${label}`;
    worksheet.getCell("A1").font = { name: "Aptos Display", size: 16, bold: true, color: { argb: "FFFFFFFF" } };
    worksheet.getCell("A1").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0866FF" } };
    worksheet.getCell("A1").alignment = { vertical: "middle", horizontal: "left" };
    worksheet.getRow(1).height = 32;

    worksheet.mergeCells("A2:E2");
    worksheet.getCell("A2").value = `Saúde Ocupacional  •  Data de temporalidade: ${formatDate(date)}  •  ${records.length} registros`;
    worksheet.getCell("A2").font = { name: "Aptos", size: 10, color: { argb: "FF5D5D6A" } };
    worksheet.getCell("A2").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEAF3FF" } };
    worksheet.getCell("A2").alignment = { vertical: "middle", horizontal: "left" };
    worksheet.getRow(2).height = 23;

    worksheet.addTable({
      name: "TabelaDocumentos",
      ref: "A4",
      headerRow: true,
      totalsRow: false,
      style: { theme: "TableStyleMedium2", showRowStripes: true, showColumnStripes: false },
      columns: [
        { name: "Número da etiqueta", filterButton: true },
        { name: "Nome do funcionário", filterButton: true },
        { name: "Matrícula", filterButton: true },
        { name: "Setor", filterButton: true },
        { name: "Data de temporalidade", filterButton: true },
      ],
      rows: records.map((row) => [label.trim(), row.nome, row.matricula, row.setor, formatDate(date)]),
    });

    worksheet.columns = [
      { width: 23 },
      { width: 32 },
      { width: 15 },
      { width: 25 },
      { width: 17 },
    ];

    worksheet.getRow(4).height = 26;
    worksheet.getRow(4).eachCell((cell) => {
      cell.font = { name: "Aptos", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0866FF" } };
      cell.alignment = { vertical: "middle", horizontal: "left" };
      cell.border = { bottom: { style: "thin", color: { argb: "FF004EC2" } } };
    });

    records.forEach((_, index) => {
      const excelRow = worksheet.getRow(index + 5);
      excelRow.height = 23;
      excelRow.eachCell((cell) => {
        cell.font = { name: "Aptos", size: 10, color: { argb: "FF24242B" } };
        cell.alignment = { vertical: "middle", horizontal: "left" };
        cell.border = { bottom: { style: "hair", color: { argb: "FFE4E1EF" } } };
        if (index % 2 === 1) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF4F8FF" } };
      });
    });
    worksheet.getColumn(3).numFmt = "@";


    return workbook;
}
export function createHistoryWorkbook(history: CompletedBox[]) {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = "Arquivo Ágil — projeto autoral de Jean";
    workbook.title = "Relatório de caixas processadas";
    const worksheet = workbook.addWorksheet("Resumo de caixas", { views: [{ state: "frozen", ySplit: 1 }] });
    worksheet.addTable({
      name: "TabelaCaixasProcessadas",
      ref: "A1",
      headerRow: true,
      style: { theme: "TableStyleMedium2", showRowStripes: true },
      columns: [
        { name: "Número da etiqueta", filterButton: true },
        { name: "Data de temporalidade", filterButton: true },
        { name: "Documentos", filterButton: true },
        { name: "Exportado em", filterButton: true },
      ],
      rows: history.map((item) => [item.etiqueta, formatDate(item.data), item.recordCount, new Intl.DateTimeFormat("pt-BR", { dateStyle: "short", timeStyle: "short" }).format(new Date(item.exportedAt))]),
    });
    worksheet.columns = [{ width: 25 }, { width: 18 }, { width: 14 }, { width: 22 }];
    worksheet.getRow(1).height = 26;
    worksheet.getRow(1).eachCell((cell) => {
      cell.font = { name: "Aptos", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0866FF" } };
    });

    return workbook;
}
