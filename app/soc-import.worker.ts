import ExcelJS from "exceljs";
import { extractSocEmployees } from "./lib/data";

self.addEventListener("message", async (event: MessageEvent<ArrayBuffer>) => {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(event.data);
    self.postMessage({ result: extractSocEmployees(workbook) });
  } catch (error) {
    self.postMessage({ error: error instanceof Error ? error.message : "Não foi possível ler o arquivo. Use uma planilha .xlsx sem senha." });
  }
});
