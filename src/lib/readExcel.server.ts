import * as XLSX from "xlsx";
import fs from "fs";
import path from "path";

export function readExcel(file: string, sheetIndex = 2) {
  const filePath = path.resolve("./", file);

  if (!fs.existsSync(filePath)) {
    throw new Error(`Excel file not found: ${filePath}`);
  }

  const workbook = XLSX.readFile(filePath);
  const sheet = workbook.Sheets[workbook.SheetNames[sheetIndex]];
  return XLSX.utils.sheet_to_json(sheet, { header: 1 });
}

