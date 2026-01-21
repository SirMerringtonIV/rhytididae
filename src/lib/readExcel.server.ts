// src/lib/readExcel.server.ts
import * as XLSX from "xlsx";
import fs from "fs";
import path from "path";

/**
 * Reads an Excel file and returns the sheet as an array of arrays.
 * @param file - path to .xlsx relative to project root
 * @param sheetIndex - 0-based index of sheet to read
 */
export function readExcel(file: string, sheetIndex = 2) {
  const filePath = path.resolve("./", file);

  if (!fs.existsSync(filePath)) {
    throw new Error(`Excel file not found: ${filePath}`);
  }

  const workbook = XLSX.readFile(filePath);
  const sheet = workbook.Sheets[workbook.SheetNames[sheetIndex]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  return rows;
}

