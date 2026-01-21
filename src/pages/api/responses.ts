import type { APIRoute } from 'astro';
import * as XLSX from 'xlsx';
import fs from 'fs';
import path from 'path';

export const get: APIRoute = ({ url }) => {
  const fileParam = url.searchParams.get('file');
  if (!fileParam) {
    return new Response(JSON.stringify({ error: 'No file specified' }), { status: 400 });
  }

  const file = path.resolve('./src/data', fileParam);

  if (!fs.existsSync(file)) {
    return new Response(JSON.stringify({ error: 'File not found' }), { status: 404 });
  }

  const workbook = XLSX.readFile(file);
  const sheet = workbook.Sheets[workbook.SheetNames[2]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  return new Response(JSON.stringify(rows), {
    status: 200,
    headers: { 'Content-Type': 'application/json' },
  });
};