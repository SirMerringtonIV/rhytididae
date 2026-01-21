// scripts/generate-weekly-mdx.js
// Run with: node scripts/generate-weekly-mdx.js

import fs from 'fs';
import path from 'path';
import XLSX from 'xlsx';

// --- CONFIG ---
const excelDir = path.resolve('./data');       // Where your Excel files live
const jsonDir = path.resolve('./src/data');    // Where JSON will be saved
const blogDir = path.resolve('./src/content/blog'); // Where MDX will be saved

const sheetIndex = 2; // 0-based index: 2 = third sheet

// --- ENSURE FOLDERS EXIST ---
if (!fs.existsSync(jsonDir)) fs.mkdirSync(jsonDir, { recursive: true });
if (!fs.existsSync(blogDir)) fs.mkdirSync(blogDir, { recursive: true });

// --- LIST EXCEL FILES ---
const files = fs.existsSync(excelDir) ? fs.readdirSync(excelDir).filter(f => f.endsWith('.xlsx')) : [];
if (files.length === 0) {
  console.log('No Excel files found in', excelDir);
  process.exit(0);
}

console.log('Found Excel files:', files);

// --- PROCESS EACH FILE ---
files.forEach(file => {
  console.log('\nProcessing file:', file);

  const excelPath = path.join(excelDir, file);
  const workbook = XLSX.readFile(excelPath);

  // --- READ THIRD SHEET ONLY ---
  const sheetNames = workbook.SheetNames;
  if (sheetNames.length <= sheetIndex) {
    console.log('  ⚠ Skipping file — it does not have a third sheet.');
    return;
  }

  const sheet = workbook.Sheets[sheetNames[sheetIndex]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  console.log(`  Rows found in third sheet: ${rows.length}`);

  // --- GENERATE JSON ---
  const jsonFileName = file.replace('.xlsx', '.json');
  const jsonPath = path.join(jsonDir, jsonFileName);
  fs.writeFileSync(jsonPath, JSON.stringify(rows, null, 2));
  console.log('  Generated JSON:', jsonPath);

  // --- GENERATE MDX ---
  const today = new Date();
  const yyyy = today.getFullYear();
  const mm = String(today.getMonth() + 1).padStart(2, '0'); // Months 0-11
  const dd = String(today.getDate()).padStart(2, '0');
  const dateStr = `${yyyy}-${mm}-${dd}`;

  // MDX filename: week-YYYY-MM-DD-[originalname].mdx
  const baseName = file.replace('.xlsx', '');
  const mdxFileName = `week-${dateStr}-${baseName}.mdx`;
  const mdxPath = path.join(blogDir, mdxFileName);

  if (!fs.existsSync(mdxPath)) {
    const mdxContent = `---
title: "Weekly Responses - ${dateStr}"
pubDate: "${dateStr}"
---

import Responses from '../../components/Responses.astro';

<Responses file="${jsonFileName}" />
`;
    fs.writeFileSync(mdxPath, mdxContent);
    console.log('  Generated MDX:', mdxPath);
  } else {
    console.log('  MDX already exists:', mdxPath);
  }
});

console.log('\n✅ All done!');
