import fs from 'fs';
import path from 'path';
import XLSX from 'xlsx';

// --- CONFIG ---
const excelDir = path.resolve('./data');       // Source Excel files
const jsonDir = path.resolve('./public/data/'); // JSON output for Astro
const blogDir = path.resolve('./src/content/blog'); // MDX output

const sheetIndex = 2; // 0-based: 2 = third sheet

// --- ENSURE FOLDERS EXIST ---
if (!fs.existsSync(jsonDir)) fs.mkdirSync(jsonDir, { recursive: true });
if (!fs.existsSync(blogDir)) fs.mkdirSync(blogDir, { recursive: true });

// --- LIST EXCEL FILES ---
const files = fs.existsSync(excelDir)
  ? fs.readdirSync(excelDir).filter(f => f.endsWith('.xlsx'))
  : [];
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
  
  // --- READ TITLE FROM SHEET 1, CELL B1 ---
  const sheet1 = workbook.Sheets[workbook.SheetNames[0]]; // first sheet
  const titleCell = sheet1?.B1?.v;
  const descriptionCell = sheet1?.B5?.v;

  const title =
    typeof titleCell === 'string' && titleCell.trim()
      ? titleCell.trim()
      : null;
	  
  const description =
    typeof descriptionCell === 'string' && descriptionCell.trim()
      ? descriptionCell.trim()
      : '';

  // --- READ THIRD SHEET ONLY ---
  const sheetNames = workbook.SheetNames;
  if (sheetNames.length <= sheetIndex) {
    console.log('  ⚠ Skipping file — it does not have a third sheet.');
    return;
  }

  const sheet = workbook.Sheets[sheetNames[sheetIndex]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  console.log(`  Rows found in third sheet: ${rows.length}`);

  // --- GENERATE JSON IN /public/data ---
  const jsonFileName = file.replace('.xlsx', '.json');
  const jsonPath = path.join(jsonDir, jsonFileName);
  fs.writeFileSync(jsonPath, JSON.stringify(rows, null, 2));
  console.log('  Generated JSON:', jsonPath);

  // --- GENERATE MDX ---
  const today = new Date();
  const yyyy = today.getFullYear();
  const mm = String(today.getMonth() + 1).padStart(2, '0');
  const dd = String(today.getDate()).padStart(2, '0');
  const dateStr = `${yyyy}-${mm}-${dd}`;

  const baseName = file.replace('.xlsx', '');
  const mdxFileName = `${baseName}.mdx`;
  const mdxPath = path.join(blogDir, mdxFileName);

  if (!fs.existsSync(mdxPath)) {
    const mdxContent = `---
title: "${title ?? `Weekly Responses - ${dateStr}`}"
description: "${description}"
pubDate: "${dateStr}"
heroImage: ""
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