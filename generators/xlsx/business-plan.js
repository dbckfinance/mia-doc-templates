import ExcelJS from 'exceljs';
import fs from 'node:fs';
import { BRAND } from '../../shared/branding.js';

const inputPath = process.argv[2] || './input.json';
const input = JSON.parse(fs.readFileSync(inputPath, 'utf8'));

const wb = new ExcelJS.Workbook();
wb.creator = 'M&IA';

// Revenue & P&L sheet
const pl = wb.addWorksheet('P&L');
pl.columns = [
  { header: '€k', key: 'label', width: 25 },
  ...(input.years || ['Y1', 'Y2', 'Y3', 'Y4', 'Y5']).map((y) => ({ header: y, key: y, width: 14 })),
];
const rows = input.plRows || [
  { label: 'Revenue', values: [0, 0, 0, 0, 0] },
  { label: 'COGS', values: [0, 0, 0, 0, 0] },
  { label: 'Gross Profit', values: [0, 0, 0, 0, 0] },
  { label: 'EBITDA', values: [0, 0, 0, 0, 0] },
  { label: 'Net Income', values: [0, 0, 0, 0, 0] },
];
rows.forEach((row) => {
  const r = { label: row.label };
  (input.years || ['Y1', 'Y2', 'Y3', 'Y4', 'Y5']).forEach((y, i) => {
    r[y] = row.values[i] || 0;
  });
  pl.addRow(r);
});

// Style header
pl.getRow(1).font = { bold: true, color: { argb: 'FF' + BRAND.primary.replace('#', '') } };

const outputName = input.fileName || 'business-plan';
await wb.xlsx.writeFile(`./${outputName}.xlsx`);
console.log(`OK: wrote ${outputName}.xlsx`);
