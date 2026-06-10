import ExcelJS from 'exceljs';
import fs from 'node:fs';
import { BRAND } from '../../shared/branding.js';

const inputPath = process.argv[2] || './input.json';
const input = JSON.parse(fs.readFileSync(inputPath, 'utf8'));

const wb = new ExcelJS.Workbook();
wb.creator = 'M&IA';

// Sources & Uses
const su = wb.addWorksheet('Sources & Uses');
su.columns = [
  { header: 'Sources', key: 'source', width: 25 },
  { header: '€k', key: 'sourceAmt', width: 14 },
  { header: 'Uses', key: 'use', width: 25 },
  { header: '€k', key: 'useAmt', width: 14 },
];
const sources = input.sources || [
  { source: 'Senior Debt', sourceAmt: 0 },
  { source: 'Mezzanine', sourceAmt: 0 },
  { source: 'Equity', sourceAmt: 0 },
];
const uses = input.uses || [
  { use: 'Enterprise Value', useAmt: 0 },
  { use: 'Transaction Fees', useAmt: 0 },
];
const maxLen = Math.max(sources.length, uses.length);
for (let i = 0; i < maxLen; i++) {
  su.addRow({
    source: sources[i]?.source || '',
    sourceAmt: sources[i]?.sourceAmt || '',
    use: uses[i]?.use || '',
    useAmt: uses[i]?.useAmt || '',
  });
}

// Returns Analysis
const returns = wb.addWorksheet('Returns');
returns.columns = [
  { header: 'Exit Year', key: 'year', width: 12 },
  { header: 'Exit EV (€k)', key: 'ev', width: 15 },
  { header: 'Equity Value (€k)', key: 'equity', width: 18 },
  { header: 'MoM', key: 'mom', width: 8 },
  { header: 'IRR', key: 'irr', width: 8 },
];
(input.exitScenarios || []).forEach((s) => returns.addRow(s));

su.getRow(1).font = { bold: true, color: { argb: 'FF' + BRAND.primary.replace('#', '') } };
returns.getRow(1).font = { bold: true, color: { argb: 'FF' + BRAND.primary.replace('#', '') } };

const outputName = input.fileName || 'lbo';
await wb.xlsx.writeFile(`./${outputName}.xlsx`);
console.log(`OK: wrote ${outputName}.xlsx`);
