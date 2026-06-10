import ExcelJS from 'exceljs';
import fs from 'node:fs';
import { BRAND } from '../../shared/branding.js';
import { wacc, npv, terminalValue, formatEur } from '../../shared/financial-helpers.js';

const inputPath = process.argv[2] || './input.json';
const input = JSON.parse(fs.readFileSync(inputPath, 'utf8'));

const wb = new ExcelJS.Workbook();
wb.creator = 'M&IA';

// Assumptions sheet
const assumptions = wb.addWorksheet('Assumptions');
assumptions.columns = [
  { header: 'Parameter', key: 'param', width: 30 },
  { header: 'Value', key: 'value', width: 15 },
];
const params = input.assumptions || {
  revenueGrowth: '5%', ebitdaMargin: '20%', taxRate: '25%',
  riskFreeRate: '3%', equityRiskPremium: '5%', beta: 1.0,
  costOfDebt: '4%', debtWeight: '30%', terminalGrowth: '2%',
};
Object.entries(params).forEach(([k, v]) => assumptions.addRow({ param: k, value: v }));

// DCF sheet
const dcf = wb.addWorksheet('DCF');
dcf.columns = [
  { header: '€k', key: 'label', width: 25 },
  ...(input.years || ['Y1', 'Y2', 'Y3', 'Y4', 'Y5']).map((y) => ({ header: y, key: y, width: 14 })),
  { header: 'Terminal', key: 'terminal', width: 14 },
];

const fcfRows = input.fcfRows || [
  { label: 'EBITDA', values: [0, 0, 0, 0, 0] },
  { label: 'D&A', values: [0, 0, 0, 0, 0] },
  { label: 'Capex', values: [0, 0, 0, 0, 0] },
  { label: 'Change in WC', values: [0, 0, 0, 0, 0] },
  { label: 'Free Cash Flow', values: [0, 0, 0, 0, 0] },
];
fcfRows.forEach((row) => {
  const r = { label: row.label };
  (input.years || ['Y1', 'Y2', 'Y3', 'Y4', 'Y5']).forEach((y, i) => {
    r[y] = row.values[i] || 0;
  });
  dcf.addRow(r);
});

dcf.getRow(1).font = { bold: true, color: { argb: 'FF' + BRAND.primary.replace('#', '') } };

const outputName = input.fileName || 'dcf';
await wb.xlsx.writeFile(`./${outputName}.xlsx`);
console.log(`OK: wrote ${outputName}.xlsx`);
