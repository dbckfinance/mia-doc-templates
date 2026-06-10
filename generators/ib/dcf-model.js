// generators/ib/dcf-model.js
// Modèle DCF complet (xlsx) avec formules Excel vivantes.

import ExcelJS from 'exceljs';
import { runCli, isCliInvocation } from '../../shared/cli.js';
import { BRAND, argb } from '../../shared/branding.js';
import { FMT } from '../../shared/xlsx-model.js';

export const metadata = {
  id: 'dcf-model',
  name: 'Modèle DCF',
  vertical: 'ib',
  outputType: 'xlsx',
  estimatedPages: '4-6 sheets',
  requiredInput: ['company'],
  optionalInput: ['revenue0', 'growthRates', 'ebitdaMargin', 'taxRate', 'capexPct', 'nwcPct', 'wacc', 'terminalGrowth', 'netDebt', 'shares'],
};

const HEADER_STYLE = (cell) => {
  cell.font = { name: BRAND.font, size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: argb(BRAND.navy) } };
  cell.alignment = { horizontal: 'center', vertical: 'middle' };
};

const INPUT_STYLE = (cell) => {
  cell.font = { name: BRAND.font, size: 10, color: { argb: 'FF1F4E99' } };
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
};

export async function generate(input = {}) {
  const company = input.company || 'Société';
  const years = input.projectionYears || 5;
  const growth = input.growthRates || Array(years).fill(0.05);
  const a = {
    revenue0: input.revenue0 ?? 100,
    ebitdaMargin: input.ebitdaMargin ?? 0.2,
    daPct: input.daPct ?? 0.04,
    taxRate: input.taxRate ?? 0.25,
    capexPct: input.capexPct ?? 0.05,
    nwcPct: input.nwcPct ?? 0.01,
    wacc: input.wacc ?? 0.09,
    terminalGrowth: input.terminalGrowth ?? 0.02,
    netDebt: input.netDebt ?? 0,
    shares: input.shares ?? null,
  };

  const wb = new ExcelJS.Workbook();
  wb.creator = BRAND.brandName;

  // ---- Assumptions ----
  const wsA = wb.addWorksheet('Hypothèses');
  wsA.getColumn(1).width = 40;
  wsA.getColumn(2).width = 18;
  wsA.addRow([`DCF — ${company}`]).font = { name: BRAND.font, size: 14, bold: true, color: { argb: argb(BRAND.navy) } };
  wsA.addRow([BRAND.confidentialityFr]).font = { name: BRAND.font, size: 9, italic: true, color: { argb: argb(BRAND.negative) } };
  wsA.addRow([]);
  const assumptionRows = [
    ['Chiffre d\'affaires année 0 (m€)', a.revenue0, FMT.number],
    ['Marge EBITDA', a.ebitdaMargin, FMT.percent],
    ['D&A (% CA)', a.daPct, FMT.percent],
    ['Taux d\'impôt', a.taxRate, FMT.percent],
    ['Capex (% CA)', a.capexPct, FMT.percent],
    ['Variation BFR (% CA)', a.nwcPct, FMT.percent],
    ['WACC', a.wacc, FMT.percent],
    ['Croissance à l\'infini (g)', a.terminalGrowth, FMT.percent],
    ['Dette nette (m€)', a.netDebt, FMT.number],
  ];
  for (const [label, value, fmt] of assumptionRows) {
    const r = wsA.addRow([label, value]);
    r.getCell(1).font = { name: BRAND.font, size: 10, bold: true, color: { argb: argb(BRAND.navy) } };
    INPUT_STYLE(r.getCell(2));
    r.getCell(2).numFmt = fmt;
  }
  // Named refs (row offsets: assumptions start row 4)
  const REF = {
    revenue0: 'Hypothèses!$B$4',
    ebitdaMargin: 'Hypothèses!$B$5',
    daPct: 'Hypothèses!$B$6',
    taxRate: 'Hypothèses!$B$7',
    capexPct: 'Hypothèses!$B$8',
    nwcPct: 'Hypothèses!$B$9',
    wacc: 'Hypothèses!$B$10',
    g: 'Hypothèses!$B$11',
    netDebt: 'Hypothèses!$B$12',
  };

  // ---- FCF projection ----
  const ws = wb.addWorksheet('DCF');
  ws.getColumn(1).width = 34;
  for (let c = 2; c <= years + 1; c++) ws.getColumn(c).width = 14;
  ws.addRow([`Projection des flux — ${company} (m€)`]).font = { name: BRAND.font, size: 13, bold: true, color: { argb: argb(BRAND.navy) } };
  ws.addRow([]);

  const headerRow = ws.addRow(['', ...Array.from({ length: years }, (_, i) => `Année ${i + 1}`)]);
  headerRow.eachCell((cell, col) => { if (col > 1) HEADER_STYLE(cell); });

  const col = (i) => String.fromCharCode(66 + i); // B, C, D…
  const growthRow = ws.addRow(['Croissance CA', ...growth]);
  growthRow.eachCell((cell, c) => { if (c > 1) { INPUT_STYLE(cell); cell.numFmt = FMT.percent; } });
  const gRowN = growthRow.number;

  const rows = [
    { label: "Chiffre d'affaires", f: (i) => i === 0
        ? `${REF.revenue0}*(1+${col(0)}${gRowN})`
        : `${col(i - 1)}${gRowN + 1}*(1+${col(i)}${gRowN})`, fmt: FMT.number },
    { label: 'EBITDA', f: (i) => `${col(i)}${gRowN + 1}*${REF.ebitdaMargin}`, fmt: FMT.number },
    { label: 'D&A', f: (i) => `-${col(i)}${gRowN + 1}*${REF.daPct}`, fmt: FMT.number },
    { label: 'EBIT', f: (i) => `${col(i)}${gRowN + 2}+${col(i)}${gRowN + 3}`, fmt: FMT.number },
    { label: 'Impôt sur EBIT', f: (i) => `-${col(i)}${gRowN + 4}*${REF.taxRate}`, fmt: FMT.number },
    { label: 'NOPAT', f: (i) => `${col(i)}${gRowN + 4}+${col(i)}${gRowN + 5}`, fmt: FMT.number },
    { label: '(+) D&A', f: (i) => `-${col(i)}${gRowN + 3}`, fmt: FMT.number },
    { label: '(-) Capex', f: (i) => `-${col(i)}${gRowN + 1}*${REF.capexPct}`, fmt: FMT.number },
    { label: '(-) Var. BFR', f: (i) => `-${col(i)}${gRowN + 1}*${REF.nwcPct}`, fmt: FMT.number },
    { label: 'Free Cash Flow', f: (i) => `${col(i)}${gRowN + 6}+${col(i)}${gRowN + 7}+${col(i)}${gRowN + 8}+${col(i)}${gRowN + 9}`, fmt: FMT.number, bold: true },
    { label: 'Facteur d\'actualisation', f: (i) => `1/(1+${REF.wacc})^${i + 1}`, fmt: '0.000' },
    { label: 'FCF actualisé', f: (i) => `${col(i)}${gRowN + 10}*${col(i)}${gRowN + 11}`, fmt: FMT.number, bold: true },
  ];

  for (const spec of rows) {
    const r = ws.addRow([spec.label]);
    r.getCell(1).font = { name: BRAND.font, size: 10, bold: Boolean(spec.bold), color: { argb: argb(BRAND.navy) } };
    for (let i = 0; i < years; i++) {
      const cell = r.getCell(i + 2);
      cell.value = { formula: spec.f(i) };
      cell.numFmt = spec.fmt;
      cell.font = { name: BRAND.font, size: 10, bold: Boolean(spec.bold) };
    }
  }

  const fcfRowN = gRowN + 10;
  const dfcfRowN = gRowN + 12;
  const lastCol = col(years - 1);

  // ---- Valuation ----
  ws.addRow([]);
  const valStart = ws.lastRow.number + 1;
  const valRows = [
    ['Somme des FCF actualisés', `SUM(${col(0)}${dfcfRowN}:${lastCol}${dfcfRowN})`],
    ['Valeur terminale (Gordon)', `${lastCol}${fcfRowN}*(1+${REF.g})/(${REF.wacc}-${REF.g})`],
    ['VT actualisée', `B${valStart + 1}*${lastCol}${gRowN + 11}`],
    ["Valeur d'entreprise", `B${valStart}+B${valStart + 2}`],
    ['(-) Dette nette', `-${REF.netDebt}`],
    ['Valeur des fonds propres', `B${valStart + 3}+B${valStart + 4}`],
  ];
  if (a.shares) valRows.push(['Valeur par action (€)', `B${valStart + 5}/${a.shares}`]);
  for (const [label, formula] of valRows) {
    const r = ws.addRow([label]);
    r.getCell(1).font = { name: BRAND.font, size: 10, bold: true, color: { argb: argb(BRAND.darkBlue) } };
    const cell = r.getCell(2);
    cell.value = { formula };
    cell.numFmt = FMT.number;
    cell.font = { name: BRAND.font, size: 10, bold: true };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: argb(BRAND.lightBlue) } };
  }

  // ---- Sensitivity (computed values, WACC x g) ----
  const wsS = wb.addWorksheet('Sensibilités');
  wsS.getColumn(1).width = 16;
  wsS.addRow(['Sensibilité VE (m€) — WACC vs croissance terminale']).font = { name: BRAND.font, size: 12, bold: true, color: { argb: argb(BRAND.navy) } };
  wsS.addRow([]);
  const waccRange = [-0.01, -0.005, 0, 0.005, 0.01].map((d) => a.wacc + d);
  const gRange = [-0.005, -0.0025, 0, 0.0025, 0.005].map((d) => a.terminalGrowth + d);
  const hdr = wsS.addRow(['WACC \\ g', ...gRange]);
  hdr.eachCell((cell, c) => { HEADER_STYLE(cell); if (c > 1) cell.numFmt = FMT.percent; });

  // numeric simulation for sensitivity values
  const simulate = (w, g) => {
    let rev = a.revenue0;
    let sumDfcf = 0;
    let fcfLast = 0;
    for (let i = 0; i < years; i++) {
      rev *= 1 + (growth[i] ?? growth[growth.length - 1] ?? 0.03);
      const ebitda = rev * a.ebitdaMargin;
      const da = rev * a.daPct;
      const ebit = ebitda - da;
      const nopat = ebit * (1 - a.taxRate);
      const fcf = nopat + da - rev * a.capexPct - rev * a.nwcPct;
      sumDfcf += fcf / Math.pow(1 + w, i + 1);
      fcfLast = fcf;
    }
    if (w <= g) return null;
    const tv = (fcfLast * (1 + g)) / (w - g);
    return sumDfcf + tv / Math.pow(1 + w, years);
  };
  for (const w of waccRange) {
    const row = wsS.addRow([w, ...gRange.map((g) => {
      const v = simulate(w, g);
      return v == null ? '—' : Math.round(v * 10) / 10;
    })]);
    row.getCell(1).numFmt = FMT.percent;
    row.getCell(1).font = { name: BRAND.font, size: 10, bold: true };
    row.eachCell((cell, c) => { if (c > 1 && typeof cell.value === 'number') cell.numFmt = FMT.number; });
  }

  const buffer = await wb.xlsx.writeBuffer();
  return Buffer.from(buffer);
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
