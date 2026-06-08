/**
 * Alphabet Inc. (GOOGL) — 5-Year DCF Valuation Model
 * Investment-banking grade Excel model with interconnected formulas.
 * Data source: Alphabet FY2025 10-K / Q4 2025 earnings release (Feb 2026)
 */
import ExcelJS from 'exceljs';
import fs from 'node:fs';
import path from 'node:path';

const OUTPUT = path.resolve('artifacts/Alphabet_Inc_DCF_5Y_Model-xlsx.xlsx');

// ── Colour palette (investment banking conventions) ──────────────────────────
const C = {
  input: 'FF0070C0',       // blue  — hard inputs
  formula: 'FF000000',     // black — same-sheet formulas
  crossRef: 'FF008000',    // green — cross-sheet references
  header: 'FF0A2540',      // navy section headers
  headerFont: 'FFFFFFFF',
  flag: 'FFFF0000',        // red flags
  ok: 'FF008000',
  sensitivity: 'FF1F4E79',
  lightFill: 'FFF2F7FB',
  altFill: 'FFE8F0FE',
};

const FONT = 'Calibri';
const MODEL_DATE = 'June 8, 2026';

// ── FY2025A base-year inputs ($ millions unless noted) ───────────────────────
// Source: Alphabet Q4/FY2025 earnings release & 10-K (Dec 31, 2025)
const INPUTS = {
  company: 'Alphabet Inc.',
  ticker: 'GOOGL / GOOG',
  baseYear: 'FY2025A',
  baseRevenue: 402836,
  // Revenue growth — decelerating from 15% FY25 actual toward mature tech rate
  revGrowth: [0.12, 0.10, 0.09, 0.08, 0.07],
  operatingMargin: 0.32,
  daPctRevenue: 0.0525,      // D&A $21,136M / Revenue $402,836M
  capexPctRevenue: [0.20, 0.19, 0.18, 0.17, 0.16], // tapering from elevated AI capex
  nwcPctRevenue: -0.0584,    // negative NWC typical for Alphabet
  taxRate: 0.21,
  sharesOutstanding: 12230,  // diluted, millions
  cashAndEquivalents: 126843,
  totalDebt: 72000,
  minorityInterests: 0,
  terminalGrowth: 0.025,
  // WACC inputs
  riskFreeRate: 0.04,
  equityRiskPremium: 0.055,
  beta: 1.05,
  preTaxCostOfDebt: 0.05,
  targetEquityWeight: 0.95,
  targetDebtWeight: 0.05,
  midYearDiscounting: 1, // 1 = Yes, 0 = No
};

const FORECAST_YEARS = [2026, 2027, 2028, 2029, 2030];
const WACC_RANGE = [];
for (let w = 0.07; w <= 0.1101; w += 0.005) WACC_RANGE.push(Math.round(w * 1000) / 1000);
const G_RANGE = [];
for (let g = 0.015; g <= 0.0351; g += 0.005) G_RANGE.push(Math.round(g * 1000) / 1000);

// ── Helpers ──────────────────────────────────────────────────────────────────
function styleCell(cell, { font, fill, numFmt, bold, align, border } = {}) {
  if (font) cell.font = { name: FONT, size: font.size ?? 11, bold: bold ?? font.bold ?? false, color: { argb: font.color ?? C.formula } };
  if (fill) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fill } };
  if (numFmt) cell.numFmt = numFmt;
  if (align) cell.alignment = align;
  if (border) cell.border = border;
}

function sectionHeader(ws, row, col, text, span = 6) {
  const cell = ws.getCell(row, col);
  cell.value = text;
  styleCell(cell, { font: { color: C.headerFont, size: 12, bold: true }, fill: C.header, bold: true });
  if (span > 1) ws.mergeCells(row, col, row, col + span - 1);
}

function inputCell(ws, row, col, value, numFmt, note) {
  const cell = ws.getCell(row, col);
  cell.value = value;
  styleCell(cell, { font: { color: C.input }, numFmt });
  if (note) {
    const n = ws.getCell(row, col + 2);
    n.value = note;
    styleCell(n, { font: { color: 'FF666666', size: 9 } });
  }
  return cell;
}

function formulaCell(ws, row, col, formula, numFmt, crossRef = false) {
  const cell = ws.getCell(row, col);
  cell.value = { formula };
  styleCell(cell, { font: { color: crossRef ? C.crossRef : C.formula }, numFmt });
  return cell;
}

function labelCell(ws, row, col, text, bold = false) {
  const cell = ws.getCell(row, col);
  cell.value = text;
  styleCell(cell, { font: { color: C.formula, bold }, bold });
}

const thinBorder = {
  top: { style: 'thin', color: { argb: 'FFD0D0D0' } },
  bottom: { style: 'thin', color: { argb: 'FFD0D0D0' } },
  left: { style: 'thin', color: { argb: 'FFD0D0D0' } },
  right: { style: 'thin', color: { argb: 'FFD0D0D0' } },
};

const FMT_USD_M = '$#,##0;($#,##0);"-"';
const FMT_USD_B = '$#,##0.0,,"B";($#,##0.0,,"B");"-"';
const FMT_PCT = '0.0%';
const FMT_PCT2 = '0.00%';
const FMT_SHARE = '$#,##0.00';
const FMT_MULT = '0.0"x"';

// ── Build workbook ───────────────────────────────────────────────────────────
const wb = new ExcelJS.Workbook();
wb.creator = 'M&IA DCF Model';
wb.created = new Date();
wb.modified = new Date();

// ══════════════════════════════════════════════════════════════════════════════
// 1. COVER
// ══════════════════════════════════════════════════════════════════════════════
const cover = wb.addWorksheet('Cover');
cover.views = [{ state: 'frozen', ySplit: 1 }];
cover.getColumn(1).width = 4;
cover.getColumn(2).width = 38;
cover.getColumn(3).width = 22;

sectionHeader(cover, 2, 2, 'DISCOUNTED CASH FLOW VALUATION MODEL', 2);
cover.mergeCells(3, 2, 3, 3);
const titleCell = cover.getCell(3, 2);
titleCell.value = INPUTS.company;
styleCell(titleCell, { font: { color: C.header, size: 22, bold: true }, bold: true, align: { horizontal: 'left' } });

labelCell(cover, 5, 2, 'Ticker:'); cover.getCell(5, 3).value = INPUTS.ticker;
labelCell(cover, 6, 2, 'Model Date:'); cover.getCell(6, 3).value = MODEL_DATE;
labelCell(cover, 7, 2, 'Base Year:'); cover.getCell(7, 3).value = INPUTS.baseYear;
labelCell(cover, 8, 2, 'Forecast Horizon:'); cover.getCell(8, 3).value = '5 Years (FY2026E – FY2030E)';
labelCell(cover, 9, 2, 'Currency:'); cover.getCell(9, 3).value = 'USD ($ millions)';

sectionHeader(cover, 12, 2, 'MODEL NAVIGATION', 2);
const nav = [
  ['Inputs', 'Historical base year & operating assumptions'],
  ['WACC', 'CAPM-based weighted average cost of capital'],
  ['Forecast', 'Unlevered free cash flow build'],
  ['Discounting', 'Period discount factors & PV of UFCFs'],
  ['Terminal Value', 'Gordon Growth perpetuity method'],
  ['Valuation Summary', 'EV → Equity Value bridge & implied multiples'],
  ['Sensitivity', 'Two-way tables: WACC vs. terminal growth'],
  ['Checks', 'Sanity checks & model flags'],
];
nav.forEach(([tab, desc], i) => {
  const r = 13 + i;
  const c1 = cover.getCell(r, 2);
  c1.value = tab;
  styleCell(c1, { font: { color: C.input, bold: true }, bold: true });
  cover.getCell(r, 3).value = desc;
});

sectionHeader(cover, 23, 2, 'DISCLAIMER', 2);
cover.mergeCells(24, 2, 28, 3);
const disc = cover.getCell(24, 2);
disc.value = 'This model is prepared for illustrative purposes only and does not constitute investment advice. '
  + 'Projections are based on management-reported financials (FY2025A) and analyst assumptions that may differ '
  + 'from company guidance. The recipient should conduct independent due diligence before making any investment decision. '
  + 'Strictly Confidential — Not for distribution.';
disc.alignment = { wrapText: true, vertical: 'top' };
styleCell(disc, { font: { color: 'FF666666', size: 9 } });

sectionHeader(cover, 30, 2, 'COLOUR LEGEND', 2);
[
  ['Blue font', 'Hard-coded inputs (editable)'],
  ['Black font', 'Formulas (same sheet)'],
  ['Green font', 'Cross-sheet references'],
].forEach(([k, v], i) => {
  const r = 31 + i;
  styleCell(cover.getCell(r, 2), { font: { color: k.includes('Blue') ? C.input : k.includes('Green') ? C.crossRef : C.formula, bold: true }, bold: true });
  cover.getCell(r, 2).value = k;
  cover.getCell(r, 3).value = v;
});

// ══════════════════════════════════════════════════════════════════════════════
// 2. INPUTS
// ══════════════════════════════════════════════════════════════════════════════
const inp = wb.addWorksheet('Inputs');
inp.views = [{ state: 'frozen', xSplit: 2, ySplit: 3 }];
[inp.getColumn(1), inp.getColumn(2), inp.getColumn(3), inp.getColumn(4), inp.getColumn(5)].forEach((c, i) => {
  c.width = i === 1 ? 34 : i === 4 ? 18 : 16;
});

sectionHeader(inp, 1, 1, 'ALPHABET INC. — MODEL INPUTS', 5);
labelCell(inp, 2, 1, 'Parameter', true);
labelCell(inp, 2, 2, 'Value', true);
labelCell(inp, 2, 4, 'Source / Notes', true);

// Row map (fixed for cross-references)
const R = {
  baseYear: 4,
  baseRevenue: 5,
  revGrowthStart: 7,   // rows 7-11
  opMargin: 13,
  daPct: 14,
  capexStart: 16,      // rows 16-20
  nwcPct: 22,
  taxRate: 24,
  shares: 26,
  cash: 28,
  debt: 29,
  minority: 30,
  netDebt: 31,
  terminalG: 33,
};

labelCell(inp, R.baseYear, 1, 'Base Year (Actual)');
inputCell(inp, R.baseYear, 2, INPUTS.baseYear);

labelCell(inp, R.baseRevenue, 1, 'Base Year Revenue ($M)');
inputCell(inp, R.baseRevenue, 2, INPUTS.baseRevenue, FMT_USD_M, 'FY2025 consolidated revenue');

labelCell(inp, 6, 1, 'Revenue Growth Assumptions', true);
FORECAST_YEARS.forEach((yr, i) => {
  const r = R.revGrowthStart + i;
  labelCell(inp, r, 1, `  FY${yr}E Revenue Growth`);
  inputCell(inp, r, 2, INPUTS.revGrowth[i], FMT_PCT);
});

labelCell(inp, R.opMargin, 1, 'Operating Margin (EBIT %)');
inputCell(inp, R.opMargin, 2, INPUTS.operatingMargin, FMT_PCT, 'FY2025 operating margin 32.0%');

labelCell(inp, R.daPct, 1, 'D&A (% of Revenue)');
inputCell(inp, R.daPct, 2, INPUTS.daPctRevenue, FMT_PCT, 'FY2025 D&A $21,136M');

labelCell(inp, 15, 1, 'Capex (% of Revenue)', true);
FORECAST_YEARS.forEach((yr, i) => {
  const r = R.capexStart + i;
  labelCell(inp, r, 1, `  FY${yr}E Capex %`);
  inputCell(inp, r, 2, INPUTS.capexPctRevenue[i], FMT_PCT, 'Tapering from elevated AI infrastructure spend');
});

labelCell(inp, R.nwcPct, 1, 'Net Working Capital (% of Revenue)');
inputCell(inp, R.nwcPct, 2, INPUTS.nwcPctRevenue, FMT_PCT, 'Negative NWC — cash-collection model');

labelCell(inp, R.taxRate, 1, 'Corporate Tax Rate');
inputCell(inp, R.taxRate, 2, INPUTS.taxRate, FMT_PCT, 'U.S. federal statutory assumption');

labelCell(inp, R.shares, 1, 'Diluted Shares Outstanding (M)');
inputCell(inp, R.shares, 2, INPUTS.sharesOutstanding, '#,##0.0', 'FY2025 weighted-average diluted');

labelCell(inp, R.cash, 1, 'Cash & Marketable Securities ($M)');
inputCell(inp, R.cash, 2, INPUTS.cashAndEquivalents, FMT_USD_M, 'FY2025 balance sheet');

labelCell(inp, R.debt, 1, 'Total Financial Debt ($M)');
inputCell(inp, R.debt, 2, INPUTS.totalDebt, FMT_USD_M, 'Long-term debt & commercial paper');

labelCell(inp, R.minority, 1, 'Minority Interests ($M)');
inputCell(inp, R.minority, 2, INPUTS.minorityInterests, FMT_USD_M, 'N/A — fully consolidated');

labelCell(inp, R.netDebt, 1, 'Net Debt / (Net Cash) ($M)');
formulaCell(inp, R.netDebt, 2, `B${R.debt}-B${R.cash}`, FMT_USD_M, true);

labelCell(inp, R.terminalG, 1, 'Terminal Growth Rate (g)');
inputCell(inp, R.terminalG, 2, INPUTS.terminalGrowth, FMT_PCT2, 'Perpetuity growth assumption');

// ══════════════════════════════════════════════════════════════════════════════
// 3. WACC
// ══════════════════════════════════════════════════════════════════════════════
const wacc = wb.addWorksheet('WACC');
wacc.views = [{ state: 'frozen', xSplit: 2, ySplit: 3 }];
[wacc.getColumn(1), wacc.getColumn(2), wacc.getColumn(3)].forEach((c, i) => { c.width = i === 0 ? 34 : 16; });

sectionHeader(wacc, 1, 1, 'WEIGHTED AVERAGE COST OF CAPITAL (CAPM)', 3);

const WR = {
  rf: 4,
  erp: 5,
  beta: 6,
  costEquity: 7,
  preTaxDebt: 9,
  taxRate: 10,
  afterTaxDebt: 11,
  eqWeight: 13,
  debtWeight: 14,
  wacc: 15,
  midYearToggle: 17,
  midYearLabel: 18,
};

labelCell(wacc, WR.rf, 1, 'Risk-Free Rate (Rf)');
inputCell(wacc, WR.rf, 2, INPUTS.riskFreeRate, FMT_PCT2);

labelCell(wacc, WR.erp, 1, 'Equity Risk Premium (ERP)');
inputCell(wacc, WR.erp, 2, INPUTS.equityRiskPremium, FMT_PCT2);

labelCell(wacc, WR.beta, 1, 'Levered Beta (β)');
inputCell(wacc, WR.beta, 2, INPUTS.beta, '0.00');

labelCell(wacc, WR.costEquity, 1, 'Cost of Equity (Ke = Rf + β × ERP)');
formulaCell(wacc, WR.costEquity, 2, `B${WR.rf}+B${WR.beta}*B${WR.erp}`, FMT_PCT2);

labelCell(wacc, WR.preTaxDebt, 1, 'Pre-Tax Cost of Debt (Kd)');
inputCell(wacc, WR.preTaxDebt, 2, INPUTS.preTaxCostOfDebt, FMT_PCT2);

labelCell(wacc, WR.taxRate, 1, 'Tax Rate (for Kd after-tax)');
formulaCell(wacc, WR.taxRate, 2, `Inputs!B${R.taxRate}`, FMT_PCT, true);

labelCell(wacc, WR.afterTaxDebt, 1, 'After-Tax Cost of Debt (Kd × (1-T))');
formulaCell(wacc, WR.afterTaxDebt, 2, `B${WR.preTaxDebt}*(1-B${WR.taxRate})`, FMT_PCT2);

labelCell(wacc, WR.eqWeight, 1, 'Target Equity Weight (E / (D+E))');
inputCell(wacc, WR.eqWeight, 2, INPUTS.targetEquityWeight, FMT_PCT);

labelCell(wacc, WR.debtWeight, 1, 'Target Debt Weight (D / (D+E))');
inputCell(wacc, WR.debtWeight, 2, INPUTS.targetDebtWeight, FMT_PCT);

labelCell(wacc, WR.wacc, 1, 'WACC', true);
formulaCell(wacc, WR.wacc, 2, `B${WR.eqWeight}*B${WR.costEquity}+B${WR.debtWeight}*B${WR.afterTaxDebt}`, FMT_PCT2, false);
styleCell(wacc.getCell(WR.wacc, 2), { font: { color: C.formula, bold: true, size: 12 }, bold: true, fill: C.lightFill });

labelCell(wacc, WR.midYearToggle, 1, 'Mid-Year Discounting (1=Yes, 0=No)');
inputCell(wacc, WR.midYearToggle, 2, INPUTS.midYearDiscounting, '0');

labelCell(wacc, WR.midYearLabel, 1, 'Discount Exponent Adjustment');
formulaCell(wacc, WR.midYearLabel, 2, `IF(B${WR.midYearToggle}=1,0.5,0)`, '0.0');

// ══════════════════════════════════════════════════════════════════════════════
// 4. FORECAST (FCF Build)
// ══════════════════════════════════════════════════════════════════════════════
const fc = wb.addWorksheet('Forecast');
fc.views = [{ state: 'frozen', xSplit: 2, ySplit: 4 }];
fc.getColumn(1).width = 32;
FORECAST_YEARS.forEach((_, i) => { fc.getColumn(i + 2).width = 14; });
fc.getColumn(7).width = 14; // base year col

sectionHeader(fc, 1, 1, 'UNLEVERED FREE CASH FLOW BUILD', FORECAST_YEARS.length + 2);
labelCell(fc, 3, 1, '$ Millions', true);
FORECAST_YEARS.forEach((yr, i) => {
  const c = fc.getCell(3, i + 2);
  c.value = `FY${yr}E`;
  styleCell(c, { font: { color: C.headerFont, bold: true }, fill: C.header, bold: true, align: { horizontal: 'center' } });
});
styleCell(fc.getCell(3, 7), { font: { color: C.headerFont, bold: true }, fill: C.header, bold: true, align: { horizontal: 'center' } });
fc.getCell(3, 7).value = INPUTS.baseYear;

const FR = {
  revenue: 5,
  growth: 6,
  ebit: 7,
  opMargin: 8,
  tax: 9,
  nopat: 10,
  da: 11,
  capex: 12,
  nwc: 13,
  deltaNwc: 14,
  ufcf: 15,
  fcfMargin: 16,
};

// Revenue
labelCell(fc, FR.revenue, 1, 'Revenue', true);
formulaCell(fc, FR.revenue, 2, `Inputs!B${R.baseRevenue}*(1+Inputs!B${R.revGrowthStart})`, FMT_USD_M, true);
for (let i = 1; i < 5; i++) {
  formulaCell(fc, FR.revenue, i + 2, `${fc.getCell(FR.revenue, i + 1).address}*(1+Inputs!B${R.revGrowthStart + i})`, FMT_USD_M, true);
}
formulaCell(fc, FR.revenue, 7, `Inputs!B${R.baseRevenue}`, FMT_USD_M, true);

// Growth rates display
labelCell(fc, FR.growth, 1, '  YoY Revenue Growth');
FORECAST_YEARS.forEach((_, i) => {
  formulaCell(fc, FR.growth, i + 2, `Inputs!B${R.revGrowthStart + i}`, FMT_PCT, true);
});
fc.getCell(FR.growth, 7).value = '—';

// EBIT (Operating Income)
labelCell(fc, FR.ebit, 1, 'EBIT (Operating Income)', true);
for (let i = 0; i < 5; i++) {
  formulaCell(fc, FR.ebit, i + 2, `${fc.getCell(FR.revenue, i + 2).address}*Inputs!B${R.opMargin}`, FMT_USD_M, true);
}
formulaCell(fc, FR.ebit, 7, `${fc.getCell(FR.revenue, 7).address}*Inputs!B${R.opMargin}`, FMT_USD_M, true);

labelCell(fc, FR.opMargin, 1, '  Operating Margin');
for (let i = 0; i < 5; i++) {
  formulaCell(fc, FR.opMargin, i + 2, `Inputs!B${R.opMargin}`, FMT_PCT, true);
}
formulaCell(fc, FR.opMargin, 7, `Inputs!B${R.opMargin}`, FMT_PCT, true);

// Taxes on EBIT
labelCell(fc, FR.tax, 1, 'Less: Taxes on EBIT');
for (let i = 0; i < 5; i++) {
  formulaCell(fc, FR.tax, i + 2, `${fc.getCell(FR.ebit, i + 2).address}*Inputs!B${R.taxRate}`, FMT_USD_M, true);
}
formulaCell(fc, FR.tax, 7, `${fc.getCell(FR.ebit, 7).address}*Inputs!B${R.taxRate}`, FMT_USD_M, true);

// NOPAT
labelCell(fc, FR.nopat, 1, 'NOPAT', true);
for (let i = 0; i < 5; i++) {
  formulaCell(fc, FR.nopat, i + 2, `${fc.getCell(FR.ebit, i + 2).address}-${fc.getCell(FR.tax, i + 2).address}`, FMT_USD_M, true);
}
formulaCell(fc, FR.nopat, 7, `${fc.getCell(FR.ebit, 7).address}-${fc.getCell(FR.tax, 7).address}`, FMT_USD_M, true);

// D&A
labelCell(fc, FR.da, 1, 'Plus: D&A');
for (let i = 0; i < 5; i++) {
  formulaCell(fc, FR.da, i + 2, `${fc.getCell(FR.revenue, i + 2).address}*Inputs!B${R.daPct}`, FMT_USD_M, true);
}
formulaCell(fc, FR.da, 7, `${fc.getCell(FR.revenue, 7).address}*Inputs!B${R.daPct}`, FMT_USD_M, true);

// Capex (negative)
labelCell(fc, FR.capex, 1, 'Less: Capex');
for (let i = 0; i < 5; i++) {
  formulaCell(fc, FR.capex, i + 2, `-${fc.getCell(FR.revenue, i + 2).address}*Inputs!B${R.capexStart + i}`, FMT_USD_M, true);
}
formulaCell(fc, FR.capex, 7, `-${fc.getCell(FR.revenue, 7).address}*Inputs!B${R.capexStart}`, FMT_USD_M, true);

// NWC level
labelCell(fc, FR.nwc, 1, 'Net Working Capital');
formulaCell(fc, FR.nwc, 2, `${fc.getCell(FR.revenue, 2).address}*Inputs!B${R.nwcPct}`, FMT_USD_M, true);
for (let i = 1; i < 5; i++) {
  formulaCell(fc, FR.nwc, i + 2, `${fc.getCell(FR.revenue, i + 2).address}*Inputs!B${R.nwcPct}`, FMT_USD_M, true);
}
formulaCell(fc, FR.nwc, 7, `${fc.getCell(FR.revenue, 7).address}*Inputs!B${R.nwcPct}`, FMT_USD_M, true);

// ΔNWC (negative = source of cash)
labelCell(fc, FR.deltaNwc, 1, 'Less: Δ Net Working Capital');
formulaCell(fc, FR.deltaNwc, 2, `-(${fc.getCell(FR.nwc, 2).address}-${fc.getCell(FR.nwc, 7).address})`, FMT_USD_M, true);
for (let i = 1; i < 5; i++) {
  formulaCell(fc, FR.deltaNwc, i + 2, `-(${fc.getCell(FR.nwc, i + 2).address}-${fc.getCell(FR.nwc, i + 1).address})`, FMT_USD_M, true);
}
fc.getCell(FR.deltaNwc, 7).value = '—';

// UFCF
labelCell(fc, FR.ufcf, 1, 'Unlevered Free Cash Flow (UFCF)', true);
for (let i = 0; i < 5; i++) {
  const col = i + 2;
  formulaCell(fc, FR.ufcf, col,
    `${fc.getCell(FR.nopat, col).address}+${fc.getCell(FR.da, col).address}+${fc.getCell(FR.capex, col).address}+${fc.getCell(FR.deltaNwc, col).address}`,
    FMT_USD_M, true);
  styleCell(fc.getCell(FR.ufcf, col), { fill: C.altFill, font: { color: C.crossRef, bold: true }, bold: true });
}
formulaCell(fc, FR.ufcf, 7,
  `${fc.getCell(FR.nopat, 7).address}+${fc.getCell(FR.da, 7).address}+${fc.getCell(FR.capex, 7).address}`,
  FMT_USD_M, true);

labelCell(fc, FR.fcfMargin, 1, '  UFCF Margin (% of Revenue)');
for (let i = 0; i < 5; i++) {
  formulaCell(fc, FR.fcfMargin, i + 2, `IF(${fc.getCell(FR.revenue, i + 2).address}=0,0,${fc.getCell(FR.ufcf, i + 2).address}/${fc.getCell(FR.revenue, i + 2).address})`, FMT_PCT, true);
}

// ══════════════════════════════════════════════════════════════════════════════
// 5. DISCOUNTING
// ══════════════════════════════════════════════════════════════════════════════
const disc2 = wb.addWorksheet('Discounting');
disc2.views = [{ state: 'frozen', xSplit: 2, ySplit: 4 }];
disc2.getColumn(1).width = 32;
FORECAST_YEARS.forEach((_, i) => { disc2.getColumn(i + 2).width = 14; });

sectionHeader(disc2, 1, 1, 'DISCOUNTING — PV OF EXPLICIT PERIOD UFCFs', 6);

labelCell(disc2, 3, 1, '', true);
FORECAST_YEARS.forEach((yr, i) => {
  disc2.getCell(3, i + 2).value = `FY${yr}E`;
  styleCell(disc2.getCell(3, i + 2), { font: { color: C.headerFont, bold: true }, fill: C.header, bold: true, align: { horizontal: 'center' } });
});

const DR = {
  period: 5,
  wacc: 6,
  midAdj: 7,
  discountFactor: 8,
  ufcf: 9,
  pvUfcf: 10,
  sumPv: 12,
};

labelCell(disc2, DR.period, 1, 'Forecast Period (Year #)');
FORECAST_YEARS.forEach((_, i) => { disc2.getCell(DR.period, i + 2).value = i + 1; });

labelCell(disc2, DR.wacc, 1, 'WACC');
FORECAST_YEARS.forEach((_, i) => {
  formulaCell(disc2, DR.wacc, i + 2, `WACC!B${WR.wacc}`, FMT_PCT2, true);
});

labelCell(disc2, DR.midAdj, 1, 'Mid-Year Adjustment');
FORECAST_YEARS.forEach((_, i) => {
  formulaCell(disc2, DR.midAdj, i + 2, `WACC!B${WR.midYearLabel}`, '0.0', true);
});

labelCell(disc2, DR.discountFactor, 1, 'Discount Factor', true);
FORECAST_YEARS.forEach((_, i) => {
  const period = i + 1;
  formulaCell(disc2, DR.discountFactor, i + 2,
    `1/(1+${disc2.getCell(DR.wacc, i + 2).address})^(${period}-${disc2.getCell(DR.midAdj, i + 2).address})`,
    '0.0000', true);
});

labelCell(disc2, DR.ufcf, 1, 'UFCF ($M)');
FORECAST_YEARS.forEach((_, i) => {
  formulaCell(disc2, DR.ufcf, i + 2, `Forecast!${fc.getCell(FR.ufcf, i + 2).address}`, FMT_USD_M, true);
});

labelCell(disc2, DR.pvUfcf, 1, 'PV of UFCF ($M)', true);
FORECAST_YEARS.forEach((_, i) => {
  formulaCell(disc2, DR.pvUfcf, i + 2,
    `${disc2.getCell(DR.ufcf, i + 2).address}*${disc2.getCell(DR.discountFactor, i + 2).address}`,
    FMT_USD_M, true);
  styleCell(disc2.getCell(DR.pvUfcf, i + 2), { fill: C.altFill });
});

labelCell(disc2, DR.sumPv, 1, 'Sum of PV of Explicit UFCFs', true);
formulaCell(disc2, DR.sumPv, 2, `SUM(B${DR.pvUfcf}:F${DR.pvUfcf})`, FMT_USD_M, true);
styleCell(disc2.getCell(DR.sumPv, 2), { font: { color: C.crossRef, bold: true, size: 12 }, bold: true, fill: C.lightFill });

// ══════════════════════════════════════════════════════════════════════════════
// 6. TERMINAL VALUE
// ══════════════════════════════════════════════════════════════════════════════
const tv = wb.addWorksheet('Terminal Value');
tv.views = [{ state: 'frozen', xSplit: 1, ySplit: 3 }];
tv.getColumn(1).width = 38;
tv.getColumn(2).width = 18;

sectionHeader(tv, 1, 1, 'TERMINAL VALUE — GORDON GROWTH METHOD', 2);

const TR = {
  ufcfY5: 4,
  terminalG: 5,
  wacc: 6,
  tv: 7,
  midAdj: 8,
  period: 9,
  pvTv: 10,
};

labelCell(tv, TR.ufcfY5, 1, 'UFCF — Year 5 ($M)');
formulaCell(tv, TR.ufcfY5, 2, `Forecast!${fc.getCell(FR.ufcf, 6).address}`, FMT_USD_M, true);

labelCell(tv, TR.terminalG, 1, 'Terminal Growth Rate (g)');
formulaCell(tv, TR.terminalG, 2, `Inputs!B${R.terminalG}`, FMT_PCT2, true);

labelCell(tv, TR.wacc, 1, 'WACC');
formulaCell(tv, TR.wacc, 2, `WACC!B${WR.wacc}`, FMT_PCT2, true);

labelCell(tv, TR.tv, 1, 'Terminal Value at End of Year 5 ($M)', true);
formulaCell(tv, TR.tv, 2, `B${TR.ufcfY5}*(1+B${TR.terminalG})/(B${TR.wacc}-B${TR.terminalG})`, FMT_USD_M, true);
styleCell(tv.getCell(TR.tv, 2), { fill: C.altFill, font: { bold: true } });

labelCell(tv, TR.midAdj, 1, 'Mid-Year Adjustment');
formulaCell(tv, TR.midAdj, 2, `WACC!B${WR.midYearLabel}`, '0.0', true);

labelCell(tv, TR.period, 1, 'Discount Period (Years)');
tv.getCell(TR.period, 2).value = 5;

labelCell(tv, TR.pvTv, 1, 'PV of Terminal Value ($M)', true);
formulaCell(tv, TR.pvTv, 2, `B${TR.tv}/(1+B${TR.wacc})^(B${TR.period}-B${TR.midAdj})`, FMT_USD_M, true);
styleCell(tv.getCell(TR.pvTv, 2), { font: { color: C.crossRef, bold: true, size: 12 }, bold: true, fill: C.lightFill });

// ══════════════════════════════════════════════════════════════════════════════
// 7. VALUATION SUMMARY
// ══════════════════════════════════════════════════════════════════════════════
const vs = wb.addWorksheet('Valuation Summary');
vs.views = [{ state: 'frozen', xSplit: 1, ySplit: 3 }];
vs.getColumn(1).width = 38;
vs.getColumn(2).width = 18;
vs.getColumn(3).width = 18;

sectionHeader(vs, 1, 1, 'VALUATION SUMMARY', 3);

const VR = {
  pvExplicit: 4,
  pvTerminal: 5,
  enterpriseValue: 6,
  cash: 8,
  debt: 9,
  minority: 10,
  equityValue: 11,
  shares: 13,
  pricePerShare: 14,
  // Multiples
  evEbitdaEntry: 17,
  evRevEntry: 18,
  evEbitdaTerminal: 19,
  evRevTerminal: 20,
  ebitdaEntry: 22,
  ebitdaTerminal: 23,
};

labelCell(vs, VR.pvExplicit, 1, 'PV of Explicit Period UFCFs ($M)');
formulaCell(vs, VR.pvExplicit, 2, `Discounting!B${DR.sumPv}`, FMT_USD_M, true);

labelCell(vs, VR.pvTerminal, 1, 'PV of Terminal Value ($M)');
formulaCell(vs, VR.pvTerminal, 2, `'Terminal Value'!B${TR.pvTv}`, FMT_USD_M, true);

labelCell(vs, VR.enterpriseValue, 1, 'Enterprise Value ($M)', true);
formulaCell(vs, VR.enterpriseValue, 2, `B${VR.pvExplicit}+B${VR.pvTerminal}`, FMT_USD_M, true);
styleCell(vs.getCell(VR.enterpriseValue, 2), { font: { bold: true, size: 12 }, bold: true, fill: C.lightFill });

sectionHeader(vs, 7, 1, 'BRIDGE TO EQUITY VALUE', 2);

labelCell(vs, VR.cash, 1, '(+) Cash & Marketable Securities ($M)');
formulaCell(vs, VR.cash, 2, `Inputs!B${R.cash}`, FMT_USD_M, true);

labelCell(vs, VR.debt, 1, '(-) Total Debt ($M)');
formulaCell(vs, VR.debt, 2, `Inputs!B${R.debt}`, FMT_USD_M, true);

labelCell(vs, VR.minority, 1, '(-) Minority Interests ($M)');
formulaCell(vs, VR.minority, 2, `Inputs!B${R.minority}`, FMT_USD_M, true);

labelCell(vs, VR.equityValue, 1, 'Equity Value ($M)', true);
formulaCell(vs, VR.equityValue, 2, `B${VR.enterpriseValue}+B${VR.cash}-B${VR.debt}-B${VR.minority}`, FMT_USD_M, true);
styleCell(vs.getCell(VR.equityValue, 2), { font: { bold: true, size: 12 }, bold: true, fill: C.altFill });

labelCell(vs, VR.shares, 1, 'Diluted Shares Outstanding (M)');
formulaCell(vs, VR.shares, 2, `Inputs!B${R.shares}`, '#,##0.0', true);

labelCell(vs, VR.pricePerShare, 1, 'Equity Value Per Share', true);
formulaCell(vs, VR.pricePerShare, 2, `B${VR.equityValue}/B${VR.shares}`, FMT_SHARE, true);
styleCell(vs.getCell(VR.pricePerShare, 2), { font: { color: C.input, bold: true, size: 14 }, bold: true, fill: 'FFFFF2CC' });

sectionHeader(vs, 16, 1, 'IMPLIED MULTIPLES', 3);
labelCell(vs, 16, 2, 'Entry Year (FY2026E)', true);
labelCell(vs, 16, 3, 'Terminal Year (FY2030E)', true);

labelCell(vs, VR.evEbitdaEntry, 1, 'EV / EBITDA');
labelCell(vs, VR.evRevEntry, 1, 'EV / Revenue');
labelCell(vs, VR.evEbitdaTerminal, 1, 'EV / EBITDA');
labelCell(vs, VR.evRevTerminal, 1, 'EV / Revenue');

// EBITDA helper rows (hidden-style below)
labelCell(vs, VR.ebitdaEntry, 1, 'EBITDA — Entry Year ($M)');
formulaCell(vs, VR.ebitdaEntry, 2, `Forecast!${fc.getCell(FR.ebit, 2).address}+Forecast!${fc.getCell(FR.da, 2).address}`, FMT_USD_M, true);

labelCell(vs, VR.ebitdaTerminal, 1, 'EBITDA — Terminal Year ($M)');
formulaCell(vs, VR.ebitdaTerminal, 2, `Forecast!${fc.getCell(FR.ebit, 6).address}+Forecast!${fc.getCell(FR.da, 6).address}`, FMT_USD_M, true);

formulaCell(vs, VR.evEbitdaEntry, 2, `IF(B${VR.ebitdaEntry}=0,0,B${VR.enterpriseValue}/B${VR.ebitdaEntry})`, FMT_MULT, true);
formulaCell(vs, VR.evRevEntry, 2, `IF(Forecast!${fc.getCell(FR.revenue, 2).address}=0,0,B${VR.enterpriseValue}/Forecast!${fc.getCell(FR.revenue, 2).address})`, FMT_MULT, true);
formulaCell(vs, VR.evEbitdaTerminal, 2, `IF(B${VR.ebitdaTerminal}=0,0,B${VR.enterpriseValue}/B${VR.ebitdaTerminal})`, FMT_MULT, true);
formulaCell(vs, VR.evRevTerminal, 2, `IF(Forecast!${fc.getCell(FR.revenue, 6).address}=0,0,B${VR.enterpriseValue}/Forecast!${fc.getCell(FR.revenue, 6).address})`, FMT_MULT, true);

// Fix multiples row labels
vs.getCell(VR.evEbitdaEntry, 1).value = 'EV / EBITDA';
vs.getCell(VR.evRevEntry, 1).value = 'EV / Revenue';
// Duplicate labels removed — use row 17-18 properly
vs.getCell(17, 1).value = 'EV / EBITDA';
vs.getCell(18, 1).value = 'EV / Revenue';
formulaCell(vs, 17, 2, `IF(B${VR.ebitdaEntry}=0,0,B${VR.enterpriseValue}/B${VR.ebitdaEntry})`, FMT_MULT, true);
formulaCell(vs, 17, 3, `IF(B${VR.ebitdaTerminal}=0,0,B${VR.enterpriseValue}/B${VR.ebitdaTerminal})`, FMT_MULT, true);
formulaCell(vs, 18, 2, `IF(Forecast!${fc.getCell(FR.revenue, 2).address}=0,0,B${VR.enterpriseValue}/Forecast!${fc.getCell(FR.revenue, 2).address})`, FMT_MULT, true);
formulaCell(vs, 18, 3, `IF(Forecast!${fc.getCell(FR.revenue, 6).address}=0,0,B${VR.enterpriseValue}/Forecast!${fc.getCell(FR.revenue, 6).address})`, FMT_MULT, true);

// ══════════════════════════════════════════════════════════════════════════════
// 8. SENSITIVITY
// ══════════════════════════════════════════════════════════════════════════════
const sens = wb.addWorksheet('Sensitivity');
sens.views = [{ state: 'frozen', xSplit: 2, ySplit: 5 }];
sens.getColumn(1).width = 14;
G_RANGE.forEach((_, i) => { sens.getColumn(i + 3).width = 12; });

sectionHeader(sens, 1, 1, 'SENSITIVITY ANALYSIS — EQUITY VALUE PER SHARE', G_RANGE.length + 2);
sens.mergeCells(2, 1, 2, G_RANGE.length + 2);
sens.getCell(2, 1).value = 'WACC (rows) vs. Terminal Growth Rate g (columns)  |  $ per share';
styleCell(sens.getCell(2, 1), { font: { color: 'FF666666', size: 10 } });

// Build sensitivity formula for share price given WACC (wCell) and g (gCell)
function sharePriceFormula(wRef, gRef) {
  // Sum PV of UFCFs with alternate WACC
  const pvParts = FORECAST_YEARS.map((_, i) => {
    const ufcfRef = `Forecast!${fc.getCell(FR.ufcf, i + 2).address}`;
    const period = i + 1;
    const midRef = `WACC!$B$${WR.midYearLabel}`;
    return `${ufcfRef}/(1+${wRef})^(${period}-${midRef})`;
  });
  const pvSum = pvParts.join('+');
  const ufcfY5 = `Forecast!${fc.getCell(FR.ufcf, 6).address}`;
  const midRef = `WACC!$B$${WR.midYearLabel}`;
  const tvFormula = `${ufcfY5}*(1+${gRef})/(${wRef}-${gRef})`;
  const pvTv = `(${tvFormula})/(1+${wRef})^(5-${midRef})`;
  const ev = `(${pvSum})+(${pvTv})`;
  const equity = `(${ev})+Inputs!$B$${R.cash}-Inputs!$B$${R.debt}-Inputs!$B$${R.minority}`;
  return `(${equity})/Inputs!$B$${R.shares}`;
}

// Header row for g values
labelCell(sens, 4, 1, 'WACC \\ g', true);
labelCell(sens, 4, 2, '', true);
G_RANGE.forEach((g, i) => {
  const cell = sens.getCell(4, i + 3);
  cell.value = g;
  styleCell(cell, { font: { color: C.input, bold: true }, numFmt: FMT_PCT2, bold: true, fill: C.lightFill, align: { horizontal: 'center' } });
});

// Data rows
WACC_RANGE.forEach((w, ri) => {
  const row = 5 + ri;
  const wCell = sens.getCell(row, 2);
  wCell.value = w;
  styleCell(wCell, { font: { color: C.input, bold: true }, numFmt: FMT_PCT2, bold: true, fill: C.lightFill, align: { horizontal: 'center' } });
  sens.getCell(row, 1).value = '';

  G_RANGE.forEach((g, ci) => {
    const gCol = sens.getCell(4, ci + 3).address;
    const wAddr = wCell.address;
    const formula = sharePriceFormula(wAddr, gCol);
    const cell = sens.getCell(row, ci + 3);
    cell.value = { formula };
    styleCell(cell, { font: { color: C.formula }, numFmt: FMT_SHARE, align: { horizontal: 'center' }, border: thinBorder });
    if (ri % 2 === 0) styleCell(cell, { fill: C.lightFill });
  });
});

// Highlight base case
const baseWaccIdx = WACC_RANGE.findIndex((w) => Math.abs(w - 0.095) < 0.001) >= 0
  ? WACC_RANGE.findIndex((w) => Math.abs(w - 0.095) < 0.001)
  : 5;
const baseGIdx = G_RANGE.findIndex((g) => Math.abs(g - 0.025) < 0.001);
if (baseGIdx >= 0) {
  const baseCell = sens.getCell(5 + baseWaccIdx, 3 + baseGIdx);
  styleCell(baseCell, { font: { color: C.input, bold: true }, bold: true, fill: 'FFFFF2CC', border: {
    top: { style: 'medium', color: { argb: C.input } },
    bottom: { style: 'medium', color: { argb: C.input } },
    left: { style: 'medium', color: { argb: C.input } },
    right: { style: 'medium', color: { argb: C.input } },
  } });
}

// Second table — EV sensitivity
const evStartRow = 5 + WACC_RANGE.length + 3;
sectionHeader(sens, evStartRow, 1, 'SENSITIVITY ANALYSIS — ENTERPRISE VALUE ($M)', G_RANGE.length + 2);

function evFormula(wRef, gRef) {
  const pvParts = FORECAST_YEARS.map((_, i) => {
    const ufcfRef = `Forecast!${fc.getCell(FR.ufcf, i + 2).address}`;
    const period = i + 1;
    const midRef = `WACC!$B$${WR.midYearLabel}`;
    return `${ufcfRef}/(1+${wRef})^(${period}-${midRef})`;
  });
  const pvSum = pvParts.join('+');
  const ufcfY5 = `Forecast!${fc.getCell(FR.ufcf, 6).address}`;
  const midRef = `WACC!$B$${WR.midYearLabel}`;
  const tvFormula = `${ufcfY5}*(1+${gRef})/(${wRef}-${gRef})`;
  const pvTv = `(${tvFormula})/(1+${wRef})^(5-${midRef})`;
  return `(${pvSum})+(${pvTv})`;
}

const evHeaderRow = evStartRow + 2;
labelCell(sens, evHeaderRow, 1, 'WACC \\ g', true);
G_RANGE.forEach((g, i) => {
  const cell = sens.getCell(evHeaderRow, i + 3);
  cell.value = g;
  styleCell(cell, { font: { color: C.input, bold: true }, numFmt: FMT_PCT2, bold: true, fill: C.lightFill, align: { horizontal: 'center' } });
});

WACC_RANGE.forEach((w, ri) => {
  const row = evHeaderRow + 1 + ri;
  const wCell = sens.getCell(row, 2);
  wCell.value = w;
  styleCell(wCell, { font: { color: C.input, bold: true }, numFmt: FMT_PCT2, bold: true, fill: C.lightFill, align: { horizontal: 'center' } });

  G_RANGE.forEach((g, ci) => {
    const gCol = sens.getCell(evHeaderRow, ci + 3).address;
    const formula = evFormula(wCell.address, gCol);
    const cell = sens.getCell(row, ci + 3);
    cell.value = { formula };
    styleCell(cell, { font: { color: C.formula }, numFmt: FMT_USD_M, align: { horizontal: 'center' }, border: thinBorder });
    if (ri % 2 === 0) styleCell(cell, { fill: C.lightFill });
  });
});

// ══════════════════════════════════════════════════════════════════════════════
// 9. CHECKS
// ══════════════════════════════════════════════════════════════════════════════
const chk = wb.addWorksheet('Checks');
chk.views = [{ state: 'frozen', xSplit: 1, ySplit: 3 }];
chk.getColumn(1).width = 42;
chk.getColumn(2).width = 16;
chk.getColumn(3).width = 12;
chk.getColumn(4).width = 36;

sectionHeader(chk, 1, 1, 'MODEL SANITY CHECKS', 4);
labelCell(chk, 3, 1, 'Check', true);
labelCell(chk, 3, 2, 'Value', true);
labelCell(chk, 3, 3, 'Status', true);
labelCell(chk, 3, 4, 'Threshold / Notes', true);

const CR = {
  fcfMargin: 4,
  tvPctEv: 5,
  waccCheck: 6,
  waccGtG: 7,
  eqWeight: 8,
  evPositive: 9,
};

// UFCF margin Y5
labelCell(chk, CR.fcfMargin, 1, 'UFCF Margin — Year 5');
formulaCell(chk, CR.fcfMargin, 2, `Forecast!${fc.getCell(FR.fcfMargin, 6).address}`, FMT_PCT, true);
formulaCell(chk, CR.fcfMargin, 3, `IF(AND(B${CR.fcfMargin}>=5%,B${CR.fcfMargin}<=45%),"OK","FLAG")`, '@', true);
chk.getCell(CR.fcfMargin, 4).value = 'Expected range: 5% – 45%';

// TV % of EV
labelCell(chk, CR.tvPctEv, 1, 'Terminal Value % of Enterprise Value');
formulaCell(chk, CR.tvPctEv, 2, `'Terminal Value'!B${TR.pvTv}/'Valuation Summary'!B${VR.enterpriseValue}`, FMT_PCT, true);
formulaCell(chk, CR.tvPctEv, 3, `IF(AND(B${CR.tvPctEv}>=40%,B${CR.tvPctEv}<=95%),"OK","FLAG")`, '@', true);
chk.getCell(CR.tvPctEv, 4).value = 'Typical range: 40% – 95% of EV';

// WACC reasonableness
labelCell(chk, CR.waccCheck, 1, 'WACC within reasonable range');
formulaCell(chk, CR.waccCheck, 2, `WACC!B${WR.wacc}`, FMT_PCT2, true);
formulaCell(chk, CR.waccCheck, 3, `IF(AND(B${CR.waccCheck}>=6%,B${CR.waccCheck}<=14%),"OK","FLAG")`, '@', true);
chk.getCell(CR.waccCheck, 4).value = 'Expected range: 6% – 14%';

// WACC > g
labelCell(chk, CR.waccGtG, 1, 'WACC > Terminal Growth (g)');
formulaCell(chk, CR.waccGtG, 2, `WACC!B${WR.wacc}-Inputs!B${R.terminalG}`, FMT_PCT2, true);
formulaCell(chk, CR.waccGtG, 3, `IF(B${CR.waccGtG}>0,"OK","FLAG")`, '@', true);
chk.getCell(CR.waccGtG, 4).value = 'Must be positive for Gordon Growth validity';

// Capital structure weights sum to 1
labelCell(chk, CR.eqWeight, 1, 'Capital Structure Weights Sum to 100%');
formulaCell(chk, CR.eqWeight, 2, `WACC!B${WR.eqWeight}+WACC!B${WR.debtWeight}`, FMT_PCT, true);
formulaCell(chk, CR.eqWeight, 3, `IF(ABS(B${CR.eqWeight}-100%)<0.1%,"OK","FLAG")`, '@', true);
chk.getCell(CR.eqWeight, 4).value = 'E weight + D weight = 100%';

// EV positive
labelCell(chk, CR.evPositive, 1, 'Enterprise Value > 0');
formulaCell(chk, CR.evPositive, 2, `'Valuation Summary'!B${VR.enterpriseValue}`, FMT_USD_M, true);
formulaCell(chk, CR.evPositive, 3, `IF(B${CR.evPositive}>0,"OK","FLAG")`, '@', true);
chk.getCell(CR.evPositive, 4).value = 'EV must be positive';

// Conditional formatting for status column
for (let r = CR.fcfMargin; r <= CR.evPositive; r++) {
  const statusCell = chk.getCell(r, 3);
  // colour applied via formula display — user sees OK/FLAG
}

// ── Write file ───────────────────────────────────────────────────────────────
fs.mkdirSync(path.dirname(OUTPUT), { recursive: true });
await wb.xlsx.writeFile(OUTPUT);

const stats = fs.statSync(OUTPUT);
console.log(`OK: wrote ${OUTPUT} (${stats.size} bytes)`);
if (stats.size < 5120) {
  console.error('WARNING: file size below 5 KB threshold');
  process.exit(1);
}
