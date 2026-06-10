// shared/xlsx-model.js
// Declarative Excel workbook composer (exceljs) with M&IA styling.
//
// buildWorkbook(spec) → Buffer where spec = {
//   title, confidential?, sheets: [
//     { name, columns?: [{ header, key, width }],
//       headerRows?: [[..]],          // banner rows above the table
//       table?: { headers, rows },    // simple data table
//       cells?: [{ ref, value, formula?, style? }],   // free-form
//       rows?: [[..]],                // raw rows appended in order
//       sectionTitle?,
//       freezeHeader? }
//   ]
// }

import ExcelJS from 'exceljs';
import { BRAND, argb } from './branding.js';

const HEADER_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: argb(BRAND.navy) } };
const BAND_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: argb(BRAND.lightGray) } };
const TOTAL_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: argb(BRAND.lightBlue) } };
const THIN_BORDER = {
  top: { style: 'thin', color: { argb: 'FFCCCCCC' } },
  bottom: { style: 'thin', color: { argb: 'FFCCCCCC' } },
  left: { style: 'thin', color: { argb: 'FFCCCCCC' } },
  right: { style: 'thin', color: { argb: 'FFCCCCCC' } },
};

export const FMT = {
  number: '#,##0.0',
  integer: '#,##0',
  currency: '#,##0.0\\ "€"',
  currencyK: '#,##0,\\ "k€"',
  currencyM: '#,##0.0,,\\ "m€"',
  percent: '0.0%',
  multiple: '0.0"x"',
  date: 'dd/mm/yyyy',
};

function styleTitleRow(row) {
  row.font = { name: BRAND.font, size: 14, bold: true, color: { argb: argb(BRAND.navy) } };
  row.height = 22;
}

function styleHeaderRow(row) {
  row.eachCell((cell) => {
    cell.font = { name: BRAND.font, size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = HEADER_FILL;
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = THIN_BORDER;
  });
  row.height = 20;
}

function styleDataRow(row, index, { totalRow = false } = {}) {
  row.eachCell((cell, colNumber) => {
    cell.font = {
      name: BRAND.font, size: 10,
      bold: totalRow || colNumber === 1,
      color: { argb: typeof cell.value === 'number' && cell.value < 0 ? argb(BRAND.negative) : argb(BRAND.text) },
    };
    if (totalRow) cell.fill = TOTAL_FILL;
    else if (index % 2 === 1) cell.fill = BAND_FILL;
    cell.border = THIN_BORDER;
    cell.alignment = { vertical: 'middle', horizontal: colNumber === 1 ? 'left' : 'right' };
    if (typeof cell.value === 'number') cell.numFmt = Number.isInteger(cell.value) && Math.abs(cell.value) >= 1000 ? FMT.integer : FMT.number;
  });
}

export async function buildWorkbook(spec) {
  const wb = new ExcelJS.Workbook();
  wb.creator = BRAND.brandName;
  wb.created = new Date();

  for (const sheetSpec of spec.sheets || []) {
    const ws = wb.addWorksheet(sheetSpec.name.slice(0, 31), {
      views: sheetSpec.freezeHeader ? [{ state: 'frozen', ySplit: sheetSpec.tableStartRow || 3 }] : undefined,
      properties: { defaultColWidth: 14 },
    });

    // Banner: title + confidentiality
    const titleRow = ws.addRow([sheetSpec.sectionTitle || spec.title || sheetSpec.name]);
    styleTitleRow(titleRow);
    if (spec.confidential !== false) {
      const conf = ws.addRow([BRAND.confidentialityFr]);
      conf.font = { name: BRAND.font, size: 9, italic: true, color: { argb: argb(BRAND.negative) } };
    }
    ws.addRow([]);

    if (sheetSpec.columns?.length) {
      sheetSpec.columns.forEach((col, i) => {
        ws.getColumn(i + 1).width = col.width || 16;
      });
    }

    for (const banner of sheetSpec.headerRows || []) {
      const r = ws.addRow(banner);
      r.font = { name: BRAND.font, size: 11, bold: true, color: { argb: argb(BRAND.darkBlue) } };
    }

    if (sheetSpec.table) {
      const headerRow = ws.addRow(sheetSpec.table.headers);
      styleHeaderRow(headerRow);
      const totalIdx = sheetSpec.table.totalRowIndex;
      sheetSpec.table.rows.forEach((row, i) => {
        const r = ws.addRow(row);
        styleDataRow(r, i, { totalRow: totalIdx != null && i === totalIdx });
      });
      if (!sheetSpec.columns?.length) {
        // Auto width from headers
        sheetSpec.table.headers.forEach((h, i) => {
          ws.getColumn(i + 1).width = Math.max(14, String(h).length + 6);
        });
        ws.getColumn(1).width = Math.max(26, ws.getColumn(1).width || 0);
      }
    }

    for (const row of sheetSpec.rows || []) {
      const r = ws.addRow(row);
      r.font = { name: BRAND.font, size: 10, color: { argb: argb(BRAND.text) } };
    }

    for (const c of sheetSpec.cells || []) {
      const cell = ws.getCell(c.ref);
      if (c.formula) cell.value = { formula: c.formula };
      else cell.value = c.value;
      if (c.numFmt) cell.numFmt = c.numFmt;
      if (c.style) Object.assign(cell, c.style);
      if (c.bold || c.color || c.size) {
        cell.font = {
          name: BRAND.font,
          size: c.size || 10,
          bold: Boolean(c.bold),
          color: { argb: argb(c.color || BRAND.text) },
        };
      }
    }
  }

  const buffer = await wb.xlsx.writeBuffer();
  return Buffer.from(buffer);
}

/** Build an assumptions sheet spec from a key/value object. */
export function assumptionsSheet(assumptions, { name = 'Assumptions', title = 'Hypothèses' } = {}) {
  return {
    name,
    sectionTitle: title,
    table: {
      headers: ['Hypothèse', 'Valeur'],
      rows: Object.entries(assumptions).map(([k, v]) => [k, v]),
    },
    columns: [{ width: 40 }, { width: 22 }],
  };
}
