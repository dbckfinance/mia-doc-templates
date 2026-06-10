// shared/table-builder.js
// Financial table builders for docx (variance highlighting, zebra rows,
// header shading) and shared formatting logic for xlsx.

import {
  Table, TableRow, TableCell, Paragraph, TextRun,
  AlignmentType, WidthType, ShadingType, BorderStyle,
} from 'docx';
import { BRAND } from './branding.js';
import { formatNumber } from './financial-helpers.js';

function cell(text, opts = {}) {
  return new TableCell({
    children: [new Paragraph({
      children: [new TextRun({
        text: String(text ?? '—'),
        size: opts.size || 18,
        font: BRAND.font,
        bold: opts.bold || false,
        color: opts.color || BRAND.text,
        italics: opts.italic || false,
      })],
      alignment:
        opts.align === 'right' ? AlignmentType.RIGHT
          : opts.align === 'center' ? AlignmentType.CENTER
            : AlignmentType.LEFT,
    })],
    shading: opts.shading ? { type: ShadingType.SOLID, color: opts.shading } : undefined,
    verticalAlign: 'center',
    width: opts.width ? { size: opts.width, type: WidthType.PERCENTAGE } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
  });
}

const THIN = { style: BorderStyle.SINGLE, size: 2, color: 'CCCCCC' };

/**
 * Build a branded docx table.
 * spec = { headers: [..], rows: [[..]], colWidths?, highlightNegatives?, totalRowIndex? }
 */
export function dataTable(spec) {
  const { headers = [], rows = [], colWidths, highlightNegatives = true, totalRowIndex } = spec;

  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) => cell(h, {
      bold: true, color: BRAND.white, shading: BRAND.navy,
      align: i === 0 ? 'left' : 'center', width: colWidths?.[i],
    })),
  });

  const bodyRows = rows.map((row, ri) => new TableRow({
    children: row.map((val, ci) => {
      const isNumeric = typeof val === 'number';
      const display = isNumeric ? formatNumber(val) : val;
      const text = String(display ?? '—');
      const negative = highlightNegatives
        && ci > 0
        && (isNumeric ? val < 0 : /^\(|^-\d|^-\s?\d/.test(text.trim()));
      const isTotal = totalRowIndex != null && ri === totalRowIndex;
      return cell(display, {
        align: ci === 0 ? 'left' : 'right',
        bold: ci === 0 || isTotal,
        color: negative ? BRAND.negative : undefined,
        shading: isTotal ? BRAND.lightBlue : ri % 2 === 1 ? BRAND.lightGray : undefined,
        width: colWidths?.[ci],
      });
    }),
  }));

  return new Table({
    rows: [headerRow, ...bodyRows],
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: THIN, bottom: THIN, left: THIN, right: THIN,
      insideHorizontal: THIN, insideVertical: THIN,
    },
  });
}

/** Two-column key/facts table (label bold navy, value plain). */
export function factsTable(pairs) {
  const rows = pairs.map(([label, value], ri) => new TableRow({
    children: [
      cell(label, { bold: true, color: BRAND.navy, width: 35, shading: ri % 2 === 1 ? BRAND.lightGray : undefined }),
      cell(value, { width: 65, shading: ri % 2 === 1 ? BRAND.lightGray : undefined }),
    ],
  }));
  return new Table({
    rows,
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: THIN, bottom: THIN, left: THIN, right: THIN,
      insideHorizontal: THIN, insideVertical: THIN,
    },
  });
}

/** Checklist table: [{ item, status?, owner?, notes? }] */
export function checklistTable(items, { lang = 'fr' } = {}) {
  const headers = lang === 'fr'
    ? ['#', 'Élément', 'Statut', 'Responsable', 'Commentaires']
    : ['#', 'Item', 'Status', 'Owner', 'Notes'];
  const rows = items.map((it, i) => [
    i + 1,
    it.item || it.label || String(it),
    it.status || '☐',
    it.owner || '—',
    it.notes || '',
  ]);
  return dataTable({ headers, rows, colWidths: [5, 45, 12, 15, 23], highlightNegatives: false });
}
