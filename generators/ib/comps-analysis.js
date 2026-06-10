// generators/ib/comps-analysis.js
// Comparables boursiers et transactionnels (xlsx + pptx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildWorkbook } from '../../shared/xlsx-model.js';
import { buildDeck } from '../../shared/pptx-deck.js';
import { formatMultiple, formatPercent } from '../../shared/financial-helpers.js';

export const metadata = {
  id: 'comps-analysis',
  name: 'Analyse de Comparables',
  vertical: 'ib',
  outputType: 'multi',
  outputs: ['xlsx', 'pptx'],
  estimatedPages: '3 sheets + 4 slides',
  requiredInput: ['target'],
  optionalInput: ['tradingComps', 'transactionComps', 'targetMetrics'],
};

const DEFAULT_TRADING = [
  { name: 'Comp A', evEbitda: 9.5, evSales: 1.8, pe: 16.2, margin: 0.19 },
  { name: 'Comp B', evEbitda: 11.2, evSales: 2.3, pe: 19.5, margin: 0.21 },
  { name: 'Comp C', evEbitda: 8.1, evSales: 1.4, pe: 13.8, margin: 0.16 },
  { name: 'Comp D', evEbitda: 10.4, evSales: 2.0, pe: 17.9, margin: 0.20 },
];

const DEFAULT_TRANSACTIONS = [
  { name: 'Deal X (2024)', evEbitda: 11.8, evSales: 2.4 },
  { name: 'Deal Y (2023)', evEbitda: 10.2, evSales: 2.0 },
  { name: 'Deal Z (2023)', evEbitda: 12.5, evSales: 2.6 },
];

function stats(values) {
  const v = values.filter((x) => Number.isFinite(x)).sort((a, b) => a - b);
  if (!v.length) return { min: null, median: null, mean: null, max: null };
  const median = v.length % 2 ? v[(v.length - 1) / 2] : (v[v.length / 2 - 1] + v[v.length / 2]) / 2;
  return {
    min: v[0],
    median,
    mean: v.reduce((s, x) => s + x, 0) / v.length,
    max: v[v.length - 1],
  };
}

export async function generate(input = {}) {
  const target = input.target || 'Cible';
  const trading = input.tradingComps?.length ? input.tradingComps : DEFAULT_TRADING;
  const transactions = input.transactionComps?.length ? input.transactionComps : DEFAULT_TRANSACTIONS;

  const tEbitda = stats(trading.map((c) => c.evEbitda));
  const tSales = stats(trading.map((c) => c.evSales));
  const txEbitda = stats(transactions.map((c) => c.evEbitda));

  const xlsx = await buildWorkbook({
    title: `Comps — ${target}`,
    sheets: [
      {
        name: 'Trading Comps',
        sectionTitle: `Comparables boursiers — ${target}`,
        table: {
          headers: ['Société', 'VE/EBITDA', 'VE/CA', 'P/E', 'Marge EBITDA'],
          rows: [
            ...trading.map((c) => [c.name, fm(c.evEbitda), fm(c.evSales), fm(c.pe), c.margin != null ? formatPercent(c.margin) : '—']),
            ['Médiane', fm(tEbitda.median), fm(tSales.median), '—', '—'],
            ['Moyenne', fm(tEbitda.mean), fm(tSales.mean), '—', '—'],
          ],
          totalRowIndex: trading.length,
        },
        columns: [{ width: 26 }, { width: 14 }, { width: 12 }, { width: 12 }, { width: 16 }],
        freezeHeader: true,
      },
      {
        name: 'Transaction Comps',
        sectionTitle: 'Comparables transactionnels',
        table: {
          headers: ['Transaction', 'VE/EBITDA', 'VE/CA'],
          rows: [
            ...transactions.map((c) => [c.name, fm(c.evEbitda), fm(c.evSales)]),
            ['Médiane', fm(txEbitda.median), fm(stats(transactions.map((c) => c.evSales)).median)],
          ],
          totalRowIndex: transactions.length,
        },
        columns: [{ width: 30 }, { width: 14 }, { width: 12 }],
        freezeHeader: true,
      },
      {
        name: 'Synthèse',
        sectionTitle: 'Synthèse des multiples',
        table: {
          headers: ['Méthode', 'Min', 'Médiane', 'Max'],
          rows: [
            ['Trading VE/EBITDA', fm(tEbitda.min), fm(tEbitda.median), fm(tEbitda.max)],
            ['Trading VE/CA', fm(tSales.min), fm(tSales.median), fm(tSales.max)],
            ['Transactions VE/EBITDA', fm(txEbitda.min), fm(txEbitda.median), fm(txEbitda.max)],
          ],
        },
        columns: [{ width: 28 }, { width: 12 }, { width: 12 }, { width: 12 }],
      },
    ],
  });

  const pptx = await buildDeck({
    title: `Analyse de comparables — ${target}`,
    subtitle: 'Trading & Transaction comps',
    confidential: true,
    slides: [
      { type: 'section', title: 'Comparables boursiers' },
      {
        type: 'table',
        title: 'Trading comps',
        headers: ['Société', 'VE/EBITDA', 'VE/CA', 'P/E'],
        rows: trading.map((c) => [c.name, fm(c.evEbitda), fm(c.evSales), fm(c.pe)]),
      },
      {
        type: 'table',
        title: 'Transaction comps',
        headers: ['Transaction', 'VE/EBITDA', 'VE/CA'],
        rows: transactions.map((c) => [c.name, fm(c.evEbitda), fm(c.evSales)]),
      },
      {
        type: 'chart',
        title: 'Multiples VE/EBITDA',
        chart: {
          kind: 'bar',
          categories: [...trading.map((c) => c.name), 'Médiane trading', 'Médiane transactions'],
          series: [{ name: 'VE/EBITDA', values: [...trading.map((c) => c.evEbitda), tEbitda.median, txEbitda.median] }],
        },
      },
    ],
  });

  return [
    { fileName: 'comps-analysis.xlsx', buffer: xlsx, ext: 'xlsx' },
    { fileName: 'comps-analysis.pptx', buffer: pptx, ext: 'pptx' },
  ];
}

function fm(v) {
  return v == null ? '—' : formatMultiple(v, { decimals: 1 });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
