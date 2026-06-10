// generators/equity-research/peer-comparison.js
// Matrice de comparaison des pairs (xlsx + pptx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildWorkbook } from '../../shared/xlsx-model.js';
import { buildDeck } from '../../shared/pptx-deck.js';
import { formatPercent, formatMultiple } from '../../shared/financial-helpers.js';

export const metadata = {
  id: 'peer-comparison',
  name: 'Comparaison des Pairs',
  vertical: 'equity-research',
  outputType: 'multi',
  outputs: ['xlsx', 'pptx'],
  estimatedPages: '2 sheets + 2 slides',
  requiredInput: ['ticker'],
  optionalInput: ['company', 'peers'],
};

const DEFAULT_PEERS = [
  { name: 'Société analysée', evEbitda: 8.2, pe: 14.5, growth: 0.08, margin: 0.21, roce: 0.15, isSubject: true },
  { name: 'Pair A', evEbitda: 10.1, pe: 18.2, growth: 0.06, margin: 0.19, roce: 0.13 },
  { name: 'Pair B', evEbitda: 9.4, pe: 16.8, growth: 0.07, margin: 0.22, roce: 0.16 },
  { name: 'Pair C', evEbitda: 11.6, pe: 21.0, growth: 0.10, margin: 0.24, roce: 0.18 },
  { name: 'Pair D', evEbitda: 7.8, pe: 13.2, growth: 0.03, margin: 0.16, roce: 0.10 },
];

export async function generate(input = {}) {
  const ticker = input.ticker || 'TICKER';
  const company = input.company || ticker;
  const peers = input.peers?.length ? input.peers : DEFAULT_PEERS.map((p) => p.isSubject ? { ...p, name: company } : p);

  const median = (key) => {
    const vals = peers.filter((p) => !p.isSubject).map((p) => p[key]).filter(Number.isFinite).sort((a, b) => a - b);
    if (!vals.length) return null;
    return vals.length % 2 ? vals[(vals.length - 1) / 2] : (vals[vals.length / 2 - 1] + vals[vals.length / 2]) / 2;
  };

  const xlsx = await buildWorkbook({
    title: `Peer comparison — ${company}`,
    sheets: [
      {
        name: 'Matrice',
        sectionTitle: `Comparaison des pairs — ${company}`,
        table: {
          headers: ['Société', 'VE/EBITDA', 'P/E', 'Croissance CA', 'Marge EBITDA', 'ROCE'],
          rows: [
            ...peers.map((p) => [
              p.name,
              fm(p.evEbitda),
              fm(p.pe),
              fp(p.growth),
              fp(p.margin),
              fp(p.roce),
            ]),
            ['Médiane pairs', fm(median('evEbitda')), fm(median('pe')), fp(median('growth')), fp(median('margin')), fp(median('roce'))],
          ],
          totalRowIndex: peers.length,
        },
        columns: [{ width: 26 }, { width: 12 }, { width: 10 }, { width: 14 }, { width: 14 }, { width: 10 }],
        freezeHeader: true,
      },
      {
        name: 'Prime-décote',
        sectionTitle: 'Prime / décote vs médiane des pairs',
        table: {
          headers: ['Multiple', company, 'Médiane pairs', 'Prime / (décote)'],
          rows: ['evEbitda', 'pe'].map((k) => {
            const subject = peers.find((p) => p.isSubject) || peers[0];
            const med = median(k);
            const delta = subject[k] && med ? subject[k] / med - 1 : null;
            return [k === 'evEbitda' ? 'VE/EBITDA' : 'P/E', fm(subject[k]), fm(med), delta != null ? formatPercent(delta) : '—'];
          }),
        },
        columns: [{ width: 16 }, { width: 14 }, { width: 16 }, { width: 18 }],
      },
    ],
  });

  const pptx = await buildDeck({
    title: `Comparaison des pairs — ${company}`,
    subtitle: ticker,
    confidential: false,
    slides: [
      {
        type: 'table',
        title: 'Matrice des pairs',
        headers: ['Société', 'VE/EBITDA', 'P/E', 'Croissance', 'Marge'],
        rows: peers.map((p) => [p.name, fm(p.evEbitda), fm(p.pe), fp(p.growth), fp(p.margin)]),
      },
      {
        type: 'chart',
        title: 'VE/EBITDA vs pairs',
        chart: {
          kind: 'bar',
          categories: peers.map((p) => p.name),
          series: [{ name: 'VE/EBITDA', values: peers.map((p) => p.evEbitda) }],
        },
      },
    ],
  });

  return [
    { fileName: 'peer-comparison.xlsx', buffer: xlsx, ext: 'xlsx' },
    { fileName: 'peer-comparison.pptx', buffer: pptx, ext: 'pptx' },
  ];
}

function fm(v) { return v == null ? '—' : formatMultiple(v, { decimals: 1 }); }
function fp(v) { return v == null ? '—' : formatPercent(v); }

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
