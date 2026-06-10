// generators/ib/sources-uses.js
// Tableau Sources & Emplois (xlsx + slide pptx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildWorkbook } from '../../shared/xlsx-model.js';
import { buildDeck } from '../../shared/pptx-deck.js';
import { formatPercent } from '../../shared/financial-helpers.js';

export const metadata = {
  id: 'sources-uses',
  name: 'Sources & Emplois',
  vertical: 'ib',
  outputType: 'multi',
  outputs: ['xlsx', 'pptx'],
  estimatedPages: '1 sheet + 1 slide',
  requiredInput: ['deal'],
  optionalInput: ['sources', 'uses', 'currency'],
};

const DEFAULT_SOURCES = [
  { label: 'Dette senior (TLB)', amount: 300 },
  { label: 'Dette mezzanine', amount: 80 },
  { label: 'Fonds propres sponsor', amount: 250 },
  { label: 'Réinvestissement management', amount: 30 },
];

const DEFAULT_USES = [
  { label: 'Prix d\'acquisition (equity)', amount: 560 },
  { label: 'Refinancement dette existante', amount: 70 },
  { label: 'Frais de transaction', amount: 22 },
  { label: 'Trésorerie au bilan', amount: 8 },
];

export async function generate(input = {}) {
  const deal = input.deal || 'Projet';
  const currency = input.currency || 'm€';
  const sources = input.sources?.length ? input.sources : DEFAULT_SOURCES;
  const uses = input.uses?.length ? input.uses : DEFAULT_USES;
  const totalSources = sources.reduce((s, x) => s + x.amount, 0);
  const totalUses = uses.reduce((s, x) => s + x.amount, 0);

  const rows = [];
  const n = Math.max(sources.length, uses.length);
  for (let i = 0; i < n; i++) {
    rows.push([
      sources[i]?.label || '', sources[i]?.amount ?? '', sources[i] ? formatPercent(sources[i].amount / totalSources) : '',
      uses[i]?.label || '', uses[i]?.amount ?? '', uses[i] ? formatPercent(uses[i].amount / totalUses) : '',
    ]);
  }
  rows.push(['Total sources', totalSources, '100%', 'Total emplois', totalUses, '100%']);

  const xlsx = await buildWorkbook({
    title: `Sources & Emplois — ${deal}`,
    sheets: [
      {
        name: 'Sources & Emplois',
        sectionTitle: `Sources & Emplois — ${deal} (${currency})`,
        table: {
          headers: ['Sources', currency, '%', 'Emplois', currency, '%'],
          rows,
          totalRowIndex: rows.length - 1,
        },
        columns: [{ width: 30 }, { width: 12 }, { width: 10 }, { width: 32 }, { width: 12 }, { width: 10 }],
      },
    ],
  });

  const pptx = await buildDeck({
    title: `Sources & Emplois — ${deal}`,
    subtitle: `Montants en ${currency}`,
    confidential: true,
    slides: [
      {
        type: 'table',
        title: `Sources & Emplois (${currency})`,
        headers: ['Sources', currency, 'Emplois', currency],
        rows: [
          ...Array.from({ length: n }, (_, i) => [
            sources[i]?.label || '', sources[i]?.amount ?? '',
            uses[i]?.label || '', uses[i]?.amount ?? '',
          ]),
          ['Total', totalSources, 'Total', totalUses],
        ],
      },
    ],
  });

  return [
    { fileName: 'sources-uses.xlsx', buffer: xlsx, ext: 'xlsx' },
    { fileName: 'sources-uses.pptx', buffer: pptx, ext: 'pptx' },
  ];
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
