// generators/hedge-fund/attribution-analysis.js
// Attribution de performance (xlsx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildWorkbook } from '../../shared/xlsx-model.js';
import { formatPercent } from '../../shared/financial-helpers.js';

export const metadata = {
  id: 'attribution-analysis',
  name: 'Attribution de Performance',
  vertical: 'hedge-fund',
  outputType: 'xlsx',
  estimatedPages: '3 sheets',
  requiredInput: ['fund'],
  optionalInput: ['period', 'bySector', 'byPosition', 'byStrategy', 'totalReturn'],
};

const DEFAULT_SECTORS = [
  { name: 'Technologie', allocation: 0.009, selection: 0.014, total: 0.023 },
  { name: 'Santé', allocation: -0.002, selection: 0.008, total: 0.006 },
  { name: 'Industrie', allocation: 0.003, selection: -0.005, total: -0.002 },
  { name: 'Consommation', allocation: 0.001, selection: 0.004, total: 0.005 },
];

const DEFAULT_POSITIONS = [
  { name: 'Position A', contribution: 0.012 },
  { name: 'Position B', contribution: 0.008 },
  { name: 'Position C', contribution: 0.006 },
  { name: 'Position D', contribution: -0.004 },
  { name: 'Position E', contribution: -0.006 },
];

export async function generate(input = {}) {
  const fund = input.fund || 'Fonds';
  const period = input.period || 'Mois en cours';
  const bySector = input.bySector?.length ? input.bySector : DEFAULT_SECTORS;
  const byPosition = input.byPosition?.length ? input.byPosition : DEFAULT_POSITIONS;
  const total = input.totalReturn ?? bySector.reduce((s, x) => s + (x.total || 0), 0);

  const winners = [...byPosition].sort((a, b) => b.contribution - a.contribution).slice(0, 5);
  const losers = [...byPosition].sort((a, b) => a.contribution - b.contribution).slice(0, 5);

  return buildWorkbook({
    title: `Attribution — ${fund}`,
    sheets: [
      {
        name: 'Par secteur',
        sectionTitle: `Attribution par secteur — ${period}`,
        table: {
          headers: ['Secteur', 'Effet allocation', 'Effet sélection', 'Total'],
          rows: [
            ...bySector.map((s) => [s.name, formatPercent(s.allocation), formatPercent(s.selection), formatPercent(s.total)]),
            ['Total portefeuille', '—', '—', formatPercent(total)],
          ],
          totalRowIndex: bySector.length,
        },
        columns: [{ width: 22 }, { width: 16 }, { width: 16 }, { width: 12 }],
        freezeHeader: true,
      },
      {
        name: 'Top contributeurs',
        sectionTitle: 'Meilleurs et pires contributeurs',
        table: {
          headers: ['Rang', 'Contributeurs positifs', 'Contribution', 'Contributeurs négatifs', 'Contribution'],
          rows: Array.from({ length: Math.max(winners.length, losers.length) }, (_, i) => [
            i + 1,
            winners[i]?.name || '',
            winners[i] ? formatPercent(winners[i].contribution) : '',
            losers[i]?.name || '',
            losers[i] ? formatPercent(losers[i].contribution) : '',
          ]),
        },
        columns: [{ width: 8 }, { width: 26 }, { width: 14 }, { width: 26 }, { width: 14 }],
      },
      {
        name: 'Par stratégie',
        sectionTitle: 'Attribution par poche / stratégie',
        table: {
          headers: ['Stratégie', 'Contribution'],
          rows: (input.byStrategy?.length ? input.byStrategy : [
            { name: 'Long fondamental', contribution: 0.021 },
            { name: 'Short alpha', contribution: 0.006 },
            { name: 'Couvertures', contribution: -0.005 },
            { name: 'Situations spéciales', contribution: 0.010 },
          ]).map((s) => [s.name, formatPercent(s.contribution)]),
        },
        columns: [{ width: 28 }, { width: 14 }],
      },
    ],
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
