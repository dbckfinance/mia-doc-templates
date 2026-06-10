// generators/hedge-fund/portfolio-report.js
// Rapport mensuel de portefeuille (pptx + xlsx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildDeck } from '../../shared/pptx-deck.js';
import { buildWorkbook } from '../../shared/xlsx-model.js';
import { formatPercent } from '../../shared/financial-helpers.js';

export const metadata = {
  id: 'portfolio-report',
  name: 'Rapport Mensuel de Portefeuille',
  vertical: 'hedge-fund',
  outputType: 'multi',
  outputs: ['pptx', 'xlsx'],
  estimatedPages: '5 slides + 2 sheets',
  requiredInput: ['fund'],
  optionalInput: ['period', 'performance', 'positions', 'exposures', 'commentary'],
};

const DEFAULT_POSITIONS = [
  { name: 'Position A', sector: 'Tech', weight: 0.08, pnlContribution: 0.012, direction: 'Long' },
  { name: 'Position B', sector: 'Santé', weight: 0.06, pnlContribution: 0.008, direction: 'Long' },
  { name: 'Position C', sector: 'Industrie', weight: 0.05, pnlContribution: -0.004, direction: 'Long' },
  { name: 'Position D', sector: 'Conso', weight: -0.04, pnlContribution: 0.006, direction: 'Short' },
];

export async function generate(input = {}) {
  const fund = input.fund || 'Fonds';
  const period = input.period || new Date().toLocaleDateString('fr-FR', { month: 'long', year: 'numeric' });
  const perf = input.performance || { month: 0.018, ytd: 0.072, sinceInception: 0.34, benchmark: 0.011 };
  const positions = input.positions?.length ? input.positions : DEFAULT_POSITIONS;
  const exposures = input.exposures || { gross: 1.45, net: 0.42, long: 0.93, short: -0.51 };

  const pptx = await buildDeck({
    title: `${fund} — Rapport mensuel`,
    subtitle: period,
    confidential: true,
    slides: [
      {
        type: 'facts',
        title: 'Performance',
        facts: [
          { label: 'Mois', value: formatPercent(perf.month) },
          { label: 'YTD', value: formatPercent(perf.ytd) },
          { label: 'Depuis création', value: formatPercent(perf.sinceInception) },
          { label: 'Indice de référence (mois)', value: formatPercent(perf.benchmark) },
        ],
      },
      {
        type: 'facts',
        title: 'Expositions',
        facts: [
          { label: 'Brute', value: formatPercent(exposures.gross) },
          { label: 'Nette', value: formatPercent(exposures.net) },
          { label: 'Long', value: formatPercent(exposures.long) },
          { label: 'Short', value: formatPercent(exposures.short) },
        ],
      },
      {
        type: 'table',
        title: 'Principales positions',
        headers: ['Position', 'Secteur', 'Sens', 'Poids', 'Contribution P&L'],
        rows: positions.map((p) => [p.name, p.sector || '—', p.direction || '—', formatPercent(p.weight), formatPercent(p.pnlContribution)]),
      },
      {
        type: 'chart',
        title: 'Contribution P&L par position',
        chart: {
          kind: 'bar',
          categories: positions.map((p) => p.name),
          series: [{ name: 'Contribution', values: positions.map((p) => Math.round((p.pnlContribution || 0) * 10000) / 100) }],
          axisFormat: '0.00"%"',
        },
      },
      {
        type: 'content',
        title: 'Commentaire de gestion',
        bullets: input.commentary || [
          'Mois positif porté par la sélection de titres dans la technologie',
          'Réduction de l\'exposition nette en fin de période',
          'Pipeline d\'idées concentré sur les situations spéciales',
        ],
      },
    ],
  });

  const xlsx = await buildWorkbook({
    title: `${fund} — ${period}`,
    sheets: [
      {
        name: 'Performance',
        sectionTitle: `Performance — ${period}`,
        table: {
          headers: ['Métrique', 'Valeur'],
          rows: [
            ['Performance mensuelle', formatPercent(perf.month)],
            ['Performance YTD', formatPercent(perf.ytd)],
            ['Performance depuis création', formatPercent(perf.sinceInception)],
            ['Indice de référence (mois)', formatPercent(perf.benchmark)],
            ['Exposition brute', formatPercent(exposures.gross)],
            ['Exposition nette', formatPercent(exposures.net)],
          ],
        },
        columns: [{ width: 32 }, { width: 16 }],
      },
      {
        name: 'Positions',
        sectionTitle: 'Détail des positions',
        table: {
          headers: ['Position', 'Secteur', 'Sens', 'Poids', 'Contribution P&L'],
          rows: positions.map((p) => [p.name, p.sector || '—', p.direction || '—', formatPercent(p.weight), formatPercent(p.pnlContribution)]),
        },
        columns: [{ width: 24 }, { width: 16 }, { width: 10 }, { width: 12 }, { width: 18 }],
        freezeHeader: true,
      },
    ],
  });

  return [
    { fileName: 'portfolio-report.pptx', buffer: pptx, ext: 'pptx' },
    { fileName: 'portfolio-report.xlsx', buffer: xlsx, ext: 'xlsx' },
  ];
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
