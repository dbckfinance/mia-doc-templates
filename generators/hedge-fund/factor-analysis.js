// generators/hedge-fund/factor-analysis.js
// Rapport d'expositions factorielles (xlsx + pptx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildWorkbook } from '../../shared/xlsx-model.js';
import { buildDeck } from '../../shared/pptx-deck.js';
import { formatPercent } from '../../shared/financial-helpers.js';

export const metadata = {
  id: 'factor-analysis',
  name: 'Analyse Factorielle',
  vertical: 'hedge-fund',
  outputType: 'multi',
  outputs: ['xlsx', 'pptx'],
  estimatedPages: '2 sheets + 2 slides',
  requiredInput: ['fund'],
  optionalInput: ['period', 'factors', 'r2', 'alphaAnnualized'],
};

const DEFAULT_FACTORS = [
  { name: 'Marché (beta)', exposure: 0.42, contribution: 0.011 },
  { name: 'Taille (SMB)', exposure: 0.18, contribution: 0.002 },
  { name: 'Value (HML)', exposure: -0.12, contribution: -0.001 },
  { name: 'Momentum', exposure: 0.25, contribution: 0.006 },
  { name: 'Qualité', exposure: 0.30, contribution: 0.004 },
  { name: 'Volatilité', exposure: -0.08, contribution: 0.001 },
];

export async function generate(input = {}) {
  const fund = input.fund || 'Fonds';
  const period = input.period || 'Trimestre en cours';
  const factors = input.factors?.length ? input.factors : DEFAULT_FACTORS;
  const r2 = input.r2 ?? 0.61;
  const alpha = input.alphaAnnualized ?? 0.045;

  const xlsx = await buildWorkbook({
    title: `Analyse factorielle — ${fund}`,
    sheets: [
      {
        name: 'Expositions',
        sectionTitle: `Expositions factorielles — ${period}`,
        table: {
          headers: ['Facteur', 'Exposition (beta)', 'Contribution au rendement'],
          rows: factors.map((f) => [f.name, f.exposure, formatPercent(f.contribution)]),
        },
        columns: [{ width: 24 }, { width: 18 }, { width: 24 }],
        freezeHeader: true,
      },
      {
        name: 'Synthèse',
        sectionTitle: 'Synthèse du modèle',
        table: {
          headers: ['Métrique', 'Valeur'],
          rows: [
            ['R² du modèle factoriel', formatPercent(r2)],
            ['Alpha annualisé (non expliqué)', formatPercent(alpha)],
            ['Part systématique du risque', formatPercent(r2)],
            ['Part idiosyncratique', formatPercent(1 - r2)],
          ],
        },
        columns: [{ width: 34 }, { width: 14 }],
      },
    ],
  });

  const pptx = await buildDeck({
    title: `Analyse factorielle — ${fund}`,
    subtitle: period,
    confidential: true,
    slides: [
      {
        type: 'chart',
        title: 'Expositions factorielles (beta)',
        chart: {
          kind: 'bar',
          categories: factors.map((f) => f.name),
          series: [{ name: 'Exposition', values: factors.map((f) => f.exposure) }],
        },
      },
      {
        type: 'facts',
        title: 'Lecture du profil',
        facts: [
          { label: 'R² du modèle', value: formatPercent(r2) },
          { label: 'Alpha annualisé', value: formatPercent(alpha) },
          { label: 'Biais dominant', value: factors.reduce((m, f) => Math.abs(f.exposure) > Math.abs(m.exposure) ? f : m, factors[0]).name },
        ],
      },
    ],
  });

  return [
    { fileName: 'factor-analysis.xlsx', buffer: xlsx, ext: 'xlsx' },
    { fileName: 'factor-analysis.pptx', buffer: pptx, ext: 'pptx' },
  ];
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
