// generators/audit/qoe-report.js
// Rapport Quality of Earnings (docx + xlsx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildReport } from '../../shared/docx-report.js';
import { buildWorkbook } from '../../shared/xlsx-model.js';

export const metadata = {
  id: 'qoe-report',
  name: 'Quality of Earnings (QoE)',
  vertical: 'audit',
  outputType: 'multi',
  outputs: ['docx', 'xlsx'],
  estimatedPages: '10-20 pages + 2 sheets',
  requiredInput: ['company'],
  optionalInput: ['reportedEbitda', 'adjustments', 'period', 'workingCapital', 'netDebtItems'],
};

const DEFAULT_ADJUSTMENTS = [
  { label: 'Rémunération dirigeants au-dessus du marché', amount: 1.2, category: 'Normalisation' },
  { label: 'Honoraires exceptionnels (litige)', amount: 0.8, category: 'Non-récurrent' },
  { label: 'Loyer intra-groupe sous-évalué', amount: -0.5, category: 'Normalisation' },
  { label: 'Crédit COVID non récurrent', amount: -0.6, category: 'Non-récurrent' },
];

export async function generate(input = {}) {
  const company = input.company || 'Société';
  const period = input.period || 'LTM';
  const reported = input.reportedEbitda ?? 24.0;
  const adjustments = input.adjustments?.length ? input.adjustments : DEFAULT_ADJUSTMENTS;
  const totalAdj = adjustments.reduce((s, a) => s + a.amount, 0);
  const adjusted = reported + totalAdj;

  const docx = await buildReport({
    docTitle: 'Rapport Quality of Earnings',
    docSubtitle: `${company} — ${period}`,
    vertical: 'audit',
    docId: 'qoe-report',
    confidential: true,
    sections: [
      {
        heading: 'Synthèse',
        blocks: [
          { type: 'kv', label: 'EBITDA reporté', value: `${reported.toFixed(1)} m€` },
          { type: 'kv', label: 'Total ajustements', value: `${totalAdj >= 0 ? '+' : ''}${totalAdj.toFixed(1)} m€` },
          { type: 'kv', label: 'EBITDA ajusté (QoE)', value: `${adjusted.toFixed(1)} m€` },
          { type: 'p', text: `Notre revue de la qualité des résultats de ${company} sur la période ${period} conduit à un EBITDA ajusté de ${adjusted.toFixed(1)} m€, contre ${reported.toFixed(1)} m€ reporté.` },
        ],
      },
      {
        heading: 'Pont EBITDA reporté → ajusté',
        blocks: [
          {
            type: 'table',
            headers: ['Ajustement', 'Catégorie', 'Impact (m€)'],
            rows: [
              ['EBITDA reporté', '—', reported.toFixed(1)],
              ...adjustments.map((a) => [a.label, a.category || '—', (a.amount >= 0 ? '+' : '') + a.amount.toFixed(1)]),
              ['EBITDA ajusté', '—', adjusted.toFixed(1)],
            ],
          },
        ],
      },
      {
        heading: 'Analyse des ajustements',
        blocks: adjustments.map((a) => ({
          type: 'p',
          text: `${a.label} (${(a.amount >= 0 ? '+' : '') + a.amount.toFixed(1)} m€) : ${a.rationale || 'ajustement identifié lors de nos travaux de revue des comptes de gestion et des pièces justificatives.'}`,
        })),
      },
      {
        heading: 'BFR normatif et dette nette',
        blocks: [
          { type: 'p', text: input.workingCapital?.commentary || 'Le BFR normatif a été estimé sur la base de la moyenne mensuelle observée sur 24 mois, corrigée des éléments exceptionnels. Les éléments assimilés à de la dette (debt-like items) sont détaillés dans le classeur joint.' },
        ],
      },
      {
        heading: 'Limites des travaux',
        blocks: [
          { type: 'p', text: 'Nos travaux ne constituent ni un audit ni un examen limité au sens des normes professionnelles. Ils reposent sur les informations communiquées par la direction, dont nous n\'avons pas vérifié l\'exhaustivité de manière indépendante.' },
        ],
      },
    ],
  });

  const xlsx = await buildWorkbook({
    title: `QoE — ${company}`,
    sheets: [
      {
        name: 'Pont EBITDA',
        sectionTitle: `Pont EBITDA reporté → ajusté (${period}, m€)`,
        table: {
          headers: ['Élément', 'Catégorie', 'm€'],
          rows: [
            ['EBITDA reporté', '—', reported],
            ...adjustments.map((a) => [a.label, a.category || '—', a.amount]),
            ['EBITDA ajusté', '—', Math.round(adjusted * 10) / 10],
          ],
          totalRowIndex: adjustments.length + 1,
        },
        columns: [{ width: 48 }, { width: 18 }, { width: 12 }],
      },
      {
        name: 'Debt-like items',
        sectionTitle: 'Éléments assimilés à de la dette',
        table: {
          headers: ['Élément', 'm€'],
          rows: (input.netDebtItems?.length ? input.netDebtItems : [
            { label: 'Provisions pour litiges', amount: 1.5 },
            { label: 'Engagements de retraite non financés', amount: 2.1 },
            { label: 'Earn-out à payer', amount: 0.9 },
          ]).map((i) => [i.label, i.amount]),
        },
        columns: [{ width: 48 }, { width: 12 }],
      },
    ],
  });

  return [
    { fileName: 'qoe-report.docx', buffer: docx, ext: 'docx' },
    { fileName: 'qoe-bridge.xlsx', buffer: xlsx, ext: 'xlsx' },
  ];
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
