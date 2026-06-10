// generators/equity-research/model-update.js
// Révision d'estimations (xlsx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildWorkbook } from '../../shared/xlsx-model.js';
import { formatPercent } from '../../shared/financial-helpers.js';

export const metadata = {
  id: 'model-update',
  name: "Révision d'Estimations",
  vertical: 'equity-research',
  outputType: 'xlsx',
  estimatedPages: '2-3 sheets',
  requiredInput: ['ticker'],
  optionalInput: ['company', 'revisions', 'rationale', 'valuationImpact'],
};

const DEFAULT_REVISIONS = [
  { metric: 'Chiffre d\'affaires N+1 (m€)', before: 2050, after: 2120 },
  { metric: 'EBITDA N+1 (m€)', before: 410, after: 435 },
  { metric: 'Marge EBITDA N+1', before: 0.20, after: 0.205, isPct: true },
  { metric: 'BPA N+1 (€)', before: 4.10, after: 4.38 },
  { metric: 'Chiffre d\'affaires N+2 (m€)', before: 2180, after: 2270 },
  { metric: 'BPA N+2 (€)', before: 4.55, after: 4.92 },
];

export async function generate(input = {}) {
  const ticker = input.ticker || 'TICKER';
  const company = input.company || ticker;
  const revisions = input.revisions?.length ? input.revisions : DEFAULT_REVISIONS;

  return buildWorkbook({
    title: `Révision d'estimations — ${company}`,
    sheets: [
      {
        name: 'Révisions',
        sectionTitle: `Révision des estimations — ${company} (${ticker})`,
        table: {
          headers: ['Indicateur', 'Avant', 'Après', 'Révision'],
          rows: revisions.map((r) => {
            const change = r.before ? (r.after / r.before - 1) : null;
            return [
              r.metric,
              r.isPct ? formatPercent(r.before) : r.before,
              r.isPct ? formatPercent(r.after) : r.after,
              r.isPct ? `${((r.after - r.before) * 100).toFixed(1)} pt` : (change != null ? formatPercent(change) : '—'),
            ];
          }),
        },
        columns: [{ width: 36 }, { width: 14 }, { width: 14 }, { width: 12 }],
        freezeHeader: true,
      },
      {
        name: 'Justification',
        sectionTitle: 'Justification des révisions',
        table: {
          headers: ['#', 'Motif'],
          rows: (input.rationale || [
            'Dynamique commerciale supérieure aux attentes au dernier trimestre',
            'Effet mix favorable sur les marges',
            'Révision du périmètre suite aux dernières acquisitions',
          ]).map((r, i) => [i + 1, r]),
        },
        columns: [{ width: 6 }, { width: 70 }],
      },
      {
        name: 'Impact valorisation',
        sectionTitle: 'Impact sur la valorisation',
        table: {
          headers: ['Élément', 'Valeur'],
          rows: input.valuationImpact?.length
            ? input.valuationImpact.map((v) => [v.label, v.value])
            : [
                ['Objectif de cours précédent', 'n.d.'],
                ['Nouvel objectif de cours', 'n.d.'],
                ['Recommandation', 'Inchangée'],
              ],
        },
        columns: [{ width: 32 }, { width: 20 }],
      },
    ],
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
