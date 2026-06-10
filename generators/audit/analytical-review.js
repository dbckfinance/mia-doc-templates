// generators/audit/analytical-review.js
// Revue analytique (xlsx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildWorkbook } from '../../shared/xlsx-model.js';
import { variance, formatPercent } from '../../shared/financial-helpers.js';

export const metadata = {
  id: 'analytical-review',
  name: 'Revue Analytique',
  vertical: 'audit',
  outputType: 'xlsx',
  estimatedPages: '2-3 sheets',
  requiredInput: ['company'],
  optionalInput: ['lines', 'currentLabel', 'priorLabel', 'threshold'],
};

const DEFAULT_LINES = [
  { label: 'Chiffre d\'affaires', current: 1250, prior: 1100 },
  { label: 'Marge brute', current: 540, prior: 495 },
  { label: 'Charges de personnel', current: -310, prior: -270 },
  { label: 'Autres charges externes', current: -185, prior: -150 },
  { label: 'EBITDA', current: 240, prior: 225 },
  { label: 'Dotations aux amortissements', current: -65, prior: -60 },
  { label: 'Résultat d\'exploitation', current: 175, prior: 165 },
  { label: 'Résultat net', current: 118, prior: 112 },
];

export async function generate(input = {}) {
  const company = input.company || 'Société';
  const lines = input.lines?.length ? input.lines : DEFAULT_LINES;
  const currentLabel = input.currentLabel || 'N';
  const priorLabel = input.priorLabel || 'N-1';
  const threshold = input.threshold ?? 0.10;

  const enriched = lines.map((l) => {
    const v = variance(l.current, l.prior);
    const significant = v.pct != null && Math.abs(v.pct) >= threshold;
    return { ...l, abs: v.abs, pct: v.pct, significant };
  });

  return buildWorkbook({
    title: `Revue analytique — ${company}`,
    sheets: [
      {
        name: 'Revue analytique',
        sectionTitle: `Revue analytique ${currentLabel} vs ${priorLabel} — ${company} (m€)`,
        table: {
          headers: ['Poste', currentLabel, priorLabel, 'Variation', 'Variation %', `Significatif (>${formatPercent(threshold)})`],
          rows: enriched.map((l) => [
            l.label,
            l.current,
            l.prior,
            l.abs != null ? Math.round(l.abs * 10) / 10 : '—',
            l.pct != null ? formatPercent(l.pct) : '—',
            l.significant ? 'OUI — à investiguer' : 'Non',
          ]),
        },
        columns: [{ width: 32 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 24 }],
        freezeHeader: true,
      },
      {
        name: 'Investigations',
        sectionTitle: 'Postes à investiguer',
        table: {
          headers: ['Poste', 'Variation %', 'Explication obtenue', 'Corroboration', 'Conclusion'],
          rows: enriched.filter((l) => l.significant).map((l) => [
            l.label,
            l.pct != null ? formatPercent(l.pct) : '—',
            'À compléter (entretien direction)',
            'À compléter (pièces justificatives)',
            'À conclure',
          ]),
        },
        columns: [{ width: 32 }, { width: 12 }, { width: 36 }, { width: 36 }, { width: 16 }],
      },
    ],
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
