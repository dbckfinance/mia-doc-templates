// generators/equity-research/target-price-update.js
// Changement de recommandation / objectif de cours (docx, 1 page)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildReport } from '../../shared/docx-report.js';

export const metadata = {
  id: 'target-price-update',
  name: "Changement de Recommandation / Objectif",
  vertical: 'equity-research',
  outputType: 'docx',
  estimatedPages: '1',
  requiredInput: ['ticker'],
  optionalInput: ['company', 'previousRating', 'newRating', 'previousTarget', 'newTarget', 'currentPrice', 'rationale', 'analyst'],
};

export async function generate(input = {}) {
  const ticker = input.ticker || 'TICKER';
  const company = input.company || ticker;
  const prevRating = input.previousRating || 'Neutre';
  const newRating = input.newRating || 'Achat';
  const prevTarget = input.previousTarget;
  const newTarget = input.newTarget;
  const current = input.currentPrice;
  const upside = current && newTarget ? `${((newTarget / current - 1) * 100).toFixed(1)}%` : 'n.d.';
  const ratingChanged = prevRating !== newRating;

  return buildReport({
    docTitle: `${company} — ${ratingChanged ? `${prevRating} → ${newRating}` : 'Objectif de cours révisé'}`,
    docSubtitle: `${ticker} | Objectif : ${prevTarget ?? 'n.d.'} → ${newTarget ?? 'n.d.'} (potentiel ${upside})`,
    vertical: 'equity-research',
    docId: 'target-price-update',
    confidential: false,
    sections: [
      {
        heading: 'Changement',
        blocks: [
          { type: 'facts', facts: [
            { label: 'Recommandation précédente', value: prevRating },
            { label: 'Nouvelle recommandation', value: newRating },
            { label: 'Objectif précédent', value: prevTarget != null ? String(prevTarget) : 'n.d.' },
            { label: 'Nouvel objectif', value: newTarget != null ? String(newTarget) : 'n.d.' },
            { label: 'Cours actuel', value: current != null ? String(current) : 'n.d.' },
            { label: 'Potentiel', value: upside },
          ] },
        ],
      },
      {
        heading: 'Justification',
        blocks: [
          { type: 'bullets', items: input.rationale || [
            'Révision à la hausse de nos estimations de résultats',
            'Amélioration du profil de génération de cash',
            'Re-rating justifié par la réduction du risque d\'exécution',
          ] },
        ],
      },
      {
        heading: 'Signature',
        blocks: [
          { type: 'kv', label: 'Analyste', value: input.analyst || 'M&IA Research' },
          { type: 'kv', label: 'Date', value: new Date().toLocaleDateString('fr-FR') },
        ],
      },
    ],
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
