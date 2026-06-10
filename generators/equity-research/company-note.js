// generators/equity-research/company-note.js
// Note rapide sur une société (docx, 1-2 pages)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildReport } from '../../shared/docx-report.js';

export const metadata = {
  id: 'company-note',
  name: 'Note Société (Flash)',
  vertical: 'equity-research',
  outputType: 'docx',
  estimatedPages: '1-2',
  requiredInput: ['ticker'],
  optionalInput: ['company', 'rating', 'targetPrice', 'event', 'analysis', 'implication', 'analyst'],
};

export async function generate(input = {}) {
  const ticker = input.ticker || 'TICKER';
  const company = input.company || ticker;

  return buildReport({
    docTitle: `${company} — Note flash`,
    docSubtitle: `${ticker} | ${input.rating || 'Achat'} | Objectif : ${input.targetPrice ?? 'n.d.'}`,
    vertical: 'equity-research',
    docId: 'company-note',
    confidential: false,
    sections: [
      {
        heading: 'Événement',
        blocks: [
          { type: 'p', text: input.event || 'Annonce de la société (communiqué, contrat, nomination, opération) résumée en quelques lignes factuelles.' },
        ],
      },
      {
        heading: 'Notre analyse',
        blocks: [
          { type: 'p', text: input.analysis || 'Lecture de l\'événement : portée stratégique, impact financier estimé, comparaison avec nos hypothèses actuelles.' },
        ],
      },
      {
        heading: 'Implication pour la recommandation',
        blocks: [
          { type: 'p', text: input.implication || `Sans impact sur notre scénario central à ce stade. Nous maintenons notre recommandation ${input.rating || 'Achat'} et notre objectif de cours.` },
          { type: 'kv', label: 'Analyste', value: input.analyst || 'M&IA Research' },
          { type: 'kv', label: 'Date', value: new Date().toLocaleDateString('fr-FR') },
        ],
      },
    ],
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
