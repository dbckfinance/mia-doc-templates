// generators/equity-research/earnings-update.js
// Note post-résultats (docx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildReport } from '../../shared/docx-report.js';

export const metadata = {
  id: 'earnings-update',
  name: 'Note Post-Résultats',
  vertical: 'equity-research',
  outputType: 'docx',
  estimatedPages: '2-5',
  requiredInput: ['ticker'],
  optionalInput: ['company', 'period', 'rating', 'targetPrice', 'results', 'takeaways', 'estimateChanges', 'analyst'],
};

export async function generate(input = {}) {
  const ticker = input.ticker || 'TICKER';
  const company = input.company || ticker;
  const period = input.period || 'T3';
  const results = input.results || [
    { metric: 'Chiffre d\'affaires', actual: '512 m€', consensus: '498 m€', delta: '+2,8%' },
    { metric: 'EBITDA', actual: '98 m€', consensus: '92 m€', delta: '+6,5%' },
    { metric: 'BPA', actual: '1,24 €', consensus: '1,15 €', delta: '+7,8%' },
  ];

  return buildReport({
    docTitle: `${company} — Résultats ${period}`,
    docSubtitle: `${ticker} | ${input.rating || 'Achat'} | Objectif : ${input.targetPrice ?? 'n.d.'}`,
    vertical: 'equity-research',
    docId: 'earnings-update',
    confidential: false,
    sections: [
      {
        heading: 'Résultats vs attentes',
        blocks: [
          {
            type: 'table',
            headers: ['Indicateur', 'Publié', 'Consensus', 'Écart'],
            rows: results.map((r) => [r.metric, r.actual, r.consensus, r.delta]),
          },
        ],
      },
      {
        heading: 'Points clés',
        blocks: [
          { type: 'bullets', items: input.takeaways || [
            'Publication supérieure aux attentes sur l\'ensemble des lignes',
            'Guidance annuelle relevée par le management',
            'Génération de cash solide, désendettement plus rapide que prévu',
          ] },
        ],
      },
      {
        heading: 'Révision de nos estimations',
        blocks: input.estimateChanges?.length
          ? [{
              type: 'table',
              headers: ['Indicateur', 'Avant', 'Après', 'Variation'],
              rows: input.estimateChanges.map((e) => [e.metric, e.before, e.after, e.change]),
            }]
          : [{ type: 'p', text: 'Nous ajustons nos estimations pour refléter la dynamique observée et la nouvelle guidance. Le détail figure dans le modèle mis à jour.' }],
      },
      {
        heading: 'Opinion',
        blocks: [
          { type: 'p', text: input.opinion || `Nous confirmons notre recommandation ${input.rating || 'Achat'} sur ${company}. La publication conforte notre thèse d'investissement et le momentum bénéficiaire reste favorable.` },
        ],
      },
    ],
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
