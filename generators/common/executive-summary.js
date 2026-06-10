// generators/common/executive-summary.js
// Résumé exécutif générique (docx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildReport } from '../../shared/docx-report.js';

export const metadata = {
  id: 'executive-summary',
  name: 'Résumé Exécutif',
  vertical: 'common',
  outputType: 'docx',
  estimatedPages: '2-4',
  requiredInput: ['title'],
  optionalInput: ['context', 'keyPoints', 'recommendations', 'nextSteps', 'facts', 'author'],
};

export async function generate(input = {}) {
  const title = input.title || 'Résumé exécutif';

  return buildReport({
    docTitle: title,
    docSubtitle: input.subtitle || 'Résumé exécutif',
    vertical: 'common',
    docId: 'executive-summary',
    confidential: input.confidential !== false,
    sections: [
      {
        heading: 'Contexte',
        blocks: [
          { type: 'p', text: input.context || 'Contexte et objet du document : rappel de la demande, du périmètre et de la méthodologie employée.' },
          ...(input.facts?.length ? [{ type: 'facts', facts: input.facts }] : []),
        ],
      },
      {
        heading: 'Points clés',
        blocks: [
          { type: 'bullets', items: input.keyPoints || [
            'Premier enseignement majeur de l\'analyse',
            'Deuxième enseignement majeur',
            'Troisième enseignement majeur',
          ] },
        ],
      },
      {
        heading: 'Recommandations',
        blocks: [
          { type: 'bullets', items: input.recommendations || [
            'Recommandation principale',
            'Recommandation secondaire',
          ] },
        ],
      },
      {
        heading: 'Prochaines étapes',
        blocks: [
          { type: 'bullets', items: input.nextSteps || [
            'Validation des conclusions avec les parties prenantes',
            'Lancement du plan d\'action',
          ] },
          ...(input.author ? [{ type: 'kv', label: 'Auteur', value: input.author }] : []),
        ],
      },
    ],
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
