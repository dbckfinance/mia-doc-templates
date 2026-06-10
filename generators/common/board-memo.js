// generators/common/board-memo.js
// Mémo au conseil / note de décision (docx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildReport } from '../../shared/docx-report.js';

export const metadata = {
  id: 'board-memo',
  name: 'Mémo au Conseil (Note de Décision)',
  vertical: 'common',
  outputType: 'docx',
  estimatedPages: '3-6',
  requiredInput: ['subject'],
  optionalInput: ['company', 'decisionRequested', 'background', 'options', 'recommendation', 'risks', 'financialImpact', 'author'],
};

export async function generate(input = {}) {
  const subject = input.subject || 'Objet de la décision';

  return buildReport({
    docTitle: `Note au conseil — ${subject}`,
    docSubtitle: input.company || '',
    vertical: 'common',
    docId: 'board-memo',
    confidential: true,
    sections: [
      {
        heading: 'Décision demandée',
        blocks: [
          { type: 'p', text: input.decisionRequested || `Il est demandé au conseil d'approuver : ${subject}.` },
        ],
      },
      {
        heading: 'Contexte',
        blocks: [
          { type: 'p', text: input.background || 'Rappel du contexte stratégique et des événements ayant conduit à la présente proposition.' },
        ],
      },
      {
        heading: 'Options étudiées',
        blocks: input.options?.length
          ? [{
              type: 'table',
              headers: ['Option', 'Avantages', 'Inconvénients'],
              rows: input.options.map((o) => [o.name, (o.pros || []).join(' ; '), (o.cons || []).join(' ; ')]),
            }]
          : [{ type: 'bullets', items: [
              'Option 1 : statu quo — écartée (perte d\'opportunité)',
              'Option 2 : proposition recommandée — détaillée ci-dessous',
              'Option 3 : alternative — écartée (risque d\'exécution)',
            ] }],
      },
      {
        heading: 'Recommandation',
        blocks: [
          { type: 'p', text: input.recommendation || 'La direction recommande l\'approbation de la proposition, qui maximise la création de valeur au regard des risques identifiés.' },
        ],
      },
      {
        heading: 'Impact financier',
        blocks: input.financialImpact?.length
          ? [{ type: 'facts', facts: input.financialImpact }]
          : [{ type: 'p', text: 'Synthèse de l\'impact financier attendu (investissement, retour attendu, impact sur la liquidité et les covenants).' }],
      },
      {
        heading: 'Risques et facteurs d\'atténuation',
        blocks: [
          { type: 'bullets', items: input.risks || [
            'Risque d\'exécution — atténué par la gouvernance projet proposée',
            'Risque financier — limité par les conditions suspensives négociées',
          ] },
          ...(input.author ? [{ type: 'kv', label: 'Préparé par', value: input.author }] : []),
        ],
      },
    ],
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
