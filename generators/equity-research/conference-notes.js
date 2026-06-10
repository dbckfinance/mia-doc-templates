// generators/equity-research/conference-notes.js
// Notes de conférence / field trip (docx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildReport } from '../../shared/docx-report.js';

export const metadata = {
  id: 'conference-notes',
  name: 'Notes de Conférence / Field Trip',
  vertical: 'equity-research',
  outputType: 'docx',
  estimatedPages: '2-5',
  requiredInput: ['event'],
  optionalInput: ['date', 'companies', 'keyTakeaways', 'implications', 'analyst'],
};

export async function generate(input = {}) {
  const event = input.event || 'Conférence';
  const companies = input.companies || [];

  return buildReport({
    docTitle: `Notes — ${event}`,
    docSubtitle: input.date || new Date().toLocaleDateString('fr-FR'),
    vertical: 'equity-research',
    docId: 'conference-notes',
    confidential: false,
    sections: [
      {
        heading: 'Synthèse',
        blocks: [
          { type: 'bullets', items: input.keyTakeaways || [
            'Ton général constructif des managements rencontrés',
            'Confirmation des tendances de demande évoquées lors des dernières publications',
            'Discipline sur les prix maintenue malgré la pression concurrentielle',
          ] },
        ],
      },
      ...(companies.length
        ? companies.map((c) => ({
            heading: `${c.name}${c.ticker ? ` (${c.ticker})` : ''}`,
            blocks: [
              ...(c.speakers ? [{ type: 'kv', label: 'Intervenants', value: c.speakers }] : []),
              { type: 'bullets', items: c.notes || ['Notes de la rencontre à compléter'] },
              ...(c.implication ? [{ type: 'kv', label: 'Implication', value: c.implication }] : []),
            ],
          }))
        : [{
            heading: 'Rencontres',
            blocks: [{ type: 'p', text: 'Détail des rencontres société par société (management, points abordés, signaux faibles).' }],
          }]),
      {
        heading: 'Implications pour nos recommandations',
        blocks: [
          { type: 'bullets', items: input.implications || [
            'Pas de changement de recommandation à l\'issue de l\'événement',
            'Points de vigilance à suivre lors des prochaines publications',
          ] },
          { type: 'kv', label: 'Analyste', value: input.analyst || 'M&IA Research' },
        ],
      },
    ],
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
