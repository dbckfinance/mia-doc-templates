// generators/common/engagement-letter.js
// Lettre de mission / mandat (docx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildReport } from '../../shared/docx-report.js';

export const metadata = {
  id: 'engagement-letter',
  name: 'Lettre de Mission / Mandat',
  vertical: 'common',
  outputType: 'docx',
  estimatedPages: '4-8',
  requiredInput: ['client'],
  optionalInput: ['advisor', 'missionType', 'scope', 'fees', 'duration', 'exclusivity', 'terminationTerms'],
};

export async function generate(input = {}) {
  const client = input.client || 'Client';
  const advisor = input.advisor || 'M&IA Advisory';
  const missionType = input.missionType || 'mandat de conseil en cession';

  return buildReport({
    docTitle: 'Lettre de mission',
    docSubtitle: `${advisor} / ${client}`,
    vertical: 'common',
    docId: 'engagement-letter',
    confidential: true,
    sections: [
      {
        heading: 'Article 1 — Objet de la mission',
        blocks: [
          { type: 'p', text: `${client} (le « Client ») confie à ${advisor} (le « Conseil ») un ${missionType}, dans les conditions définies par la présente lettre de mission.` },
        ],
      },
      {
        heading: 'Article 2 — Périmètre des prestations',
        blocks: [
          { type: 'bullets', items: input.scope || [
            'Analyse stratégique et financière préalable',
            'Préparation de la documentation marketing (teaser, mémorandum d\'information)',
            'Identification et approche des contreparties potentielles',
            'Organisation et coordination du processus (data room, management presentations)',
            'Assistance à la négociation jusqu\'à la signature de la documentation définitive',
          ] },
        ],
      },
      {
        heading: 'Article 3 — Rémunération',
        blocks: input.fees?.length
          ? [{
              type: 'table',
              headers: ['Composante', 'Montant / taux', 'Exigibilité'],
              rows: input.fees.map((f) => [f.label, f.amount, f.due || '—']),
            }]
          : [{ type: 'p', text: 'La rémunération du Conseil comprend des honoraires fixes (retainer) et des honoraires de succès calculés sur la valeur de la transaction, dont le détail figure en annexe. Les honoraires de succès ne sont dus qu\'en cas de réalisation effective de l\'opération.' }],
      },
      {
        heading: 'Article 4 — Durée et exclusivité',
        blocks: [
          { type: 'kv', label: 'Durée de la mission', value: input.duration || '12 mois renouvelables par accord écrit' },
          { type: 'kv', label: 'Exclusivité', value: input.exclusivity || 'Le Client confie la mission au Conseil à titre exclusif pendant toute la durée du mandat.' },
        ],
      },
      {
        heading: 'Article 5 — Confidentialité',
        blocks: [
          { type: 'p', text: 'Chaque partie s\'engage à préserver la confidentialité des informations échangées dans le cadre de la mission, pendant sa durée et les 24 mois suivant son terme.' },
        ],
      },
      {
        heading: 'Article 6 — Résiliation',
        blocks: [
          { type: 'p', text: input.terminationTerms || 'Chaque partie peut résilier la mission moyennant un préavis écrit de 30 jours. En cas de réalisation d\'une opération avec une contrepartie approchée par le Conseil dans les 18 mois suivant la résiliation, les honoraires de succès restent dus (tail period).' },
        ],
      },
      {
        heading: 'Article 7 — Responsabilité et droit applicable',
        blocks: [
          { type: 'p', text: 'La responsabilité du Conseil est limitée au montant des honoraires effectivement perçus. La présente lettre est régie par le droit français ; tout litige relève de la compétence du Tribunal de commerce de Paris.' },
        ],
      },
      {
        heading: 'Signatures',
        blocks: [
          { type: 'kv', label: `Pour ${client}`, value: '___________________ (nom, qualité, date)' },
          { type: 'kv', label: `Pour ${advisor}`, value: '___________________ (nom, qualité, date)' },
        ],
      },
    ],
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
