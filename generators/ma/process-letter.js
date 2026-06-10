// generators/ma/process-letter.js
// Process letter (docx) — instructions aux candidats acquéreurs

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildLetter } from '../../shared/docx-report.js';
import { disclaimerFor } from '../../shared/disclaimer-library.js';

export const metadata = {
  id: 'process-letter',
  name: 'Process Letter',
  vertical: 'ma',
  outputType: 'docx',
  estimatedPages: '3-6',
  requiredInput: ['project'],
  optionalInput: ['recipient', 'phase', 'deadline', 'offerRequirements', 'rules', 'contacts'],
};

export async function generate(input = {}) {
  const project = input.project || 'Projet';
  const phase = input.phase || 'Phase 1 — Offres indicatives';

  return buildLetter({
    docTitle: `Process Letter — ${project}`,
    subject: `${project} — ${phase}`,
    recipientLines: input.recipient ? [input.recipient] : ['Aux candidats acquéreurs'],
    date: input.date,
    lang: 'fr',
    sections: [
      {
        paragraphs: [
          input.intro || `Dans le cadre du processus de cession relatif au ${project}, nous vous remercions de l'intérêt que vous portez à cette opportunité. La présente lettre a pour objet de préciser les modalités de remise de votre offre.`,
        ],
      },
      {
        heading: 'Calendrier',
        keyValues: [
          ['Date limite de remise des offres', input.deadline || 'À confirmer'],
          ['Format', input.format || 'Offre écrite en français ou anglais, format PDF'],
          ['Destinataire', input.submitTo || 'M&IA — équipe Transaction'],
        ],
      },
      {
        heading: "Contenu attendu de l'offre",
        bullets: input.offerRequirements || [
          'Prix proposé (valeur d\'entreprise) et hypothèses de passage à la valeur des titres',
          'Structure et conditions de financement envisagées',
          'Périmètre de la transaction et conditions suspensives',
          'Approbations internes requises et calendrier de décision',
          'Due diligences complémentaires souhaitées',
          'Intentions concernant le management et les équipes',
        ],
      },
      {
        heading: 'Règles du processus',
        bullets: input.rules || [
          'Toute communication relative au processus doit passer exclusivement par M&IA',
          'Il est strictement interdit de contacter la société, ses dirigeants, clients ou fournisseurs',
          'Le vendeur se réserve le droit de modifier le processus ou d\'y mettre fin à tout moment',
          'Les frais engagés par les candidats restent à leur charge exclusive',
        ],
      },
      {
        heading: 'Contacts',
        bullets: input.contacts || ['M&IA — équipe Transaction'],
      },
    ],
    signatureLines: [input.signature || 'M&IA', 'Conseil du vendeur'],
    disclaimer: disclaimerFor('ma', metadata.id),
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
