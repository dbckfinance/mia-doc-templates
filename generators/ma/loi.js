// generators/ma/loi.js
// Lettre d'intention / LOI (docx) — draft à faire valider par un juriste

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildLetter } from '../../shared/docx-report.js';
import { disclaimerFor } from '../../shared/disclaimer-library.js';
import { formatCurrency } from '../../shared/financial-helpers.js';

export const metadata = {
  id: 'loi',
  name: "Lettre d'Intention (LOI)",
  vertical: 'ma',
  outputType: 'docx',
  estimatedPages: '4-8',
  requiredInput: ['buyer', 'target'],
  optionalInput: ['price', 'structure', 'conditions', 'exclusivity', 'timeline', 'financing'],
};

export async function generate(input = {}) {
  const buyer = input.buyer || '[Acquéreur]';
  const target = input.target || '[Société Cible]';
  const seller = input.seller || '[Vendeur]';

  return buildLetter({
    docTitle: `Lettre d'Intention — ${target}`,
    subject: `Lettre d'intention relative à l'acquisition de ${target}`,
    recipientLines: [seller],
    senderLines: [buyer],
    date: input.date,
    lang: 'fr',
    sections: [
      {
        paragraphs: [
          `Nous vous remercions de l'opportunité qui nous est offerte d'étudier l'acquisition de ${target} (la « Société »). La présente lettre d'intention (la « Lettre ») a pour objet de formaliser notre intérêt et les principaux termes et conditions de notre offre indicative, laquelle demeure non engageante.`,
        ],
      },
      {
        heading: '1. Prix et structure',
        keyValues: [
          ['Valeur d\'entreprise proposée', input.price ? formatCurrency(input.price, { compact: true }) : '[À compléter]'],
          ['Structure', input.structure || "Acquisition de 100% des titres en numéraire"],
          ['Mécanisme d\'ajustement', input.adjustment || 'Locked box / completion accounts à discuter'],
        ],
      },
      {
        heading: '2. Financement',
        paragraphs: [
          input.financing || `L'acquisition serait financée par une combinaison de fonds propres et de dette d'acquisition. Nous confirmons disposer des ressources nécessaires à la réalisation de l'opération.`,
        ],
      },
      {
        heading: '3. Conditions',
        bullets: input.conditions || [
          'Réalisation satisfaisante des due diligences (financière, juridique, fiscale, sociale)',
          'Obtention des autorisations réglementaires requises',
          'Approbation des organes de gouvernance de l\'acquéreur',
          'Négociation d\'une documentation juridique satisfaisante',
        ],
      },
      {
        heading: '4. Exclusivité',
        paragraphs: [
          input.exclusivity || `Nous sollicitons une période d'exclusivité de huit (8) semaines à compter de l'acceptation de la présente Lettre afin de mener à bien nos due diligences et de finaliser la documentation.`,
        ],
      },
      {
        heading: '5. Calendrier indicatif',
        bullets: input.timeline || [
          'Semaines 1-4 : due diligences confirmatoires',
          'Semaines 5-6 : négociation de la documentation',
          'Semaines 7-8 : signing',
          'Closing : sous réserve des conditions suspensives',
        ],
      },
      {
        heading: '6. Caractère non engageant',
        paragraphs: [
          `La présente Lettre ne constitue pas un engagement ferme d'acquérir. Seule la signature d'une documentation définitive engagera les parties, à l'exception des stipulations relatives à la confidentialité et à l'exclusivité qui ont un caractère obligatoire.`,
        ],
      },
    ],
    signatureLines: [`Pour ${buyer}`, '', '_______________________', input.signatory || '[Nom, Fonction]'],
    disclaimer: disclaimerFor('ma', metadata.id),
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
