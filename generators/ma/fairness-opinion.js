// generators/ma/fairness-opinion.js
// Attestation d'équité / fairness opinion letter (docx)

import { runCli, isCliInvocation } from '../../shared/cli.js';
import { buildReport } from '../../shared/docx-report.js';
import { disclaimerFor } from '../../shared/disclaimer-library.js';
import { formatCurrency, formatMultiple } from '../../shared/financial-helpers.js';

export const metadata = {
  id: 'fairness-opinion',
  name: "Attestation d'Équité (Fairness Opinion)",
  vertical: 'ma',
  outputType: 'docx',
  estimatedPages: '10-20',
  requiredInput: ['target', 'offer'],
  optionalInput: ['client', 'methods', 'conclusion', 'limitations'],
};

export async function generate(input = {}) {
  const target = input.target || 'Société';
  const offer = input.offer || {};
  const sections = [];

  sections.push({
    heading: 'Contexte de la mission',
    paragraphs: [
      input.context || `${input.client || 'Le conseil d\'administration'} a sollicité M&IA afin d'émettre une opinion sur le caractère équitable, d'un point de vue financier, des conditions proposées dans le cadre de l'offre portant sur ${target}.`,
    ],
    facts: [
      ['Société visée', target],
      ['Initiateur', offer.bidder || '—'],
      ['Prix offert', offer.price ? formatCurrency(offer.price, { compact: false }) : '—'],
      ['Prime induite', offer.premium || '—'],
      ['Nature de l\'offre', offer.type || '—'],
    ],
  });

  sections.push({
    heading: 'Diligences effectuées',
    bullets: input.diligences || [
      'Analyse des comptes historiques et du plan d\'affaires de la société',
      'Entretiens avec le management',
      'Analyse des conditions de marché et des transactions comparables',
      'Mise en œuvre d\'une approche multicritères de valorisation',
    ],
  });

  if (input.methods?.length) {
    sections.push({
      heading: 'Approche de valorisation multicritères',
      table: {
        headers: ['Méthode', 'Valeur basse', 'Valeur haute', 'Prime / (décote) induite'],
        rows: input.methods.map((m) => [
          m.name,
          m.low ?? '—',
          m.high ?? '—',
          m.premium ?? '—',
        ]),
      },
      paragraphs: input.methodsCommentary ? [input.methodsCommentary] : [],
    });
  }

  sections.push({
    heading: 'Conclusion',
    paragraphs: [
      input.conclusion || `Sur la base des travaux décrits ci-avant, et sous réserve des limites exposées, nous sommes d'avis que les conditions financières proposées sont équitables pour les actionnaires de ${target}.`,
    ],
  });

  sections.push({
    heading: 'Limites de l\'opinion',
    bullets: input.limitations || [
      'L\'opinion repose sur les informations communiquées par la société, dont l\'exactitude n\'a pas fait l\'objet d\'un audit indépendant',
      'L\'opinion est émise à la date du présent rapport et ne tient pas compte d\'évènements postérieurs',
      'L\'opinion ne constitue pas une recommandation d\'apporter ou non les titres à l\'offre',
    ],
  });

  return buildReport({
    docTitle: "Attestation d'Équité",
    docSubtitle: target,
    project: input.project,
    date: input.date,
    author: 'M&IA — Expert indépendant',
    lang: 'fr',
    sections,
    disclaimer: disclaimerFor('ma', metadata.id),
  });
}

if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
